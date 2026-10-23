"""The two original files, for a manual comparison side by side.

One workbook, one sheet per file: the original columns and rows as RIM-Soft
exported them, without the matched operations, and a coloured «الحالة» column:
red (on this side only), orange (possibly the same operation), yellow (a
different amount), grey (outside the chosen period). The original files are
never modified: only a new file is written, where the user chooses.
"""

from __future__ import annotations

from copy import copy
from io import BytesIO
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from binaa_reconciler.compare.engine import (
    SIDE_AR,
    Kind,
    Line,
    Pair,
    Result,
    amount_text,
    mirror,
)
from binaa_reconciler.compare.files import Loaded
from binaa_reconciler.compare.texts import reason

RED = PatternFill("solid", fgColor="F8CBAD")
ORANGE = PatternFill("solid", fgColor="FCE4D6")
YELLOW = PatternFill("solid", fgColor="FFF2CC")
GREY = PatternFill("solid", fgColor="E7E6E6")
_FORBIDDEN = str.maketrans(dict.fromkeys("[]:*?/\\", " "))


def _other(pair: Pair, line: Line) -> str:
    """The other side of a pair, as «dd/mm amount (+ ...)»."""
    side = pair.b if line in pair.a else pair.a
    if not side:  # a reversal in the same file
        side = tuple(x for x in pair.lines if x is not line)
    return " + ".join(f"{x.day:%d/%m} {amount_text(x.amount)}" for x in side)


def statuses(result: Result, i: int) -> tuple[set[int], dict[int, tuple[str, PatternFill]]]:
    """Rows to drop (matched) and the status of each remaining row of file i."""
    a, b = result.names
    me, other = (a, b) if i == 0 else (b, a)
    matched: set[int] = set()
    for x, y in result.matched:
        matched.add((x, y)[i].row)
    for pair in result.confirmed:
        matched.update(x.row for x in (pair.a if i == 0 else pair.b))
        if pair.kind is Kind.REVERSAL:
            matched.update(x.row for x in pair.lines if x.key[0] == "AB"[i])
    status: dict[int, tuple[str, PatternFill]] = {}
    for pair in result.possible:
        for x in pair.lines:
            if x.key[0] == "AB"[i]:
                text = f"قد يكون: {reason(pair, result.names)} — المقابل: {_other(pair, x)}"
                status[x.row] = (text, ORANGE)
    for pair in result.close:
        for x in pair.lines:
            if x.key[0] == "AB"[i]:
                text = f"{reason(pair, result.names)} — المقابل عند {other}: {_other(pair, x)}"
                status[x.row] = (text, YELLOW)
    for x in result.only[i]:
        note = result.notes.get(x.key)
        text = f"عند {me} وليس عند {other}" + (f" — {note}" if note else "")
        status[x.row] = (text, RED)
    return matched, status


def fixes(result: Result, i: int) -> dict[int, str]:
    """The suggested correction in RIM-Soft for each remaining row of file i.
    Who is right is decided together from the documents: when it is not
    known, both possibilities are written."""
    a, b = result.names
    me, other = (a, b) if i == 0 else (b, a)
    mine = "AB"[i]
    out: dict[int, str] = {}
    for pair in result.possible:
        for x in pair.lines:
            if x.key[0] != mine:
                continue
            if pair.kind is Kind.REVERSAL:
                out[x.row] = "لا تصحيح: الحركة وعكسها أثرهما صفر"
            elif pair.kind is Kind.SAME:
                out[x.row] = "لا تصحيح في المبلغ: التاريخ فقط مختلف (يُعدَّل إن أردتم التطابق)"
            elif pair.kind is Kind.SUM:
                out[x.row] = "لا تصحيح: المجموع مساوٍ — يكفي التأكد من الأوراق"
            else:  # SIDE: the same side on both files
                side, opposite = SIDE_AR[x.side], SIDE_AR[mirror(x.side)]
                out[x.row] = (
                    f"أحدكما سجّلها في الجهة الخطأ: إن أخطأ {me} يحوّل قيده من {side} "
                    f"إلى {opposite} — وإن أخطأ {other} يحوّل قيده من {side} إلى {opposite}"
                )
    for pair in result.close:
        for x in pair.lines:
            if x.key[0] != mine:
                continue
            theirs = pair.b if x in pair.a else pair.a
            o = amount_text(sum(y.amount for y in theirs))
            m = amount_text(x.amount)
            out[x.row] = (
                f"اتفقوا على المبلغ الصحيح: إن كان {o} يعدّل {me} قيده إلى {o} — "
                f"وإن كان {m} يعدّل {other} قيده إلى {m}"
            )
    for x in result.only[i]:
        out[x.row] = (
            f"إن كانت صحيحة: يسجلها {other} {SIDE_AR[mirror(x.side)]} {amount_text(x.amount)} "
            f"بتاريخ {x.day:%d/%m/%Y} — وإن كانت خطأ: يلغيها {me}"
        )
    return out


def _copy_sheet(target: Worksheet, loaded: Loaded, result: Result, i: int) -> None:
    source_wb = load_workbook(BytesIO(loaded.data), data_only=True)
    source = source_wb.worksheets[loaded.sheet_index]
    matched, status = statuses(result, i)
    fix = fixes(result, i)
    transaction_rows = {x.row for x in loaded.lines}
    second_rows = {v: k for k, v in loaded.continuation.items()}  # 2nd row -> 1st
    width = source.max_column
    status_col = width + 1
    out = 0
    for n, row in enumerate(source.iter_rows(), start=1):
        first = second_rows.get(n, n)  # a 2nd row follows its operation
        if n > loaded.header_row and (first not in transaction_rows or first in matched):
            continue  # matched (or a total / balance line: no longer right)
        out += 1
        for cell in row:
            value: Any = cell.value
            new = target.cell(out, cell.column or 1, value)
            if cell.has_style:
                new.number_format = cell.number_format
                new.font = copy(cell.font)
        if n == loaded.header_row:
            for col, title in ((status_col, "الحالة"), (status_col + 1, "التصحيح المقترح")):
                target.cell(out, col, title).font = Font(bold=True)
        elif n > loaded.header_row:
            text, fill = status.get(first, ("خارج الفترة المختارة", GREY))
            if n == first:
                target.cell(out, status_col, text)
                target.cell(out, status_col + 1, fix.get(first, "يُراجع في الفترة التالية"))
            for col in range(1, status_col + 2):
                target.cell(out, col).fill = fill
    for col in range(1, width + 1):
        letter = get_column_letter(col)
        target.column_dimensions[letter].width = source.column_dimensions[letter].width or 14
    target.column_dimensions[get_column_letter(status_col)].width = 60
    target.column_dimensions[get_column_letter(status_col + 1)].width = 90
    target.freeze_panes = f"A{loaded.header_row + 1}"


def build(result: Result, a: Loaded, b: Loaded) -> bytes:
    wb = Workbook()
    if wb.active is not None:
        wb.remove(wb.active)
    for i, loaded in enumerate((a, b)):
        ws: Worksheet = wb.create_sheet(result.names[i].translate(_FORBIDDEN)[:31] or "AB"[i])
        _copy_sheet(ws, loaded, result, i)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save(result: Result, a: Loaded, b: Loaded, target: Path) -> Path:
    target.write_bytes(build(result, a, b))
    return target


__all__ = ["build", "fixes", "save", "statuses"]
