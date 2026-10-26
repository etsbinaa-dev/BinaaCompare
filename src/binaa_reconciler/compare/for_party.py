"""One file for both parties, open on both screens during the phone call.

One right-to-left sheet, sorted by date, one numbered line per operation to
settle: the amount as each party has it side by side (an empty, red cell is
what is missing in that party's books; two different amounts are yellow; a
possible pair is orange), «لكم» / «عليكم» for the other party, and an empty
«الرد» column to fill in during the call. No other totals: the two files rarely
start and end on the same day; only the amount gaps are added up. Nothing
else of our books.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from binaa_reconciler.compare.engine import Kind, Line, Result, units
from binaa_reconciler.compare.manual import ORANGE, RED, YELLOW, arabic
from binaa_reconciler.core.models import LedgerSide

HEAD = PatternFill("solid", fgColor="D9E1F2")
CELL = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"


@dataclass(frozen=True)
class Row:
    day: date
    label: str
    ours: int | None  # whole units, as in our books (None: missing)
    theirs: int | None  # as in theirs
    owed: str  # «لكم» / «عليكم», for the other party
    note: str
    fill: PatternFill  # of the amounts


def _for_them(x: Line, ours: bool) -> str:
    """In our books a debit is owed by them; in theirs, a debit is owed by us."""
    return "عليكم" if (x.side is LedgerSide.DEBIT) == ours else "لكم"


def _total(lines: tuple[Line, ...]) -> int:
    return units(sum(x.amount for x in lines))


def rows(result: Result) -> list[Row]:
    out = [
        Row(x.day, x.label, units(x.amount), None, _for_them(x, True), "", RED)
        for x in result.only[0]
    ]
    out += [
        Row(y.day, y.label, None, units(y.amount), _for_them(y, False), "", RED)
        for y in result.only[1]
    ]
    for pair in result.close:
        gap = abs(_total(pair.a) - _total(pair.b))
        out.append(
            Row(
                pair.b[0].day,
                pair.b[0].label or pair.a[0].label,
                _total(pair.a),
                _total(pair.b),
                _for_them(pair.b[0], False),
                f"فرق {gap:,}",
                YELLOW,
            )
        )
    for pair in result.possible:
        if not (pair.a and pair.b) or pair.kind is Kind.REVERSAL:
            continue  # a movement and its reversal in one file: not theirs to check
        days = sorted({x.day for x in pair.lines})
        note = "قد تكون نفس العملية"
        if pair.kind is Kind.SAME:
            note += f" — فرق {(days[-1] - days[0]).days} يوم في التاريخ"
        elif pair.kind is Kind.SUM:
            note += f" — مجموع {max(len(pair.a), len(pair.b))} عمليات"
        elif pair.kind is Kind.SIDE:
            note += " — نفس الجهة عند الطرفين"
        label = " / ".join(x.label for x in pair.b if x.label) or pair.a[0].label
        out.append(
            Row(
                days[0],
                label,
                _total(pair.a),
                _total(pair.b),
                _for_them(pair.b[0], False),
                note,
                ORANGE,
            )
        )
    out.sort(key=lambda r: r.day)
    return out


def build(result: Result) -> bytes:
    us, them = (arabic(n) for n in result.names)
    wb = Workbook()
    ws = wb.active if wb.active is not None else wb.create_sheet()
    ws.title = f"{us} و{them}"[:31]
    ws.sheet_view.rightToLeft = True
    heads = (
        "رقم",
        "التاريخ",
        "البيان",
        f"عند {us}",
        f"عند {them}",
        "لكم / عليكم",
        "ملاحظة",
        "الرد",
    )
    for col, width in zip("ABCDEFGH", (6, 12, 34, 14, 14, 9, 30, 22), strict=True):
        ws.column_dimensions[col].width = width
    ws.append([f"مقارنة الحساب بين {us} و{them}"])
    ws["A1"].font = Font(bold=True, size=14)
    if result.period:
        ws.append([f"الفترة: {result.period[0]:%d/%m/%Y} — {result.period[1]:%d/%m/%Y}"])
    ws.append(
        ["الخانة الحمراء الفارغة: غير مسجلة عند صاحبها — الصفراء: مبلغ مختلف — البرتقالية: للتأكد"]
    )
    ws.append([])
    ws.append(list(heads))
    for cell in ws[ws.max_row]:
        cell.font, cell.fill, cell.alignment, cell.border = Font(bold=True), HEAD, CELL, BOX
    ws.freeze_panes = f"A{ws.max_row + 1}"
    for n, r in enumerate(rows(result), start=1):
        ws.append([n, r.day, r.label, r.ours, r.theirs, r.owed, r.note, None])
        line = ws[ws.max_row]
        for cell in line:
            cell.alignment, cell.border = CELL, BOX
        line[1].number_format = "DD/MM/YYYY"
        for cell in line[3:5]:
            cell.number_format = MONEY
            cell.fill = r.fill if (r.fill is not RED or cell.value is None) else PatternFill()
    gaps = [r for r in rows(result) if r.fill is YELLOW]
    if gaps:  # the one total that means something: the amount gaps
        total = sum(abs((r.ours or 0) - (r.theirs or 0)) for r in gaps)
        ws.append([])
        ws.append([f"مجموع فروق المبالغ ({len(gaps)} عملية): {total:,}"])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save(result: Result, target: Path) -> Path:
    target.write_bytes(build(result))
    return target


__all__ = ["Row", "build", "rows", "save"]
