"""One file for both parties, open on both screens during the phone call.

One right-to-left sheet, sorted by date, one numbered line per operation to
settle: the amount as each party has it side by side (an empty, red cell is
what is missing in that party's books; two different amounts are yellow; a
possible pair is orange), in whose favour it is («لصالح بناء»: clear for
both, whoever reads the file), and an empty
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
    our_label: str  # each party's own description (empty where missing)
    their_label: str
    ours: int | None  # whole units, as in our books (None: missing)
    theirs: int | None  # as in theirs
    owed: str  # «لصالح <name>»: who the operation favours
    note: str
    fill: PatternFill  # of the amounts


def _favour(x: Line, ours: bool, names: tuple[str, str]) -> str:
    """In our books a debit is owed by them (in our favour); in theirs, a debit
    is owed by us (in their favour)."""
    us, them = (arabic(n) for n in names)
    return f"لصالح {us if (x.side is LedgerSide.DEBIT) == ours else them}"


def _labels(lines: tuple[Line, ...]) -> str:
    return " / ".join(x.label for x in lines if x.label)


def _total(lines: tuple[Line, ...]) -> int:
    return units(sum(x.amount for x in lines))


def rows(result: Result) -> list[Row]:
    out = [
        Row(x.day, x.label, "", units(x.amount), None, _favour(x, True, result.names), "", RED)
        for x in result.only[0]
    ]
    out += [
        Row(y.day, "", y.label, None, units(y.amount), _favour(y, False, result.names), "", RED)
        for y in result.only[1]
    ]
    for pair in result.close:
        gap = abs(_total(pair.a) - _total(pair.b))
        out.append(
            Row(
                pair.b[0].day,
                _labels(pair.a),
                _labels(pair.b),
                _total(pair.a),
                _total(pair.b),
                _favour(pair.b[0], False, result.names),
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
        out.append(
            Row(
                days[0],
                _labels(pair.a),
                _labels(pair.b),
                _total(pair.a),
                _total(pair.b),
                _favour(pair.b[0], False, result.names),
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
        f"بيان {us}",
        f"عند {us}",
        f"بيان {them}",
        f"عند {them}",
        "لصالح",
        "ملاحظة",
        "الرد",
    )
    for col, width in zip("ABCDEFGHI", (6, 12, 28, 13, 28, 13, 14, 26, 18), strict=True):
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
        ws.append([n, r.day, r.our_label, r.ours, r.their_label, r.theirs, r.owed, r.note, None])
        line = ws[ws.max_row]
        for cell in line:
            cell.alignment, cell.border = CELL, BOX
        line[1].number_format = "DD/MM/YYYY"
        for cell in (line[3], line[5]):
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
