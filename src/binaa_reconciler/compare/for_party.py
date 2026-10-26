"""One file for both parties, open on both screens during the phone call.

One right-to-left sheet, sorted by date, one line per operation to
settle: each party's own date, description and amount(s) side by side, as in
its books (red, empty: missing there — with in whose favour it is, clear for
both readers; yellow: a different amount; orange: to confirm), and an empty
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
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.styles import Side as Edge

from binaa_reconciler.compare.engine import Kind, Line, Result, units
from binaa_reconciler.compare.manual import ORANGE, RED, YELLOW, arabic
from binaa_reconciler.core.models import LedgerSide

HEAD = PatternFill("solid", fgColor="D9E1F2")
CELL = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
THIN = Edge(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"


@dataclass(frozen=True)
class Part:
    """One party's part of a line, exactly as in its books (empty: missing)."""

    lines: tuple[Line, ...] = ()

    @property
    def day(self) -> str:
        return " ، ".join(sorted({f"{x.day:%d/%m/%Y}" for x in self.lines}))

    @property
    def label(self) -> str:
        return " / ".join(x.label for x in self.lines if x.label)

    @property
    def amount(self) -> int | str | None:
        """One amount as a number; several (a sum) written one by one, so each
        can be found in RIM-Soft."""
        if not self.lines:
            return None
        if len(self.lines) == 1:
            return units(self.lines[0].amount)
        return " + ".join(f"{units(x.amount):,}" for x in self.lines)


@dataclass(frozen=True)
class Row:
    ours: Part
    theirs: Part
    note: str
    fill: PatternFill  # of the amounts

    @property
    def first_day(self) -> date:
        return min(x.day for x in (*self.ours.lines, *self.theirs.lines))


def _favour(x: Line, ours: bool, names: tuple[str, str]) -> str:
    """In our books a debit is owed by them (in our favour); in theirs, a debit
    is owed by us (in their favour)."""
    us, them = (arabic(n) for n in names)
    return f"لصالح {us if (x.side is LedgerSide.DEBIT) == ours else them}"


def rows(result: Result) -> list[Row]:
    us, them = (arabic(n) for n in result.names)
    out = [
        Row(Part((x,)), Part(), f"ناقصة عند {them} — {_favour(x, True, result.names)}", RED)
        for x in result.only[0]
    ]
    out += [
        Row(Part(), Part((y,)), f"ناقصة عند {us} — {_favour(y, False, result.names)}", RED)
        for y in result.only[1]
    ]
    for pair in result.close:
        gap = abs(sum(x.amount for x in pair.a) - sum(x.amount for x in pair.b))
        out.append(Row(Part(pair.a), Part(pair.b), f"فرق في المبلغ: {units(gap):,}", YELLOW))
    for pair in result.possible:
        if not (pair.a and pair.b) or pair.kind is Kind.REVERSAL:
            continue  # a movement and its reversal in one file: not theirs to check
        note = "قد تكون نفس العملية"
        if pair.kind is Kind.SAME:
            note += " — بتاريخ مختلف"
        elif pair.kind is Kind.SUM:
            note += " — مجموع عدة عمليات"
        elif pair.kind is Kind.SIDE:
            note += " — بنفس الجهة عند الطرفين"
        out.append(Row(Part(pair.a), Part(pair.b), note, ORANGE))
    out.sort(key=lambda r: r.first_day)
    return out


def build(result: Result) -> bytes:
    us, them = (arabic(n) for n in result.names)
    wb = Workbook()
    ws = wb.active if wb.active is not None else wb.create_sheet()
    ws.title = f"{us} و{them}"[:31]
    ws.sheet_view.rightToLeft = True
    heads = (
        f"تاريخ {us}",
        f"بيان {us}",
        f"مبلغ {us}",
        f"تاريخ {them}",
        f"بيان {them}",
        f"مبلغ {them}",
        "ملاحظة",
        "الرد",
    )
    for col, width in zip("ABCDEFGH", (12, 26, 14, 12, 26, 14, 30, 18), strict=True):
        ws.column_dimensions[col].width = width

    def wide(text: str, size: int = 11) -> None:
        """A line of text across the whole table (never cut by a narrow column)."""
        ws.append([text])
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=8)
        ws.cell(ws.max_row, 1).font = Font(bold=size > 11, size=size)
        ws.cell(ws.max_row, 1).alignment = Alignment(horizontal="right", readingOrder=2)

    wide(f"مقارنة الحساب بين {us} و{them}", 14)
    if result.period:
        wide(f"الفترة: {result.period[0]:%d/%m/%Y} — {result.period[1]:%d/%m/%Y}")
    wide("الأحمر: ناقصة عند أحد الطرفين — الأصفر: مبلغ مختلف — البرتقالي: للتأكد")
    ws.append([])
    ws.append(list(heads))
    for cell in ws[ws.max_row]:
        cell.font, cell.fill, cell.alignment, cell.border = Font(bold=True), HEAD, CELL, BOX
    ws.freeze_panes = f"A{ws.max_row + 1}"
    for r in rows(result):
        o, t = r.ours, r.theirs
        ws.append([o.day, o.label, o.amount, t.day, t.label, t.amount, r.note, None])
        line = ws[ws.max_row]
        for cell in line:
            cell.alignment, cell.border = CELL, BOX
        for cells, side in ((line[0:3], o), (line[3:6], t)):
            cells[2].number_format = MONEY
            if r.fill is not RED or not side.lines:  # red: the missing side only
                for cell in cells:
                    cell.fill = r.fill
    gaps = [r for r in rows(result) if r.fill is YELLOW]
    if gaps:  # the one total that means something: the amount gaps
        total = sum(
            abs(sum(x.amount for x in r.ours.lines) - sum(x.amount for x in r.theirs.lines))
            for r in gaps
        )
        total = units(total)
        ws.append([])
        wide(f"مجموع فروق المبالغ ({len(gaps)} عملية): {total:,}", 12)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save(result: Result, target: Path) -> Path:
    target.write_bytes(build(result))
    return target


__all__ = ["Row", "build", "rows", "save"]
