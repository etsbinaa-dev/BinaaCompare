"""One file for both parties, open on both screens during the phone call.

Five columns, right to left, sorted by date: the date and description of
whoever has the operation, each party's amount, and a short note. Only what
needs the call: what is missing on one side (red, with in whose favour it
is — clear for both readers) and the amount gaps (yellow, with the other
date and description when they differ). The possible pairs are decided in
the application beforehand. Only the amount gaps are added up: the two files
rarely start and end on the same day. Nothing else of our books.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill
from openpyxl.styles import Side as Edge

from binaa_reconciler.compare.engine import Line, Result, units
from binaa_reconciler.compare.manual import RED, YELLOW
from binaa_reconciler.compare.texts import arabic
from binaa_reconciler.core.models import LedgerSide

HEAD = PatternFill("solid", fgColor="D9E1F2")
CELL = Alignment(horizontal="right", vertical="center", wrap_text=True, readingOrder=2)
THIN = Edge(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
MONEY = "#,##0"
MISSING = PatternFill("solid", fgColor="FF8B8B")


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
    for pair in result.close:  # the possible pairs are decided in the application
        gap = abs(sum(x.amount for x in pair.a) - sum(x.amount for x in pair.b))
        note = f"فرق في المبلغ: {units(gap):,}"
        if Part(pair.b).day != Part(pair.a).day:
            note += f" — عند {them} بتاريخ {Part(pair.b).day}"
        if Part(pair.b).label:
            note += f" — بيان {them}: {Part(pair.b).label}"
        out.append(Row(Part(pair.a), Part(pair.b), note, YELLOW))
    out.sort(key=lambda r: r.first_day)
    return out


def build(result: Result) -> bytes:
    us, them = (arabic(n) for n in result.names)
    wb = Workbook()
    ws = wb.active if wb.active is not None else wb.create_sheet()
    ws.title = f"{us} و{them}"[:31]
    ws.sheet_view.rightToLeft = True
    heads = ("التاريخ", "البيان", f"مبلغ {us}", f"مبلغ {them}", "ملاحظة")
    for col, width in zip("ABCDE", (12, 34, 15, 15, 44), strict=True):
        ws.column_dimensions[col].width = width

    def wide(text: str, size: int = 11) -> None:
        """A line of text across the whole table (never cut by a narrow column)."""
        ws.append([text])
        ws.merge_cells(start_row=ws.max_row, start_column=1, end_row=ws.max_row, end_column=5)
        ws.cell(ws.max_row, 1).font = Font(bold=size > 11, size=size)
        ws.cell(ws.max_row, 1).alignment = Alignment(horizontal="right", readingOrder=2)

    wide(f"مقارنة الحساب بين {us} و{them}", 14)
    if result.period:
        wide(f"الفترة: {result.period[0]:%d/%m/%Y} — {result.period[1]:%d/%m/%Y}")
    wide("الأحمر: ناقصة عند أحد الطرفين — الأصفر: مبلغ مختلف")
    ws.append([])
    ws.append(list(heads))
    for cell in ws[ws.max_row]:
        cell.font, cell.fill, cell.alignment, cell.border = Font(bold=True), HEAD, CELL, BOX
    ws.freeze_panes = f"A{ws.max_row + 1}"
    for r in rows(result):
        o, t = r.ours, r.theirs  # the date and description of whoever has it (ours first)
        ws.append([o.day or t.day, o.label or t.label, o.amount, t.amount, r.note])
        line = ws[ws.max_row]
        for cell in line:
            cell.alignment, cell.border = CELL, BOX
        for cell in line:  # the whole line in its colour: seen at a glance
            cell.fill = r.fill
        for cell, part in ((line[2], o), (line[3], t)):
            cell.number_format = MONEY
            if r.fill is RED and not part.lines:  # the missing amount: darker, marked
                cell.value, cell.fill = "ناقصة", MISSING
                cell.font = Font(bold=True, color="9C0006")
    gaps = [r for r in rows(result) if r.fill is YELLOW]
    if gaps:  # the one total that means something: the amount gaps
        total = sum(
            abs(sum(x.amount for x in r.ours.lines) - sum(x.amount for x in r.theirs.lines))
            for r in gaps
        )
        total = units(total)
        ws.append([])
        wide(f"مجموع الفروق في الأسطر الصفراء ({len(gaps)} عملية): {total:,}", 12)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save(result: Result, target: Path) -> Path:
    target.write_bytes(build(result))
    return target


__all__ = ["Row", "build", "rows", "save"]
