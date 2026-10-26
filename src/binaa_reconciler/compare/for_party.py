"""A short file to send to the other party: only what concerns them.

One right-to-left sheet, written from their point of view («عليكم» / «لكم»):
what we recorded and they did not, the amount gaps, and what they recorded
that we could not find. Nothing else of our books; written only where the
user chooses.
"""

from __future__ import annotations

from io import BytesIO
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from binaa_reconciler.compare.engine import Line, Result, units
from binaa_reconciler.compare.manual import arabic
from binaa_reconciler.core.models import LedgerSide

HEAD = PatternFill("solid", fgColor="D9E1F2")
RIGHT = Alignment(horizontal="right", vertical="top", wrap_text=True, readingOrder=2)
MONEY = "#,##0"


def _for_them(x: Line, ours: bool) -> str:
    """«عليكم» / «لكم» seen by the other party. In our books a debit is owed by
    them; in theirs a debit is owed by us."""
    owed_by_them = (x.side is LedgerSide.DEBIT) == ours
    return "عليكم" if owed_by_them else "لكم"


def build(result: Result) -> bytes:
    us, them = (arabic(n) for n in result.names)
    wb = Workbook()
    ws = wb.active if wb.active is not None else wb.create_sheet()
    ws.title = them[:31]
    ws.sheet_view.rightToLeft = True
    widths = {"A": 13, "B": 34, "C": 15, "D": 9, "E": 30}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.append([f"مقارنة الحساب بين {us} و{them}"])
    ws["A1"].font = Font(bold=True, size=14)
    if result.period:
        ws.append([f"الفترة: {result.period[0]:%d/%m/%Y} — {result.period[1]:%d/%m/%Y}"])

    def section(title: str, rows: list[list[object]]) -> None:
        if not rows:
            return
        ws.append([])
        ws.append([f"{title} ({len(rows)})"])
        ws.cell(ws.max_row, 1).font = Font(bold=True, size=12)
        ws.append(["التاريخ", "البيان", "المبلغ", "", "ملاحظة"])
        for cell in ws[ws.max_row]:
            cell.font, cell.fill = Font(bold=True), HEAD
        for row in rows:
            ws.append(row)
            ws.cell(ws.max_row, 1).number_format = "DD/MM/YYYY"
            ws.cell(ws.max_row, 3).number_format = MONEY
            for cell in ws[ws.max_row]:
                cell.alignment = RIGHT

    section(
        "سجّلناها عندنا ولم نجدها عندكم — هل سجلتموها؟",
        [[x.day, x.label, units(x.amount), _for_them(x, ours=True), ""] for x in result.only[0]],
    )
    section(
        "مسجّلة عند الطرفين بمبلغ مختلف",
        [
            [
                pair.b[0].day if pair.b else pair.a[0].day,
                (pair.b[0].label if pair.b else "") or (pair.a[0].label if pair.a else ""),
                units(sum(x.amount for x in pair.b)),
                _for_them(pair.b[0], ours=False) if pair.b else "",
                f"عندنا {sum(units(x.amount) for x in pair.a):,}",
            ]
            for pair in result.close
        ],
    )
    section(
        "سجّلتموها عندكم ولم نجدها عندنا",
        [[x.day, x.label, units(x.amount), _for_them(x, ours=False), ""] for x in result.only[1]],
    )
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def save(result: Result, target: Path) -> Path:
    target.write_bytes(build(result))
    return target


__all__ = ["build", "save"]
