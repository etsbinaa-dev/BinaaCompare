"""Read one RIM-Soft file into comparison lines (locally, never sent anywhere)."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from binaa_reconciler.app.texts import ISSUES_AR
from binaa_reconciler.compare.engine import Line
from binaa_reconciler.core.ledger_columns import ColumnRole
from binaa_reconciler.core.models import Severity
from binaa_reconciler.importers.rimsoft_xlsx import import_rimsoft_xlsx


@dataclass(frozen=True)
class Loaded:
    lines: tuple[Line, ...]
    problems: tuple[str, ...]  # rows that could not be read ("R93: سطر بلا مبلغ")
    merged: int  # operations cut over two rows and rejoined
    error: str | None = None  # the file could not be read at all
    # to rebuild the original sheet for the manual comparison (kept in memory)
    data: bytes = field(default=b"", repr=False)
    sheet_index: int = 0
    header_row: int = 1
    continuation: dict[int, int] = field(default_factory=dict)  # row -> its 2nd row
    date_col: int | None = None  # 0-based column of the dates


def load(path: Path | bytes, prefix: str) -> Loaded:
    try:
        data = path if isinstance(path, bytes) else Path(path).read_bytes()
        result = import_rimsoft_xlsx(data)
    except Exception:  # an unreadable file is a message, never a crash
        return Loaded((), (), 0, "تعذّرت قراءة الملف — تأكد أنه ملف Excel من RIM-Soft")
    lines = tuple(Line.of(t, prefix) for t in result.transactions if t.ledger_side is not None)
    problems = tuple(
        f"السطر {i.row_ref.row}: {ISSUES_AR.get(i.code, i.code.value)}"
        for i in result.issues
        if i.severity is Severity.BLOCKING and i.row_ref is not None
    )
    merged = sum(1 for i in result.issues if i.code.value == "SPLIT_ROW_MERGED")
    if not lines:
        return Loaded((), problems, merged, "لم يُعثر في الملف على أي عملية")
    report = result.report
    continuation = {
        i.row_ref.row - 1: i.row_ref.row
        for i in result.issues
        if i.code.value == "SPLIT_ROW_MERGED" and i.row_ref is not None
    }
    return Loaded(
        lines,
        problems,
        merged,
        data=data,
        sheet_index=report.sheet_index or 0,
        header_row=report.header_row or 1,
        date_col=report.columns.get(ColumnRole.DATE),
        continuation=continuation,
    )


__all__ = ["Loaded", "load"]
