"""Structured import report (UI-ready, free of personal data).

`RowRecord` accounts for every sheet row: nothing disappears without a
row id, a class and a reason. `ImportReport.summary()` returns only counts,
codes, row numbers and column letters, so it can be printed or displayed
without exposing labels, names, phone numbers or amounts.
"""

from __future__ import annotations

from collections import Counter
from collections.abc import Mapping
from dataclasses import dataclass, field

from binaa_reconciler.core.dates import DateOrder
from binaa_reconciler.core.ledger_columns import ColumnRole
from binaa_reconciler.core.models import IssueCode, RowClass, Severity, ValidationIssue


def column_letter(index: int) -> str:
    """0 -> A, 25 -> Z, 26 -> AA."""
    letters = ""
    n = index + 1
    while n:
        n, rem = divmod(n - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


@dataclass(frozen=True)
class RowRecord:
    row_id: str  # "R<sheet row number>"
    row_number: int
    row_class: RowClass
    reason: str
    issues: tuple[IssueCode, ...] = ()
    transaction_id: str | None = None


@dataclass(frozen=True)
class ImportReport:
    file_sha256: str
    sheet_index: int | None
    sheet_count: int
    header_row: int | None
    columns: Mapping[ColumnRole, int]
    date_order: DateOrder | None
    rows: tuple[RowRecord, ...]
    issues: tuple[ValidationIssue, ...]
    sheet_name: str | None = field(default=None, repr=False)

    # ------------------------------------------------------------ counts
    @property
    def total_rows(self) -> int:
        return len(self.rows)

    @property
    def class_counts(self) -> dict[RowClass, int]:
        counts = Counter(r.row_class for r in self.rows)
        return {cls: counts.get(cls, 0) for cls in RowClass}

    def count(self, row_class: RowClass) -> int:
        return sum(1 for r in self.rows if r.row_class is row_class)

    @property
    def transactions(self) -> int:
        return self.count(RowClass.TRANSACTION)

    @property
    def headers(self) -> int:
        return self.count(RowClass.HEADER) + self.count(RowClass.HEADER_REPEAT)

    @property
    def totals(self) -> int:
        return self.count(RowClass.TOTAL)

    @property
    def opening_balances(self) -> int:
        return self.count(RowClass.OPENING_BALANCE)

    @property
    def closing_balances(self) -> int:
        return self.count(RowClass.CLOSING_BALANCE)

    @property
    def empty(self) -> int:
        return self.count(RowClass.EMPTY)

    @property
    def invalid(self) -> int:
        return self.count(RowClass.INVALID)

    def _severity(self, severity: Severity) -> int:
        return sum(1 for i in self.issues if i.severity is severity)

    @property
    def warnings(self) -> int:
        return self._severity(Severity.WARNING)

    @property
    def blocking_errors(self) -> int:
        return self._severity(Severity.BLOCKING)

    @property
    def infos(self) -> int:
        return self._severity(Severity.INFO)

    @property
    def is_blocked(self) -> bool:
        return self.blocking_errors > 0

    # ------------------------------------------------------------ export
    def summary(self) -> dict[str, object]:
        """Personal-data-free summary for CLI / UI."""
        grouped: dict[tuple[str, str], list[ValidationIssue]] = {}
        for issue in self.issues:
            grouped.setdefault((issue.code.value, issue.severity.value), []).append(issue)
        by_code = [
            {
                "code": code,
                "severity": severity,
                "count": len(items),
                "rows": [i.row_id for i in items if i.row_id is not None],
            }
            for (code, severity), items in grouped.items()
        ]
        return {
            "file_sha256": self.file_sha256,
            "sheet_index": self.sheet_index,
            "sheet_count": self.sheet_count,
            "header_row": self.header_row,
            "columns": {role.value: column_letter(i) for role, i in self.columns.items()},
            "date_order": self.date_order.value if self.date_order else None,
            "total_rows": self.total_rows,
            "transactions": self.transactions,
            "headers": self.headers,
            "totals": self.totals,
            "opening_balances": self.opening_balances,
            "closing_balances": self.closing_balances,
            "empty": self.empty,
            "invalid": self.invalid,
            "preamble": self.count(RowClass.PREAMBLE),
            "warnings": self.warnings,
            "blocking_errors": self.blocking_errors,
            "infos": self.infos,
            "blocked": self.is_blocked,
            "issues": by_code,
        }
