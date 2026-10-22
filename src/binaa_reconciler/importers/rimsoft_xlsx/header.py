"""Header row detection and column mapping.

A header row is accepted only when it contains each required role
(Date, Libellé, Débit, Crédit) exactly once. Anything less is reported, never
guessed. `ColumnMapping` is also the object a future "choose columns" dialog
will build and pass back to the importer.
"""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from dataclasses import dataclass, field

from binaa_reconciler.core.ledger_columns import REQUIRED_ROLES, ColumnRole, role_for_header
from binaa_reconciler.core.models import IssueCode


@dataclass(frozen=True)
class ColumnMapping:
    """Where the header is and which column (0-based) holds each role."""

    header_row: int  # 1-based sheet row number
    columns: Mapping[ColumnRole, int]

    def __post_init__(self) -> None:
        if self.header_row < 1:
            raise ValueError("header_row_must_be_positive")
        missing = [r for r in REQUIRED_ROLES if r not in self.columns]
        if missing:
            raise ValueError("missing_required_roles")
        indexes = list(self.columns.values())
        if any(i < 0 for i in indexes) or len(set(indexes)) != len(indexes):
            raise ValueError("invalid_column_indexes")

    def column(self, role: ColumnRole) -> int | None:
        return self.columns.get(role)


@dataclass(frozen=True)
class HeaderDetection:
    """Result of scanning a sheet for its header row."""

    mapping: ColumnMapping | None
    issue: IssueCode | None = None
    best_row: int | None = None  # closest candidate when detection failed
    missing_roles: tuple[ColumnRole, ...] = ()
    duplicated_roles: tuple[ColumnRole, ...] = ()
    recognized_roles: tuple[ColumnRole, ...] = field(default=())

    @property
    def found(self) -> bool:
        return self.mapping is not None


def _roles_in_row(row: Sequence[object]) -> dict[ColumnRole, list[int]]:
    found: dict[ColumnRole, list[int]] = {}
    for col, value in enumerate(row):
        role = role_for_header(value)
        if role is not None:
            found.setdefault(role, []).append(col)
    return found


def detect_header(
    rows: Sequence[Sequence[object]],
    max_scan: int,
    required: tuple[ColumnRole, ...] = REQUIRED_ROLES,
) -> HeaderDetection:
    """Find the first row (within `max_scan`) holding every required role exactly once.

    `required` defaults to the RIM-Soft roles; other importers pass their own.
    """
    best: tuple[int, int, dict[ColumnRole, list[int]]] | None = None  # (score, row, roles)
    for index, row in enumerate(rows[:max_scan], start=1):
        roles = _roles_in_row(row)
        required_hits = sum(1 for r in required if r in roles)
        if required_hits == len(required):
            duplicated = tuple(r for r, cols in roles.items() if len(cols) > 1)
            if duplicated:
                return HeaderDetection(
                    None,
                    IssueCode.AMBIGUOUS_HEADER,
                    best_row=index,
                    duplicated_roles=duplicated,
                    recognized_roles=tuple(roles),
                )
            columns = {role: cols[0] for role, cols in roles.items()}
            return HeaderDetection(
                ColumnMapping(header_row=index, columns=columns),
                recognized_roles=tuple(roles),
            )
        if required_hits and (best is None or required_hits > best[0]):
            best = (required_hits, index, roles)
    if best is None:
        return HeaderDetection(None, IssueCode.HEADER_NOT_FOUND, missing_roles=required)
    _, row_number, roles = best
    return HeaderDetection(
        None,
        IssueCode.HEADER_NOT_FOUND,
        best_row=row_number,
        missing_roles=tuple(r for r in required if r not in roles),
        recognized_roles=tuple(roles),
    )
