"""Classification of RIM-Soft ledger rows (pure; shared by every importer).

Every row gets a class and a reason. Non-transaction rows (header repeats,
totals, opening balances, empty rows) are skipped *visibly*; invalid rows are
reported with a code. Nothing is dropped silently.

Cells are typed values as read from a spreadsheet (str, int, float, Decimal,
date, datetime, None...) or plain strings (synthetic CSV).

Balances are split into opening and closing. Explicit phrases decide
("Solde initial", "Report à nouveau" -> opening; "Solde final", "Nouveau
solde" -> closing). A dated balance ("Solde au 31/03/2026") is an opening
balance only before the first transaction row, and a closing balance
anywhere after it. Neither ever becomes a transaction.

Summary rows are detected conservatively. A label is a summary only when it
*is* a marker phrase, optionally followed by filler words, dates or numbers:
"Total", "Sous-total", "Report à nouveau", "Solde au 01/03/2026". A label that
merely *starts* with such a word ("Solde facture 12", "Total Energies
carburant") stays a transaction; for strong markers a WARNING is raised so
the user can double check.
"""

from __future__ import annotations

import re
from collections.abc import Callable
from dataclasses import dataclass
from datetime import date, datetime, time
from decimal import Decimal
from enum import StrEnum

from binaa_reconciler.core.dates import DateParseError
from binaa_reconciler.core.ledger_columns import role_for_header
from binaa_reconciler.core.models import IssueCode, LedgerSide, RowClass
from binaa_reconciler.core.money import AmountParseError, amount_from_number, parse_amount
from binaa_reconciler.normalize.text import normalize_text


class RowReason(StrEnum):
    TRANSACTION = "transaction"
    EMPTY = "empty_row"
    PAGE_MARKER = "page_marker"
    HEADER_REPEAT = "repeated_header"
    TOTAL = "total_marker"
    CLOSING_BALANCE = "closing_balance_marker"
    OPENING_BALANCE = "opening_balance_marker"
    DATED_BALANCE_OPENING = "dated_balance_before_transactions"
    DATED_BALANCE_CLOSING = "dated_balance_after_transactions"
    INVALID = "invalid_row"
    UNRELIABLE_FORMULA = "unreliable_formula"
    ZERO_AMOUNT = "zero_amount"  # 0 written in the amount columns: nothing moved


class _Marker(StrEnum):
    TOTAL = "total"
    CLOSING = "closing"
    OPENING = "opening"
    DATED = "dated"  # opening or closing depending on position


# Longest phrases first so "sous total" wins over "total".
_MARKER_PHRASES: tuple[tuple[str, _Marker], ...] = tuple(
    sorted(
        (
            *((p, _Marker.TOTAL) for p in (
                "total", "totaux", "sous total", "total general", "cumul", "cumuls",
                "a reporter",
            )),
            *((p, _Marker.CLOSING) for p in (
                "solde", "solde final", "solde de cloture", "solde cloture", "nouveau solde",
                "solde fin de periode",
            )),
            *((p, _Marker.OPENING) for p in (
                "solde initial", "solde anterieur", "solde d ouverture", "solde debut",
                "solde a nouveau", "report", "reports", "report a nouveau",
                "reports a nouveau", "a nouveau", "anouveau",
            )),
            *((p, _Marker.DATED) for p in (
                "solde au", "solde du", "solde a la date du", "solde en date du",
                "solde arrete au",
            )),
        ),
        key=lambda item: -len(item[0]),
    )
)  # fmt: skip

_FILLER_WORDS = frozenset(
    {
        "au", "du", "le", "la", "les", "de", "des", "d", "l", "a", "en", "fin", "mru", "um",
        "general", "generale", "generaux", "periode", "mois", "exercice", "page", "debiteur",
        "crediteur", "janvier", "fevrier", "mars", "avril", "mai", "juin", "juillet", "aout",
        "septembre", "octobre", "novembre", "decembre",
    }
)  # fmt: skip

# "Page 2", "Page 2 / 5", "Page 2 sur 5", "p. 2" -- only when nothing else is on the row.
_PAGE_MARKER_RE = re.compile(r"^(?:page|p)\s*\d+(?:\s*(?:sur|of)?\s*\d+)?$")

DateCellParser = Callable[[object], date]


@dataclass(frozen=True)
class LedgerRowResult:
    row_class: RowClass
    issue: IssueCode | None = None
    tx_date: date | None = None
    side: LedgerSide | None = None
    amount_minor: int | None = None
    reason: RowReason = RowReason.TRANSACTION
    warning: IssueCode | None = None


def is_blank(value: object) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def cell_text(value: object) -> str:
    """Text of a cell for labels and identifiers (numbers without spurious ".0")."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, datetime | date | time):
        return value.isoformat()
    return str(value)


def parse_amount_cell(value: object) -> int | None:
    """Minor units, or None when blank or zero. Raises AmountParseError when unreadable."""
    if is_blank(value):
        return None
    if isinstance(value, str):
        amount = parse_amount(value)
    elif isinstance(value, bool):
        raise AmountParseError("not_a_number")
    elif isinstance(value, int | float | Decimal):
        amount = amount_from_number(value)
    else:
        raise AmountParseError("not_a_number")
    return None if amount == 0 else amount


def _summary_marker(label: str) -> tuple[str, _Marker, bool] | None:
    """(phrase, kind, is_pure) when the normalized label starts with a marker phrase."""
    for phrase, kind in _MARKER_PHRASES:
        if label == phrase or label.startswith(phrase + " "):
            rest = label[len(phrase) :].split()
            pure = all(tok.isdigit() or tok in _FILLER_WORDS for tok in rest)
            return phrase, kind, pure
    return None


def classify_ledger_row(
    date_cell: object,
    label_cell: object,
    debit_cell: object,
    credit_cell: object,
    parse_date_cell: DateCellParser,
    after_transactions: bool = False,
) -> LedgerRowResult:
    """Classify one ledger row.

    `after_transactions` tells whether a transaction (or would-be transaction)
    row appeared earlier in the sheet; it only decides between opening and
    closing for a dated balance such as "Solde au 31/03/2026".
    """
    cells = (date_cell, label_cell, debit_cell, credit_cell)
    if all(is_blank(c) for c in cells):
        return LedgerRowResult(RowClass.EMPTY, reason=RowReason.EMPTY)

    if sum(1 for c in cells if role_for_header(c) is not None) >= 2:
        return LedgerRowResult(RowClass.HEADER_REPEAT, reason=RowReason.HEADER_REPEAT)

    label = normalize_text(cell_text(label_cell))
    if (
        _PAGE_MARKER_RE.match(label)
        and is_blank(date_cell)
        and is_blank(debit_cell)
        and is_blank(credit_cell)
    ):
        return LedgerRowResult(RowClass.EMPTY, reason=RowReason.PAGE_MARKER)
    warning: IssueCode | None = None
    marker = _summary_marker(label)
    if marker is not None:
        phrase, kind, pure = marker
        if kind is _Marker.TOTAL and (pure or is_blank(date_cell)):
            return LedgerRowResult(RowClass.TOTAL, reason=RowReason.TOTAL)
        if kind is _Marker.CLOSING and pure:
            return LedgerRowResult(RowClass.CLOSING_BALANCE, reason=RowReason.CLOSING_BALANCE)
        if kind is _Marker.OPENING and pure:
            return LedgerRowResult(RowClass.OPENING_BALANCE, reason=RowReason.OPENING_BALANCE)
        if kind is _Marker.DATED and pure:
            if after_transactions:
                return LedgerRowResult(
                    RowClass.CLOSING_BALANCE, reason=RowReason.DATED_BALANCE_CLOSING
                )
            return LedgerRowResult(RowClass.OPENING_BALANCE, reason=RowReason.DATED_BALANCE_OPENING)
        # Not a summary row. A bare "solde ..." ("Solde facture 12") is a common real
        # label; any other marker at the start of a transaction label is worth a check.
        if phrase != "solde":
            warning = IssueCode.LABEL_RESEMBLES_SUMMARY

    def invalid(code: IssueCode) -> LedgerRowResult:
        return LedgerRowResult(RowClass.INVALID, code, reason=RowReason.INVALID, warning=warning)

    try:
        debit = parse_amount_cell(debit_cell)
        credit = parse_amount_cell(credit_cell)
    except AmountParseError as exc:
        ambiguous = exc.code == "ambiguous_separator"
        return invalid(IssueCode.AMBIGUOUS_AMOUNT if ambiguous else IssueCode.BAD_AMOUNT)

    if debit is not None and credit is not None:
        return invalid(IssueCode.BOTH_SIDES)
    if debit is not None:
        side, amount = LedgerSide.DEBIT, debit
    elif credit is not None:
        side, amount = LedgerSide.CREDIT, credit
    elif not (is_blank(debit_cell) and is_blank(credit_cell)):
        # 0 written in the amount column(s) (RIM-Soft prints 0,00 on informative
        # lines): no money moved, so the row is skipped like an empty one
        return LedgerRowResult(RowClass.EMPTY, reason=RowReason.ZERO_AMOUNT)
    else:  # neither side carries an amount
        return invalid(IssueCode.NO_AMOUNT)
    if amount < 0:
        return invalid(IssueCode.NEGATIVE_AMOUNT)

    if is_blank(date_cell):
        return invalid(IssueCode.BAD_DATE)
    try:
        tx_date = parse_date_cell(date_cell)
    except DateParseError as exc:
        ambiguous = exc.code == "ambiguous_order"
        return invalid(IssueCode.AMBIGUOUS_DATE_ORDER if ambiguous else IssueCode.BAD_DATE)

    return LedgerRowResult(RowClass.TRANSACTION, None, tx_date, side, amount, warning=warning)
