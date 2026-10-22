"""Core domain models.

Privacy rule: fields that may contain personal data (names, phones, labels,
free text, notes) are excluded from `repr`, so an accidental log of a model
never prints personal data.

Two concepts are deliberately separate:
* `Tier` is the engine's *suggestion* strength (HIGH / REVIEW).
* `TransactionStatus` is the *decided* state. A transaction becomes MATCHED
  only through a user decision, never through a suggestion.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from datetime import date, datetime, time
from enum import StrEnum


class Source(StrEnum):
    BANKILY = "BANKILY"
    RIMSOFT = "RIMSOFT"


class Direction(StrEnum):
    """Direction of money relative to the Bankily wallet."""

    IN = "IN"
    OUT = "OUT"


class LedgerSide(StrEnum):
    """Accounting side of a RIM-Soft ledger line."""

    DEBIT = "DEBIT"
    CREDIT = "CREDIT"


class TransactionStatus(StrEnum):
    MATCHED = "matched"
    PENDING = "pending"
    UNMATCHED = "unmatched"
    EXPLAINED = "explained"


class RowClass(StrEnum):
    TRANSACTION = "TRANSACTION"
    HEADER = "HEADER"  # the detected column header row
    PREAMBLE = "PREAMBLE"  # non-empty rows above the header (titles, period...)
    HEADER_REPEAT = "HEADER_REPEAT"
    TOTAL = "TOTAL"
    OPENING_BALANCE = "OPENING_BALANCE"
    CLOSING_BALANCE = "CLOSING_BALANCE"
    EMPTY = "EMPTY"
    INVALID = "INVALID"


class Severity(StrEnum):
    BLOCKING = "BLOCKING"
    WARNING = "WARNING"
    INFO = "INFO"


class IssueCode(StrEnum):
    BAD_DATE = "BAD_DATE"
    AMBIGUOUS_DATE_ORDER = "AMBIGUOUS_DATE_ORDER"
    BAD_AMOUNT = "BAD_AMOUNT"
    AMBIGUOUS_AMOUNT = "AMBIGUOUS_AMOUNT"
    NO_AMOUNT = "NO_AMOUNT"
    NEGATIVE_AMOUNT = "NEGATIVE_AMOUNT"
    BOTH_SIDES = "BOTH_SIDES"
    BALANCE_CHAIN_BREAK = "BALANCE_CHAIN_BREAK"
    DUPLICATE_ROW = "DUPLICATE_ROW"
    NON_TRANSACTION_ROW = "NON_TRANSACTION_ROW"
    # File / sheet / header level (P1)
    FILE_UNREADABLE = "FILE_UNREADABLE"
    SHEET_NOT_FOUND = "SHEET_NOT_FOUND"
    MULTIPLE_CANDIDATE_SHEETS = "MULTIPLE_CANDIDATE_SHEETS"
    HEADER_NOT_FOUND = "HEADER_NOT_FOUND"
    AMBIGUOUS_HEADER = "AMBIGUOUS_HEADER"
    INVALID_COLUMN_MAPPING = "INVALID_COLUMN_MAPPING"
    NO_TRANSACTIONS = "NO_TRANSACTIONS"
    # Row level (P1)
    LABEL_RESEMBLES_SUMMARY = "LABEL_RESEMBLES_SUMMARY"
    SPLIT_ROW_MERGED = "SPLIT_ROW_MERGED"  # one operation cut over two rows, rejoined
    CONTENT_OUTSIDE_COLUMNS = "CONTENT_OUTSIDE_COLUMNS"
    FORMULA_WITHOUT_CACHED_VALUE = "FORMULA_WITHOUT_CACHED_VALUE"
    FORMULA_ERROR_VALUE = "FORMULA_ERROR_VALUE"
    # Bankily PDF structure (P2A)
    NO_TEXT_LAYER = "NO_TEXT_LAYER"
    EMPTY_PAGE = "EMPTY_PAGE"
    PAGE_WITHOUT_HEADER = "PAGE_WITHOUT_HEADER"
    ORPHAN_LINE = "ORPHAN_LINE"
    UNALIGNED_TEXT = "UNALIGNED_TEXT"
    OPENING_BALANCE_MISMATCH = "OPENING_BALANCE_MISMATCH"
    # Bankily balances (P2A)
    BAD_BALANCE = "BAD_BALANCE"
    BALANCE_MISSING = "BALANCE_MISSING"
    BALANCE_CELL_MISMATCH = "BALANCE_CELL_MISMATCH"
    NO_OPENING_BALANCE = "NO_OPENING_BALANCE"
    CLOSING_BALANCE_MISMATCH = "CLOSING_BALANCE_MISMATCH"
    BAD_TIME = "BAD_TIME"
    # Reconciliation (P2B)
    INVALID_SCOPE = "INVALID_SCOPE"
    NO_START_DATE = "NO_START_DATE"
    MISSING_ACCOUNT_PROFILE = "MISSING_ACCOUNT_PROFILE"
    DUPLICATE_TRANSACTION_ID = "DUPLICATE_TRANSACTION_ID"
    COVERAGE_MISMATCH = "COVERAGE_MISMATCH"
    LARGE_AMBIGUOUS_CLUSTER = "LARGE_AMBIGUOUS_CLUSTER"
    GROUP_AMBIGUOUS = "GROUP_AMBIGUOUS"
    GROUP_SEARCH_LIMIT = "GROUP_SEARCH_LIMIT"


class CandidateKind(StrEnum):
    """Cardinality, always written Bankily:RIM-Soft."""

    ONE_TO_ONE = "1:1"
    MANY_TO_ONE = "N:1"  # several Bankily lines -> one RIM-Soft line
    ONE_TO_MANY = "1:N"  # one Bankily line -> several RIM-Soft lines


class Tier(StrEnum):
    HIGH = "HIGH"
    REVIEW = "REVIEW"


class Polarity(StrEnum):
    POSITIVE = "+"
    NEGATIVE = "-"
    NEUTRAL = "0"


class DecisionAction(StrEnum):
    CONFIRM = "CONFIRM"
    REJECT = "REJECT"
    MANUAL_MATCH = "MANUAL_MATCH"
    EXPLAIN = "EXPLAIN"
    UNDO = "UNDO"
    SAVE_ALIAS = "SAVE_ALIAS"


class AliasKeyType(StrEnum):
    PHONE = "PHONE"
    NAME = "NAME"


def kind_for(n_bankily: int, n_rimsoft: int) -> CandidateKind:
    """Cardinality for a group. M:N is not supported and raises."""
    if n_bankily < 1 or n_rimsoft < 1:
        raise ValueError("empty_side")
    if n_bankily == 1 and n_rimsoft == 1:
        return CandidateKind.ONE_TO_ONE
    if n_rimsoft == 1:
        return CandidateKind.MANY_TO_ONE
    if n_bankily == 1:
        return CandidateKind.ONE_TO_MANY
    raise ValueError("many_to_many_not_supported")


@dataclass(frozen=True)
class RowRef:
    """Location of a row in its source file (for the user, not for matching)."""

    row: int
    page: int | None = None


@dataclass(frozen=True)
class ImportBatch:
    id: str
    source: Source
    file_name: str = field(repr=False)
    file_sha256: str
    imported_at: datetime
    profile_name: str | None = None
    period_start: date | None = None
    period_end: date | None = None
    row_counts: Mapping[RowClass, int] = field(default_factory=dict)


@dataclass(frozen=True)
class Transaction:
    """A normalized transaction from either source.

    Bankily lines carry `direction`; RIM-Soft lines carry `ledger_side`.
    The mapping between them is the job of `AccountProfile`.
    """

    id: str
    batch_id: str
    source: Source
    fingerprint: str
    row_ref: RowRef
    tx_date: date
    amount_minor: int
    tx_time: time | None = None
    direction: Direction | None = None
    ledger_side: LedgerSide | None = None
    reference: str | None = field(default=None, repr=False)
    piece_number: str | None = field(default=None, repr=False)
    ecriture_number: str | None = field(default=None, repr=False)
    journal: str | None = field(default=None, repr=False)
    phone_norm: str | None = field(default=None, repr=False)
    name_raw: str | None = field(default=None, repr=False)
    name_key: str | None = field(default=None, repr=False)
    text_raw: str = field(default="", repr=False)
    text_norm: str = field(default="", repr=False)
    balance_minor: int | None = None
    status: TransactionStatus = TransactionStatus.PENDING
    # Set by `core.currency.normalize_transactions`: the amount as written in the
    # file and its unit ("MRU" / "MRO"); `amount_minor` is then the internal value.
    source_amount_minor: int | None = None
    source_currency: str | None = None

    def __post_init__(self) -> None:
        if self.amount_minor <= 0:
            raise ValueError("amount_must_be_positive")
        if self.source is Source.BANKILY and (
            self.direction is None or self.ledger_side is not None
        ):
            raise ValueError("bankily_requires_direction_only")
        if self.source is Source.RIMSOFT and (
            self.ledger_side is None or self.direction is not None
        ):
            raise ValueError("rimsoft_requires_ledger_side_only")


@dataclass(frozen=True)
class ValidationIssue:
    """A problem found before matching. `message` must never contain data values."""

    batch_id: str
    source: Source
    row_ref: RowRef | None
    severity: Severity
    code: IssueCode
    row_id: str | None = None
    message: str = ""
    group: str | None = None  # e.g. duplicate group id

    @property
    def is_blocking(self) -> bool:
        return self.severity is Severity.BLOCKING


@dataclass(frozen=True)
class Evidence:
    """One human-readable reason contributing to a candidate's score."""

    code: str
    points: int
    polarity: Polarity
    label: str = field(default="", repr=False)


@dataclass(frozen=True)
class Candidate:
    """A *suggestion* produced by the engine. It never changes any status."""

    bankily_ids: tuple[str, ...]
    rimsoft_ids: tuple[str, ...]
    kind: CandidateKind
    score: int
    tier: Tier
    evidence: tuple[Evidence, ...] = ()
    margin_to_next: int | None = None

    def __post_init__(self) -> None:
        if kind_for(len(self.bankily_ids), len(self.rimsoft_ids)) is not self.kind:
            raise ValueError("kind_does_not_match_ids")
        if self.kind is not CandidateKind.ONE_TO_ONE and self.tier is Tier.HIGH:
            # Group suggestions always require review; never high confidence.
            raise ValueError("group_candidate_cannot_be_high")


@dataclass(frozen=True)
class MatchGroup:
    """A match confirmed by the user (1:1, N:1 or 1:N)."""

    id: str
    bankily_ids: tuple[str, ...]
    rimsoft_ids: tuple[str, ...]
    kind: CandidateKind
    confirmed_at: datetime
    confirmed_by: str
    decision_id: str
    note: str = field(default="", repr=False)

    def __post_init__(self) -> None:
        if kind_for(len(self.bankily_ids), len(self.rimsoft_ids)) is not self.kind:
            raise ValueError("kind_does_not_match_ids")
        if set(self.bankily_ids) & set(self.rimsoft_ids):
            raise ValueError("ids_overlap")


@dataclass(frozen=True)
class Decision:
    """One user decision. Decisions are append-only (see `DecisionLog`)."""

    id: str
    ts: datetime
    user: str
    action: DecisionAction
    bankily_ids: tuple[str, ...] = ()
    rimsoft_ids: tuple[str, ...] = ()
    score_snapshot: int | None = None
    evidence_snapshot: tuple[Evidence, ...] = ()
    undoes: str | None = None
    note: str = field(default="", repr=False)

    def __post_init__(self) -> None:
        if (self.action is DecisionAction.UNDO) != (self.undoes is not None):
            raise ValueError("undo_requires_target")
        if self.ts.tzinfo is None:
            raise ValueError("decision_timestamp_must_be_aware")


@dataclass(frozen=True)
class RejectedPair:
    """A pair the user rejected; the engine must not suggest it again."""

    bankily_fingerprint: str
    rimsoft_fingerprint: str
    rejected_at: datetime
    decision_id: str


@dataclass(frozen=True)
class Alias:
    """User-confirmed link: a Bankily sender key -> a RIM-Soft label/customer."""

    key_type: AliasKeyType
    key_norm: str = field(repr=False)
    rimsoft_label_norm: str = field(repr=False)
    confirmations: int = 1
    last_used: datetime | None = None

    def __post_init__(self) -> None:
        if self.confirmations < 1:
            raise ValueError("confirmations_must_be_positive")
        if not self.key_norm or not self.rimsoft_label_norm:
            raise ValueError("alias_keys_required")


class DecisionLog:
    """In-memory append-only decision journal (persistence comes in a later phase)."""

    def __init__(self) -> None:
        self._entries: list[Decision] = []
        self._ids: set[str] = set()

    def append(self, decision: Decision) -> None:
        if decision.id in self._ids:
            raise ValueError("duplicate_decision_id")
        if decision.undoes is not None and decision.undoes not in self._ids:
            raise ValueError("undo_target_unknown")
        self._entries.append(decision)
        self._ids.add(decision.id)

    @property
    def entries(self) -> tuple[Decision, ...]:
        return tuple(self._entries)

    def active(self) -> tuple[Decision, ...]:
        """Decisions not cancelled by a later UNDO (UNDO entries excluded)."""
        undone = {d.undoes for d in self._entries if d.undoes is not None}
        return tuple(
            d for d in self._entries if d.action is not DecisionAction.UNDO and d.id not in undone
        )
