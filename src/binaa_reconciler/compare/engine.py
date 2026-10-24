"""Compare two RIM-Soft ledgers of one account between two parties.

Each party records the other: what one writes on the debit side the other
writes on the credit side (a mirror). The comparison is plain and predictable:

1. Match: the same amount on mirrored sides, dated at most 3 days apart
   (same day first, then 1, 2, 3 days). Each line is used once.
2. Possible pairs among what is left, for the user to confirm («مطابقة») or
   reject («ليست هي»): a movement and its reversal in the same file, the same
   amount up to 15 days apart, the same amount on the same side (an entry on
   the wrong side?), two lines of one day making one line. Close amounts (at
   most 2 %, a price difference?) are a separate list. A confirmed pair is
   moved to the matches; a rejected one frees its lines.
3. What has no counterpart at all is listed per side, for the chosen period
   (by default the days both files cover); lines outside it are only counted.

Amounts are shown in whole units. Pure functions: no file, no screen.
"""

from __future__ import annotations

from collections.abc import Callable, Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date
from enum import StrEnum

from binaa_reconciler.core.models import LedgerSide, Transaction

MATCH_DAYS = 3
SAME_AMOUNT_DAYS = 15
NEAR_DAYS = 7
SUM_DAYS = 2  # the single line may be recorded up to 2 days after/before
MAX_SUM_PARTS = 8  # at most 8 lines of one day make one line of the other file
MAX_SUM_LINES = 40  # lines of one day tried
MAX_SUM_STATES = 50_000
CLOSE_PERCENT = 2

DEBIT, CREDIT = LedgerSide.DEBIT, LedgerSide.CREDIT
SIDE_AR = {DEBIT: "مدين", CREDIT: "دائن"}


class Kind(StrEnum):
    REVERSAL = "REVERSAL"  # a movement and its reversal in the same file
    SAME = "SAME"  # the same amount, mirrored sides, some days apart
    SIDE = "SIDE"  # the same amount on the same side in both files
    SUM = "SUM"  # two lines of one day making one line of the other file
    CLOSE = "CLOSE"  # a close amount (price difference?)


@dataclass(frozen=True)
class Line:
    """One ledger line, as read from the file."""

    key: str  # unique within its file (e.g. "A12")
    row: int  # row number in Excel
    day: date
    side: LedgerSide
    amount: int  # minor units (1/100)
    label: str = ""
    ref: str = ""

    @staticmethod
    def of(tx: Transaction, prefix: str) -> Line:
        if tx.ledger_side is None:
            raise ValueError("not_a_ledger_line")
        return Line(
            key=f"{prefix}{tx.row_ref.row}",
            row=tx.row_ref.row,
            day=tx.tx_date,
            side=tx.ledger_side,
            amount=tx.amount_minor,
            label=" ".join((tx.text_raw or "").split()),
            ref=tx.piece_number or tx.reference or "",
        )


@dataclass(frozen=True)
class Pair:
    """A possible pair: lines of the first file (a) and of the second (b)."""

    kind: Kind
    a: tuple[Line, ...]
    b: tuple[Line, ...]

    @property
    def key(self) -> str:
        return "|".join(sorted(x.key for x in (*self.a, *self.b)))

    @property
    def lines(self) -> tuple[Line, ...]:
        return (*self.a, *self.b)

    @property
    def day(self) -> date:
        return min(x.day for x in self.lines)

    @property
    def difference(self) -> int:
        """Amount of the first file minus the second (signed, minor units)."""
        return sum(x.amount for x in self.a) - sum(x.amount for x in self.b)


@dataclass
class Result:
    names: tuple[str, str]
    period: tuple[date, date] | None
    matched: list[tuple[Line, Line]] = field(default_factory=list)  # automatic
    confirmed: list[Pair] = field(default_factory=list)  # confirmed by the user
    possible: list[Pair] = field(default_factory=list)  # to confirm or reject
    close: list[Pair] = field(default_factory=list)  # close amounts, to check
    only: tuple[list[Line], list[Line]] = field(default_factory=lambda: ([], []))
    notes: dict[str, str] = field(default_factory=dict)  # line key -> note
    outside: tuple[int, int] = (0, 0)  # lines outside the period (not listed)

    def total(self, i: int) -> int:
        return sum(x.amount for x in self.only[i])

    @property
    def matched_count(self) -> int:
        return len(self.matched) + len(self.confirmed)


def mirror(side: LedgerSide) -> LedgerSide:
    return CREDIT if side is DEBIT else DEBIT


def units(minor: int) -> int:
    """Whole units (no decimals), rounded half up."""
    return (abs(minor) + 50) // 100 * (1 if minor >= 0 else -1)


def amount_text(minor: int) -> str:
    """4,140,000: whole units, one block of digits (safe in right-to-left text)."""
    return f"{units(minor):,}"


def _gap(a: Line, b: Line) -> int:
    return abs((a.day - b.day).days)


def _pair(
    left: Sequence[Line],
    right: Sequence[Line],
    fits: Callable[[Line, Line], bool],
    rank: Callable[[Line, Line], tuple[object, ...]],
    used: set[str],
) -> list[tuple[Line, Line]]:
    """Greedy, deterministic pairing: the best-ranked fitting pairs first."""
    options = [
        (rank(a, b), a.key, b.key, a, b)
        for a in left
        if a.key not in used
        for b in right
        if b.key not in used and fits(a, b)
    ]
    out = []
    for *_, a, b in sorted(options, key=lambda o: o[:3]):
        if a.key not in used and b.key not in used:
            used.update((a.key, b.key))
            out.append((a, b))
    return out


def _by_amount(lines: Iterable[Line]) -> dict[int, list[Line]]:
    groups: dict[int, list[Line]] = {}
    for x in lines:
        groups.setdefault(x.amount, []).append(x)
    return groups


def common_period(a: Sequence[Line], b: Sequence[Line]) -> tuple[date, date] | None:
    """The days both files cover (all days of both when they do not overlap)."""
    if not a or not b:
        return None
    start = max(min(x.day for x in a), min(x.day for x in b))
    end = min(max(x.day for x in a), max(x.day for x in b))
    if start > end:
        start = min(min(x.day for x in a), min(x.day for x in b))
        end = max(max(x.day for x in a), max(x.day for x in b))
    return start, end


def _key(*lines: Line) -> str:
    return "|".join(sorted(x.key for x in lines))


def _sum_of(
    parts: Sequence[Line], target: Line, ok: Callable[..., bool]
) -> tuple[Line, ...] | None:
    """2 to MAX_SUM_PARTS lines of one day whose amounts make `target` exactly
    (the nearest day first), or None. Bounded: a busy day never stalls."""
    by_day: dict[date, list[Line]] = {}
    for x in parts:
        by_day.setdefault(x.day, []).append(x)
    for day in sorted(by_day, key=lambda d: (abs((d - target.day).days), d)):
        lines = sorted(by_day[day], key=lambda x: (-x.amount, x.key))[:MAX_SUM_LINES]
        reach: dict[int, tuple[Line, ...]] = {0: ()}
        for x in lines:
            for total, combo in list(reach.items()):
                new = total + x.amount
                if new > target.amount or len(combo) >= MAX_SUM_PARTS or new in reach:
                    continue
                reach[new] = (*combo, x)
                if len(reach) > MAX_SUM_STATES:
                    break
        combo = reach.get(target.amount, ())
        if len(combo) >= 2 and ok(*combo, target):
            return tuple(sorted(combo, key=lambda x: (x.day, x.key)))
    return None


def _possible(
    a: Sequence[Line], b: Sequence[Line], refused: set[str], used: set[str]
) -> list[Pair]:
    """Possible pairs among free lines, the most likely kind first."""
    out: list[Pair] = []

    def free(lines: Sequence[Line]) -> list[Line]:
        return [x for x in lines if x.key not in used]

    def ok(*lines: Line) -> bool:
        return _key(*lines) not in refused

    for lines, first in ((a, True), (b, False)):  # a movement and its reversal
        for x, y in _pair(
            free(lines),
            free(lines),
            lambda p, q: (
                p.key < q.key
                and p.amount == q.amount
                and p.side is not q.side
                and _gap(p, q) <= NEAR_DAYS
                and ok(p, q)
            ),
            lambda p, q: (_gap(p, q), p.day),
            used,
        ):
            out.append(Pair(Kind.REVERSAL, (x, y) if first else (), () if first else (x, y)))
    for first in (True, False):  # several lines of one day making one line
        targets, parts_pool = (free(b), free(a)) if first else (free(a), free(b))
        for t in sorted(targets, key=lambda x: (x.day, x.amount, x.key)):
            if t.key in used:
                continue
            found = _sum_of(
                [
                    x
                    for x in parts_pool
                    if x.key not in used
                    and _gap(x, t) <= SUM_DAYS
                    and x.side is mirror(t.side)
                    and x.amount < t.amount
                ],
                t,
                ok,
            )
            if found is None:
                continue
            used.update(x.key for x in (*found, t))
            pair = Pair(Kind.SUM, found, (t,)) if first else Pair(Kind.SUM, (t,), found)
            out.append(pair)
    for x, y in _pair(  # the same amount, mirrored sides, up to 15 days apart
        free(a),
        free(b),
        lambda p, q: (
            p.amount == q.amount
            and q.side is mirror(p.side)
            and _gap(p, q) <= SAME_AMOUNT_DAYS
            and ok(p, q)
        ),
        lambda p, q: (_gap(p, q), p.day),
        used,
    ):
        out.append(Pair(Kind.SAME, (x,), (y,)))
    for x, y in _pair(  # the same amount on the same side: an entry error?
        free(a),
        free(b),
        lambda p, q: (
            p.amount == q.amount and q.side is p.side and _gap(p, q) <= NEAR_DAYS and ok(p, q)
        ),
        lambda p, q: (_gap(p, q), p.day),
        used,
    ):
        out.append(Pair(Kind.SIDE, (x,), (y,)))
    for x, y in _pair(  # a close amount, mirrored sides, up to 7 days apart
        free(a),
        free(b),
        lambda p, q: (
            q.side is mirror(p.side)
            and _gap(p, q) <= NEAR_DAYS
            and 0 < abs(p.amount - q.amount) * 100 <= CLOSE_PERCENT * min(p.amount, q.amount)
            and ok(p, q)
        ),
        lambda p, q: (abs(p.amount - q.amount), _gap(p, q), p.day),
        used,
    ):
        out.append(Pair(Kind.CLOSE, (x,), (y,)))
    return out


def compare(
    a: Sequence[Line],
    b: Sequence[Line],
    names: tuple[str, str],
    period: tuple[date, date] | None = None,
    confirmed: Sequence[Pair] = (),
    rejected: Iterable[str] = (),
) -> Result:
    """`confirmed`: pairs the user confirmed; `rejected`: keys of pairs the user
    rejected (both kept by the window for the session only). Matching uses
    every line; the lists show what is dated within `period` (default: the
    common period)."""
    period = period or common_period(a, b)
    result = Result(names=names, period=period)
    used: set[str] = set()

    # 1. automatic matches: same amount, mirrored sides, at most 3 days apart
    b_by_amount = _by_amount(b)
    for amount, a_lines in sorted(_by_amount(a).items()):
        result.matched += _pair(
            a_lines,
            b_by_amount.get(amount, []),
            lambda x, y: y.side is mirror(x.side) and _gap(x, y) <= MATCH_DAYS,
            lambda x, y: (_gap(x, y), x.day),
            used,
        )
    result.matched.sort(key=lambda p: (p[0].day, p[0].amount))

    def inside(day: date) -> bool:
        return period is not None and period[0] <= day <= period[1]

    # 2. possible pairs: confirmed ones first become matches; rejected ones are
    # never proposed again (their lines may then pair otherwise or stay alone)
    refused = set(rejected)
    for pair in confirmed:
        if not any(x.key in used for x in pair.lines):
            result.confirmed.append(pair)
            used.update(x.key for x in pair.lines)
    pairs = _possible(a, b, refused, used)
    for pair in pairs:
        if not any(inside(x.day) for x in pair.lines):
            continue
        used.update(x.key for x in pair.lines)
        (result.close if pair.kind is Kind.CLOSE else result.possible).append(pair)

    # 3. alone: no counterpart at all
    left = [x for x in a if x.key not in used]
    right = [x for x in b if x.key not in used]
    result.only = (
        sorted((x for x in left if inside(x.day)), key=lambda x: (x.day, x.amount)),
        sorted((x for x in right if inside(x.day)), key=lambda x: (x.day, x.amount)),
    )
    result.outside = (
        sum(not inside(x.day) for x in left),
        sum(not inside(x.day) for x in right),
    )
    for i, lines in enumerate(result.only):  # the same amount alone on the other side
        others = _by_amount(result.only[1 - i])
        for x in lines:
            if same := others.get(x.amount):
                days = "، ".join(f"{y.day:%d/%m/%Y}" for y in same[:3])
                result.notes[x.key] = f"⚠️ نفس المبلغ عند {names[1 - i]} بتاريخ {days}"
    for i, lines in enumerate(result.only):
        other = b if i == 0 else a
        if not other:
            continue
        first, last = min(x.day for x in other), max(x.day for x in other)
        for x in lines:
            edge = ""
            if x.day < first:
                edge = f"قبل أول عملية في ملف {names[1 - i]} ({first:%d/%m/%Y})"
            elif x.day > last:
                edge = f"بعد آخر عملية في ملف {names[1 - i]} ({last:%d/%m/%Y})"
            if edge:
                result.notes[x.key] = " — ".join(filter(None, (edge, result.notes.get(x.key))))
    return result


__all__ = [
    "Kind",
    "Line",
    "Pair",
    "Result",
    "amount_text",
    "common_period",
    "compare",
    "mirror",
    "units",
]
