"""Review of a day's journal entries: «Copier écritures» from the RIM-Soft
Comptabilité screen (all accounts), pasted into a spreadsheet and saved.

Columns used: Date, Compte, Libellé, NoPiece, Débit, Crédit. Checks: a piece
whose debits and credits differ, a piece of one line, the same entry twice,
an amount far from the account's usual one, and an account on the side it
seldom takes for this kind of piece (OD, T, A…: a receipt written as a debit
on a customer?). What is usual is learned on this computer from the days
reviewed before. Nothing is changed; the user decides.
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from pathlib import Path

from binaa_reconciler.compare.prices import _number, _rows

RED, YELLOW = "red", "yellow"
FACTOR = 5
LEARNED = 5
RARE = 0.1


@dataclass(frozen=True)
class Entry:
    row: int
    day: str
    account: str
    label: str
    piece: str
    debit: float
    credit: float

    @property
    def amount(self) -> float:
        return self.debit or self.credit

    @property
    def side(self) -> str:
        return "D" if self.debit else "C"

    @property
    def kind(self) -> str:
        """OD / T / A…: the letters of the piece number (26OD0341 → OD)."""
        found = re.match(r"\d*([A-Za-z]+)", self.piece)
        return found.group(1).upper() if found else ""


@dataclass(frozen=True)
class JournalFinding:
    entry: Entry
    level: str
    text: str


@dataclass
class JournalHistory:
    amounts: dict[str, list[float]] = field(default_factory=dict)
    sides: dict[str, Counter[str]] = field(default_factory=dict)  # "account|kind" → D/C
    seen: set[str] = field(default_factory=set)  # piece|row signature

    def learn(self, entries: Iterable[Entry]) -> None:
        for e in entries:
            key = f"{e.day}|{e.piece}|{e.account}|{e.amount}|{e.side}"
            if not e.amount or key in self.seen:
                continue
            self.seen.add(key)
            self.amounts.setdefault(e.account, []).append(e.amount)
            del self.amounts[e.account][:-300]
            self.sides.setdefault(f"{e.account}|{e.kind}", Counter())[e.side] += 1

    def to_json(self) -> dict[str, object]:
        return {
            "amounts": self.amounts,
            "sides": {k: dict(c) for k, c in self.sides.items()},
            "seen": sorted(self.seen)[-50000:],
        }

    @staticmethod
    def from_json(data: object) -> JournalHistory:
        h = JournalHistory()
        if not isinstance(data, dict):
            return h
        try:
            h.amounts = {a: [float(x) for x in v] for a, v in data.get("amounts", {}).items()}
            h.sides = {k: Counter(c) for k, c in data.get("sides", {}).items()}
            h.seen = set(data.get("seen", []))
        except (TypeError, ValueError, AttributeError):
            return JournalHistory()
        return h


def _text(value: object) -> str:
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return str(value if value is not None else "").strip()


def read_entries(path: Path) -> list[Entry]:
    rows = _rows(path)
    found = [
        (h, [str(c or "").strip().lower() for c in header])
        for h, header in enumerate(rows[:10])
        if {"compte", "nopiece"} <= {str(c or "").strip().lower() for c in header}
    ]
    if not found:
        raise ValueError("no_header")
    h, names = found[0]
    col = {n: k for k, n in enumerate(names)}
    debit = next(k for n, k in col.items() if n.startswith("d") and "bit" in n)
    credit = next(k for n, k in col.items() if n.startswith("cr"))
    label = next((k for n, k in col.items() if n.startswith("libell")), col["compte"])
    day = col.get("date", col["compte"])
    out = []
    for n, row in enumerate(rows[h + 1 :], start=h + 2):

        def cell(k: int, row: list[object] = row) -> object:
            return row[k] if k < len(row) else None

        account = _text(cell(col["compte"]))
        d, c = _number(cell(debit)) or 0.0, _number(cell(credit)) or 0.0
        if not account or (not d and not c):
            continue
        out.append(
            Entry(
                n, _text(cell(day)), account, _text(cell(label)), _text(cell(col["nopiece"])), d, c
            )
        )
    return out


def check_entries(
    entries: Sequence[Entry], history: JournalHistory | None = None
) -> list[JournalFinding]:
    learned = JournalHistory.from_json((history or JournalHistory()).to_json())
    learned.learn(entries)
    notes: dict[int, list[tuple[str, str]]] = {}

    def note(e: Entry, level: str, text: str) -> None:
        notes.setdefault(e.row, []).append((level, text))

    pieces: dict[str, list[Entry]] = {}
    for e in entries:
        pieces.setdefault(e.piece, []).append(e)
    for piece, group in pieces.items():
        d = sum(e.debit for e in group)
        c = sum(e.credit for e in group)
        if len(group) == 1:
            note(group[0], RED, f"القطعة {piece} سطر واحد بلا مقابل")
        elif abs(d - c) > 0.5:
            for e in group:
                note(e, RED, f"القطعة {piece} غير متوازنة: مدين {d:,.0f} ودائن {c:,.0f}")

    same: dict[tuple[str, str, float, str, str], list[Entry]] = {}
    for e in entries:
        same.setdefault((e.day, e.account, e.amount, e.side, e.label.lower()), []).append(e)
    for group in same.values():
        if len({e.piece for e in group}) > 1:
            for e in group:
                others = ", ".join(sorted({g.piece for g in group if g is not e}))
                note(e, YELLOW, f"نفس القيد في قطعة أخرى ({others}) — هل هو مكرر؟")

    for e in entries:
        past = list(learned.amounts.get(e.account, []))
        if e.amount in past:
            past.remove(e.amount)
        if len(past) >= LEARNED:
            usual = sorted(past)[len(past) // 2]
            if usual and e.amount >= usual * FACTOR:
                note(
                    e,
                    RED,
                    f"المبلغ أكبر بـ {round(e.amount / usual)} مرات من المعتاد ({usual:,.0f})",
                )
            elif usual and e.amount * FACTOR <= usual:
                note(
                    e,
                    RED,
                    f"المبلغ أصغر بـ {round(usual / e.amount)} مرات من المعتاد ({usual:,.0f})",
                )
        sides = learned.sides.get(f"{e.account}|{e.kind}", Counter())
        total = sum(sides.values()) - 1
        if total >= LEARNED and (sides[e.side] - 1) / total < RARE:
            usual_side = "مدين" if e.side == "C" else "دائن"
            mine = "مدين" if e.side == "D" else "دائن"
            note(e, RED, f"الجهة مقلوبة؟ {mine} — وهذا الحساب في قطع {e.kind} عادةً {usual_side}")

    rank = {RED: 0, YELLOW: 1}
    by_row = {e.row: e for e in entries}
    out = []
    for row, items in sorted(notes.items()):
        level = min((lv for lv, _ in items), key=rank.__getitem__)
        out.append(
            JournalFinding(by_row[row], level, " — ".join(dict.fromkeys(t for _, t in items)))
        )
    out.sort(key=lambda f: (rank[f.level], f.entry.row))
    return out


__all__ = [
    "RED",
    "YELLOW",
    "Entry",
    "JournalFinding",
    "JournalHistory",
    "check_entries",
    "read_entries",
]
