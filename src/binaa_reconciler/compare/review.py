"""Review of a day's sales: the RIM-Soft «Rapport ventes» printed to PDF.

Every line of the report (BR / FC / FT, number, date, description, HT,
remise, TVA, TTC, avance, remarques) is read from the PDF: the characters of
one line are taken in order and cut at the column separators (a «|» written
in the text or a thin vertical line drawn on the page). Then a few checks
point out what is probably an entry error; nothing is changed, the user
decides. Local only: the PDF is read on this computer, never sent anywhere.
"""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field
from datetime import date, datetime
from itertools import pairwise
from pathlib import Path

RED, YELLOW, GREY = "red", "yellow", "grey"
_DATE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
_TYPE = re.compile(r"^[A-Z]{2,3}$")
_INVOICE = re.compile(r"(?<![A-Za-z0-9])f\s*(\d{4,6})\b", re.IGNORECASE)
_ACCOUNT = re.compile(r"^\d{2}[0-9A-Za-z]*$")
SEQUENCE_JUMP = 20  # an invoice number this far from both neighbours is out of sequence
FACTOR = 5  # an amount 5 times above or below the account's usual one: 100 bags for 1000?
LEARNED = 5  # entries of an account needed before it is judged
RARE = 0.1  # a kind of entry (BR / FT…) below 10 % of an account's entries is unusual


@dataclass
class History:
    """What each account usually has, learned from the reports reviewed
    before (kept on this computer): its kinds of entry and its amounts."""

    kinds: dict[str, Counter[str]] = field(default_factory=dict)
    amounts: dict[str, list[int]] = field(default_factory=dict)
    seen: set[str] = field(default_factory=set)  # piece numbers already learned

    def learn(self, sales: Iterable[Sale]) -> None:
        for s in sales:
            if not s.account or s.number in self.seen or not s.amount:
                continue
            self.seen.add(s.number)
            self.kinds.setdefault(s.account, Counter())[s.kind] += 1
            self.amounts.setdefault(s.account, []).append(s.amount)
            del self.amounts[s.account][:-300]  # the recent ones only

    def to_json(self) -> dict[str, object]:
        return {
            "kinds": {a: dict(c) for a, c in self.kinds.items()},
            "amounts": self.amounts,
            "seen": sorted(self.seen)[-20000:],
        }

    @staticmethod
    def from_json(data: object) -> History:
        h = History()
        if not isinstance(data, dict):
            return h
        try:
            h.kinds = {a: Counter(c) for a, c in data.get("kinds", {}).items()}
            h.amounts = {a: [int(x) for x in v] for a, v in data.get("amounts", {}).items()}
            h.seen = set(data.get("seen", []))
        except (TypeError, ValueError, AttributeError):
            return History()
        return h


@dataclass(frozen=True)
class Sale:
    kind: str  # BR, FC, FT…
    number: str  # RIM-Soft piece number (26T06843)
    day: date
    label: str
    amount: int | None  # TTC, whole units; None: empty
    page: int = 1

    @property
    def account(self) -> str:
        first = self.label.split()[0] if self.label.split() else ""
        return first if _ACCOUNT.match(first) else ""

    @property
    def invoice(self) -> int | None:
        """The hand-written invoice number («f 15360»), if any."""
        found = _INVOICE.search(self.label)
        return int(found.group(1)) if found else None


@dataclass(frozen=True)
class Finding:
    sale: Sale
    level: str  # RED / YELLOW / GREY
    text: str


@dataclass
class Review:
    sales: list[Sale] = field(default_factory=list)
    findings: list[Finding] = field(default_factory=list)
    missing: list[int] = field(default_factory=list)  # invoice numbers not found
    unread: int = 0  # lines that look like sales but could not be read

    def count(self, level: str) -> int:
        return sum(f.level == level for f in self.findings)


# --------------------------------------------------------------- reading


def amount(text: str) -> int | None:
    """«3 408 000» → 3408000; empty → None."""
    digits = re.sub(r"[^\d-]", "", text)
    return int(digits) if digits not in ("", "-") else None


def parse_row(fields: Sequence[str], page: int = 1) -> Sale | None:
    """One report line cut into its columns: type, number, date, description,
    HT, remise, TVA, TTC (the rest is not used)."""
    cells = [" ".join(f.split()) for f in fields]
    while cells and not cells[0]:
        cells.pop(0)
    if len(cells) < 8 or not _TYPE.match(cells[0]) or not _DATE.match(cells[2]):
        return None
    day = datetime.strptime(cells[2], "%d/%m/%Y").date()
    ttc = amount(cells[7])
    if ttc is None:
        ttc = amount(cells[4])  # no TTC: the HT
    return Sale(cells[0], cells[1], day, cells[3], ttc, page)


def _rows(
    chars: Iterable[tuple[float, float, str]], tolerance: float = 2.5
) -> list[tuple[float, list[tuple[float, str]]]]:
    """(x, y, char) → lines, top to bottom, each line's chars left to right."""
    lines: list[tuple[float, list[tuple[float, str]]]] = []
    for x, y, c in sorted(chars, key=lambda t: -t[1]):
        if lines and abs(lines[-1][0] - y) <= tolerance:
            lines[-1][1].append((x, c))
        else:
            lines.append((y, [(x, c)]))
    return [(y, sorted(line)) for y, line in lines]


def split_line(line: Sequence[tuple[float, str]], separators: Sequence[float]) -> list[str]:
    """Cut one line at «|» characters and at drawn vertical separators; an
    empty column between two separators is an empty field. A drawn line next
    to a «|» character is the same separator."""
    bars = [x for x, c in line if c == "|"]
    seps = sorted(x for x in separators if all(abs(x - b) > 3 for b in bars))
    fields: list[str] = []
    current: list[str] = []
    k = 0
    for x, c in line:
        while k < len(seps) and seps[k] < x:
            fields.append("".join(current))
            current = []
            k += 1
        if c == "|":
            fields.append("".join(current))
            current = []
        else:
            current.append(c)
    fields.append("".join(current))
    fields.extend("" for _ in seps[k:])  # empty columns at the end of the line
    return [f.strip() for f in fields]


def read_pdf(path: Path) -> tuple[list[Sale], int]:
    """The sales of the report, and how many lines starting like a sale
    could not be read."""
    from pdfminer.high_level import extract_pages
    from pdfminer.layout import LTChar, LTCurve, LTLine, LTRect

    sales: list[Sale] = []
    unread = 0
    for number, page in enumerate(extract_pages(str(path), laparams=None), start=1):
        chars: list[tuple[float, float, str]] = []
        verticals: list[tuple[float, float, float]] = []  # x, y0, y1
        stack: list[object] = [page]
        while stack:
            item = stack.pop()
            if isinstance(item, LTChar):
                chars.append((item.x0, (item.y0 + item.y1) / 2, item.get_text()))
            elif isinstance(item, LTLine | LTRect | LTCurve):
                if item.width <= 1.5 and item.height >= 4:
                    verticals.append(((item.x0 + item.x1) / 2, item.y0, item.y1))
            elif isinstance(item, Iterable):
                stack.extend(item)
        for y, line in _rows(chars):
            seps = [x for x, y0, y1 in verticals if y0 - 2 <= y <= y1 + 2]
            fields = split_line(line, seps)
            sale = parse_row(fields, number)
            if sale is not None:
                sales.append(sale)
            elif _TYPE.match(fields[0].strip()) and any(_DATE.match(f.strip()) for f in fields):
                unread += 1
    return sales, unread


# --------------------------------------------------------------- checks


def check(sales: Sequence[Sale], unread: int = 0, history: History | None = None) -> Review:
    review = Review(list(sales), unread=unread)
    learned = History.from_json((history or History()).to_json())  # a copy
    learned.learn(sales)  # this report counts too (each entry without itself, below)
    notes: dict[int, list[tuple[str, str]]] = {}

    def note(i: int, level: str, text: str) -> None:
        notes.setdefault(i, []).append((level, text))

    # invoices only (FT / FC…); delivery notes (BR) have their own numbers
    invoiced = [(i, s) for i, s in enumerate(sales) if s.kind != "BR" and s.invoice]
    counts = Counter(s.invoice for _, s in invoiced)
    for i, s in invoiced:
        if counts[s.invoice] > 1:
            others = [t.number for _, t in invoiced if t.invoice == s.invoice and t is not s]
            note(i, RED, f"رقم الفاتورة f {s.invoice} مكرر (أيضًا في {', '.join(others)})")

    order = sorted(invoiced, key=lambda t: t[1].number)
    outliers: set[int] = set()
    guessed: set[int] = set()
    for k in range(1, len(order) - 1):
        (_, before), (i, s), (_, after) = order[k - 1], order[k], order[k + 1]
        if not (s.invoice and before.invoice and after.invoice):
            continue
        if (
            abs(s.invoice - before.invoice) > SEQUENCE_JUMP
            and abs(s.invoice - after.invoice) > SEQUENCE_JUMP
            and abs(after.invoice - before.invoice) <= SEQUENCE_JUMP
        ):
            outliers.add(i)
            guess = before.invoice + 1
            guessed.add(guess)  # most likely the number meant: not «missing»
            note(
                i,
                RED,
                f"رقم الفاتورة f {s.invoice} خارج التسلسل"
                f" (بين f {before.invoice} و f {after.invoice}) — هل هو f {guess}؟",
            )
    known = sorted({s.invoice for i, s in invoiced if i not in outliers and s.invoice})
    for prev, nxt in pairwise(known):
        if 1 < nxt - prev <= SEQUENCE_JUMP:
            review.missing.extend(n for n in range(prev + 1, nxt) if n not in guessed)

    for i, s in enumerate(sales):
        if not s.amount:
            if "annul" in s.label.lower():
                note(i, GREY, "فاتورة ملغاة (بلا مبلغ)")
            else:
                note(i, RED, "بلا مبلغ")
        if s.kind != "BR" and not s.account:
            note(i, YELLOW, "لا يبدأ البيان برقم حساب")

    same: dict[tuple[str, int | None, date], list[int]] = {}
    for i, s in enumerate(sales):
        if s.amount and not s.invoice:
            same.setdefault((s.account, s.amount, s.day), []).append(i)
    for group in same.values():
        if len(group) > 1:
            for i in group:
                note(i, YELLOW, "نفس الحساب والمبلغ والتاريخ مرتين — هل هي مكررة؟")

    for i, s in enumerate(sales):
        if not s.account or not s.amount:
            continue
        kinds = learned.kinds.get(s.account, Counter())
        total = sum(kinds.values())
        if total - 1 >= LEARNED:
            mine = kinds[s.kind] - (1 if s.number in learned.seen else 0)
            if mine / (total - 1) < RARE:
                usual = kinds.most_common(1)[0][0]
                note(i, RED, f"نوع غير معتاد لهذا الحساب: {s.kind} — عادةً {usual}")
        past = list(learned.amounts.get(s.account, []))
        if s.amount in past:
            past.remove(s.amount)
        if len(past) >= LEARNED:
            usual_amount = sorted(past)[len(past) // 2]
            if s.amount >= usual_amount * FACTOR:
                times = round(s.amount / usual_amount)
                note(
                    i,
                    RED,
                    f"المبلغ أكبر بـ {times} مرات من المعتاد ({usual_amount:,}) — هل الكمية صحيحة؟",
                )
            elif s.amount * FACTOR <= usual_amount:
                times = round(usual_amount / s.amount)
                note(
                    i,
                    RED,
                    f"المبلغ أصغر بـ {times} مرات من المعتاد ({usual_amount:,}) — هل الكمية صحيحة؟",
                )

    rank = {RED: 0, YELLOW: 1, GREY: 2}
    for i, items in sorted(notes.items()):
        level = min((lv for lv, _ in items), key=rank.__getitem__)
        review.findings.append(Finding(sales[i], level, " — ".join(t for _, t in items)))
    review.findings.sort(key=lambda f: (rank[f.level], f.sale.number))
    return review


def review_pdf(path: Path, history: History | None = None) -> Review:
    sales, unread = read_pdf(path)
    return check(sales, unread, history)


__all__ = [
    "GREY",
    "RED",
    "YELLOW",
    "Finding",
    "History",
    "Review",
    "Sale",
    "check",
    "parse_row",
    "read_pdf",
    "review_pdf",
    "split_line",
]
