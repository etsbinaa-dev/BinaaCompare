"""The few sentences of BinaaCompare (shared by the screen and the Excel file)."""

from __future__ import annotations

from collections.abc import Sequence
from datetime import timedelta

from binaa_reconciler.compare.engine import (
    SIDE_AR,
    Kind,
    Line,
    Pair,
    Result,
    amount_text,
    mirror,
)
from binaa_reconciler.core.models import LedgerSide

ARABIC_NAMES = {"binaa": "بناء", "clinique": "كلينك", "clinic": "كلينك", "cliniq": "كلينك"}


def arabic(name: str) -> str:
    """The party's name in Arabic letters when known: no Latin word inside an
    Arabic sentence, so the sentence never breaks (else the name as typed)."""
    return ARABIC_NAMES.get(name.strip().casefold(), name)


def arabic_text(text: str, names: tuple[str, str]) -> str:
    """The parties' names written in Arabic letters inside a sentence."""
    for name in names:
        text = text.replace(name, arabic(name))
    return text


def title(result: Result) -> str:
    a, b = result.names
    if result.period is None:
        return f"مقارنة {a} ↔ {b}"
    start, end = result.period
    return f"مقارنة {a} ↔ {b} — الفترة {start:%d/%m/%Y} → {end:%d/%m/%Y}"


def headline(result: Result) -> str:
    a, b = result.names
    n_a, n_b = (len(x) for x in result.only)
    left = len(result.possible) + len(result.close)
    if not n_a and not n_b and not left:
        return f"✅ كل العمليات متطابقة ({result.matched_count} عملية)"
    parts = [f"✅ متطابقة: {result.matched_count}"]
    if result.possible:
        parts.append(f"🔎 قد تكون نفس العملية: {len(result.possible)}")
    if result.close:
        parts.append(f"💲 فرق في المبلغ: {len(result.close)}")
    parts.append(f"❌ عند {a} فقط: {n_a}")
    parts.append(f"❌ عند {b} فقط: {n_b}")
    if result.settled:
        parts.append(f"✔️ تمت تسويتها: {len(result.settled)}")
    return " — ".join(parts)


def footnote(result: Result) -> str:
    """The lines left out of the lists because they are outside the period."""
    a, b = result.names
    out_a, out_b = result.outside
    if not out_a and not out_b:
        return ""
    return (
        f"غير معروضة لأنها خارج الفترة المختارة: {out_a} عملية عند {a}، {out_b} عملية عند {b} "
        "(غيّر «من / إلى» لعرضها)"
    )


def total_text(result: Result, i: int) -> str:
    return f"المجموع: {amount_text(result.total(i))}"


def _where(result: Result) -> dict[str, str]:
    """What became of each line: matched with what, in which list."""
    a, b = result.names
    out: dict[str, str] = {}
    for x, y in result.matched:
        out[x.key] = f"مطابقة مع {b} {y.day:%d/%m}"
        out[y.key] = f"مطابقة مع {a} {x.day:%d/%m}"
    for pair in result.confirmed:
        for x in pair.lines:
            out[x.key] = "مطابقة (أكدتها أنت)"
    for pair in result.possible:
        for x in pair.lines:
            out[x.key] = "في «قد تكون نفس العملية»"
    for pair in result.close:
        for x in pair.lines:
            out[x.key] = "في «فرق في المبلغ»"
    for i, lines in enumerate(result.only):
        for x in lines:
            out[x.key] = f"عند {result.names[i]} فقط"
    for x in result.settled:
        out[x.key] = "تمت تسويتها"
    return out


def same_amount(
    result: Result, line: Line, mine: Sequence[Line], others: Sequence[Line], i: int
) -> str:
    """One short conclusion under a line alone on side i: how many times this
    amount appears on each side around its date (15 days before and after),
    and whether the other side's ones are all taken by other matches."""
    me, other = arabic(result.names[i]), arabic(result.names[1 - i])
    near = [line.day - timedelta(days=15), line.day + timedelta(days=15)]

    def around(lines: Sequence[Line], side: LedgerSide) -> list[Line]:
        return [
            x
            for x in lines
            if x.amount == line.amount and x.side is side and near[0] <= x.day <= near[1]
        ]

    ours = around(mine, line.side)
    theirs = around(others, mirror(line.side))
    amount = amount_text(line.amount)
    span = f"بين {near[0]:%d/%m} و{near[1]:%d/%m}"
    if not theirs:
        return f"المبلغ {amount} غير موجود عند {other} {span}"
    taken = {k for x, y in result.matched for k in (x.key, y.key)}
    taken |= {x.key for p in result.confirmed for x in p.lines}
    free = [x for x in theirs if x.key not in taken]
    count = f"المبلغ {amount} {span}: {me} {len(ours)} مرة — {other} {len(theirs)} مرة"
    if not free:
        return f"{count}\nكلها مطابقة: عملية زائدة عند {me} أو ناقصة عند {other}"
    days = "، ".join(f"{x.day:%d/%m}" for x in free[:3])
    return f"{count}\nغير المطابقة عند {other}: {days}"


def reason(pair: Pair, names: tuple[str, str]) -> str:
    a, b = names
    if pair.kind is Kind.REVERSAL:
        where = a if pair.a else b
        return f"حركة وعكسها في ملف {where} (صافي الأثر صفر)"
    if pair.kind is Kind.SAME:
        days = abs((pair.a[0].day - pair.b[0].day).days)
        return f"نفس المبلغ — فرق {days} يوم في التاريخ"
    if pair.kind is Kind.SIDE:
        side = SIDE_AR[pair.a[0].side]
        return f"نفس المبلغ لكن {side} في الطرفين — تحقق من جهة القيد"
    if pair.kind is Kind.SUM:
        n = max(len(pair.a), len(pair.b))
        count = "عمليتين" if n == 2 else f"{n} عمليات"
        return f"مجموع {count} من يوم واحد يساوي عملية واحدة"
    return f"فرق في المبلغ: {amount_text(abs(pair.difference))}"


def side_cells(lines: tuple[Line, ...]) -> tuple[str, str, str]:
    """Date, amount and description of one side of a pair ("—" when empty)."""
    if not lines:
        return "—", "—", "—"
    days = sorted({f"{x.day:%d/%m/%Y}" for x in lines})
    amount = " + ".join(amount_text(x.amount) for x in lines)
    label = " / ".join(x.label for x in lines if x.label)
    return " ، ".join(days), amount, label


__all__ = [
    "ARABIC_NAMES",
    "arabic",
    "arabic_text",
    "footnote",
    "headline",
    "reason",
    "same_amount",
    "side_cells",
    "title",
    "total_text",
]
