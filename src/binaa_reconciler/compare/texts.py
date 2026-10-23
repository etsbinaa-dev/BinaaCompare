"""The few sentences of BinaaCompare (shared by the screen and the Excel file)."""

from __future__ import annotations

from binaa_reconciler.compare.engine import SIDE_AR, Kind, Line, Pair, Result, amount_text


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


__all__ = ["footnote", "headline", "reason", "side_cells", "title", "total_text"]
