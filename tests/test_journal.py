"""Journal entries review (synthetic entries shaped like «Copier écritures»)."""

from __future__ import annotations

from openpyxl import Workbook

from binaa_reconciler.compare.journal import (
    RED,
    YELLOW,
    Entry,
    JournalHistory,
    check_entries,
    read_entries,
)

HEAD = ["NoOrdre", "Date", "Compte", "Libellé", "NoPiece", "Débit", "Crédit", "Dev.Deb", "Dev.Cred"]


def e(row, account, piece, debit=0, credit=0, label="SYN", day="29/09/2026"):
    return Entry(row, day, account, label, piece, debit, credit)


def test_the_pasted_columns_are_read(tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(HEAD)
    ws.append(["030274", "29/09/2026", 411981, "SYN a", "26OD0341", 250000, 0, 0, 0])
    ws.append(["030275", "29/09/2026", "55SYN", "SYN a", "26OD0341", 0, 250000, 0, 0])
    wb.save(tmp_path / "j.xlsx")
    first, second = read_entries(tmp_path / "j.xlsx")
    assert (first.account, first.piece, first.kind, first.side, first.amount) == (
        "411981",
        "26OD0341",
        "OD",
        "D",
        250000,
    )
    assert (second.row, second.side) == (3, "C")


def test_unbalanced_lonely_and_repeated_entries():
    entries = [
        e(2, "41SYN", "26OD1", debit=500),
        e(3, "55SYN", "26OD1", credit=450),  # does not balance
        e(4, "41SYN", "26OD2", debit=70),  # alone
        e(5, "41ONE", "26OD3", debit=900, label="same"),
        e(6, "55SYN", "26OD3", credit=900, label="same"),
        e(7, "41ONE", "26OD4", debit=900, label="same"),  # the same entry again
        e(8, "55SYN", "26OD4", credit=900, label="same"),
    ]
    found = {f.entry.row: (f.level, f.text) for f in check_entries(entries)}
    assert found[2] == (RED, "القطعة 26OD1 غير متوازنة: مدين 500 ودائن 450")
    assert found[4] == (RED, "القطعة 26OD2 سطر واحد بلا مقابل")
    assert found[5][0] == YELLOW and "26OD4" in found[5][1]


def test_learned_usual_amount_and_side_of_each_account():
    history = JournalHistory()
    history.learn(
        [e(10 + k, "41SYN", f"26A{k}", credit=3_400_000) for k in range(6)]
        + [e(20 + k, "6SYN", f"26A{k}", debit=3_400_000) for k in range(6)]
    )
    today = [
        e(2, "41SYN", "26A90", credit=340_000),  # 10 times less: 100 bags for 1000?
        e(3, "6SYN", "26A90", debit=340_000),
        e(4, "41SYN", "26A91", debit=3_400_000),  # the reverse side
        e(5, "6SYN", "26A91", credit=3_400_000),
    ]
    found = {f.entry.row: f.text for f in check_entries(today, history)}
    assert found[2].startswith("المبلغ أصغر بـ 10 مرات من المعتاد (3,400,000)")
    assert "الجهة مقلوبة؟ مدين — وهذا الحساب في قطع A عادةً دائن" in found[4]
