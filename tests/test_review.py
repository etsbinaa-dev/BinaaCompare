"""Review of a day's sales (synthetic PDFs shaped like the RIM-Soft report)."""

from __future__ import annotations

from datetime import date

import pytest

from binaa_reconciler.compare.review import (
    GREY,
    RED,
    YELLOW,
    Sale,
    check,
    parse_row,
    review_pdf,
    split_line,
)

D = date(2026, 9, 27)


def sale(number: str, label: str, amount: int | None, kind: str = "FT") -> Sale:
    return Sale(kind, number, D, label, amount)


def test_a_report_line_is_cut_into_its_columns():
    fields = ["FT ", "26T00843 ", "27/09/2026 ", "55SYN BANK f 15360 syn ", " 18 000", " ", " "]
    fields += [" 18 000", " ", ""]
    s = parse_row(fields)
    assert s == Sale("FT", "26T00843", D, "55SYN BANK f 15360 syn", 18_000)
    assert (s.account, s.invoice) == ("55SYN", 15_360)
    assert (
        parse_row(["Sous total des factures ", " 5 635 280", " 2 000", " ", " 5 633 280"]) is None
    )
    line = [(float(x), c) for x, c in enumerate("FT 26T1 27/09/2026 SYN 5")]
    assert split_line(line, [2.5, 7.5, 18.5, 22.5, 25]) == [
        "FT",
        "26T1",
        "27/09/2026",
        "SYN",
        "5",
        "",
    ]


def test_the_checks_point_out_what_looks_wrong():
    sales = [
        sale("26T00843", "55SYN BANK f 15360 a", 18_000),
        sale("26T00844", "41SYN ONE f 15361 b", 117_000),
        sale("26T00845", "41SYN TWO f 15145 c", 122_000),  # 15362 typed wrong
        sale("26T00846", "41SYN TRE f 15363 d", 186_000),
        sale("26T00847", "41SYN FOR f 15363 e", 186_000),  # the same invoice twice
        sale("26T00848", "41SYN FIV f 15366 f", 72_300),  # 15364, 15365 missing
        sale("26T00849", "41F SYN f 15367 annule", None),  # a cancelled invoice
        sale("26T00850", "41SYN SIX f 15368", None),  # no amount
        sale("26T00851", "client sans compte f 15369", 3_000),
        sale("26A00937", "40001 SYN 60t BR 26A00587", 3_408_000, kind="BR"),
    ]
    r = check(sales)
    found = {f.sale.number: (f.level, f.text) for f in r.findings}
    assert found["26T00845"] == (
        RED,
        "رقم الفاتورة f 15145 خارج التسلسل (بين f 15361 و f 15363) — هل هو f 15362؟",
    )
    assert found["26T00846"][0] == RED and "f 15363 مكرر (أيضًا في 26T00847)" in found["26T00846"][1]
    assert found["26T00849"] == (GREY, "فاتورة ملغاة (بلا مبلغ)")
    assert found["26T00850"] == (RED, "بلا مبلغ")
    assert found["26T00851"] == (YELLOW, "لا يبدأ البيان برقم حساب")
    assert r.missing == [15_364, 15_365]
    assert "26A00937" not in found and "26T00843" not in found  # nothing to say
    assert [f.level for f in r.findings][:2] == [RED, RED]  # the most serious first


def test_the_same_entry_twice_without_invoice_number():
    r = check([sale("26T1", "41SYN ONE cash", 5_000), sale("26T2", "41SYN ONE cash", 5_000)])
    assert {f.level for f in r.findings} == {YELLOW}
    distinct = check([sale("26T1", "41SYN f 15001", 5_000), sale("26T2", "41SYN f 15002", 5_000)])
    assert distinct.findings == []  # two invoices of the same amount: normal


ROWS = [
    ("BR", "26A00937", "27/09/2026", "40001 SYN CLINIC 60t bl 3", "3 408 000", "", "", "3 408 000"),
    ("FT", "26T00843", "27/09/2026", "55SYN BANK f 15360 syn", "18 000", "", "", "18 000"),
    ("FT", "26T00844", "27/09/2026", "41SYN ONE f 15145 one", "117 000", "", "", "117 000"),
    ("FT", "26T00845", "27/09/2026", "41SYN TWO f 15362 two", "122 000", "2 000", "", "120 000"),
    ("FT", "26T00846", "27/09/2026", "41F SYN f 15363 annule", "", "", "", ""),
]
X = (40, 70, 130, 190, 400, 460, 500, 540, 600)


def _pdf(path, pipes: bool):
    canvas = pytest.importorskip("reportlab.pdfgen.canvas")
    c = canvas.Canvas(str(path))
    c.setFont("Helvetica", 8)
    c.drawString(40, 800, "Facture")
    y = 760
    for row in ROWS:
        for k, text in enumerate(row):
            if k >= 4:  # amounts right-aligned in their column
                c.drawRightString(X[k + 1] - 4, y, text)
            else:
                c.drawString(X[k] + 4, y, text)
        for x in X[1:]:  # the column separators: «|» written, or lines drawn
            if pipes:
                c.drawString(x, y, "|")
            else:
                c.line(x, y - 3, x, y + 9)
        y -= 16
    c.drawString(300, y, "Sous total des factures a terme 257 000")
    c.showPage()
    c.save()
    return path


@pytest.mark.parametrize("pipes", [True, False], ids=["pipes in the text", "drawn lines"])
def test_the_report_is_read_from_the_pdf(tmp_path, pipes):
    r = review_pdf(_pdf(tmp_path / "ventes.pdf", pipes))
    assert [(s.kind, s.number, s.amount) for s in r.sales] == [
        ("BR", "26A00937", 3_408_000),
        ("FT", "26T00843", 18_000),
        ("FT", "26T00844", 117_000),
        ("FT", "26T00845", 120_000),
        ("FT", "26T00846", None),
    ]
    assert r.sales[2].label == "41SYN ONE f 15145 one" and r.unread == 0
    found = {f.sale.number: f.level for f in r.findings}
    assert found == {"26T00844": RED, "26T00846": GREY}
