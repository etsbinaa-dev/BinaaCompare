"""Sale prices check (synthetic product lines shaped like the RIM-Soft copy)."""

from __future__ import annotations

from openpyxl import Workbook

from binaa_reconciler.compare.prices import PriceHistory, check_prices, read_lines

HEAD = ["Code", "Désignation", "Recu", "Vendu", "Différence", "PrixVente", "PrixProduit"]


def _xlsx(path, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(HEAD)
    for row in rows:
        ws.append(row)
    wb.save(path)
    return path


def test_a_price_far_from_the_usual_one_is_pointed_out_with_its_line(tmp_path):
    rows = [
        ["P0001", "SYN CIMENT", 0, 40, -40, 3050, 0],
        ["P0001", "SYN CIMENT", 0, 15, -15, 3000, 0],
        ["P0001", "SYN CIMENT", 0, 20, -20, 2925, 0],
        ["P0001", "SYN CIMENT", 0, 10, -10, 305, 0],  # a zero forgotten
        ["P0050", "SYN FER", 0, 2, -2, 40000, 0],  # one price only: not judged
        ["P0001", "SYN CIMENT", 0, 5, -5, 3025, 0],
    ]
    lines = read_lines(_xlsx(tmp_path / "lignes.xlsx", rows))
    assert [(x.row, x.code, x.quantity, x.price) for x in lines][3] == (5, "P0001", 10, 305)
    (f,) = check_prices(lines)
    assert f.line.row == 5 and f.text == "السعر 305 والمعتاد 3,025 — 10 مرات أصغر"


def test_prices_learned_before_judge_a_new_file_and_a_tab_text_is_read(tmp_path):
    history = PriceHistory()
    history.learn(
        read_lines(_xlsx(tmp_path / "a.xlsx", [["P0050", "SYN FER", 0, 2, -2, 40000, 0]] * 3))
    )
    text = "\t".join(HEAD) + "\nP0050\tSYN FER\t0\t1\t-1\t52 000\t0\n"
    (tmp_path / "b.txt").write_text(text, encoding="cp1252")
    (f,) = check_prices(read_lines(tmp_path / "b.txt"), PriceHistory.from_json(history.to_json()))
    assert f.text == "السعر 52,000 والمعتاد 40,000 — 30% أكبر"
