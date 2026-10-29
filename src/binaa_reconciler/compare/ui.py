"""BinaaCompare window: two files, one button, two lists.

Everything happens on this one screen: choose the two RIM-Soft files, press
«قارن», read the two lists, export to Excel. The ✓ / ✗ decisions are saved
on this computer (memory.py).
"""

from __future__ import annotations

import contextlib
import os
import sys
import tkinter as tk
from datetime import date, datetime
from functools import partial
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import ClassVar

from binaa_reconciler.compare import for_party, manual, memory, prices, review
from binaa_reconciler.compare.engine import (
    SIDE_AR,
    Kind,
    Pair,
    Result,
    amount_text,
    common_period,
    compare,
)
from binaa_reconciler.compare.files import Loaded, load
from binaa_reconciler.compare.texts import (
    arabic,
    arabic_text,
    footnote,
    headline,
    reason,
    same_amount,
    side_cells,
    title,
    total_text,
)

FONT = ("Segoe UI", 11)
CARDS = (  # key, title, (background, text colour): the colours of the Excel files too
    ("matched", "✅ متطابقة", ("#E8F5E9", "#2E7D32")),
    ("possible", "🔎 قد تكون نفس العملية", ("#FFF3E0", "#B45309")),
    ("close", "💲 فرق في المبلغ", ("#FFF8E1", "#8D6E00")),
    ("A", "❌ عند الطرف الأول فقط", ("#FDECEA", "#A01010")),
    ("B", "❌ عند الطرف الثاني فقط", ("#FDECEA", "#A01010")),
)
STRIPE = "#F5F7FA"  # every other line of the lists
BIG = ("Segoe UI", 13, "bold")
GREY, ORANGE, GREEN, RED = "#555555", "#b45309", "#2e7d32", "#a01010"
COLUMNS = ("التاريخ", "المبلغ", "الجهة", "الوصف", "المرجع", "ملاحظة")
WIDTHS = (90, 110, 55, 300, 110, 330)
PAIR_COLUMNS = ("تاريخ", "مبلغ", "وصف", "تاريخ", "مبلغ", "وصف", "الفرق", "السبب")
PAIR_WIDTHS = (90, 120, 200, 90, 120, 200, 90, 300)
PRIVACY = "🔒 كل شيء على هذا الجهاز فقط — لا يُرسل أي ملف إلى الإنترنت — قراراتك تُحفظ هنا"


class Dialogs:
    """The system dialogs (replaced in tests)."""

    def ask_open(self, last: Path | None = None) -> Path | None:
        """Opens where the last file was (and proposes it)."""
        name = filedialog.askopenfilename(
            title="اختر ملف RIM-Soft",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("الكل", "*.*")],
            initialdir=str(last.parent) if last else None,
            initialfile=last.name if last else None,
        )
        return Path(name) if name else None

    def ask_open_pdf(self) -> Path | None:
        name = filedialog.askopenfilename(
            title="اختر تقرير المبيعات (PDF)", filetypes=[("PDF", "*.pdf"), ("الكل", "*.*")]
        )
        return Path(name) if name else None

    def ask_open_prices(self) -> Path | None:
        name = filedialog.askopenfilename(
            title="اختر ملف أسطر المبيعات",
            filetypes=[("Excel / نص", "*.xlsx *.xlsm *.csv *.txt"), ("الكل", "*.*")],
        )
        return Path(name) if name else None

    def ask_save(self, suggested: str) -> Path | None:
        name = filedialog.asksaveasfilename(
            title="حفظ ملف Excel",
            defaultextension=".xlsx",
            initialfile=suggested,
            filetypes=[("Excel", "*.xlsx")],
        )
        return Path(name) if name else None

    def info(self, title: str, text: str) -> None:
        messagebox.showinfo(title, text)

    def open_file(self, path: Path) -> None:
        """Open the saved workbook in Excel (Windows); elsewhere: nothing."""
        if sys.platform.startswith("win"):
            with contextlib.suppress(OSError):
                os.startfile(path)  # type: ignore[attr-defined]  # noqa: S606

    def error(self, title: str, text: str) -> None:
        messagebox.showerror(title, text)


def label(parent: tk.Misc, text: str = "", **kw: object) -> ttk.Label:
    widget = ttk.Label(parent, text=text, anchor="e", justify="right")
    if kw:
        widget.configure(**kw)  # type: ignore[call-overload]
    return widget


class FileBox(ttk.LabelFrame):
    """One party: its name and its file."""

    def __init__(self, parent: tk.Misc, app: CompareApp, title_text: str, name: str, key: str):
        super().__init__(parent, text=title_text, labelanchor="ne", padding=10)
        self.app, self.key = app, key
        self.loaded: Loaded | None = None
        self.name_var = tk.StringVar(value=name)
        self.status_var = tk.StringVar(value="لم يتم اختيار ملف")
        top = ttk.Frame(self)
        top.pack(fill="x")
        label(top, "الاسم:").pack(side="right")
        ttk.Entry(top, textvariable=self.name_var, justify="right", width=18).pack(
            side="right", padx=6
        )
        ttk.Button(top, text="اختيار الملف…", command=self.choose).pack(side="left")
        self.status = label(self, textvariable=self.status_var, foreground=GREY)
        self.status.pack(fill="x", pady=(6, 0))

    @property
    def name(self) -> str:
        return self.name_var.get().strip() or self.key

    def choose(self) -> None:
        path = self.app.dialogs.ask_open(memory.last_file(self.key))
        if path is not None:
            self.load(path)

    def load(self, path: Path, again: bool = False) -> None:
        """`again`: the file of last time, reopened at start."""
        self.app.configure(cursor="watch")
        self.app.update_idletasks()
        try:
            self.loaded = load(path, self.key)
        finally:
            self.app.configure(cursor="")
        if self.loaded.error:
            self.status_var.set(f"❌ {path.name}: {self.loaded.error}")
            self.status.configure(foreground=RED)
            self.loaded = None
        else:
            text = f"✅ {path.name}: {len(self.loaded.lines)} عملية"
            if again:  # which export it is: its date, to check it is the latest
                saved = datetime.fromtimestamp(path.stat().st_mtime)
                text = f"📂 الملف السابق ({saved:%d/%m/%Y %H:%M}) — " + text
            with contextlib.suppress(OSError):
                memory.remember_file(self.key, path)
            if self.loaded.merged:
                text += f" — {self.loaded.merged} عملية مقسومة على سطرين دُمجت"
            if self.loaded.problems:
                text += f" — ⚠️ {len(self.loaded.problems)} سطر لم يُقرأ"
            self.status_var.set(text)
            self.status.configure(foreground=ORANGE if self.loaded.problems else GREEN)
        self.app.restore = True  # other files: their own saved decisions
        self.app.propose_period()
        self.app.refresh_buttons()


class CompareApp(tk.Tk):
    def __init__(self, dialogs: Dialogs | None = None) -> None:
        super().__init__()
        self.dialogs = dialogs or Dialogs()
        self.result: Result | None = None
        self.sources: tuple[Loaded, Loaded] | None = None
        self.title("مقارنة حسابين — BinaaCompare")
        self.option_add("*Font", FONT)
        self.minsize(760, 520)
        self._fit(1200, 780)
        accent = _modern_theme(self)
        style = ttk.Style(self)
        style.configure("Big.TButton", font=BIG, padding=(18, 8))
        style.configure("Treeview", rowheight=28, font=FONT)
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        self.go_style = "Accent.TButton" if accent else "Big.TButton"
        if accent:
            style.configure("Accent.TButton", font=BIG, padding=(22, 8))

        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)
        head = ttk.Frame(root)
        head.pack(fill="x")
        label(head, "مقارنة حسابين", font=("Segoe UI", 15, "bold")).pack(side="right")
        label(head, PRIVACY, foreground=GREY).pack(side="left")

        files = ttk.Frame(root)
        files.pack(fill="x", pady=4)
        self.box_a = FileBox(files, self, "الطرف الأول", "binaa", "A")
        self.box_b = FileBox(files, self, "الطرف الثاني", "clinique", "B")
        self.box_a.pack(side="right", fill="x", expand=True, padx=(0, 6))
        self.box_b.pack(side="right", fill="x", expand=True)
        self.after_idle(self._reopen_last_files)

        actions = ttk.Frame(root)
        actions.pack(fill="x")
        self.compare_button = ttk.Button(
            actions, text="قارن", style=self.go_style, command=self.run
        )
        self.compare_button.pack(side="right")
        self.from_var, self.to_var = tk.StringVar(), tk.StringVar()
        label(actions, "الفترة من:").pack(side="right", padx=(16, 4))
        ttk.Entry(actions, textvariable=self.from_var, width=11, justify="center").pack(
            side="right"
        )
        label(actions, "إلى:").pack(side="right", padx=(8, 4))
        ttk.Entry(actions, textvariable=self.to_var, width=11, justify="center").pack(side="right")
        self.manual_button = ttk.Button(actions, text="تصدير Excel", command=self.export_manual)
        self.manual_button.pack(side="left")
        self.party_button = ttk.Button(
            actions, text="ملف مشترك للمكالمة", command=self.export_for_party
        )
        self.party_button.pack(side="left", padx=(0, 6))
        ttk.Button(actions, text="🔍 مراجعة المبيعات", command=self.review_sales).pack(
            side="left", padx=(0, 6)
        )
        ttk.Button(actions, text="💲 مراجعة الأسعار", command=self.review_prices).pack(
            side="left", padx=(0, 6)
        )

        self.headline_var = tk.StringVar()  # the same counts in one sentence (window title)
        self.note_var = tk.StringVar()
        label(root, textvariable=self.note_var, foreground=GREY, wraplength=1150).pack(
            side="bottom", fill="x", pady=(4, 0)
        )
        cards = ttk.Frame(root)
        cards.pack(fill="x", pady=(10, 2))
        self.cards: dict[str, tk.StringVar] = {}
        self.card_titles: dict[str, tk.StringVar] = {}
        for key, title_text, (bg, fg) in CARDS:
            card = tk.Frame(cards, bg=bg, padx=12, pady=6, cursor="hand2")
            card.pack(side="right", fill="x", expand=True, padx=4)
            number = tk.StringVar(value="—")
            self.cards[key] = number
            big = tk.Label(card, textvariable=number, bg=bg, fg=fg, font=("Segoe UI", 20, "bold"))
            self.card_titles[key] = tk.StringVar(value=title_text)
            small = tk.Label(
                card, textvariable=self.card_titles[key], bg=bg, fg=fg, font=("Segoe UI", 11)
            )
            big.pack()
            small.pack()
            for widget in (card, big, small):
                widget.bind("<Button-1>", partial(self._open_card, key))

        # four lists, one tab each: the pairs to confirm, the close amounts, and
        # what each side has alone
        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill="both", expand=True, pady=(6, 0))
        self.pair_trees: dict[str, ttk.Treeview] = {}
        self.pair_frames: dict[str, ttk.Frame] = {}
        self.sorted_by: dict[str, tuple[str, bool]] = {}  # tree -> (column, descending)
        self.undo_buttons: list[ttk.Button] = []
        self.match_buttons: dict[str, ttk.Button] = {}
        # ✓ / ✗ in order (the last is undone first), saved on this computer after
        # each one; read back the first time the two files are compared
        self.decisions: list[memory.Decision] = []
        self.restore = True
        self.close_total = tk.StringVar()
        for kind in ("possible", "close"):
            frame = ttk.Frame(self.tabs, padding=6)
            self.tabs.add(frame, text="")
            tree = self._tree(frame, PAIR_COLUMNS, PAIR_WIDTHS, stretch=(2, 5, 7))
            tree.bind("<Double-1>", partial(self._double_click, kind))
            buttons = ttk.Frame(frame)  # at the bottom: never pushed off a small screen
            buttons.pack(side="bottom", fill="x", pady=(6, 0))
            match = ttk.Button(
                buttons,
                text="✓ مطابقة",
                style="Big.TButton",
                command=partial(self.decide, kind, True),
            )
            match.pack(side="right")
            self.match_buttons[kind] = match
            ttk.Button(buttons, text="✗ ليست هي", command=partial(self.decide, kind, False)).pack(
                side="right", padx=8
            )
            undo = ttk.Button(buttons, text="↶ تراجع عن آخر قرار", command=self.undo)
            undo.pack(side="left")
            self.undo_buttons.append(undo)
            label(
                buttons,
                "اختر سطرًا ثم «مطابقة» (أو اضغط عليه مرتين) — يختفي وينتقل إلى المتطابقة",
                foreground=GREY,
            ).pack(side="right", padx=8)
            if kind == "close":  # the reason is always "a different amount": not shown
                tree.configure(displaycolumns=[f"c{k}" for k in range(6, -1, -1)])
                label(frame, textvariable=self.close_total, font=BIG).pack(side="bottom", fill="x")
            self._details(frame, tree)
            self._list_last(tree)
            self.pair_trees[kind] = tree
            self.pair_frames[kind] = frame
        self.frames: list[ttk.Frame] = []
        self.trees: list[ttk.Treeview] = []
        self.totals: list[tk.StringVar] = []
        for i in range(2):
            frame = ttk.Frame(self.tabs, padding=6)
            self.tabs.add(frame, text="")
            tree = self._tree(frame, COLUMNS, WIDTHS, stretch=(3, 5))
            buttons = ttk.Frame(frame)  # at the bottom: never pushed off a small screen
            buttons.pack(side="bottom", fill="x", pady=(6, 0))
            ttk.Button(
                buttons, text="✔ تمت تسويتها", style="Big.TButton", command=partial(self.settle, i)
            ).pack(side="right")
            undo = ttk.Button(buttons, text="↶ تراجع عن آخر قرار", command=self.undo)
            undo.pack(side="left")
            self.undo_buttons.append(undo)
            label(
                buttons,
                "صُحّحت في RIM-Soft؟ اخترها (أو عدة أسطر) ثم «تمت تسويتها» — تختفي من القائمة",
                foreground=GREY,
            ).pack(side="right", padx=8)
            total = tk.StringVar()
            label(frame, textvariable=total, foreground=GREY).pack(side="bottom", fill="x")
            same = tk.StringVar()  # why alone: the same amount on the other side
            label(frame, textvariable=same, foreground=ORANGE, wraplength=900).pack(
                side="bottom", fill="x"
            )
            tree.bind("<<TreeviewSelect>>", partial(self._show_same_amount, i, same), add="+")
            self._details(frame, tree)
            self._list_last(tree)
            self.frames.append(frame)
            self.trees.append(tree)
            self.totals.append(total)
        # tabs right to left in the order of the cards above them
        for frame in (self.pair_frames["possible"], self.pair_frames["close"], *self.frames):
            self.tabs.insert(0, frame)
        self.refresh_buttons()

    @staticmethod
    def _list_last(tree: ttk.Treeview) -> None:
        """The list takes what is left once the buttons and lines under it are
        placed: on a small screen the list shrinks, the buttons stay visible."""
        holder = tree.master
        if isinstance(holder, ttk.Frame):
            holder.pack_forget()
            holder.pack(side="top", fill="both", expand=True)

    @staticmethod
    def _details(frame: ttk.Frame, tree: ttk.Treeview) -> None:
        """The selected line in full under the list (descriptions are cut)."""
        text = tk.StringVar()
        label(frame, textvariable=text, wraplength=1100).pack(side="bottom", fill="x", pady=(4, 0))

        def show(_event: object = None) -> None:
            selected = tree.selection()
            values = tree.item(selected[0], "values")[::-1] if selected else ()
            text.set("  |  ".join(str(v) for v in values if str(v).strip()))

        tree.bind("<<TreeviewSelect>>", show, add="+")

    def _open_card(self, key: str, _event: object = None) -> None:
        """A click on a card opens its list (the matched ones have none)."""
        frames = {
            "possible": self.pair_frames.get("possible"),
            "close": self.pair_frames.get("close"),
            "A": self.frames[0] if self.frames else None,
            "B": self.frames[1] if len(self.frames) > 1 else None,
        }
        if (frame := frames.get(key)) is not None:
            self.tabs.select(frame)  # type: ignore[no-untyped-call]

    def _reopen_last_files(self) -> None:
        """The two files of last time are loaded again (another can be chosen)."""
        for box in (self.box_a, self.box_b):
            if box.loaded is None and (last := memory.last_file(box.key)) is not None:
                box.load(last, again=True)

    def _show_same_amount(self, i: int, text: tk.StringVar, _event: object = None) -> None:
        """Under a line alone on side i: every line of the other file with the
        same amount, and what became of it."""
        selected = self.trees[i].selection()
        if self.result is None or self.sources is None or not selected:
            text.set("")
            return
        line = next((x for x in self.result.only[i] if x.key == selected[0]), None)
        if line is None:
            text.set("")
            return
        mine, others = self.sources[i].lines, self.sources[1 - i].lines
        text.set(same_amount(self.result, line, mine, others, i))

    def settle(self, i: int) -> None:
        """✔ the selected lines of side i were settled by hand (corrected in
        RIM-Soft): set aside, saved like the other decisions."""
        if self.result is None:
            return
        by_key = {x.key: x for x in self.result.only[i]}
        chosen = [by_key[k] for k in self.trees[i].selection() if k in by_key]
        if not chosen:
            return
        for x in chosen:
            pair = Pair(Kind.SETTLED, (x,), ()) if i == 0 else Pair(Kind.SETTLED, (), (x,))
            self.decisions.append((True, pair))
        self._save()
        self.run()

    def undo(self) -> None:
        """Take back the last ✓ / ✗ (the pair is proposed again)."""
        if not self.decisions:
            return
        self.decisions.pop()
        self._save()
        self.run()

    def _save(self) -> None:
        if self.result is not None:
            with contextlib.suppress(OSError):
                memory.save(self.result.names, self.decisions)

    def _double_click(self, kind: str, _event: object) -> None:
        self.decide(kind, True)

    def _tree(
        self,
        frame: ttk.Frame,
        columns: tuple[str, ...],
        widths: tuple[int, ...],
        stretch: tuple[int, ...],
    ) -> ttk.Treeview:
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        ids = [f"c{k}" for k in range(len(columns))]
        tree = ttk.Treeview(holder, columns=ids[::-1], show="headings", height=6)
        for k, (heading, width) in enumerate(zip(columns, widths, strict=True)):
            tree.heading(
                f"c{k}", text=heading, anchor="e", command=partial(self.sort, tree, f"c{k}")
            )
            tree.column(f"c{k}", width=width, anchor="e", stretch=k in stretch)
        bar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        bar.pack(side="left", fill="y")
        tree.pack(side="right", fill="both", expand=True)
        return tree

    def sort(self, tree: ttk.Treeview, column: str, again: bool = False) -> None:
        """A click on a column title sorts by it; a second click, the other way.
        Dates by date, amounts by value, the rest alphabetically."""
        last = self.sorted_by.get(str(tree))
        if again:
            if last is None:
                return
            column, descending = last
        else:
            descending = last == (column, False) if last else False
            if last is None or last[0] != column:
                descending = self._numeric(tree, column)  # amounts: largest first
        self.sorted_by[str(tree)] = (column, descending)
        items = list(tree.get_children())
        items.sort(key=lambda i: _sort_key(tree.set(i, column)), reverse=descending)
        for n, item in enumerate(items):
            tree.move(item, "", n)
        arrow = " ▼" if descending else " ▲"
        for c in tree["columns"]:
            text = str(tree.heading(c, "text")).removesuffix(" ▲").removesuffix(" ▼")
            tree.heading(c, text=text + (arrow if c == column else ""))

    @staticmethod
    def _numeric(tree: ttk.Treeview, column: str) -> bool:
        first = next(iter(tree.get_children()), None)
        return first is not None and _sort_key(tree.set(first, column))[0] == 1

    def _fit(self, width: int, height: int) -> None:
        """Never larger than the screen: on Windows the window opens maximized
        (the taskbar and a 125-150 % display zoom are then taken into account)."""
        w = min(width, self.winfo_screenwidth() - 40)
        h = min(height, self.winfo_screenheight() - 90)
        self.geometry(f"{w}x{h}+10+10")
        if sys.platform.startswith("win"):
            with contextlib.suppress(tk.TclError):
                self.state("zoomed")

    def propose_period(self) -> None:
        """Each time a file is chosen: the days both files cover (editable)."""
        a, b = self.box_a.loaded, self.box_b.loaded
        period = common_period(a.lines, b.lines) if a and b else None
        if period is not None:
            self.from_var.set(f"{period[0]:%d/%m/%Y}")
            self.to_var.set(f"{period[1]:%d/%m/%Y}")

    def refresh_buttons(self) -> None:
        ready = self.box_a.loaded is not None and self.box_b.loaded is not None
        self.compare_button.state(["!disabled"] if ready else ["disabled"])
        self.manual_button.state(["!disabled"] if self.result is not None else ["disabled"])
        self.party_button.state(["!disabled"] if self.result is not None else ["disabled"])
        for button in self.undo_buttons:
            button.state(["!disabled"] if self.decisions else ["disabled"])

    def run(self) -> None:
        a, b = self.box_a.loaded, self.box_b.loaded
        if a is None or b is None:
            return
        start, end = _date(self.from_var.get()), _date(self.to_var.get())
        if start is None or end is None or start > end:
            self.dialogs.error("الفترة", "اكتب التاريخين بهذا الشكل: 01/09/2026")
            return
        if a.data == b.data:
            self.dialogs.error("الملف نفسه", "اخترت الملف نفسه للطرفين — اختر ملف كل طرف")
            return
        self.sources = (a, b)  # the files this result was made from
        names = (self.box_a.name, self.box_b.name)
        restored = 0
        if self.restore:  # files just chosen: the decisions saved for them
            self.restore = False
            self.decisions = memory.load(names, a.lines, b.lines)
            restored = len(self.decisions)
        self.result = compare(
            a.lines,
            b.lines,
            names,
            (start, end),
            confirmed=[pair for ok, pair in self.decisions if ok],
            rejected={pair.key for ok, pair in self.decisions if not ok},
        )
        self.show(self.result)
        if restored:
            self.note_var.set(
                f"💾 استُرجع {restored} قرارًا محفوظًا من المرة السابقة\n" + self.note_var.get()
            )
        self.refresh_buttons()

    def decide(self, kind: str, accepted: bool) -> None:
        """✓ the pair is the same operation (it disappears into the matches);
        ✗ it is not (its lines go back to the lists, never proposed together
        again). Kept for this session only."""
        if self.result is None:
            return
        tree = self.pair_trees[kind]
        selected = tree.selection()
        if not selected:
            return
        pairs = {p.key: p for p in (*self.result.possible, *self.result.close)}
        pair = pairs.get(selected[0])
        if pair is None:
            return
        self.decisions.append((accepted, pair))
        self._save()
        self.run()

    def show(self, result: Result) -> None:
        self.headline_var.set(headline(result))
        self.title(f"{headline(result)} — BinaaCompare")
        for key, count in (
            ("matched", result.matched_count),
            ("possible", len(result.possible)),
            ("close", len(result.close)),
            ("A", len(result.only[0])),
            ("B", len(result.only[1])),
        ):
            self.cards[key].set(str(count))
        a_ar, b_ar = (arabic(n) for n in result.names)
        self.card_titles["A"].set(f"عند {a_ar} فقط")
        self.card_titles["B"].set(f"عند {b_ar} فقط")
        notes = [title(result)]
        if extra := footnote(result):
            notes.append(extra)
        counts = [len(box.loaded.lines) for box in (self.box_a, self.box_b) if box.loaded]
        if len(counts) == 2 and max(counts) > 2 * min(counts):
            big = self.box_a if counts[0] > counts[1] else self.box_b
            notes.append(
                f"⚠️ عدد العمليات مختلف كثيرًا ({counts[0]} مقابل {counts[1]}) — "
                f"تأكد أن ملف {big.name} هو كشف الحساب بين الطرفين فقط ولنفس الفترة"
            )
        for box in (self.box_a, self.box_b):
            if box.loaded is not None and box.loaded.problems:
                rows = "، ".join(p.split(":")[0].replace("السطر ", "") for p in box.loaded.problems)
                notes.append(f"⚠️ سطور لم تُقرأ في ملف {box.name}: {rows}")
        self.note_var.set(arabic_text("   •   ".join(notes), result.names))
        a, b = result.names
        by_gap = sorted(result.close, key=lambda p: (abs(p.difference), p.day))
        gaps = sum(abs(p.difference) for p in result.close)
        self.close_total.set(
            f"{len(result.close)} عملية — مجموع الفروق: {amount_text(gaps)}" if result.close else ""
        )
        for kind, pairs, text in (
            ("possible", result.possible, "🔎 قد تكون نفس العملية"),
            ("close", by_gap, "💲 فرق في المبلغ"),
        ):
            self.tabs.tab(self.pair_frames[kind], text=f"{text} ({len(pairs)})")  # type: ignore[no-untyped-call]
            tree = self.pair_trees[kind]
            tree.delete(*tree.get_children())
            for pair in pairs:
                da, ma, la = side_cells(pair.a)
                db, mb, lb = side_cells(pair.b)
                gap = amount_text(abs(pair.difference)) if pair.a and pair.b else ""
                row = (da, ma, la, db, mb, lb, gap, reason(pair, result.names))
                tree.insert("", "end", iid=pair.key, values=row[::-1], tags=_stripe(tree))
            for k, name in ((0, a), (1, a), (2, a), (3, b), (4, b), (5, b)):
                tree.heading(f"c{k}", text=f"{PAIR_COLUMNS[k]} {arabic(name)}", anchor="e")
        for i in range(2):
            lines = result.only[i]
            alone = f"❌ عند {arabic(result.names[i])} فقط ({len(lines)})"
            self.tabs.tab(self.frames[i], text=alone)  # type: ignore[no-untyped-call]
            tree = self.trees[i]
            tree.delete(*tree.get_children())
            for x in lines:
                values = (
                    f"{x.day:%d/%m/%Y}",
                    amount_text(x.amount),
                    SIDE_AR[x.side],
                    x.label,
                    x.ref,
                    arabic_text(result.notes.get(x.key, ""), result.names),
                )
                tree.insert("", "end", iid=x.key, values=values[::-1], tags=_stripe(tree))
            self.totals[i].set(total_text(result, i))
            shown = [k for k in range(len(COLUMNS)) if k != 4 or any(x.ref for x in lines)]
            tree.configure(displaycolumns=[f"c{k}" for k in reversed(shown)])  # no empty «المرجع»
        for tree in (*self.pair_trees.values(), *self.trees):
            self.sort(tree, "", again=True)  # keep the order the user chose
        current = self.tabs.select()  # type: ignore[no-untyped-call]
        empty = {
            str(f)
            for f, t in zip(self._tab_frames(), self._tab_trees(), strict=True)
            if not t.get_children()
        }
        if current in empty:  # an empty list open: show the next one with lines
            for frame, tree in zip(self._tab_frames(), self._tab_trees(), strict=True):
                if tree.get_children():
                    self.tabs.select(frame)  # type: ignore[no-untyped-call]
                    break

    def _tab_frames(self) -> list[ttk.Frame]:
        return [self.pair_frames["possible"], self.pair_frames["close"], *self.frames]

    def _tab_trees(self) -> list[ttk.Treeview]:
        return [self.pair_trees["possible"], self.pair_trees["close"], *self.trees]

    def review_sales(self) -> ReviewWindow | None:
        """The day's sales report (PDF from RIM-Soft): what looks like an entry
        error, the most serious first. Works without the two files."""
        path = self.dialogs.ask_open_pdf()
        if path is None:
            return None
        self.configure(cursor="watch")
        self.update_idletasks()
        history = review.History.from_json(memory.sales_history())
        try:
            result = review.review_pdf(path, history)
        except Exception:
            self.dialogs.error("تعذّرت القراءة", "تأكد أنه تقرير المبيعات من RIM-Soft بصيغة PDF")
            return None
        finally:
            self.configure(cursor="")
        if not result.sales:
            self.dialogs.error("لا توجد فواتير", "لم يُعثر في الملف على أي فاتورة")
            return None
        history.learn(result.sales)  # each report reviewed teaches the usual entries
        with contextlib.suppress(OSError):
            memory.save_sales_history(history.to_json())
        return ReviewWindow(self, path.name, result)

    def review_prices(self) -> PricesWindow | None:
        """The product lines of a day (copied from RIM-Soft): a price far from
        the product's usual one. The prices are learned on this computer."""
        path = self.dialogs.ask_open_prices()
        if path is None:
            return None
        history = prices.PriceHistory.from_json(memory.price_history())
        try:
            lines = prices.read_lines(path)
        except Exception:
            self.dialogs.error("تعذّرت القراءة", "يجب أن يحتوي الملف على العمودين Code و PrixVente")
            return None
        found = prices.check_prices(lines, history)
        history.learn(lines)
        with contextlib.suppress(OSError):
            memory.save_price_history(history.to_json())
        return PricesWindow(self, path.name, lines, found)

    def export_for_party(self) -> Path | None:
        """A short file for the other party: only what concerns them."""
        if self.result is None:
            return None
        other = arabic(self.result.names[1])
        target = self.dialogs.ask_save(f"ملف مشترك مع {other}.xlsx")
        if target is None:
            return None
        try:
            for_party.save(self.result, target)
        except OSError:
            self.dialogs.error("تعذّر الحفظ", "أغلق الملف إن كان مفتوحًا في Excel ثم أعد المحاولة")
            return None
        self.dialogs.open_file(target)
        return target

    def export_manual(self) -> Path | None:
        """The two original files without the matched operations, with a status
        on each remaining line, to compare by hand together."""
        if self.result is None or self.sources is None:
            return None
        a, b = self.result.names
        target = self.dialogs.ask_save(f"للمقارنة اليدوية {a} {b}.xlsx")
        if target is None:
            return None
        try:
            manual.save(self.result, *self.sources, target)
        except OSError:
            self.dialogs.error("تعذّر الحفظ", "أغلق الملف إن كان مفتوحًا في Excel ثم أعد المحاولة")
            return None
        self.dialogs.open_file(target)
        return target


class ReviewWindow(tk.Toplevel):
    """The checks of one sales report: one line per suspicious entry."""

    COLORS: ClassVar[dict[str, tuple[str, str]]] = {
        review.RED: ("#FDECEA", "🔴"),
        review.YELLOW: ("#FFF8E1", "🟡"),
        review.GREY: ("#F0F0F0", "⚪"),
    }

    def __init__(self, app: CompareApp, name: str, result: review.Review) -> None:
        super().__init__(app)
        self.title(f"مراجعة المبيعات — {name}")
        self.geometry("1100x620")
        self.result = result
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        days = sorted({s.day for s in result.sales})
        span = f"{days[0]:%d/%m/%Y}" + (f" → {days[-1]:%d/%m/%Y}" if len(days) > 1 else "")
        red, yellow = result.count(review.RED), result.count(review.YELLOW)
        if not red and not yellow:
            head = f"✅ {len(result.sales)} فاتورة ({span}) — لا شيء مشبوه"
        else:
            head = f"{len(result.sales)} فاتورة ({span}) — 🔴 {red} خطأ محتمل — 🟡 {yellow} للتحقق"
        self.headline = label(frame, head, font=BIG)
        self.headline.pack(fill="x")
        extra = []
        if result.missing:
            numbers = "، ".join(f"f {n}" for n in result.missing[:15])
            extra.append(f"⚠️ أرقام فواتير غير موجودة في التقرير: {numbers}")
        if result.unread:
            extra.append(f"⚠️ {result.unread} سطر لم يُقرأ")
        self.extra = label(frame, "\n".join(extra), foreground=RED, wraplength=1050)
        self.extra.pack(fill="x", pady=(4, 6))
        columns = ("الرقم", "التاريخ", "البيان", "المبلغ", "الملاحظة")
        ids = [f"c{k}" for k in range(len(columns))]
        holder = ttk.Frame(frame)
        holder.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(holder, columns=ids[::-1], show="headings")
        for k, (heading, width) in enumerate(zip(columns, (100, 95, 330, 110, 430), strict=True)):
            self.tree.heading(f"c{k}", text=heading, anchor="e")
            self.tree.column(f"c{k}", width=width, anchor="e", stretch=k in (2, 4))
        bar = ttk.Scrollbar(holder, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=bar.set)
        bar.pack(side="left", fill="y")
        self.tree.pack(side="right", fill="both", expand=True)
        for level, (bg, _) in self.COLORS.items():
            self.tree.tag_configure(level, background=bg)
        for k, f in enumerate(result.findings):
            s = f.sale
            values = (
                f"{self.COLORS[f.level][1]} {s.number}",
                f"{s.day:%d/%m/%Y}",
                s.label,
                amount_text(s.amount * 100) if s.amount else "—",
                f.text,
            )
            self.tree.insert("", "end", iid=str(k), values=values[::-1], tags=(f.level,))


class PricesWindow(tk.Toplevel):
    """The prices far from the usual ones, with their line in the file."""

    def __init__(
        self,
        app: CompareApp,
        name: str,
        lines: list[prices.PriceLine],
        found: list[prices.PriceFinding],
    ) -> None:
        super().__init__(app)
        self.title(f"مراجعة الأسعار — {name}")
        self.geometry("900x520")
        frame = ttk.Frame(self, padding=12)
        frame.pack(fill="both", expand=True)
        if found:
            head = f"{len(lines)} سطر — 🔴 {len(found)} سعر غير معتاد"
        else:
            head = f"✅ {len(lines)} سطر — كل الأسعار معتادة"
        self.headline = label(frame, head, font=BIG)
        self.headline.pack(fill="x")
        label(
            frame,
            "السعر المعتاد يُتعلَّم من الملفات السابقة: كلما راجعت أكثر صار أدق",
            foreground=GREY,
        ).pack(fill="x", pady=(2, 6))
        columns = ("السطر في الملف", "الرمز", "المنتج", "الكمية", "الملاحظة")
        ids = [f"c{k}" for k in range(len(columns))]
        self.tree = ttk.Treeview(frame, columns=ids[::-1], show="headings")
        for k, (heading, width) in enumerate(zip(columns, (110, 80, 220, 80, 380), strict=True)):
            self.tree.heading(f"c{k}", text=heading, anchor="e")
            self.tree.column(f"c{k}", width=width, anchor="e", stretch=k in (2, 4))
        self.tree.pack(fill="both", expand=True)
        self.tree.tag_configure("red", background="#FDECEA")
        for k, f in enumerate(found):
            x = f.line
            values = (str(x.row), x.code, x.name, f"{x.quantity:,.0f}", f.text)
            self.tree.insert("", "end", iid=str(k), values=values[::-1], tags=("red",))


def _stripe(tree: ttk.Treeview) -> tuple[str, ...]:
    """Every other line lightly shaded: easier to follow a line across."""
    tree.tag_configure("stripe", background=STRIPE)
    return ("stripe",) if len(tree.get_children()) % 2 else ()


def _modern_theme(root: tk.Tk) -> bool:
    """The Windows 11 look (Sun Valley theme, shipped with the program, offline);
    the plain look when it is not there."""
    try:
        import sv_ttk

        sv_ttk.set_theme("light", root)
    except Exception:
        return False
    return True


def _sort_key(text: str) -> tuple[int, object]:
    """(0, date) / (1, amount) / (2, text): dates, then amounts, then words."""
    text = text.strip()
    if parsed := _date(text[:10]):
        return 0, parsed
    digits = text.replace(",", "").replace(" ", "")
    parts = digits.split("+")
    if parts and all(p.isdigit() for p in parts):
        return 1, sum(int(p) for p in parts)
    return 2, text.casefold()


def _date(text: str) -> date | None:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        with contextlib.suppress(ValueError):
            return datetime.strptime(text.strip(), fmt).date()
    return None


def main() -> None:
    CompareApp().mainloop()


__all__ = ["CompareApp", "Dialogs", "main"]
