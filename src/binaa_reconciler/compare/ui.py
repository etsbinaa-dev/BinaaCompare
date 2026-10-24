"""BinaaCompare window: two files, one button, two lists.

Everything happens on this one screen: choose the two RIM-Soft files, press
«قارن», read the two lists, export to Excel. Nothing is saved between runs.
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

from binaa_reconciler.compare import manual
from binaa_reconciler.compare.engine import (
    SIDE_AR,
    Pair,
    Result,
    amount_text,
    common_period,
    compare,
)
from binaa_reconciler.compare.files import Loaded, load
from binaa_reconciler.compare.texts import (
    footnote,
    headline,
    reason,
    side_cells,
    title,
    total_text,
)

FONT = ("Segoe UI", 11)
BIG = ("Segoe UI", 13, "bold")
GREY, ORANGE, GREEN, RED = "#555555", "#b45309", "#2e7d32", "#a01010"
COLUMNS = ("التاريخ", "المبلغ", "الجهة", "الوصف", "المرجع", "ملاحظة")
WIDTHS = (90, 110, 55, 300, 110, 330)
PAIR_COLUMNS = ("تاريخ", "مبلغ", "وصف", "تاريخ", "مبلغ", "وصف", "الفرق", "السبب")
PAIR_WIDTHS = (90, 120, 200, 90, 120, 200, 90, 300)
PRIVACY = "🔒 المعالجة على هذا الجهاز فقط — لا يُرسل أي ملف إلى الإنترنت ولا يُحفظ شيء"


class Dialogs:
    """The system dialogs (replaced in tests)."""

    def ask_open(self) -> Path | None:
        name = filedialog.askopenfilename(
            title="اختر ملف RIM-Soft", filetypes=[("Excel", "*.xlsx *.xlsm"), ("الكل", "*.*")]
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
        path = self.app.dialogs.ask_open()
        if path is not None:
            self.load(path)

    def load(self, path: Path) -> None:
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
            if self.loaded.merged:
                text += f" — {self.loaded.merged} عملية مقسومة على سطرين دُمجت"
            if self.loaded.problems:
                text += f" — ⚠️ {len(self.loaded.problems)} سطر لم يُقرأ"
            self.status_var.set(text)
            self.status.configure(foreground=ORANGE if self.loaded.problems else GREEN)
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
        style = ttk.Style(self)
        style.configure("Big.TButton", font=BIG, padding=(18, 8))

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

        actions = ttk.Frame(root)
        actions.pack(fill="x")
        self.compare_button = ttk.Button(
            actions, text="قارن", style="Big.TButton", command=self.run
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

        self.headline_var = tk.StringVar()
        self.note_var = tk.StringVar()
        self.headline = label(root, textvariable=self.headline_var, font=BIG, wraplength=1100)
        self.headline.pack(fill="x", pady=(6, 0))
        label(root, textvariable=self.note_var, foreground=GREY, wraplength=1100).pack(fill="x")

        # four lists, one tab each: the pairs to confirm, the close amounts, and
        # what each side has alone
        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill="both", expand=True, pady=(6, 0))
        self.pair_trees: dict[str, ttk.Treeview] = {}
        self.pair_frames: dict[str, ttk.Frame] = {}
        self.undo_buttons: list[ttk.Button] = []
        self.history: list[tuple[bool, str]] = []  # (confirmed?, pair key), last last
        self.close_total = tk.StringVar()
        for kind in ("possible", "close"):
            frame = ttk.Frame(self.tabs, padding=6)
            self.tabs.add(frame, text="")
            tree = self._tree(frame, PAIR_COLUMNS, PAIR_WIDTHS, stretch=(2, 5, 7))
            tree.bind("<Double-1>", partial(self._double_click, kind))
            buttons = ttk.Frame(frame)
            buttons.pack(fill="x", pady=(6, 0))
            ttk.Button(
                buttons,
                text="✓ مطابقة",
                style="Big.TButton",
                command=partial(self.decide, kind, True),
            ).pack(side="right")
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
            self._details(frame, tree)
            if kind == "close":
                label(frame, textvariable=self.close_total, font=BIG).pack(fill="x")
            self.pair_trees[kind] = tree
            self.pair_frames[kind] = frame
        self.frames: list[ttk.Frame] = []
        self.trees: list[ttk.Treeview] = []
        self.totals: list[tk.StringVar] = []
        for _ in range(2):
            frame = ttk.Frame(self.tabs, padding=6)
            self.tabs.add(frame, text="")
            tree = self._tree(frame, COLUMNS, WIDTHS, stretch=(3, 5))
            self._details(frame, tree)
            total = tk.StringVar()
            label(frame, textvariable=total, foreground=GREY).pack(fill="x")
            self.frames.append(frame)
            self.trees.append(tree)
            self.totals.append(total)
        self.confirmed: dict[str, Pair] = {}
        self.rejected: set[str] = set()
        self.refresh_buttons()

    @staticmethod
    def _details(frame: ttk.Frame, tree: ttk.Treeview) -> None:
        """The selected line in full under the list (descriptions are cut)."""
        text = tk.StringVar()
        label(frame, textvariable=text, wraplength=1100).pack(fill="x", pady=(4, 0))

        def show(_event: object = None) -> None:
            selected = tree.selection()
            values = tree.item(selected[0], "values")[::-1] if selected else ()
            text.set("  |  ".join(str(v) for v in values if str(v).strip()))

        tree.bind("<<TreeviewSelect>>", show, add="+")

    def undo(self) -> None:
        """Take back the last ✓ / ✗ (the pair is proposed again)."""
        if not self.history:
            return
        confirmed, key = self.history.pop()
        if confirmed:
            self.confirmed.pop(key, None)
        else:
            self.rejected.discard(key)
        self.run()

    def _double_click(self, kind: str, _event: object) -> None:
        self.decide(kind, True)

    @staticmethod
    def _tree(
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
            tree.heading(f"c{k}", text=heading, anchor="e")
            tree.column(f"c{k}", width=width, anchor="e", stretch=k in stretch)
        bar = ttk.Scrollbar(holder, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=bar.set)
        bar.pack(side="left", fill="y")
        tree.pack(side="right", fill="both", expand=True)
        return tree

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
        for button in self.undo_buttons:
            button.state(["!disabled"] if self.history else ["disabled"])

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
        self.result = compare(
            a.lines,
            b.lines,
            (self.box_a.name, self.box_b.name),
            (start, end),
            confirmed=list(self.confirmed.values()),
            rejected=self.rejected,
        )
        self.show(self.result)
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
        if accepted:
            self.confirmed[pair.key] = pair
        else:
            self.rejected.add(pair.key)
        self.history.append((accepted, pair.key))
        self.run()

    def show(self, result: Result) -> None:
        self.headline_var.set(headline(result))
        clean = not any((*result.only, result.possible, result.close))
        self.headline.configure(foreground=GREEN if clean else ORANGE)
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
        self.note_var.set("\n".join(notes))
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
                tree.insert("", "end", iid=pair.key, values=row[::-1])
            for k, name in ((0, a), (1, a), (2, a), (3, b), (4, b), (5, b)):
                tree.heading(f"c{k}", text=f"{PAIR_COLUMNS[k]} {name}", anchor="e")
        for i in range(2):
            lines = result.only[i]
            self.tabs.tab(self.frames[i], text=f"❌ عند {result.names[i]} فقط ({len(lines)})")  # type: ignore[no-untyped-call]
            tree = self.trees[i]
            tree.delete(*tree.get_children())
            for x in lines:
                values = (
                    f"{x.day:%d/%m/%Y}",
                    amount_text(x.amount),
                    SIDE_AR[x.side],
                    x.label,
                    x.ref,
                    result.notes.get(x.key, ""),
                )
                tree.insert("", "end", iid=x.key, values=values[::-1])
            self.totals[i].set(total_text(result, i))

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


def _date(text: str) -> date | None:
    for fmt in ("%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y"):
        with contextlib.suppress(ValueError):
            return datetime.strptime(text.strip(), fmt).date()
    return None


def main() -> None:
    CompareApp().mainloop()


__all__ = ["CompareApp", "Dialogs", "main"]
