"""RIM-Soft XLSX ledger importer.

Privacy and safety:
* The file is read once into memory and parsed from there: no copy is written,
  nothing is cached, and the file handle is released immediately (important on
  Windows, where an open handle locks the file for Excel).
* External links are not loaded (`keep_links=False`); formulas are never
  evaluated. A formula cell is used only through the value Excel cached when
  the file was saved. A formula without a cached value, or whose value is an
  Excel error (#DIV/0!, #REF!...), is reported instead of being read as empty
  or zero: the workbook is read twice (values and formulas) to tell them apart.
* No logging. Issues carry codes and row ids only.

Pipeline: read workbook -> choose sheet -> detect header (or use a manual
`ColumnMapping`) -> infer the date order of the whole column -> classify every
row -> build transactions -> flag duplicates -> report.
"""

from __future__ import annotations

import hashlib
import re
from collections.abc import Sequence
from dataclasses import dataclass, field
from datetime import UTC, date, datetime
from io import BytesIO
from pathlib import Path

from openpyxl import load_workbook

from binaa_reconciler.core.dates import (
    DateOrder,
    DateParseError,
    excel_serial_to_datetime,
    infer_date_order,
    parse_datetime,
)
from binaa_reconciler.core.fingerprint import occurrence_indexes, rimsoft_fingerprint
from binaa_reconciler.core.ledger_columns import ColumnRole
from binaa_reconciler.core.models import (
    ImportBatch,
    IssueCode,
    LedgerSide,
    RowClass,
    RowRef,
    Severity,
    Source,
    Transaction,
    ValidationIssue,
)
from binaa_reconciler.importers.rimsoft_xlsx.header import (
    ColumnMapping,
    HeaderDetection,
    detect_header,
)
from binaa_reconciler.importers.rimsoft_xlsx.report import ImportReport, RowRecord
from binaa_reconciler.normalize.phones import extract_mr_phones
from binaa_reconciler.normalize.text import normalize_text, to_ascii_digits
from binaa_reconciler.validation.duplicates import find_duplicate_groups
from binaa_reconciler.validation.row_rules import (
    LedgerRowResult,
    RowReason,
    cell_text,
    classify_ledger_row,
    is_blank,
)

MAX_HEADER_SCAN_ROWS = 30
# Plausible Excel serial dates for a ledger: 1954-10-04 .. 2119-01-10.
EXCEL_SERIAL_MIN = 20_000
EXCEL_SERIAL_MAX = 80_000

_DATE_TEXT_RE = re.compile(
    r"^(\d{1,4})[/\-.](\d{1,2})[/\-.](\d{1,4})(?:\s+\d{1,2}[:hH]\d{2}(?::\d{2})?)?$"
)

Row = tuple[object, ...]
CellKey = tuple[int, int]  # (1-based row number, 0-based column index)

_EXCEL_ERRORS = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA",
     "#SPILL!", "#CALC!"}
)  # fmt: skip


@dataclass(frozen=True)
class _SheetData:
    title: str
    rows: list[Row]
    # Cells whose value cannot be trusted: formula without cached value, or error value.
    unreliable: dict[CellKey, IssueCode]


@dataclass(frozen=True)
class RimsoftImport:
    batch: ImportBatch
    transactions: tuple[Transaction, ...]
    report: ImportReport
    mapping: ColumnMapping | None
    detection: HeaderDetection | None = field(default=None, repr=False)

    @property
    def issues(self) -> tuple[ValidationIssue, ...]:
        return self.report.issues

    @property
    def is_blocked(self) -> bool:
        return self.report.is_blocked


# --------------------------------------------------------------------------- dates


def _date_text(value: str) -> str:
    return " ".join(to_ascii_digits(value).split())


def infer_column_date_order(values: Sequence[object]) -> tuple[DateOrder | None, bool]:
    """(order, needed) for the text dates of a column.

    `needed` is False when no text date requires a day/month order (native
    Excel dates, serials and ISO "yyyy-mm-dd" texts are unambiguous).
    """
    candidates: list[str] = []
    for value in values:
        if not isinstance(value, str):
            continue
        m = _DATE_TEXT_RE.match(_date_text(value))
        if m and len(m.group(1)) != 4:
            candidates.append(m.group(0).split(" ")[0])
    if not candidates:
        return None, False
    return infer_date_order(candidates), True


def parse_date_cell(value: object, order: DateOrder | None) -> date:
    """Parse one date cell. Naive values are Mauritanian time (UTC+0)."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, bool):
        raise DateParseError("bad_format")
    if isinstance(value, int | float):
        if not EXCEL_SERIAL_MIN <= value <= EXCEL_SERIAL_MAX:
            raise DateParseError("bad_excel_serial")
        return excel_serial_to_datetime(float(value))[0]
    if isinstance(value, str):
        text = _date_text(value)
        if text.isdigit() and len(text) == 5:
            return parse_date_cell(int(text), order)
        m = _DATE_TEXT_RE.match(text)
        if not m:
            raise DateParseError("bad_format")
        if len(m.group(1)) == 4:
            return parse_datetime(text, DateOrder.YMD)[0]
        if order is None:
            raise DateParseError("ambiguous_order")
        return parse_datetime(text, order)[0]
    raise DateParseError("bad_format")


# --------------------------------------------------------------------------- workbook


def _read_workbook(data: bytes) -> list[_SheetData]:
    values_wb = load_workbook(BytesIO(data), read_only=True, data_only=True, keep_links=False)
    formulas_wb = load_workbook(BytesIO(data), read_only=True, data_only=False, keep_links=False)
    try:
        sheets: list[_SheetData] = []
        for ws, fws in zip(values_wb.worksheets, formulas_wb.worksheets, strict=True):
            if not hasattr(ws, "iter_rows") or not hasattr(fws, "iter_rows"):
                continue
            rows: list[Row] = []
            unreliable: dict[CellKey, IssueCode] = {}
            for r, cells in enumerate(ws.iter_rows(), start=1):
                values: list[object] = []
                for c, cell in enumerate(cells):
                    value = cell.value
                    if cell.data_type == "e" or (isinstance(value, str) and value in _EXCEL_ERRORS):
                        unreliable[(r, c)] = IssueCode.FORMULA_ERROR_VALUE
                    values.append(value)
                rows.append(tuple(values))
            for r, cells in enumerate(fws.iter_rows(), start=1):
                for c, cell in enumerate(cells):
                    if cell.data_type != "f":
                        continue
                    cached = rows[r - 1][c] if r <= len(rows) and c < len(rows[r - 1]) else None
                    if cached is None:
                        unreliable.setdefault((r, c), IssueCode.FORMULA_WITHOUT_CACHED_VALUE)
            sheets.append(_SheetData(ws.title, rows, unreliable))
        return sheets
    finally:
        values_wb.close()
        formulas_wb.close()


def _cell(row: Row, index: int | None) -> object:
    if index is None or index >= len(row):
        return None
    return row[index]


def _optional_text(row: Row, mapping: ColumnMapping, role: ColumnRole) -> str | None:
    text = cell_text(_cell(row, mapping.column(role)))
    return text or None


def _sheet_issue(batch_id: str, code: IssueCode) -> ValidationIssue:
    return ValidationIssue(batch_id, Source.RIMSOFT, None, Severity.BLOCKING, code)


# --------------------------------------------------------------------------- main entry


@dataclass
class _Context:
    batch_id: str
    file_name: str
    imported_at: datetime
    digest: str = ""
    sheet_count: int = 0
    sheet_index: int | None = None
    sheet_name: str | None = None
    mapping: ColumnMapping | None = None
    detection: HeaderDetection | None = None
    date_order: DateOrder | None = None

    def build(
        self,
        issues: Sequence[ValidationIssue],
        rows: Sequence[RowRecord] = (),
        transactions: Sequence[Transaction] = (),
    ) -> RimsoftImport:
        report = ImportReport(
            file_sha256=self.digest,
            sheet_index=self.sheet_index,
            sheet_count=self.sheet_count,
            header_row=self.mapping.header_row if self.mapping else None,
            columns=dict(self.mapping.columns) if self.mapping else {},
            date_order=self.date_order,
            rows=tuple(rows),
            issues=tuple(issues),
            sheet_name=self.sheet_name,
        )
        dates = [t.tx_date for t in transactions]
        batch = ImportBatch(
            id=self.batch_id,
            source=Source.RIMSOFT,
            file_name=self.file_name,
            file_sha256=self.digest,
            imported_at=self.imported_at,
            period_start=min(dates) if dates else None,
            period_end=max(dates) if dates else None,
            row_counts={k: v for k, v in report.class_counts.items() if v},
        )
        return RimsoftImport(batch, tuple(transactions), report, self.mapping, self.detection)

    def blocked(self, code: IssueCode) -> RimsoftImport:
        return self.build([_sheet_issue(self.batch_id, code)])


def _choose_sheet(
    sheets: list[tuple[str, list[Row]]], sheet: str | int | None, manual_mapping: bool
) -> int | IssueCode:
    """Index of the sheet to import, or the blocking issue explaining why not."""
    if sheet is not None:
        names = [name for name, _ in sheets]
        if isinstance(sheet, int):
            return sheet if 0 <= sheet < len(sheets) else IssueCode.SHEET_NOT_FOUND
        return names.index(sheet) if sheet in names else IssueCode.SHEET_NOT_FOUND
    if len(sheets) == 1:
        return 0
    if manual_mapping:
        return IssueCode.MULTIPLE_CANDIDATE_SHEETS
    found = [
        i for i, (_, rows) in enumerate(sheets) if detect_header(rows, MAX_HEADER_SCAN_ROWS).found
    ]
    if len(found) == 1:
        return found[0]
    return IssueCode.MULTIPLE_CANDIDATE_SHEETS if found else IssueCode.HEADER_NOT_FOUND


def import_rimsoft_xlsx(
    source: Path | bytes,
    *,
    sheet: str | int | None = None,
    mapping: ColumnMapping | None = None,
    date_order: DateOrder | None = None,
    batch_id: str = "RIMSOFT",
    imported_at: datetime | None = None,
) -> RimsoftImport:
    """Import a RIM-Soft XLSX export.

    `sheet` (name or 0-based index), `mapping` and `date_order` are manual
    overrides for when automatic detection is not certain; the importer never
    guesses them itself.
    """
    ctx = _Context(
        batch_id=batch_id,
        file_name=source.name if isinstance(source, Path) else "",
        imported_at=imported_at or datetime.now(UTC),
    )
    try:
        data = source.read_bytes() if isinstance(source, Path) else source
    except OSError:
        return ctx.blocked(IssueCode.FILE_UNREADABLE)
    ctx.digest = hashlib.sha256(data).hexdigest()
    try:
        sheets = _read_workbook(data)
    except Exception:  # corrupted file or not an XLSX: report it, never crash
        return ctx.blocked(IssueCode.FILE_UNREADABLE)
    if not sheets:
        return ctx.blocked(IssueCode.FILE_UNREADABLE)
    ctx.sheet_count = len(sheets)

    chosen = _choose_sheet([(s.title, s.rows) for s in sheets], sheet, mapping is not None)
    if isinstance(chosen, IssueCode):
        return ctx.blocked(chosen)
    ctx.sheet_index = chosen
    ctx.sheet_name, rows = sheets[chosen].title, sheets[chosen].rows

    if mapping is None:
        ctx.detection = detect_header(rows, MAX_HEADER_SCAN_ROWS)
        if ctx.detection.mapping is None:
            return ctx.blocked(ctx.detection.issue or IssueCode.HEADER_NOT_FOUND)
        mapping = ctx.detection.mapping
    elif mapping.header_row > len(rows):
        return ctx.blocked(IssueCode.INVALID_COLUMN_MAPPING)
    ctx.mapping = mapping
    return _import_rows(ctx, rows, sheets[chosen].unreliable, mapping, date_order)


def _import_rows(
    ctx: _Context,
    rows: list[Row],
    unreliable: dict[CellKey, IssueCode],
    mapping: ColumnMapping,
    date_order: DateOrder | None,
) -> RimsoftImport:
    batch_id, header_row = ctx.batch_id, mapping.header_row
    date_col = mapping.column(ColumnRole.DATE)
    mapped = set(mapping.columns.values())
    issues: list[ValidationIssue] = []
    records: dict[int, RowRecord] = {}

    def issue(n: int, severity: Severity, code: IssueCode, group: str | None = None) -> None:
        issues.append(
            ValidationIssue(
                batch_id,
                Source.RIMSOFT,
                RowRef(row=n),
                severity,
                code,
                row_id=f"R{n}",
                group=group,
            )
        )

    def record(n: int, cls: RowClass, reason: str, codes: Sequence[IssueCode] = ()) -> None:
        tx_id = f"{batch_id}-R{n}" if cls is RowClass.TRANSACTION else None
        records[n] = RowRecord(f"R{n}", n, cls, reason, tuple(codes), tx_id)

    # -- date order, decided once for the whole column --------------------------
    order = date_order
    if order is None:
        order, needed = infer_column_date_order([_cell(r, date_col) for r in rows[header_row:]])
        if needed and order is None:
            issues.append(_sheet_issue(batch_id, IssueCode.AMBIGUOUS_DATE_ORDER))
    ctx.date_order = order

    # -- one operation cut over two rows (a line break in the label) ------------
    # The first row keeps the date and the label but no amount; the rest of its
    # cells (reference, debit, credit...) went to the next row, shifted left. The
    # two are rejoined when, and only when, the result is a valid transaction.
    def classify(cells: Row) -> LedgerRowResult:
        return classify_ledger_row(
            _cell(cells, date_col),
            _cell(cells, mapping.column(ColumnRole.LABEL)),
            _cell(cells, mapping.column(ColumnRole.DEBIT)),
            _cell(cells, mapping.column(ColumnRole.CREDIT)),
            lambda value: parse_date_cell(value, order),
        )

    rows = list(rows)
    continuation: set[int] = set()  # 1-based rows rejoined into the row above
    for i in range(header_row, len(rows) - 1):
        first, second = rows[i], rows[i + 1]
        if (n := i + 1) in continuation or classify(first).issue is not IssueCode.NO_AMOUNT:
            continue
        filled = [k for k, c in enumerate(first) if not is_blank(c)]
        rest = [k for k, c in enumerate(second) if not is_blank(c)]
        if not filled or not rest or not is_blank(_cell(second, 0)):
            continue
        merged = tuple(first[: filled[-1] + 1]) + tuple(second[rest[0] :])
        if (n, filled[-1] + 1) in unreliable or classify(merged).row_class is not (
            RowClass.TRANSACTION
        ):
            continue
        rows[i] = merged
        continuation.add(n + 1)

    # -- classify every row -------------------------------------------------------
    pending: list[tuple[int, Row, date, LedgerSide, int]] = []
    seen_transaction_row = False  # decides opening vs closing for "Solde au <date>"
    for n, row in enumerate(rows, start=1):
        if n in continuation:
            issue(n, Severity.WARNING, IssueCode.SPLIT_ROW_MERGED)
            record(n, RowClass.EMPTY, "continuation_of_previous_row", [IssueCode.SPLIT_ROW_MERGED])
            continue
        if n < header_row:
            if all(is_blank(c) for c in row):
                record(n, RowClass.EMPTY, RowReason.EMPTY.value)
            else:
                issue(n, Severity.INFO, IssueCode.NON_TRANSACTION_ROW)
                record(n, RowClass.PREAMBLE, "above_header", [IssueCode.NON_TRANSACTION_ROW])
            continue
        if n == header_row:
            record(n, RowClass.HEADER, "detected_header")
            continue

        res = classify_ledger_row(
            _cell(row, date_col),
            _cell(row, mapping.column(ColumnRole.LABEL)),
            _cell(row, mapping.column(ColumnRole.DEBIT)),
            _cell(row, mapping.column(ColumnRole.CREDIT)),
            lambda value: parse_date_cell(value, order),
            after_transactions=seen_transaction_row,
        )
        codes: list[IssueCode] = []
        if res.warning is not None:
            issue(n, Severity.WARNING, res.warning)
            codes.append(res.warning)

        # A mapped cell holding an Excel error or a formula without cached value is
        # never read as empty/zero. On a row that is (or could be) a transaction the
        # row becomes INVALID (BLOCKING); on a summary row it is only reported.
        formula_codes = {unreliable[(n, c)] for c in mapped if (n, c) in unreliable}
        if formula_codes:
            code = (
                IssueCode.FORMULA_ERROR_VALUE
                if IssueCode.FORMULA_ERROR_VALUE in formula_codes
                else IssueCode.FORMULA_WITHOUT_CACHED_VALUE
            )
            if res.row_class in (RowClass.TRANSACTION, RowClass.INVALID, RowClass.EMPTY):
                seen_transaction_row = True
                issue(n, Severity.BLOCKING, code)
                record(n, RowClass.INVALID, RowReason.UNRELIABLE_FORMULA.value, [*codes, code])
                continue
            issue(n, Severity.INFO, code)
            codes.append(code)

        if res.row_class in (RowClass.TRANSACTION, RowClass.INVALID):
            seen_transaction_row = True

        if res.row_class is RowClass.EMPTY:
            if any(not is_blank(c) for i, c in enumerate(row) if i not in mapped):
                issue(n, Severity.INFO, IssueCode.CONTENT_OUTSIDE_COLUMNS)
                record(
                    n,
                    RowClass.EMPTY,
                    "content_outside_columns",
                    [IssueCode.CONTENT_OUTSIDE_COLUMNS],
                )
            else:
                record(n, RowClass.EMPTY, res.reason.value)
        elif res.row_class is RowClass.INVALID:
            code = res.issue or IssueCode.BAD_AMOUNT
            issue(n, Severity.BLOCKING, code)
            record(n, RowClass.INVALID, res.reason.value, [*codes, code])
        elif res.row_class is not RowClass.TRANSACTION:
            issue(n, Severity.INFO, IssueCode.NON_TRANSACTION_ROW)
            record(n, res.row_class, res.reason.value, [*codes, IssueCode.NON_TRANSACTION_ROW])
        elif res.tx_date is None or res.side is None or res.amount_minor is None:
            raise RuntimeError("classifier_returned_incomplete_transaction")
        else:
            record(n, RowClass.TRANSACTION, res.reason.value, codes)
            pending.append((n, row, res.tx_date, res.side, res.amount_minor))

    # -- transactions, duplicates, fingerprints -------------------------------
    labels = [cell_text(_cell(r, mapping.column(ColumnRole.LABEL))) for _, r, _, _, _ in pending]
    label_norms = [normalize_text(label) for label in labels]
    pieces = [_optional_text(r, mapping, ColumnRole.PIECE) for _, r, _, _, _ in pending]
    entries = [_optional_text(r, mapping, ColumnRole.ECRITURE) for _, r, _, _, _ in pending]
    keys = [
        (d, side, amount, label_norms[i], pieces[i] or "", entries[i] or "")
        for i, (_, _, d, side, amount) in enumerate(pending)
    ]
    occurrences = occurrence_indexes(keys)
    duplicate_group = {
        i: f"DUP-{g}" for g, group in enumerate(find_duplicate_groups(keys), start=1) for i in group
    }

    transactions: list[Transaction] = []
    for i, (n, row, tx_date, side, amount) in enumerate(pending):
        if i in duplicate_group:
            issue(n, Severity.WARNING, IssueCode.DUPLICATE_ROW, duplicate_group[i])
            rec = records[n]
            records[n] = RowRecord(
                rec.row_id,
                n,
                rec.row_class,
                rec.reason,
                (*rec.issues, IssueCode.DUPLICATE_ROW),
                rec.transaction_id,
            )
        phones = extract_mr_phones(labels[i])
        transactions.append(
            Transaction(
                id=f"{batch_id}-R{n}",
                batch_id=batch_id,
                source=Source.RIMSOFT,
                fingerprint=rimsoft_fingerprint(
                    piece_number=pieces[i],
                    tx_date=tx_date,
                    amount_minor=amount,
                    side=side,
                    label_norm=label_norms[i],
                    occurrence_index=occurrences[i],
                    ecriture_number=entries[i],
                ),
                row_ref=RowRef(row=n),
                tx_date=tx_date,
                amount_minor=amount,
                ledger_side=side,
                reference=_optional_text(row, mapping, ColumnRole.REFERENCE),
                piece_number=pieces[i],
                ecriture_number=entries[i],
                journal=_optional_text(row, mapping, ColumnRole.JOURNAL),
                phone_norm=phones[0] if phones else None,
                text_raw=labels[i],
                text_norm=label_norms[i],
            )
        )

    if not transactions:
        issues.append(
            ValidationIssue(
                batch_id, Source.RIMSOFT, None, Severity.WARNING, IssueCode.NO_TRANSACTIONS
            )
        )
    return ctx.build(issues, [records[n] for n in sorted(records)], transactions)
