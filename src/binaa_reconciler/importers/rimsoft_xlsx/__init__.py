"""RIM-Soft ledger importer for XLSX files (local, read-only)."""

from binaa_reconciler.importers.rimsoft_xlsx.header import ColumnMapping, HeaderDetection
from binaa_reconciler.importers.rimsoft_xlsx.importer import (
    MAX_HEADER_SCAN_ROWS,
    RimsoftImport,
    import_rimsoft_xlsx,
)
from binaa_reconciler.importers.rimsoft_xlsx.report import ImportReport, RowRecord

__all__ = [
    "MAX_HEADER_SCAN_ROWS",
    "ColumnMapping",
    "HeaderDetection",
    "ImportReport",
    "RimsoftImport",
    "RowRecord",
    "import_rimsoft_xlsx",
]
