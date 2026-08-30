"""Preserve modern Excel 365 features while editing with openpyxl."""

from openpyxl365.workbook import (
    Excel365Workbook,
    UnsupportedWorkbookMetadataError,
    inspect_excel365_dynamic_arrays,
    normalize_excel365_formula,
)

__all__ = [
    "Excel365Workbook",
    "UnsupportedWorkbookMetadataError",
    "inspect_excel365_dynamic_arrays",
    "normalize_excel365_formula",
]

__version__ = "0.1.1"
