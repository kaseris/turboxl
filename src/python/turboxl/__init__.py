"""Fast XLSX-to-CSV conversion backed by the native TurboXL extension."""

from ._turboxl import (
    CsvOptions,
    DateMode,
    MergedHandling,
    Newline,
    SharedStringsMode,
    SheetKind,
    SheetMetadata,
    SheetVisibility,
    _read_sheet_to_python,
    get_sheet_list,
    get_visible_sheets,
    read_multiple_sheets,
    read_sheet_to_csv,
    read_sheet_to_file,
    read_specific_sheet,
)

__all__ = [
    "CsvOptions",
    "DateMode",
    "MergedHandling",
    "Newline",
    "SharedStringsMode",
    "SheetKind",
    "SheetMetadata",
    "SheetVisibility",
    "get_sheet_list",
    "get_visible_sheets",
    "read_multiple_sheets",
    "read_sheet_to_csv",
    "read_sheet_to_file",
    "read_specific_sheet",
]
