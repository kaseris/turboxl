"""Fast XLSX-to-CSV conversion backed by the native TurboXL extension."""

import os

from ._turboxl import (
    CsvOptions,
    DateMode,
    MergedHandling,
    Newline,
    SharedStringsMode,
    SheetKind,
    SheetMetadata,
    SheetVisibility,
    Workbook,
    Sheet,
    _load_workbook_bytes,
    _load_workbook_path,
    _read_sheet_to_python,
    get_sheet_list,
    get_visible_sheets,
    read_multiple_sheets,
    read_sheet_to_csv,
    read_sheet_to_file,
    read_specific_sheet,
)


def load_workbook(source, max_cells=10_000_000):
    """Open an XLSX workbook for repeated typed worksheet reads.

    ``source`` may be a path, ``os.PathLike``, bytes-like XLSX archive, or a
    seekable binary stream. Stream contents are read once and its cursor is
    restored before parsing; the stream is never closed. ``max_cells`` limits
    each ``Sheet.to_python`` result. Use the returned object as a context
    manager or call ``close()``; retained sheets remain usable if only the
    Workbook wrapper is collected, but explicit close invalidates them.
    """
    if isinstance(max_cells, bool) or not isinstance(max_cells, int):
        raise TypeError("max_cells must be an integer")
    if max_cells <= 0:
        raise ValueError("max_cells must be greater than zero")
    if isinstance(source, (bytes, bytearray, memoryview)):
        if isinstance(source, memoryview) and (not source.contiguous or source.itemsize != 1):
            raise TypeError("source buffer must be a contiguous byte-oriented buffer")
        try:
            data = bytes(source)
        except (TypeError, ValueError) as error:
            raise TypeError("source buffer must be contiguous bytes") from error
        return _load_workbook_bytes(data, max_cells)
    if isinstance(source, (str, os.PathLike)):
        return _load_workbook_path(os.fspath(source), max_cells)
    read = getattr(source, "read", None)
    seek = getattr(source, "seek", None)
    tell = getattr(source, "tell", None)
    if not callable(read) or not callable(seek) or not callable(tell):
        raise TypeError("source must be a path, bytes-like object, or seekable binary stream")
    position = tell()
    try:
        seek(0)
        data = read()
    finally:
        seek(position)
    if not isinstance(data, (bytes, bytearray, memoryview)):
        raise TypeError("binary stream read() must return bytes")
    return _load_workbook_bytes(bytes(data), max_cells)

__all__ = [
    "CsvOptions",
    "DateMode",
    "MergedHandling",
    "Newline",
    "SharedStringsMode",
    "SheetKind",
    "SheetMetadata",
    "SheetVisibility",
    "Workbook",
    "Sheet",
    "load_workbook",
    "get_sheet_list",
    "get_visible_sheets",
    "read_multiple_sheets",
    "read_sheet_to_csv",
    "read_sheet_to_file",
    "read_specific_sheet",
]
