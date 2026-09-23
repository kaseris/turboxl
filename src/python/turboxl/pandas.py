"""Explicit, optional pandas Excel reader registration for TurboXL."""

from __future__ import annotations

import importlib
import re

from . import Workbook, load_workbook

_reader_class = None
_SUPPORTED = {(2, 2), (2, 3), (3, 0)}


def register() -> None:
    """Register ``engine="turboxl"`` with this process's pandas installation.

    Registration changes pandas' private Excel engine map. Call it explicitly
    before using ``pandas.read_excel`` or ``pandas.ExcelFile``.
    """
    try:
        pd = importlib.import_module("pandas")
        base = importlib.import_module("pandas.io.excel._base")
    except ImportError as error:
        raise ImportError(
            "turboxl.pandas.register() requires pandas 2.2, 2.3, or 3.0; "
            "install turboxl[pandas]"
        ) from error

    version = re.match(r"^(\d+)\.(\d+)(?:\.|$)", pd.__version__)
    if version is None or (int(version[1]), int(version[2])) not in _SUPPORTED:
        raise RuntimeError(
            f"pandas {pd.__version__} is unsupported by TurboXL's private Excel "
            "reader adapter; use pandas 2.2, 2.3, or 3.0"
        )

    engines = getattr(pd.ExcelFile, "_engines", None)
    reader_base = getattr(base, "BaseExcelReader", None)
    if (
        not isinstance(engines, dict)
        or not isinstance(reader_base, type)
        or not all(
            callable(getattr(reader_base, name, None))
            for name in (
                "load_workbook",
                "get_sheet_by_name",
                "get_sheet_by_index",
                "get_sheet_data",
                "raise_if_bad_sheet_by_name",
                "raise_if_bad_sheet_by_index",
            )
        )
    ):
        raise RuntimeError(
            "This pandas build has an incompatible private Excel reader API; "
            "turboxl.pandas supports pandas 2.2, 2.3, and 3.0"
        )

    global _reader_class
    current = engines.get("turboxl")
    if current is not None:
        if current is _reader_class:
            return
        raise RuntimeError("pandas Excel engine 'turboxl' is already registered by another provider")

    if _reader_class is None or not issubclass(_reader_class, reader_base):
        class TurboXLReader(reader_base):
            @property
            def _workbook_class(self):
                return Workbook

            def load_workbook(self, filepath_or_buffer, engine_kwargs):
                return load_workbook(filepath_or_buffer, **engine_kwargs)

            @property
            def sheet_names(self):
                return self.book.sheet_names

            def get_sheet_by_name(self, name):
                self.raise_if_bad_sheet_by_name(name)
                return self.book.get_sheet_by_name(name)

            def get_sheet_by_index(self, index):
                self.raise_if_bad_sheet_by_index(index)
                return self.book.get_sheet_by_index(index)

            def get_sheet_data(self, sheet, file_rows_needed=None):
                rows = sheet.to_python(
                    skip_empty_area=False, nrows=file_rows_needed
                )
                # XLSX files often style empty cells far beyond the data.
                # pandas' built-in readers exclude those trailing cells, while
                # retaining empty leading rows and columns for headers/indexes.
                last_row = -1
                last_column = -1
                for row_index, row in enumerate(rows):
                    occupied = [index for index, value in enumerate(row)
                                if value is not None and value != ""]
                    if occupied:
                        last_row = row_index
                        last_column = max(last_column, occupied[-1])
                return [row[:last_column + 1] for row in rows[:last_row + 1]]

            def close(self):
                if getattr(self, "_turboxl_closed", False):
                    return
                self._turboxl_closed = True
                try:
                    book = getattr(self, "book", None)
                    if book is not None:
                        book.close()
                finally:
                    handles = getattr(self, "handles", None)
                    if handles is not None:
                        handles.close()

        _reader_class = TurboXLReader

    engines["turboxl"] = _reader_class
