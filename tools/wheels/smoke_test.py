"""Minimal installed-wheel and architecture smoke test for cibuildwheel."""

from __future__ import annotations

import struct
import sys

import turboxl


def main() -> None:
    expected_bits = int(sys.argv[1]) if len(sys.argv) > 1 else None
    actual_bits = struct.calcsize("P") * 8

    if expected_bits is not None and actual_bits != expected_bits:
        raise RuntimeError(
            f"wheel architecture mismatch: expected {expected_bits}-bit Python, "
            f"got {actual_bits}-bit"
        )
    if not callable(turboxl.read_sheet_to_csv):
        raise RuntimeError("turboxl.read_sheet_to_csv is missing or not callable")

    print(f"turboxl import OK on {actual_bits}-bit {sys.platform}")


if __name__ == "__main__":
    main()
