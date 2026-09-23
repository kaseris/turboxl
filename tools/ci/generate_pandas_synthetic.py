#!/usr/bin/env python3
"""Generate the three typed fixture families for pandas diagnostic timings."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import sys

FAMILIES = ("dense-inline", "dense-shared", "sparse-mixed")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--rows", type=int, default=60_000)
    parser.add_argument("--output-dir", type=Path, required=True)
    args = parser.parse_args()
    if args.rows < 1:
        parser.error("--rows must be positive")
    root = args.output_dir.resolve()
    root.mkdir(parents=True, exist_ok=True)
    generator = Path(__file__).resolve().parents[1] / "generate_benchmark_fixture.py"
    workbooks = []
    for family in FAMILIES:
        path = root / f"{family}.xlsx"
        subprocess.run(
            [
                sys.executable,
                str(generator),
                str(path),
                "--rows",
                str(args.rows),
                "--family",
                family,
            ],
            check=True,
        )
        workbooks.append(
            {
                "id": family,
                "path": path.name,
                "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
                "source": "tools/generate_benchmark_fixture.py",
                "publisher": "TurboXL",
                "license": "MIT",
                "kind": "synthetic",
                "sheet_name": 0,
                "read_options": {"header": None},
            }
        )
    (root / "manifest.json").write_text(
        json.dumps({"workbooks": workbooks}, indent=2) + "\n"
    )


if __name__ == "__main__":
    main()
