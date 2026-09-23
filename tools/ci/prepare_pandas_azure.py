#!/usr/bin/env python3
"""Validate and package a frozen pandas XLSX corpus for Azure benchmark VMs."""

import argparse
import importlib.util
import json
from pathlib import Path
import zipfile


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--bundle", required=True, type=Path)
    parser.add_argument("--wheel", required=True, type=Path)
    args = parser.parse_args()
    if "manylinux" not in args.wheel.name or "x86_64.whl" not in args.wheel.name:
        parser.error("wheel must be a Linux x86-64 wheel")
    with zipfile.ZipFile(args.wheel) as wheel_archive:
        metadata = [
            name
            for name in wheel_archive.namelist()
            if name.endswith(".dist-info/WHEEL")
        ]
        if (
            len(metadata) != 1
            or "manylinux" not in wheel_archive.read(metadata[0]).decode()
        ):
            parser.error("wheel metadata must contain a manylinux platform tag")
    source = Path(__file__).resolve().parents[1] / "benchmark_pandas.py"
    spec = importlib.util.spec_from_file_location("benchmark_pandas", source)
    benchmark = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(benchmark)
    items = benchmark.validate_manifest(args.manifest.resolve())
    packaged = []
    with zipfile.ZipFile(args.bundle, "w", compression=zipfile.ZIP_STORED) as archive:
        for index, item in enumerate(items):
            name = f"data/{index:03d}.xlsx"
            archive.write(item["_resolved_path"], name)
            packaged.append(
                {
                    **{k: v for k, v in item.items() if not k.startswith("_")},
                    "path": name,
                }
            )
        archive.writestr("manifest.json", json.dumps({"workbooks": packaged}, indent=2))
    print(f"Packaged {len(items)} validated workbooks into {args.bundle}")


if __name__ == "__main__":
    main()
