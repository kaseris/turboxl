#!/usr/bin/env python3
"""Validate and package a frozen pandas XLSX corpus for Azure benchmark VMs."""

import argparse
import importlib.util
import json
from pathlib import Path
import re
import zipfile


def supports_pandas_vm(tag: str) -> bool:
    """Check a wheel tag against Ubuntu 24.04's CPython 3.12 x86-64 runtime."""
    try:
        python, abi, platform = tag.split("-")
    except ValueError:
        return False
    platform_match = re.fullmatch(r"manylinux_(\d+)_(\d+)_x86_64", platform)
    if not platform_match or tuple(map(int, platform_match.groups())) > (2, 39):
        return False
    if python == "py3" and abi == "none":
        return True
    python_match = re.fullmatch(r"cp3(\d+)", python)
    if not python_match:
        return False
    minor = int(python_match.group(1))
    return (abi == "abi3" and 10 <= minor <= 12) or (
        python == "cp312" and abi == "cp312"
    )


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
        if len(metadata) != 1:
            parser.error("wheel must contain one WHEEL metadata file")
        wheel_tags = {
            line.removeprefix("Tag: ")
            for line in wheel_archive.read(metadata[0]).decode().splitlines()
            if line.startswith("Tag: ")
        }
        filename_tag = "-".join(args.wheel.stem.split("-")[-3:])
        if filename_tag not in wheel_tags:
            parser.error("wheel filename tag does not match WHEEL metadata")
        if not supports_pandas_vm(filename_tag):
            parser.error(
                "wheel is incompatible with the pandas VM "
                "(Ubuntu 24.04, CPython 3.12, x86-64)"
            )
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
