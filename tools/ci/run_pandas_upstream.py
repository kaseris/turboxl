#!/usr/bin/env python3
"""Run upstream pandas XLSX reader parametrization with the TurboXL wheel.

Requires git, an internet connection, and a wheel built from the candidate
commit. All upstream sources and environments live under --work-dir.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import subprocess
import sys
import venv
from pathlib import Path

VERSIONS = ("2.2.3", "2.3.3", "3.0.0")


def call(command: list[str], *, cwd: Path | None = None, check: bool = True):
    result = subprocess.run(command, cwd=cwd, capture_output=True, text=True)
    if check and result.returncode:
        raise RuntimeError(f"{' '.join(command)}\n{result.stderr[-3000:]}")
    return result


def patch_reader_tests(path: Path) -> None:
    source = path.read_text()
    if '    pytest.param("turboxl"),\n' in source:
        return
    anchor = '    pytest.param("calamine", marks=td.skip_if_no("python_calamine")),\n'
    if source.count(anchor) != 1:
        raise RuntimeError(f"upstream fixture anchor changed: {path}")
    source = source.replace(anchor, anchor + '    pytest.param("turboxl"),\n')
    anchor = '    if engine == "openpyxl" and read_ext == ".xls":\n'
    if source.count(anchor) != 1:
        raise RuntimeError(f"upstream extension anchor changed: {path}")
    source = source.replace(
        anchor,
        '    if engine == "turboxl" and read_ext != ".xlsx":\n'
        "        return False\n" + anchor,
    )
    path.write_text(source)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--wheel", required=True, type=Path)
    parser.add_argument("--work-dir", required=True, type=Path)
    parser.add_argument("--json-output", required=True, type=Path)
    args = parser.parse_args()
    wheel = args.wheel.resolve()
    if not wheel.is_file():
        parser.error("wheel does not exist")
    work = args.work_dir.resolve()
    work.mkdir(parents=True, exist_ok=True)
    report: dict = {
        "wheel": str(wheel),
        "wheel_sha256": hashlib.sha256(wheel.read_bytes()).hexdigest(),
        "versions": {},
    }
    for version in VERSIONS:
        source = work / f"pandas-{version}"
        if not source.exists():
            call(
                [
                    "git",
                    "clone",
                    "--depth",
                    "1",
                    "--branch",
                    f"v{version}",
                    "https://github.com/pandas-dev/pandas.git",
                    str(source),
                ]
            )
        test_file = source / "pandas/tests/io/excel/test_readers.py"
        patch_reader_tests(test_file)
        environment = work / f"venv-{version}"
        if not environment.exists():
            venv.create(environment, with_pip=True)
        python = environment / (
            "Scripts/python.exe" if sys.platform == "win32" else "bin/python"
        )
        call(
            [
                str(python),
                "-m",
                "pip",
                "install",
                "--quiet",
                f"pandas=={version}",
                "pytest",
                "hypothesis",
                "openpyxl==3.1.5",
                "python-calamine==0.8.2",
                "numpy<2.4",
            ]
        )
        call(
            [
                str(python),
                "-m",
                "pip",
                "install",
                "--quiet",
                "--force-reinstall",
                "--no-deps",
                str(wheel),
            ]
        )
        installed_root = Path(
            call(
                [
                    str(python),
                    "-c",
                    "import pandas, pathlib; print(pathlib.Path(pandas.__file__).parent)",
                ]
            ).stdout.strip()
        )
        installed_test = installed_root / "tests/io/excel/test_readers.py"
        shutil.copy2(test_file, installed_test)
        shutil.copytree(
            source / "pandas/tests/io/data",
            installed_root / "tests/io/data",
            dirs_exist_ok=True,
        )
        plugin = work / f"register_{version.replace('.', '_')}.py"
        plugin.write_text("import turboxl.pandas\nturboxl.pandas.register()\n")
        # The wheel includes upstream data and conftest. Copy only the patched
        # reader test from the pinned tag, avoiding duplicate pandas packages.
        selection = (
            "turboxl and not test_read_from_http_url and not test_read_from_s3_object"
        )
        result = call(
            [
                str(python),
                "-m",
                "pytest",
                "-q",
                "-ra",
                "-k",
                selection,
                "-p",
                plugin.stem,
                str(installed_test),
            ],
            cwd=work,
            check=False,
        )
        report["versions"][version] = {
            "returncode": result.returncode,
            "stdout": result.stdout,
            "stderr": result.stderr,
            "upstream_tag": f"v{version}",
            "upstream_commit": call(
                ["git", "rev-parse", "HEAD"], cwd=source
            ).stdout.strip(),
            "test_file": "pandas/tests/io/excel/test_readers.py",
            "exclusions": [
                "non-XLSX engine/extension pairs",
                "HTTP URL: requires pytest-httpserver fixture",
                "S3: requires external S3 fixture and credentials",
            ],
            "expected_failures": [],
        }
        args.json_output.write_text(json.dumps(report, indent=2) + "\n")
    return 0 if all(v["returncode"] == 0 for v in report["versions"].values()) else 1


if __name__ == "__main__":
    raise SystemExit(main())
