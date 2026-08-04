#!/usr/bin/env python3
"""Benchmark TurboXL against python-calamine and verify CSV parity."""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import itertools
import json
import os
import statistics
import subprocess
import sys
import tempfile
import time
from datetime import date, datetime, time as dt_time
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from typing import Any


ENGINES = ("turboxl", "calamine")


def format_number_like_turboxl(value: int | float | bool) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if value != value:
        return "#NUM!"
    if value == float("inf"):
        return "#DIV/0!"
    if value == float("-inf"):
        return "-#DIV/0!"
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    rendered = f"{value:.6f}".rstrip("0").rstrip(".")
    return rendered if rendered else "0"


def normalize_cell(value: Any) -> str:
    """Render a calamine cell as closely as possible to TurboXL's CSV output."""
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float, bool)):
        return format_number_like_turboxl(value)
    if isinstance(value, datetime):
        return value.isoformat(timespec="milliseconds")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dt_time):
        return value.isoformat(timespec="milliseconds")
    return str(value)


def package_version(distribution: str) -> str:
    try:
        return version(distribution)
    except PackageNotFoundError:
        return "unknown"


def peak_rss_mib() -> float | None:
    """Return peak RSS for this worker (resource is unavailable on Windows)."""
    try:
        import resource
    except ImportError:
        return None

    rss = float(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    divisor = 1024.0 * 1024.0 if sys.platform == "darwin" else 1024.0
    return rss / divisor


def calamine_to_csv(path: str, sheet_index: int) -> tuple[str, int]:
    import python_calamine

    workbook = python_calamine.load_workbook(path)
    sheet = workbook.get_sheet_by_index(sheet_index)
    output = io.StringIO()
    writer = csv.writer(output, lineterminator="\n")
    rows = 0
    for row in sheet.iter_rows():
        rendered = [normalize_cell(value) for value in row]
        while rendered and rendered[-1] == "":
            rendered.pop()
        writer.writerow(rendered)
        rows += 1
    return output.getvalue(), rows


def turboxl_to_csv(
    path: str, sheet_index: int, max_entry_size_mib: int
) -> tuple[str, int]:
    import turboxl

    options = turboxl.CsvOptions()
    options.max_entry_size = max_entry_size_mib * 1024 * 1024
    output = turboxl.read_sheet_to_csv(path, sheet_index, options)
    return output, output.count("\n")


def worker(args: argparse.Namespace) -> int:
    # Imports happen before the timer so startup/import cost does not distort parsing.
    if args.worker == "turboxl":
        import turboxl  # noqa: F401

        distribution = "turboxl"
        converter = lambda path, sheet: turboxl_to_csv(
            path, sheet, args.max_entry_size_mib
        )
    else:
        import python_calamine  # noqa: F401

        distribution = "python-calamine"
        converter = calamine_to_csv

    started = time.perf_counter()
    csv_text, rows = converter(args.xlsx, args.sheet_index)
    elapsed = time.perf_counter() - started
    encoded = csv_text.encode("utf-8")

    if args.output_csv:
        Path(args.output_csv).write_bytes(encoded)

    print(
        json.dumps(
            {
                "engine": args.worker,
                "version": package_version(distribution),
                "seconds": elapsed,
                "peak_rss_mib": peak_rss_mib(),
                "rows": rows,
                "bytes": len(encoded),
                "sha256": hashlib.sha256(encoded).hexdigest(),
            }
        )
    )
    return 0


def run_worker(
    python: str,
    script: Path,
    xlsx: Path,
    sheet_index: int,
    engine: str,
    max_entry_size_mib: int,
    output_csv: Path | None = None,
    profile_turboxl: bool = False,
) -> tuple[dict[str, Any], str]:
    command = [
        python,
        str(script),
        str(xlsx),
        "--sheet-index",
        str(sheet_index),
        "--worker",
        engine,
        "--max-entry-size-mib",
        str(max_entry_size_mib),
    ]
    if output_csv is not None:
        command.extend(("--output-csv", str(output_csv)))

    environment = os.environ.copy()
    if engine == "turboxl" and profile_turboxl:
        environment["TURBOXL_PROFILE_TIMINGS"] = "1"

    completed = subprocess.run(
        command,
        capture_output=True,
        text=True,
        env=environment,
        check=False,
    )
    if completed.returncode:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise RuntimeError(f"{engine} worker failed:\n{detail}")

    try:
        result = json.loads(completed.stdout)
    except json.JSONDecodeError as error:
        raise RuntimeError(
            f"{engine} worker returned invalid output:\n{completed.stdout}"
        ) from error
    return result, completed.stderr.strip()


def first_differing_line(left: Path, right: Path) -> tuple[int, str, str] | None:
    with left.open(encoding="utf-8", newline="") as left_file, right.open(
        encoding="utf-8", newline=""
    ) as right_file:
        lines = itertools.zip_longest(left_file, right_file, fillvalue=None)
        for number, (left_line, right_line) in enumerate(lines, start=1):
            if left_line != right_line:
                left_value = "<EOF>" if left_line is None else left_line.rstrip("\r\n")
                right_value = (
                    "<EOF>" if right_line is None else right_line.rstrip("\r\n")
                )
                return number, left_value[:200], right_value[:200]
    return None


def describe(values: list[float]) -> str:
    return (
        f"median={statistics.median(values):.3f}s "
        f"min={min(values):.3f}s max={max(values):.3f}s"
    )


def controller(args: argparse.Namespace, parser: argparse.ArgumentParser) -> int:
    xlsx = Path(args.xlsx).expanduser().resolve()
    if not xlsx.is_file():
        parser.error(f"workbook does not exist: {xlsx}")
    if args.rounds < 1:
        parser.error("--rounds must be at least 1")
    if args.warmups < 0:
        parser.error("--warmups cannot be negative")
    if args.max_entry_size_mib < 1:
        parser.error("--max-entry-size-mib must be at least 1")

    python = str(Path(args.python).expanduser())
    script = Path(__file__).resolve()
    try:
        python_version = subprocess.run(
            [python, "--version"],
            capture_output=True,
            text=True,
            check=True,
        )
    except (FileNotFoundError, subprocess.CalledProcessError) as error:
        parser.error(f"cannot execute Python interpreter {python!r}: {error}")

    print(f"workbook: {xlsx} ({xlsx.stat().st_size / (1024 * 1024):.1f} MiB)")
    print(f"python:   {python_version.stdout.strip() or python_version.stderr.strip()}")
    print(f"sheet:    {args.sheet_index}")
    print(f"TurboXL ZIP-entry limit: {args.max_entry_size_mib} MiB")

    try:
        for warmup in range(args.warmups):
            order = ENGINES if warmup % 2 == 0 else tuple(reversed(ENGINES))
            print(f"warm-up {warmup + 1}/{args.warmups}: {', '.join(order)}")
            for engine in order:
                run_worker(
                    python,
                    script,
                    xlsx,
                    args.sheet_index,
                    engine,
                    args.max_entry_size_mib,
                    profile_turboxl=args.profile_turboxl,
                )

        results: dict[str, list[dict[str, Any]]] = {engine: [] for engine in ENGINES}
        with tempfile.TemporaryDirectory(prefix="turboxl-benchmark-") as temp_dir:
            temp = Path(temp_dir)
            parity_files = {
                "turboxl": temp / "turboxl.csv",
                "calamine": temp / "calamine.csv",
            }

            print("\nMEASUREMENTS")
            for round_number in range(1, args.rounds + 1):
                order = (
                    ENGINES
                    if round_number % 2 == 1
                    else tuple(reversed(ENGINES))
                )
                for engine in order:
                    output_csv = (
                        parity_files[engine]
                        if round_number == args.rounds
                        else None
                    )
                    result, diagnostics = run_worker(
                        python,
                        script,
                        xlsx,
                        args.sheet_index,
                        engine,
                        args.max_entry_size_mib,
                        output_csv,
                        args.profile_turboxl,
                    )
                    results[engine].append(result)
                    memory = result["peak_rss_mib"]
                    memory_text = "n/a" if memory is None else f"{memory:.1f} MiB"
                    print(
                        f"round={round_number} engine={engine:<8} "
                        f"time={result['seconds']:.3f}s peak_rss={memory_text} "
                        f"rows={result['rows']:,}"
                    )
                    if diagnostics:
                        for line in diagnostics.splitlines():
                            if line.startswith("turboxl_timing_ms"):
                                print(f"  {line}")

            turbo = results["turboxl"]
            calamine = results["calamine"]
            turbo_last = turbo[-1]
            calamine_last = calamine[-1]
            parity = turbo_last["sha256"] == calamine_last["sha256"]

            print("\nPARITY")
            for engine, result in (("turboxl", turbo_last), ("calamine", calamine_last)):
                print(
                    f"{engine:<8} bytes={result['bytes']:,} "
                    f"sha256={result['sha256']}"
                )
            print(f"exact_csv_match={parity}")
            if not parity:
                difference = first_differing_line(
                    parity_files["turboxl"], parity_files["calamine"]
                )
                if difference:
                    line, turbo_value, calamine_value = difference
                    print(f"first_diff_line={line}")
                    print(f"turboxl:  {turbo_value}")
                    print(f"calamine: {calamine_value}")

        turbo_times = [float(result["seconds"]) for result in results["turboxl"]]
        calamine_times = [float(result["seconds"]) for result in results["calamine"]]
        turbo_median = statistics.median(turbo_times)
        calamine_median = statistics.median(calamine_times)

        print("\nSUMMARY")
        print(f"turboxl  {describe(turbo_times)} version={turbo[0]['version']}")
        print(f"calamine {describe(calamine_times)} version={calamine[0]['version']}")
        if turbo_median:
            print(f"speed_ratio_calamine_over_turboxl={calamine_median / turbo_median:.2f}x")

        for engine in ENGINES:
            memory_values = [
                float(result["peak_rss_mib"])
                for result in results[engine]
                if result["peak_rss_mib"] is not None
            ]
            if memory_values:
                print(
                    f"{engine}_median_peak_rss_mib="
                    f"{statistics.median(memory_values):.1f}"
                )
    except RuntimeError as error:
        print(f"error: {error}", file=sys.stderr)
        if "ModuleNotFoundError" in str(error):
            print(
                f"Install both engines into {python!r}, for example:\n"
                f"  {python} -m pip install turboxl python-calamine",
                file=sys.stderr,
            )
        return 2

    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "Benchmark TurboXL against python-calamine using isolated processes "
            "and verify their normalized CSV output."
        )
    )
    parser.add_argument("xlsx", help="path to an .xlsx workbook")
    parser.add_argument("--sheet-index", type=int, default=0)
    parser.add_argument("--rounds", type=int, default=5)
    parser.add_argument("--warmups", type=int, default=1)
    parser.add_argument(
        "--max-entry-size-mib",
        type=int,
        default=512,
        help=(
            "TurboXL expanded ZIP-entry safety limit; the benchmark workbook's "
            "worksheet XML is about 296 MiB (default: 512)"
        ),
    )
    parser.add_argument(
        "--python",
        default=sys.executable,
        help="Python interpreter containing turboxl and python-calamine",
    )
    parser.add_argument(
        "--profile-turboxl",
        action="store_true",
        help="show TurboXL's internal timing diagnostics when supported",
    )
    parser.add_argument("--worker", choices=ENGINES, help=argparse.SUPPRESS)
    parser.add_argument("--output-csv", help=argparse.SUPPRESS)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.worker:
        return worker(args)
    return controller(args, parser)


if __name__ == "__main__":
    raise SystemExit(main())
