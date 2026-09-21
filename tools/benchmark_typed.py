#!/usr/bin/env python3
"""Benchmark TurboXL's typed vertical slice against python-calamine."""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import math
import os
import platform
import statistics
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


FAMILIES = ("dense-inline", "dense-shared", "sparse-mixed")
ENGINES = ("turboxl", "calamine")
TIMING_PREFIX = "turboxl_typed_timing_ms "


def peak_rss_mib() -> float | None:
    try:
        import resource
    except ImportError:
        return None
    value = float(resource.getrusage(resource.RUSAGE_SELF).ru_maxrss)
    return value / (1024.0 * 1024.0 if sys.platform == "darwin" else 1024.0)


def package_version(distribution: str) -> str:
    try:
        return importlib.metadata.version(distribution)
    except importlib.metadata.PackageNotFoundError:
        return "unknown"


def canonical_cell(value: object) -> tuple[str, object]:
    if value is None:
        return ("empty", None)
    if isinstance(value, bool):
        return ("bool", value)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        number = float(value)
        if math.isnan(number):
            return ("number", "nan")
        if math.isinf(number):
            return ("number", "inf" if number > 0 else "-inf")
        return ("number", number.hex())
    # Calamine 0.8 represents sparse holes as empty strings while the typed
    # vertical slice uses None. Both mean an empty Excel cell.
    if value == "":
        return ("empty", None)
    if isinstance(value, str):
        return ("string", value)
    raise TypeError(f"unsupported benchmark value: {type(value).__name__}")


def describe_matrix(rows: list[list[object]]) -> dict[str, object]:
    width = len(rows[0]) if rows else 0
    if any(len(row) != width for row in rows):
        raise ValueError("engine returned non-rectangular rows")
    digest = hashlib.sha256()
    for row in rows:
        encoded = json.dumps(
            [canonical_cell(value) for value in row],
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")
        digest.update(len(encoded).to_bytes(8, "little"))
        digest.update(encoded)
    return {
        "rows": len(rows),
        "columns": width,
        "cells": len(rows) * width,
        "sha256": digest.hexdigest(),
    }


def worker(args: argparse.Namespace) -> int:
    phases: dict[str, float] = {}
    started = time.perf_counter()
    if args.engine == "turboxl":
        import turboxl

        rows = turboxl._read_sheet_to_python(
            args.xlsx,
            0,
            skip_empty_area=args.skip_empty_area,
            nrows=args.nrows,
            max_cells=args.max_cells,
        )
        distribution = "turboxl"
    else:
        import python_calamine

        load_started = time.perf_counter()
        workbook = python_calamine.load_workbook(args.xlsx)
        sheet = workbook.get_sheet_by_index(0)
        phases["load_seconds"] = time.perf_counter() - load_started
        materialize_started = time.perf_counter()
        rows = sheet.to_python(
            skip_empty_area=args.skip_empty_area,
            nrows=args.nrows,
        )
        phases["materialize_seconds"] = time.perf_counter() - materialize_started
        distribution = "python-calamine"
    elapsed = time.perf_counter() - started
    result = {
        "engine": args.engine,
        "version": package_version(distribution),
        "seconds": elapsed,
        "peak_rss_mib": peak_rss_mib(),
        "phases": phases,
        **describe_matrix(rows),
    }
    print(json.dumps(result, separators=(",", ":")))
    return 0


def parse_turboxl_timings(stderr: str) -> dict[str, float]:
    for line in stderr.splitlines():
        if not line.startswith(TIMING_PREFIX):
            continue
        values: dict[str, float] = {}
        for item in line[len(TIMING_PREFIX) :].split():
            key, value = item.split("=", 1)
            if key in {"native", "boxing", "total"}:
                values[f"{key}_seconds"] = float(value) / 1000.0
        return values
    return {}


def run_worker(
    python: str,
    script: Path,
    fixture: Path,
    engine: str,
    args: argparse.Namespace,
) -> dict[str, Any]:
    environment = os.environ.copy()
    if engine == "turboxl":
        environment["TURBOXL_PROFILE_TYPED_TIMINGS"] = "1"
    command = [
        python,
        str(script),
        "--worker",
        "--engine",
        engine,
        "--xlsx",
        str(fixture),
        "--max-cells",
        str(args.max_cells),
    ]
    if args.skip_empty_area:
        command.append("--skip-empty-area")
    if args.nrows is not None:
        command.extend(["--nrows", str(args.nrows)])
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
    result = json.loads(completed.stdout)
    if engine == "turboxl":
        result["phases"].update(parse_turboxl_timings(completed.stderr))
    return result


def generate_fixtures(args: argparse.Namespace, script: Path) -> dict[str, Path]:
    directory = Path(args.fixtures_dir).expanduser().resolve()
    directory.mkdir(parents=True, exist_ok=True)
    generator = script.with_name("generate_benchmark_fixture.py")
    fixtures: dict[str, Path] = {}
    for family in FAMILIES:
        output = directory / f"typed-{family}.xlsx"
        subprocess.run(
            [
                args.python,
                str(generator),
                str(output),
                "--rows",
                str(args.rows),
                "--family",
                family,
            ],
            check=True,
        )
        fixtures[family] = output
    return fixtures


def median(values: list[dict[str, Any]]) -> float:
    return statistics.median(float(value["seconds"]) for value in values)


def git_revision() -> str:
    completed = subprocess.run(
        ["git", "rev-parse", "HEAD"],
        capture_output=True,
        text=True,
        check=False,
    )
    return completed.stdout.strip() if completed.returncode == 0 else "unknown"


def load_baseline(args: argparse.Namespace) -> dict[str, Any] | None:
    if not args.baseline_json:
        return None
    baseline = json.loads(Path(args.baseline_json).read_text())
    expected = {
        "platform": platform.platform(),
        "machine": platform.machine(),
        "python": platform.python_version(),
        "rows": args.rows,
        "warmups": args.warmups,
        "rounds": args.rounds,
    }
    actual = baseline.get("environment", {})
    mismatches = [
        key for key, value in expected.items() if actual.get(key) != value
    ]
    if actual.get("skip_empty_area", False) != args.skip_empty_area:
        mismatches.append("skip_empty_area")
    if actual.get("nrows") != args.nrows:
        mismatches.append("nrows")
    if mismatches:
        raise ValueError(
            "baseline environment differs for: " + ", ".join(mismatches)
        )
    return baseline


def controller(args: argparse.Namespace) -> int:
    if args.rows < 1 or args.rounds < 1 or args.warmups < 0:
        raise ValueError("rows and rounds must be positive; warmups cannot be negative")
    if args.nrows is not None and args.nrows < 0:
        raise ValueError("nrows cannot be negative")
    if args.max_cells < 1:
        raise ValueError("max-cells must be positive")
    if args.max_regression < 0:
        raise ValueError("max-regression cannot be negative")
    script = Path(__file__).resolve()
    fixtures = generate_fixtures(args, script)
    baseline = load_baseline(args)
    report: dict[str, Any] = {
        "environment": {
            "platform": platform.platform(),
            "machine": platform.machine(),
            "python": platform.python_version(),
            "rows": args.rows,
            "warmups": args.warmups,
            "rounds": args.rounds,
            "skip_empty_area": args.skip_empty_area,
            "nrows": args.nrows,
            "max_cells": args.max_cells,
            "git_revision": git_revision(),
        },
        "families": {},
    }
    all_parity = True
    all_gate = True
    all_no_regression = True

    for family, fixture in fixtures.items():
        print(f"\n{family}: {fixture.name} ({fixture.stat().st_size / 1024 / 1024:.1f} MiB)")
        for warmup in range(args.warmups):
            order = ENGINES if warmup % 2 == 0 else tuple(reversed(ENGINES))
            for engine in order:
                run_worker(args.python, script, fixture, engine, args)

        results: dict[str, list[dict[str, Any]]] = {engine: [] for engine in ENGINES}
        for round_number in range(args.rounds):
            order = ENGINES if round_number % 2 == 0 else tuple(reversed(ENGINES))
            for engine in order:
                result = run_worker(args.python, script, fixture, engine, args)
                results[engine].append(result)
                print(
                    f"  round={round_number + 1} engine={engine:<8} "
                    f"time={result['seconds']:.4f}s rss={result['peak_rss_mib']}"
                )

        turbo_median = median(results["turboxl"])
        calamine_median = median(results["calamine"])
        representative = {engine: values[-1] for engine, values in results.items()}
        parity = (
            representative["turboxl"]["rows"] == representative["calamine"]["rows"]
            and representative["turboxl"]["columns"]
            == representative["calamine"]["columns"]
            and representative["turboxl"]["sha256"]
            == representative["calamine"]["sha256"]
        )
        advantage = (calamine_median - turbo_median) / calamine_median
        passes_gate = parity and advantage >= 0.10
        print(
            f"  parity={parity} turbo_median={turbo_median:.4f}s "
            f"calamine_median={calamine_median:.4f}s advantage={advantage:.1%} "
            f"gate={passes_gate}"
        )
        family_report = {
            "fixture": fixture.name,
            "parity": parity,
            "turboxl_median_seconds": turbo_median,
            "calamine_median_seconds": calamine_median,
            "turboxl_advantage": advantage,
            "passes_10_percent_gate": passes_gate,
            "results": results,
        }
        if baseline:
            baseline_median = float(
                baseline["families"][family]["turboxl_median_seconds"]
            )
            regression = (turbo_median - baseline_median) / baseline_median
            passes_regression = regression <= args.max_regression
            family_report["baseline_turboxl_median_seconds"] = baseline_median
            family_report["turboxl_regression"] = regression
            family_report["passes_regression_gate"] = passes_regression
            all_no_regression = all_no_regression and passes_regression
            print(
                f"  baseline={baseline_median:.4f}s regression={regression:.1%} "
                f"regression_gate={passes_regression}"
            )
        report["families"][family] = family_report
        all_parity = all_parity and parity
        all_gate = all_gate and passes_gate

    report["exact_compatible_value_parity"] = all_parity
    report["typed_adapter_may_proceed"] = all_gate
    report["max_regression"] = args.max_regression
    report["no_performance_regression"] = all_no_regression
    if args.json_output:
        Path(args.json_output).write_text(json.dumps(report, indent=2) + "\n")
    print(
        f"\noverall parity={all_parity} typed_adapter_may_proceed={all_gate} "
        f"no_performance_regression={all_no_regression}"
    )
    if not all_parity:
        return 3
    if not all_gate:
        return 4
    return 0 if all_no_regression else 5


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--python", default=sys.executable)
    parser.add_argument("--fixtures-dir", default="typed-benchmark-fixtures")
    parser.add_argument("--rows", type=int, default=60_000)
    parser.add_argument("--warmups", type=int, default=2)
    parser.add_argument("--rounds", type=int, default=9)
    parser.add_argument("--json-output")
    parser.add_argument("--skip-empty-area", action="store_true")
    parser.add_argument("--nrows", type=int)
    parser.add_argument("--max-cells", type=int, default=10_000_000)
    parser.add_argument("--baseline-json")
    parser.add_argument("--max-regression", type=float, default=0.05)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--engine", choices=ENGINES, help=argparse.SUPPRESS)
    parser.add_argument("--xlsx", help=argparse.SUPPRESS)
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    if args.worker:
        if not args.engine or not args.xlsx:
            parser.error("--worker requires --engine and --xlsx")
        return worker(args)
    try:
        return controller(args)
    except (OSError, RuntimeError, subprocess.CalledProcessError, ValueError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
