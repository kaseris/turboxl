#!/usr/bin/env python3
"""End-to-end pandas Excel engine comparison using a frozen XLSX manifest.

Manifest format: {"workbooks": [{"id": str, "path": str, "sha256": str,
"source": str, "publisher": str, "license": str, "sheet_name": str|int,
"read_options": object, "kind": "real"|"synthetic"}]}.
Paths are relative to the manifest. Each worker imports and registers before
timing; only pd.read_excel is timed. Workers serialize frames after timing so
the controller can compare pandas values, indexes, columns and dtypes exactly.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.metadata
import json
import os
import pickle
import platform
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from urllib.parse import urlparse
from urllib.request import url2pathname

ENGINES = ("turboxl", "calamine")


def digest(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def validate_manifest(path: Path) -> list[dict]:
    document = json.loads(path.read_text())
    items = document.get("workbooks")
    if not isinstance(items, list) or not items:
        raise ValueError("manifest must contain nonempty workbooks list")
    ids: set[str] = set()
    for item in items:
        if not isinstance(item, dict):
            raise ValueError("workbook entry must be an object")
        for field in ("id", "path", "sha256", "source", "publisher", "license", "kind"):
            if not isinstance(item.get(field), str) or not item[field]:
                raise ValueError(f"workbook requires nonempty {field}")
        if item["id"] in ids:
            raise ValueError(f"duplicate workbook id: {item['id']}")
        ids.add(item["id"])
        if item["kind"] not in ("real", "synthetic"):
            raise ValueError("kind must be real or synthetic")
        if not isinstance(item.get("sheet_name", 0), (str, int)):
            raise ValueError("sheet_name must be string or integer")
        if not isinstance(item.get("read_options", {}), dict):
            raise ValueError("read_options must be an object")
        if "engine" in item.get("read_options", {}):
            raise ValueError("engine is controlled by the benchmark")
        workbook = (path.parent / item["path"]).resolve()
        if workbook.suffix.lower() != ".xlsx" or not workbook.is_file():
            raise ValueError(f"missing XLSX: {workbook}")
        if digest(workbook) != item["sha256"].lower():
            raise ValueError(f"SHA256 mismatch: {workbook}")
        item["_resolved_path"] = str(workbook)
    return items


def peak_rss_mib() -> float | None:
    if sys.platform == "win32":
        import ctypes
        from ctypes import wintypes

        class ProcessMemoryCounters(ctypes.Structure):
            _fields_ = [
                ("cb", wintypes.DWORD),
                ("PageFaultCount", wintypes.DWORD),
                ("PeakWorkingSetSize", ctypes.c_size_t),
                ("WorkingSetSize", ctypes.c_size_t),
                ("QuotaPeakPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPagedPoolUsage", ctypes.c_size_t),
                ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t),
                ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
                ("PagefileUsage", ctypes.c_size_t),
                ("PeakPagefileUsage", ctypes.c_size_t),
            ]

        counters = ProcessMemoryCounters(cb=ctypes.sizeof(ProcessMemoryCounters))
        if ctypes.windll.psapi.GetProcessMemoryInfo(
            ctypes.windll.kernel32.GetCurrentProcess(),
            ctypes.byref(counters),
            counters.cb,
        ):
            return counters.PeakWorkingSetSize / (1024 * 1024)
        return None
    import resource

    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    return peak / (1024 * 1024 if sys.platform == "darwin" else 1024)


def worker(args: argparse.Namespace) -> None:
    import pandas as pd

    if args.engine == "turboxl":
        import turboxl.pandas

        turboxl.pandas.register()
    item = json.loads(args.item)
    started = time.perf_counter()
    frame = pd.read_excel(
        item["_resolved_path"],
        engine=args.engine,
        sheet_name=item.get("sheet_name", 0),
        **item.get("read_options", {}),
    )
    elapsed = time.perf_counter() - started
    rss = peak_rss_mib()
    with Path(args.frame).open("wb") as stream:
        pickle.dump(frame, stream, protocol=pickle.HIGHEST_PROTOCOL)
    print(
        json.dumps(
            {"seconds": elapsed, "peak_rss_mib": rss, "shape": list(frame.shape)}
        )
    )


def run_worker(args: argparse.Namespace, item: dict, engine: str, frame: Path) -> dict:
    command = [
        args.python,
        str(Path(__file__).resolve()),
        "--worker",
        "--engine",
        engine,
        "--item",
        json.dumps(item),
        "--frame",
        str(frame),
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=False)
    if result.returncode:
        raise RuntimeError(
            f"{item['id']} {engine} worker failed: {result.stderr.strip()}"
        )
    return json.loads(result.stdout)


def compare_frames(left: Path, right: Path) -> str | None:
    import pandas as pd

    with left.open("rb") as stream:
        turbo = pickle.load(stream)
    with right.open("rb") as stream:
        calamine = pickle.load(stream)
    try:
        pd.testing.assert_frame_equal(turbo, calamine, check_exact=True)
    except AssertionError as error:
        return str(error)
    return None


def performance_gate(results: list[dict]) -> tuple[float | None, bool]:
    real = [item for item in results if item["kind"] == "real"]
    if not real:
        return None, False
    complete = [item for item in real if item.get("advantage") is not None]
    if not complete:
        return None, False
    median = statistics.median(item["advantage"] for item in complete)
    publishers = {item.get("publisher") for item in real}
    return median, (
        len(real) >= 12
        and len(publishers) >= 3
        and len(complete) == len(real)
        and all(item["parity"] for item in real)
        and median >= 0.20
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--wheel", type=Path)
    parser.add_argument("--warmups", type=int, default=2)
    parser.add_argument("--rounds", type=int, default=9)
    parser.add_argument("--python", default=sys.executable)
    parser.add_argument("--json-output", type=Path)
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--engine", choices=ENGINES, help=argparse.SUPPRESS)
    parser.add_argument("--item", help=argparse.SUPPRESS)
    parser.add_argument("--frame", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        worker(args)
        return 0
    if not args.manifest or not args.wheel or not args.json_output:
        parser.error("--manifest, --wheel and --json-output are required")
    if args.rounds < 1 or args.warmups < 0:
        parser.error("rounds must be positive; warmups cannot be negative")
    items = validate_manifest(args.manifest.resolve())
    import pandas as pd
    import python_calamine  # noqa: F401 - verifies the comparison engine is installed
    import turboxl  # noqa: F401 - verifies the candidate wheel is installed

    if (
        pd.__version__ != "3.0.0"
        or importlib.metadata.version("python-calamine") != "0.8.2"
    ):
        raise RuntimeError(
            "benchmark requires pandas==3.0.0 and python-calamine==0.8.2"
        )
    origin = importlib.metadata.distribution("turboxl").read_text("direct_url.json")
    if not origin:
        raise RuntimeError("TurboXL must be installed from the candidate wheel")
    installed_url = urlparse(json.loads(origin)["url"])
    if (
        installed_url.scheme != "file"
        or Path(url2pathname(installed_url.path)).resolve() != args.wheel.resolve()
    ):
        raise RuntimeError("installed TurboXL does not match --wheel")
    wheel_hash = digest(args.wheel)
    report = {
        "environment": {
            "platform": platform.platform(),
            "machine": platform.machine(),
            "processor": platform.processor(),
            "cloud_provider": os.environ.get("BENCHMARK_CLOUD_PROVIDER"),
            "cloud_region": os.environ.get("BENCHMARK_CLOUD_REGION"),
            "cloud_vm_size": os.environ.get("BENCHMARK_CLOUD_VM_SIZE"),
            "python": platform.python_version(),
            "pandas": pd.__version__,
            "calamine": importlib.metadata.version("python-calamine"),
            "turboxl": importlib.metadata.version("turboxl"),
            "wheel_sha256": wheel_hash,
            "git_revision": subprocess.run(
                ["git", "rev-parse", "HEAD"],
                capture_output=True,
                text=True,
                check=False,
            ).stdout.strip(),
            "git_dirty": bool(
                subprocess.run(
                    ["git", "status", "--porcelain"],
                    capture_output=True,
                    text=True,
                    check=False,
                ).stdout.strip()
            ),
            "warmups": args.warmups,
            "rounds": args.rounds,
        },
        "workbooks": [],
        "real_median_advantage": None,
        "passes_20_percent_gate": False,
    }
    with tempfile.TemporaryDirectory() as temp:
        for item in items:
            paths = {engine: Path(temp) / f"{engine}.pickle" for engine in ENGINES}
            samples = {engine: [] for engine in ENGINES}
            try:
                for round_index in range(args.warmups + args.rounds):
                    order = (
                        ENGINES if round_index % 2 == 0 else tuple(reversed(ENGINES))
                    )
                    for engine in order:
                        result = run_worker(args, item, engine, paths[engine])
                        if round_index >= args.warmups:
                            samples[engine].append(result)
            except (RuntimeError, OSError, ValueError) as error:
                report["workbooks"].append(
                    {
                        **{k: v for k, v in item.items() if not k.startswith("_")},
                        "samples": samples,
                        "medians": None,
                        "advantage": None,
                        "parity": False,
                        "error": str(error),
                    }
                )
                continue
            mismatch = compare_frames(paths["turboxl"], paths["calamine"])
            medians = {
                engine: statistics.median(s["seconds"] for s in samples[engine])
                for engine in ENGINES
            }
            advantage = (medians["calamine"] - medians["turboxl"]) / medians["calamine"]
            report["workbooks"].append(
                {
                    **{k: v for k, v in item.items() if not k.startswith("_")},
                    "samples": samples,
                    "medians": medians,
                    "advantage": advantage,
                    "parity": mismatch is None,
                    "mismatch": mismatch,
                }
            )
    report["real_median_advantage"], report["passes_20_percent_gate"] = (
        performance_gate(report["workbooks"])
    )
    args.json_output.write_text(json.dumps(report, indent=2) + "\n")
    return 0 if all(x["parity"] for x in report["workbooks"]) else 3


if __name__ == "__main__":
    raise SystemExit(main())
