#!/usr/bin/env python3
import argparse
import csv
import hashlib
import io
import os
import statistics
import subprocess
import sys
import time
from datetime import date, datetime, time as dt_time

import python_calamine
import turboxl


def format_number_like_turboxl(value):
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if not isinstance(value, float):
        return str(value)
    if value != value:
        return "#NUM!"
    if value == float("inf"):
        return "#DIV/0!"
    if value == float("-inf"):
        return "-#DIV/0!"
    if value == int(value) and abs(value) < 1e15:
        return str(int(value))
    s = f"{value:.6f}".rstrip("0").rstrip(".")
    return s if s else "0"


def normalize_cell(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return value
    if isinstance(value, (int, float, bool)):
        return format_number_like_turboxl(value)
    if isinstance(value, datetime):
        return value.replace(microsecond=0).isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, dt_time):
        return value.replace(microsecond=0).isoformat()
    return str(value)


def calamine_to_csv(path, sheet_index):
    wb = python_calamine.load_workbook(path)
    sh = wb.get_sheet_by_index(sheet_index)
    out = io.StringIO()
    writer = csv.writer(out, lineterminator="\n")
    row_count = 0
    for row in sh.iter_rows():
        writer.writerow([normalize_cell(v) for v in row])
        row_count += 1
    return out.getvalue(), row_count


def turboxl_to_csv(path, sheet_index, profile=False):
    env = os.environ.copy()
    if profile:
        env["TURBOXL_PROFILE_TIMINGS"] = "1"
    code = (
        "import turboxl,sys\n"
        "p=sys.argv[1]\n"
        "i=int(sys.argv[2])\n"
        "print(turboxl.read_sheet_to_csv(p, i), end='')\n"
    )
    proc = subprocess.run(
        [sys.executable, "-c", code, path, str(sheet_index)],
        check=True,
        capture_output=True,
        text=True,
        env=env,
    )
    csv_text = proc.stdout
    timing_line = ""
    for line in proc.stderr.splitlines():
        if line.startswith("turboxl_timing_ms"):
            timing_line = line
            break
    return csv_text, timing_line


def sha256_text(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()


def first_diff_line(a, b):
    al = a.splitlines()
    bl = b.splitlines()
    n = min(len(al), len(bl))
    for i in range(n):
        if al[i] != bl[i]:
            return i + 1, al[i][:160], bl[i][:160]
    if len(al) != len(bl):
        return n + 1, "<EOF>" if n >= len(al) else al[n][:160], "<EOF>" if n >= len(bl) else bl[n][:160]
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("xlsx")
    ap.add_argument("--sheet-index", type=int, default=0)
    ap.add_argument("--rounds", type=int, default=3)
    args = ap.parse_args()

    t_times = []
    c_times = []
    last_t = ""
    last_c = ""
    last_timing = ""

    for i in range(1, args.rounds + 1):
        t0 = time.perf_counter()
        t_csv, timing = turboxl_to_csv(args.xlsx, args.sheet_index, profile=True)
        t_dt = time.perf_counter() - t0

        t0 = time.perf_counter()
        c_csv, c_rows = calamine_to_csv(args.xlsx, args.sheet_index)
        c_dt = time.perf_counter() - t0

        t_times.append(t_dt)
        c_times.append(c_dt)
        last_t = t_csv
        last_c = c_csv
        last_timing = timing

        print(f"round={i} turboxl_sec={t_dt:.3f} calamine_sec={c_dt:.3f} rows={c_rows}")
        if timing:
            print(timing)

    t_hash = sha256_text(last_t)
    c_hash = sha256_text(last_c)
    print("\nPARITY")
    print(f"turboxl_bytes={len(last_t)} sha256={t_hash}")
    print(f"calamine_bytes={len(last_c)} sha256={c_hash}")
    print(f"hash_match={t_hash == c_hash}")
    diff = first_diff_line(last_t, last_c)
    if diff:
        ln, ta, ca = diff
        print(f"first_diff_line={ln}")
        print(f"turboxl: {ta}")
        print(f"calamine: {ca}")

    print("\nSUMMARY")
    print(f"turboxl_median_sec={statistics.median(t_times):.3f}")
    print(f"calamine_median_sec={statistics.median(c_times):.3f}")


if __name__ == "__main__":
    main()
