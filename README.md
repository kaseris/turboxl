# TurboXL

<p align="center">
  <img src="assets/logo.svg" alt="TurboXL Logo" width="400"/>
</p>
Fast, read-only XLSX to CSV converter with C++20 core and Python bindings.

## Performance

The following end-to-end XLSX-to-CSV results were measured on Azure x86-64
VMs in UK South. The deterministic
[input workbook](benchmarks/turboxl-large-benchmark.xlsx) is 21.7 MB and
contains 150,001 rows, 18 columns, and approximately 2.7 million cells. Each
engine reads the first worksheet and materializes equivalent UTF-8 CSV output
in memory.

### Execution time

Lower is better. The multiplier is the competing engine's elapsed time divided
by TurboXL's elapsed time on the same machine.

| OS | Azure VM | Host CPU | TurboXL | Calamine | OpenPyXL |
| --- | --- | --- | ---: | ---: | ---: |
| Linux | `Standard_D4as_v6` | AMD EPYC 9V74 | **2.752s** | 3.578s (1.30x) | 18.027s (6.55x) |
| Linux | `Standard_D4s_v6` | Intel Xeon Platinum 8573C | **3.126s** | 3.526s (1.13x) | 17.972s (5.75x) |
| Windows | `Standard_D4as_v6` | AMD EPYC 9V74 | **3.097s** | 4.204s (1.36x) | 21.095s (6.81x) |
| Windows | `Standard_D4s_v6` | Intel Xeon Platinum 8573C | **3.544s** | 4.252s (1.20x) | 20.800s (5.87x) |

### Peak process memory

| OS | Azure VM | TurboXL | Calamine | OpenPyXL |
| --- | --- | ---: | ---: | ---: |
| Linux | `Standard_D4as_v6` | **177.5 MiB** | 401.6 MiB | 259.9 MiB |
| Linux | `Standard_D4s_v6` | **177.5 MiB** | 401.6 MiB | 260.0 MiB |
| Windows | `Standard_D4as_v6` | **172.2 MiB** | 373.2 MiB | 235.2 MiB |
| Windows | `Standard_D4s_v6` | **172.4 MiB** | 373.6 MiB | 235.6 MiB |

### Methodology

- Versions: TurboXL 0.3.0, python-calamine 0.8.2, and OpenPyXL 3.1.5.
- Linux: Ubuntu 22.04, Python 3.10.12, one warm-up plus seven measured rounds;
  reported values are medians.
- Windows: Windows Server 2022, Python 3.10.11, one warm-up plus one measured
  validation round. These Windows figures are preliminary until repeated with
  the same seven-round protocol.
- Engine order is shuffled deterministically between rounds. Each measurement
  runs in a fresh child process.
- All engines produced the same row count, byte count, and SHA-256 output hash.
- Linux memory is maximum resident set size; Windows memory is peak working set.
  Treat cross-OS memory comparisons as directional rather than exact.

The raw [Linux](benchmarks/results/linux) and
[Windows](benchmarks/results/windows) JSON and CSV reports are included in the
repository.

Reproduce the benchmark with [`tools/azure_benchmark_once.sh`](tools/azure_benchmark_once.sh)
and [`tools/cloud_benchmark.py`](tools/cloud_benchmark.py):

```bash
# Linux
./tools/azure_benchmark_once.sh --subscription "SUBSCRIPTION"

# Windows
./tools/azure_benchmark_once.sh --subscription "SUBSCRIPTION" --os windows
```

Implementation features include zlib-ng support, release-mode compiler
optimizations, arena-based shared strings, and chunked ZIP reading.

### Typed workbook API

Open an XLSX once when reading typed values from several worksheets:

```python
from pathlib import Path
import turboxl

with turboxl.load_workbook(Path("report.xlsx")) as workbook:
    print(workbook.sheet_names)  # worksheets, including hidden worksheets
    rows = workbook.get_sheet_by_name("Data").to_python(nrows=100)
```

`load_workbook` accepts paths, `os.PathLike`, workbook bytes and seekable
binary streams. Streams are read once and restored to their original cursor;
they are never closed. `sheets_metadata` includes chartsheets and other sheet
kinds for inspection, while name and index lookup select worksheets only.
`Sheet.to_python()` returns rectangular rows containing `None`, `bool`, `int`,
`float`, `str`, `datetime.datetime`, and `datetime.time`. It accepts
`skip_empty_area` and `nrows`; `max_cells` on `load_workbook` limits each read.
Call `close()` or use a context manager when done. Reads and lookups after
close raise `RuntimeError` and retained sheets are also invalidated.

### Typed worksheet benchmark

Issues #100 and #103 provide a private, path-based typed extraction slice to
measure the cost of producing rectangular Python `list[list]` data before
committing to a pandas adapter. The extractor can crop to the used range, stop
at a physical-row limit, and rejects dense results above 10,000,000 cells by
default. `_read_sheet_to_python` remains benchmark scaffolding; use
`load_workbook` for the supported public API.

The private helper accepts keyword-only `skip_empty_area`, `nrows`, and
`max_cells` arguments. With `skip_empty_area=False`, output is anchored at A1;
with it enabled, leading empty rows and columns are removed. Missing rows and
columns inside the selected rectangle remain present, and every returned row
has the same width.

Cells are materialized with pandas-compatible Python scalars: empty and Excel
error cells become `None`, exact integral numbers become `int`, other numbers
remain `float`, and text and booleans become `str` and `bool`. Styled Excel
dates and datetimes become `datetime.datetime`, while styled times become
`datetime.time`; cached formula results are returned without evaluating
formulas. Both Excel date epochs are supported at microsecond precision, with
1900 serial 60 normalized to `1900-02-28`. Styled temporal values outside
Python's representable range remain numeric.

The September 20, 2026 run used macOS 15.7.4 on arm64, CPython 3.14.7,
TurboXL 0.3.0 from this checkout, and python-calamine 0.8.2. Each deterministic
fixture contained 60,000 data rows. Results are medians of nine measurements
after two warm-ups, with each engine run in a fresh process and engine order
alternated between rounds.

| Fixture family | TurboXL total | Native extraction | Python boxing | Calamine total | TurboXL advantage | Peak RSS (TurboXL / Calamine) |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Dense inline strings | 0.0900s | 0.0750s | 0.0080s | 0.1547s | 41.8% | 85.0 / 95.3 MiB |
| Dense shared strings | 0.0868s | 0.0717s | 0.0080s | 0.1396s | 37.8% | 82.3 / 95.0 MiB |
| Sparse mixed primitives | 0.0248s | 0.0140s | 0.0057s | 0.0295s | 16.0% | 55.7 / 55.1 MiB |

All three fixture families produced identical rectangular shapes and compatible
scalar values. The typed-only worksheet scanner retained the existing libxml2
reader as a compatibility fallback and achieved the required 10% median speed
advantage on every family. The performance gate therefore passes and the
conditional pandas-adapter work in #108 may proceed.

Reproduce the run with:

```bash
python tools/benchmark_typed.py \
  --rows 60000 --warmups 2 --rounds 9 \
  --fixtures-dir typed-benchmark-fixtures \
  --json-output typed-benchmark-results.json
```

To reject a median slowdown greater than 5% on any fixture family, capture a
same-machine result before the change and pass it back as the baseline:

```bash
python tools/benchmark_typed.py \
  --rows 60000 --warmups 2 --rounds 9 \
  --baseline-json typed-benchmark-baseline.json \
  --json-output typed-benchmark-results.json
```

The raw report is committed at
[`benchmarks/results/typed/macos-arm64-260920.json`](benchmarks/results/typed/macos-arm64-260920.json).

The bounded-extraction change was rechecked on the same machine against a
fresh pre-change run. Median total times were 0.0902s, 0.0882s, and 0.0237s;
these were respectively 1.8%, 1.3%, and 5.3% faster than the baseline. TurboXL
retained advantages of 41.0%, 36.5%, and 18.2% over calamine, so both the 10%
comparison gate and the maximum 5% regression gate passed. The candidate
report is committed at
[`benchmarks/results/typed/macos-arm64-260921-issue103.json`](benchmarks/results/typed/macos-arm64-260921-issue103.json).

## What It Does

- ✅ Read XLSX files and convert to CSV
- ✅ Handle shared strings, numbers, dates, booleans
- ✅ Process multiple worksheets
- ✅ Memory-efficient XLSX-to-CSV conversion
- ✅ Cross-platform (Linux, macOS, Windows)

## What It Doesn't Do

- ❌ Write or modify XLSX files
- ❌ Formula evaluation (uses cached values)
- ❌ Charts, images, pivot tables
- ❌ Password-protected files

## Quick Start

### Python

```python
import turboxl

# Convert first sheet
csv_data = turboxl.read_sheet_to_csv("data.xlsx")

# Stream directly to disk with atomic replacement (accepts pathlib.Path too)
turboxl.read_sheet_to_file("data.xlsx", "data.csv")

# Convert specific sheet
csv_data = turboxl.read_sheet_to_csv("data.xlsx", sheet="Sheet2")

# Custom options
options = turboxl.CsvOptions()
options.delimiter = ";"
options.date_mode = turboxl.DateMode.ISO

csv_data = turboxl.read_sheet_to_csv(
    "data.xlsx",
    sheet=0,
    options=options,
)

# Inspect every workbook entry. Integer sheet selectors count worksheets only;
# chartsheets and other entry kinds remain visible here as metadata.
for sheet in turboxl.get_sheet_list("data.xlsx"):
    print(sheet.name, sheet.kind, sheet.visibility)

# Save to file
with open("output.csv", "w", encoding="utf-8") as f:
    f.write(csv_data)
```

### C++

```cpp
#include <xlsxcsv.hpp>
#include <iostream>

int main() {
    try {
        std::string csv = xlsxcsv::readSheetToCsv("data.xlsx");
        xlsxcsv::readSheetToFile("data.xlsx", "data.csv");
        std::cout << csv << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
    }
    return 0;
}
```

## Building

### Prerequisites

Install system dependencies (used via pkg-config/CMake):

```bash
# macOS (Recommended for best performance)
brew install libxml2 minizip-ng zlib-ng cmake pkg-config

# Ubuntu/Debian (Recommended for best performance)
sudo apt-get install -y libxml2-dev libminizip-dev cmake build-essential pkg-config
# For zlib-ng on Ubuntu/Debian, build from source:
# git clone https://github.com/zlib-ng/zlib-ng.git
# cd zlib-ng && cmake -B build && cmake --build build -j && sudo cmake --install build

# Windows (vcpkg)
vcpkg install --triplet x64-windows-static-md
```

**Performance Note:** The build system automatically detects and uses zlib-ng
when available, falling back to standard zlib otherwise. Measure the impact on
your own workbook and deployment target.

### Build C++ Core (library only)

Build the C++ core without Python bindings (no Python/nanobind required):

```bash
# From repo root
cmake -S . -B build \
  -DCMAKE_BUILD_TYPE=Release \
  -DBUILD_TESTS=OFF \
  -DBUILD_PYTHON=OFF \
  -DBUILD_CLI=OFF
cmake --build build -j4
```

Artifacts:

- Static library: `build/libturboxl_core.a`

**Build Modes:**

- **Release** (Recommended): Enables `-O3 -march=native -flto` optimizations
- **Debug**: Enables debugging symbols and assertions

### Build Options

- `BUILD_TESTS=ON/OFF` - Build test suite (default: ON)
- `BUILD_PYTHON=ON/OFF` - Build Python bindings (default: ON)
- `BUILD_CLI=ON/OFF` - Build command-line tool (default: OFF)

---

## Python Wheel

TurboXL ships a PEP 517/518 build powered by scikit-build-core. The wheel
installs a regular `turboxl` Python package backed by the native
`turboxl._turboxl` extension, built in Release mode using CMake. Existing
`import turboxl` calls continue to use the same public API.

### Python prerequisites

```bash
python3 -m pip install -U pip build scikit-build-core nanobind
```

System dependencies listed above (libxml2, minizip-ng, zlib-ng, cmake, compiler) must be installed and discoverable by CMake/pkg-config.

### Build the wheel

```bash
# From repo root
python3 -m build -w
```

Outputs go to `dist/`, for example:

- `dist/turboxl-0.3.0-<python>-<abi>-<platform>.whl`

Install the built wheel locally:

```bash
pip install dist/turboxl-*.whl
```

Tips:

- Parallel CMake build: `CMAKE_BUILD_PARALLEL_LEVEL=4 python3 -m build -w`
- macOS arch (defaults to the host architecture): to override, you can pass
  `--config-setting=cmake.define.CMAKE_OSX_ARCHITECTURES="arm64;x86_64"` to `python -m build`.

## Requirements

- **C++**: C++20 compiler (GCC 10+, Clang 12+, MSVC 2019+)
- **Build**: CMake 3.20+
- **Python**: 3.10+ (CPython; free-threaded builds are not supported)

## API Reference

### Python

```python
turboxl.read_sheet_to_csv(
    xlsx_path: str,
    sheet: Union[str, int] = -1,
    options: turboxl.CsvOptions = turboxl.CsvOptions(),
) -> str

turboxl.read_sheet_to_file(
    xlsx_path: str,
    output_path: Union[str, os.PathLike],
    sheet: Union[str, int] = -1,
    options: turboxl.CsvOptions = turboxl.CsvOptions(),
) -> None
```

Configure CSV formatting with a `turboxl.CsvOptions` instance. Enum-valued
fields use the exported enums, such as `turboxl.DateMode.ISO`,
`turboxl.Newline.CRLF`, and `turboxl.MergedHandling.PROPAGATE`; they do not
accept string values.

### C++

```cpp
struct CsvOptions {
    std::string sheetByName;
    int sheetByIndex = -1;
    char delimiter = ',';
    bool includeBom = false;
    // ... more options
};

std::string readSheetToCsv(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector = -1,
    const CsvOptions& options = {}
);

void readSheetToFile(
    const std::string& xlsxPath,
    const std::filesystem::path& outputPath,
    const std::variant<std::string, int>& sheetSelector = -1,
    const CsvOptions& options = {}
);
```

## License

MIT License - see [LICENSE](LICENSE) file for details.

## CI and releases

See [CONTRIBUTING.md](CONTRIBUTING.md) for the complete feature and release
workflow.

CI builds all 12 wheels, runs native Debug tests, and gates Windows releases on
exact CSV parity and median performance against pinned `python-calamine`. These
checks run on every PR and push to `main`. The stable branch-protection check is
**CI passed**. The four wheel targets
are Linux x64 (glibc 2.28+), Windows x64, and macOS 15+ on Intel and Apple Silicon.
Each gets CPython 3.10 and 3.11 wheels plus a CPython 3.12 ABI3 wheel, tested on
3.12, 3.13, and 3.14. Windows 32-bit is no longer supported.

Python package versions come from the CMake `project()` version, including in
source archives without Git metadata. Wheels use portable CPU flags and Release
IPO; native Debug tests do not. Python wheels contain the package, extension, and runtime
libraries; a normal CMake install still supplies native development files.
The platform dependency scripts live in `tools/ci/`. CI pins Python build tools
using `tools/ci/constraints.txt`; Homebrew and distro packages remain rolling
inputs, so builds are not claimed to be bit-for-bit reproducible.

### One-time repository setup

1. Make **CI passed** required on `main`.
2. Keep the existing `PYPI_API_TOKEN` repository secret available to the release
   workflow. The token is used only by the PyPI publication job after all builds
   and release validation succeed.
3. Protect release tags against changes and deletion.

### Rehearse and release

Run **CI → Run workflow** on the intended branch first. This builds and tests the
complete distribution set without publishing. Merge a reviewed CMake version bump
and wait for CI, then tag that exact commit:

```bash
git switch main
git pull --ff-only
# Replace X.Y.Z with the version already recorded in CMakeLists.txt.
git tag -a vX.Y.Z -m "Release vX.Y.Z"
git push origin refs/tags/vX.Y.Z
```

Only the tag push triggers publication. The tag must match CMake and point to a
commit in `main` history. A release publishes the validated source archive and
12 wheels to PyPI, then attaches those same files and SHA256SUMS to a GitHub
Release with generated notes. No workflow automatically creates a tag.

If publication fails, use **Re-run failed jobs**, preserving the successful build
artifacts (retained for 30 days). Existing PyPI files are skipped only after their
SHA256 hashes match the retained artifacts. Differences fail closed; never move a
released tag or overwrite a version. If build artifacts have expired and cannot
be recovered, fix the problem and release a new version. Re-running only the
GitHub Release job is safe after a successful PyPI upload.

### Local checks

```bash
python -m pip install -c tools/ci/constraints.txt build scikit-build-core nanobind packaging
python -m unittest discover -s tools/ci -v
cmake -S . -B out/native -DBUILD_PYTHON=OFF -DBUILD_TESTS=ON \
  -DCMAKE_BUILD_TYPE=Debug -DTURBOXL_ENABLE_IPO=OFF -DTURBOXL_PORTABLE_BUILD=ON
cmake --build out/native --config Debug --parallel 4
ctest --test-dir out/native -C Debug --output-on-failure
python -m build --sdist --no-isolation
```

Native tests use Python's standard-library `zipfile` to create fixtures; no Unix
`zip` executable is required. Fixture creation errors fail the tests.
