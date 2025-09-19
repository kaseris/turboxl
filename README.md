# TurboXL

<p align="center">
  <img src="assets/logo.svg" alt="TurboXL Logo" width="400"/>
</p>
Fast, read-only XLSX to CSV converter with C++20 core and Python bindings.

## Performance

**Real-world benchmarks** on Chicago Crime dataset (21.9MB, 146,574 rows):

| Metric         | TurboXL         | OpenPyXL       | Improvement      |
| -------------- | --------------- | -------------- | ---------------- |
| **Speed**      | 2.4s            | 63.1s          | **26.7x faster** |
| **Memory**     | 33.5MB          | 66.9MB         | **2.0x less**    |
| **Throughput** | 62,040 rows/sec | 2,321 rows/sec | **26.7x faster** |

_Dataset: [Chicago Crimes 2025](https://data.cityofchicago.org/Public-Safety/Crimes-2025/t7ek-mgzi/about_data)_

🚀 **Recent Optimizations Implemented:**

- **zlib-ng integration** - Up to 2.5x faster ZIP decompression
- **Release build optimizations** - `-O3 -march=native -flto` for GCC/Clang, `/O2 /GL /arch:AVX2` for MSVC
- **Arena-based shared strings** - Memory-efficient string storage
- **Chunked ZIP reading** - 512 KiB buffer optimization

## What It Does

- ✅ Read XLSX files and convert to CSV
- ✅ Handle shared strings, numbers, dates, booleans
- ✅ Process multiple worksheets
- ✅ Memory-efficient streaming (33.5MB for 146k rows)
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

# Convert specific sheet
csv_data = turboxl.read_sheet_to_csv("data.xlsx", sheet="Sheet2")

# Custom options
csv_data = turboxl.read_sheet_to_csv(
    "data.xlsx",
    sheet=0,
    delimiter=";",
    date_mode="iso"
)

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
brew install libxml2 minizip-ng zlib-ng cmake pybind11 pkg-config

# Ubuntu/Debian (Recommended for best performance)
sudo apt-get install -y libxml2-dev libminizip-dev cmake build-essential pkg-config
# For zlib-ng on Ubuntu/Debian, build from source:
# git clone https://github.com/zlib-ng/zlib-ng.git
# cd zlib-ng && cmake -B build && cmake --build build -j && sudo cmake --install build

# Windows (vcpkg)
vcpkg install libxml2 minizip-ng zlib-ng
```

**Performance Note:** Installing `zlib-ng` provides significant performance improvements (up to 2.5x faster decompression). The build system automatically detects and uses zlib-ng if available, falling back to standard zlib otherwise.

### Build C++ Core (library only)

Build the C++ core without Python bindings (no Python/pybind11 required):

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

TurboXL ships a PEP 517/518 build powered by scikit-build-core. The wheel builds the C++ core and Python extension in Release mode using CMake.

### Python prerequisites

```bash
python3 -m pip install -U pip build scikit-build-core pybind11
```

System dependencies listed above (libxml2, minizip-ng, zlib-ng, cmake, compiler) must be installed and discoverable by CMake/pkg-config.

### Build the wheel

```bash
# From repo root
python3 -m build -w
```

Outputs go to `dist/`, for example:

- `dist/turboxl-0.1.0-<python>-<abi>-<platform>.whl`

Install the built wheel locally:

```bash
pip install python/dist/turboxl-*.whl
```

Tips:

- Parallel CMake build: `CMAKE_BUILD_PARALLEL_LEVEL=4 python3 -m build -w`
- macOS arch (defaults to arm64 via `pyproject.toml`): to override, you can pass
  `--config-setting=cmake.define.CMAKE_OSX_ARCHITECTURES="arm64;x86_64"` to `python -m build`.

## Requirements

- **C++**: C++20 compiler (GCC 10+, Clang 12+, MSVC 2019+)
- **Build**: CMake 3.20+
- **Python**: 3.8-3.12 (for Python bindings)

## API Reference

### Python

```python
turboxl.read_sheet_to_csv(
    xlsx_path: str,
    sheet: Union[str, int] = None,  # First sheet if None
    delimiter: str = ",",
    newline: Literal["LF", "CRLF"] = "LF",
    include_bom: bool = False,
    date_mode: Literal["iso", "rawNumber"] = "iso"
) -> str
```

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
    const CsvOptions& opts = {}
);
```

## License

MIT License - see [LICENSE](LICENSE) file for details.
