# TurboXL

Fast, read-only XLSX to CSV converter with C++20 core and Python bindings.

## Performance

**Real-world benchmarks** on Chicago Crime dataset (21.9MB, 146,574 rows):

| Metric | TurboXL | OpenPyXL | Improvement |
|--------|---------|----------|-------------|
| **Speed** | 8.4s | 64.7s | **7.7x faster** |
| **Memory** | 33.5MB | 66.9MB | **2x less** |
| **Throughput** | 17,457 rows/sec | 2,266 rows/sec | **7.7x faster** |

*Dataset: [Chicago Crimes 2025](https://data.cityofchicago.org/Public-Safety/Crimes-2025/t7ek-mgzi/about_data)*

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

Install dependencies:

```bash
# macOS
brew install libxml2 minizip-ng cmake

# Ubuntu/Debian
sudo apt-get install libxml2-dev libminizip-dev cmake build-essential

# Windows (vcpkg)
vcpkg install libxml2 minizip-ng
```

### Build Steps

```bash
mkdir build && cd build
cmake ..
make -j4
```

### Build Options

- `BUILD_TESTS=ON/OFF` - Build test suite (default: ON)
- `BUILD_PYTHON=ON/OFF` - Build Python bindings (default: ON)
- `BUILD_CLI=ON/OFF` - Build command-line tool (default: OFF)

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