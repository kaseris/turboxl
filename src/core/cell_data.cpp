#include "xlsxcsv/core.hpp"
#include <regex>
#include <sstream>
#include <cctype>

namespace xlsxcsv::core {

// CellCoordinate implementation
std::optional<CellCoordinate> CellCoordinate::fromReference(const std::string& ref) {
    if (ref.empty()) {
        return std::nullopt;
    }
    
    // Parse Excel reference format: [A-Z]+[1-9][0-9]*
    // Examples: A1, BC42, XFD1048576
    
    size_t i = 0;
    int column = 0;
    
    // Parse column letters (A=1, B=2, ..., Z=26, AA=27, etc.)
    while (i < ref.length() && std::isalpha(ref[i])) {
        char c = std::toupper(ref[i]);
        if (c < 'A' || c > 'Z') {
            return std::nullopt;
        }
        column = column * 26 + (c - 'A' + 1);
        ++i;
    }
    
    if (column == 0 || i == ref.length()) {
        return std::nullopt;  // No column letters or no row number
    }
    
    // Parse row number (1-based)
    int row = 0;
    while (i < ref.length() && std::isdigit(ref[i])) {
        row = row * 10 + (ref[i] - '0');
        ++i;
    }
    
    if (row == 0 || i != ref.length()) {
        return std::nullopt;  // Invalid row number or extra characters
    }
    
    CellCoordinate coord;
    coord.row = row;
    coord.column = column;
    return coord;
}

std::string CellCoordinate::toReference() const {
    if (row <= 0 || column <= 0) {
        return "";
    }
    
    std::string result;
    
    // Convert column number to letters (1=A, 2=B, ..., 26=Z, 27=AA, etc.)
    int col = column;
    while (col > 0) {
        col--; // Make 0-based for modulo operation
        result = char('A' + (col % 26)) + result;
        col /= 26;
    }
    
    // Append row number
    result += std::to_string(row);
    
    return result;
}

// CellData helper method implementations
std::string CellData::getString() const {
    if (std::holds_alternative<std::string>(value)) {
        return std::get<std::string>(value);
    }
    return "";
}

double CellData::getNumber() const {
    if (std::holds_alternative<double>(value)) {
        return std::get<double>(value);
    }
    return 0.0;
}

bool CellData::getBoolean() const {
    if (std::holds_alternative<bool>(value)) {
        return std::get<bool>(value);
    }
    return false;
}

int CellData::getSharedStringIndex() const {
    if (std::holds_alternative<int>(value)) {
        return std::get<int>(value);
    }
    return 0;
}

// RowData implementation
const CellData* RowData::findCell(int column) const {
    for (const auto& cell : cells) {
        if (cell.coordinate.column == column) {
            return &cell;
        }
    }
    return nullptr;
}

CellData* RowData::findCell(int column) {
    for (auto& cell : cells) {
        if (cell.coordinate.column == column) {
            return &cell;
        }
    }
    return nullptr;
}

// MergedCellRange implementation
std::optional<MergedCellRange> MergedCellRange::fromReference(const std::string& ref) {
    // Parse Excel range format: "A1:C3"
    size_t colonPos = ref.find(':');
    if (colonPos == std::string::npos) {
        return std::nullopt;
    }
    
    std::string startRef = ref.substr(0, colonPos);
    std::string endRef = ref.substr(colonPos + 1);
    
    auto startCoord = CellCoordinate::fromReference(startRef);
    auto endCoord = CellCoordinate::fromReference(endRef);
    
    if (!startCoord || !endCoord) {
        return std::nullopt;
    }
    
    MergedCellRange range;
    range.topLeft = startCoord.value();
    range.bottomRight = endCoord.value();
    
    // Ensure topLeft is actually top-left
    if (range.topLeft.row > range.bottomRight.row || 
        range.topLeft.column > range.bottomRight.column) {
        return std::nullopt;
    }
    
    return range;
}

std::string MergedCellRange::toReference() const {
    return topLeft.toReference() + ":" + bottomRight.toReference();
}

bool MergedCellRange::contains(const CellCoordinate& coord) const {
    return coord.row >= topLeft.row && coord.row <= bottomRight.row &&
           coord.column >= topLeft.column && coord.column <= bottomRight.column;
}

std::vector<CellCoordinate> MergedCellRange::getAllCoordinates() const {
    std::vector<CellCoordinate> coords;
    for (int row = topLeft.row; row <= bottomRight.row; ++row) {
        for (int col = topLeft.column; col <= bottomRight.column; ++col) {
            coords.push_back({row, col});
        }
    }
    return coords;
}

// WorksheetMetadata implementation
const MergedCellRange* WorksheetMetadata::findMergedCellRange(const CellCoordinate& coord) const {
    for (const auto& range : mergedCells) {
        if (range.contains(coord)) {
            return &range;
        }
    }
    return nullptr;
}

bool WorksheetMetadata::isColumnHidden(int column) const {
    for (const auto& col : columnInfo) {
        if (col.columnIndex == column) {
            return col.hidden;
        }
    }
    return false; // Not found = not hidden
}

} // namespace xlsxcsv::core