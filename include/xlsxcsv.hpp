#pragma once

#include <string>
#include <variant>
#include <cstdint>
#include <vector>
#include <map>

namespace xlsxcsv {

/**
 * @brief Options for CSV conversion
 */
struct CsvOptions {
    // Sheet selection
    std::string sheetByName;           // Select sheet by name
    int sheetByIndex = -1;             // Select sheet by index (-1 = auto)
    
    // CSV formatting
    char delimiter = ',';               // Field delimiter
    enum class Newline { LF, CRLF };    // Line ending style
    Newline newline = Newline::LF;      // Default to LF
    
    // Output options
    bool includeBom = false;            // Include UTF-8 BOM
    enum class DateMode { ISO, RAW };   // Date formatting mode
    DateMode dateMode = DateMode::ISO;  // Default to ISO format
    bool quoteAll = false;              // Quote all fields
    
    // Shared strings handling
    enum class SharedStringsMode { AUTO, IN_MEMORY, EXTERNAL };
    SharedStringsMode sharedStringsMode = SharedStringsMode::AUTO;
    
    // Merged cells handling
    enum class MergedHandling { NONE, PROPAGATE };
    MergedHandling mergedHandling = MergedHandling::NONE;
    
    // Hidden content handling
    bool includeHiddenRows = true;      // Include hidden rows (default: true)
    bool includeHiddenColumns = true;   // Include hidden columns (default: true)
    
    // Security limits
    uint32_t maxEntries = 10000;                    // Max ZIP entries
    uint64_t maxEntrySize = 256 * 1024 * 1024;     // Max entry size (256 MB)
    uint64_t maxTotalUncompressed = 2ULL * 1024 * 1024 * 1024; // Max total (2 GB)
};

/**
 * @brief Metadata about a worksheet without reading its content
 */
struct SheetMetadata {
    std::string name = "";        // Sheet name
    int sheetId = 0;             // Sheet ID number
    bool visible = false;        // Sheet visibility (false for hidden/veryHidden)
    std::string target = "";     // Internal target path (e.g., "worksheets/sheet1.xml")
};

/**
 * @brief Convert a worksheet from XLSX to CSV
 * 
 * @param xlsxPath Path to the XLSX file
 * @param sheetSelector Sheet name or index (-1 for first sheet)
 * @param options CSV conversion options
 * @return CSV string
 * @throws std::runtime_error on file errors or parsing failures
 */
std::string readSheetToCsv(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector,
    const CsvOptions& options = {}
);

/**
 * @brief Convenience function to read first sheet with default options
 * 
 * @param xlsxPath Path to the XLSX file
 * @return CSV string
 */
std::string readSheetToCsv(const std::string& xlsxPath);

/**
 * @brief Get metadata for all sheets in an XLSX file without reading sheet content
 * 
 * This function only parses the workbook structure, making it very fast for
 * discovering what sheets are available and their visibility status.
 * 
 * @param xlsxPath Path to the XLSX file
 * @return Vector of sheet metadata
 * @throws std::runtime_error on file errors or parsing failures
 */
std::vector<SheetMetadata> getSheetList(const std::string& xlsxPath);

/**
 * @brief Get metadata for only visible sheets in an XLSX file
 * 
 * Convenience function that filters out hidden and very hidden sheets.
 * 
 * @param xlsxPath Path to the XLSX file
 * @return Vector of visible sheet metadata
 * @throws std::runtime_error on file errors or parsing failures
 */
std::vector<SheetMetadata> getVisibleSheets(const std::string& xlsxPath);

/**
 * @brief Convert a specific worksheet to CSV by name
 * 
 * @param xlsxPath Path to the XLSX file
 * @param sheetName Name of the sheet to convert
 * @param options CSV conversion options
 * @return CSV string
 * @throws std::runtime_error on file errors, parsing failures, or if sheet not found
 */
std::string readSpecificSheet(
    const std::string& xlsxPath,
    const std::string& sheetName,
    const CsvOptions& options = {}
);

/**
 * @brief Convert multiple worksheets to CSV by name
 * 
 * More efficient than calling readSpecificSheet multiple times as it reuses
 * the ZIP file parsing and workbook structure.
 * 
 * @param xlsxPath Path to the XLSX file
 * @param sheetNames Vector of sheet names to convert
 * @param options CSV conversion options
 * @return Map of sheet name to CSV string
 * @throws std::runtime_error on file errors or parsing failures
 */
std::map<std::string, std::string> readMultipleSheets(
    const std::string& xlsxPath,
    const std::vector<std::string>& sheetNames,
    const CsvOptions& options = {}
);

} // namespace xlsxcsv
