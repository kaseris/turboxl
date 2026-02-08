#include "xlsxcsv/core.hpp"
#include "xlsxcsv.hpp"  // For CsvOptions
#include <sstream>
#include <iomanip>
#include <cmath>
#include <chrono>
#include <unordered_map>

namespace xlsxcsv::core {

// Excel date constants
[[maybe_unused]] constexpr double EXCEL_EPOCH_1900 = 1.0;    // January 1, 1900 (Excel day 1)
constexpr double EXCEL_EPOCH_1904 = 1462.0; // January 1, 1904 (Mac Excel)
constexpr double SECONDS_PER_DAY = 86400.0;

// Date conversion class
class DateConverter {
public:
    static std::string convertExcelSerial(double serialDate, 
                                         DateSystem dateSystem,
                                         [[maybe_unused]] const std::string& formatCode = "") {
        
        // Handle zero and negative values
        if (serialDate <= 0.0) {
            return "1900-01-01";
        }
        
        // Adjust for date system
        double adjustedSerial = serialDate;
        if (dateSystem == DateSystem::Date1904) {
            adjustedSerial += EXCEL_EPOCH_1904;
        }
        
        // Excel 1900 date system has a bug - it considers 1900 a leap year
        // We need to account for this when converting
        if (dateSystem == DateSystem::Date1900 && serialDate >= 60.0) {
            adjustedSerial -= 1.0; // Account for the phantom Feb 29, 1900
        }
        
        // Convert to days since Unix epoch (January 1, 1970)
        // Excel epoch 1900 = December 30, 1899 (not January 1, 1900 due to the bug)
        constexpr double DAYS_BETWEEN_1899_AND_1970 = 25567.0;
        double daysSinceUnixEpoch = adjustedSerial - DAYS_BETWEEN_1899_AND_1970;
        
        // Convert to seconds and create time_point
        int64_t secondsSinceEpoch = static_cast<int64_t>(daysSinceUnixEpoch * SECONDS_PER_DAY);
        auto timePoint = std::chrono::system_clock::from_time_t(secondsSinceEpoch);
        
        // Get fractional part for time
        double fractionalPart = adjustedSerial - std::floor(adjustedSerial);
        int hours = static_cast<int>(fractionalPart * 24.0);
        int minutes = static_cast<int>((fractionalPart * 24.0 - hours) * 60.0);
        int seconds = static_cast<int>(((fractionalPart * 24.0 - hours) * 60.0 - minutes) * 60.0);
        
        // Convert to tm structure for formatting
        time_t timeT = std::chrono::system_clock::to_time_t(timePoint);
        std::tm* tm = std::gmtime(&timeT);
        
        if (!tm) {
            return "1900-01-01";
        }
        
        // Override time components with calculated values
        tm->tm_hour = hours;
        tm->tm_min = minutes;
        tm->tm_sec = seconds;
        
        // Determine output format based on format code analysis
        bool hasDatePart = fractionalPart < 0.999;  // Has date if not purely time
        bool hasTimePart = fractionalPart > 0.001;  // Has time if significant fractional part
        
        std::ostringstream oss;
        
        if (hasDatePart && hasTimePart) {
            // Date and time
            oss << std::put_time(tm, "%Y-%m-%dT%H:%M:%S");
        } else if (hasTimePart) {
            // Time only
            oss << std::put_time(tm, "%H:%M:%S");
        } else {
            // Date only
            oss << std::put_time(tm, "%Y-%m-%d");
        }
        
        return oss.str();
    }
    
};

// Main data conversion class
class DataConverter {
public:
    static std::string convertCellValue(const CellData& cell,
                                       const SharedStringsProvider* sharedStrings,
                                       const StylesRegistry* styles,
                                       DateSystem dateSystem = DateSystem::Date1900) {
        
        // Handle empty cells
        if (cell.isEmpty()) {
            return "";
        }
        
        // Handle different cell types
        switch (cell.type) {
            case CellType::Boolean:
                return cell.getBoolean() ? "TRUE" : "FALSE";
                
            case CellType::Error:
                return formatErrorValue(cell.getString());
                
            case CellType::InlineString:
            case CellType::String:
                return cell.getString();
                
            case CellType::SharedString:
                if (sharedStrings && cell.isSharedStringIndex()) {
                    auto str = sharedStrings->tryGetString(static_cast<size_t>(cell.getSharedStringIndex()));
                    return str.value_or("");
                }
                return cell.getString(); // Fallback to resolved string
                
            case CellType::Number:
                return convertNumericValue(cell.getNumber(), cell.styleIndex, styles, dateSystem);
                
            case CellType::Unknown:
            default:
                return cell.getString(); // Best effort conversion
        }
    }

private:
    static std::string formatErrorValue(const std::string& errorCode) {
        // Return Excel error codes as-is
        if (errorCode.empty()) return "#N/A";
        return errorCode;
    }
    
    static std::string convertNumericValue(double value, 
                                         int styleIndex,
                                         const StylesRegistry* styles,
                                         DateSystem dateSystem) {
        
        // Check if this should be formatted as a date/time
        if (styles && styleIndex > 0 && styles->isDateTimeStyle(styleIndex)) {
            return DateConverter::convertExcelSerial(value, dateSystem);
        }
        
        // Format as regular number
        return formatNumericValue(value);
    }
    
    static std::string formatNumericValue(double value) {
        // Handle special cases
        if (std::isnan(value)) return "#NUM!";
        if (std::isinf(value)) return value > 0 ? "#DIV/0!" : "-#DIV/0!";
        
        // Check if it's effectively an integer
        if (value == std::floor(value) && std::abs(value) < 1e15) {
            // Format as integer
            return std::to_string(static_cast<long long>(value));
        }
        
        // Use stream formatting for maximum standard-library portability.
        std::ostringstream oss;
        oss << std::fixed << std::setprecision(6) << value;
        std::string result = oss.str();
        
        // Remove trailing zeros and decimal point if not needed
        if (result.find('.') != std::string::npos) {
            result.erase(result.find_last_not_of('0') + 1);
            if (result.back() == '.') {
                result.pop_back();
            }
        }
        
        return result;
    }
};

// CSV Row Handler implementation
class CsvRowCollectorImpl {
public:
    explicit CsvRowCollectorImpl(const SharedStringsProvider* sharedStrings = nullptr,
                               const StylesRegistry* styles = nullptr,
                               DateSystem dateSystem = DateSystem::Date1900,
                               const void* options = nullptr)
        : m_sharedStrings(sharedStrings)
        , m_styles(styles) 
        , m_dateSystem(dateSystem)
        , m_options(static_cast<const ::xlsxcsv::CsvOptions*>(options)) {
        
        // Set delimiter from options or default
        m_delimiter = (m_options && m_options->delimiter != '\0') ? m_options->delimiter : ',';
    }
    
    void handleRow(const RowData& row) {
        // Check if row should be skipped due to hidden row filtering
        if (row.hidden && m_options && !m_options->includeHiddenRows) {
            return; // Skip hidden row
        }
        
        if (row.cells.empty()) {
            // Empty row
            m_csvOutput.push_back('\n');
            ++m_rowCount;
            return;
        }

        // Find max column to handle sparse data
        int maxColumn = 0;
        for (const auto& cell : row.cells) {
            maxColumn = std::max(maxColumn, cell.coordinate.column);
        }

        bool firstField = true;
        std::size_t cellIndex = 0;

        // Generate CSV row with proper spacing for sparse data.
        // Row cells are parsed in document order, so advance a single cursor.
        for (int col = 1; col <= maxColumn; ++col) {
            // Check if column should be skipped due to hidden column filtering
            if (m_worksheetMetadata.isColumnHidden(col) && m_options && !m_options->includeHiddenColumns) {
                continue; // Skip hidden column
            }

            const CellData* cell = nullptr;
            while (cellIndex < row.cells.size() && row.cells[cellIndex].coordinate.column < col) {
                ++cellIndex;
            }
            if (cellIndex < row.cells.size() && row.cells[cellIndex].coordinate.column == col) {
                cell = &row.cells[cellIndex];
                ++cellIndex;
            }

            std::string cellValue;

            if (cell) {
                cellValue = DataConverter::convertCellValue(*cell, m_sharedStrings, m_styles, m_dateSystem);

                // If this cell is the top-left of a merged range, cache its value
                if (m_options && m_options->mergedHandling == ::xlsxcsv::CsvOptions::MergedHandling::PROPAGATE) {
                    const MergedCellRange* mergedRange = m_worksheetMetadata.findMergedCellRange(cell->coordinate);
                    if (mergedRange && mergedRange->topLeft.row == cell->coordinate.row && 
                        mergedRange->topLeft.column == cell->coordinate.column) {
                        // This is the top-left cell of a merged range - cache the value
                        m_mergedCellValues[mergedRange->toReference()] = cellValue;
                    }
                }
            } else {
                // Check for merged cell propagation
                cellValue = handleMergedCell(CellCoordinate{row.rowNumber, col});
            }

            if (!firstField) {
                m_csvOutput.push_back(m_delimiter);
            }
            firstField = false;
            appendEscapedCsvField(cellValue);
        }

        m_csvOutput.push_back('\n');
        ++m_rowCount;
    }
    
    void handleError(const std::string& message) {
        m_errorMessages.push_back(message);
    }
    
    void handleWorksheetMetadata(const WorksheetMetadata& metadata) {
        m_worksheetMetadata = metadata;
    }
    
    std::string getCsvString() const {
        return m_csvOutput;
    }
    
    const std::vector<std::string>& getErrors() const {
        return m_errorMessages;
    }
    
    size_t getRowCount() const {
        return m_rowCount;
    }

private:
    void appendEscapedCsvField(const std::string& field) {
        // Check if field needs quoting
        bool needsQuoting = field.find(m_delimiter) != std::string::npos ||
                           field.find('"') != std::string::npos ||
                           field.find('\n') != std::string::npos ||
                           field.find('\r') != std::string::npos;

        if (!needsQuoting) {
            m_csvOutput.append(field);
            return;
        }

        m_csvOutput.push_back('"');
        for (char ch : field) {
            if (ch == '"') {
                m_csvOutput.push_back('"');
            }
            m_csvOutput.push_back(ch);
        }
        m_csvOutput.push_back('"');
    }
    
    std::string handleMergedCell(const CellCoordinate& coord) {
        // Check if merged cell propagation is enabled
        if (!m_options || m_options->mergedHandling != ::xlsxcsv::CsvOptions::MergedHandling::PROPAGATE) {
            return ""; // No propagation, return empty
        }
        
        // Find merged cell range that contains this coordinate
        const MergedCellRange* mergedRange = m_worksheetMetadata.findMergedCellRange(coord);
        if (!mergedRange) {
            return ""; // Not in a merged range
        }
        
        // Look for cached value for this merged range
        auto it = m_mergedCellValues.find(mergedRange->toReference());
        if (it != m_mergedCellValues.end()) {
            return it->second; // Return cached value
        }
        
        // Find the value from the top-left cell of the merged range
        // Note: We would need to store cell data to look this up
        // For now, return empty string as we can't retroactively look up cell values
        // This limitation could be addressed by caching all cell data during processing
        return "";
    }
    
    const SharedStringsProvider* m_sharedStrings;
    const StylesRegistry* m_styles;
    DateSystem m_dateSystem;
    const ::xlsxcsv::CsvOptions* m_options;
    char m_delimiter;
    
    WorksheetMetadata m_worksheetMetadata;
    std::unordered_map<std::string, std::string> m_mergedCellValues; // Cache for merged cell values
    std::string m_csvOutput;
    size_t m_rowCount = 0;
    std::vector<std::string> m_errorMessages;
};

// CsvRowCollector PIMPL wrapper
class CsvRowCollector::Impl : public CsvRowCollectorImpl {
public:
    using CsvRowCollectorImpl::CsvRowCollectorImpl;
};

CsvRowCollector::CsvRowCollector(const SharedStringsProvider* sharedStrings,
                               const StylesRegistry* styles,
                               DateSystem dateSystem,
                               const void* csvOptions)
    : m_impl(std::make_unique<Impl>(sharedStrings, styles, dateSystem, csvOptions)) {
}

CsvRowCollector::~CsvRowCollector() = default;

void CsvRowCollector::handleRow(const RowData& row) {
    m_impl->handleRow(row);
}

void CsvRowCollector::handleError(const std::string& message) {
    m_impl->handleError(message);
}

void CsvRowCollector::handleWorksheetMetadata(const WorksheetMetadata& metadata) {
    m_impl->handleWorksheetMetadata(metadata);
}

std::string CsvRowCollector::getCsvString() const {
    return m_impl->getCsvString();
}

const std::vector<std::string>& CsvRowCollector::getErrors() const {
    return m_impl->getErrors();
}

size_t CsvRowCollector::getRowCount() const {
    return m_impl->getRowCount();
}

} // namespace xlsxcsv::core
