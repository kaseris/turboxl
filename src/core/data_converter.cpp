#include "xlsxcsv/core.hpp"
#include "xlsxcsv.hpp"  // For CsvOptions
#include <sstream>
#include <iomanip>
#include <cmath>
#include <chrono>
#include <unordered_map>

namespace xlsxcsv::core {

constexpr int64_t SECONDS_PER_DAY = 86400;

// Date conversion class
class DateConverter {
public:
    static std::string convertExcelSerial(double serialDate, 
                                         DateSystem dateSystem,
                                         NumberFormatType formatType) {
        // Round once at second precision. Repeated floating-point truncation of the
        // hour, minute, and second components turns values such as 13:45:30 into
        // 13:45:29 and fails to carry values that round across midnight.
        int64_t serialSeconds = std::llround(serialDate * SECONDS_PER_DAY);
        int64_t serialDay = serialSeconds / SECONDS_PER_DAY;
        int64_t secondsOfDay = serialSeconds % SECONDS_PER_DAY;
        if (secondsOfDay < 0) {
            secondsOfDay += SECONDS_PER_DAY;
            --serialDay;
        }

        std::string date;
        if (dateSystem == DateSystem::Date1900 && serialDay == 60) {
            // Excel intentionally preserves Lotus 1-2-3's fictional leap day.
            date = "1900-02-29";
        } else {
            const auto epoch = dateSystem == DateSystem::Date1904
                ? std::chrono::sys_days{std::chrono::year{1904}/1/1}
                : std::chrono::sys_days{std::chrono::year{1899}/12/31};
            const int64_t leapDayAdjustment =
                dateSystem == DateSystem::Date1900 && serialDay > 60 ? 1 : 0;
            const std::chrono::year_month_day civilDate{
                epoch + std::chrono::days{serialDay - leapDayAdjustment}};

            std::ostringstream dateStream;
            dateStream << std::setfill('0')
                       << std::setw(4) << static_cast<int>(civilDate.year()) << '-'
                       << std::setw(2) << static_cast<unsigned>(civilDate.month()) << '-'
                       << std::setw(2) << static_cast<unsigned>(civilDate.day());
            date = dateStream.str();
        }

        const int hours = static_cast<int>(secondsOfDay / 3600);
        const int minutes = static_cast<int>((secondsOfDay % 3600) / 60);
        const int seconds = static_cast<int>(secondsOfDay % 60);

        std::ostringstream timeStream;
        timeStream << std::setfill('0')
                   << std::setw(2) << hours << ':'
                   << std::setw(2) << minutes << ':'
                   << std::setw(2) << seconds;

        std::ostringstream oss;
        if (formatType == NumberFormatType::Time) {
            oss << timeStream.str();
        } else if (formatType == NumberFormatType::DateTime) {
            oss << date << 'T' << timeStream.str();
        } else {
            oss << date;
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
        if (!std::isfinite(value)) {
            return formatNumericValue(value);
        }

        if (styles && styleIndex > 0 && styles->isDateTimeStyle(styleIndex)) {
            const auto style = styles->getCellStyle(styleIndex);
            if (style && (style->numberFormat.type == NumberFormatType::Date ||
                          style->numberFormat.type == NumberFormatType::Time ||
                          style->numberFormat.type == NumberFormatType::DateTime)) {
                return DateConverter::convertExcelSerial(
                    value, dateSystem, style->numberFormat.type);
            }
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
                // RAW mode preserves the numeric serial even when a date style exists.
                const auto* styles = (m_options && m_options->dateMode == ::xlsxcsv::CsvOptions::DateMode::RAW)
                    ? nullptr : m_styles;
                cellValue = DataConverter::convertCellValue(*cell, m_sharedStrings, styles, m_dateSystem);

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
