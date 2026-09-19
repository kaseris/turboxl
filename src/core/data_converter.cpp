#include "xlsxcsv/core.hpp"
#include "xlsxcsv.hpp"  // For CsvOptions
#include <sstream>
#include <iomanip>
#include <cmath>
#include <chrono>
#include <unordered_map>
#include <charconv>
#include <array>
#include <cstdio>
#include <system_error>
#include <ostream>

namespace xlsxcsv::core {

constexpr int64_t SECONDS_PER_DAY = 86400;

// Date conversion class
class DateConverter {
public:
    static void convertExcelSerial(double serialDate,
                                   DateSystem dateSystem,
                                   NumberFormatType formatType,
                                   std::string& output) {
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

        char dateBuffer[16] = {};
        if (dateSystem == DateSystem::Date1900 && serialDay == 60) {
            // Excel intentionally preserves Lotus 1-2-3's fictional leap day.
            output = "1900-02-29";
        } else {
            const auto epoch = dateSystem == DateSystem::Date1904
                ? std::chrono::sys_days{std::chrono::year{1904}/1/1}
                : std::chrono::sys_days{std::chrono::year{1899}/12/31};
            const int64_t leapDayAdjustment =
                dateSystem == DateSystem::Date1900 && serialDay > 60 ? 1 : 0;
            const std::chrono::year_month_day civilDate{
                epoch + std::chrono::days{serialDay - leapDayAdjustment}};

            const int written = std::snprintf(
                dateBuffer, sizeof(dateBuffer), "%04d-%02u-%02u",
                static_cast<int>(civilDate.year()),
                static_cast<unsigned>(civilDate.month()),
                static_cast<unsigned>(civilDate.day()));
            output.assign(dateBuffer, static_cast<size_t>(written));
        }

        const int hours = static_cast<int>(secondsOfDay / 3600);
        const int minutes = static_cast<int>((secondsOfDay % 3600) / 60);
        const int seconds = static_cast<int>(secondsOfDay % 60);

        char timeBuffer[16] = {};
        const int timeLength = std::snprintf(
            timeBuffer, sizeof(timeBuffer), "%02d:%02d:%02d", hours, minutes, seconds);
        if (formatType == NumberFormatType::Time) {
            output.assign(timeBuffer, static_cast<size_t>(timeLength));
            return;
        } else if (formatType == NumberFormatType::DateTime) {
            output.push_back('T');
            output.append(timeBuffer, static_cast<size_t>(timeLength));
        }
    }
    
};

// Main data conversion class
class DataConverter {
public:
    static std::string convertCellValue(const CellData& cell,
                                       const SharedStringsProvider* sharedStrings,
                                       const StylesRegistry* styles,
                                       DateSystem dateSystem = DateSystem::Date1900) {
        std::string output;
        convertCellValueTo(cell, sharedStrings, styles, dateSystem, output);
        return output;
    }

    static void convertCellValueTo(const CellData& cell,
                                   const SharedStringsProvider* sharedStrings,
                                   const StylesRegistry* styles,
                                   DateSystem dateSystem,
                                   std::string& output) {
        output.clear();
        // Handle empty cells
        if (cell.isEmpty()) {
            return;
        }
        
        // Handle different cell types
        switch (cell.type) {
            case CellType::Boolean:
                output = cell.getBoolean() ? "TRUE" : "FALSE";
                return;
                
            case CellType::Error:
                output = formatErrorValue(cell.getString());
                return;
                
            case CellType::InlineString:
            case CellType::String:
                output = cell.getString();
                return;
                
            case CellType::SharedString:
                if (sharedStrings && cell.isSharedStringIndex()) {
                    auto str = sharedStrings->tryGetString(static_cast<size_t>(cell.getSharedStringIndex()));
                    output = str.value_or("");
                    return;
                }
                output = cell.getString();
                return;
                
            case CellType::Number:
                convertNumericValue(cell.getNumber(), cell.styleIndex, styles, dateSystem, output);
                return;
                
            case CellType::Unknown:
            default:
                output = cell.getString();
                return;
        }
    }

private:
    static std::string formatErrorValue(const std::string& errorCode) {
        // Return Excel error codes as-is
        if (errorCode.empty()) return "#N/A";
        return errorCode;
    }
    
    static void convertNumericValue(double value,
                                    int styleIndex,
                                    const StylesRegistry* styles,
                                    DateSystem dateSystem,
                                    std::string& output) {
        
        // Check if this should be formatted as a date/time
        if (!std::isfinite(value)) {
            formatNumericValue(value, output);
            return;
        }

        if (styles && styleIndex > 0 && styles->isDateTimeStyle(styleIndex)) {
            const auto type = styles->getNumberFormatTypeForStyle(styleIndex);
            if (type == NumberFormatType::Date || type == NumberFormatType::Time ||
                type == NumberFormatType::DateTime) {
                DateConverter::convertExcelSerial(value, dateSystem, type, output);
                return;
            }
        }
        
        // Format as regular number
        formatNumericValue(value, output);
    }
    
    static void formatNumericValue(double value, std::string& output) {
        // Handle special cases
        if (std::isnan(value)) { output = "#NUM!"; return; }
        if (std::isinf(value)) { output = value > 0 ? "#DIV/0!" : "-#DIV/0!"; return; }
        
        // Check if it's effectively an integer
        if (value == std::floor(value) && std::abs(value) < 1e15) {
            // Format as integer
            std::array<char, 32> buffer{};
            const auto converted = std::to_chars(
                buffer.data(), buffer.data() + buffer.size(), static_cast<long long>(value));
            output.assign(buffer.data(), converted.ptr);
            return;
        }

        std::array<char, 384> buffer{};
        const auto converted = std::to_chars(
            buffer.data(), buffer.data() + buffer.size(), value,
            std::chars_format::fixed, 6);
        if (converted.ec != std::errc{}) {
            throw XlsxError("Failed to format numeric cell value");
        }
        output.assign(buffer.data(), converted.ptr);
        
        // Remove trailing zeros and decimal point if not needed
        if (output.find('.') != std::string::npos) {
            output.erase(output.find_last_not_of('0') + 1);
            if (output.back() == '.') {
                output.pop_back();
            }
        }
    }
};

// CSV Row Handler implementation
class CsvRowCollectorImpl {
public:
    explicit CsvRowCollectorImpl(const SharedStringsProvider* sharedStrings = nullptr,
                               const StylesRegistry* styles = nullptr,
                               DateSystem dateSystem = DateSystem::Date1900,
                               const void* options = nullptr,
                               std::ostream* outputStream = nullptr)
        : m_sharedStrings(sharedStrings)
        , m_styles(styles) 
        , m_dateSystem(dateSystem)
        , m_options(static_cast<const ::xlsxcsv::CsvOptions*>(options))
        , m_outputStream(outputStream) {
        
        // Set delimiter from options or default
        m_delimiter = (m_options && m_options->delimiter != '\0') ? m_options->delimiter : ',';
        m_newline = (m_options && m_options->newline == ::xlsxcsv::CsvOptions::Newline::CRLF)
            ? "\r\n" : "\n";
        if (m_options && m_options->includeBom) {
            m_csvOutput.append("\xEF\xBB\xBF", 3);
        }
    }
    
    void handleRow(const RowData& row) {
        // Worksheet XML is sparse: rows with no cells are commonly omitted
        // entirely. Preserve their physical positions as blank CSV records.
        if (row.rowNumber > m_lastPhysicalRow + 1) {
            const size_t missingRows = static_cast<size_t>(row.rowNumber - m_lastPhysicalRow - 1);
            for (size_t i = 0; i < missingRows; ++i) m_csvOutput.append(m_newline);
            m_rowCount += missingRows;
        }
        m_lastPhysicalRow = std::max(m_lastPhysicalRow, row.rowNumber);

        // Check if row should be skipped due to hidden row filtering
        if (row.hidden && m_options && !m_options->includeHiddenRows) {
            return; // Skip hidden row
        }
        
        if (row.cells.empty()) {
            // Empty row
            m_csvOutput.append(m_newline);
            ++m_rowCount;
            flushIfNeeded();
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
            if (m_options && !m_options->includeHiddenColumns &&
                isColumnHidden(col)) {
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
            std::string_view cellView;
            bool hasCellView = false;

            if (cell) {
                // RAW mode preserves the numeric serial even when a date style exists.
                const auto* styles = (m_options && m_options->dateMode == ::xlsxcsv::CsvOptions::DateMode::RAW)
                    ? nullptr : m_styles;
                const bool needsOwnedValue = m_options &&
                    m_options->mergedHandling == ::xlsxcsv::CsvOptions::MergedHandling::PROPAGATE;
                if (!needsOwnedValue &&
                    (cell->type == CellType::InlineString || cell->type == CellType::String) &&
                    std::holds_alternative<std::string>(cell->value)) {
                    cellView = std::get<std::string>(cell->value);
                    hasCellView = true;
                } else if (!needsOwnedValue && cell->type == CellType::SharedString &&
                           m_sharedStrings && cell->isSharedStringIndex()) {
                    auto view = m_sharedStrings->tryGetStringView(
                        static_cast<size_t>(cell->getSharedStringIndex()));
                    if (view) {
                        cellView = *view;
                        hasCellView = true;
                    } else {
                        DataConverter::convertCellValueTo(
                            *cell, m_sharedStrings, styles, m_dateSystem, cellValue);
                    }
                } else if (!needsOwnedValue && cell->type == CellType::Boolean) {
                    cellView = cell->getBoolean() ? std::string_view("TRUE") : std::string_view("FALSE");
                    hasCellView = true;
                } else {
                    DataConverter::convertCellValueTo(
                        *cell, m_sharedStrings, styles, m_dateSystem, cellValue);
                }

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
            appendEscapedCsvField(hasCellView ? cellView : std::string_view(cellValue));
        }

        m_csvOutput.append(m_newline);
        ++m_rowCount;
        flushIfNeeded();
    }

    bool acceptsStreamingCells() const {
        return !m_options ||
            m_options->mergedHandling == ::xlsxcsv::CsvOptions::MergedHandling::NONE;
    }

    void beginRow(int rowNumber, bool hidden) {
        if (rowNumber > m_lastPhysicalRow + 1) {
            const size_t missingRows = static_cast<size_t>(rowNumber - m_lastPhysicalRow - 1);
            for (size_t i = 0; i < missingRows; ++i) m_csvOutput.append(m_newline);
            m_rowCount += missingRows;
        }
        m_lastPhysicalRow = std::max(m_lastPhysicalRow, rowNumber);
        m_streamSkipRow = hidden && m_options && !m_options->includeHiddenRows;
        m_streamFirstField = true;
        m_streamNextColumn = 1;
    }

    void handleCell(CellData&& cell) {
        if (m_streamSkipRow) return;
        const int column = cell.coordinate.column;
        for (int gap = m_streamNextColumn; gap < column; ++gap) {
            if (m_options && !m_options->includeHiddenColumns &&
                isColumnHidden(gap)) continue;
            emitStreamingField({});
        }
        m_streamNextColumn = column + 1;
        if (m_options && !m_options->includeHiddenColumns &&
            isColumnHidden(column)) return;

        const auto* styles = (m_options &&
            m_options->dateMode == ::xlsxcsv::CsvOptions::DateMode::RAW) ? nullptr : m_styles;
        if ((cell.type == CellType::InlineString || cell.type == CellType::String) &&
            std::holds_alternative<std::string>(cell.value)) {
            emitStreamingField(std::get<std::string>(cell.value));
        } else if (cell.type == CellType::SharedString && m_sharedStrings &&
                   cell.isSharedStringIndex()) {
            auto view = m_sharedStrings->tryGetStringView(
                static_cast<size_t>(cell.getSharedStringIndex()));
            if (view) emitStreamingField(*view);
            else {
                DataConverter::convertCellValueTo(
                    cell, m_sharedStrings, styles, m_dateSystem, m_cellScratch);
                emitStreamingField(m_cellScratch);
            }
        } else if (cell.type == CellType::Boolean) {
            emitStreamingField(cell.getBoolean() ? std::string_view("TRUE")
                                                  : std::string_view("FALSE"));
        } else {
            DataConverter::convertCellValueTo(
                cell, m_sharedStrings, styles, m_dateSystem, m_cellScratch);
            emitStreamingField(m_cellScratch);
        }
    }

    void endRow() {
        if (m_streamSkipRow) return;
        m_csvOutput.append(m_newline);
        ++m_rowCount;
        flushIfNeeded();
    }
    
    void handleError(const std::string& message) {
        m_errorMessages.push_back(message);
    }
    
    void handleWorksheetMetadata(const WorksheetMetadata& metadata) {
        m_worksheetMetadata = metadata;
        m_hiddenColumns.clear();
        if (m_options && !m_options->includeHiddenColumns) {
            int maxColumn = 0;
            for (const auto& column : metadata.columnInfo) {
                maxColumn = std::max(maxColumn, column.columnIndex);
            }
            m_hiddenColumns.assign(static_cast<size_t>(maxColumn + 1), 0);
            for (const auto& column : metadata.columnInfo) {
                if (column.hidden && column.columnIndex >= 0) {
                    m_hiddenColumns[static_cast<size_t>(column.columnIndex)] = 1;
                }
            }
        }
    }
    
    std::string getCsvString() const {
        return m_csvOutput;
    }

    std::string takeCsvString() {
        return std::move(m_csvOutput);
    }

    void finalize() {
        if (!m_outputStream) return;
        flushOutput();
        m_outputStream->flush();
        if (!*m_outputStream) throw XlsxError("Failed to write CSV output");
    }
    
    const std::vector<std::string>& getErrors() const {
        return m_errorMessages;
    }
    
    size_t getRowCount() const {
        return m_rowCount;
    }

private:
    bool isColumnHidden(int column) const {
        return column >= 0 && static_cast<size_t>(column) < m_hiddenColumns.size() &&
            m_hiddenColumns[static_cast<size_t>(column)] != 0;
    }

    void emitStreamingField(std::string_view value) {
        if (!m_streamFirstField) m_csvOutput.push_back(m_delimiter);
        m_streamFirstField = false;
        appendEscapedCsvField(value);
    }

    void flushIfNeeded() {
        if (m_outputStream && m_csvOutput.size() >= 256 * 1024) flushOutput();
    }

    void flushOutput() {
        if (m_csvOutput.empty()) return;
        m_outputStream->write(m_csvOutput.data(), static_cast<std::streamsize>(m_csvOutput.size()));
        if (!*m_outputStream) throw XlsxError("Failed to write CSV output");
        m_csvOutput.clear();
    }

    void appendEscapedCsvField(std::string_view field) {
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
    std::string_view m_newline;
    std::ostream* m_outputStream;
    
    WorksheetMetadata m_worksheetMetadata;
    std::vector<uint8_t> m_hiddenColumns;
    std::string m_cellScratch;
    std::unordered_map<std::string, std::string> m_mergedCellValues; // Cache for merged cell values
    std::string m_csvOutput;
    size_t m_rowCount = 0;
    int m_lastPhysicalRow = 0;
    bool m_streamSkipRow = false;
    bool m_streamFirstField = true;
    int m_streamNextColumn = 1;
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
                               const void* csvOptions,
                               std::ostream* outputStream)
    : m_impl(std::make_unique<Impl>(sharedStrings, styles, dateSystem, csvOptions, outputStream)) {
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

bool CsvRowCollector::acceptsStreamingCells() const {
    return m_impl->acceptsStreamingCells();
}

void CsvRowCollector::beginRow(int rowNumber, bool hidden) {
    m_impl->beginRow(rowNumber, hidden);
}

void CsvRowCollector::handleCell(CellData&& cell) {
    m_impl->handleCell(std::move(cell));
}

void CsvRowCollector::endRow() {
    m_impl->endRow();
}

std::string CsvRowCollector::getCsvString() const {
    return m_impl->getCsvString();
}

std::string CsvRowCollector::takeCsvString() {
    return m_impl->takeCsvString();
}

void CsvRowCollector::finalize() {
    m_impl->finalize();
}

const std::vector<std::string>& CsvRowCollector::getErrors() const {
    return m_impl->getErrors();
}

size_t CsvRowCollector::getRowCount() const {
    return m_impl->getRowCount();
}

} // namespace xlsxcsv::core
