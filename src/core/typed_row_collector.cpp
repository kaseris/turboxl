#include "typed_reader.hpp"

#include <algorithm>
#include <limits>
#include <sstream>
#include <utility>

namespace xlsxcsv::internal {

namespace {

TypedCellValue convertCell(const core::CellData& cell,
                           const core::SharedStringsProvider* sharedStrings,
                           std::vector<std::string>& errors) {
    if (cell.isEmpty()) {
        return std::monostate{};
    }
    if (cell.type == core::CellType::SharedString && cell.isSharedStringIndex()) {
        if (sharedStrings) {
            auto value = sharedStrings->tryGetString(
                static_cast<std::size_t>(cell.getSharedStringIndex()));
            if (value) {
                return std::move(*value);
            }
        }
        errors.push_back(
            "Shared string index out of range: " +
            std::to_string(cell.getSharedStringIndex()));
        return std::monostate{};
    }
    if (std::holds_alternative<bool>(cell.value)) {
        return std::get<bool>(cell.value);
    }
    if (std::holds_alternative<double>(cell.value)) {
        return std::get<double>(cell.value);
    }
    if (std::holds_alternative<std::string>(cell.value)) {
        return std::get<std::string>(cell.value);
    }
    errors.push_back("Unsupported typed cell value at " + cell.coordinate.toReference());
    return std::monostate{};
}

} // namespace

class TypedRowCollector::Impl {
public:
    struct DenseStoredRow {
        int number;
        TypedRow values;
    };

    struct SparseStoredRow {
        int number;
        std::vector<std::pair<std::size_t, TypedCellValue>> values;
    };

    explicit Impl(
        const core::SharedStringsProvider* sharedStrings,
        TypedReadOptions options)
        : sharedStrings(sharedStrings), options(std::move(options)) {
        if (this->options.maxCells == 0) {
            throw std::invalid_argument("max_cells must be greater than zero");
        }
        completed = this->options.nrows && *this->options.nrows == 0;
    }

    static std::size_t inclusiveSpan(std::size_t first, std::size_t last) {
        if (last < first || last - first == std::numeric_limits<std::size_t>::max()) {
            throw std::runtime_error("Worksheet dense range dimensions overflow");
        }
        return last - first + 1;
    }

    void checkCellLimit(std::size_t height, std::size_t width) const {
        if (height == 0 || width == 0) return;
        const bool overflow = width > std::numeric_limits<std::size_t>::max() / height;
        const std::size_t cells = overflow ? 0 : height * width;
        if (overflow || cells > options.maxCells) {
            std::ostringstream message;
            message << "Worksheet dense range requires " << height << " x " << width
                    << " = ";
            if (overflow) message << "overflow";
            else message << cells;
            message << " cells, exceeding max_cells=" << options.maxCells;
            throw std::runtime_error(message.str());
        }
    }

    std::size_t cropEndRow() const {
        if (!options.nrows || !cropOriginRow) {
            return std::numeric_limits<std::size_t>::max();
        }
        const auto count = *options.nrows;
        if (count == 0) return *cropOriginRow - 1;
        if (*cropOriginRow > std::numeric_limits<std::size_t>::max() - (count - 1)) {
            return std::numeric_limits<std::size_t>::max();
        }
        return *cropOriginRow + count - 1;
    }

    bool shouldParseRow(int rowNumber) const {
        if (completed || rowNumber <= 0) return false;
        if (!options.nrows) return true;
        const auto row = static_cast<std::size_t>(rowNumber);
        const bool allowed = !options.skipEmptyArea
            ? row <= *options.nrows
            : !cropOriginRow || row <= cropEndRow();
        if (!allowed) limitBoundaryObserved = true;
        return allowed;
    }

    void beginRow(int rowNumber) {
        if (rowNumber <= 0) {
            errors.emplace_back("Worksheet row number must be positive");
            rowOpen = false;
            return;
        }
        if (rowNumber <= lastSeenRow) {
            errors.push_back(
                "Worksheet rows are not in ascending order at row " +
                std::to_string(rowNumber));
            rowOpen = false;
            return;
        }
        lastSeenRow = rowNumber;
        if (!shouldParseRow(rowNumber)) {
            completed = true;
            rowOpen = false;
            return;
        }
        currentRowNumber = rowNumber;
        pendingReserve = 0;
        rowOpen = true;
    }

    void reserveColumns(std::size_t hint) {
        if (!rowOpen || options.skipEmptyArea || hint == 0) return;
        const auto height = static_cast<std::size_t>(currentRowNumber);
        pendingReserve = std::min(
            {hint, options.maxCells / height, static_cast<std::size_t>(16'384)});
    }

    void addValue(int columnNumber, TypedCellValue&& value) {
        if (!rowOpen || columnNumber <= 0) {
            errors.emplace_back("Invalid primitive cell column");
            return;
        }
        const auto row = static_cast<std::size_t>(currentRowNumber);
        const auto column = static_cast<std::size_t>(columnNumber);

        if (options.skipEmptyArea) {
            const auto nextMinRow = cropMinRow ? std::min(*cropMinRow, row) : row;
            const auto nextMaxRow = cropMaxRow ? std::max(*cropMaxRow, row) : row;
            const auto nextMinColumn = cropMinColumn ? std::min(*cropMinColumn, column) : column;
            const auto nextMaxColumn = cropMaxColumn ? std::max(*cropMaxColumn, column) : column;
            if (!cropOriginRow) cropOriginRow = row;
            if (row > cropEndRow()) {
                completed = true;
                rowOpen = false;
                return;
            }
            checkCellLimit(
                inclusiveSpan(nextMinRow, nextMaxRow),
                inclusiveSpan(nextMinColumn, nextMaxColumn));
            if (sparseRows.empty() || sparseRows.back().number != currentRowNumber) {
                sparseRows.push_back({currentRowNumber, {}});
            }
            sparseRows.back().values.emplace_back(column, std::move(value));
            cropMinRow = nextMinRow;
            cropMaxRow = nextMaxRow;
            cropMinColumn = nextMinColumn;
            cropMaxColumn = nextMaxColumn;
            columnCount = inclusiveSpan(*cropMinColumn, *cropMaxColumn);
            return;
        }

        const auto nextLastRow = std::max(lastCellRow, row);
        const auto nextColumnCount = std::max(columnCount, column);
        checkCellLimit(nextLastRow, nextColumnCount);
        if (denseRows.empty() || denseRows.back().number != currentRowNumber) {
            denseRows.push_back({currentRowNumber, {}});
            if (pendingReserve > 0) denseRows.back().values.reserve(pendingReserve);
        }
        auto& current = denseRows.back().values;
        if (current.size() < column) current.resize(column);
        current[column - 1] = std::move(value);
        lastCellRow = nextLastRow;
        columnCount = nextColumnCount;
    }

    void addCell(const core::CellData& cell) {
        if (!rowOpen) {
            return;
        }
        if (cell.coordinate.row != currentRowNumber || cell.coordinate.column <= 0) {
            errors.push_back("Invalid cell coordinate: " + cell.coordinate.toReference());
            return;
        }
        addValue(
            cell.coordinate.column,
            convertCell(cell, sharedStrings, errors));
    }

    void finishRow() {
        if (!rowOpen) return;
        rowOpen = false;
        if (!options.nrows) return;
        const auto row = static_cast<std::size_t>(currentRowNumber);
        if (!options.skipEmptyArea) {
            completed = row >= *options.nrows;
        } else if (cropOriginRow) {
            completed = row >= cropEndRow();
        }
    }

    void finalize() {
        if (finalized) return;
        finalized = true;

        if (options.skipEmptyArea) {
            if (!cropMinRow || !cropMaxRow || !cropMinColumn || !cropMaxColumn) return;
            auto outputMaxRow = *cropMaxRow;
            if (options.nrows &&
                (limitBoundaryObserved ||
                 static_cast<std::size_t>(lastSeenRow) >= cropEndRow())) {
                outputMaxRow = std::max(outputMaxRow, cropEndRow());
            }
            const auto height = inclusiveSpan(*cropMinRow, outputMaxRow);
            const auto width = inclusiveSpan(*cropMinColumn, *cropMaxColumn);
            checkCellLimit(height, width);
            rows.assign(height, TypedRow(width));
            for (auto& stored : sparseRows) {
                auto& output = rows[static_cast<std::size_t>(stored.number) - *cropMinRow];
                for (auto& [column, value] : stored.values) {
                    output[column - *cropMinColumn] = std::move(value);
                }
            }
            return;
        }

        if (lastCellRow == 0 || columnCount == 0) return;
        auto outputLastRow = std::max(
            lastCellRow, static_cast<std::size_t>(lastSeenRow));
        if (options.nrows && limitBoundaryObserved) {
            outputLastRow = std::max(outputLastRow, *options.nrows);
        }
        checkCellLimit(outputLastRow, columnCount);
        rows.assign(outputLastRow, TypedRow{});
        for (auto& stored : denseRows) {
            const auto rowIndex = static_cast<std::size_t>(stored.number - 1);
            if (rowIndex < rows.size()) rows[rowIndex] = std::move(stored.values);
        }
        for (auto& row : rows) row.resize(columnCount);
    }

    const core::SharedStringsProvider* sharedStrings;
    TypedReadOptions options;
    TypedWorksheet rows;
    std::vector<DenseStoredRow> denseRows;
    std::vector<SparseStoredRow> sparseRows;
    std::optional<std::size_t> cropOriginRow;
    std::optional<std::size_t> cropMinRow;
    std::optional<std::size_t> cropMaxRow;
    std::optional<std::size_t> cropMinColumn;
    std::optional<std::size_t> cropMaxColumn;
    int currentRowNumber = 0;
    int lastSeenRow = 0;
    std::size_t lastCellRow = 0;
    std::size_t columnCount = 0;
    std::size_t pendingReserve = 0;
    bool rowOpen = false;
    bool completed = false;
    mutable bool limitBoundaryObserved = false;
    bool finalized = false;
    std::vector<std::string> errors;
};

TypedRowCollector::TypedRowCollector(
    const core::SharedStringsProvider* sharedStrings,
    TypedReadOptions options)
    : m_impl(std::make_unique<Impl>(sharedStrings, std::move(options))) {}

TypedRowCollector::~TypedRowCollector() = default;
TypedRowCollector::TypedRowCollector(TypedRowCollector&&) noexcept = default;
TypedRowCollector& TypedRowCollector::operator=(TypedRowCollector&&) noexcept = default;

void TypedRowCollector::handleRow(const core::RowData& row) {
    m_impl->beginRow(row.rowNumber);
    for (const auto& cell : row.cells) {
        m_impl->addCell(cell);
    }
    m_impl->finishRow();
}

void TypedRowCollector::handleError(const std::string& message) {
    m_impl->errors.push_back(message);
}

bool TypedRowCollector::acceptsStreamingCells() const {
    return true;
}

void TypedRowCollector::beginRow(int rowNumber, [[maybe_unused]] bool hidden) {
    m_impl->beginRow(rowNumber);
}

void TypedRowCollector::handleCell(core::CellData&& cell) {
    m_impl->addCell(cell);
}

void TypedRowCollector::endRow() {
    m_impl->finishRow();
}

void TypedRowCollector::beginPrimitiveRow(
    int rowNumber, [[maybe_unused]] bool hidden, std::size_t columnReserveHint) {
    m_impl->beginRow(rowNumber);
    m_impl->reserveColumns(columnReserveHint);
}

void TypedRowCollector::addEmpty(int column) {
    m_impl->addValue(column, std::monostate{});
}

void TypedRowCollector::addBoolean(int column, bool value) {
    m_impl->addValue(column, value);
}

void TypedRowCollector::addNumber(int column, double value) {
    m_impl->addValue(column, value);
}

void TypedRowCollector::addString(int column, std::string&& value) {
    m_impl->addValue(column, std::move(value));
}

void TypedRowCollector::addSharedString(int column, std::size_t index) {
    if (m_impl->sharedStrings) {
        auto value = m_impl->sharedStrings->tryGetString(index);
        if (value) {
            m_impl->addValue(column, std::move(*value));
            return;
        }
    }
    m_impl->errors.push_back(
        "Shared string index out of range: " + std::to_string(index));
    m_impl->addValue(column, std::monostate{});
}

void TypedRowCollector::endPrimitiveRow() {
    m_impl->finishRow();
}

bool TypedRowCollector::shouldParseRow(int rowNumber) const {
    return m_impl->shouldParseRow(rowNumber);
}

bool TypedRowCollector::shouldContinueParsing() const {
    return !m_impl->completed;
}

TypedWorksheet TypedRowCollector::takeRows() {
    m_impl->finalize();
    return std::move(m_impl->rows);
}

const std::vector<std::string>& TypedRowCollector::getErrors() const {
    return m_impl->errors;
}

std::size_t TypedRowCollector::getColumnCount() const {
    return m_impl->columnCount;
}

} // namespace xlsxcsv::internal
