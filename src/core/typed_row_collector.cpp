#include "typed_reader.hpp"

#include <algorithm>
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
    explicit Impl(const core::SharedStringsProvider* sharedStrings)
        : sharedStrings(sharedStrings) {}

    void beginRow(int rowNumber) {
        if (rowNumber <= 0) {
            errors.emplace_back("Worksheet row number must be positive");
            currentRow = nullptr;
            return;
        }
        if (rowNumber <= lastPhysicalRow) {
            errors.push_back(
                "Worksheet rows are not in ascending order at row " +
                std::to_string(rowNumber));
            currentRow = nullptr;
            return;
        }
        while (lastPhysicalRow + 1 < rowNumber) {
            rows.emplace_back();
            ++lastPhysicalRow;
        }
        rows.emplace_back();
        currentRow = &rows.back();
        currentRowNumber = rowNumber;
        lastPhysicalRow = rowNumber;
    }

    void reserveColumns(std::size_t hint) {
        if (currentRow && hint > 0) {
            currentRow->reserve(hint);
        }
    }

    void addValue(int columnNumber, TypedCellValue&& value) {
        if (!currentRow || columnNumber <= 0) {
            errors.emplace_back("Invalid primitive cell column");
            return;
        }
        const auto column = static_cast<std::size_t>(columnNumber);
        if (currentRow->size() < column) {
            currentRow->resize(column);
        }
        (*currentRow)[column - 1] = std::move(value);
        columnCount = std::max(columnCount, column);
    }

    void addCell(const core::CellData& cell) {
        if (!currentRow) {
            return;
        }
        if (cell.coordinate.row != currentRowNumber || cell.coordinate.column <= 0) {
            errors.push_back("Invalid cell coordinate: " + cell.coordinate.toReference());
            return;
        }
        const auto column = static_cast<std::size_t>(cell.coordinate.column);
        if (currentRow->size() < column) {
            currentRow->resize(column);
        }
        (*currentRow)[column - 1] = convertCell(cell, sharedStrings, errors);
        columnCount = std::max(columnCount, column);
    }

    void finalize() {
        if (finalized) {
            return;
        }
        for (auto& row : rows) {
            row.resize(columnCount);
        }
        finalized = true;
    }

    const core::SharedStringsProvider* sharedStrings;
    TypedWorksheet rows;
    TypedRow* currentRow = nullptr;
    int currentRowNumber = 0;
    int lastPhysicalRow = 0;
    std::size_t columnCount = 0;
    bool finalized = false;
    std::vector<std::string> errors;
};

TypedRowCollector::TypedRowCollector(const core::SharedStringsProvider* sharedStrings)
    : m_impl(std::make_unique<Impl>(sharedStrings)) {}

TypedRowCollector::~TypedRowCollector() = default;
TypedRowCollector::TypedRowCollector(TypedRowCollector&&) noexcept = default;
TypedRowCollector& TypedRowCollector::operator=(TypedRowCollector&&) noexcept = default;

void TypedRowCollector::handleRow(const core::RowData& row) {
    m_impl->beginRow(row.rowNumber);
    for (const auto& cell : row.cells) {
        m_impl->addCell(cell);
    }
    m_impl->currentRow = nullptr;
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
    m_impl->currentRow = nullptr;
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
    m_impl->currentRow = nullptr;
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
