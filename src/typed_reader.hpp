#pragma once

#include "xlsxcsv/core.hpp"
#include "core/primitive_cell_handler.hpp"

#include <cstddef>
#include <memory>
#include <optional>
#include <string>
#include <variant>
#include <vector>

namespace xlsxcsv::internal {

using TypedCellValue = std::variant<std::monostate, bool, double, std::string>;
using TypedRow = std::vector<TypedCellValue>;
using TypedWorksheet = std::vector<TypedRow>;

struct TypedReadOptions {
    bool skipEmptyArea = false;
    std::optional<std::size_t> nrows;
    std::size_t maxCells = 10'000'000;
};

class TypedRowCollector final : public core::SheetRowHandler,
                                public core::SheetCellHandler,
                                public PrimitiveCellHandler,
                                public WorksheetRowControl {
public:
    explicit TypedRowCollector(
        const core::SharedStringsProvider* sharedStrings = nullptr,
        TypedReadOptions options = {});
    ~TypedRowCollector() override;

    TypedRowCollector(const TypedRowCollector&) = delete;
    TypedRowCollector& operator=(const TypedRowCollector&) = delete;
    TypedRowCollector(TypedRowCollector&&) noexcept;
    TypedRowCollector& operator=(TypedRowCollector&&) noexcept;

    void handleRow(const core::RowData& row) override;
    void handleError(const std::string& message) override;
    bool acceptsStreamingCells() const override;
    void beginRow(int rowNumber, bool hidden) override;
    void handleCell(core::CellData&& cell) override;
    void endRow() override;

    void beginPrimitiveRow(
        int rowNumber, bool hidden, std::size_t columnReserveHint) override;
    void addEmpty(int column) override;
    void addBoolean(int column, bool value) override;
    void addNumber(int column, double value) override;
    void addString(int column, std::string&& value) override;
    void addSharedString(int column, std::size_t index) override;
    void endPrimitiveRow() override;

    bool shouldParseRow(int rowNumber) const override;
    bool shouldContinueParsing() const override;

    TypedWorksheet takeRows();
    const std::vector<std::string>& getErrors() const;
    std::size_t getColumnCount() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

TypedWorksheet readSheetToTyped(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector = 0,
    const TypedReadOptions& options = {});

} // namespace xlsxcsv::internal
