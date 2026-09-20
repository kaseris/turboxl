#pragma once

#include "xlsxcsv/core.hpp"

#include <cstddef>
#include <memory>
#include <string>
#include <variant>
#include <vector>

namespace xlsxcsv::internal {

using TypedCellValue = std::variant<std::monostate, bool, double, std::string>;
using TypedRow = std::vector<TypedCellValue>;
using TypedWorksheet = std::vector<TypedRow>;

class TypedRowCollector final : public core::SheetRowHandler,
                                public core::SheetCellHandler {
public:
    explicit TypedRowCollector(const core::SharedStringsProvider* sharedStrings = nullptr);
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

    TypedWorksheet takeRows();
    const std::vector<std::string>& getErrors() const;
    std::size_t getColumnCount() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

TypedWorksheet readSheetToTyped(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector = 0);

} // namespace xlsxcsv::internal
