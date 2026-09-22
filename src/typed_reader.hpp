#pragma once

#include "xlsxcsv/core.hpp"
#include "core/primitive_cell_handler.hpp"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <variant>
#include <vector>
#include <mutex>

namespace xlsxcsv::internal {

struct TypedDateTime {
    int year;
    unsigned month;
    unsigned day;
    int hour;
    int minute;
    int second;
    int microsecond;
};

struct TypedTime {
    int hour;
    int minute;
    int second;
    int microsecond;
};

struct PendingStyledNumber {
    double value;
    int styleIndex;
};

using TypedCellValue = std::variant<
    std::monostate, bool, std::int64_t, double, std::string,
    TypedDateTime, TypedTime, PendingStyledNumber>;
using TypedRow = std::vector<TypedCellValue>;
using TypedWorksheet = std::vector<TypedRow>;

struct TypedReadOptions {
    bool skipEmptyArea = false;
    std::optional<std::size_t> nrows;
    std::size_t maxCells = 10'000'000;
};

// Internal owner used by the Python facade.  It deliberately keeps the ZIP
// package and its derived providers alive so multiple sheet reads do not reopen
// the workbook.
class TypedWorkbookSession {
public:
    explicit TypedWorkbookSession(const std::string& path, std::size_t maxCells);
    explicit TypedWorkbookSession(core::ByteVector data, std::size_t maxCells);
    ~TypedWorkbookSession();
    TypedWorkbookSession(const TypedWorkbookSession&) = delete;
    TypedWorkbookSession& operator=(const TypedWorkbookSession&) = delete;

    std::vector<core::SheetInfo> sheets() const;
    std::optional<core::SheetInfo> sheetByName(const std::string& name) const;
    std::optional<core::SheetInfo> sheetByIndex(int index) const;
    TypedWorksheet read(const core::SheetInfo& sheet, TypedReadOptions options);
    void close();
    bool isOpen() const;

private:
    void open();
    void ensureSharedStrings();
    void ensureStyles();
    mutable std::mutex m_mutex;
    core::OpcPackage m_package;
    core::Workbook m_workbook;
    core::SharedStringsProvider m_sharedStrings;
    core::StylesRegistry m_styles;
    bool m_sharedStringsInitialized = false;
    bool m_stylesInitialized = false;
    bool m_open = false;
    std::size_t m_maxCells;
};

class TypedRowCollector final : public core::SheetRowHandler,
                                public core::SheetCellHandler,
                                public PrimitiveCellHandler,
                                public WorksheetRowControl {
public:
    explicit TypedRowCollector(
        const core::SharedStringsProvider* sharedStrings = nullptr,
        TypedReadOptions options = {},
        const core::StylesRegistry* styles = nullptr,
        core::DateSystem dateSystem = core::DateSystem::Date1900);
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
    void addNumber(int column, double value, int styleIndex) override;
    void addError(int column) override;
    void addString(int column, std::string&& value) override;
    void addSharedString(int column, std::size_t index) override;
    void endPrimitiveRow() override;

    bool shouldParseRow(int rowNumber) const override;
    bool shouldContinueParsing() const override;

    bool hasStyledNumbers() const;
    void setStyles(const core::StylesRegistry* styles);

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
