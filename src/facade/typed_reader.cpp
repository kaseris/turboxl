#include "typed_reader.hpp"
#include "core/fast_typed_sheet_reader.hpp"

#include <optional>
#include <sstream>
#include <stdexcept>

namespace xlsxcsv::internal {

namespace {

std::optional<core::SheetInfo> selectSheet(
    const core::Workbook& workbook,
    const std::variant<std::string, int>& selector) {
    if (std::holds_alternative<std::string>(selector)) {
        return workbook.findSheet(std::get<std::string>(selector));
    }
    const int index = std::get<int>(selector);
    return workbook.findSheet(index == -1 ? 0 : index);
}

std::string joinErrors(const std::vector<std::string>& errors) {
    std::ostringstream message;
    for (std::size_t index = 0; index < errors.size(); ++index) {
        if (index != 0) {
            message << "; ";
        }
        message << errors[index];
    }
    return message.str();
}

} // namespace

TypedWorksheet readSheetToTyped(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector,
    const TypedReadOptions& options) {
    try {
        if (options.maxCells == 0) {
            throw std::invalid_argument("max_cells must be greater than zero");
        }
        core::OpcPackage package;
        package.open(xlsxPath);

        core::Workbook workbook;
        workbook.open(package);

        const auto sheet = selectSheet(workbook, sheetSelector);
        if (!sheet) {
            if (std::holds_alternative<std::string>(sheetSelector)) {
                throw std::runtime_error(
                    "Sheet not found: " + std::get<std::string>(sheetSelector));
            }
            throw std::runtime_error(
                "Sheet index out of range: " +
                std::to_string(std::get<int>(sheetSelector)));
        }
        if (options.nrows && *options.nrows == 0) return {};

        core::SharedStringsProvider sharedStrings;
        try {
            sharedStrings.parse(package);
        } catch (const core::XlsxError&) {
            // Workbooks without sharedStrings.xml are valid.
        }

        const auto* strings = sharedStrings.isOpen() ? &sharedStrings : nullptr;
        TypedRowCollector fastCollector(
            strings, options, nullptr, workbook.getDateSystem());
        if (tryReadTypedWorksheetFast(package, sheet->target, fastCollector) &&
            fastCollector.getErrors().empty()) {
            core::StylesRegistry styles;
            if (fastCollector.hasStyledNumbers()) {
                try {
                    styles.parse(package, core::StylesRegistry::ParseMode::CsvOnly);
                    fastCollector.setStyles(&styles);
                } catch (const core::XlsxError&) {
                    // Workbooks without styles.xml are valid.
                }
            }
            return fastCollector.takeRows();
        }

        TypedRowCollector fallbackCollector(
            strings, options, nullptr, workbook.getDateSystem());
        core::SheetStreamReader reader;
        reader.parseSheet(package, sheet->target, fallbackCollector, strings);
        if (!fallbackCollector.getErrors().empty()) {
            throw std::runtime_error(
                "Sheet parsing errors: " + joinErrors(fallbackCollector.getErrors()));
        }
        core::StylesRegistry styles;
        if (fallbackCollector.hasStyledNumbers()) {
            try {
                styles.parse(package, core::StylesRegistry::ParseMode::CsvOnly);
                fallbackCollector.setStyles(&styles);
            } catch (const core::XlsxError&) {
                // Workbooks without styles.xml are valid.
            }
        }
        return fallbackCollector.takeRows();
    } catch (const core::XlsxError& error) {
        throw std::runtime_error("XLSX parsing error: " + std::string(error.what()));
    } catch (const std::exception& error) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(error.what()));
    }
}

} // namespace xlsxcsv::internal
