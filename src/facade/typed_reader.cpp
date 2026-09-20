#include "typed_reader.hpp"

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
    const std::variant<std::string, int>& sheetSelector) {
    try {
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

        core::SharedStringsProvider sharedStrings;
        try {
            sharedStrings.parse(package);
        } catch (const core::XlsxError&) {
            // Workbooks without sharedStrings.xml are valid.
        }

        TypedRowCollector collector(sharedStrings.isOpen() ? &sharedStrings : nullptr);
        core::SheetStreamReader reader;
        reader.parseSheet(
            package, sheet->target, collector,
            sharedStrings.isOpen() ? &sharedStrings : nullptr);
        if (!collector.getErrors().empty()) {
            throw std::runtime_error(
                "Sheet parsing errors: " + joinErrors(collector.getErrors()));
        }
        return collector.takeRows();
    } catch (const core::XlsxError& error) {
        throw std::runtime_error("XLSX parsing error: " + std::string(error.what()));
    } catch (const std::exception& error) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(error.what()));
    }
}

} // namespace xlsxcsv::internal
