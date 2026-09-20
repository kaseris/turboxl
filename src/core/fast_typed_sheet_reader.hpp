#pragma once

#include "xlsxcsv/core.hpp"

#include <string>
#include <vector>

namespace xlsxcsv::internal {

class TypedRowCollector;

// Returns false when the document uses XML features outside the deliberately
// small worksheet fast path. Callers must then use SheetStreamReader.
bool tryParseTypedWorksheetFast(
    const std::vector<std::uint8_t>& xml,
    TypedRowCollector& collector);

bool tryReadTypedWorksheetFast(
    const core::OpcPackage& package,
    const std::string& sheetPath,
    TypedRowCollector& collector);

} // namespace xlsxcsv::internal
