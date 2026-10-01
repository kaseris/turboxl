#include "core/fast_typed_sheet_reader.hpp"

#include "core/fast_xml.hpp"
#include "typed_reader.hpp"

#include <algorithm>
#include <charconv>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <string_view>
#include <vector>

namespace xlsxcsv::internal {
namespace {

using core::fastxml::Token;
using core::fastxml::TokenKind;
using core::fastxml::XmlCursor;
using core::fastxml::appendXmlText;
using core::fastxml::findAttribute;
using core::fastxml::localName;

bool parsePositiveInt(std::string_view text, int& value) {
    if (text.empty()) return false;
    const auto [end, error] = std::from_chars(text.data(), text.data() + text.size(), value);
    return error == std::errc{} && end == text.data() + text.size() && value > 0;
}

bool parseNonnegativeInt(std::string_view text, int& value) {
    if (text.empty()) return false;
    const auto [end, error] = std::from_chars(text.data(), text.data() + text.size(), value);
    return error == std::errc{} && end == text.data() + text.size() && value >= 0;
}

bool parseColumn(std::string_view reference, int& column) {
    column = 0;
    std::size_t position = 0;
    while (position < reference.size() && reference[position] >= 'A' && reference[position] <= 'Z') {
        if (column > (std::numeric_limits<int>::max() - 26) / 26) return false;
        column = column * 26 + reference[position] - 'A' + 1;
        ++position;
    }
    return column > 0 && position < reference.size() &&
           reference[position] >= '1' && reference[position] <= '9';
}

core::CellType parseType(std::string_view tag) {
    std::string_view value;
    if (!findAttribute(tag, "t", value) || value == "n") return core::CellType::Number;
    if (value == "b") return core::CellType::Boolean;
    if (value == "e") return core::CellType::Error;
    if (value == "s") return core::CellType::SharedString;
    if (value == "str") return core::CellType::String;
    if (value == "inlineStr") return core::CellType::InlineString;
    return core::CellType::Unknown;
}

bool emitCell(
    TypedRowCollector& collector,
    int column,
    core::CellType type,
    int styleIndex,
    bool hasValue,
    std::string&& value) {
    if (column <= 0) return false;
    if (!hasValue) {
        collector.addEmpty(column);
        return true;
    }
    switch (type) {
        case core::CellType::Boolean:
            collector.addBoolean(column, value == "1");
            return true;
        case core::CellType::Number: {
            char* end = nullptr;
            const double number = std::strtod(value.c_str(), &end);
            if (end != value.data() + value.size()) {
                collector.addEmpty(column);
            } else {
                collector.addNumber(column, number, styleIndex);
            }
            return true;
        }
        case core::CellType::SharedString: {
            std::size_t index = 0;
            const auto [end, error] = std::from_chars(
                value.data(), value.data() + value.size(), index);
            if (error != std::errc{} || end != value.data() + value.size()) {
                collector.addEmpty(column);
            } else {
                collector.addSharedString(column, index);
            }
            return true;
        }
        case core::CellType::Error:
            collector.addError(column);
            return true;
        case core::CellType::String:
        case core::CellType::InlineString:
        case core::CellType::Unknown:
            collector.addString(column, std::move(value));
            return true;
    }
    return false;
}

std::string normalizedSheetPath(std::string path) {
    if (!path.empty() && path.front() == '/') {
        path.erase(0, 1);
    } else if (path.rfind("xl/", 0) != 0) {
        path.insert(0, "xl/");
    }
    return path;
}

} // namespace

bool tryParseTypedWorksheetFast(
    const std::vector<std::uint8_t>& bytes,
    TypedRowCollector& collector) {
    if (!collector.shouldContinueParsing()) return true;
    if (bytes.empty()) return false;
    const std::string_view xml(
        reinterpret_cast<const char*>(bytes.data()), bytes.size());
    if ((bytes.size() >= 2 && ((bytes[0] == 0xff && bytes[1] == 0xfe) ||
                               (bytes[0] == 0xfe && bytes[1] == 0xff))) ||
        xml.find("<!DOCTYPE") != std::string_view::npos) {
        return false;
    }

    const auto declarationEnd = xml.substr(0, std::min<std::size_t>(xml.size(), 256)).find("?>");
    if (xml.rfind("<?xml", 0) == 0 && declarationEnd != std::string_view::npos) {
        const auto declaration = xml.substr(0, declarationEnd);
        const auto encoding = declaration.find("encoding");
        if (encoding != std::string_view::npos) {
            const auto quote = declaration.find_first_of("\"'", encoding + 8);
            if (quote == std::string_view::npos) return false;
            const auto endQuote = declaration.find(declaration[quote], quote + 1);
            if (endQuote == std::string_view::npos) return false;
            const auto value = declaration.substr(quote + 1, endQuote - quote - 1);
            if (value != "UTF-8" && value != "utf-8" && value != "UTF8" && value != "utf8") {
                return false;
            }
        }
    }

    XmlCursor cursor(xml);
    bool inSheetData = false;
    bool sawWorksheet = false;
    bool inRow = false;
    bool inCell = false;
    bool captureValue = false;
    bool captureInlineText = false;
    bool hasValue = false;
    int previousRow = 0;
    int rowNumber = 0;
    int column = 0;
    int styleIndex = 0;
    core::CellType type = core::CellType::Unknown;
    std::string value;
    std::vector<std::string_view> elements;

    while (true) {
        const Token token = cursor.next();
        if (token.kind == TokenKind::Invalid) return false;
        if (token.kind == TokenKind::Eof) break;

        if (token.kind == TokenKind::Start) {
            if (!token.empty) elements.push_back(token.name);
            if (!sawWorksheet && token.name == "worksheet") {
                sawWorksheet = true;
            }
            if (!inSheetData && token.name == "sheetData") {
                inSheetData = !token.empty;
                continue;
            }
            if (!inSheetData) continue;
            if (!inRow && token.name == "row") {
                std::string_view rowReference;
                rowNumber = previousRow + 1;
                if (findAttribute(token.raw, "r", rowReference) &&
                    !parsePositiveInt(rowReference, rowNumber)) {
                    return false;
                }
                if (!collector.shouldParseRow(rowNumber)) return true;
                std::size_t reserveHint = 0;
                std::string_view spans;
                if (findAttribute(token.raw, "spans", spans)) {
                    const auto colon = spans.find(':');
                    int last = 0;
                    if (colon != std::string_view::npos &&
                        parsePositiveInt(spans.substr(colon + 1), last)) {
                        reserveHint = static_cast<std::size_t>(last);
                    }
                }
                collector.beginPrimitiveRow(rowNumber, false, reserveHint);
                previousRow = rowNumber;
                inRow = !token.empty;
                if (token.empty) collector.endPrimitiveRow();
                continue;
            }
            if (inRow && !inCell && token.name == "c") {
                std::string_view reference;
                if (!findAttribute(token.raw, "r", reference) ||
                    !parseColumn(reference, column)) {
                    return false;
                }
                type = parseType(token.raw);
                std::string_view style;
                styleIndex = 0;
                if (findAttribute(token.raw, "s", style) &&
                    !parseNonnegativeInt(style, styleIndex)) {
                    return false;
                }
                value.clear();
                hasValue = false;
                inCell = !token.empty;
                if (token.empty) {
                    collector.addEmpty(column);
                }
                continue;
            }
            if (inCell && token.name == "v") {
                captureValue = !token.empty;
                hasValue = true;
                continue;
            }
            if (inCell && type == core::CellType::InlineString && token.name == "t") {
                captureInlineText = !token.empty;
                hasValue = true;
                continue;
            }
        } else if (token.kind == TokenKind::End) {
            if (elements.empty() || elements.back() != token.name) return false;
            elements.pop_back();
            if (inCell && token.name == "v") {
                captureValue = false;
            } else if (inCell && token.name == "t") {
                captureInlineText = false;
            } else if (inCell && token.name == "c") {
                if (!emitCell(
                        collector, column, type, styleIndex, hasValue, std::move(value))) {
                    return false;
                }
                value.clear();
                inCell = false;
                captureValue = false;
                captureInlineText = false;
            } else if (inRow && token.name == "row") {
                if (inCell) return false;
                collector.endPrimitiveRow();
                inRow = false;
                if (!collector.shouldContinueParsing()) return true;
            } else if (inSheetData && token.name == "sheetData") {
                if (inRow || inCell) return false;
                inSheetData = false;
            }
        } else if ((token.kind == TokenKind::Text || token.kind == TokenKind::CData) &&
                   inCell && (captureValue || captureInlineText)) {
            if (!appendXmlText(value, token.raw, token.kind == TokenKind::CData)) return false;
        }
    }
    return sawWorksheet && elements.empty() && !inSheetData && !inRow && !inCell;
}

bool tryReadTypedWorksheetFast(
    const core::OpcPackage& package,
    const std::string& sheetPath,
    TypedRowCollector& collector) {
    if (!collector.shouldContinueParsing()) return true;
    const auto bytes = package.getZipReader().readEntry(normalizedSheetPath(sheetPath));
    return tryParseTypedWorksheetFast(bytes, collector);
}

} // namespace xlsxcsv::internal
