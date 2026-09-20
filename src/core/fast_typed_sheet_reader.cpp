#include "core/fast_typed_sheet_reader.hpp"

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

enum class TokenKind { Start, End, Text, CData, Eof, Invalid };

struct Token {
    TokenKind kind = TokenKind::Invalid;
    std::string_view name;
    std::string_view raw;
    bool empty = false;
};

std::string_view localName(std::string_view name) {
    const auto colon = name.rfind(':');
    return colon == std::string_view::npos ? name : name.substr(colon + 1);
}

bool isNameCharacter(char value) {
    return (value >= 'a' && value <= 'z') ||
           (value >= 'A' && value <= 'Z') ||
           (value >= '0' && value <= '9') ||
           value == '_' || value == '-' || value == ':' || value == '.';
}

class XmlCursor {
public:
    explicit XmlCursor(std::string_view xml) : xml(xml) {}

    Token next() {
        while (position < xml.size()) {
            if (xml[position] != '<') {
                const auto end = xml.find('<', position);
                const auto finish = end == std::string_view::npos ? xml.size() : end;
                Token token{TokenKind::Text, {}, xml.substr(position, finish - position)};
                position = finish;
                return token;
            }
            if (xml.substr(position, 4) == "<!--") {
                const auto end = xml.find("-->", position + 4);
                if (end == std::string_view::npos) return invalid();
                position = end + 3;
                continue;
            }
            if (xml.substr(position, 2) == "<?") {
                const auto end = xml.find("?>", position + 2);
                if (end == std::string_view::npos) return invalid();
                position = end + 2;
                continue;
            }
            if (xml.substr(position, 9) == "<![CDATA[") {
                const auto end = xml.find("]]>", position + 9);
                if (end == std::string_view::npos) return invalid();
                Token token{
                    TokenKind::CData, {}, xml.substr(position + 9, end - position - 9)};
                position = end + 3;
                return token;
            }
            if (xml.substr(position, 2) == "<!") {
                // DTDs and custom entities are intentionally handled by libxml2.
                return invalid();
            }

            const auto start = position;
            bool closing = false;
            ++position;
            if (position < xml.size() && xml[position] == '/') {
                closing = true;
                ++position;
            }
            while (position < xml.size() &&
                   (xml[position] == ' ' || xml[position] == '\t' ||
                    xml[position] == '\r' || xml[position] == '\n')) {
                ++position;
            }
            const auto nameStart = position;
            while (position < xml.size() && isNameCharacter(xml[position])) ++position;
            if (nameStart == position) return invalid();
            const auto name = xml.substr(nameStart, position - nameStart);

            char quote = 0;
            while (position < xml.size()) {
                const char value = xml[position];
                if (quote != 0) {
                    if (value == quote) quote = 0;
                } else if (value == '\'' || value == '"') {
                    quote = value;
                } else if (value == '>') {
                    break;
                }
                ++position;
            }
            if (position == xml.size() || quote != 0) return invalid();
            auto beforeEnd = position;
            while (beforeEnd > start &&
                   (xml[beforeEnd - 1] == ' ' || xml[beforeEnd - 1] == '\t' ||
                    xml[beforeEnd - 1] == '\r' || xml[beforeEnd - 1] == '\n')) {
                --beforeEnd;
            }
            const bool empty = !closing && beforeEnd > start && xml[beforeEnd - 1] == '/';
            ++position;
            return Token{
                closing ? TokenKind::End : TokenKind::Start,
                localName(name),
                xml.substr(start, position - start),
                empty};
        }
        return Token{TokenKind::Eof, {}, {}, false};
    }

private:
    Token invalid() {
        position = xml.size();
        return Token{TokenKind::Invalid, {}, {}, false};
    }

    std::string_view xml;
    std::size_t position = 0;
};

bool findAttribute(
    std::string_view tag, std::string_view wanted, std::string_view& result) {
    std::size_t position = 1;
    if (position < tag.size() && tag[position] == '/') ++position;
    while (position < tag.size() && isNameCharacter(tag[position])) ++position;

    while (position < tag.size()) {
        while (position < tag.size() &&
               (tag[position] == ' ' || tag[position] == '\t' ||
                tag[position] == '\r' || tag[position] == '\n')) {
            ++position;
        }
        if (position >= tag.size() || tag[position] == '>' || tag[position] == '/') break;
        const auto nameStart = position;
        while (position < tag.size() && isNameCharacter(tag[position])) ++position;
        if (nameStart == position) return false;
        const auto name = localName(tag.substr(nameStart, position - nameStart));
        while (position < tag.size() &&
               (tag[position] == ' ' || tag[position] == '\t' ||
                tag[position] == '\r' || tag[position] == '\n')) {
            ++position;
        }
        if (position >= tag.size() || tag[position] != '=') return false;
        ++position;
        while (position < tag.size() &&
               (tag[position] == ' ' || tag[position] == '\t' ||
                tag[position] == '\r' || tag[position] == '\n')) {
            ++position;
        }
        if (position >= tag.size() || (tag[position] != '\'' && tag[position] != '"')) {
            return false;
        }
        const char quote = tag[position++];
        const auto valueStart = position;
        const auto valueEnd = tag.find(quote, position);
        if (valueEnd == std::string_view::npos) return false;
        if (name == wanted) {
            result = tag.substr(valueStart, valueEnd - valueStart);
            return true;
        }
        position = valueEnd + 1;
    }
    return false;
}

bool parsePositiveInt(std::string_view text, int& value) {
    if (text.empty()) return false;
    const auto [end, error] = std::from_chars(text.data(), text.data() + text.size(), value);
    return error == std::errc{} && end == text.data() + text.size() && value > 0;
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

bool appendCodepoint(std::string& output, std::uint32_t codepoint) {
    if (codepoint == 0 || codepoint > 0x10ffff ||
        (codepoint >= 0xd800 && codepoint <= 0xdfff)) {
        return false;
    }
    if (codepoint <= 0x7f) {
        output.push_back(static_cast<char>(codepoint));
    } else if (codepoint <= 0x7ff) {
        output.push_back(static_cast<char>(0xc0 | (codepoint >> 6)));
        output.push_back(static_cast<char>(0x80 | (codepoint & 0x3f)));
    } else if (codepoint <= 0xffff) {
        output.push_back(static_cast<char>(0xe0 | (codepoint >> 12)));
        output.push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3f)));
        output.push_back(static_cast<char>(0x80 | (codepoint & 0x3f)));
    } else {
        output.push_back(static_cast<char>(0xf0 | (codepoint >> 18)));
        output.push_back(static_cast<char>(0x80 | ((codepoint >> 12) & 0x3f)));
        output.push_back(static_cast<char>(0x80 | ((codepoint >> 6) & 0x3f)));
        output.push_back(static_cast<char>(0x80 | (codepoint & 0x3f)));
    }
    return true;
}

bool appendXmlText(std::string& output, std::string_view text, bool cdata) {
    if (cdata || text.find('&') == std::string_view::npos) {
        output.append(text);
        return true;
    }
    std::size_t position = 0;
    while (position < text.size()) {
        const auto ampersand = text.find('&', position);
        if (ampersand == std::string_view::npos) {
            output.append(text.substr(position));
            return true;
        }
        output.append(text.substr(position, ampersand - position));
        const auto semicolon = text.find(';', ampersand + 1);
        if (semicolon == std::string_view::npos) return false;
        const auto entity = text.substr(ampersand + 1, semicolon - ampersand - 1);
        if (entity == "amp") output.push_back('&');
        else if (entity == "lt") output.push_back('<');
        else if (entity == "gt") output.push_back('>');
        else if (entity == "quot") output.push_back('"');
        else if (entity == "apos") output.push_back('\'');
        else if (entity.size() > 1 && entity[0] == '#') {
            std::uint32_t codepoint = 0;
            const bool hexadecimal = entity.size() > 2 && (entity[1] == 'x' || entity[1] == 'X');
            const auto digits = entity.substr(hexadecimal ? 2 : 1);
            if (digits.empty()) return false;
            const auto [end, error] = std::from_chars(
                digits.data(), digits.data() + digits.size(), codepoint, hexadecimal ? 16 : 10);
            if (error != std::errc{} || end != digits.data() + digits.size() ||
                !appendCodepoint(output, codepoint)) {
                return false;
            }
        } else {
            return false;
        }
        position = semicolon + 1;
    }
    return true;
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
                collector.addNumber(column, number);
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
                if (!emitCell(collector, column, type, hasValue, std::move(value))) return false;
                value.clear();
                inCell = false;
                captureValue = false;
                captureInlineText = false;
            } else if (inRow && token.name == "row") {
                if (inCell) return false;
                collector.endPrimitiveRow();
                inRow = false;
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
    const auto bytes = package.getZipReader().readEntry(normalizedSheetPath(sheetPath));
    return tryParseTypedWorksheetFast(bytes, collector);
}

} // namespace xlsxcsv::internal
