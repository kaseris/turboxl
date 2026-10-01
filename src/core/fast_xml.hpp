#pragma once

// The lexer sits in the worksheet scanner's per-cell loop. Now that several
// translation units call it, the optimizer no longer inlines it there on its
// own (measured ~3% slower on dense worksheets), so inlining is forced.
#if defined(_MSC_VER)
#define XLSXCSV_FASTXML_INLINE __forceinline
#else
#define XLSXCSV_FASTXML_INLINE inline __attribute__((always_inline))
#endif

// Minimal non-validating XML lexer shared by the worksheet, styles and shared
// strings fast paths. It deliberately rejects DTDs and custom entities so
// callers can fall back to libxml2 for anything outside this subset.

#include <charconv>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

namespace xlsxcsv::core::fastxml {

enum class TokenKind { Start, End, Text, CData, Eof, Invalid };

struct Token {
    TokenKind kind = TokenKind::Invalid;
    std::string_view name;
    std::string_view raw;
    bool empty = false;
};

inline std::string_view localName(std::string_view name) {
    const auto colon = name.rfind(':');
    return colon == std::string_view::npos ? name : name.substr(colon + 1);
}

inline bool isNameCharacter(char value) {
    return (value >= 'a' && value <= 'z') ||
           (value >= 'A' && value <= 'Z') ||
           (value >= '0' && value <= '9') ||
           value == '_' || value == '-' || value == ':' || value == '.';
}

class XmlCursor {
public:
    explicit XmlCursor(std::string_view xml) : xml(xml) {}

    XLSXCSV_FASTXML_INLINE Token next() {
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

enum class AttributeLookup { Found, Missing, Malformed };

// Looks up an attribute in a start tag. With `Exact`, `wanted` is compared
// with the qualified attribute name; otherwise only the local name is. The
// choice is a template parameter so the per-cell worksheet path pays nothing
// for it.
template <bool Exact = false>
inline AttributeLookup lookupAttribute(
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
        if (nameStart == position) return AttributeLookup::Malformed;
        const auto qualified = tag.substr(nameStart, position - nameStart);
        const auto name = Exact ? qualified : localName(qualified);
        while (position < tag.size() &&
               (tag[position] == ' ' || tag[position] == '\t' ||
                tag[position] == '\r' || tag[position] == '\n')) {
            ++position;
        }
        if (position >= tag.size() || tag[position] != '=') return AttributeLookup::Malformed;
        ++position;
        while (position < tag.size() &&
               (tag[position] == ' ' || tag[position] == '\t' ||
                tag[position] == '\r' || tag[position] == '\n')) {
            ++position;
        }
        if (position >= tag.size() || (tag[position] != '\'' && tag[position] != '"')) {
            return AttributeLookup::Malformed;
        }
        const char quote = tag[position++];
        const auto valueStart = position;
        const auto valueEnd = tag.find(quote, position);
        if (valueEnd == std::string_view::npos) return AttributeLookup::Malformed;
        if (name == wanted) {
            result = tag.substr(valueStart, valueEnd - valueStart);
            return AttributeLookup::Found;
        }
        position = valueEnd + 1;
    }
    return AttributeLookup::Missing;
}

inline bool findAttribute(
    std::string_view tag, std::string_view wanted, std::string_view& result) {
    return lookupAttribute(tag, wanted, result) == AttributeLookup::Found;
}

inline bool appendCodepoint(std::string& output, std::uint32_t codepoint) {
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

inline bool appendXmlText(std::string& output, std::string_view text, bool cdata) {
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

// Qualified element name of a start or end tag such as "<x:sheet a='1'/>".
inline std::string_view qualifiedName(std::string_view tag) {
    std::size_t begin = 1;
    if (begin < tag.size() && tag[begin] == '/') ++begin;
    while (begin < tag.size() &&
           (tag[begin] == ' ' || tag[begin] == '\t' || tag[begin] == '\r' || tag[begin] == '\n')) {
        ++begin;
    }
    std::size_t end = begin;
    while (end < tag.size() && isNameCharacter(tag[end])) ++end;
    return tag.substr(begin, end - begin);
}

// Reads an attribute by exact qualified name into `out`, with entities
// decoded. `present` reports whether it exists. Returns false when libxml2
// must handle the document: malformed syntax, unsupported entities, or
// literal whitespace that XML attribute-value normalization would rewrite.
inline bool readAttribute(
    std::string_view tag, std::string_view name, std::string& out, bool& present) {
    std::string_view raw;
    switch (lookupAttribute<true>(tag, name, raw)) {
        case AttributeLookup::Missing:
            present = false;
            return true;
        case AttributeLookup::Malformed:
            return false;
        case AttributeLookup::Found:
            break;
    }
    present = true;
    if (raw.find_first_of("\t\r\n") != std::string_view::npos) return false;
    out.clear();
    return appendXmlText(out, raw, false);
}

// Verifies that end tags close the elements opened by start tags, as a
// well-formedness check libxml2 would otherwise enforce.
class TagStack {
public:
    void open(std::string_view qualified) { names.push_back(qualified); }
    bool close(std::string_view qualified) {
        if (names.empty() || names.back() != qualified) return false;
        names.pop_back();
        return true;
    }
    bool empty() const { return names.empty(); }

private:
    std::vector<std::string_view> names;
};

// Calls onStart(qualifiedName, rawTag) for every start tag. Returns false when
// the document is malformed or onStart declines it, so callers can fall back
// to libxml2.
template <class OnStart>
bool forEachStartTag(std::string_view xml, OnStart&& onStart) {
    XmlCursor cursor(xml);
    TagStack stack;
    bool sawRoot = false;
    for (;;) {
        const auto token = cursor.next();
        switch (token.kind) {
            case TokenKind::Invalid:
                return false;
            case TokenKind::Eof:
                return sawRoot && stack.empty();
            case TokenKind::Text:
            case TokenKind::CData:
                break;
            case TokenKind::End:
                if (!stack.close(qualifiedName(token.raw))) return false;
                break;
            case TokenKind::Start: {
                sawRoot = true;
                const auto qualified = qualifiedName(token.raw);
                if (!onStart(qualified, token.raw)) return false;
                if (!token.empty) stack.open(qualified);
                break;
            }
        }
    }
}

// Reads every <Relationship Id Type Target> of an OPC relationships part.
// onRelationship(id, type, target) is called only for elements carrying all
// three attributes, matching the libxml2-based readers.
template <class OnRelationship>
bool parseRelationshipsFast(std::string_view xml, OnRelationship&& onRelationship) {
    std::string id, type, target;
    return forEachStartTag(xml, [&](std::string_view name, std::string_view tag) {
        if (name != "Relationship") return true;
        bool hasId = false, hasType = false, hasTarget = false;
        if (!readAttribute(tag, "Id", id, hasId) ||
            !readAttribute(tag, "Type", type, hasType) ||
            !readAttribute(tag, "Target", target, hasTarget)) {
            return false;
        }
        if (hasId && hasType && hasTarget) onRelationship(id, type, target);
        return true;
    });
}

} // namespace xlsxcsv::core::fastxml
