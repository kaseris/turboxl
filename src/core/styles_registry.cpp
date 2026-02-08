#include "xlsxcsv/core.hpp"
#include <libxml/xmlreader.h>
#include <map>
#include <vector>
#include <regex>
#include <algorithm>

namespace xlsxcsv::core {

class StylesRegistry::Impl {
public:
    Impl() = default;
    
    ~Impl() {
        close();
    }
    
    void parse(const OpcPackage& package) {
        if (m_isOpen) {
            close();
        }
        
        // Find styles.xml path through relationships
        std::string stylesPath = "xl/styles.xml";
        
        if (!package.getZipReader().hasEntry(stylesPath)) {
            throw XlsxError("Missing styles.xml in XLSX package");
        }
        
        auto xmlData = package.getZipReader().readEntry(stylesPath);
        parseStylesXml(xmlData);
        
        m_isOpen = true;
    }
    
    void close() {
        m_isOpen = false;
        m_numberFormats.clear();
        m_fonts.clear();
        m_fills.clear();
        m_borders.clear();
        m_cellStyles.clear();
        m_dateTimeStyleMask.clear();
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::optional<CellStyle> getCellStyle(int styleIndex) const {
        if (!m_isOpen || styleIndex < 0 || static_cast<size_t>(styleIndex) >= m_cellStyles.size()) {
            return std::nullopt;
        }
        return m_cellStyles[styleIndex];
    }
    
    std::optional<NumberFormat> getNumberFormat(int formatId) const {
        if (!m_isOpen) {
            return std::nullopt;
        }
        
        auto it = m_numberFormats.find(formatId);
        if (it != m_numberFormats.end()) {
            return it->second;
        }
        
        // Check for built-in formats
        return getBuiltInNumberFormat(formatId);
    }
    
    NumberFormatType detectNumberFormatType(const std::string& formatCode) const {
        if (formatCode.empty() || formatCode == "General") {
            return NumberFormatType::General;
        }
        
        // Date/time patterns - be more specific about date patterns
        bool hasDatePattern = std::regex_search(formatCode, std::regex(R"([yY])")); // Year
        hasDatePattern = hasDatePattern || std::regex_search(formatCode, std::regex(R"(d+)")); // Day (dd, ddd, etc)
        // Month pattern (MM, MMM) but not AM/PM 
        hasDatePattern = hasDatePattern || (std::regex_search(formatCode, std::regex(R"(M+)")) && 
                                           formatCode.find("AM/PM") == std::string::npos);
        
        bool hasTimePattern = std::regex_search(formatCode, std::regex(R"([hHsS])")); // Hour, second
        hasTimePattern = hasTimePattern || (std::regex_search(formatCode, std::regex(R"(m+)")) && 
                                           std::regex_search(formatCode, std::regex(R"([hHsS])"))); // Minutes with hours/seconds
        
        if (hasDatePattern && hasTimePattern) {
            return NumberFormatType::DateTime;
        }
        
        if (hasDatePattern) {
            return NumberFormatType::Date;
        }
        
        if (hasTimePattern) {
            return NumberFormatType::Time;
        }
        
        // Percentage
        if (formatCode.find('%') != std::string::npos) {
            return NumberFormatType::Percentage;
        }
        
        // Currency/accounting
        if (formatCode.find('$') != std::string::npos || 
            formatCode.find("\xC2\xA4") != std::string::npos ||  // UTF-8 for ¤
            std::regex_search(formatCode, std::regex(R"(\[Currency\])"))) {
            return NumberFormatType::Currency;
        }
        
        // Scientific notation
        if (std::regex_search(formatCode, std::regex(R"([eE][+-])"))) {
            return NumberFormatType::Scientific;
        }
        
        // Fraction
        if (formatCode.find('/') != std::string::npos) {
            return NumberFormatType::Fraction;
        }
        
        // Text format
        if (formatCode.find('@') != std::string::npos) {
            return NumberFormatType::Text;
        }
        
        // Decimal numbers (has decimal point)
        if (formatCode.find('.') != std::string::npos) {
            return NumberFormatType::Decimal;
        }
        
        // Integer (contains digits, zeros, formatting, but no decimal point)
        if (std::regex_search(formatCode, std::regex(R"([0#])"))) {
            return NumberFormatType::Integer;
        }
        
        return NumberFormatType::Custom;
    }
    
    bool isDateTimeFormat(int formatId) const {
        auto format = getNumberFormat(formatId);
        if (!format.has_value()) {
            return false;
        }
        
        auto type = format->type;
        return type == NumberFormatType::Date || 
               type == NumberFormatType::Time || 
               type == NumberFormatType::DateTime;
    }
    
    bool isDateTimeFormat(const std::string& formatCode) const {
        auto type = detectNumberFormatType(formatCode);
        return type == NumberFormatType::Date || 
               type == NumberFormatType::Time || 
               type == NumberFormatType::DateTime;
    }

    bool isDateTimeStyle(int styleIndex) const {
        if (!m_isOpen || styleIndex < 0 || static_cast<size_t>(styleIndex) >= m_dateTimeStyleMask.size()) {
            return false;
        }
        return m_dateTimeStyleMask[static_cast<size_t>(styleIndex)] != 0;
    }
    
    size_t getStyleCount() const {
        return m_cellStyles.size();
    }
    
    size_t getNumberFormatCount() const {
        return m_numberFormats.size();
    }

private:
    void parseStylesXml(const ByteVector& xmlData) {
        // Initialize libxml2 reader
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            xmlData.size(),
            nullptr,
            nullptr,
            XML_PARSE_NOENT | XML_PARSE_NONET | XML_PARSE_COMPACT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for styles.xml");
        }
        
        // Parse the XML
        int result = xmlTextReaderRead(reader);
        while (result == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name) {
                    if (xmlStrcmp(name, BAD_CAST "numFmts") == 0) {
                        parseNumberFormats(reader);
                    } else if (xmlStrcmp(name, BAD_CAST "fonts") == 0) {
                        parseFonts(reader);
                    } else if (xmlStrcmp(name, BAD_CAST "fills") == 0) {
                        parseFills(reader);
                    } else if (xmlStrcmp(name, BAD_CAST "borders") == 0) {
                        parseBorders(reader);
                    } else if (xmlStrcmp(name, BAD_CAST "cellXfs") == 0) {
                        parseCellXfs(reader);
                    }
                    xmlFree(name);
                }
            }
            result = xmlTextReaderRead(reader);
        }
        
        xmlFreeTextReader(reader);
        
        if (result < 0) {
            throw XlsxError("Error parsing styles.xml");
        }
    }
    
    void parseNumberFormats(xmlTextReaderPtr reader) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "numFmt") == 0) {
                    NumberFormat format;
                    
                    xmlChar* numFmtId = xmlTextReaderGetAttribute(reader, BAD_CAST "numFmtId");
                    xmlChar* formatCode = xmlTextReaderGetAttribute(reader, BAD_CAST "formatCode");
                    
                    if (numFmtId && formatCode) {
                        format.formatId = std::atoi(reinterpret_cast<const char*>(numFmtId));
                        format.formatCode = reinterpret_cast<const char*>(formatCode);
                        format.type = detectNumberFormatType(format.formatCode);
                        format.isBuiltIn = false;
                        
                        m_numberFormats[format.formatId] = format;
                    }
                    
                    if (numFmtId) xmlFree(numFmtId);
                    if (formatCode) xmlFree(formatCode);
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    void parseFonts(xmlTextReaderPtr reader) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "font") == 0) {
                    FontInfo font;
                    parseFontElement(reader, font);
                    m_fonts.push_back(font);
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    void parseFontElement(xmlTextReaderPtr reader, FontInfo& font) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name) {
                    if (xmlStrcmp(name, BAD_CAST "name") == 0) {
                        xmlChar* val = xmlTextReaderGetAttribute(reader, BAD_CAST "val");
                        if (val) {
                            font.name = reinterpret_cast<const char*>(val);
                            xmlFree(val);
                        }
                    } else if (xmlStrcmp(name, BAD_CAST "sz") == 0) {
                        xmlChar* val = xmlTextReaderGetAttribute(reader, BAD_CAST "val");
                        if (val) {
                            font.size = std::atof(reinterpret_cast<const char*>(val));
                            xmlFree(val);
                        }
                    } else if (xmlStrcmp(name, BAD_CAST "b") == 0) {
                        font.bold = true;
                    } else if (xmlStrcmp(name, BAD_CAST "i") == 0) {
                        font.italic = true;
                    } else if (xmlStrcmp(name, BAD_CAST "u") == 0) {
                        font.underline = true;
                    } else if (xmlStrcmp(name, BAD_CAST "color") == 0) {
                        xmlChar* rgb = xmlTextReaderGetAttribute(reader, BAD_CAST "rgb");
                        if (rgb) {
                            font.color = reinterpret_cast<const char*>(rgb);
                            xmlFree(rgb);
                        }
                    }
                    xmlFree(name);
                }
            }
        }
    }
    
    void parseFills(xmlTextReaderPtr reader) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "fill") == 0) {
                    FillInfo fill;
                    parseFillElement(reader, fill);
                    m_fills.push_back(fill);
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    void parseFillElement(xmlTextReaderPtr reader, FillInfo& fill) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "patternFill") == 0) {
                    xmlChar* patternType = xmlTextReaderGetAttribute(reader, BAD_CAST "patternType");
                    if (patternType) {
                        fill.patternType = reinterpret_cast<const char*>(patternType);
                        xmlFree(patternType);
                    }
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    void parseBorders(xmlTextReaderPtr reader) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "border") == 0) {
                    BorderInfo border;
                    parseBorderElement(reader, border);
                    m_borders.push_back(border);
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    void parseBorderElement(xmlTextReaderPtr reader, BorderInfo& border) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name) {
                    xmlChar* style = xmlTextReaderGetAttribute(reader, BAD_CAST "style");
                    std::string styleStr = style ? reinterpret_cast<const char*>(style) : "none";
                    
                    if (xmlStrcmp(name, BAD_CAST "left") == 0) {
                        border.left = styleStr;
                    } else if (xmlStrcmp(name, BAD_CAST "right") == 0) {
                        border.right = styleStr;
                    } else if (xmlStrcmp(name, BAD_CAST "top") == 0) {
                        border.top = styleStr;
                    } else if (xmlStrcmp(name, BAD_CAST "bottom") == 0) {
                        border.bottom = styleStr;
                    } else if (xmlStrcmp(name, BAD_CAST "diagonal") == 0) {
                        border.diagonal = styleStr;
                    }
                    
                    if (style) xmlFree(style);
                    xmlFree(name);
                }
            }
        }
    }
    
    void parseCellXfs(xmlTextReaderPtr reader) {
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth && xmlTextReaderNodeType(reader) == XML_READER_TYPE_END_ELEMENT) {
                break;
            }
            
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "xf") == 0) {
                    CellStyle style;
                    style.styleIndex = static_cast<int>(m_cellStyles.size());
                    
                    xmlChar* numFmtId = xmlTextReaderGetAttribute(reader, BAD_CAST "numFmtId");
                    xmlChar* fontId = xmlTextReaderGetAttribute(reader, BAD_CAST "fontId");
                    xmlChar* fillId = xmlTextReaderGetAttribute(reader, BAD_CAST "fillId");
                    xmlChar* borderId = xmlTextReaderGetAttribute(reader, BAD_CAST "borderId");
                    
                    if (numFmtId) {
                        int formatId = std::atoi(reinterpret_cast<const char*>(numFmtId));
                        
                        // Look up the format directly from our parsed formats or built-in formats
                        auto it = m_numberFormats.find(formatId);
                        if (it != m_numberFormats.end()) {
                            style.numberFormat = it->second;
                        } else {
                            // Try built-in formats
                            auto builtIn = getBuiltInNumberFormat(formatId);
                            if (builtIn.has_value()) {
                                style.numberFormat = *builtIn;
                            } else {
                                // Default format
                                style.numberFormat.formatId = formatId;
                                style.numberFormat.formatCode = "General";
                                style.numberFormat.type = NumberFormatType::General;
                                style.numberFormat.isBuiltIn = true;
                            }
                        }
                        
                        xmlFree(numFmtId);
                    }
                    
                    if (fontId) {
                        int fid = std::atoi(reinterpret_cast<const char*>(fontId));
                        if (fid >= 0 && static_cast<size_t>(fid) < m_fonts.size()) {
                            style.font = m_fonts[fid];
                        }
                        xmlFree(fontId);
                    }
                    
                    if (fillId) {
                        int fid = std::atoi(reinterpret_cast<const char*>(fillId));
                        if (fid >= 0 && static_cast<size_t>(fid) < m_fills.size()) {
                            style.fill = m_fills[fid];
                        }
                        xmlFree(fillId);
                    }
                    
                    if (borderId) {
                        int bid = std::atoi(reinterpret_cast<const char*>(borderId));
                        if (bid >= 0 && static_cast<size_t>(bid) < m_borders.size()) {
                            style.border = m_borders[bid];
                        }
                        xmlFree(borderId);
                    }
                    
                    const NumberFormatType type = style.numberFormat.type;
                    const bool isDateTime = (type == NumberFormatType::Date) ||
                                            (type == NumberFormatType::Time) ||
                                            (type == NumberFormatType::DateTime);
                    m_dateTimeStyleMask.push_back(isDateTime ? 1 : 0);
                    m_cellStyles.push_back(style);
                }
                
                if (name) xmlFree(name);
            }
        }
    }
    
    std::optional<NumberFormat> getBuiltInNumberFormat(int formatId) const {
        // Excel built-in number formats
        static const std::map<int, std::pair<std::string, NumberFormatType>> builtInFormats = {
            {0, {"General", NumberFormatType::General}},
            {1, {"0", NumberFormatType::Integer}},
            {2, {"0.00", NumberFormatType::Decimal}},
            {3, {"#,##0", NumberFormatType::Integer}},
            {4, {"#,##0.00", NumberFormatType::Decimal}},
            {9, {"0%", NumberFormatType::Percentage}},
            {10, {"0.00%", NumberFormatType::Percentage}},
            {11, {"0.00E+00", NumberFormatType::Scientific}},
            {12, {"# ?/?", NumberFormatType::Fraction}},
            {13, {"# ?\x3f/?\x3f", NumberFormatType::Fraction}},
            {14, {"mm-dd-yy", NumberFormatType::Date}},
            {15, {"d-mmm-yy", NumberFormatType::Date}},
            {16, {"d-mmm", NumberFormatType::Date}},
            {17, {"mmm-yy", NumberFormatType::Date}},
            {18, {"h:mm AM/PM", NumberFormatType::Time}},
            {19, {"h:mm:ss AM/PM", NumberFormatType::Time}},
            {20, {"h:mm", NumberFormatType::Time}},
            {21, {"h:mm:ss", NumberFormatType::Time}},
            {22, {"m/d/yy h:mm", NumberFormatType::DateTime}},
            {37, {"#,##0 ;(#,##0)", NumberFormatType::Currency}},
            {38, {"#,##0 ;[Red](#,##0)", NumberFormatType::Currency}},
            {39, {"#,##0.00;(#,##0.00)", NumberFormatType::Currency}},
            {40, {"#,##0.00;[Red](#,##0.00)", NumberFormatType::Currency}},
            {49, {"@", NumberFormatType::Text}}
        };
        
        auto it = builtInFormats.find(formatId);
        if (it != builtInFormats.end()) {
            NumberFormat format;
            format.formatId = formatId;
            format.formatCode = it->second.first;
            format.type = it->second.second;
            format.isBuiltIn = true;
            return format;
        }
        
        return std::nullopt;
    }
    
    bool m_isOpen = false;
    std::map<int, NumberFormat> m_numberFormats;
    std::vector<FontInfo> m_fonts;
    std::vector<FillInfo> m_fills;
    std::vector<BorderInfo> m_borders;
    std::vector<CellStyle> m_cellStyles;
    std::vector<uint8_t> m_dateTimeStyleMask;
};

// StylesRegistry implementation
StylesRegistry::StylesRegistry() : m_impl(std::make_unique<Impl>()) {}

StylesRegistry::~StylesRegistry() = default;

StylesRegistry::StylesRegistry(StylesRegistry&&) noexcept = default;
StylesRegistry& StylesRegistry::operator=(StylesRegistry&&) noexcept = default;

void StylesRegistry::parse(const OpcPackage& package) {
    m_impl->parse(package);
}

bool StylesRegistry::isOpen() const {
    return m_impl->isOpen();
}

void StylesRegistry::close() {
    m_impl->close();
}

std::optional<CellStyle> StylesRegistry::getCellStyle(int styleIndex) const {
    return m_impl->getCellStyle(styleIndex);
}

std::optional<NumberFormat> StylesRegistry::getNumberFormat(int formatId) const {
    return m_impl->getNumberFormat(formatId);
}

NumberFormatType StylesRegistry::detectNumberFormatType(const std::string& formatCode) const {
    return m_impl->detectNumberFormatType(formatCode);
}

bool StylesRegistry::isDateTimeFormat(int formatId) const {
    return m_impl->isDateTimeFormat(formatId);
}

bool StylesRegistry::isDateTimeFormat(const std::string& formatCode) const {
    return m_impl->isDateTimeFormat(formatCode);
}

bool StylesRegistry::isDateTimeStyle(int styleIndex) const {
    return m_impl->isDateTimeStyle(styleIndex);
}

size_t StylesRegistry::getStyleCount() const {
    return m_impl->getStyleCount();
}

size_t StylesRegistry::getNumberFormatCount() const {
    return m_impl->getNumberFormatCount();
}

} // namespace xlsxcsv::core
