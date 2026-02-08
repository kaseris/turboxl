#include "xlsxcsv/core.hpp"
#include <libxml/xmlreader.h>
#include <libxml/xmlstring.h>
#include <stdexcept>
#include <memory>
#include <sstream>
#include <charconv>
#include <cstring>
#include <cstdlib>
#include <system_error>

namespace xlsxcsv::core {

class SheetStreamReader::Impl {
public:
    Impl() = default;
    ~Impl() = default;
    
    void parseSheet(const OpcPackage& package, 
                   const std::string& sheetPath,
                   SheetRowHandler& handler,
                   const SharedStringsProvider* sharedStrings,
                   const StylesRegistry* styles) {
        
        // Read worksheet XML from package
        // The sheetPath is relative to xl/ directory, so prefix it
        std::string fullPath = sheetPath;
        if (fullPath.find("xl/") != 0) {
            fullPath = "xl/" + fullPath;
        }
        auto xmlData = package.getZipReader().readEntry(fullPath);
        parseSheetData(xmlData, handler, sharedStrings, styles);
    }
    
    void parseSheetData(const std::vector<uint8_t>& xmlData,
                       SheetRowHandler& handler,
                       const SharedStringsProvider* sharedStrings,
                       const StylesRegistry* styles) {
        
        if (xmlData.empty()) {
            handler.handleError("Empty worksheet data");
            return;
        }
        
        // Create XML reader from memory
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            static_cast<int>(xmlData.size()),
            nullptr, nullptr, XML_PARSE_NOENT | XML_PARSE_NOCDATA);
        
        if (!reader) {
            handler.handleError("Failed to create XML reader for worksheet");
            return;
        }
        
        // Parse worksheet
        try {
            parseWorksheetXml(reader, handler, sharedStrings, styles);
        } catch (const std::exception& e) {
            handler.handleError("Worksheet parsing error: " + std::string(e.what()));
        }
        
        xmlFreeTextReader(reader);
    }

private:
    static bool parseCellReference(const char* ref, CellCoordinate& out) {
        if (!ref || *ref == '\0') {
            return false;
        }

        int column = 0;
        int row = 0;
        const char* p = ref;

        while (*p >= 'A' && *p <= 'Z') {
            column = (column * 26) + (*p - 'A' + 1);
            ++p;
        }
        if (column == 0) {
            return false;
        }

        if (*p < '1' || *p > '9') {
            return false;
        }
        while (*p >= '0' && *p <= '9') {
            row = (row * 10) + (*p - '0');
            ++p;
        }

        if (*p != '\0' || row <= 0) {
            return false;
        }

        out.row = row;
        out.column = column;
        return true;
    }

    static bool parseInt(const char* s, int& out) {
        if (!s || *s == '\0') {
            return false;
        }
        const char* begin = s;
        const char* end = begin + std::strlen(s);
        auto [ptr, ec] = std::from_chars(begin, end, out);
        return ec == std::errc{} && ptr == end;
    }

    void parseWorksheetXml(xmlTextReaderPtr reader,
                          SheetRowHandler& handler,
                          const SharedStringsProvider* sharedStrings,
                          const StylesRegistry* styles) {
        
        WorksheetMetadata metadata;
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT) {
                if (strcmp(name, "row") == 0) {
                    // Parse row element
                    parseRow(reader, handler, sharedStrings, styles);
                } else if (strcmp(name, "mergeCells") == 0) {
                    // Parse merged cells section
                    parseMergedCells(reader, metadata);
                    // Send updated metadata immediately after parsing merged cells
                    handler.handleWorksheetMetadata(metadata);
                } else if (strcmp(name, "cols") == 0) {
                    // Parse column definitions
                    parseColumns(reader, metadata);
                    // Send updated metadata immediately after parsing cols
                    handler.handleWorksheetMetadata(metadata);
                }
            }
        }
        
        // Send metadata to handler before processing is complete
        handler.handleWorksheetMetadata(metadata);
        
        if (ret != 0) {
            throw std::runtime_error("XML parsing error in worksheet");
        }
    }
    
    void parseRow(xmlTextReaderPtr reader,
                  SheetRowHandler& handler,
                  const SharedStringsProvider* sharedStrings,
                  const StylesRegistry* styles) {
        
        int rowNumber = 1; // Default to row 1
        bool isHidden = false;
        int spanReserveHint = 0;

        if (xmlTextReaderMoveToFirstAttribute(reader) == 1) {
            do {
                const char* attrName = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
                const char* attrValue = reinterpret_cast<const char*>(xmlTextReaderConstValue(reader));
                if (!attrName || !attrValue) {
                    continue;
                }

                if (attrName[0] == 'r' && attrName[1] == '\0') {
                    int parsedRow = 0;
                    if (parseInt(attrValue, parsedRow) && parsedRow > 0) {
                        rowNumber = parsedRow;
                    }
                    continue;
                }

                if (attrName[0] == 'h' && std::strcmp(attrName, "hidden") == 0) {
                    isHidden = (attrValue[0] == '1' && attrValue[1] == '\0') || std::strcmp(attrValue, "true") == 0;
                    continue;
                }

                if (attrName[0] == 's' && std::strcmp(attrName, "spans") == 0) {
                    const char* colon = std::strchr(attrValue, ':');
                    if (!colon) {
                        continue;
                    }
                    int first = 0;
                    int last = 0;
                    std::string firstPart(attrValue, static_cast<size_t>(colon - attrValue));
                    if (parseInt(firstPart.c_str(), first) && parseInt(colon + 1, last) && last >= first) {
                        spanReserveHint = std::min(last - first + 1, 16384);
                    }
                }
            } while (xmlTextReaderMoveToNextAttribute(reader) == 1);
            xmlTextReaderMoveToElement(reader);
        }

        RowData rowData;
        rowData.rowNumber = rowNumber;
        rowData.hidden = isHidden;
        if (spanReserveHint > 0) {
            rowData.cells.reserve(static_cast<size_t>(spanReserveHint));
        }
        
        // Parse cells in this row
        if (xmlTextReaderIsEmptyElement(reader)) {
            // Empty row
            handler.handleRow(rowData);
            return;
        }
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT && strcmp(name, "c") == 0) {
                // Parse cell
                auto cell = parseCell(reader, sharedStrings, styles);
                if (cell.has_value()) {
                    rowData.cells.push_back(std::move(*cell));
                }
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && strcmp(name, "row") == 0) {
                // End of row
                break;
            }
        }
        
        handler.handleRow(rowData);
    }
    
    std::optional<CellData> parseCell(xmlTextReaderPtr reader,
                                      const SharedStringsProvider* sharedStrings,
                                      [[maybe_unused]] const StylesRegistry* styles) {
        
        CellData cell;
        bool hasTypeAttribute = false;
        
        if (xmlTextReaderMoveToFirstAttribute(reader) == 1) {
            do {
                const char* attrName = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
                const char* attrValue = reinterpret_cast<const char*>(xmlTextReaderConstValue(reader));
                if (!attrName || !attrValue) {
                    continue;
                }

                if (attrName[0] == 'r' && attrName[1] == '\0') {
                    parseCellReference(attrValue, cell.coordinate);
                    continue;
                }

                if (attrName[0] == 't' && attrName[1] == '\0') {
                    hasTypeAttribute = true;
                    if (attrValue[0] == 'b' && attrValue[1] == '\0') {
                        cell.type = CellType::Boolean;
                    } else if (attrValue[0] == 'e' && attrValue[1] == '\0') {
                        cell.type = CellType::Error;
                    } else if (attrValue[0] == 'n' && attrValue[1] == '\0') {
                        cell.type = CellType::Number;
                    } else if (attrValue[0] == 's' && attrValue[1] == '\0') {
                        cell.type = CellType::SharedString;
                    } else if (std::strcmp(attrValue, "str") == 0) {
                        cell.type = CellType::String;
                    } else if (std::strcmp(attrValue, "inlineStr") == 0) {
                        cell.type = CellType::InlineString;
                    } else {
                        cell.type = CellType::Unknown;
                    }
                    continue;
                }

                if (attrName[0] == 's' && attrName[1] == '\0') {
                    int parsedStyle = 0;
                    if (parseInt(attrValue, parsedStyle) && parsedStyle >= 0) {
                        cell.styleIndex = parsedStyle;
                    }
                    continue;
                }
            } while (xmlTextReaderMoveToNextAttribute(reader) == 1);
            xmlTextReaderMoveToElement(reader);
        }

        if (!hasTypeAttribute) {
            // No type attribute; default Excel type is numeric.
            cell.type = CellType::Number;
        }
        
        // Parse cell content
        if (xmlTextReaderIsEmptyElement(reader)) {
            // Empty cell
            cell.value = std::monostate{};
            return cell;
        }
        
        // Read cell content (v or is elements)
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT) {
                if (strcmp(name, "v") == 0) {
                    // Cell value
                    std::string valueStr = readElementText(reader);
                    cell.value = convertCellValue(valueStr, cell.type, sharedStrings);
                } else if (strcmp(name, "is") == 0) {
                    // Inline string
                    cell.value = parseInlineString(reader);
                    cell.type = CellType::InlineString;
                }
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && strcmp(name, "c") == 0) {
                // End of cell
                break;
            }
        }
        
        return cell;
    }
    
    CellValue convertCellValue(const std::string& valueStr, 
                              CellType type,
                              [[maybe_unused]] const SharedStringsProvider* sharedStrings) {
        
        if (valueStr.empty()) {
            return std::monostate{};
        }
        
        switch (type) {
            case CellType::Boolean: {
                // Excel booleans: "0" = false, "1" = true
                return valueStr == "1";
            }
            
            case CellType::Number: {
                const char* begin = valueStr.data();
                char* parseEnd = nullptr;
                double parsed = std::strtod(begin, &parseEnd);
                if (parseEnd == begin + valueStr.size()) {
                    return parsed;
                }
                return std::monostate{};
            }
            
            case CellType::SharedString: {
                int index = 0;
                const char* begin = valueStr.data();
                const char* end = begin + valueStr.size();
                auto [ptr, ec] = std::from_chars(begin, end, index);
                if (ec == std::errc{} && ptr == end) {
                    // Keep shared-string index and defer lookup to CSV conversion.
                    return index;
                }
                return std::monostate{};
            }
            
            case CellType::Error:
            case CellType::String:
            case CellType::InlineString:
            default:
                return valueStr;
        }
    }
    
    std::string readElementText(xmlTextReaderPtr reader) {
        std::string result;
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (nodeType == XML_READER_TYPE_TEXT || nodeType == XML_READER_TYPE_CDATA) {
                const char* text = reinterpret_cast<const char*>(xmlTextReaderConstValue(reader));
                if (text) {
                    result += text;
                }
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT) {
                // End of element
                break;
            }
        }
        
        return result;
    }
    
    std::string parseInlineString(xmlTextReaderPtr reader) {
        // For now, just extract text content
        // TODO: Handle rich text formatting if needed
        return readElementText(reader);
    }
    
    void parseMergedCells(xmlTextReaderPtr reader, WorksheetMetadata& metadata) {
        // Parse <mergeCells> section
        if (xmlTextReaderIsEmptyElement(reader)) {
            return; // No merged cells
        }
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT && strcmp(name, "mergeCell") == 0) {
                // Parse individual merged cell range
                xmlChar* refAttr = xmlTextReaderGetAttribute(reader, BAD_CAST "ref");
                if (refAttr) {
                    std::string refStr = reinterpret_cast<const char*>(refAttr);
                    auto range = MergedCellRange::fromReference(refStr);
                    if (range.has_value()) {
                        metadata.mergedCells.push_back(range.value());
                    }
                    xmlFree(refAttr);
                }
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && strcmp(name, "mergeCells") == 0) {
                // End of mergeCells section
                break;
            }
        }
    }
    
    void parseColumns(xmlTextReaderPtr reader, WorksheetMetadata& metadata) {
        // Parse <cols> section for column information
        if (xmlTextReaderIsEmptyElement(reader)) {
            return; // No column definitions
        }
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT && strcmp(name, "col") == 0) {
                // Parse individual column definition
                ColumnInfo colInfo;
                
                // Get column range (min and max)
                xmlChar* minAttr = xmlTextReaderGetAttribute(reader, BAD_CAST "min");
                xmlChar* maxAttr = xmlTextReaderGetAttribute(reader, BAD_CAST "max");
                
                int minCol = 1, maxCol = 1;
                if (minAttr) {
                    minCol = std::atoi(reinterpret_cast<const char*>(minAttr));
                    xmlFree(minAttr);
                }
                if (maxAttr) {
                    maxCol = std::atoi(reinterpret_cast<const char*>(maxAttr));
                    xmlFree(maxAttr);
                }
                
                // Check if hidden
                bool isHidden = false;
                xmlChar* hiddenAttr = xmlTextReaderGetAttribute(reader, BAD_CAST "hidden");
                if (hiddenAttr) {
                    std::string hiddenStr = reinterpret_cast<const char*>(hiddenAttr);
                    isHidden = (hiddenStr == "1" || hiddenStr == "true");
                    xmlFree(hiddenAttr);
                }
                
                // Get width if available
                double width = 0.0;
                xmlChar* widthAttr = xmlTextReaderGetAttribute(reader, BAD_CAST "width");
                if (widthAttr) {
                    width = std::stod(reinterpret_cast<const char*>(widthAttr));
                    xmlFree(widthAttr);
                }
                
                // Add column info for all columns in the range
                for (int col = minCol; col <= maxCol; ++col) {
                    colInfo.columnIndex = col;
                    colInfo.hidden = isHidden;
                    colInfo.width = width;
                    metadata.columnInfo.push_back(colInfo);
                }
                
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && strcmp(name, "cols") == 0) {
                // End of cols section
                break;
            }
        }
    }
};

// SheetStreamReader implementation

SheetStreamReader::SheetStreamReader() : m_impl(std::make_unique<Impl>()) {
}

SheetStreamReader::~SheetStreamReader() = default;

SheetStreamReader::SheetStreamReader(SheetStreamReader&&) noexcept = default;
SheetStreamReader& SheetStreamReader::operator=(SheetStreamReader&&) noexcept = default;

void SheetStreamReader::parseSheet(const OpcPackage& package, 
                                  const std::string& sheetPath,
                                  SheetRowHandler& handler,
                                  const SharedStringsProvider* sharedStrings,
                                  const StylesRegistry* styles) {
    m_impl->parseSheet(package, sheetPath, handler, sharedStrings, styles);
}

void SheetStreamReader::parseSheetData(const std::vector<uint8_t>& xmlData,
                                      SheetRowHandler& handler,
                                      const SharedStringsProvider* sharedStrings,
                                      const StylesRegistry* styles) {
    m_impl->parseSheetData(xmlData, handler, sharedStrings, styles);
}

} // namespace xlsxcsv::core
