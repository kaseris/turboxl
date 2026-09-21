#include "xlsxcsv/core.hpp"
#include "core/primitive_cell_handler.hpp"
#include <libxml/xmlreader.h>
#include <libxml/xmlstring.h>
#include <stdexcept>
#include <memory>
#include <sstream>
#include <charconv>
#include <cstring>
#include <cstdlib>
#include <system_error>
#include <exception>

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
        
        // Relationship targets may be package-absolute ("/xl/...") or
        // relative to xl/workbook.xml ("worksheets/..."). ZIP entry names
        // never start with '/', so normalize absolute targets first.
        std::string fullPath = sheetPath;
        if (!fullPath.empty() && fullPath.front() == '/') {
            fullPath.erase(0, 1);
        } else if (fullPath.find("xl/") != 0) {
            fullPath = "xl/" + fullPath;
        }
        auto entry = package.getZipReader().openEntryStream(fullPath);
        StreamContext context{entry.get(), {}};
        xmlTextReaderPtr reader = xmlReaderForIO(
            &readStream, &closeStream, &context, nullptr, nullptr,
            XML_PARSE_NOENT | XML_PARSE_NOCDATA | XML_PARSE_NONET | XML_PARSE_COMPACT);
        if (!reader) {
            throw XlsxError("Failed to create streaming XML reader for worksheet");
        }
        try {
            parseWorksheetXml(reader, handler, sharedStrings, styles);
        } catch (const std::exception& e) {
            handler.handleError("Worksheet parsing error: " + std::string(e.what()));
        }
        xmlFreeTextReader(reader);
        if (context.error) {
            std::rethrow_exception(context.error);
        }
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
            nullptr, nullptr,
            XML_PARSE_NOENT | XML_PARSE_NOCDATA | XML_PARSE_NONET | XML_PARSE_COMPACT);
        
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
    struct StreamContext {
        ZipEntryStream* stream;
        std::exception_ptr error;
    };

    static int readStream(void* rawContext, char* buffer, int length) noexcept {
        auto* context = static_cast<StreamContext*>(rawContext);
        try {
            return static_cast<int>(context->stream->read(
                buffer, static_cast<size_t>(length)));
        } catch (...) {
            context->error = std::current_exception();
            return -1;
        }
    }

    static int closeStream([[maybe_unused]] void* rawContext) noexcept {
        return 0;
    }

    static bool parseIntRange(const char* begin, const char* end, int& out) {
        if (!begin || !end || begin >= end) {
            return false;
        }
        auto [ptr, ec] = std::from_chars(begin, end, out);
        return ec == std::errc{} && ptr == end;
    }

    static bool parseCellColumn(const char* ref, int& outColumn) {
        if (!ref || *ref == '\0') {
            return false;
        }

        int column = 0;
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
        outColumn = column;
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
        auto* rowControl = dynamic_cast<internal::WorksheetRowControl*>(&handler);
        bool intentionallyStopped = rowControl && !rowControl->shouldContinueParsing();
        
        int ret = intentionallyStopped ? 0 : 1;
        while (!intentionallyStopped && (ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT) {
                if (strcmp(name, "row") == 0) {
                    xmlChar* rowReference = xmlTextReaderGetAttribute(
                        reader, reinterpret_cast<const xmlChar*>("r"));
                    int rowNumber = 0;
                    const bool hasRowNumber = rowReference &&
                        parseInt(reinterpret_cast<const char*>(rowReference), rowNumber) &&
                        rowNumber > 0;
                    if (rowReference) xmlFree(rowReference);
                    if (rowControl && hasRowNumber &&
                        !rowControl->shouldParseRow(rowNumber)) {
                        intentionallyStopped = true;
                        break;
                    }
                    // Parse row element
                    parseRow(reader, handler, sharedStrings, styles);
                    if (rowControl && !rowControl->shouldContinueParsing()) {
                        intentionallyStopped = true;
                        break;
                    }
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
        
        if (ret != 0 && !intentionallyStopped) {
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
                    if (parseIntRange(attrValue, colon, first) &&
                        parseInt(colon + 1, last) &&
                        last >= first) {
                        spanReserveHint = std::min(last - first + 1, 16384);
                    }
                }
            } while (xmlTextReaderMoveToNextAttribute(reader) == 1);
            xmlTextReaderMoveToElement(reader);
        }

        auto* primitiveHandler = dynamic_cast<internal::PrimitiveCellHandler*>(&handler);
        const bool streamPrimitives = primitiveHandler != nullptr;
        auto* cellHandler = dynamic_cast<SheetCellHandler*>(&handler);
        const bool streamCells = !streamPrimitives && cellHandler && cellHandler->acceptsStreamingCells();
        std::optional<RowData> rowData;
        if (streamPrimitives) {
            primitiveHandler->beginPrimitiveRow(
                rowNumber, isHidden, static_cast<std::size_t>(spanReserveHint));
        } else if (streamCells) {
            cellHandler->beginRow(rowNumber, isHidden);
        } else {
            rowData.emplace();
            rowData->rowNumber = rowNumber;
            rowData->hidden = isHidden;
            if (spanReserveHint > 0) {
                rowData->cells.reserve(static_cast<size_t>(spanReserveHint));
            }
        }
        
        // Parse cells in this row
        if (xmlTextReaderIsEmptyElement(reader)) {
            if (streamPrimitives) primitiveHandler->endPrimitiveRow();
            else if (streamCells) cellHandler->endRow();
            else handler.handleRow(*rowData);
            return;
        }
        
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            int nodeType = xmlTextReaderNodeType(reader);
            
            if (!name) continue;
            
            if (nodeType == XML_READER_TYPE_ELEMENT && strcmp(name, "c") == 0) {
                // Parse cell
                if (streamPrimitives) {
                    parsePrimitiveCell(reader, *primitiveHandler);
                } else {
                    auto cell = parseCell(reader, rowNumber, sharedStrings, styles);
                    if (cell.has_value()) {
                        if (streamCells) cellHandler->handleCell(std::move(*cell));
                        else rowData->cells.push_back(std::move(*cell));
                    }
                }
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && strcmp(name, "row") == 0) {
                // End of row
                break;
            }
        }
        
        if (streamPrimitives) primitiveHandler->endPrimitiveRow();
        else if (streamCells) cellHandler->endRow();
        else handler.handleRow(*rowData);
    }

    void parsePrimitiveCell(
        xmlTextReaderPtr reader, internal::PrimitiveCellHandler& handler) {
        int column = 0;
        CellType type = CellType::Number;

        if (xmlTextReaderMoveToFirstAttribute(reader) == 1) {
            do {
                const char* attrName = reinterpret_cast<const char*>(
                    xmlTextReaderConstName(reader));
                const char* attrValue = reinterpret_cast<const char*>(
                    xmlTextReaderConstValue(reader));
                if (!attrName || !attrValue) continue;

                if (attrName[0] == 'r' && attrName[1] == '\0') {
                    parseCellColumn(attrValue, column);
                } else if (attrName[0] == 't' && attrName[1] == '\0') {
                    if (attrValue[0] == 'b' && attrValue[1] == '\0') {
                        type = CellType::Boolean;
                    } else if (attrValue[0] == 'e' && attrValue[1] == '\0') {
                        type = CellType::Error;
                    } else if (attrValue[0] == 'n' && attrValue[1] == '\0') {
                        type = CellType::Number;
                    } else if (attrValue[0] == 's' && attrValue[1] == '\0') {
                        type = CellType::SharedString;
                    } else if (std::strcmp(attrValue, "str") == 0) {
                        type = CellType::String;
                    } else if (std::strcmp(attrValue, "inlineStr") == 0) {
                        type = CellType::InlineString;
                    } else {
                        type = CellType::Unknown;
                    }
                }
            } while (xmlTextReaderMoveToNextAttribute(reader) == 1);
            xmlTextReaderMoveToElement(reader);
        }

        if (xmlTextReaderIsEmptyElement(reader)) {
            handler.addEmpty(column);
            return;
        }

        bool emitted = false;
        int ret;
        while ((ret = xmlTextReaderRead(reader)) == 1) {
            const char* name = reinterpret_cast<const char*>(xmlTextReaderConstName(reader));
            const int nodeType = xmlTextReaderNodeType(reader);
            if (!name) continue;

            if (nodeType == XML_READER_TYPE_ELEMENT && std::strcmp(name, "v") == 0) {
                std::string value = readElementText(reader);
                emitPrimitiveValue(handler, column, type, std::move(value));
                emitted = true;
            } else if (nodeType == XML_READER_TYPE_ELEMENT && std::strcmp(name, "is") == 0) {
                handler.addString(column, parseInlineString(reader));
                emitted = true;
            } else if (nodeType == XML_READER_TYPE_END_ELEMENT && std::strcmp(name, "c") == 0) {
                break;
            }
        }
        if (!emitted) handler.addEmpty(column);
    }

    static void emitPrimitiveValue(
        internal::PrimitiveCellHandler& handler,
        int column,
        CellType type,
        std::string&& value) {
        if (value.empty()) {
            handler.addEmpty(column);
            return;
        }
        switch (type) {
            case CellType::Boolean:
                handler.addBoolean(column, value == "1");
                return;
            case CellType::Number: {
                const char* begin = value.data();
                char* parsedEnd = nullptr;
                const double number = std::strtod(begin, &parsedEnd);
                if (parsedEnd == begin + value.size()) {
                    handler.addNumber(column, number);
                } else {
                    handler.addEmpty(column);
                }
                return;
            }
            case CellType::SharedString: {
                std::size_t index = 0;
                const char* begin = value.data();
                const char* end = begin + value.size();
                const auto [parsedEnd, error] = std::from_chars(begin, end, index);
                if (error == std::errc{} && parsedEnd == end) {
                    handler.addSharedString(column, index);
                } else {
                    handler.addEmpty(column);
                }
                return;
            }
            case CellType::Error:
            case CellType::String:
            case CellType::InlineString:
            case CellType::Unknown:
                handler.addString(column, std::move(value));
                return;
        }
    }
    
    std::optional<CellData> parseCell(xmlTextReaderPtr reader,
                                      int rowNumber,
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
                    int column = 0;
                    if (parseCellColumn(attrValue, column)) {
                        cell.coordinate.row = rowNumber;
                        cell.coordinate.column = column;
                    }
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
