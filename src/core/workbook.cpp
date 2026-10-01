#include "xlsxcsv/core.hpp"
#include "core/fast_xml.hpp"
#include <libxml/xmlreader.h>
#include <map>
#include <algorithm>
#include <charconv>
#include <string_view>

namespace xlsxcsv::core {

class Workbook::Impl {
public:
    Impl() = default;
    
    ~Impl() {
        close();
    }
    
    void open(const OpcPackage& package) {
        if (m_isOpen) {
            close();
        }
        
        // Store reference to the package (must remain valid while workbook is open)
        m_package = &package;
        
        // Parse workbook.xml to get sheets and properties
        parseWorkbook();
        
        // Parse workbook relationships to map r:id to targets
        parseWorkbookRelationships();
        
        // Update sheet targets using relationship mapping
        updateSheetTargets();
        
        m_isOpen = true;
    }
    
    void close() {
        m_package = nullptr;
        m_isOpen = false;
        m_sheets.clear();
        m_relationships.clear();
        m_properties = WorkbookProperties{};
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::vector<SheetInfo> getSheets() const {
        if (!m_isOpen) {
            throw XlsxError("Workbook is not open");
        }
        return m_sheets;
    }
    
    std::optional<SheetInfo> findSheet(const std::string& name) const {
        if (!m_isOpen) {
            throw XlsxError("Workbook is not open");
        }
        
        auto it = std::find_if(m_sheets.begin(), m_sheets.end(),
                              [&name](const SheetInfo& sheet) {
                                  return sheet.name == name;
                              });
        
        if (it == m_sheets.end()) return std::nullopt;
        if (it->kind != SheetKind::Worksheet) {
            throw XlsxError("Sheet is not a worksheet: " + name);
        }
        return *it;
        
        return std::nullopt;
    }
    
    std::optional<SheetInfo> findSheet(int index) const {
        if (!m_isOpen) {
            throw XlsxError("Workbook is not open");
        }
        
        if (index < 0) return std::nullopt;
        int worksheetIndex = 0;
        for (const auto& sheet : m_sheets) {
            if (sheet.kind != SheetKind::Worksheet) continue;
            if (worksheetIndex == index) return sheet;
            ++worksheetIndex;
        }
        return std::nullopt;
    }
    
    size_t getSheetCount() const {
        if (!m_isOpen) {
            return 0;
        }
        return static_cast<size_t>(std::count_if(
            m_sheets.begin(), m_sheets.end(), [](const SheetInfo& sheet) {
                return sheet.kind == SheetKind::Worksheet;
            }));
    }
    
    const WorkbookProperties& getProperties() const {
        if (!m_isOpen) {
            throw XlsxError("Workbook is not open");
        }
        return m_properties;
    }
    
    DateSystem getDateSystem() const {
        return getProperties().dateSystem;
    }
    
    std::string resolveRelationshipTarget(const std::string& relationshipId) const {
        if (!m_isOpen) {
            throw XlsxError("Workbook is not open");
        }
        
        auto it = m_relationships.find(relationshipId);
        if (it != m_relationships.end()) {
            return it->second.target;
        }
        
        throw XlsxError("Relationship not found: " + relationshipId);
    }

private:
    struct Relationship {
        std::string id;
        std::string type;
        std::string target;
    };
    
    // Reads workbook properties and <sheet> entries without libxml2. Returns
    // false when the document needs the full parser; the caller discards
    // partial results.
    bool parseWorkbookFast(const ByteVector& xmlData) {
        static constexpr std::string_view relationshipsNamespace =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        bool sawRoot = false;
        bool rootDeclaresRelationships = false;
        std::string name, sheetId, relationshipId, state, date1904;
        return fastxml::forEachStartTag(
            std::string_view(reinterpret_cast<const char*>(xmlData.data()), xmlData.size()),
            [&](std::string_view element, std::string_view tag) {
                if (!sawRoot) {
                    sawRoot = true;
                    const std::string prefix = "xmlns:r=";
                    const std::string uri(relationshipsNamespace);
                    rootDeclaresRelationships =
                        tag.find(prefix + "\"" + uri + "\"") != std::string_view::npos ||
                        tag.find(prefix + "'" + uri + "'") != std::string_view::npos;
                }
                if (element == "workbookPr") {
                    bool present = false;
                    if (!fastxml::readAttribute(tag, "date1904", date1904, present)) return false;
                    m_properties.dateSystem = present && (date1904 == "1" || date1904 == "true")
                        ? DateSystem::Date1904 : DateSystem::Date1900;
                    return true;
                }
                if (element != "sheet") return true;

                bool hasName = false, hasId = false, hasRelationship = false, hasState = false;
                if (!fastxml::readAttribute(tag, "name", name, hasName) ||
                    !fastxml::readAttribute(tag, "sheetId", sheetId, hasId) ||
                    !fastxml::readAttribute(tag, "r:id", relationshipId, hasRelationship) ||
                    !fastxml::readAttribute(tag, "state", state, hasState)) {
                    return false;
                }
                // libxml2 resolves "r:id" through the in-scope prefix binding,
                // so only trust the textual match when the conventional
                // declaration is present on the root element.
                if (hasRelationship && !rootDeclaresRelationships) return false;

                SheetInfo sheet;
                if (hasName) sheet.name = name;
                if (hasId) {
                    int value = 0;
                    const auto [end, error] = std::from_chars(
                        sheetId.data(), sheetId.data() + sheetId.size(), value);
                    if (error != std::errc{} || end != sheetId.data() + sheetId.size()) return false;
                    sheet.sheetId = value;
                }
                if (hasRelationship) sheet.relationshipId = relationshipId;
                sheet.visibility = SheetVisibility::Visible;
                if (hasState) {
                    if (state == "hidden") sheet.visibility = SheetVisibility::Hidden;
                    if (state == "veryHidden") sheet.visibility = SheetVisibility::VeryHidden;
                }
                sheet.visible = sheet.visibility == SheetVisibility::Visible;
                m_sheets.push_back(std::move(sheet));
                return true;
            });
    }

    void parseWorkbook() {
        std::string workbookPath = m_package->findWorkbookPath();
        auto xmlData = m_package->getZipReader().readEntry(workbookPath);
        if (parseWorkbookFast(xmlData)) return;
        m_sheets.clear();
        m_properties = WorkbookProperties{};
        
        // Initialize libxml2 reader
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            xmlData.size(),
            nullptr,
            nullptr,
            XML_PARSE_NOENT | XML_PARSE_NONET | XML_PARSE_COMPACT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for workbook.xml");
        }
        
        // Parse the XML
        int result = xmlTextReaderRead(reader);
        while (result == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name) {
                    if (xmlStrcmp(name, BAD_CAST "workbookPr") == 0) {
                        parseWorkbookProperties(reader);
                    } else if (xmlStrcmp(name, BAD_CAST "sheet") == 0) {
                        parseSheetElement(reader);
                    }
                    xmlFree(name);
                }
            }
            result = xmlTextReaderRead(reader);
        }
        
        xmlFreeTextReader(reader);
        
        if (result < 0) {
            throw XlsxError("Error parsing workbook.xml");
        }
    }
    
    void parseWorkbookProperties(xmlTextReaderPtr reader) {
        xmlChar* date1904Attr = xmlTextReaderGetAttribute(reader, BAD_CAST "date1904");
        
        if (date1904Attr) {
            std::string date1904Value = reinterpret_cast<const char*>(date1904Attr);
            if (date1904Value == "1" || date1904Value == "true") {
                m_properties.dateSystem = DateSystem::Date1904;
            } else {
                m_properties.dateSystem = DateSystem::Date1900;
            }
            xmlFree(date1904Attr);
        } else {
            // Default to 1900 date system if not specified
            m_properties.dateSystem = DateSystem::Date1900;
        }
    }
    
    void parseSheetElement(xmlTextReaderPtr reader) {
        SheetInfo sheet;
        
        // Get sheet attributes
        xmlChar* name = xmlTextReaderGetAttribute(reader, BAD_CAST "name");
        xmlChar* sheetId = xmlTextReaderGetAttribute(reader, BAD_CAST "sheetId");
        xmlChar* rId = xmlTextReaderGetAttribute(reader, BAD_CAST "r:id");
        xmlChar* state = xmlTextReaderGetAttribute(reader, BAD_CAST "state");
        
        if (name) {
            sheet.name = reinterpret_cast<const char*>(name);
            xmlFree(name);
        }
        
        if (sheetId) {
            sheet.sheetId = std::atoi(reinterpret_cast<const char*>(sheetId));
            xmlFree(sheetId);
        }
        
        if (rId) {
            sheet.relationshipId = reinterpret_cast<const char*>(rId);
            xmlFree(rId);
        }
        
        // Check visibility state
        sheet.visibility = SheetVisibility::Visible;
        if (state) {
            std::string stateValue = reinterpret_cast<const char*>(state);
            if (stateValue == "hidden") sheet.visibility = SheetVisibility::Hidden;
            if (stateValue == "veryHidden") {
                sheet.visibility = SheetVisibility::VeryHidden;
            }
            xmlFree(state);
        }
        sheet.visible = sheet.visibility == SheetVisibility::Visible;
        
        m_sheets.push_back(sheet);
    }
    
    void parseWorkbookRelationships() {
        const std::string relsPath = "xl/_rels/workbook.xml.rels";
        
        if (!m_package->getZipReader().hasEntry(relsPath)) {
            throw XlsxError("Missing workbook relationships file: " + relsPath);
        }
        
        auto xmlData = m_package->getZipReader().readEntry(relsPath);

        const bool scanned = fastxml::parseRelationshipsFast(
            std::string_view(reinterpret_cast<const char*>(xmlData.data()), xmlData.size()),
            [&](const std::string& id, const std::string& type, const std::string& target) {
                m_relationships[id] = Relationship{id, type, target};
            });
        if (scanned) return;
        m_relationships.clear();

        // Initialize libxml2 reader
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            xmlData.size(),
            nullptr,
            nullptr,
            XML_PARSE_NOENT | XML_PARSE_NONET | XML_PARSE_COMPACT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for workbook relationships");
        }
        
        // Parse the XML
        int result = xmlTextReaderRead(reader);
        while (result == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && xmlStrcmp(name, BAD_CAST "Relationship") == 0) {
                    xmlChar* id = xmlTextReaderGetAttribute(reader, BAD_CAST "Id");
                    xmlChar* type = xmlTextReaderGetAttribute(reader, BAD_CAST "Type");
                    xmlChar* target = xmlTextReaderGetAttribute(reader, BAD_CAST "Target");
                    
                    if (id && type && target) {
                        Relationship rel;
                        rel.id = reinterpret_cast<const char*>(id);
                        rel.type = reinterpret_cast<const char*>(type);
                        rel.target = reinterpret_cast<const char*>(target);
                        
                        m_relationships[rel.id] = rel;
                    }
                    
                    if (id) xmlFree(id);
                    if (type) xmlFree(type);
                    if (target) xmlFree(target);
                }
                
                if (name) xmlFree(name);
            }
            result = xmlTextReaderRead(reader);
        }
        
        xmlFreeTextReader(reader);
        
        if (result < 0) {
            throw XlsxError("Error parsing workbook relationships");
        }
    }
    
    void updateSheetTargets() {
        for (auto& sheet : m_sheets) {
            auto it = m_relationships.find(sheet.relationshipId);
            if (it != m_relationships.end()) {
                sheet.target = it->second.target;
                const auto separator = it->second.type.find_last_of('/');
                const auto type = it->second.type.substr(separator + 1);
                if (type == "worksheet") {
                    sheet.kind = SheetKind::Worksheet;
                } else if (type == "chartsheet") {
                    sheet.kind = SheetKind::Chartsheet;
                } else {
                    sheet.kind = SheetKind::Other;
                }
            } else {
                throw XlsxError("Relationship not found for sheet: " + sheet.name + " (r:id=" + sheet.relationshipId + ")");
            }
        }
    }
    
    const OpcPackage* m_package = nullptr;
    bool m_isOpen = false;
    std::vector<SheetInfo> m_sheets;
    std::map<std::string, Relationship> m_relationships;
    WorkbookProperties m_properties;
};

// Workbook implementation
Workbook::Workbook() : m_impl(std::make_unique<Impl>()) {}

Workbook::~Workbook() = default;

Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;

void Workbook::open(const OpcPackage& package) {
    m_impl->open(package);
}

void Workbook::close() {
    m_impl->close();
}

bool Workbook::isOpen() const {
    return m_impl->isOpen();
}

std::vector<SheetInfo> Workbook::getSheets() const {
    return m_impl->getSheets();
}

std::optional<SheetInfo> Workbook::findSheet(const std::string& name) const {
    return m_impl->findSheet(name);
}

std::optional<SheetInfo> Workbook::findSheet(int index) const {
    return m_impl->findSheet(index);
}

size_t Workbook::getSheetCount() const {
    return m_impl->getSheetCount();
}

const WorkbookProperties& Workbook::getProperties() const {
    return m_impl->getProperties();
}

DateSystem Workbook::getDateSystem() const {
    return m_impl->getDateSystem();
}

std::string Workbook::resolveRelationshipTarget(const std::string& relationshipId) const {
    return m_impl->resolveRelationshipTarget(relationshipId);
}

} // namespace xlsxcsv::core
