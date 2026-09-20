#include "xlsxcsv/core.hpp"
#include <libxml/xmlreader.h>
#include <map>
#include <algorithm>

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
    
    void parseWorkbook() {
        std::string workbookPath = m_package->findWorkbookPath();
        auto xmlData = m_package->getZipReader().readEntry(workbookPath);
        
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
