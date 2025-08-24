#include "xlsxcsv/core.hpp"
#include <libxml/xmlreader.h>
#include <memory>
#include <map>
#include <regex>

namespace xlsxcsv::core {

class OpcPackage::Impl {
public:
    Impl() = default;
    
    ~Impl() {
        close();
    }
    
    void open(const std::string& path) {
        if (m_zipReader.isOpen()) {
            close();
        }
        
        // Open the ZIP file using our secure ZipReader
        m_zipReader.open(path);
        
        // Parse the OPC package structure
        parseContentTypes();
        parseMainRelationships();
        
        m_isOpen = true;
    }
    
    void close() {
        m_zipReader.close();
        m_isOpen = false;
        m_contentTypes.clear();
        m_relationships.clear();
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::string findWorkbookPath() {
        if (!m_isOpen) {
            throw XlsxError("OPC package is not open");
        }
        
        // Look for the main document relationship
        for (const auto& rel : m_relationships) {
            if (rel.second.type.find("officeDocument") != std::string::npos) {
                return rel.second.target;
            }
        }
        
        throw XlsxError("Workbook not found in OPC package relationships");
    }
    
    std::vector<std::string> getContentTypes() {
        if (!m_isOpen) {
            throw XlsxError("OPC package is not open");
        }
        
        std::vector<std::string> types;
        for (const auto& entry : m_contentTypes) {
            types.push_back(entry.second);
        }
        return types;
    }
    
    const ZipReader& getZipReader() const {
        if (!m_isOpen) {
            throw XlsxError("OPC package is not open");
        }
        return m_zipReader;
    }

private:
    struct Relationship {
        std::string id;
        std::string type;
        std::string target;
    };
    
    void parseContentTypes() {
        const std::string contentTypesPath = "[Content_Types].xml";
        
        if (!m_zipReader.hasEntry(contentTypesPath)) {
            throw XlsxError("Missing [Content_Types].xml in OPC package");
        }
        
        auto xmlData = m_zipReader.readEntry(contentTypesPath);
        parseXmlForContentTypes(xmlData);
    }
    
    void parseMainRelationships() {
        const std::string relsPath = "_rels/.rels";
        
        if (!m_zipReader.hasEntry(relsPath)) {
            throw XlsxError("Missing _rels/.rels in OPC package");
        }
        
        auto xmlData = m_zipReader.readEntry(relsPath);
        parseXmlForRelationships(xmlData);
    }
    
    void parseXmlForContentTypes(const ByteVector& xmlData) {
        // Initialize libxml2 reader
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            xmlData.size(),
            nullptr,
            nullptr,
            XML_PARSE_NOENT | XML_PARSE_NONET | XML_PARSE_COMPACT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for [Content_Types].xml");
        }
        
        // Parse the XML
        int result = xmlTextReaderRead(reader);
        while (result == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                xmlChar* name = xmlTextReaderName(reader);
                
                if (name && (xmlStrcmp(name, BAD_CAST "Override") == 0 ||
                            xmlStrcmp(name, BAD_CAST "Default") == 0)) {
                    
                    xmlChar* partName = xmlTextReaderGetAttribute(reader, BAD_CAST "PartName");
                    xmlChar* extension = xmlTextReaderGetAttribute(reader, BAD_CAST "Extension");
                    xmlChar* contentType = xmlTextReaderGetAttribute(reader, BAD_CAST "ContentType");
                    
                    if (contentType) {
                        std::string key;
                        if (partName) {
                            key = reinterpret_cast<const char*>(partName);
                            // Remove leading slash if present
                            if (!key.empty() && key[0] == '/') {
                                key.erase(0, 1);
                            }
                        } else if (extension) {
                            key = std::string("*.") + reinterpret_cast<const char*>(extension);
                        }
                        
                        if (!key.empty()) {
                            m_contentTypes[key] = reinterpret_cast<const char*>(contentType);
                        }
                    }
                    
                    if (partName) xmlFree(partName);
                    if (extension) xmlFree(extension);
                    if (contentType) xmlFree(contentType);
                }
                
                if (name) xmlFree(name);
            }
            result = xmlTextReaderRead(reader);
        }
        
        xmlFreeTextReader(reader);
        
        if (result < 0) {
            throw XlsxError("Error parsing [Content_Types].xml");
        }
    }
    
    void parseXmlForRelationships(const ByteVector& xmlData) {
        // Initialize libxml2 reader
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            xmlData.size(),
            nullptr,
            nullptr,
            XML_PARSE_NOENT | XML_PARSE_NONET | XML_PARSE_COMPACT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for _rels/.rels");
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
            throw XlsxError("Error parsing _rels/.rels");
        }
        
        if (m_relationships.empty()) {
            throw XlsxError("No relationships found in _rels/.rels");
        }
    }
    
    ZipReader m_zipReader;
    bool m_isOpen = false;
    std::map<std::string, std::string> m_contentTypes;
    std::map<std::string, Relationship> m_relationships;
};

// OpcPackage implementation
OpcPackage::OpcPackage() : m_impl(std::make_unique<Impl>()) {}

OpcPackage::~OpcPackage() = default;

OpcPackage::OpcPackage(OpcPackage&&) noexcept = default;
OpcPackage& OpcPackage::operator=(OpcPackage&&) noexcept = default;

void OpcPackage::open(const std::string& path) {
    m_impl->open(path);
}

void OpcPackage::close() {
    m_impl->close();
}

bool OpcPackage::isOpen() const {
    return m_impl->isOpen();
}

std::string OpcPackage::findWorkbookPath() const {
    return m_impl->findWorkbookPath();
}

std::vector<std::string> OpcPackage::getContentTypes() const {
    return m_impl->getContentTypes();
}

const ZipReader& OpcPackage::getZipReader() const {
    return m_impl->getZipReader();
}

} // namespace xlsxcsv::core
