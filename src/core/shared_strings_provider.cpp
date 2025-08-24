#include "xlsxcsv/core.hpp"
#include <libxml/xmlreader.h>
#include <vector>
#include <fstream>
#include <memory>
#include <sstream>
#include <regex>
#include <filesystem>
#include <algorithm>

namespace xlsxcsv::core {

class SharedStringsProvider::Impl {
public:
    Impl() : m_config(), m_isOpen(false), m_activeMode(SharedStringsMode::Auto), 
             m_stringCount(0), m_memoryUsage(0), m_arenaCapacity(0), m_isUsingDisk(false) {}
    
    explicit Impl(const SharedStringsConfig& config) : m_config(config), m_isOpen(false), 
                                                       m_activeMode(config.mode), m_stringCount(0), 
                                                       m_memoryUsage(0), m_arenaCapacity(0), m_isUsingDisk(false) {}
    
    ~Impl() {
        close();
    }
    
    void parse(const OpcPackage& package) {
        if (m_isOpen) {
            close();
        }
        
        // Check if sharedStrings.xml exists
        std::string sharedStringsPath = "xl/sharedStrings.xml";
        if (!package.getZipReader().hasEntry(sharedStringsPath)) {
            // No shared strings - this is valid for some XLSX files
            m_isOpen = true;
            m_stringCount = 0;
            return;
        }
        
        m_xmlData = package.getZipReader().readEntry(sharedStringsPath);
        parseSharedStringsXml(m_xmlData);
        
        m_isOpen = true;
    }
    
    void close() {
        m_isOpen = false;
        m_arena.clear();
        m_offsets.clear();
        m_arenaCapacity = 0;
        m_stringCount = 0;
        m_memoryUsage = 0;
        
        if (m_isUsingDisk && !m_diskFilePath.empty()) {
            std::filesystem::remove(m_diskFilePath);
            m_diskFilePath.clear();
        }
        m_isUsingDisk = false;
        m_activeMode = m_config.mode;
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::string getString(size_t index) const {
        auto result = tryGetString(index);
        if (!result.has_value()) {
            throw XlsxError("Shared string index " + std::to_string(index) + " out of range");
        }
        return result.value();
    }
    
    std::optional<std::string> tryGetString(size_t index) const {
        if (!m_isOpen || index >= m_stringCount) {
            return std::nullopt;
        }
        
        if (m_stringCount == 0) {
            return std::nullopt;
        }
        
        if (m_isUsingDisk) {
            return readStringFromDisk(index);
        } else {
            return getStringFromArena(index);
        }
    }
    
    std::optional<std::string> getStringFromArena(size_t index) const {
        if (index >= m_offsets.size()) {
            return std::nullopt;
        }
        
        uint32_t offset = m_offsets[index];
        if (offset >= m_arena.size()) {
            return std::nullopt;
        }
        
        // Strings are null-terminated in arena
        const char* str = reinterpret_cast<const char*>(&m_arena[offset]);
        return std::string(str);
    }
    
    size_t getStringCount() const {
        return m_stringCount;
    }
    
    bool hasStrings() const {
        return m_isOpen && m_stringCount > 0;
    }
    
    const SharedStringsConfig& getConfig() const {
        return m_config;
    }
    
    SharedStringsMode getActiveMode() const {
        return m_activeMode;
    }
    
    size_t getMemoryUsage() const {
        return m_memoryUsage;
    }
    
    bool isUsingDisk() const {
        return m_isUsingDisk;
    }

private:
    void parseSharedStringsXml(const ByteVector& xmlData) {
        xmlTextReaderPtr reader = xmlReaderForMemory(
            reinterpret_cast<const char*>(xmlData.data()),
            static_cast<int>(xmlData.size()),
            nullptr, nullptr,
            XML_PARSE_RECOVER | XML_PARSE_NONET | XML_PARSE_NOENT
        );
        
        if (!reader) {
            throw XlsxError("Failed to create XML reader for sharedStrings.xml");
        }
        
        try {
            parseWithReader(reader);
        } catch (...) {
            xmlFreeTextReader(reader);
            throw;
        }
        
        xmlFreeTextReader(reader);
    }
    
    void parseWithReader(xmlTextReaderPtr reader) {
        size_t estimatedSize = 0;
        m_stringCount = 0;
        
        // First pass: count strings and estimate memory usage
        while (xmlTextReaderRead(reader) == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                auto nodeName = getNodeName(reader);
                if (nodeName == "sst") {
                    // Get count attribute if available
                    auto countAttr = getAttribute(reader, "count");
                    if (countAttr.has_value()) {
                        try {
                            m_stringCount = std::stoull(countAttr.value());
                        } catch (...) {
                            // If parsing fails, we'll count manually
                        }
                    }
                } else if (nodeName == "si") {
                    if (m_stringCount == 0) {
                        m_stringCount++; // Count manually if no count attribute
                    }
                    // Rough estimation: average 50 bytes per string
                    estimatedSize += 50;
                }
            }
        }
        
        // Decide storage mode
        decideStorageMode(estimatedSize);
        
        // Re-create reader for actual parsing (the original reader is consumed by first pass)
        auto reader2 = xmlReaderForMemory(
            reinterpret_cast<const char*>(m_xmlData.data()),
            static_cast<int>(m_xmlData.size()),
            nullptr, nullptr,
            XML_PARSE_RECOVER | XML_PARSE_NONET | XML_PARSE_NOENT
        );
        
        if (!reader2) {
            throw XlsxError("Failed to re-create XML reader for sharedStrings.xml");
        }
        
        try {
            parseStrings(reader2);
        } catch (...) {
            xmlFreeTextReader(reader2);
            throw;
        }
        
        xmlFreeTextReader(reader2);
    }
    
    void decideStorageMode(size_t estimatedSize) {
        switch (m_config.mode) {
            case SharedStringsMode::InMemory:
                m_activeMode = SharedStringsMode::InMemory;
                m_isUsingDisk = false;
                break;
            case SharedStringsMode::External:
                m_activeMode = SharedStringsMode::External;
                initializeDiskStorage();
                break;
            case SharedStringsMode::Auto:
                if (estimatedSize > m_config.memoryThreshold) {
                    m_activeMode = SharedStringsMode::External;
                    initializeDiskStorage();
                } else {
                    m_activeMode = SharedStringsMode::InMemory;
                    m_isUsingDisk = false;
                }
                break;
        }
        
        if (!m_isUsingDisk && m_stringCount > 0) {
            // Initialize arena for estimated size
            m_arenaCapacity = std::max(INITIAL_ARENA_SIZE, estimatedSize * 2);
            m_arena.reserve(m_arenaCapacity);
            m_offsets.reserve(m_stringCount + 1); // +1 for sentinel
        }
    }
    
    void initializeDiskStorage() {
        m_isUsingDisk = true;
        
        // Create temporary file for disk storage
        auto tempDir = std::filesystem::temp_directory_path();
        m_diskFilePath = tempDir / ("turboxl_strings_" + std::to_string(reinterpret_cast<uintptr_t>(this)) + ".tmp");
        
        m_diskFile.open(m_diskFilePath, std::ios::binary | std::ios::out | std::ios::in | std::ios::trunc);
        if (!m_diskFile.is_open()) {
            throw XlsxError("Failed to create temporary file for shared strings storage");
        }
    }
    
    void parseStrings(xmlTextReaderPtr reader) {
        size_t currentIndex = 0;
        
        while (xmlTextReaderRead(reader) == 1) {
            if (xmlTextReaderNodeType(reader) == XML_READER_TYPE_ELEMENT) {
                auto nodeName = getNodeName(reader);
                if (nodeName == "si") {
                    auto stringValue = parseStringItem(reader);
                    storeString(currentIndex, stringValue);
                    currentIndex++;
                }
            }
        }
        
        m_stringCount = currentIndex;
        
        if (m_isUsingDisk) {
            m_diskFile.flush();
        }
    }
    
    std::string parseStringItem(xmlTextReaderPtr reader) {
        std::ostringstream result;
        int depth = xmlTextReaderDepth(reader);
        
        // Read until we close the <si> element
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth) {
                break; // We've closed the <si> element
            }
            
            int nodeType = xmlTextReaderNodeType(reader);
            if (nodeType == XML_READER_TYPE_ELEMENT) {
                auto nodeName = getNodeName(reader);
                if (nodeName == "t") {
                    // Text element
                    auto text = getElementText(reader);
                    if (text.has_value()) {
                        result << text.value();
                    }
                } else if (nodeName == "r" && m_config.flattenRichText) {
                    // Rich text run - extract text from <t> children
                    auto richText = parseRichTextRun(reader);
                    result << richText;
                }
            }
        }
        
        auto resultStr = result.str();
        
        // Truncate if too long
        if (resultStr.length() > m_config.maxStringLength) {
            resultStr = resultStr.substr(0, m_config.maxStringLength);
        }
        
        return resultStr;
    }
    
    std::string parseRichTextRun(xmlTextReaderPtr reader) {
        std::ostringstream result;
        int depth = xmlTextReaderDepth(reader);
        
        while (xmlTextReaderRead(reader) == 1) {
            int currentDepth = xmlTextReaderDepth(reader);
            if (currentDepth <= depth) {
                break;
            }
            
            int nodeType = xmlTextReaderNodeType(reader);
            if (nodeType == XML_READER_TYPE_ELEMENT) {
                auto nodeName = getNodeName(reader);
                if (nodeName == "t") {
                    auto text = getElementText(reader);
                    if (text.has_value()) {
                        result << text.value();
                    }
                }
            }
        }
        
        return result.str();
    }
    
    void storeString(size_t index, const std::string& value) {
        if (m_isUsingDisk) {
            storeStringToDisk(index, value);
        } else {
            storeStringToArena(index, value);
        }
    }
    
    void storeStringToArena(size_t index, const std::string& value) {
        // Ensure arena has enough space (doubling strategy)
        size_t needed = value.size() + 1; // +1 for null terminator
        if (m_arena.size() + needed > m_arena.capacity()) {
            // Double arena capacity when needed
            size_t newCapacity = std::max(m_arena.capacity() * 2, m_arena.size() + needed);
            m_arena.reserve(newCapacity);
            m_arenaCapacity = newCapacity;
        }
        
        // Store offset before adding string
        uint32_t offset = static_cast<uint32_t>(m_arena.size());
        if (index >= m_offsets.size()) {
            m_offsets.resize(index + 1);
        }
        m_offsets[index] = offset;
        
        // Append string to arena with null terminator
        m_arena.insert(m_arena.end(), value.begin(), value.end());
        m_arena.push_back('\0');
        
        m_memoryUsage += needed;
    }
    
    void storeStringToDisk(size_t index, const std::string& value) {
        // Simple format: [length][string] for each entry
        uint32_t length = static_cast<uint32_t>(value.size());
        
        m_diskFile.seekp(0, std::ios::end);
        m_diskFile.write(reinterpret_cast<const char*>(&length), sizeof(length));
        m_diskFile.write(value.data(), length);
        
        // Store offset for quick lookup
        if (index >= m_diskOffsets.size()) {
            m_diskOffsets.resize(index + 1);
        }
        m_diskOffsets[index] = m_diskFile.tellp() - static_cast<std::streampos>(sizeof(length) + length);
    }
    
    std::optional<std::string> readStringFromDisk(size_t index) const {
        if (!m_isUsingDisk || index >= m_diskOffsets.size()) {
            return std::nullopt;
        }
        
        auto& file = const_cast<std::fstream&>(m_diskFile);
        file.seekg(m_diskOffsets[index]);
        
        uint32_t length;
        file.read(reinterpret_cast<char*>(&length), sizeof(length));
        if (file.gcount() != sizeof(length)) {
            return std::nullopt;
        }
        
        std::string result(length, '\0');
        file.read(result.data(), length);
        if (file.gcount() != static_cast<std::streamsize>(length)) {
            return std::nullopt;
        }
        
        return result;
    }
    
    std::string getNodeName(xmlTextReaderPtr reader) {
        xmlChar* name = xmlTextReaderLocalName(reader);
        if (!name) return "";
        
        std::string result(reinterpret_cast<const char*>(name));
        xmlFree(name);
        return result;
    }
    
    std::optional<std::string> getAttribute(xmlTextReaderPtr reader, const std::string& attrName) {
        xmlChar* value = xmlTextReaderGetAttribute(reader, reinterpret_cast<const xmlChar*>(attrName.c_str()));
        if (!value) return std::nullopt;
        
        std::string result(reinterpret_cast<const char*>(value));
        xmlFree(value);
        return result;
    }
    
    std::optional<std::string> getElementText(xmlTextReaderPtr reader) {
        xmlChar* value = xmlTextReaderReadString(reader);
        if (!value) return std::nullopt;
        
        std::string result(reinterpret_cast<const char*>(value));
        xmlFree(value);
        return result;
    }

private:
    SharedStringsConfig m_config;
    bool m_isOpen;
    SharedStringsMode m_activeMode;
    size_t m_stringCount;
    size_t m_memoryUsage;
    
    // Arena-based storage (performance optimization)
    std::vector<uint8_t> m_arena;          // Single arena buffer for all strings
    std::vector<uint32_t> m_offsets;       // Start offset of each string in arena
    size_t m_arenaCapacity;                // Current arena capacity
    static constexpr size_t INITIAL_ARENA_SIZE = 8 * 1024 * 1024;  // 8MB initial
    
    // Disk storage
    bool m_isUsingDisk;
    std::filesystem::path m_diskFilePath;
    mutable std::fstream m_diskFile;
    std::vector<std::streampos> m_diskOffsets;
    
    // Keep XML data for re-parsing
    ByteVector m_xmlData;
};

// SharedStringsProvider implementation

SharedStringsProvider::SharedStringsProvider() : m_impl(std::make_unique<Impl>()) {}

SharedStringsProvider::SharedStringsProvider(const SharedStringsConfig& config) 
    : m_impl(std::make_unique<Impl>(config)) {}

SharedStringsProvider::~SharedStringsProvider() = default;

SharedStringsProvider::SharedStringsProvider(SharedStringsProvider&&) noexcept = default;

SharedStringsProvider& SharedStringsProvider::operator=(SharedStringsProvider&&) noexcept = default;

void SharedStringsProvider::parse(const OpcPackage& package) {
    m_impl->parse(package);
}

void SharedStringsProvider::close() {
    m_impl->close();
}

bool SharedStringsProvider::isOpen() const {
    return m_impl->isOpen();
}

std::string SharedStringsProvider::getString(size_t index) const {
    return m_impl->getString(index);
}

std::optional<std::string> SharedStringsProvider::tryGetString(size_t index) const {
    return m_impl->tryGetString(index);
}

size_t SharedStringsProvider::getStringCount() const {
    return m_impl->getStringCount();
}

bool SharedStringsProvider::hasStrings() const {
    return m_impl->hasStrings();
}

const SharedStringsConfig& SharedStringsProvider::getConfig() const {
    return m_impl->getConfig();
}

SharedStringsMode SharedStringsProvider::getActiveMode() const {
    return m_impl->getActiveMode();
}

size_t SharedStringsProvider::getMemoryUsage() const {
    return m_impl->getMemoryUsage();
}

bool SharedStringsProvider::isUsingDisk() const {
    return m_impl->isUsingDisk();
}

} // namespace xlsxcsv::core
