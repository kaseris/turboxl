#include "xlsxcsv/core.hpp"
#include <minizip/unzip.h>
#include <algorithm>
#include <filesystem>
#include <regex>

namespace fs = std::filesystem;

namespace xlsxcsv::core {

class ZipReader::Impl {
public:
    explicit Impl(const ZipSecurityLimits& limits) 
        : m_limits(limits), m_unzFile(nullptr) {}
    
    ~Impl() {
        close();
    }
    
    void open(const std::string& path) {
        if (m_unzFile) {
            close();
        }
        
        if (!fs::exists(path)) {
            throw XlsxError("ZIP file does not exist: " + path);
        }
        
        m_unzFile = unzOpen64(path.c_str());
        if (!m_unzFile) {
            throw XlsxError("Failed to open ZIP file: " + path);
        }
        
        validateZipSecurity();
        m_isOpen = true;
    }
    
    void close() {
        if (m_unzFile) {
            unzClose(m_unzFile);
            m_unzFile = nullptr;
        }
        m_isOpen = false;
        m_entries.clear();
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::vector<ZipEntry> listEntries() {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        if (!m_entries.empty()) {
            return m_entries; // Return cached entries
        }
        
        int result = unzGoToFirstFile(m_unzFile);
        if (result != UNZ_OK && result != UNZ_END_OF_LIST_OF_FILE) {
            throw XlsxError("Failed to navigate to first ZIP entry");
        }
        
        while (result == UNZ_OK) {
            unz_file_info64 fileInfo;
            char filename[1024] = {0};
            
            result = unzGetCurrentFileInfo64(m_unzFile, &fileInfo, filename, sizeof(filename), nullptr, 0, nullptr, 0);
            if (result != UNZ_OK) {
                break;
            }
            
            ZipEntry entry;
            entry.path = sanitizePath(filename);
            entry.compressedSize = fileInfo.compressed_size;
            entry.uncompressedSize = fileInfo.uncompressed_size;
            entry.isEncrypted = (fileInfo.flag & 1) != 0; // UNZ_FLAG_ENCRYPTED
            
            // Skip entries with suspicious paths
            if (entry.path.empty() || isPathSuspicious(entry.path)) {
                result = unzGoToNextFile(m_unzFile);
                continue;
            }
            
            // Check security limits
            if (entry.uncompressedSize > m_limits.maxEntrySize) {
                throw XlsxError("ZIP entry exceeds size limit: " + entry.path);
            }
            
            if (entry.isEncrypted) {
                throw XlsxError("Encrypted ZIP entries are not supported: " + entry.path);
            }
            
            m_entries.push_back(entry);
            result = unzGoToNextFile(m_unzFile);
        }
        
        if (m_entries.size() > m_limits.maxEntries) {
            throw XlsxError("ZIP file contains too many entries: " + 
                          std::to_string(m_entries.size()) + " > " + 
                          std::to_string(m_limits.maxEntries));
        }
        
        return m_entries;
    }
    
    bool hasEntry(const std::string& path) {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        auto entries = listEntries();
        return std::find_if(entries.begin(), entries.end(),
                           [&path](const ZipEntry& entry) {
                               return entry.path == path;
                           }) != entries.end();
    }
    
    ByteVector readEntry(const std::string& path) {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        if (isPathSuspicious(path)) {
            throw XlsxError("Suspicious path rejected: " + path);
        }
        
        int result = unzLocateFile(m_unzFile, path.c_str(), 0);
        if (result != UNZ_OK) {
            throw XlsxError("ZIP entry not found: " + path);
        }
        
        unz_file_info64 fileInfo;
        result = unzGetCurrentFileInfo64(m_unzFile, &fileInfo, nullptr, 0, nullptr, 0, nullptr, 0);
        if (result != UNZ_OK) {
            throw XlsxError("Failed to get ZIP entry info: " + path);
        }
        
        if (fileInfo.uncompressed_size > m_limits.maxEntrySize) {
            throw XlsxError("ZIP entry exceeds size limit: " + path);
        }
        
        if (fileInfo.flag & 1) { // UNZ_FLAG_ENCRYPTED
            throw XlsxError("Encrypted ZIP entries are not supported: " + path);
        }
        
        result = unzOpenCurrentFile(m_unzFile);
        if (result != UNZ_OK) {
            throw XlsxError("Failed to open ZIP entry: " + path);
        }
        
        // Use chunked reading with 512 KiB buffers for better I/O efficiency
        static constexpr size_t BUFFER_SIZE = 512 * 1024; // 512 KiB chunks
        ByteVector data;
        data.reserve(fileInfo.uncompressed_size);
        
        ByteVector buffer(BUFFER_SIZE);
        
        while (true) {
            int bytesRead = unzReadCurrentFile(m_unzFile, buffer.data(), buffer.size());
            if (bytesRead < 0) {
                unzCloseCurrentFile(m_unzFile);
                throw XlsxError("Failed to read ZIP entry: " + path);
            }
            
            if (bytesRead == 0) {
                break; // End of file
            }
            
            // Append the chunk to our data
            data.insert(data.end(), buffer.begin(), buffer.begin() + bytesRead);
        }
        
        unzCloseCurrentFile(m_unzFile);
        return data;
    }
    
    std::string readEntryAsString(const std::string& path) {
        auto data = readEntry(path);
        return std::string(data.begin(), data.end());
    }
    
    const ZipSecurityLimits& getSecurityLimits() const {
        return m_limits;
    }

private:
    void validateZipSecurity() {
        // Get total uncompressed size to check against limit
        size_t totalUncompressed = 0;
        size_t entryCount = 0;
        
        int result = unzGoToFirstFile(m_unzFile);
        while (result == UNZ_OK) {
            unz_file_info64 fileInfo;
            result = unzGetCurrentFileInfo64(m_unzFile, &fileInfo, nullptr, 0, nullptr, 0, nullptr, 0);
            if (result != UNZ_OK) {
                break;
            }
            
            totalUncompressed += fileInfo.uncompressed_size;
            entryCount++;
            
            if (entryCount > m_limits.maxEntries) {
                throw XlsxError("ZIP file contains too many entries");
            }
            
            if (totalUncompressed > m_limits.maxTotalUncompressed) {
                throw XlsxError("ZIP file total uncompressed size exceeds limit");
            }
            
            result = unzGoToNextFile(m_unzFile);
        }
    }
    
    std::string sanitizePath(const std::string& path) {
        // Normalize path separators and remove dangerous sequences
        std::string sanitized = path;
        
        // Replace backslashes with forward slashes
        std::replace(sanitized.begin(), sanitized.end(), '\\', '/');
        
        // Remove leading slashes
        while (!sanitized.empty() && sanitized[0] == '/') {
            sanitized.erase(0, 1);
        }
        
        return sanitized;
    }
    
    bool isPathSuspicious(const std::string& path) {
        // Check for path traversal attempts
        if (path.find("..") != std::string::npos) {
            return true;
        }
        
        // Check for absolute paths
        if (!path.empty() && path[0] == '/') {
            return true;
        }
        
        // Check for null bytes
        if (path.find('\0') != std::string::npos) {
            return true;
        }
        
        // Check for extremely long paths
        if (path.length() > 1024) {
            return true;
        }
        
        return false;
    }
    
    ZipSecurityLimits m_limits;
    unzFile m_unzFile;
    bool m_isOpen = false;
    std::vector<ZipEntry> m_entries; // Cache entries after first list
};

// ZipReader implementation
ZipReader::ZipReader(const ZipSecurityLimits& limits) 
    : m_impl(std::make_unique<Impl>(limits)) {}

ZipReader::~ZipReader() = default;

ZipReader::ZipReader(ZipReader&&) noexcept = default;
ZipReader& ZipReader::operator=(ZipReader&&) noexcept = default;

void ZipReader::open(const std::string& path) {
    m_impl->open(path);
}

void ZipReader::close() {
    m_impl->close();
}

bool ZipReader::isOpen() const {
    return m_impl->isOpen();
}

std::vector<ZipEntry> ZipReader::listEntries() const {
    return m_impl->listEntries();
}

bool ZipReader::hasEntry(const std::string& path) const {
    return m_impl->hasEntry(path);
}

ByteVector ZipReader::readEntry(const std::string& path) const {
    return m_impl->readEntry(path);
}

std::string ZipReader::readEntryAsString(const std::string& path) const {
    return m_impl->readEntryAsString(path);
}

const ZipSecurityLimits& ZipReader::getSecurityLimits() const {
    return m_impl->getSecurityLimits();
}

} // namespace xlsxcsv::core
