#include "xlsxcsv/core.hpp"
#if __has_include(<minizip-ng/unzip.h>)
#  include <minizip-ng/unzip.h>
#elif __has_include(<minizip/unzip.h>)
#  include <minizip/unzip.h>
#else
#  error "minizip unzip.h header not found"
#endif
#include <algorithm>
#include <filesystem>
#include <regex>
#include <unordered_map>
#include <limits>

namespace fs = std::filesystem;

namespace xlsxcsv::core {

class ZipEntryStream::Impl {
public:
    Impl(const std::string& archivePath,
         const unz64_file_pos& position,
         size_t expectedSize)
        : m_expectedSize(expectedSize) {
        m_file = unzOpen64(archivePath.c_str());
        if (!m_file) {
            throw XlsxError("Failed to reopen ZIP file for streaming: " + archivePath);
        }
        if (unzGoToFilePos64(m_file, &position) != UNZ_OK ||
            unzOpenCurrentFile(m_file) != UNZ_OK) {
            unzClose(m_file);
            m_file = nullptr;
            throw XlsxError("Failed to open indexed ZIP entry for streaming");
        }
        m_currentOpen = true;
    }

    ~Impl() { close(); }

    size_t read(void* buffer, size_t size) {
        if (!m_currentOpen || size == 0) return 0;
        const size_t remaining = m_expectedSize - m_bytesRead;
        if (remaining == 0) {
            unsigned char extra = 0;
            const int result = unzReadCurrentFile(m_file, &extra, 1);
            if (result < 0) throw XlsxError("Failed while streaming ZIP entry");
            if (result > 0) throw XlsxError("ZIP entry exceeds its declared size");
            return 0;
        }
        const size_t requested = std::min({size, remaining,
            static_cast<size_t>(std::numeric_limits<int>::max())});
        const int result = unzReadCurrentFile(m_file, buffer, static_cast<uint32_t>(requested));
        if (result < 0) throw XlsxError("Failed while streaming ZIP entry");
        if (result == 0 && m_bytesRead != m_expectedSize) {
            throw XlsxError("ZIP entry ended before its declared size");
        }
        m_bytesRead += static_cast<size_t>(result);
        return static_cast<size_t>(result);
    }

    size_t size() const { return m_expectedSize; }
    size_t bytesRead() const { return m_bytesRead; }

private:
    void close() noexcept {
        if (m_currentOpen) {
            unzCloseCurrentFile(m_file);
            m_currentOpen = false;
        }
        if (m_file) {
            unzClose(m_file);
            m_file = nullptr;
        }
    }

    unzFile m_file = nullptr;
    size_t m_expectedSize = 0;
    size_t m_bytesRead = 0;
    bool m_currentOpen = false;
};

ZipEntryStream::ZipEntryStream(std::unique_ptr<Impl> impl) : m_impl(std::move(impl)) {}
ZipEntryStream::~ZipEntryStream() = default;
ZipEntryStream::ZipEntryStream(ZipEntryStream&&) noexcept = default;
ZipEntryStream& ZipEntryStream::operator=(ZipEntryStream&&) noexcept = default;
size_t ZipEntryStream::read(void* buffer, size_t size) { return m_impl->read(buffer, size); }
size_t ZipEntryStream::size() const { return m_impl->size(); }
size_t ZipEntryStream::bytesRead() const { return m_impl->bytesRead(); }

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
        
        m_archivePath = path;
        validateAndIndex();
        m_isOpen = true;
    }
    
    void close() {
        if (m_unzFile) {
            unzClose(m_unzFile);
            m_unzFile = nullptr;
        }
        m_isOpen = false;
        m_entries.clear();
        m_records.clear();
        m_entryIndex.clear();
        m_archivePath.clear();
    }
    
    bool isOpen() const {
        return m_isOpen;
    }
    
    std::vector<ZipEntry> listEntries() {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        return m_entries;
    }
    
    bool hasEntry(const std::string& path) {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        return m_entryIndex.find(path) != m_entryIndex.end();
    }
    
    ByteVector readEntry(const std::string& path) {
        if (!m_isOpen) {
            throw XlsxError("ZIP file is not open");
        }
        
        if (isPathSuspicious(path)) {
            throw XlsxError("Suspicious path rejected: " + path);
        }
        
        auto record = findRecord(path);
        int result = unzGoToFilePos64(m_unzFile, &record.position);
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
        
        ByteVector data(static_cast<size_t>(fileInfo.uncompressed_size));
        size_t offset = 0;
        while (offset < data.size()) {
            const size_t request = std::min(data.size() - offset,
                static_cast<size_t>(std::numeric_limits<int>::max()));
            int bytesRead = unzReadCurrentFile(m_unzFile, data.data() + offset,
                                               static_cast<uint32_t>(request));
            if (bytesRead < 0) {
                unzCloseCurrentFile(m_unzFile);
                throw XlsxError("Failed to read ZIP entry: " + path);
            }
            if (bytesRead == 0) {
                unzCloseCurrentFile(m_unzFile);
                throw XlsxError("ZIP entry ended before its declared size: " + path);
            }
            offset += static_cast<size_t>(bytesRead);
        }
        unsigned char extra = 0;
        const int trailingRead = unzReadCurrentFile(m_unzFile, &extra, 1);
        if (trailingRead < 0) {
            unzCloseCurrentFile(m_unzFile);
            throw XlsxError("Failed to finish reading ZIP entry: " + path);
        }
        if (trailingRead > 0) {
            unzCloseCurrentFile(m_unzFile);
            throw XlsxError("ZIP entry exceeds its declared size: " + path);
        }
        unzCloseCurrentFile(m_unzFile);
        return data;
    }

    struct StreamSpec {
        std::string archivePath;
        unz64_file_pos position;
        size_t uncompressedSize;
    };

    StreamSpec streamSpec(const std::string& path) const {
        const auto& record = findRecord(path);
        return {m_archivePath, record.position, record.entry.uncompressedSize};
    }
    
    std::string readEntryAsString(const std::string& path) {
        auto data = readEntry(path);
        return std::string(data.begin(), data.end());
    }
    
    const ZipSecurityLimits& getSecurityLimits() const {
        return m_limits;
    }

private:
    struct EntryRecord {
        ZipEntry entry;
        unz64_file_pos position{};
    };

    const EntryRecord& findRecord(const std::string& path) const {
        const auto it = m_entryIndex.find(path);
        if (it == m_entryIndex.end()) throw XlsxError("ZIP entry not found: " + path);
        return m_records[it->second];
    }

    void validateAndIndex() {
        size_t totalUncompressed = 0;
        size_t entryCount = 0;
        m_entries.clear();
        m_records.clear();
        m_entryIndex.clear();
        int result = unzGoToFirstFile(m_unzFile);
        while (result == UNZ_OK) {
            unz_file_info64 fileInfo;
            char filename[1024] = {0};
            result = unzGetCurrentFileInfo64(m_unzFile, &fileInfo, filename,
                sizeof(filename), nullptr, 0, nullptr, 0);
            if (result != UNZ_OK) {
                throw XlsxError("Failed to inspect ZIP entry");
            }
            EntryRecord record;
            record.entry.path = sanitizePath(filename);
            record.entry.compressedSize = static_cast<size_t>(fileInfo.compressed_size);
            record.entry.uncompressedSize = static_cast<size_t>(fileInfo.uncompressed_size);
            record.entry.isEncrypted = (fileInfo.flag & 1) != 0;
            if (record.entry.path.empty() || isPathSuspicious(record.entry.path)) {
                throw XlsxError("Suspicious ZIP entry path rejected: " + std::string(filename));
            }
            if (record.entry.uncompressedSize > m_limits.maxEntrySize) {
                throw XlsxError("ZIP entry exceeds size limit: " + record.entry.path);
            }
            if (record.entry.isEncrypted) {
                throw XlsxError("Encrypted ZIP entries are not supported: " + record.entry.path);
            }
            if (unzGetFilePos64(m_unzFile, &record.position) != UNZ_OK) {
                throw XlsxError("Failed to index ZIP entry: " + record.entry.path);
            }
            if (record.entry.uncompressedSize > m_limits.maxTotalUncompressed ||
                totalUncompressed > m_limits.maxTotalUncompressed - record.entry.uncompressedSize) {
                throw XlsxError("ZIP file total uncompressed size exceeds limit");
            }
            totalUncompressed += record.entry.uncompressedSize;
            entryCount++;
            if (entryCount > m_limits.maxEntries) {
                throw XlsxError("ZIP file contains too many entries");
            }
            if (!m_entryIndex.emplace(record.entry.path, m_records.size()).second) {
                throw XlsxError("Duplicate ZIP entry: " + record.entry.path);
            }
            m_entries.push_back(record.entry);
            m_records.push_back(std::move(record));
            result = unzGoToNextFile(m_unzFile);
        }
        if (result != UNZ_END_OF_LIST_OF_FILE) {
            throw XlsxError("Failed while indexing ZIP archive");
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
    std::string m_archivePath;
    std::vector<ZipEntry> m_entries;
    std::vector<EntryRecord> m_records;
    std::unordered_map<std::string, size_t> m_entryIndex;
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

std::unique_ptr<ZipEntryStream> ZipReader::openEntryStream(const std::string& path) const {
    auto spec = m_impl->streamSpec(path);
    auto impl = std::make_unique<ZipEntryStream::Impl>(
        spec.archivePath, spec.position, spec.uncompressedSize);
    return std::unique_ptr<ZipEntryStream>(new ZipEntryStream(std::move(impl)));
}

const ZipSecurityLimits& ZipReader::getSecurityLimits() const {
    return m_impl->getSecurityLimits();
}

} // namespace xlsxcsv::core
