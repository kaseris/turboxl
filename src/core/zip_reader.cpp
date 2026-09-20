#include "xlsxcsv/core.hpp"
#if __has_include(<minizip-ng/unzip.h>)
#  include <minizip-ng/unzip.h>
#elif __has_include(<minizip/unzip.h>)
#  include <minizip/unzip.h>
#else
#  error "minizip unzip.h header not found"
#endif
#if __has_include(<minizip-ng/ioapi.h>)
#  include <minizip-ng/ioapi.h>
#  define TURBOXL_MINIZIP_FILE_CALLBACKS 1
#elif __has_include(<minizip/ioapi.h>)
#  include <minizip/ioapi.h>
#  define TURBOXL_MINIZIP_FILE_CALLBACKS 1
#elif __has_include(<minizip/mz_strm_mem.h>)
#  include <minizip/mz.h>
#  include <minizip/mz_strm_mem.h>
#  define TURBOXL_MINIZIP_NATIVE_MEMORY_STREAM 1
#elif __has_include(<minizip-ng/mz_strm_mem.h>)
#  include <minizip-ng/mz.h>
#  include <minizip-ng/mz_strm_mem.h>
#  define TURBOXL_MINIZIP_NATIVE_MEMORY_STREAM 1
#else
#  error "minizip memory stream support not found"
#endif
#include <algorithm>
#include <array>
#include <cstring>
#include <filesystem>
#include <regex>
#include <unordered_map>
#include <limits>

namespace fs = std::filesystem;

namespace xlsxcsv::core {

namespace {

#if defined(TURBOXL_MINIZIP_FILE_CALLBACKS)
struct MemoryCursor {
    std::shared_ptr<const ByteVector> data;
    ZPOS64_T offset = 0;
    bool error = false;
};

voidpf ZCALLBACK openMemoryFile(voidpf opaque, const void*, int mode) {
    const bool readable = (mode & ZLIB_FILEFUNC_MODE_READ) != 0;
    const bool writable = (mode & ZLIB_FILEFUNC_MODE_WRITE) != 0;
    if (!readable || writable || opaque == nullptr) return nullptr;
    auto* data = static_cast<std::shared_ptr<const ByteVector>*>(opaque);
    return new MemoryCursor{*data};
}

uLong ZCALLBACK readMemoryFile(voidpf, voidpf stream, void* buffer, uLong size) {
    auto* cursor = static_cast<MemoryCursor*>(stream);
    if (!cursor || !buffer) return 0;
    const auto dataSize = static_cast<ZPOS64_T>(cursor->data->size());
    if (cursor->offset > dataSize) {
        cursor->error = true;
        return 0;
    }
    const auto remaining = dataSize - cursor->offset;
    const auto count = std::min<ZPOS64_T>(remaining, size);
    if (count != 0) {
        std::memcpy(buffer, cursor->data->data() + cursor->offset,
                    static_cast<size_t>(count));
        cursor->offset += count;
    }
    return static_cast<uLong>(count);
}

uLong ZCALLBACK writeMemoryFile(voidpf, voidpf stream, const void*, uLong) {
    if (auto* cursor = static_cast<MemoryCursor*>(stream)) cursor->error = true;
    return 0;
}

ZPOS64_T ZCALLBACK tellMemoryFile(voidpf, voidpf stream) {
    const auto* cursor = static_cast<MemoryCursor*>(stream);
    return cursor ? cursor->offset : 0;
}

long ZCALLBACK seekMemoryFile(voidpf, voidpf stream, ZPOS64_T offset, int origin) {
    auto* cursor = static_cast<MemoryCursor*>(stream);
    if (!cursor) return -1;
    const auto dataSize = static_cast<ZPOS64_T>(cursor->data->size());
    const auto signedOffset = static_cast<std::int64_t>(offset);
    std::int64_t base = 0;
    switch (origin) {
        case ZLIB_FILEFUNC_SEEK_SET:
            if (offset > dataSize) {
                cursor->error = true;
                return -1;
            }
            cursor->offset = offset;
            return 0;
        case ZLIB_FILEFUNC_SEEK_CUR:
            base = static_cast<std::int64_t>(cursor->offset);
            break;
        case ZLIB_FILEFUNC_SEEK_END:
            base = static_cast<std::int64_t>(dataSize);
            break;
        default:
            cursor->error = true;
            return -1;
    }
    if (signedOffset < -base ||
        signedOffset > static_cast<std::int64_t>(dataSize) - base) {
        cursor->error = true;
        return -1;
    }
    cursor->offset = static_cast<ZPOS64_T>(base + signedOffset);
    return 0;
}

int ZCALLBACK closeMemoryFile(voidpf, voidpf stream) {
    delete static_cast<MemoryCursor*>(stream);
    return 0;
}

int ZCALLBACK errorMemoryFile(voidpf, voidpf stream) {
    const auto* cursor = static_cast<MemoryCursor*>(stream);
    return cursor && cursor->error ? 1 : 0;
}

zlib_filefunc64_def memoryFileFunctions(
    std::shared_ptr<const ByteVector>* data) {
    zlib_filefunc64_def functions{};
    functions.zopen64_file = openMemoryFile;
    functions.zread_file = readMemoryFile;
    functions.zwrite_file = writeMemoryFile;
    functions.ztell64_file = tellMemoryFile;
    functions.zseek64_file = seekMemoryFile;
    functions.zclose_file = closeMemoryFile;
    functions.zerror_file = errorMemoryFile;
    functions.opaque = data;
    return functions;
}
#else
struct MemoryFileFunctions {};

unzFile openMemoryArchive(
    const std::shared_ptr<const ByteVector>& data,
    MemoryFileFunctions&) {
    if (data->size() > static_cast<size_t>(INT32_MAX)) return nullptr;

    void* stream = nullptr;
    if (mz_stream_mem_create(&stream) == nullptr) return nullptr;
    mz_stream_mem_set_buffer(
        stream, const_cast<unsigned char*>(data->data()),
        static_cast<int32_t>(data->size()));
    if (mz_stream_mem_open(stream, nullptr, MZ_OPEN_MODE_READ) != MZ_OK) {
        mz_stream_mem_delete(&stream);
        return nullptr;
    }

    unzFile file = unzOpen_MZ(stream);
    if (!file) {
        mz_stream_mem_close(stream);
        mz_stream_mem_delete(&stream);
    }
    return file;
}
#endif

#if defined(TURBOXL_MINIZIP_FILE_CALLBACKS)
using MemoryFileFunctions = zlib_filefunc64_def;

unzFile openMemoryArchive(
    std::shared_ptr<const ByteVector>& data,
    MemoryFileFunctions& functions) {
    functions = memoryFileFunctions(&data);
    return unzOpen2_64(nullptr, &functions);
}
#endif

} // namespace

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

    Impl(std::shared_ptr<const ByteVector> archiveData,
         const unz64_file_pos& position,
         size_t expectedSize)
        : m_archiveData(std::move(archiveData)),
          m_expectedSize(expectedSize) {
        m_file = openMemoryArchive(m_archiveData, m_fileFunctions);
        if (!m_file) {
            throw XlsxError("Failed to reopen memory ZIP archive for streaming");
        }
        if (unzGoToFilePos64(m_file, &position) != UNZ_OK ||
            unzOpenCurrentFile(m_file) != UNZ_OK) {
            unzClose(m_file);
            m_file = nullptr;
            throw XlsxError("Failed to open indexed memory ZIP entry for streaming");
        }
        m_currentOpen = true;
    }

    ~Impl() { close(); }

    size_t read(void* buffer, size_t size) {
        if (!m_currentOpen || size == 0) return 0;
        const size_t remaining = m_expectedSize - m_bytesRead;
        if (remaining == 0) {
            verifyEnd();
            return 0;
        }

        auto* output = static_cast<unsigned char*>(buffer);
        const size_t requested = std::min(size, remaining);
        size_t copied = 0;
        while (copied < requested) {
            if (m_bufferOffset == m_bufferSize) {
                fillBuffer();
            }
            const size_t available = m_bufferSize - m_bufferOffset;
            const size_t count = std::min(requested - copied, available);
            std::memcpy(output + copied, m_buffer.data() + m_bufferOffset, count);
            m_bufferOffset += count;
            copied += count;
        }
        m_bytesRead += copied;
        return copied;
    }

    size_t size() const { return m_expectedSize; }
    size_t bytesRead() const { return m_bytesRead; }

private:
    void fillBuffer() {
        const size_t remaining = m_expectedSize - m_inflatedBytes;
        if (remaining == 0) {
            throw XlsxError("ZIP entry ended before its declared size");
        }
        const size_t requested = std::min(remaining, m_buffer.size());
        const int result = unzReadCurrentFile(
            m_file, m_buffer.data(), static_cast<uint32_t>(requested));
        if (result < 0) throw XlsxError("Failed while streaming ZIP entry");
        if (result == 0) throw XlsxError("ZIP entry ended before its declared size");
        m_bufferOffset = 0;
        m_bufferSize = static_cast<size_t>(result);
        m_inflatedBytes += m_bufferSize;
    }

    void verifyEnd() {
        if (m_endVerified) return;
        unsigned char extra = 0;
        const int result = unzReadCurrentFile(m_file, &extra, 1);
        if (result < 0) throw XlsxError("Failed while streaming ZIP entry");
        if (result > 0) throw XlsxError("ZIP entry exceeds its declared size");
        m_endVerified = true;
    }

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
    std::shared_ptr<const ByteVector> m_archiveData;
    size_t m_expectedSize = 0;
    size_t m_bytesRead = 0;
    size_t m_inflatedBytes = 0;
    std::array<unsigned char, 128 * 1024> m_buffer{};
    size_t m_bufferOffset = 0;
    size_t m_bufferSize = 0;
    bool m_currentOpen = false;
    bool m_endVerified = false;
    MemoryFileFunctions m_fileFunctions{};
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
        close();
        
        if (!fs::exists(path)) {
            throw XlsxError("ZIP file does not exist: " + path);
        }
        
        m_unzFile = unzOpen64(path.c_str());
        if (!m_unzFile) {
            throw XlsxError("Failed to open ZIP file: " + path);
        }
        
        m_archivePath = path;
        try {
            validateAndIndex();
            m_isOpen = true;
        } catch (...) {
            close();
            throw;
        }
    }

    void open(ByteVector data) {
        close();
        if (data.empty()) {
            throw XlsxError("Memory ZIP archive is empty");
        }
        if (data.size() > m_limits.maxArchiveSize) {
            throw XlsxError("Memory ZIP archive exceeds size limit");
        }
        m_archiveData = std::make_shared<const ByteVector>(std::move(data));
        m_unzFile = openMemoryArchive(m_archiveData, m_fileFunctions);
        if (!m_unzFile) {
            m_archiveData.reset();
            throw XlsxError("Failed to open memory ZIP archive");
        }
        try {
            validateAndIndex();
            m_isOpen = true;
        } catch (...) {
            close();
            throw;
        }
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
        m_archiveData.reset();
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
        std::shared_ptr<const ByteVector> archiveData;
        unz64_file_pos position;
        size_t uncompressedSize;
    };

    StreamSpec streamSpec(const std::string& path) const {
        const auto& record = findRecord(path);
        return {m_archivePath, m_archiveData, record.position,
                record.entry.uncompressedSize};
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
    std::shared_ptr<const ByteVector> m_archiveData;
    MemoryFileFunctions m_fileFunctions{};
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

void ZipReader::open(ByteVector data) {
    m_impl->open(std::move(data));
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
    auto impl = spec.archiveData
        ? std::make_unique<ZipEntryStream::Impl>(
              std::move(spec.archiveData), spec.position, spec.uncompressedSize)
        : std::make_unique<ZipEntryStream::Impl>(
              spec.archivePath, spec.position, spec.uncompressedSize);
    return std::unique_ptr<ZipEntryStream>(new ZipEntryStream(std::move(impl)));
}

const ZipSecurityLimits& ZipReader::getSecurityLimits() const {
    return m_impl->getSecurityLimits();
}

} // namespace xlsxcsv::core
