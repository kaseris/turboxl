#pragma once

#include <memory>
#include <string>
#include <vector>
#include <functional>
#include <optional>
#include <variant>

namespace xlsxcsv::core {

// Common types
using StringRef = std::string_view;
using ByteVector = std::vector<uint8_t>;

// Error handling
class XlsxError : public std::runtime_error {
public:
    explicit XlsxError(const std::string& message) : std::runtime_error(message) {}
};

// Security limits for ZIP operations
struct ZipSecurityLimits {
    size_t maxEntries = 10000;
    size_t maxEntrySize = 256 * 1024 * 1024; // 256MB
    size_t maxTotalUncompressed = 2ULL * 1024 * 1024 * 1024; // 2GB
};

// ZIP entry information
struct ZipEntry {
    std::string path;
    size_t compressedSize;
    size_t uncompressedSize;
    bool isEncrypted;
};

// Forward declarations
class ZipReader {
public:
    ZipReader(const ZipSecurityLimits& limits = {});
    ~ZipReader();
    
    // Non-copyable but movable
    ZipReader(const ZipReader&) = delete;
    ZipReader& operator=(const ZipReader&) = delete;
    ZipReader(ZipReader&&) noexcept;
    ZipReader& operator=(ZipReader&&) noexcept;
    
    void open(const std::string& path);
    void close();
    bool isOpen() const;
    
    std::vector<ZipEntry> listEntries() const;
    bool hasEntry(const std::string& path) const;
    ByteVector readEntry(const std::string& path) const;
    std::string readEntryAsString(const std::string& path) const;
    
    const ZipSecurityLimits& getSecurityLimits() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

class OpcPackage {
public:
    OpcPackage();
    ~OpcPackage();
    
    // Non-copyable but movable
    OpcPackage(const OpcPackage&) = delete;
    OpcPackage& operator=(const OpcPackage&) = delete;
    OpcPackage(OpcPackage&&) noexcept;
    OpcPackage& operator=(OpcPackage&&) noexcept;
    
    void open(const std::string& path);
    void close();
    bool isOpen() const;
    
    std::string findWorkbookPath() const;
    std::vector<std::string> getContentTypes() const;
    
    // Access to the underlying ZipReader for advanced operations
    const ZipReader& getZipReader() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

// Date system enumeration for Excel workbooks
enum class DateSystem {
    Date1900 = 0,  // Default Excel date system (Windows)
    Date1904 = 1   // Mac Excel date system
};

// Sheet information structure
struct SheetInfo {
    std::string name;           // Sheet name
    std::string relationshipId; // Relationship ID (r:id)
    std::string target;         // Target path (e.g., "worksheets/sheet1.xml")
    int sheetId;                // Sheet ID number
    bool visible;               // Sheet visibility
};

// Workbook properties
struct WorkbookProperties {
    DateSystem dateSystem = DateSystem::Date1900;
    std::string appName;        // Application name that created the workbook
    std::string lastModifiedBy; // Last modified by user
};

class Workbook {
public:
    Workbook();
    ~Workbook();
    
    // Non-copyable but movable
    Workbook(const Workbook&) = delete;
    Workbook& operator=(const Workbook&) = delete;
    Workbook(Workbook&&) noexcept;
    Workbook& operator=(Workbook&&) noexcept;
    
    void open(const OpcPackage& package);
    void close();
    bool isOpen() const;
    
    // Sheet information
    std::vector<SheetInfo> getSheets() const;
    std::optional<SheetInfo> findSheet(const std::string& name) const;
    std::optional<SheetInfo> findSheet(int index) const;
    size_t getSheetCount() const;
    
    // Workbook properties
    const WorkbookProperties& getProperties() const;
    DateSystem getDateSystem() const;
    
    // Relationship mapping
    std::string resolveRelationshipTarget(const std::string& relationshipId) const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

// Excel number format types
enum class NumberFormatType {
    General = 0,
    Integer = 1,
    Decimal = 2,
    Currency = 3,
    Accounting = 4,
    Date = 5,
    Time = 6,
    DateTime = 7,
    Percentage = 8,
    Fraction = 9,
    Scientific = 10,
    Text = 11,
    Custom = 12
};

// Font information structure
struct FontInfo {
    std::string name = "Calibri";
    double size = 11.0;
    bool bold = false;
    bool italic = false;
    bool underline = false;
    std::string color; // RGB hex string (e.g., "FF000000")
};

// Fill/background information structure
struct FillInfo {
    std::string patternType = "none";
    std::string fgColor; // Foreground color (RGB hex)
    std::string bgColor; // Background color (RGB hex)
};

// Border information structure
struct BorderInfo {
    std::string left;
    std::string right;
    std::string top;
    std::string bottom;
    std::string diagonal;
    std::string leftColor;
    std::string rightColor;
    std::string topColor;
    std::string bottomColor;
    std::string diagonalColor;
};

// Number format information
struct NumberFormat {
    int formatId = 0;
    std::string formatCode;
    NumberFormatType type = NumberFormatType::General;
    bool isBuiltIn = false;
};

// Complete cell style information
struct CellStyle {
    int styleIndex = 0;
    NumberFormat numberFormat;
    FontInfo font;
    FillInfo fill;
    BorderInfo border;
    int alignment = 0; // For future use
};

class StylesRegistry {
public:
    StylesRegistry();
    ~StylesRegistry();
    
    // Non-copyable but movable
    StylesRegistry(const StylesRegistry&) = delete;
    StylesRegistry& operator=(const StylesRegistry&) = delete;
    StylesRegistry(StylesRegistry&&) noexcept;
    StylesRegistry& operator=(StylesRegistry&&) noexcept;
    
    void parse(const OpcPackage& package);
    bool isOpen() const;
    void close();
    
    // Style lookup methods
    std::optional<CellStyle> getCellStyle(int styleIndex) const;
    std::optional<NumberFormat> getNumberFormat(int formatId) const;
    NumberFormatType detectNumberFormatType(const std::string& formatCode) const;
    
    // Utility methods
    bool isDateTimeFormat(int formatId) const;
    bool isDateTimeFormat(const std::string& formatCode) const;
    size_t getStyleCount() const;
    size_t getNumberFormatCount() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

// Shared strings storage mode
enum class SharedStringsMode {
    Auto = 0,      // Automatically choose based on size thresholds
    InMemory = 1,  // Force in-memory storage
    External = 2   // Force external (spill-to-disk) storage
};

// Configuration for SharedStringsProvider
struct SharedStringsConfig {
    SharedStringsMode mode = SharedStringsMode::Auto;
    size_t memoryThreshold = 50 * 1024 * 1024; // 50MB threshold for spill-to-disk
    size_t maxStringLength = 32767; // Excel's maximum string length
    bool flattenRichText = true;    // Flatten rich text runs to plain text
};

class SharedStringsProvider {
public:
    SharedStringsProvider();
    explicit SharedStringsProvider(const SharedStringsConfig& config);
    ~SharedStringsProvider();
    
    // Non-copyable but movable
    SharedStringsProvider(const SharedStringsProvider&) = delete;
    SharedStringsProvider& operator=(const SharedStringsProvider&) = delete;
    SharedStringsProvider(SharedStringsProvider&&) noexcept;
    SharedStringsProvider& operator=(SharedStringsProvider&&) noexcept;
    
    void parse(const OpcPackage& package);
    void close();
    bool isOpen() const;
    
    // String lookup methods
    std::string getString(size_t index) const;
    std::optional<std::string> tryGetString(size_t index) const;
    size_t getStringCount() const;
    bool hasStrings() const;
    
    // Configuration and statistics
    const SharedStringsConfig& getConfig() const;
    SharedStringsMode getActiveMode() const;
    size_t getMemoryUsage() const;
    bool isUsingDisk() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

// Cell coordinate representation
struct CellCoordinate {
    int row = 0;        // 1-based row number
    int column = 0;     // 1-based column number (A=1, B=2, etc.)
    
    // Parse from Excel reference (e.g., "A1", "BC42")
    static std::optional<CellCoordinate> fromReference(const std::string& ref);
    
    // Convert to Excel reference (e.g., "A1", "BC42")
    std::string toReference() const;
};

// Excel cell types
enum class CellType {
    Unknown = 0,
    Boolean,        // b
    Error,          // e  
    InlineString,   // inlineStr
    Number,         // n (default)
    SharedString,   // s
    String          // str (formula result)
};

// Cell value variant
using CellValue = std::variant<
    std::monostate,    // Empty cell
    bool,              // Boolean value
    double,            // Numeric value  
    std::string,       // String value (resolved or inline)
    int                // Shared string index (before resolution)
>;

// Parsed cell data
struct CellData {
    CellCoordinate coordinate;
    CellType type = CellType::Unknown;
    CellValue value;
    int styleIndex = 0;        // Style reference for date/formatting
    
    // Helper methods
    bool isEmpty() const { return std::holds_alternative<std::monostate>(value); }
    bool isBoolean() const { return std::holds_alternative<bool>(value); }
    bool isNumber() const { return std::holds_alternative<double>(value); }
    bool isString() const { return std::holds_alternative<std::string>(value); }
    bool isSharedStringIndex() const { return std::holds_alternative<int>(value) && type == CellType::SharedString; }
    
    std::string getString() const;
    double getNumber() const;
    bool getBoolean() const;
    int getSharedStringIndex() const;
};

// Row of cells - sparse representation
struct RowData {
    int rowNumber = 0;
    std::vector<CellData> cells;
    bool hidden = false;               // Row visibility (true = hidden)
    
    // Find cell by column number (1-based)
    const CellData* findCell(int column) const;
    CellData* findCell(int column);
};

// Merged cell range (e.g., A1:C3)
struct MergedCellRange {
    CellCoordinate topLeft;           // Top-left cell (e.g., A1)
    CellCoordinate bottomRight;       // Bottom-right cell (e.g., C3)
    
    // Parse from Excel range format (e.g., "A1:C3")
    static std::optional<MergedCellRange> fromReference(const std::string& ref);
    
    // Convert to Excel range format (e.g., "A1:C3")
    std::string toReference() const;
    
    // Check if a coordinate is within this range
    bool contains(const CellCoordinate& coord) const;
    
    // Get all coordinates covered by this range
    std::vector<CellCoordinate> getAllCoordinates() const;
};

// Column information
struct ColumnInfo {
    int columnIndex = 0;              // 1-based column index
    bool hidden = false;              // Column visibility (true = hidden)
    double width = 0.0;               // Column width (optional)
};

// Worksheet metadata including merged cells and hidden elements
struct WorksheetMetadata {
    std::vector<MergedCellRange> mergedCells;
    std::vector<ColumnInfo> columnInfo;
    
    // Check if a coordinate is part of any merged cell
    const MergedCellRange* findMergedCellRange(const CellCoordinate& coord) const;
    
    // Check if a column is hidden
    bool isColumnHidden(int column) const;
};

// Callback interface for row-by-row processing
class SheetRowHandler {
public:
    virtual ~SheetRowHandler() = default;
    virtual void handleRow(const RowData& row) = 0;
    virtual void handleError(const std::string& message) = 0;
    virtual void handleWorksheetMetadata([[maybe_unused]] const WorksheetMetadata& metadata) {}  // Optional
};

// Sheet streaming parser
class SheetStreamReader {
public:
    SheetStreamReader();
    ~SheetStreamReader();
    
    // Non-copyable but movable
    SheetStreamReader(const SheetStreamReader&) = delete;
    SheetStreamReader& operator=(const SheetStreamReader&) = delete;
    SheetStreamReader(SheetStreamReader&&) noexcept;
    SheetStreamReader& operator=(SheetStreamReader&&) noexcept;
    
    // Parse worksheet XML from OPC package
    void parseSheet(const OpcPackage& package, 
                    const std::string& sheetPath,
                    SheetRowHandler& handler,
                    const SharedStringsProvider* sharedStrings = nullptr,
                    const StylesRegistry* styles = nullptr);
    
    // Parse worksheet XML from raw data
    void parseSheetData(const std::vector<uint8_t>& xmlData,
                        SheetRowHandler& handler,
                        const SharedStringsProvider* sharedStrings = nullptr,
                        const StylesRegistry* styles = nullptr);

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

// Forward declarations for data conversion
class DateConverter;
class DataConverter;

// Forward declare CsvOptions as opaque pointer
// We'll use void* and cast appropriately in the implementation
class CsvOptions; // Forward declaration

// CSV Row Handler that collects data into CSV format
class CsvRowCollector : public SheetRowHandler {
public:
    explicit CsvRowCollector(const SharedStringsProvider* sharedStrings = nullptr,
                           const StylesRegistry* styles = nullptr,
                           DateSystem dateSystem = DateSystem::Date1900,
                           const void* csvOptions = nullptr);
    ~CsvRowCollector();
    
    // SheetRowHandler interface
    void handleRow(const RowData& row) override;
    void handleError(const std::string& message) override;
    void handleWorksheetMetadata(const WorksheetMetadata& metadata) override;
    
    // Get results
    std::string getCsvString() const;
    const std::vector<std::string>& getErrors() const;
    size_t getRowCount() const;

private:
    class Impl;
    std::unique_ptr<Impl> m_impl;
};

} // namespace xlsxcsv::core
