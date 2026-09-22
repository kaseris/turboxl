#include "typed_reader.hpp"
#include "core/fast_typed_sheet_reader.hpp"

#include <optional>
#include <sstream>
#include <stdexcept>

namespace xlsxcsv::internal {

namespace {

std::string joinErrors(const std::vector<std::string>& errors) {
    std::ostringstream message;
    for (std::size_t index = 0; index < errors.size(); ++index) {
        if (index != 0) {
            message << "; ";
        }
        message << errors[index];
    }
    return message.str();
}

} // namespace

TypedWorkbookSession::TypedWorkbookSession(const std::string& path, std::size_t maxCells)
    : m_maxCells(maxCells) {
    if (maxCells == 0) throw std::invalid_argument("max_cells must be greater than zero");
    m_package.open(path);
    open();
}

TypedWorkbookSession::TypedWorkbookSession(core::ByteVector data, std::size_t maxCells)
    : m_maxCells(maxCells) {
    if (maxCells == 0) throw std::invalid_argument("max_cells must be greater than zero");
    m_package.open(std::move(data));
    open();
}

TypedWorkbookSession::~TypedWorkbookSession() { close(); }

void TypedWorkbookSession::open() {
    try {
        m_workbook.open(m_package);
        m_open = true;
    } catch (...) { m_package.close(); throw; }
}

bool TypedWorkbookSession::isOpen() const {
    std::lock_guard lock(m_mutex);
    return m_open;
}

void TypedWorkbookSession::close() {
    std::lock_guard lock(m_mutex);
    if (!m_open) return;
    m_styles.close();
    m_sharedStrings.close();
    m_workbook.close();
    m_package.close();
    m_open = false;
}

std::vector<core::SheetInfo> TypedWorkbookSession::sheets() const {
    std::lock_guard lock(m_mutex);
    if (!m_open) throw std::runtime_error("Workbook is closed");
    return m_workbook.getSheets();
}

std::optional<core::SheetInfo> TypedWorkbookSession::sheetByName(const std::string& name) const {
    std::lock_guard lock(m_mutex);
    if (!m_open) throw std::runtime_error("Workbook is closed");
    return m_workbook.findSheet(name);
}

std::optional<core::SheetInfo> TypedWorkbookSession::sheetByIndex(int index) const {
    std::lock_guard lock(m_mutex);
    if (!m_open) throw std::runtime_error("Workbook is closed");
    return m_workbook.findSheet(index);
}

void TypedWorkbookSession::ensureSharedStrings() {
    if (m_sharedStringsInitialized) return;
    try {
        if (m_package.getZipReader().hasEntry("xl/sharedStrings.xml")) {
            m_sharedStrings.parse(m_package);
        }
        m_sharedStringsInitialized = true;
    } catch (...) {
        // A failed lazy parse must not leave a partially initialized provider
        // behind or make a later read silently skip shared strings.
        m_sharedStrings.close();
        throw;
    }
}

void TypedWorkbookSession::ensureStyles() {
    if (m_stylesInitialized) return;
    try {
        if (m_package.getZipReader().hasEntry("xl/styles.xml")) {
            m_styles.parse(m_package, core::StylesRegistry::ParseMode::CsvOnly);
        }
        m_stylesInitialized = true;
    } catch (...) {
        // Keep retries deterministic after malformed optional style data.
        m_styles.close();
        throw;
    }
}

TypedWorksheet TypedWorkbookSession::read(const core::SheetInfo& sheet, TypedReadOptions options) {
    std::lock_guard lock(m_mutex);
    if (!m_open) throw std::runtime_error("Workbook is closed");
    // The limit belongs to the workbook session.  Python's public Sheet API
    // deliberately has no per-read override, so never fall back to a separate
    // TypedReadOptions default here.
    options.maxCells = m_maxCells;
    if (options.maxCells == 0) throw std::invalid_argument("max_cells must be greater than zero");
    if (options.nrows && *options.nrows == 0) return {};
    ensureSharedStrings();
    const auto* strings = m_sharedStrings.isOpen() ? &m_sharedStrings : nullptr;
    TypedRowCollector fast(strings, options, nullptr, m_workbook.getDateSystem());
    if (tryReadTypedWorksheetFast(m_package, sheet.target, fast) && fast.getErrors().empty()) {
        if (fast.hasStyledNumbers()) ensureStyles();
        if (m_styles.isOpen()) fast.setStyles(&m_styles);
        return fast.takeRows();
    }
    TypedRowCollector fallback(strings, options, nullptr, m_workbook.getDateSystem());
    core::SheetStreamReader reader;
    reader.parseSheet(m_package, sheet.target, fallback, strings);
    if (!fallback.getErrors().empty()) throw std::runtime_error("Sheet parsing errors: " + joinErrors(fallback.getErrors()));
    if (fallback.hasStyledNumbers()) ensureStyles();
    if (m_styles.isOpen()) fallback.setStyles(&m_styles);
    return fallback.takeRows();
}

TypedWorksheet readSheetToTyped(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector,
    const TypedReadOptions& options) {
    try {
        if (options.maxCells == 0) {
            throw std::invalid_argument("max_cells must be greater than zero");
        }
        TypedWorkbookSession session(xlsxPath, options.maxCells);
        const auto sheet = std::holds_alternative<std::string>(sheetSelector)
            ? session.sheetByName(std::get<std::string>(sheetSelector))
            : session.sheetByIndex(std::get<int>(sheetSelector) == -1 ? 0 : std::get<int>(sheetSelector));
        if (!sheet) {
            if (std::holds_alternative<std::string>(sheetSelector)) {
                throw std::runtime_error(
                    "Sheet not found: " + std::get<std::string>(sheetSelector));
            }
            throw std::runtime_error(
                "Sheet index out of range: " +
                std::to_string(std::get<int>(sheetSelector)));
        }
        return session.read(*sheet, options);
    } catch (const core::XlsxError& error) {
        throw std::runtime_error("XLSX parsing error: " + std::string(error.what()));
    } catch (const std::exception& error) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(error.what()));
    }
}

} // namespace xlsxcsv::internal
