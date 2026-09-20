#include "xlsxcsv.hpp"
#include "xlsxcsv/core.hpp"
#include <stdexcept>
#include <sstream>
#include <algorithm>
#include <chrono>
#include <cstdlib>
#include <iostream>
#include <filesystem>
#include <fstream>
#include <random>
#include <system_error>
#ifdef _WIN32
#include <windows.h>
#endif

namespace xlsxcsv {

namespace {

std::filesystem::path normalizedPath(const std::filesystem::path& path) {
    std::error_code error;
    auto normalized = std::filesystem::weakly_canonical(path, error);
    return error ? std::filesystem::absolute(path).lexically_normal() : normalized;
}

std::filesystem::path temporaryOutputPath(const std::filesystem::path& outputPath) {
    std::random_device random;
    for (int attempt = 0; attempt < 32; ++attempt) {
        const auto suffix = std::to_string(random()) + "-" + std::to_string(attempt);
        auto temporaryName = outputPath.filename();
        temporaryName += std::filesystem::path(".turboxl-" + suffix + ".tmp");
        auto candidate = outputPath.parent_path() / temporaryName;
        if (!std::filesystem::exists(candidate)) return candidate;
    }
    throw std::runtime_error("Unable to allocate a temporary CSV output path");
}

void atomicReplace(const std::filesystem::path& source,
                   const std::filesystem::path& destination) {
#ifdef _WIN32
    if (!MoveFileExW(source.c_str(), destination.c_str(),
                     MOVEFILE_REPLACE_EXISTING | MOVEFILE_WRITE_THROUGH)) {
        throw std::system_error(static_cast<int>(GetLastError()),
            std::system_category(), "Failed to replace CSV output");
    }
#else
    std::error_code error;
    std::filesystem::rename(source, destination, error);
    if (error) throw std::system_error(error, "Failed to replace CSV output");
#endif
}

std::optional<xlsxcsv::core::SheetInfo> selectSheet(
    const xlsxcsv::core::Workbook& workbook,
    const std::variant<std::string, int>& selector) {
    if (std::holds_alternative<std::string>(selector)) {
        return workbook.findSheet(std::get<std::string>(selector));
    }
    const int index = std::get<int>(selector);
    return workbook.findSheet(index == -1 ? 0 : index);
}

} // namespace

std::string readSheetToCsv(
    const std::string& xlsxPath,
    const std::variant<std::string, int>& sheetSelector,
    const CsvOptions& options) {
    
    try {
        const bool profileTimings = []() {
            const char* v = std::getenv("TURBOXL_PROFILE_TIMINGS");
            return v && (v[0] == '1' || v[0] == 't' || v[0] == 'T' || v[0] == 'y' || v[0] == 'Y');
        }();
        const auto t0 = std::chrono::steady_clock::now();
        auto msSince = [&](const std::chrono::steady_clock::time_point& start) -> double {
            return std::chrono::duration<double, std::milli>(std::chrono::steady_clock::now() - start).count();
        };
        double t_open = 0.0;
        double t_workbook = 0.0;
        double t_styles = 0.0;
        double t_shared = 0.0;
        double t_sheet = 0.0;
        double t_csv = 0.0;
        double t_post = 0.0;

        // Create security limits from options
        xlsxcsv::core::ZipSecurityLimits limits;
        limits.maxEntries = options.maxEntries;
        limits.maxEntrySize = options.maxEntrySize;
        limits.maxTotalUncompressed = options.maxTotalUncompressed;
        
        // Phase 4 implementation: Shared strings + all previous phases
        xlsxcsv::core::OpcPackage package(limits);
        auto t = std::chrono::steady_clock::now();
        package.open(xlsxPath);
        t_open = msSince(t);
        
        // Parse workbook structure
        xlsxcsv::core::Workbook workbook;
        t = std::chrono::steady_clock::now();
        workbook.open(package);
        t_workbook = msSince(t);
        
        // Parse styles registry
        xlsxcsv::core::StylesRegistry styles;
        t = std::chrono::steady_clock::now();
        try {
            styles.parse(package, xlsxcsv::core::StylesRegistry::ParseMode::CsvOnly);
        } catch (const xlsxcsv::core::XlsxError& e) {
            // Some XLSX files might not have styles.xml, continue without styles
        }
        t_styles = msSince(t);
        
        // Parse shared strings (Phase 4)
        xlsxcsv::core::SharedStringsConfig sharedConfig;
        sharedConfig.mode = options.sharedStringsMode == CsvOptions::SharedStringsMode::AUTO ? 
            xlsxcsv::core::SharedStringsMode::Auto :
            options.sharedStringsMode == CsvOptions::SharedStringsMode::IN_MEMORY ?
            xlsxcsv::core::SharedStringsMode::InMemory :
            xlsxcsv::core::SharedStringsMode::External;
        
        xlsxcsv::core::SharedStringsProvider sharedStrings(sharedConfig);
        t = std::chrono::steady_clock::now();
        try {
            sharedStrings.parse(package);
        } catch (const xlsxcsv::core::XlsxError& e) {
            // Some XLSX files might not have sharedStrings.xml, continue without shared strings
        }
        t_shared = msSince(t);
        
        const auto targetSheet = selectSheet(workbook, sheetSelector);
        
        if (!targetSheet.has_value()) {
            if (std::holds_alternative<std::string>(sheetSelector)) {
                throw std::runtime_error(
                    "Sheet not found: " + std::get<std::string>(sheetSelector));
            }
            throw std::runtime_error(
                "Sheet index out of range: " +
                std::to_string(std::get<int>(sheetSelector)));
        }
        
        // Phase 5: Parse sheet content to CSV
        xlsxcsv::core::SheetStreamReader sheetReader;
        
        // Create CSV collector with proper configuration
        xlsxcsv::core::CsvRowCollector csvCollector(
            sharedStrings.isOpen() ? &sharedStrings : nullptr,
            styles.isOpen() ? &styles : nullptr, 
            workbook.getDateSystem(),
            &options
        );
        
        // Parse the worksheet
        t = std::chrono::steady_clock::now();
        sheetReader.parseSheet(package, targetSheet->target, csvCollector,
                              sharedStrings.isOpen() ? &sharedStrings : nullptr,
                              styles.isOpen() ? &styles : nullptr);
        t_sheet = msSince(t);
        
        // Check for parsing errors
        const auto& errors = csvCollector.getErrors();
        if (!errors.empty()) {
            std::ostringstream errorMsg;
            errorMsg << "Sheet parsing errors: ";
            for (size_t i = 0; i < errors.size(); ++i) {
                if (i > 0) errorMsg << "; ";
                errorMsg << errors[i];
            }
            throw std::runtime_error(errorMsg.str());
        }
        
        // Return CSV string
        t = std::chrono::steady_clock::now();
        std::string csvResult = csvCollector.takeCsvString();
        t_csv = msSince(t);
        t_post = 0.0;

        if (profileTimings) {
            const double totalMs = msSince(t0);
            std::cerr
                << "turboxl_timing_ms"
                << " open=" << t_open
                << " workbook=" << t_workbook
                << " styles=" << t_styles
                << " shared_strings=" << t_shared
                << " parse_sheet=" << t_sheet
                << " assemble_csv=" << t_csv
                << " postprocess=" << t_post
                << " total=" << totalMs
                << " rows=" << csvCollector.getRowCount()
                << "\n";
        }
        
        return csvResult;
    }
    catch (const xlsxcsv::core::XlsxError& e) {
        throw std::runtime_error("XLSX parsing error: " + std::string(e.what()));
    }
    catch (const std::exception& e) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(e.what()));
    }
}

std::string readSheetToCsv(const std::string& xlsxPath) {
    return readSheetToCsv(xlsxPath, -1, CsvOptions{});
}

void readSheetToFile(
    const std::string& xlsxPath,
    const std::filesystem::path& outputPath,
    const std::variant<std::string, int>& sheetSelector,
    const CsvOptions& options) {
    if (outputPath.empty()) throw std::runtime_error("CSV output path is empty");
    if (normalizedPath(xlsxPath) == normalizedPath(outputPath)) {
        throw std::runtime_error("Input workbook and CSV output paths must differ");
    }
    const auto parent = outputPath.has_parent_path()
        ? outputPath.parent_path() : std::filesystem::current_path();
    if (!std::filesystem::is_directory(parent)) {
        throw std::runtime_error("CSV output directory does not exist: " + parent.string());
    }

    const auto temporaryPath = temporaryOutputPath(outputPath);
    bool committed = false;
    try {
        xlsxcsv::core::ZipSecurityLimits limits;
        limits.maxEntries = options.maxEntries;
        limits.maxEntrySize = options.maxEntrySize;
        limits.maxTotalUncompressed = options.maxTotalUncompressed;

        xlsxcsv::core::OpcPackage package(limits);
        package.open(xlsxPath);
        xlsxcsv::core::Workbook workbook;
        workbook.open(package);
        const auto targetSheet = selectSheet(workbook, sheetSelector);
        if (!targetSheet) throw std::runtime_error("Sheet not found");

        xlsxcsv::core::StylesRegistry styles;
        try {
            styles.parse(package, xlsxcsv::core::StylesRegistry::ParseMode::CsvOnly);
        } catch (const xlsxcsv::core::XlsxError&) {
        }

        xlsxcsv::core::SharedStringsConfig sharedConfig;
        sharedConfig.mode = options.sharedStringsMode == CsvOptions::SharedStringsMode::AUTO
            ? xlsxcsv::core::SharedStringsMode::Auto
            : options.sharedStringsMode == CsvOptions::SharedStringsMode::IN_MEMORY
                ? xlsxcsv::core::SharedStringsMode::InMemory
                : xlsxcsv::core::SharedStringsMode::External;
        xlsxcsv::core::SharedStringsProvider sharedStrings(sharedConfig);
        try {
            sharedStrings.parse(package);
        } catch (const xlsxcsv::core::XlsxError&) {
        }

        std::ofstream output(temporaryPath, std::ios::binary | std::ios::trunc);
        if (!output) throw std::runtime_error("Unable to create temporary CSV output");
        xlsxcsv::core::CsvRowCollector collector(
            sharedStrings.isOpen() ? &sharedStrings : nullptr,
            styles.isOpen() ? &styles : nullptr,
            workbook.getDateSystem(), &options, &output);
        xlsxcsv::core::SheetStreamReader reader;
        reader.parseSheet(package, targetSheet->target, collector,
            sharedStrings.isOpen() ? &sharedStrings : nullptr,
            styles.isOpen() ? &styles : nullptr);
        if (!collector.getErrors().empty()) {
            throw std::runtime_error("Sheet parsing failed: " + collector.getErrors().front());
        }
        collector.finalize();
        output.close();
        if (!output) throw std::runtime_error("Failed to close CSV output");
        atomicReplace(temporaryPath, outputPath);
        committed = true;
    } catch (const xlsxcsv::core::XlsxError& error) {
        std::error_code ignored;
        if (!committed) std::filesystem::remove(temporaryPath, ignored);
        throw std::runtime_error("XLSX parsing error: " + std::string(error.what()));
    } catch (...) {
        std::error_code ignored;
        if (!committed) std::filesystem::remove(temporaryPath, ignored);
        throw;
    }
}

std::vector<SheetMetadata> getSheetList(const std::string& xlsxPath) {
    try {
        // Create security limits with defaults
        xlsxcsv::core::ZipSecurityLimits limits;
        
        // Open package and workbook (lightweight operations)
        xlsxcsv::core::OpcPackage package(limits);
        package.open(xlsxPath);
        
        xlsxcsv::core::Workbook workbook;
        workbook.open(package);
        
        // Get all sheets and convert to public metadata format
        auto sheets = workbook.getSheets();
        std::vector<SheetMetadata> result;
        result.reserve(sheets.size());
        
        for (const auto& sheet : sheets) {
            SheetMetadata metadata;
            metadata.name = sheet.name;
            metadata.sheetId = sheet.sheetId;
            metadata.visible = sheet.visible;
            metadata.target = sheet.target;
            switch (sheet.kind) {
                case core::SheetKind::Worksheet:
                    metadata.kind = SheetKind::Worksheet;
                    break;
                case core::SheetKind::Chartsheet:
                    metadata.kind = SheetKind::Chartsheet;
                    break;
                case core::SheetKind::Other:
                    metadata.kind = SheetKind::Other;
                    break;
            }
            switch (sheet.visibility) {
                case core::SheetVisibility::Visible:
                    metadata.visibility = SheetVisibility::Visible;
                    break;
                case core::SheetVisibility::Hidden:
                    metadata.visibility = SheetVisibility::Hidden;
                    break;
                case core::SheetVisibility::VeryHidden:
                    metadata.visibility = SheetVisibility::VeryHidden;
                    break;
            }
            result.push_back(metadata);
        }
        
        return result;
    }
    catch (const xlsxcsv::core::XlsxError& e) {
        throw std::runtime_error("XLSX parsing error: " + std::string(e.what()));
    }
    catch (const std::exception& e) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(e.what()));
    }
}

std::vector<SheetMetadata> getVisibleSheets(const std::string& xlsxPath) {
    auto allSheets = getSheetList(xlsxPath);
    
    std::vector<SheetMetadata> visibleSheets;
    for (const auto& sheet : allSheets) {
        if (sheet.visible) {
            visibleSheets.push_back(sheet);
        }
    }
    
    return visibleSheets;
}

std::string readSpecificSheet(
    const std::string& xlsxPath,
    const std::string& sheetName,
    const CsvOptions& options) {
    
    // Use the existing function but with the specific sheet name
    CsvOptions modifiedOptions = options;
    modifiedOptions.sheetByName = sheetName;
    modifiedOptions.sheetByIndex = -1; // Clear index to ensure name takes precedence
    
    return readSheetToCsv(xlsxPath, sheetName, modifiedOptions);
}

std::map<std::string, std::string> readMultipleSheets(
    const std::string& xlsxPath,
    const std::vector<std::string>& sheetNames,
    const CsvOptions& options) {
    
    try {
        // Create security limits from options
        xlsxcsv::core::ZipSecurityLimits limits;
        limits.maxEntries = options.maxEntries;
        limits.maxEntrySize = options.maxEntrySize;
        limits.maxTotalUncompressed = options.maxTotalUncompressed;
        
        // Open package, workbook, styles, and shared strings once (efficient reuse)
        xlsxcsv::core::OpcPackage package(limits);
        package.open(xlsxPath);
        
        xlsxcsv::core::Workbook workbook;
        workbook.open(package);
        
        xlsxcsv::core::StylesRegistry styles;
        try {
            styles.parse(package, xlsxcsv::core::StylesRegistry::ParseMode::CsvOnly);
        } catch (const xlsxcsv::core::XlsxError& e) {
            // Some XLSX files might not have styles.xml, continue without styles
        }
        
        xlsxcsv::core::SharedStringsConfig sharedConfig;
        sharedConfig.mode = options.sharedStringsMode == CsvOptions::SharedStringsMode::AUTO ? 
            xlsxcsv::core::SharedStringsMode::Auto :
            options.sharedStringsMode == CsvOptions::SharedStringsMode::IN_MEMORY ?
            xlsxcsv::core::SharedStringsMode::InMemory :
            xlsxcsv::core::SharedStringsMode::External;
        
        xlsxcsv::core::SharedStringsProvider sharedStrings(sharedConfig);
        try {
            sharedStrings.parse(package);
        } catch (const xlsxcsv::core::XlsxError& e) {
            // Some XLSX files might not have sharedStrings.xml, continue without shared strings
        }
        
        std::map<std::string, std::string> results;
        
        // Process each requested sheet
        xlsxcsv::core::SheetStreamReader sheetReader;
        
        for (const std::string& sheetName : sheetNames) {
            auto sheetInfo = workbook.findSheet(sheetName);
            if (!sheetInfo.has_value()) {
                throw std::runtime_error("Sheet not found: " + sheetName);
            }
            
            // Parse sheet content to CSV
            xlsxcsv::core::CsvRowCollector csvCollector(
                sharedStrings.isOpen() ? &sharedStrings : nullptr,
                styles.isOpen() ? &styles : nullptr, 
                workbook.getDateSystem(),
                &options
            );
            
            // Parse the worksheet
            sheetReader.parseSheet(package, sheetInfo->target, csvCollector,
                                  sharedStrings.isOpen() ? &sharedStrings : nullptr,
                                  styles.isOpen() ? &styles : nullptr);
            
            // Check for parsing errors
            const auto& errors = csvCollector.getErrors();
            if (!errors.empty()) {
                std::ostringstream errorMsg;
                errorMsg << "Sheet parsing errors for '" << sheetName << "': ";
                for (size_t i = 0; i < errors.size(); ++i) {
                    if (i > 0) errorMsg << "; ";
                    errorMsg << errors[i];
                }
                throw std::runtime_error(errorMsg.str());
            }
            
            // Get CSV result
            std::string csvResult = csvCollector.takeCsvString();
            
            results[sheetName] = csvResult;
        }
        
        return results;
    }
    catch (const xlsxcsv::core::XlsxError& e) {
        throw std::runtime_error("XLSX parsing error: " + std::string(e.what()));
    }
    catch (const std::exception& e) {
        throw std::runtime_error("Error reading XLSX file: " + std::string(e.what()));
    }
}

} // namespace xlsxcsv
