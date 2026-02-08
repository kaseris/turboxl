#include "xlsxcsv.hpp"
#include "xlsxcsv/core.hpp"
#include <stdexcept>
#include <sstream>
#include <algorithm>
#include <chrono>
#include <cstdlib>
#include <iostream>

namespace xlsxcsv {

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
        xlsxcsv::core::OpcPackage package;
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
            styles.parse(package);
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
        
        // Determine which sheet to parse
        std::optional<xlsxcsv::core::SheetInfo> targetSheet;
        auto sheets = workbook.getSheets();
        
        if (std::holds_alternative<std::string>(sheetSelector)) {
            // Find sheet by name
            std::string sheetName = std::get<std::string>(sheetSelector);
            targetSheet = workbook.findSheet(sheetName);
            if (!targetSheet.has_value()) {
                throw std::runtime_error("Sheet not found: " + sheetName);
            }
        } else {
            // Find sheet by index
            int sheetIndex = std::get<int>(sheetSelector);
            if (sheetIndex == -1) {
                // Use first sheet
                if (!sheets.empty()) {
                    targetSheet = sheets[0];
                }
            } else if (sheetIndex >= 0 && static_cast<size_t>(sheetIndex) < sheets.size()) {
                targetSheet = sheets[static_cast<size_t>(sheetIndex)];
            }
            
            if (!targetSheet.has_value()) {
                throw std::runtime_error("Sheet index out of range: " + std::to_string(sheetIndex));
            }
        }
        
        if (!targetSheet.has_value()) {
            throw std::runtime_error("No sheets found in workbook");
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
        std::string csvResult = csvCollector.getCsvString();
        t_csv = msSince(t);
        
        // Handle BOM if requested
        if (options.includeBom) {
            csvResult = "\xEF\xBB\xBF" + csvResult;
        }
        
        // Handle newline conversion if needed
        if (options.newline == CsvOptions::Newline::CRLF) {
            // Convert LF to CRLF
            std::string result;
            result.reserve(csvResult.size() * 1.1); // Reserve some extra space
            for (char c : csvResult) {
                if (c == '\n') {
                    result += "\r\n";
                } else {
                    result += c;
                }
            }
            csvResult = std::move(result);
        }
        t_post = msSince(t);

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

std::vector<SheetMetadata> getSheetList(const std::string& xlsxPath) {
    try {
        // Create security limits with defaults
        xlsxcsv::core::ZipSecurityLimits limits;
        
        // Open package and workbook (lightweight operations)
        xlsxcsv::core::OpcPackage package;
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
        xlsxcsv::core::OpcPackage package;
        package.open(xlsxPath);
        
        xlsxcsv::core::Workbook workbook;
        workbook.open(package);
        
        xlsxcsv::core::StylesRegistry styles;
        try {
            styles.parse(package);
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
            std::string csvResult = csvCollector.getCsvString();
            
            // Handle BOM if requested
            if (options.includeBom) {
                csvResult = "\xEF\xBB\xBF" + csvResult;
            }
            
            // Handle newline conversion if needed
            if (options.newline == CsvOptions::Newline::CRLF) {
                // Convert LF to CRLF
                std::string result;
                result.reserve(csvResult.size() * 1.1);
                for (char c : csvResult) {
                    if (c == '\n') {
                        result += "\r\n";
                    } else {
                        result += c;
                    }
                }
                csvResult = std::move(result);
            }
            
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
