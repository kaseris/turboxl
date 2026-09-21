#include <nanobind/nanobind.h>
#include <nanobind/stl/map.h>
#include <nanobind/stl/optional.h>
#include <nanobind/stl/string.h>
#include <nanobind/stl/variant.h>
#include <nanobind/stl/vector.h>
#include <nanobind/stl/filesystem.h>
#include "xlsxcsv.hpp"
#include "typed_reader.hpp"

#include <chrono>
#include <cstdint>
#include <cstdlib>
#include <iostream>

namespace nb = nanobind;

namespace {

nb::object boxTypedValue(const xlsxcsv::internal::TypedCellValue& value) {
    if (std::holds_alternative<std::monostate>(value)) {
        return nb::none();
    }
    if (std::holds_alternative<bool>(value)) {
        return nb::bool_(std::get<bool>(value));
    }
    if (std::holds_alternative<double>(value)) {
        return nb::float_(std::get<double>(value));
    }
    const auto& text = std::get<std::string>(value);
    return nb::str(text.data(), text.size());
}

bool profileTypedTimings() {
    const char* value = std::getenv("TURBOXL_PROFILE_TYPED_TIMINGS");
    return value && (value[0] == '1' || value[0] == 't' || value[0] == 'T' ||
                     value[0] == 'y' || value[0] == 'Y');
}

} // namespace

NB_MODULE(_turboxl, m) {
    m.doc() = "Fast XLSX to CSV converter (C++ core with Python bindings)";
    
    // Enums
    nb::enum_<xlsxcsv::CsvOptions::Newline>(m, "Newline")
        .value("LF", xlsxcsv::CsvOptions::Newline::LF)
        .value("CRLF", xlsxcsv::CsvOptions::Newline::CRLF);
    
    nb::enum_<xlsxcsv::CsvOptions::DateMode>(m, "DateMode")
        .value("ISO", xlsxcsv::CsvOptions::DateMode::ISO)
        .value("RAW", xlsxcsv::CsvOptions::DateMode::RAW);
    
    nb::enum_<xlsxcsv::CsvOptions::SharedStringsMode>(m, "SharedStringsMode")
        .value("AUTO", xlsxcsv::CsvOptions::SharedStringsMode::AUTO)
        .value("IN_MEMORY", xlsxcsv::CsvOptions::SharedStringsMode::IN_MEMORY)
        .value("EXTERNAL", xlsxcsv::CsvOptions::SharedStringsMode::EXTERNAL);
    
    nb::enum_<xlsxcsv::CsvOptions::MergedHandling>(m, "MergedHandling")
        .value("NONE", xlsxcsv::CsvOptions::MergedHandling::NONE)
        .value("PROPAGATE", xlsxcsv::CsvOptions::MergedHandling::PROPAGATE);

    nb::enum_<xlsxcsv::SheetKind>(m, "SheetKind")
        .value("WORKSHEET", xlsxcsv::SheetKind::Worksheet)
        .value("CHARTSHEET", xlsxcsv::SheetKind::Chartsheet)
        .value("OTHER", xlsxcsv::SheetKind::Other);

    nb::enum_<xlsxcsv::SheetVisibility>(m, "SheetVisibility")
        .value("VISIBLE", xlsxcsv::SheetVisibility::Visible)
        .value("HIDDEN", xlsxcsv::SheetVisibility::Hidden)
        .value("VERY_HIDDEN", xlsxcsv::SheetVisibility::VeryHidden);
    
    // SheetMetadata struct
    nb::class_<xlsxcsv::SheetMetadata>(m, "SheetMetadata")
        .def(nb::init<>())
        .def_rw("name", &xlsxcsv::SheetMetadata::name)
        .def_rw("sheet_id", &xlsxcsv::SheetMetadata::sheetId)
        .def_rw("visible", &xlsxcsv::SheetMetadata::visible)
        .def_rw("target", &xlsxcsv::SheetMetadata::target)
        .def_rw("kind", &xlsxcsv::SheetMetadata::kind)
        .def_rw("visibility", &xlsxcsv::SheetMetadata::visibility)
        .def("__repr__", [](const xlsxcsv::SheetMetadata &s) {
            return "SheetMetadata(name='" + s.name + "', sheet_id=" + std::to_string(s.sheetId) + 
                   ", visible=" + (s.visible ? "True" : "False") + ")";
        });
    
    // CsvOptions struct
    nb::class_<xlsxcsv::CsvOptions>(m, "CsvOptions")
        .def(nb::init<>())
        .def_rw("sheet_by_name", &xlsxcsv::CsvOptions::sheetByName)
        .def_rw("sheet_by_index", &xlsxcsv::CsvOptions::sheetByIndex)
        .def_rw("delimiter", &xlsxcsv::CsvOptions::delimiter)
        .def_rw("newline", &xlsxcsv::CsvOptions::newline)
        .def_rw("include_bom", &xlsxcsv::CsvOptions::includeBom)
        .def_rw("date_mode", &xlsxcsv::CsvOptions::dateMode)
        .def_rw("quote_all", &xlsxcsv::CsvOptions::quoteAll)
        .def_rw("shared_strings_mode", &xlsxcsv::CsvOptions::sharedStringsMode)
        .def_rw("merged_handling", &xlsxcsv::CsvOptions::mergedHandling)
        .def_rw("include_hidden_rows", &xlsxcsv::CsvOptions::includeHiddenRows)
        .def_rw("include_hidden_columns", &xlsxcsv::CsvOptions::includeHiddenColumns)
        .def_rw("max_entries", &xlsxcsv::CsvOptions::maxEntries)
        .def_rw("max_entry_size", &xlsxcsv::CsvOptions::maxEntrySize)
        .def_rw("max_total_uncompressed", &xlsxcsv::CsvOptions::maxTotalUncompressed);
    
    // Main function
    m.def("read_sheet_to_csv", 
        [](const std::string& xlsx_path, 
           const std::variant<std::string, int>& sheet,
           const xlsxcsv::CsvOptions& options) -> std::string {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSheetToCsv(xlsx_path, sheet, options);
        },
        nb::arg("xlsx_path"),
        nb::arg("sheet") = -1,
        nb::arg("options") = xlsxcsv::CsvOptions{},
        "Convert a worksheet from XLSX to CSV string"
    );

    m.def("read_sheet_to_file",
        [](const std::string& xlsx_path,
           const std::filesystem::path& output_path,
           const std::variant<std::string, int>& sheet,
           const xlsxcsv::CsvOptions& options) {
            nb::gil_scoped_release gil;
            xlsxcsv::readSheetToFile(xlsx_path, output_path, sheet, options);
        },
        nb::arg("xlsx_path"),
        nb::arg("output_path"),
        nb::arg("sheet") = -1,
        nb::arg("options") = xlsxcsv::CsvOptions{},
        "Convert a worksheet from XLSX directly to a CSV file"
    );
    
    // Convenience function
    m.def("read_sheet_to_csv", 
        [](const std::string& xlsx_path) -> std::string {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSheetToCsv(xlsx_path);
        },
        nb::arg("xlsx_path"),
        "Convert the first worksheet from XLSX to CSV string"
    );
    
    // Sheet discovery functions
    m.def("get_sheet_list", 
        [](const std::string& xlsx_path) -> std::vector<xlsxcsv::SheetMetadata> {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::getSheetList(xlsx_path);
        },
        nb::arg("xlsx_path"),
        "Get metadata for all sheets in an XLSX file without reading sheet content"
    );
    
    m.def("get_visible_sheets", 
        [](const std::string& xlsx_path) -> std::vector<xlsxcsv::SheetMetadata> {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::getVisibleSheets(xlsx_path);
        },
        nb::arg("xlsx_path"),
        "Get metadata for only visible sheets in an XLSX file"
    );
    
    // Selective parsing functions
    m.def("read_specific_sheet", 
        [](const std::string& xlsx_path, 
           const std::string& sheet_name,
           const xlsxcsv::CsvOptions& options) -> std::string {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSpecificSheet(xlsx_path, sheet_name, options);
        },
        nb::arg("xlsx_path"),
        nb::arg("sheet_name"),
        nb::arg("options") = xlsxcsv::CsvOptions{},
        "Convert a specific worksheet to CSV by name"
    );
    
    m.def("read_multiple_sheets", 
        [](const std::string& xlsx_path, 
           const std::vector<std::string>& sheet_names,
           const xlsxcsv::CsvOptions& options) -> std::map<std::string, std::string> {
            nb::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readMultipleSheets(xlsx_path, sheet_names, options);
        },
        nb::arg("xlsx_path"),
        nb::arg("sheet_names"),
        nb::arg("options") = xlsxcsv::CsvOptions{},
        "Convert multiple worksheets to CSV by name"
    );

    // Private vertical-slice API used only by the typed-path benchmark. The
    // public owning Workbook/Sheet API is tracked separately.
    m.def("_read_sheet_to_python",
        [](const std::string& xlsx_path,
           const std::variant<std::string, int>& sheet,
           bool skip_empty_area,
           const std::optional<std::int64_t>& nrows,
           std::int64_t max_cells) -> nb::list {
            if (nrows && *nrows < 0) {
                throw nb::value_error("nrows must be non-negative or None");
            }
            if (max_cells <= 0) {
                throw nb::value_error("max_cells must be greater than zero");
            }
            xlsxcsv::internal::TypedReadOptions options;
            options.skipEmptyArea = skip_empty_area;
            if (nrows) options.nrows = static_cast<std::size_t>(*nrows);
            options.maxCells = static_cast<std::size_t>(max_cells);

            using Clock = std::chrono::steady_clock;
            const auto totalStart = Clock::now();
            xlsxcsv::internal::TypedWorksheet nativeRows;
            const auto nativeStart = Clock::now();
            {
                nb::gil_scoped_release gil;
                nativeRows = xlsxcsv::internal::readSheetToTyped(
                    xlsx_path, sheet, options);
            }
            const auto nativeEnd = Clock::now();

            const auto boxingStart = Clock::now();
            nb::list rows;
            for (const auto& nativeRow : nativeRows) {
                nb::list row;
                for (const auto& value : nativeRow) {
                    row.append(boxTypedValue(value));
                }
                rows.append(std::move(row));
            }
            const auto boxingEnd = Clock::now();

            if (profileTypedTimings()) {
                const auto milliseconds = [](auto duration) {
                    return std::chrono::duration<double, std::milli>(duration).count();
                };
                const std::size_t columns = nativeRows.empty() ? 0 : nativeRows.front().size();
                std::cerr
                    << "turboxl_typed_timing_ms"
                    << " native=" << milliseconds(nativeEnd - nativeStart)
                    << " boxing=" << milliseconds(boxingEnd - boxingStart)
                    << " total=" << milliseconds(boxingEnd - totalStart)
                    << " rows=" << nativeRows.size()
                    << " columns=" << columns << '\n';
            }
            return rows;
        },
        nb::arg("xlsx_path"),
        nb::arg("sheet") = 0,
        nb::kw_only(),
        nb::arg("skip_empty_area") = false,
        nb::arg("nrows") = nb::none(),
        nb::arg("max_cells") = 10'000'000,
        "Private benchmark-only bounded typed worksheet extraction path"
    );
}
