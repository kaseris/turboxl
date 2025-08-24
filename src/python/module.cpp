#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include "xlsxcsv.hpp"

namespace py = pybind11;

PYBIND11_MODULE(turboxl, m) {
    m.doc() = "Fast XLSX to CSV converter (C++ core with Python bindings)";
    
    // Enums
    py::enum_<xlsxcsv::CsvOptions::Newline>(m, "Newline")
        .value("LF", xlsxcsv::CsvOptions::Newline::LF)
        .value("CRLF", xlsxcsv::CsvOptions::Newline::CRLF);
    
    py::enum_<xlsxcsv::CsvOptions::DateMode>(m, "DateMode")
        .value("ISO", xlsxcsv::CsvOptions::DateMode::ISO)
        .value("RAW", xlsxcsv::CsvOptions::DateMode::RAW);
    
    py::enum_<xlsxcsv::CsvOptions::SharedStringsMode>(m, "SharedStringsMode")
        .value("AUTO", xlsxcsv::CsvOptions::SharedStringsMode::AUTO)
        .value("IN_MEMORY", xlsxcsv::CsvOptions::SharedStringsMode::IN_MEMORY)
        .value("EXTERNAL", xlsxcsv::CsvOptions::SharedStringsMode::EXTERNAL);
    
    py::enum_<xlsxcsv::CsvOptions::MergedHandling>(m, "MergedHandling")
        .value("NONE", xlsxcsv::CsvOptions::MergedHandling::NONE)
        .value("PROPAGATE", xlsxcsv::CsvOptions::MergedHandling::PROPAGATE);
    
    // SheetMetadata struct
    py::class_<xlsxcsv::SheetMetadata>(m, "SheetMetadata")
        .def(py::init<>())
        .def_readwrite("name", &xlsxcsv::SheetMetadata::name)
        .def_readwrite("sheet_id", &xlsxcsv::SheetMetadata::sheetId)
        .def_readwrite("visible", &xlsxcsv::SheetMetadata::visible)
        .def_readwrite("target", &xlsxcsv::SheetMetadata::target)
        .def("__repr__", [](const xlsxcsv::SheetMetadata &s) {
            return "SheetMetadata(name='" + s.name + "', sheet_id=" + std::to_string(s.sheetId) + 
                   ", visible=" + (s.visible ? "True" : "False") + ")";
        });
    
    // CsvOptions struct
    py::class_<xlsxcsv::CsvOptions>(m, "CsvOptions")
        .def(py::init<>())
        .def_readwrite("sheet_by_name", &xlsxcsv::CsvOptions::sheetByName)
        .def_readwrite("sheet_by_index", &xlsxcsv::CsvOptions::sheetByIndex)
        .def_readwrite("delimiter", &xlsxcsv::CsvOptions::delimiter)
        .def_readwrite("newline", &xlsxcsv::CsvOptions::newline)
        .def_readwrite("include_bom", &xlsxcsv::CsvOptions::includeBom)
        .def_readwrite("date_mode", &xlsxcsv::CsvOptions::dateMode)
        .def_readwrite("quote_all", &xlsxcsv::CsvOptions::quoteAll)
        .def_readwrite("shared_strings_mode", &xlsxcsv::CsvOptions::sharedStringsMode)
        .def_readwrite("merged_handling", &xlsxcsv::CsvOptions::mergedHandling)
        .def_readwrite("include_hidden_rows", &xlsxcsv::CsvOptions::includeHiddenRows)
        .def_readwrite("include_hidden_columns", &xlsxcsv::CsvOptions::includeHiddenColumns)
        .def_readwrite("max_entries", &xlsxcsv::CsvOptions::maxEntries)
        .def_readwrite("max_entry_size", &xlsxcsv::CsvOptions::maxEntrySize)
        .def_readwrite("max_total_uncompressed", &xlsxcsv::CsvOptions::maxTotalUncompressed);
    
    // Main function
    m.def("read_sheet_to_csv", 
        [](const std::string& xlsx_path, 
           const std::variant<std::string, int>& sheet,
           const xlsxcsv::CsvOptions& options) -> std::string {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSheetToCsv(xlsx_path, sheet, options);
        },
        py::arg("xlsx_path"),
        py::arg("sheet") = -1,
        py::arg("options") = xlsxcsv::CsvOptions{},
        "Convert a worksheet from XLSX to CSV string"
    );
    
    // Convenience function
    m.def("read_sheet_to_csv", 
        [](const std::string& xlsx_path) -> std::string {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSheetToCsv(xlsx_path);
        },
        py::arg("xlsx_path"),
        "Convert the first worksheet from XLSX to CSV string"
    );
    
    // Sheet discovery functions
    m.def("get_sheet_list", 
        [](const std::string& xlsx_path) -> std::vector<xlsxcsv::SheetMetadata> {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::getSheetList(xlsx_path);
        },
        py::arg("xlsx_path"),
        "Get metadata for all sheets in an XLSX file without reading sheet content"
    );
    
    m.def("get_visible_sheets", 
        [](const std::string& xlsx_path) -> std::vector<xlsxcsv::SheetMetadata> {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::getVisibleSheets(xlsx_path);
        },
        py::arg("xlsx_path"),
        "Get metadata for only visible sheets in an XLSX file"
    );
    
    // Selective parsing functions
    m.def("read_specific_sheet", 
        [](const std::string& xlsx_path, 
           const std::string& sheet_name,
           const xlsxcsv::CsvOptions& options) -> std::string {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readSpecificSheet(xlsx_path, sheet_name, options);
        },
        py::arg("xlsx_path"),
        py::arg("sheet_name"),
        py::arg("options") = xlsxcsv::CsvOptions{},
        "Convert a specific worksheet to CSV by name"
    );
    
    m.def("read_multiple_sheets", 
        [](const std::string& xlsx_path, 
           const std::vector<std::string>& sheet_names,
           const xlsxcsv::CsvOptions& options) -> std::map<std::string, std::string> {
            py::gil_scoped_release gil;  // Release GIL during C++ execution
            return xlsxcsv::readMultipleSheets(xlsx_path, sheet_names, options);
        },
        py::arg("xlsx_path"),
        py::arg("sheet_names"),
        py::arg("options") = xlsxcsv::CsvOptions{},
        "Convert multiple worksheets to CSV by name"
    );
}
