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
#include <memory>
#include <limits>

namespace nb = nanobind;

namespace {

nb::object boxTypedValue(
    const xlsxcsv::internal::TypedCellValue& value,
    const nb::object& datetimeType,
    const nb::object& timeType) {
    switch (value.index()) {
        case 0:
            return nb::none();
        case 1:
            return nb::bool_(std::get<1>(value));
        case 2:
            return nb::int_(std::get<2>(value));
        case 3:
            return nb::float_(std::get<3>(value));
        case 4: {
            const auto& text = std::get<4>(value);
            return nb::str(text.data(), text.size());
        }
        case 5: {
            const auto& date = std::get<5>(value);
            return datetimeType(
                date.year, date.month, date.day, date.hour, date.minute,
                date.second, date.microsecond);
        }
        case 6: {
            const auto& time = std::get<6>(value);
            return timeType(time.hour, time.minute, time.second, time.microsecond);
        }
        default:
            return nb::float_(std::get<7>(value).value);
    }
}

bool profileTypedTimings() {
    const char* value = std::getenv("TURBOXL_PROFILE_TYPED_TIMINGS");
    return value && (value[0] == '1' || value[0] == 't' || value[0] == 'T' ||
                     value[0] == 'y' || value[0] == 'Y');
}

nb::list boxTypedWorksheet(const xlsxcsv::internal::TypedWorksheet& rows,
                           const nb::object& datetimeType, const nb::object& timeType) {
    nb::list result = nb::steal<nb::list>(PyList_New(static_cast<Py_ssize_t>(rows.size())));
    for (std::size_t i = 0; i < rows.size(); ++i) {
        nb::list row = nb::steal<nb::list>(PyList_New(static_cast<Py_ssize_t>(rows[i].size())));
        for (std::size_t j = 0; j < rows[i].size(); ++j) {
            auto value = boxTypedValue(rows[i][j], datetimeType, timeType);
            if (PyList_SetItem(row.ptr(), static_cast<Py_ssize_t>(j), value.release().ptr()) < 0)
                throw nb::python_error();
        }
        if (PyList_SetItem(result.ptr(), static_cast<Py_ssize_t>(i), row.release().ptr()) < 0)
            throw nb::python_error();
    }
    return result;
}

struct PySheet {
    std::shared_ptr<xlsxcsv::internal::TypedWorkbookSession> session;
    xlsxcsv::core::SheetInfo info;
};

struct PyWorkbook {
    std::shared_ptr<xlsxcsv::internal::TypedWorkbookSession> session;
};

xlsxcsv::SheetMetadata metadata(const xlsxcsv::core::SheetInfo& sheet) {
    return {sheet.name, sheet.sheetId, sheet.visible, sheet.target,
            static_cast<xlsxcsv::SheetKind>(sheet.kind),
            static_cast<xlsxcsv::SheetVisibility>(sheet.visibility)};
}

} // namespace

NB_MODULE(_turboxl, m) {
    m.doc() = "Fast XLSX to CSV converter (C++ core with Python bindings)";
    const nb::object datetimeModule = nb::module_::import_("datetime");
    const nb::object datetimeType = datetimeModule.attr("datetime");
    const nb::object timeType = datetimeModule.attr("time");
    
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

    nb::class_<PySheet>(m, "Sheet")
        .def("to_python", [datetimeType, timeType](const PySheet& self, bool skipEmptyArea,
                const std::optional<std::int64_t>& nrows) {
            if (nrows && *nrows < 0) throw nb::value_error("nrows must be non-negative or None");
            xlsxcsv::internal::TypedReadOptions options;
            options.skipEmptyArea = skipEmptyArea;
            if (nrows) options.nrows = static_cast<std::size_t>(*nrows);
            using Clock = std::chrono::steady_clock;
            const auto totalStart = Clock::now();
            xlsxcsv::internal::TypedWorksheet rows;
            const auto nativeStart = Clock::now();
            { nb::gil_scoped_release release; rows = self.session->read(self.info, options); }
            const auto nativeEnd = Clock::now();
            const auto boxingStart = Clock::now();
            auto result = boxTypedWorksheet(rows, datetimeType, timeType);
            const auto boxingEnd = Clock::now();
            if (profileTypedTimings()) {
                const auto milliseconds = [](auto duration) {
                    return std::chrono::duration<double, std::milli>(duration).count();
                };
                const std::size_t columns = rows.empty() ? 0 : rows.front().size();
                std::cerr
                    << "turboxl_typed_timing_ms"
                    << " native=" << milliseconds(nativeEnd - nativeStart)
                    << " boxing=" << milliseconds(boxingEnd - boxingStart)
                    << " total=" << milliseconds(boxingEnd - totalStart)
                    << " rows=" << rows.size()
                    << " columns=" << columns << '\n';
            }
            return result;
        }, nb::kw_only(), nb::arg("skip_empty_area") = false, nb::arg("nrows") = nb::none(),
        "Return dense rectangular rows of Python scalar values. The workbook's "
        "max_cells limit applies to every call; raises RuntimeError after close().");

    nb::class_<PyWorkbook>(m, "Workbook")
        .def_prop_ro("sheet_names", [](const PyWorkbook& self) {
            nb::list names;
            for (const auto& sheet : self.session->sheets())
                if (sheet.kind == xlsxcsv::core::SheetKind::Worksheet)
                    names.append(nb::str(sheet.name.data(), sheet.name.size()));
            return names;
        })
        .def_prop_ro("sheets_metadata", [](const PyWorkbook& self) {
            std::vector<xlsxcsv::SheetMetadata> result;
            for (const auto& sheet : self.session->sheets()) result.push_back(metadata(sheet));
            return result;
        })
        .def("get_sheet_by_name", [](const PyWorkbook& self, const std::string& name) {
            auto sheet = self.session->sheetByName(name);
            if (!sheet) throw nb::key_error(("Worksheet not found: " + name).c_str());
            return PySheet{self.session, *sheet};
        }, nb::arg("name"))
        .def("get_sheet_by_index", [](const PyWorkbook& self, std::int64_t index) {
            if (index < 0 || index > std::numeric_limits<int>::max()) throw nb::index_error("Worksheet index out of range");
            auto sheet = self.session->sheetByIndex(static_cast<int>(index));
            if (!sheet) throw nb::index_error("Worksheet index out of range");
            return PySheet{self.session, *sheet};
        }, nb::arg("index"))
        .def("close", [](const PyWorkbook& self) { self.session->close(); })
        .def("__enter__", [](const PyWorkbook& self) -> const PyWorkbook& {
            if (!self.session->isOpen()) throw std::runtime_error("Workbook is closed"); return self;
        }, nb::rv_policy::reference_internal)
        .def("__exit__", [](const PyWorkbook& self, nb::args) {
            self.session->close(); return false;
        });

    m.def("_load_workbook_path", [](const std::string& path, std::int64_t maxCells) {
        if (maxCells <= 0) throw nb::value_error("max_cells must be greater than zero");
        auto session = std::make_shared<xlsxcsv::internal::TypedWorkbookSession>(path, static_cast<std::size_t>(maxCells));
        return PyWorkbook{std::move(session)};
    }, nb::arg("path"), nb::arg("max_cells") = 10'000'000);
    m.def("_load_workbook_bytes", [](nb::bytes input, std::int64_t maxCells) {
        if (maxCells <= 0) throw nb::value_error("max_cells must be greater than zero");
        const auto* data = reinterpret_cast<const std::uint8_t*>(input.c_str());
        xlsxcsv::core::ByteVector bytes(data, data + input.size());
        auto session = std::make_shared<xlsxcsv::internal::TypedWorkbookSession>(std::move(bytes), static_cast<std::size_t>(maxCells));
        return PyWorkbook{std::move(session)};
    }, nb::arg("data"), nb::arg("max_cells") = 10'000'000);
    
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
        [datetimeType, timeType](const std::string& xlsx_path,
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
            nb::list rows = nb::steal<nb::list>(
                PyList_New(static_cast<Py_ssize_t>(nativeRows.size())));
            for (std::size_t rowIndex = 0; rowIndex < nativeRows.size(); ++rowIndex) {
                const auto& nativeRow = nativeRows[rowIndex];
                nb::list row = nb::steal<nb::list>(
                    PyList_New(static_cast<Py_ssize_t>(nativeRow.size())));
                for (std::size_t column = 0; column < nativeRow.size(); ++column) {
                    auto boxed = boxTypedValue(
                        nativeRow[column], datetimeType, timeType);
                    if (PyList_SetItem(
                            row.ptr(), static_cast<Py_ssize_t>(column),
                            boxed.release().ptr()) < 0) {
                        throw nb::python_error();
                    }
                }
                if (PyList_SetItem(
                        rows.ptr(), static_cast<Py_ssize_t>(rowIndex),
                        row.release().ptr()) < 0) {
                    throw nb::python_error();
                }
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
