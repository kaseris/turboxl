#pragma once
#include "fixture_config.hpp"
#include <cstdlib>
#include <filesystem>
#include <fstream>
#include <initializer_list>
#include <stdexcept>
#include <string>
#include <utility>
#include <vector>

inline std::string fixtureQuote(const std::string& value) {
#ifdef _WIN32
    return "\"" + value + "\"";
#else
    std::string out = "'";
    for (char c : value) out += c == '\'' ? "'\\''" : std::string(1, c);
    return out + "'";
#endif
}

inline void createArchive(const std::filesystem::path& root,
                          const std::filesystem::path& output,
                          std::initializer_list<std::string> members = {"."}) {
    std::string command = fixtureQuote(FIXTURE_PYTHON) + " " + fixtureQuote(FIXTURE_SCRIPT)
        + " --archive " + fixtureQuote(root.string()) + " " + fixtureQuote(output.string());
    for (const auto& member : members) command += " " + fixtureQuote(member);
#ifdef _WIN32
    command = "\"" + command + "\"";
#endif
    if (std::system(command.c_str()) != 0 || !std::filesystem::exists(output))
        throw std::runtime_error("Could not create ZIP fixture: " + output.string());
}

// Builds a minimal workbook package and returns the archive path. Each
// (path, xml) pair in `parts` is written verbatim, replacing the default part
// of the same name (e.g. "xl/workbook.xml" or "xl/styles.xml").
inline std::filesystem::path writePackageWithParts(
    const std::filesystem::path& directory, const std::string& name,
    const std::vector<std::pair<std::string, std::string>>& parts) {
    namespace fs = std::filesystem;
    const auto root = directory / (name + "_parts");
    fs::create_directories(root / "xl");
    fs::create_directories(root / "_rels");
    std::ofstream(root / "[Content_Types].xml") <<
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>)";
    std::ofstream(root / "_rels" / ".rels") <<
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
    std::ofstream(root / "xl" / "workbook.xml") <<
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>)";
    for (const auto& [partPath, partXml] : parts) {
        fs::create_directories((root / partPath).parent_path());
        // Binary mode keeps CR bytes intact on every platform.
        std::ofstream(root / partPath, std::ios::binary | std::ios::trunc) << partXml;
    }
    const auto archive = directory / (name + ".xlsx");
    createArchive(root, archive, {"[Content_Types].xml", "_rels", "xl"});
    return archive;
}

inline std::filesystem::path writePackageWithPart(const std::filesystem::path& directory,
                                                  const std::string& name,
                                                  const std::string& partPath,
                                                  const std::string& partXml) {
    return writePackageWithParts(directory, name, {{partPath, partXml}});
}
