#include <gtest/gtest.h>

#include "fixture_config.hpp"
#include "typed_reader.hpp"

#include <string>

namespace {

using xlsxcsv::internal::TypedRowCollector;
using xlsxcsv::internal::TypedWorksheet;

TypedWorksheet parseXml(const std::string& xml, TypedRowCollector& collector) {
    xlsxcsv::core::SheetStreamReader reader;
    const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
    reader.parseSheetData(bytes, collector);
    return collector.takeRows();
}

TEST(TypedRowCollectorTest, PreservesCoordinatesAndPrimitiveValues) {
    const std::string xml =
        R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
        R"(<sheetData><row r="2"><c r="A2"/><c r="B2" t="inlineStr">)"
        R"(<is><t>text</t></is></c><c r="C2" t="e"><v>#N/A</v></c>)"
        R"(<c r="D2" t="b"><v>1</v></c></row><row r="4"><c r="A4"><v>42.5</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedRowCollector collector;
    auto rows = parseXml(xml, collector);

    ASSERT_TRUE(collector.getErrors().empty());
    ASSERT_EQ(rows.size(), 4U);
    for (const auto& row : rows) {
        EXPECT_EQ(row.size(), 4U);
    }
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[0][0]));
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[1][0]));
    EXPECT_EQ(std::get<std::string>(rows[1][1]), "text");
    EXPECT_EQ(std::get<std::string>(rows[1][2]), "#N/A");
    EXPECT_TRUE(std::get<bool>(rows[1][3]));
    EXPECT_DOUBLE_EQ(std::get<double>(rows[3][0]), 42.5);
}

TEST(TypedRowCollectorTest, EmptyWorksheetProducesEmptyMatrix) {
    const std::string xml =
        R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
        R"(<sheetData/></worksheet>)";
    TypedRowCollector collector;
    const auto rows = parseXml(xml, collector);
    EXPECT_TRUE(collector.getErrors().empty());
    EXPECT_TRUE(rows.empty());
}

TEST(TypedRowCollectorTest, EmptyPhysicalRowIsRetained) {
    const std::string xml =
        R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">)"
        R"(<sheetData><row r="1"/><row r="2"><c r="C2"><v>7</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedRowCollector collector;
    const auto rows = parseXml(xml, collector);
    ASSERT_EQ(rows.size(), 2U);
    ASSERT_EQ(rows[0].size(), 3U);
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[0][2]));
    EXPECT_DOUBLE_EQ(std::get<double>(rows[1][2]), 7.0);
}

TEST(TypedRowCollectorTest, ReportsMalformedWorksheet) {
    const std::string xml = R"(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>)";
    TypedRowCollector collector;
    (void)parseXml(xml, collector);
    EXPECT_FALSE(collector.getErrors().empty());
}

TEST(TypedReaderIntegrationTest, ResolvesSharedStringsAndKeepsRawNumbers) {
    const auto rows = xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, "Data");
    ASSERT_EQ(rows.size(), 1U);
    ASSERT_EQ(rows[0].size(), 4U);
    EXPECT_EQ(std::get<std::string>(rows[0][0]), "caf\xc3\xa9, \"quoted\"");
    EXPECT_DOUBLE_EQ(std::get<double>(rows[0][1]), 42.0);
    EXPECT_DOUBLE_EQ(std::get<double>(rows[0][2]), 45306.0);
    EXPECT_DOUBLE_EQ(std::get<double>(rows[0][3]), 45306.57326388889);
}

TEST(TypedReaderIntegrationTest, ProducesRectangularSparseWorksheet) {
    const auto rows = xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, 2);
    ASSERT_EQ(rows.size(), 50U);
    for (const auto& row : rows) {
        EXPECT_EQ(row.size(), 26U);
    }
    EXPECT_EQ(std::get<std::string>(rows[0][0]), "origin");
    EXPECT_EQ(std::get<std::string>(rows[4][3]), "gap");
    EXPECT_EQ(std::get<std::string>(rows[49][25]), "far");
}

TEST(TypedReaderIntegrationTest, RejectsInvalidSheetSelectors) {
    EXPECT_THROW(
        xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, "Missing"),
        std::runtime_error);
    EXPECT_THROW(
        xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, 99),
        std::runtime_error);
}

} // namespace
