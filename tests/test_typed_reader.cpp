#include <gtest/gtest.h>

#include "fixture_config.hpp"
#include "core/fast_typed_sheet_reader.hpp"
#include "typed_reader.hpp"

#include <string>
#include <tuple>

namespace {

using xlsxcsv::internal::TypedRowCollector;
using xlsxcsv::internal::TypedReadOptions;
using xlsxcsv::internal::TypedWorksheet;

TypedWorksheet parseXml(const std::string& xml, TypedRowCollector& collector) {
    xlsxcsv::core::SheetStreamReader reader;
    const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
    reader.parseSheetData(bytes, collector);
    return collector.takeRows();
}

bool parseFastXml(const std::string& xml, TypedRowCollector& collector) {
    const std::vector<std::uint8_t> bytes(xml.begin(), xml.end());
    return xlsxcsv::internal::tryParseTypedWorksheetFast(bytes, collector);
}

xlsxcsv::core::StylesRegistry integrationStyles() {
    xlsxcsv::core::OpcPackage package;
    package.open(INTEGRATION_XLSX);
    xlsxcsv::core::StylesRegistry styles;
    styles.parse(package);
    return styles;
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
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[1][2]));
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
    EXPECT_EQ(std::get<std::int64_t>(rows[1][2]), 7);
}

TEST(TypedRowCollectorTest, CropsToUsedRangeAndPreservesInternalHoles) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="3"><c r="C3"><v>1</v></c></row>)"
        R"(<row r="5"><c r="E5" t="inlineStr"><is><t>end</t></is></c></row>)"
        R"(</sheetData></worksheet>)";
    TypedReadOptions options;
    options.skipEmptyArea = true;
    TypedRowCollector collector(nullptr, options);
    const auto rows = parseXml(xml, collector);

    ASSERT_EQ(rows.size(), 3U);
    ASSERT_EQ(rows.front().size(), 3U);
    EXPECT_EQ(std::get<std::int64_t>(rows[0][0]), 1);
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[1][1]));
    EXPECT_EQ(std::get<std::string>(rows[2][2]), "end");
}

TEST(TypedRowCollectorTest, CroppedFarCellDoesNotAllocateFromOrigin) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1048576"><c r="XFD1048576"><v>7</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedReadOptions options;
    options.skipEmptyArea = true;
    options.maxCells = 1;
    TypedRowCollector collector(nullptr, options);
    const auto rows = parseXml(xml, collector);

    ASSERT_EQ(rows.size(), 1U);
    ASSERT_EQ(rows[0].size(), 1U);
    EXPECT_EQ(std::get<std::int64_t>(rows[0][0]), 7);
}

TEST(TypedRowCollectorTest, FarCellFailsBeforeDenseAllocation) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1048576"><c r="XFD1048576"><v>7</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedReadOptions options;
    options.maxCells = 10'000'000;
    TypedRowCollector collector(nullptr, options);
    try {
        ASSERT_TRUE(parseFastXml(xml, collector));
        FAIL() << "Expected the dense cell limit to fail";
    } catch (const std::runtime_error& error) {
        const std::string message = error.what();
        EXPECT_NE(message.find("1048576 x 16384"), std::string::npos);
        EXPECT_NE(message.find("max_cells=10000000"), std::string::npos);
    }
}

TEST(TypedRowCollectorTest, RejectsZeroCellLimit) {
    TypedReadOptions options;
    options.maxCells = 0;
    EXPECT_THROW(TypedRowCollector(nullptr, options), std::invalid_argument);
}

TEST(TypedRowCollectorTest, ReportsMalformedWorksheet) {
    const std::string xml = R"(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>)";
    TypedRowCollector collector;
    (void)parseXml(xml, collector);
    EXPECT_FALSE(collector.getErrors().empty());
}

TEST(TypedRowCollectorTest, ConvertsPandasCompatiblePrimitiveScalars) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1">)"
        R"(<c r="A1"><v>42</v></c><c r="B1"><v>42.5</v></c>)"
        R"(<c r="C1" t="b"><v>1</v></c><c r="D1" t="e"><v>#N/A</v></c>)"
        R"(<c r="E1" t="str"><f>1+1</f><v>cached</v></c>)"
        R"(<c r="F1"><f>1+1</f></c><c r="G1"><v>1e20</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedRowCollector collector;
    ASSERT_TRUE(parseFastXml(xml, collector));
    const auto rows = collector.takeRows();

    ASSERT_EQ(rows.size(), 1U);
    EXPECT_EQ(std::get<std::int64_t>(rows[0][0]), 42);
    EXPECT_DOUBLE_EQ(std::get<double>(rows[0][1]), 42.5);
    EXPECT_TRUE(std::get<bool>(rows[0][2]));
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[0][3]));
    EXPECT_EQ(std::get<std::string>(rows[0][4]), "cached");
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[0][5]));
    EXPECT_DOUBLE_EQ(std::get<double>(rows[0][6]), 1e20);
}

TEST(TypedRowCollectorTest, ConvertsStyledTemporalValuesAtMicrosecondPrecision) {
    auto styles = integrationStyles();
    const std::string xml =
        R"(<worksheet><sheetData><row r="1">)"
        R"(<c r="A1" s="1"><v>59</v></c><c r="B1" s="1"><v>60</v></c>)"
        R"(<c r="C1" s="1"><v>61</v></c>)"
        R"(<c r="D1" s="2"><v>45292.0009765625</v></c>)"
        R"(<c r="E1" s="3"><v>0.999999999999</v></c>)"
        R"(<c r="F1" s="1"><v>4000000</v></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedRowCollector collector(
        nullptr, {}, &styles, xlsxcsv::core::DateSystem::Date1900);
    ASSERT_TRUE(parseFastXml(xml, collector));
    const auto rows = collector.takeRows();

    const auto& feb28 = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][0]);
    const auto& serial60 = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][1]);
    const auto& mar1 = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][2]);
    EXPECT_EQ(std::tie(feb28.year, feb28.month, feb28.day), std::make_tuple(1900, 2U, 28U));
    EXPECT_EQ(std::tie(serial60.year, serial60.month, serial60.day), std::make_tuple(1900, 2U, 28U));
    EXPECT_EQ(std::tie(mar1.year, mar1.month, mar1.day), std::make_tuple(1900, 3U, 1U));
    const auto& datetime = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][3]);
    EXPECT_EQ(datetime.microsecond, 375'000);
    const auto& time = std::get<xlsxcsv::internal::TypedTime>(rows[0][4]);
    EXPECT_EQ(std::tie(time.hour, time.minute, time.second, time.microsecond),
              std::make_tuple(0, 0, 0, 0));
    EXPECT_EQ(std::get<std::int64_t>(rows[0][5]), 4'000'000);
}

TEST(TypedRowCollectorTest, Supports1904DatesAndCellDataFallback) {
    auto styles = integrationStyles();
    TypedRowCollector collector(
        nullptr, {}, &styles, xlsxcsv::core::DateSystem::Date1904);
    xlsxcsv::core::RowData row;
    row.rowNumber = 1;
    xlsxcsv::core::CellData cell;
    cell.coordinate = {1, 1};
    cell.type = xlsxcsv::core::CellType::Number;
    cell.value = 59.0;
    cell.styleIndex = 1;
    row.cells.push_back(cell);
    collector.handleRow(row);
    const auto rows = collector.takeRows();
    const auto& date = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][0]);
    EXPECT_EQ(std::tie(date.year, date.month, date.day), std::make_tuple(1904, 2U, 29U));
}

TEST(FastTypedSheetReaderTest, ParsesNamespacedPrimitiveCellsAndEntities) {
    const std::string xml =
        R"(<?xml version="1.0"?><x:worksheet xmlns:x="urn:test"><x:sheetData>)"
        R"(<x:row r="2" spans="1:5"><x:c r="A2"><x:v>-12.5</x:v></x:c>)"
        R"(<x:c r="C2" t="b"><x:v>1</x:v></x:c>)"
        R"(<x:c r="D2" t="str"><x:v>A&amp;B&#x20AC;</x:v></x:c>)"
        R"(<x:c r="E2" t="inlineStr"><x:is><x:r><x:t>rich </x:t></x:r>)"
        R"(<x:r><x:t><![CDATA[text]]></x:t></x:r></x:is></x:c>)"
        R"(</x:row></x:sheetData></x:worksheet>)";
    TypedRowCollector collector;
    ASSERT_TRUE(parseFastXml(xml, collector));
    auto rows = collector.takeRows();

    ASSERT_TRUE(collector.getErrors().empty());
    ASSERT_EQ(rows.size(), 2U);
    ASSERT_EQ(rows[1].size(), 5U);
    EXPECT_DOUBLE_EQ(std::get<double>(rows[1][0]), -12.5);
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[1][1]));
    EXPECT_TRUE(std::get<bool>(rows[1][2]));
    EXPECT_EQ(std::get<std::string>(rows[1][3]), "A&B\xe2\x82\xac");
    EXPECT_EQ(std::get<std::string>(rows[1][4]), "rich text");
}

TEST(FastTypedSheetReaderTest, HandlesEmptyRowsAndCells) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1"/><row r="3">)"
        R"(<c r="B3"/><c r="D3" t="inlineStr"><is><t/></is></c>)"
        R"(</row></sheetData></worksheet>)";
    TypedRowCollector collector;
    ASSERT_TRUE(parseFastXml(xml, collector));
    auto rows = collector.takeRows();

    ASSERT_EQ(rows.size(), 3U);
    ASSERT_EQ(rows.front().size(), 4U);
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[2][1]));
    EXPECT_EQ(std::get<std::string>(rows[2][3]), "");
}

TEST(FastTypedSheetReaderTest, NrowsStopsBeforeMalformedTail) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row>)"
        R"(<row r="2"><c r="A2"><v>this tail is never closed)";
    TypedReadOptions options;
    options.nrows = 1;
    TypedRowCollector fastCollector(nullptr, options);
    ASSERT_TRUE(parseFastXml(xml, fastCollector));
    const auto fastRows = fastCollector.takeRows();
    ASSERT_EQ(fastRows.size(), 1U);
    EXPECT_EQ(std::get<std::int64_t>(fastRows[0][0]), 1);

    TypedRowCollector fallbackCollector(nullptr, options);
    const auto fallbackRows = parseXml(xml, fallbackCollector);
    EXPECT_TRUE(fallbackCollector.getErrors().empty());
    ASSERT_EQ(fallbackRows.size(), 1U);
    EXPECT_EQ(std::get<std::int64_t>(fallbackRows[0][0]), 1);
}

TEST(FastTypedSheetReaderTest, NrowsCountsMissingPhysicalRows) {
    const std::string xml =
        R"(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c></row>)"
        R"(<row r="5"><c r="Z5"><v>99</v></c></row></sheetData></worksheet>)";
    TypedReadOptions options;
    options.nrows = 3;
    TypedRowCollector collector(nullptr, options);
    ASSERT_TRUE(parseFastXml(xml, collector));
    const auto rows = collector.takeRows();

    ASSERT_EQ(rows.size(), 3U);
    ASSERT_EQ(rows[0].size(), 1U);
    EXPECT_EQ(std::get<std::int64_t>(rows[0][0]), 1);
    EXPECT_TRUE(std::holds_alternative<std::monostate>(rows[2][0]));
}

TEST(FastTypedSheetReaderTest, ZeroRowsDoesNotReadXml) {
    TypedReadOptions options;
    options.nrows = 0;
    TypedRowCollector collector(nullptr, options);
    EXPECT_TRUE(parseFastXml("", collector));
    EXPECT_TRUE(collector.takeRows().empty());
}

TEST(FastTypedSheetReaderTest, DeclinesUnsupportedOrMalformedXml) {
    TypedRowCollector doctypeCollector;
    EXPECT_FALSE(parseFastXml(
        R"(<!DOCTYPE worksheet [<!ENTITY x "text">]><worksheet/>)",
        doctypeCollector));

    TypedRowCollector malformedCollector;
    EXPECT_FALSE(parseFastXml(
        R"(<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c>)",
        malformedCollector));

    TypedRowCollector wrongRootCollector;
    EXPECT_FALSE(parseFastXml(R"(<not-a-worksheet/>)", wrongRootCollector));
}

TEST(TypedReaderIntegrationTest, ResolvesSharedStringsNumbersAndDates) {
    const auto rows = xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, "Data");
    ASSERT_EQ(rows.size(), 1U);
    ASSERT_EQ(rows[0].size(), 4U);
    EXPECT_EQ(std::get<std::string>(rows[0][0]), "caf\xc3\xa9, \"quoted\"");
    EXPECT_EQ(std::get<std::int64_t>(rows[0][1]), 42);
    const auto& date = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][2]);
    EXPECT_EQ(date.year, 2024);
    EXPECT_EQ(date.month, 1U);
    EXPECT_EQ(date.day, 15U);
    const auto& datetime = std::get<xlsxcsv::internal::TypedDateTime>(rows[0][3]);
    EXPECT_EQ(datetime.year, 2024);
    EXPECT_EQ(datetime.month, 1U);
    EXPECT_EQ(datetime.day, 15U);
    EXPECT_EQ(datetime.hour, 13);
    EXPECT_EQ(datetime.minute, 45);
    EXPECT_EQ(datetime.second, 30);
}

TEST(TypedReaderIntegrationTest, ProducesRectangularSparseWorksheet) {
    const auto rows = xlsxcsv::internal::readSheetToTyped(INTEGRATION_XLSX, 3);
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
