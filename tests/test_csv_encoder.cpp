#include <gtest/gtest.h>
#include "fixture_helpers.hpp"
#include "xlsxcsv/core.hpp"
#include "xlsxcsv.hpp"

namespace {

xlsxcsv::core::CellData numericCell(int column, double value, int styleIndex) {
    xlsxcsv::core::CellData cell;
    cell.coordinate = {1, column};
    cell.type = xlsxcsv::core::CellType::Number;
    cell.value = value;
    cell.styleIndex = styleIndex;
    return cell;
}

std::string encodeDateRow(xlsxcsv::core::DateSystem dateSystem,
                          std::initializer_list<xlsxcsv::core::CellData> cells) {
    xlsxcsv::core::OpcPackage package;
    package.open(INTEGRATION_XLSX);
    xlsxcsv::core::StylesRegistry styles;
    styles.parse(package);

    xlsxcsv::core::RowData row;
    row.rowNumber = 1;
    row.cells = cells;
    xlsxcsv::core::CsvRowCollector collector(nullptr, &styles, dateSystem);
    collector.handleRow(row);
    return collector.getCsvString();
}

} // namespace

TEST(CsvEncoderTest, QuotesDelimiterAndEmbeddedNewline) {
    xlsxcsv::CsvOptions options;
    options.delimiter = ';';
    xlsxcsv::core::CsvRowCollector collector(nullptr, nullptr, xlsxcsv::core::DateSystem::Date1900, &options);
    xlsxcsv::core::RowData row;
    row.rowNumber = 1;
    xlsxcsv::core::CellData cell;
    cell.coordinate = {1, 1};
    cell.type = xlsxcsv::core::CellType::String;
    cell.value = std::string("line;one\nline\"two");
    row.cells.push_back(cell);
    collector.handleRow(row);
    EXPECT_EQ(collector.getCsvString(), "\"line;one\nline\"\"two\"\n");
}

TEST(CsvEncoderTest, PreservesMissingPhysicalRows) {
    xlsxcsv::core::CsvRowCollector collector;

    xlsxcsv::core::RowData first;
    first.rowNumber = 1;
    first.cells.push_back(numericCell(1, 1.0, 0));
    collector.handleRow(first);

    xlsxcsv::core::RowData fifth;
    fifth.rowNumber = 5;
    fifth.cells.push_back(numericCell(4, 5.0, 0));
    collector.handleRow(fifth);

    xlsxcsv::core::RowData fiftieth;
    fiftieth.rowNumber = 50;
    fiftieth.cells.push_back(numericCell(26, 50.0, 0));
    collector.handleRow(fiftieth);

    std::string expected = "1\n\n\n\n,,,5\n";
    expected.append(44, '\n');
    expected += std::string(25, ',') + "50\n";
    EXPECT_EQ(collector.getCsvString(), expected);
    EXPECT_EQ(collector.getRowCount(), 50);
}

TEST(CsvEncoderTest, DoesNotReinsertExcludedHiddenRows) {
    xlsxcsv::CsvOptions options;
    options.includeHiddenRows = false;
    xlsxcsv::core::CsvRowCollector collector(nullptr, nullptr,
        xlsxcsv::core::DateSystem::Date1900, &options);

    xlsxcsv::core::RowData first;
    first.rowNumber = 1;
    first.cells.push_back(numericCell(1, 1.0, 0));
    collector.handleRow(first);

    xlsxcsv::core::RowData hidden;
    hidden.rowNumber = 3;
    hidden.hidden = true;
    hidden.cells.push_back(numericCell(1, 3.0, 0));
    collector.handleRow(hidden);

    xlsxcsv::core::RowData fourth;
    fourth.rowNumber = 4;
    fourth.cells.push_back(numericCell(1, 4.0, 0));
    collector.handleRow(fourth);

    EXPECT_EQ(collector.getCsvString(), "1\n\n4\n");
    EXPECT_EQ(collector.getRowCount(), 3);
}

TEST(CsvEncoderTest, RoundsExcelDatetimeToNearestSecond) {
    EXPECT_EQ(
        encodeDateRow(
            xlsxcsv::core::DateSystem::Date1900,
            {numericCell(1, 45306.0, 1),
             numericCell(2, 45306.57326388889, 2)}),
        "2024-01-15,2024-01-15T13:45:30\n");
}

TEST(CsvEncoderTest, CarriesRoundedDatetimeAcrossMidnight) {
    EXPECT_EQ(
        encodeDateRow(
            xlsxcsv::core::DateSystem::Date1900,
            {numericCell(1, 45306.99999999999, 2)}),
        "2024-01-16T00:00:00\n");
}

TEST(CsvEncoderTest, SupportsExcel1904DateSystem) {
    EXPECT_EQ(
        encodeDateRow(
            xlsxcsv::core::DateSystem::Date1904,
            {numericCell(1, 43844.0, 1),
             numericCell(2, 43844.57326388889, 2)}),
        "2024-01-15,2024-01-15T13:45:30\n");
}

TEST(CsvEncoderTest, PreservesExcel1900LeapDaySemantics) {
    EXPECT_EQ(
        encodeDateRow(
            xlsxcsv::core::DateSystem::Date1900,
            {numericCell(1, 59.0, 1),
             numericCell(2, 60.0, 1),
             numericCell(3, 61.0, 1)}),
        "1900-02-28,1900-02-29,1900-03-01\n");
}
