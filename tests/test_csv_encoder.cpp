#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include "xlsxcsv.hpp"

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
