#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"

TEST(SheetStreamReaderTest, ParsesSparseInlineStringsAndNumbers) {
    const std::string xml = R"(<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c><c r="C1"><v>42</v></c></row></sheetData></worksheet>)";
    xlsxcsv::core::SheetStreamReader reader;
    xlsxcsv::core::CsvRowCollector collector;
    reader.parseSheetData(std::vector<uint8_t>(xml.begin(), xml.end()), collector);
    EXPECT_TRUE(collector.getErrors().empty());
    EXPECT_EQ(collector.getRowCount(), 1u);
    EXPECT_EQ(collector.getCsvString(), "Hello,,42\n");
}
