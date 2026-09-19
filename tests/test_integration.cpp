#include <gtest/gtest.h>
#include "fixture_helpers.hpp"
#include "xlsxcsv.hpp"

TEST(IntegrationTest, EndToEndConversion) {
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX),
              "\"caf\xc3\xa9, \"\"quoted\"\"\",42,2024-01-15,2024-01-15T13:45:30\n");
    xlsxcsv::CsvOptions options;
    options.dateMode = xlsxcsv::CsvOptions::DateMode::RAW;
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, 0, options),
              "\"caf\xc3\xa9, \"\"quoted\"\"\",42,45306,45306.573264\n");
}

TEST(IntegrationTest, PreservesSparseWorksheetRows) {
    std::string expected = "origin\n\n\n\n,,,gap\n";
    expected.append(44, '\n');
    expected += std::string(25, ',') + "far\n";
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, "Sparse"), expected);
}
