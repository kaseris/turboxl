#include <gtest/gtest.h>
#include "fixture_helpers.hpp"
#include "xlsxcsv.hpp"

TEST(IntegrationTest, EndToEndConversion) {
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX),
              "\"caf\xc3\xa9, \"\"quoted\"\"\",42,2024-01-01\n");
    xlsxcsv::CsvOptions options;
    options.dateMode = xlsxcsv::CsvOptions::DateMode::RAW;
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, 0, options),
              "\"caf\xc3\xa9, \"\"quoted\"\"\",42,45292\n");
}
