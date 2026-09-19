#include <gtest/gtest.h>
#include "fixture_helpers.hpp"
#include "xlsxcsv.hpp"
#include <filesystem>
#include <fstream>

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

TEST(IntegrationTest, FileOutputMatchesStringOutputAndAtomicallyReplaces) {
    namespace fs = std::filesystem;
    const auto output = fs::temp_directory_path() / "turboxl-file-output.csv";
    {
        std::ofstream existing(output, std::ios::binary | std::ios::trunc);
        existing << "old";
    }
    xlsxcsv::CsvOptions options;
    options.includeBom = true;
    options.newline = xlsxcsv::CsvOptions::Newline::CRLF;
    const auto expected = xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, 0, options);
    xlsxcsv::readSheetToFile(INTEGRATION_XLSX, output, 0, options);
    std::ifstream file(output, std::ios::binary);
    const std::string actual((std::istreambuf_iterator<char>(file)), {});
    EXPECT_EQ(actual, expected);
    fs::remove(output);
}

TEST(IntegrationTest, FileOutputPreservesDestinationOnFailure) {
    namespace fs = std::filesystem;
    const auto output = fs::temp_directory_path() / "turboxl-file-failure.csv";
    {
        std::ofstream existing(output, std::ios::binary | std::ios::trunc);
        existing << "preserve-me";
    }
    EXPECT_THROW(xlsxcsv::readSheetToFile("missing.xlsx", output), std::runtime_error);
    std::ifstream file(output, std::ios::binary);
    const std::string actual((std::istreambuf_iterator<char>(file)), {});
    EXPECT_EQ(actual, "preserve-me");
    fs::remove(output);
}

TEST(IntegrationTest, FileOutputRejectsInputAsDestination) {
    EXPECT_THROW(
        xlsxcsv::readSheetToFile(INTEGRATION_XLSX, INTEGRATION_XLSX),
        std::runtime_error);
}
