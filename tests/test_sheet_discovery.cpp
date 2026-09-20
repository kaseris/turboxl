#include <gtest/gtest.h>
#include "fixture_helpers.hpp"
#include "xlsxcsv.hpp"
#include <filesystem>
#include <fstream>

class SheetDiscoveryTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(SheetDiscoveryTest, SheetMetadataConstruction) {
    xlsxcsv::SheetMetadata metadata;
    
    // Test default construction
    EXPECT_EQ(metadata.name, "");
    EXPECT_EQ(metadata.sheetId, 0);
    EXPECT_FALSE(metadata.visible);
    EXPECT_EQ(metadata.target, "");
    EXPECT_EQ(metadata.kind, xlsxcsv::SheetKind::Other);
    EXPECT_EQ(metadata.visibility, xlsxcsv::SheetVisibility::Hidden);
    
    // Test assignment
    metadata.name = "Test Sheet";
    metadata.sheetId = 1;
    metadata.visible = true;
    metadata.target = "worksheets/sheet1.xml";
    
    EXPECT_EQ(metadata.name, "Test Sheet");
    EXPECT_EQ(metadata.sheetId, 1);
    EXPECT_TRUE(metadata.visible);
    EXPECT_EQ(metadata.target, "worksheets/sheet1.xml");
}

TEST_F(SheetDiscoveryTest, ErrorHandlingForNonExistentFile) {
    // Test that non-existent file throws appropriate error
    EXPECT_THROW(xlsxcsv::getSheetList("non_existent_file.xlsx"), std::runtime_error);
    EXPECT_THROW(xlsxcsv::getVisibleSheets("non_existent_file.xlsx"), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readSpecificSheet("non_existent_file.xlsx", "Sheet1"), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readMultipleSheets("non_existent_file.xlsx", {"Sheet1"}), std::runtime_error);
}

TEST_F(SheetDiscoveryTest, ErrorHandlingForInvalidFile) {
    // Create a temporary non-XLSX file
    std::string tempFile = (std::filesystem::temp_directory_path() / "invalid.xlsx").string();
    {
        std::ofstream file(tempFile);
        file << "This is not an XLSX file";
    }
    
    EXPECT_THROW(xlsxcsv::getSheetList(tempFile), std::runtime_error);
    EXPECT_THROW(xlsxcsv::getVisibleSheets(tempFile), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readSpecificSheet(tempFile, "Sheet1"), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readMultipleSheets(tempFile, {"Sheet1"}), std::runtime_error);
    
    // Clean up
    std::filesystem::remove(tempFile);
}

TEST_F(SheetDiscoveryTest, MultiSheetFileDiscovery) {
    const auto sheets = xlsxcsv::getSheetList(INTEGRATION_XLSX);
    ASSERT_EQ(sheets.size(), 6u);
    EXPECT_EQ(sheets[0].kind, xlsxcsv::SheetKind::Chartsheet);
    EXPECT_EQ(sheets[1].name, "Data");
    EXPECT_EQ(sheets[1].kind, xlsxcsv::SheetKind::Worksheet);
    EXPECT_EQ(sheets[2].visibility, xlsxcsv::SheetVisibility::Hidden);
    EXPECT_EQ(sheets[3].visibility, xlsxcsv::SheetVisibility::VeryHidden);
    EXPECT_EQ(sheets[4].kind, xlsxcsv::SheetKind::Other);
    EXPECT_EQ(sheets[5].name, "Sparse");
}

TEST_F(SheetDiscoveryTest, VisibleSheetFiltering) {
    const auto sheets = xlsxcsv::getVisibleSheets(INTEGRATION_XLSX);
    ASSERT_EQ(sheets.size(), 4u);
    EXPECT_EQ(sheets[0].name, "Chart");
    EXPECT_EQ(sheets[1].name, "Data");
    EXPECT_EQ(sheets[2].name, "Dialog");
    EXPECT_EQ(sheets[3].name, "Sparse");
}

TEST_F(SheetDiscoveryTest, WorksheetIndexesExcludeOtherSheetKinds) {
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, 0),
              xlsxcsv::readSpecificSheet(INTEGRATION_XLSX, "Data"));
    EXPECT_EQ(xlsxcsv::readSheetToCsv(INTEGRATION_XLSX, 1), "secret\n");
    EXPECT_THROW(xlsxcsv::readSpecificSheet(INTEGRATION_XLSX, "Chart"),
                 std::runtime_error);
    EXPECT_THROW(xlsxcsv::readSpecificSheet(INTEGRATION_XLSX, "Dialog"),
                 std::runtime_error);
}

TEST_F(SheetDiscoveryTest, SpecificSheetReading) {
    EXPECT_EQ(xlsxcsv::readSpecificSheet(INTEGRATION_XLSX, "Hidden"), "secret\n");
    EXPECT_THROW(xlsxcsv::readSpecificSheet(INTEGRATION_XLSX, "Missing"), std::runtime_error);
}

TEST_F(SheetDiscoveryTest, BatchSheetProcessing) {
    const auto sheets = xlsxcsv::readMultipleSheets(INTEGRATION_XLSX, {"Data", "Hidden"});
    ASSERT_EQ(sheets.size(), 2u);
    EXPECT_EQ(sheets.at("Hidden"), "secret\n");
    EXPECT_EQ(sheets.at("Data"), xlsxcsv::readSheetToCsv(INTEGRATION_XLSX));
    EXPECT_THROW(xlsxcsv::readMultipleSheets(INTEGRATION_XLSX, {"Missing"}), std::runtime_error);
}

TEST_F(SheetDiscoveryTest, ErrorHandlingForMissingSheet) {
    // This test can work with any file, even if parsing fails
    // because we're testing the error handling path
    std::string tempFile = (std::filesystem::temp_directory_path() / "empty.xlsx").string();
    {
        std::ofstream file(tempFile);
        file << ""; // Empty file
    }
    
    // These should all throw because the file is invalid
    EXPECT_THROW(xlsxcsv::readSpecificSheet(tempFile, "NonExistentSheet"), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readMultipleSheets(tempFile, {"Sheet1", "Sheet2"}), std::runtime_error);
    
    // Clean up
    std::filesystem::remove(tempFile);
}

TEST_F(SheetDiscoveryTest, EmptySheetNameHandling) {
    std::string tempFile = (std::filesystem::temp_directory_path() / "empty.xlsx").string();
    {
        std::ofstream file(tempFile);
        file << ""; // Empty file
    }
    
    // Test with empty sheet name
    EXPECT_THROW(xlsxcsv::readSpecificSheet(tempFile, ""), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readMultipleSheets(tempFile, {""}), std::runtime_error);
    EXPECT_THROW(xlsxcsv::readMultipleSheets(tempFile, {}), std::runtime_error); // Empty vector
    
    // Clean up
    std::filesystem::remove(tempFile);
}
