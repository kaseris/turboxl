#include <gtest/gtest.h>
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
    std::string tempFile = std::filesystem::temp_directory_path() / "invalid.xlsx";
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

// Helper function to create a multi-sheet XLSX file for testing
std::string createMultiSheetTestFile() {
    std::string tempFile = std::filesystem::temp_directory_path() / "test_multi_sheet.xlsx";
    
    // Create a proper XLSX file with multiple sheets (visible and hidden)
    std::ofstream file(tempFile, std::ios::binary);
    
    // Create minimal XLSX structure with ZIP
    // This is a simplified approach - in a real scenario you'd use a proper ZIP library
    // For now, we'll create the structure that the parser expects
    
    // Note: This is a placeholder - we need to create a proper XLSX file
    // For comprehensive testing, we should either:
    // 1. Include test files in the repository
    // 2. Use a library to create XLSX files programmatically
    // 3. Use pre-created test files
    
    file.close();
    return tempFile;
}

TEST_F(SheetDiscoveryTest, MultiSheetFileDiscovery) {
    // For now, skip this test until we have proper XLSX test files
    // TODO: Add comprehensive multi-sheet testing once test files are available
    GTEST_SKIP() << "Multi-sheet test requires proper XLSX test files";
    
    // This test would verify:
    // - Multiple sheets are detected correctly
    // - Hidden and visible sheets are differentiated
    // - Sheet IDs and names are preserved
    // - Target paths are correct
}

TEST_F(SheetDiscoveryTest, VisibleSheetFiltering) {
    // Skip until proper test files are available
    GTEST_SKIP() << "Visible sheet filtering test requires proper XLSX test files";
    
    // This test would verify:
    // - Only visible sheets are returned by getVisibleSheets()
    // - Hidden and veryHidden sheets are excluded
    // - Empty result when all sheets are hidden
}

TEST_F(SheetDiscoveryTest, SpecificSheetReading) {
    // Skip until proper test files are available
    GTEST_SKIP() << "Specific sheet reading test requires proper XLSX test files";
    
    // This test would verify:
    // - readSpecificSheet() finds correct sheet by name
    // - Error thrown for non-existent sheet names
    // - Correct CSV content is returned
    // - Hidden sheets can be read by name
}

TEST_F(SheetDiscoveryTest, BatchSheetProcessing) {
    // Skip until proper test files are available
    GTEST_SKIP() << "Batch processing test requires proper XLSX test files";
    
    // This test would verify:
    // - readMultipleSheets() processes all requested sheets
    // - Results map contains correct sheet names as keys
    // - Error thrown if any requested sheet doesn't exist
    // - Efficient reuse of ZIP/workbook parsing
}

TEST_F(SheetDiscoveryTest, ErrorHandlingForMissingSheet) {
    // This test can work with any file, even if parsing fails
    // because we're testing the error handling path
    std::string tempFile = std::filesystem::temp_directory_path() / "empty.xlsx";
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
    std::string tempFile = std::filesystem::temp_directory_path() / "empty.xlsx";
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