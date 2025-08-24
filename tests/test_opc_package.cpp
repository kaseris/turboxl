#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include <fstream>
#include <filesystem>

namespace fs = std::filesystem;

class OpcPackageTest : public ::testing::Test {
protected:
    void SetUp() override {
        // Create test directory
        testDir = fs::temp_directory_path() / "turboxl_opc_test";
        fs::create_directories(testDir);
        
        // Create a mock XLSX file structure
        createMockXlsxFile();
    }
    
    void TearDown() override {
        // Clean up test files
        if (fs::exists(testDir)) {
            fs::remove_all(testDir);
        }
    }
    
    void createMockXlsxFile() {
        // Create the directory structure for a minimal XLSX file
        testXlsxPath = testDir / "test.xlsx";
        
        auto tempXlsxDir = testDir / "xlsx_content";
        fs::create_directories(tempXlsxDir);
        fs::create_directories(tempXlsxDir / "_rels");
        fs::create_directories(tempXlsxDir / "xl" / "_rels");
        
        // Create [Content_Types].xml
        std::ofstream contentTypes(tempXlsxDir / "[Content_Types].xml");
        contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>)";
        contentTypes.close();
        
        // Create _rels/.rels
        std::ofstream mainRels(tempXlsxDir / "_rels" / ".rels");
        mainRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
        mainRels.close();
        
        // Create xl/workbook.xml
        std::ofstream workbook(tempXlsxDir / "xl" / "workbook.xml");
        workbook << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";
        workbook.close();
        
        // Create ZIP file
        std::string cmd = "cd " + tempXlsxDir.string() + " && zip -r ../test.xlsx . > /dev/null 2>&1";
        system(cmd.c_str());
    }
    
    fs::path testDir;
    fs::path testXlsxPath;
};

TEST_F(OpcPackageTest, DefaultConstruction) {
    xlsxcsv::core::OpcPackage package;
    EXPECT_FALSE(package.isOpen());
}

TEST_F(OpcPackageTest, OpenValidXlsxFile) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    EXPECT_NO_THROW(package.open(testXlsxPath.string()));
    EXPECT_TRUE(package.isOpen());
}

TEST_F(OpcPackageTest, OpenNonExistentFile) {
    xlsxcsv::core::OpcPackage package;
    EXPECT_THROW(package.open("nonexistent.xlsx"), xlsxcsv::core::XlsxError);
    EXPECT_FALSE(package.isOpen());
}

TEST_F(OpcPackageTest, OpenInvalidFile) {
    // Create a non-ZIP file
    auto invalidFile = testDir / "invalid.xlsx";
    std::ofstream file(invalidFile);
    file << "This is not an XLSX file";
    file.close();
    
    xlsxcsv::core::OpcPackage package;
    EXPECT_THROW(package.open(invalidFile.string()), xlsxcsv::core::XlsxError);
    EXPECT_FALSE(package.isOpen());
}

TEST_F(OpcPackageTest, FindWorkbookPath) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(testXlsxPath.string());
    
    std::string workbookPath = package.findWorkbookPath();
    EXPECT_EQ(workbookPath, "xl/workbook.xml");
}

TEST_F(OpcPackageTest, GetContentTypes) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(testXlsxPath.string());
    
    auto contentTypes = package.getContentTypes();
    EXPECT_FALSE(contentTypes.empty());
    
    // Should contain at least the workbook content type
    bool foundWorkbook = false;
    for (const auto& contentType : contentTypes) {
        if (contentType.find("spreadsheetml.sheet.main") != std::string::npos) {
            foundWorkbook = true;
            break;
        }
    }
    EXPECT_TRUE(foundWorkbook);
}

TEST_F(OpcPackageTest, CloseFile) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(testXlsxPath.string());
    EXPECT_TRUE(package.isOpen());
    
    package.close();
    EXPECT_FALSE(package.isOpen());
    
    // Operations should fail after close
    EXPECT_THROW(package.findWorkbookPath(), xlsxcsv::core::XlsxError);
    EXPECT_THROW(package.getContentTypes(), xlsxcsv::core::XlsxError);
}

TEST_F(OpcPackageTest, MoveConstruction) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package1;
    package1.open(testXlsxPath.string());
    EXPECT_TRUE(package1.isOpen());
    
    xlsxcsv::core::OpcPackage package2 = std::move(package1);
    EXPECT_TRUE(package2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(OpcPackageTest, MoveAssignment) {
    if (!fs::exists(testXlsxPath)) {
        GTEST_SKIP() << "Test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package1;
    package1.open(testXlsxPath.string());
    EXPECT_TRUE(package1.isOpen());
    
    xlsxcsv::core::OpcPackage package2;
    package2 = std::move(package1);
    EXPECT_TRUE(package2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(OpcPackageTest, MissingContentTypesFile) {
    // Create a ZIP without [Content_Types].xml
    auto malformedXlsxPath = testDir / "malformed.xlsx";
    auto tempDir = testDir / "malformed_content";
    fs::create_directories(tempDir);
    
    // Create a dummy file
    std::ofstream dummy(tempDir / "dummy.txt");
    dummy << "dummy content";
    dummy.close();
    
    std::string cmd = "cd " + tempDir.string() + " && zip ../malformed.xlsx dummy.txt > /dev/null 2>&1";
    system(cmd.c_str());
    
    if (fs::exists(malformedXlsxPath)) {
        xlsxcsv::core::OpcPackage package;
        EXPECT_THROW(package.open(malformedXlsxPath.string()), xlsxcsv::core::XlsxError);
    }
}

TEST_F(OpcPackageTest, MissingMainRelationshipsFile) {
    // Create a ZIP with [Content_Types].xml but no _rels/.rels
    auto malformedXlsxPath = testDir / "malformed2.xlsx";
    auto tempDir = testDir / "malformed2_content";
    fs::create_directories(tempDir);
    
    // Create only [Content_Types].xml
    std::ofstream contentTypes(tempDir / "[Content_Types].xml");
    contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
</Types>)";
    contentTypes.close();
    
    std::string cmd = "cd " + tempDir.string() + " && zip ../malformed2.xlsx [Content_Types].xml > /dev/null 2>&1";
    system(cmd.c_str());
    
    if (fs::exists(malformedXlsxPath)) {
        xlsxcsv::core::OpcPackage package;
        EXPECT_THROW(package.open(malformedXlsxPath.string()), xlsxcsv::core::XlsxError);
    }
}
