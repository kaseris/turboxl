#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include <fstream>
#include <filesystem>

namespace fs = std::filesystem;

class WorkbookTest : public ::testing::Test {
protected:
    void SetUp() override {
        // Create test directory
        testDir = fs::temp_directory_path() / "turboxl_workbook_test";
        fs::create_directories(testDir);
        
        // Create mock XLSX files for different scenarios
        createBasicXlsxFile();
        createDate1904XlsxFile();
        createMultiSheetXlsxFile();
    }
    
    void TearDown() override {
        // Clean up test files
        if (fs::exists(testDir)) {
            fs::remove_all(testDir);
        }
    }
    
    void createBasicXlsxFile() {
        basicXlsxPath = testDir / "basic.xlsx";
        
        auto tempDir = testDir / "basic_content";
        fs::create_directories(tempDir);
        fs::create_directories(tempDir / "_rels");
        fs::create_directories(tempDir / "xl" / "_rels");
        fs::create_directories(tempDir / "xl" / "worksheets");
        
        // Create [Content_Types].xml
        std::ofstream contentTypes(tempDir / "[Content_Types].xml");
        contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>)";
        contentTypes.close();
        
        // Create _rels/.rels
        std::ofstream mainRels(tempDir / "_rels" / ".rels");
        mainRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
        mainRels.close();
        
        // Create xl/workbook.xml (1900 date system)
        std::ofstream workbook(tempDir / "xl" / "workbook.xml");
        workbook << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="0"/>
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";
        workbook.close();
        
        // Create xl/_rels/workbook.xml.rels
        std::ofstream workbookRels(tempDir / "xl" / "_rels" / "workbook.xml.rels");
        workbookRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>)";
        workbookRels.close();
        
        // Create xl/worksheets/sheet1.xml
        std::ofstream sheet1(tempDir / "xl" / "worksheets" / "sheet1.xml");
        sheet1 << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheetData/>
</worksheet>)";
        sheet1.close();
        
        // Create ZIP file
        std::string cmd = "cd " + tempDir.string() + " && zip -r ../basic.xlsx . > /dev/null 2>&1";
        system(cmd.c_str());
    }
    
    void createDate1904XlsxFile() {
        date1904XlsxPath = testDir / "date1904.xlsx";
        
        auto tempDir = testDir / "date1904_content";
        fs::create_directories(tempDir);
        fs::create_directories(tempDir / "_rels");
        fs::create_directories(tempDir / "xl" / "_rels");
        fs::create_directories(tempDir / "xl" / "worksheets");
        
        // Create [Content_Types].xml
        std::ofstream contentTypes(tempDir / "[Content_Types].xml");
        contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>)";
        contentTypes.close();
        
        // Create _rels/.rels
        std::ofstream mainRels(tempDir / "_rels" / ".rels");
        mainRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
        mainRels.close();
        
        // Create xl/workbook.xml (1904 date system)
        std::ofstream workbook(tempDir / "xl" / "workbook.xml");
        workbook << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="1"/>
    <sheets>
        <sheet name="Data" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";
        workbook.close();
        
        // Create xl/_rels/workbook.xml.rels
        std::ofstream workbookRels(tempDir / "xl" / "_rels" / "workbook.xml.rels");
        workbookRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>)";
        workbookRels.close();
        
        // Create ZIP file
        std::string cmd = "cd " + tempDir.string() + " && zip -r ../date1904.xlsx . > /dev/null 2>&1";
        system(cmd.c_str());
    }
    
    void createMultiSheetXlsxFile() {
        multiSheetXlsxPath = testDir / "multisheet.xlsx";
        
        auto tempDir = testDir / "multisheet_content";
        fs::create_directories(tempDir);
        fs::create_directories(tempDir / "_rels");
        fs::create_directories(tempDir / "xl" / "_rels");
        
        // Create [Content_Types].xml
        std::ofstream contentTypes(tempDir / "[Content_Types].xml");
        contentTypes << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>)";
        contentTypes.close();
        
        // Create _rels/.rels
        std::ofstream mainRels(tempDir / "_rels" / ".rels");
        mainRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";
        mainRels.close();
        
        // Create xl/workbook.xml (multiple sheets, one hidden)
        std::ofstream workbook(tempDir / "xl" / "workbook.xml");
        workbook << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <workbookPr date1904="0"/>
    <sheets>
        <sheet name="Summary" sheetId="1" r:id="rId1"/>
        <sheet name="Data" sheetId="2" r:id="rId2"/>
        <sheet name="Hidden" sheetId="3" r:id="rId3" state="hidden"/>
    </sheets>
</workbook>)";
        workbook.close();
        
        // Create xl/_rels/workbook.xml.rels
        std::ofstream workbookRels(tempDir / "xl" / "_rels" / "workbook.xml.rels");
        workbookRels << R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>)";
        workbookRels.close();
        
        // Create ZIP file
        std::string cmd = "cd " + tempDir.string() + " && zip -r ../multisheet.xlsx . > /dev/null 2>&1";
        system(cmd.c_str());
    }
    
    fs::path testDir;
    fs::path basicXlsxPath;
    fs::path date1904XlsxPath;
    fs::path multiSheetXlsxPath;
};

TEST_F(WorkbookTest, DefaultConstruction) {
    xlsxcsv::core::Workbook workbook;
    EXPECT_FALSE(workbook.isOpen());
    EXPECT_EQ(workbook.getSheetCount(), 0u);
}

TEST_F(WorkbookTest, OpenBasicWorkbook) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    EXPECT_NO_THROW(workbook.open(package));
    EXPECT_TRUE(workbook.isOpen());
}

TEST_F(WorkbookTest, DateSystemDetection1900) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    EXPECT_EQ(workbook.getDateSystem(), xlsxcsv::core::DateSystem::Date1900);
    EXPECT_EQ(workbook.getProperties().dateSystem, xlsxcsv::core::DateSystem::Date1900);
}

TEST_F(WorkbookTest, DateSystemDetection1904) {
    if (!fs::exists(date1904XlsxPath)) {
        GTEST_SKIP() << "Date1904 test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(date1904XlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    EXPECT_EQ(workbook.getDateSystem(), xlsxcsv::core::DateSystem::Date1904);
    EXPECT_EQ(workbook.getProperties().dateSystem, xlsxcsv::core::DateSystem::Date1904);
}

TEST_F(WorkbookTest, SingleSheetInfo) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    EXPECT_EQ(workbook.getSheetCount(), 1u);
    
    auto sheets = workbook.getSheets();
    ASSERT_EQ(sheets.size(), 1u);
    
    const auto& sheet = sheets[0];
    EXPECT_EQ(sheet.name, "Sheet1");
    EXPECT_EQ(sheet.relationshipId, "rId1");
    EXPECT_EQ(sheet.target, "worksheets/sheet1.xml");
    EXPECT_EQ(sheet.sheetId, 1);
    EXPECT_TRUE(sheet.visible);
}

TEST_F(WorkbookTest, MultiSheetInfo) {
    if (!fs::exists(multiSheetXlsxPath)) {
        GTEST_SKIP() << "Multi-sheet test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(multiSheetXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    EXPECT_EQ(workbook.getSheetCount(), 3u);
    
    auto sheets = workbook.getSheets();
    ASSERT_EQ(sheets.size(), 3u);
    
    // Check first sheet
    EXPECT_EQ(sheets[0].name, "Summary");
    EXPECT_EQ(sheets[0].relationshipId, "rId1");
    EXPECT_EQ(sheets[0].sheetId, 1);
    EXPECT_TRUE(sheets[0].visible);
    
    // Check second sheet
    EXPECT_EQ(sheets[1].name, "Data");
    EXPECT_EQ(sheets[1].relationshipId, "rId2");
    EXPECT_EQ(sheets[1].sheetId, 2);
    EXPECT_TRUE(sheets[1].visible);
    
    // Check hidden sheet
    EXPECT_EQ(sheets[2].name, "Hidden");
    EXPECT_EQ(sheets[2].relationshipId, "rId3");
    EXPECT_EQ(sheets[2].sheetId, 3);
    EXPECT_FALSE(sheets[2].visible);
}

TEST_F(WorkbookTest, FindSheetByName) {
    if (!fs::exists(multiSheetXlsxPath)) {
        GTEST_SKIP() << "Multi-sheet test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(multiSheetXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    auto summarySheet = workbook.findSheet("Summary");
    ASSERT_TRUE(summarySheet.has_value());
    EXPECT_EQ(summarySheet->name, "Summary");
    EXPECT_EQ(summarySheet->sheetId, 1);
    
    auto dataSheet = workbook.findSheet("Data");
    ASSERT_TRUE(dataSheet.has_value());
    EXPECT_EQ(dataSheet->name, "Data");
    EXPECT_EQ(dataSheet->sheetId, 2);
    
    auto hiddenSheet = workbook.findSheet("Hidden");
    ASSERT_TRUE(hiddenSheet.has_value());
    EXPECT_EQ(hiddenSheet->name, "Hidden");
    EXPECT_FALSE(hiddenSheet->visible);
    
    auto nonExistentSheet = workbook.findSheet("NonExistent");
    EXPECT_FALSE(nonExistentSheet.has_value());
}

TEST_F(WorkbookTest, FindSheetByIndex) {
    if (!fs::exists(multiSheetXlsxPath)) {
        GTEST_SKIP() << "Multi-sheet test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(multiSheetXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    auto sheet0 = workbook.findSheet(0);
    ASSERT_TRUE(sheet0.has_value());
    EXPECT_EQ(sheet0->name, "Summary");
    
    auto sheet1 = workbook.findSheet(1);
    ASSERT_TRUE(sheet1.has_value());
    EXPECT_EQ(sheet1->name, "Data");
    
    auto sheet2 = workbook.findSheet(2);
    ASSERT_TRUE(sheet2.has_value());
    EXPECT_EQ(sheet2->name, "Hidden");
    
    auto invalidSheet = workbook.findSheet(10);
    EXPECT_FALSE(invalidSheet.has_value());
    
    auto negativeSheet = workbook.findSheet(-1);
    EXPECT_FALSE(negativeSheet.has_value());
}

TEST_F(WorkbookTest, RelationshipMapping) {
    if (!fs::exists(multiSheetXlsxPath)) {
        GTEST_SKIP() << "Multi-sheet test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(multiSheetXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    
    EXPECT_EQ(workbook.resolveRelationshipTarget("rId1"), "worksheets/sheet1.xml");
    EXPECT_EQ(workbook.resolveRelationshipTarget("rId2"), "worksheets/sheet2.xml");
    EXPECT_EQ(workbook.resolveRelationshipTarget("rId3"), "worksheets/sheet3.xml");
    
    EXPECT_THROW(workbook.resolveRelationshipTarget("rId999"), xlsxcsv::core::XlsxError);
}

TEST_F(WorkbookTest, CloseWorkbook) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook;
    workbook.open(package);
    EXPECT_TRUE(workbook.isOpen());
    
    workbook.close();
    EXPECT_FALSE(workbook.isOpen());
    EXPECT_EQ(workbook.getSheetCount(), 0u);
    
    // Operations should fail after close
    EXPECT_THROW(workbook.getSheets(), xlsxcsv::core::XlsxError);
    EXPECT_THROW(workbook.getProperties(), xlsxcsv::core::XlsxError);
}

TEST_F(WorkbookTest, MoveConstruction) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook1;
    workbook1.open(package);
    EXPECT_TRUE(workbook1.isOpen());
    
    xlsxcsv::core::Workbook workbook2 = std::move(workbook1);
    EXPECT_TRUE(workbook2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(WorkbookTest, MoveAssignment) {
    if (!fs::exists(basicXlsxPath)) {
        GTEST_SKIP() << "Basic test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicXlsxPath.string());
    
    xlsxcsv::core::Workbook workbook1;
    workbook1.open(package);
    EXPECT_TRUE(workbook1.isOpen());
    
    xlsxcsv::core::Workbook workbook2;
    workbook2 = std::move(workbook1);
    EXPECT_TRUE(workbook2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}
