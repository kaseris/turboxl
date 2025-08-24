#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include <filesystem>
#include <fstream>
#include <sstream>
#include <cstdlib>

namespace fs = std::filesystem;

class StylesRegistryTest : public ::testing::Test {
protected:
    void SetUp() override {
        testDir = fs::temp_directory_path() / "turboxl_styles_test";
        fs::create_directories(testDir);
        
        basicStylesXlsxPath = testDir / "basic_styles.xlsx";
        complexStylesXlsxPath = testDir / "complex_styles.xlsx";
        dateFormatsXlsxPath = testDir / "date_formats.xlsx";
        
        createBasicStylesXlsx();
        createComplexStylesXlsx();
        createDateFormatsXlsx();
    }
    
    void TearDown() override {
        fs::remove_all(testDir);
    }
    
    void createBasicStylesXlsx() {
        std::string zipCmd = "cd \"" + testDir.string() + "\" && zip -q \"" + 
                            basicStylesXlsxPath.filename().string() + "\" ";
        
        // Create directory structure
        fs::create_directories(testDir / "xl");
        fs::create_directories(testDir / "_rels");
        
        // [Content_Types].xml
        std::ofstream(testDir / "[Content_Types].xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>)";

        // _rels/.rels
        std::ofstream(testDir / "_rels" / ".rels") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

        // xl/workbook.xml
        std::ofstream(testDir / "xl" / "workbook.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";

        // xl/styles.xml - Basic styles
        std::ofstream(testDir / "xl" / "styles.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="2">
        <numFmt numFmtId="164" formatCode="yyyy-mm-dd"/>
        <numFmt numFmtId="165" formatCode="0.00%"/>
    </numFmts>
    <fonts count="2">
        <font>
            <sz val="11"/>
            <name val="Calibri"/>
        </font>
        <font>
            <sz val="12"/>
            <name val="Arial"/>
            <b/>
            <i/>
            <color rgb="FF0000FF"/>
        </font>
    </fonts>
    <fills count="2">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFFF0000"/>
            </patternFill>
        </fill>
    </fills>
    <borders count="2">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border>
            <left style="thin"/>
            <right style="thick"/>
            <top style="medium"/>
            <bottom style="double"/>
            <diagonal/>
        </border>
    </borders>
    <cellXfs count="4">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="1" fillId="1" borderId="1"/>
        <xf numFmtId="165" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="14" fontId="1" fillId="0" borderId="1"/>
    </cellXfs>
</styleSheet>)";

        // Create ZIP
        zipCmd += "[Content_Types].xml _rels xl";
        system(zipCmd.c_str());
        
        // Clean up temp files
        fs::remove_all(testDir / "[Content_Types].xml");
        fs::remove_all(testDir / "_rels");
        fs::remove_all(testDir / "xl");
    }
    
    void createComplexStylesXlsx() {
        std::string zipCmd = "cd \"" + testDir.string() + "\" && zip -q \"" + 
                            complexStylesXlsxPath.filename().string() + "\" ";
        
        // Create directory structure  
        fs::create_directories(testDir / "xl");
        fs::create_directories(testDir / "_rels");
        
        // [Content_Types].xml
        std::ofstream(testDir / "[Content_Types].xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>)";

        // _rels/.rels
        std::ofstream(testDir / "_rels" / ".rels") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

        // xl/workbook.xml
        std::ofstream(testDir / "xl" / "workbook.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";

        // xl/styles.xml - Complex styles with multiple number formats
        std::ofstream(testDir / "xl" / "styles.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="5">
        <numFmt numFmtId="164" formatCode="[$-409]d/m/yyyy;@"/>
        <numFmt numFmtId="165" formatCode="h:mm:ss AM/PM"/>
        <numFmt numFmtId="166" formatCode="m/d/yyyy h:mm"/>
        <numFmt numFmtId="167" formatCode="$#,##0.00"/>
        <numFmt numFmtId="168" formatCode="0.00E+00"/>
    </numFmts>
    <fonts count="3">
        <font>
            <sz val="11"/>
            <name val="Calibri"/>
        </font>
        <font>
            <sz val="14"/>
            <name val="Times New Roman"/>
            <b/>
            <color rgb="FF800080"/>
        </font>
        <font>
            <sz val="10"/>
            <name val="Courier New"/>
            <u/>
        </font>
    </fonts>
    <fills count="3">
        <fill>
            <patternFill patternType="none"/>
        </fill>
        <fill>
            <patternFill patternType="solid">
                <fgColor rgb="FFFFFF00"/>
                <bgColor rgb="FF000000"/>
            </patternFill>
        </fill>
        <fill>
            <patternFill patternType="lightGray"/>
        </fill>
    </fills>
    <borders count="3">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
        <border>
            <left style="thin">
                <color rgb="FF000000"/>
            </left>
            <right style="thin">
                <color rgb="FF000000"/>
            </right>
            <top style="thin">
                <color rgb="FF000000"/>
            </top>
            <bottom style="thin">
                <color rgb="FF000000"/>
            </bottom>
            <diagonal/>
        </border>
        <border>
            <left style="thick">
                <color rgb="FFFF0000"/>
            </left>
            <right style="thick">
                <color rgb="FFFF0000"/>
            </right>
            <top style="thick">
                <color rgb="FFFF0000"/>
            </top>
            <bottom style="thick">
                <color rgb="FFFF0000"/>
            </bottom>
            <diagonal/>
        </border>
    </borders>
    <cellXfs count="6">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="164" fontId="1" fillId="1" borderId="1"/>
        <xf numFmtId="165" fontId="2" fillId="2" borderId="2"/>
        <xf numFmtId="166" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="167" fontId="1" fillId="1" borderId="1"/>
        <xf numFmtId="168" fontId="2" fillId="0" borderId="0"/>
    </cellXfs>
</styleSheet>)";

        // Create ZIP
        zipCmd += "[Content_Types].xml _rels xl";
        system(zipCmd.c_str());
        
        // Clean up temp files
        fs::remove_all(testDir / "[Content_Types].xml");
        fs::remove_all(testDir / "_rels");
        fs::remove_all(testDir / "xl");
    }
    
    void createDateFormatsXlsx() {
        std::string zipCmd = "cd \"" + testDir.string() + "\" && zip -q \"" + 
                            dateFormatsXlsxPath.filename().string() + "\" ";
        
        // Create directory structure
        fs::create_directories(testDir / "xl");
        fs::create_directories(testDir / "_rels");
        
        // [Content_Types].xml
        std::ofstream(testDir / "[Content_Types].xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
    <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>)";

        // _rels/.rels
        std::ofstream(testDir / "_rels" / ".rels") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

        // xl/workbook.xml
        std::ofstream(testDir / "xl" / "workbook.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";

        // xl/styles.xml - Focus on date/time formats
        std::ofstream(testDir / "xl" / "styles.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <numFmts count="3">
        <numFmt numFmtId="170" formatCode="yyyy-mm-dd hh:mm:ss"/>
        <numFmt numFmtId="171" formatCode="[$-en-US]mmmm d, yyyy"/>
        <numFmt numFmtId="172" formatCode="[h]:mm:ss"/>
    </numFmts>
    <fonts count="1">
        <font>
            <sz val="11"/>
            <name val="Calibri"/>
        </font>
    </fonts>
    <fills count="1">
        <fill>
            <patternFill patternType="none"/>
        </fill>
    </fills>
    <borders count="1">
        <border>
            <left/>
            <right/>
            <top/>
            <bottom/>
            <diagonal/>
        </border>
    </borders>
    <cellXfs count="6">
        <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="14" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="18" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="170" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="171" fontId="0" fillId="0" borderId="0"/>
        <xf numFmtId="172" fontId="0" fillId="0" borderId="0"/>
    </cellXfs>
</styleSheet>)";

        // Create ZIP
        zipCmd += "[Content_Types].xml _rels xl";
        system(zipCmd.c_str());
        
        // Clean up temp files
        fs::remove_all(testDir / "[Content_Types].xml");
        fs::remove_all(testDir / "_rels");
        fs::remove_all(testDir / "xl");
    }
    
    fs::path testDir;
    fs::path basicStylesXlsxPath;
    fs::path complexStylesXlsxPath;
    fs::path dateFormatsXlsxPath;
};

TEST_F(StylesRegistryTest, DefaultConstruction) {
    xlsxcsv::core::StylesRegistry registry;
    EXPECT_FALSE(registry.isOpen());
    EXPECT_EQ(registry.getStyleCount(), 0);
    EXPECT_EQ(registry.getNumberFormatCount(), 0);
}

TEST_F(StylesRegistryTest, ParseBasicStyles) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    
    EXPECT_TRUE(registry.isOpen());
    EXPECT_EQ(registry.getStyleCount(), 4); // 4 cellXfs
    EXPECT_EQ(registry.getNumberFormatCount(), 2); // 2 custom numFmts
}

TEST_F(StylesRegistryTest, GetCellStyle) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    
    // Test first cell style
    auto style0 = registry.getCellStyle(0);
    ASSERT_TRUE(style0.has_value());
    EXPECT_EQ(style0->styleIndex, 0);
    EXPECT_EQ(style0->numberFormat.formatId, 0);
    EXPECT_EQ(style0->font.name, "Calibri");
    EXPECT_EQ(style0->font.size, 11.0);
    EXPECT_FALSE(style0->font.bold);
    
    // Test second cell style (with custom format and styled font)
    auto style1 = registry.getCellStyle(1);
    ASSERT_TRUE(style1.has_value());
    EXPECT_EQ(style1->styleIndex, 1);
    EXPECT_EQ(style1->numberFormat.formatId, 164);
    EXPECT_EQ(style1->font.name, "Arial");
    EXPECT_EQ(style1->font.size, 12.0);
    EXPECT_TRUE(style1->font.bold);
    EXPECT_TRUE(style1->font.italic);
    EXPECT_EQ(style1->font.color, "FF0000FF");
    
    // Test invalid style index
    auto invalidStyle = registry.getCellStyle(100);
    EXPECT_FALSE(invalidStyle.has_value());
}

TEST_F(StylesRegistryTest, NumberFormatDetection) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    
    // Test custom number formats
    auto format164 = registry.getNumberFormat(164);
    ASSERT_TRUE(format164.has_value());
    EXPECT_EQ(format164->formatCode, "yyyy-mm-dd");
    EXPECT_EQ(format164->type, xlsxcsv::core::NumberFormatType::Date);
    EXPECT_FALSE(format164->isBuiltIn);
    
    auto format165 = registry.getNumberFormat(165);
    ASSERT_TRUE(format165.has_value());
    EXPECT_EQ(format165->formatCode, "0.00%");
    EXPECT_EQ(format165->type, xlsxcsv::core::NumberFormatType::Percentage);
    EXPECT_FALSE(format165->isBuiltIn);
    
    // Test built-in number formats
    auto format0 = registry.getNumberFormat(0);
    ASSERT_TRUE(format0.has_value());
    EXPECT_EQ(format0->formatCode, "General");
    EXPECT_EQ(format0->type, xlsxcsv::core::NumberFormatType::General);
    EXPECT_TRUE(format0->isBuiltIn);
    
    auto format14 = registry.getNumberFormat(14);
    ASSERT_TRUE(format14.has_value());
    EXPECT_EQ(format14->formatCode, "mm-dd-yy");
    EXPECT_EQ(format14->type, xlsxcsv::core::NumberFormatType::Date);
    EXPECT_TRUE(format14->isBuiltIn);
}

TEST_F(StylesRegistryTest, NumberFormatTypeDetection) {
    xlsxcsv::core::StylesRegistry registry;
    
    // Test General
    EXPECT_EQ(registry.detectNumberFormatType("General"), xlsxcsv::core::NumberFormatType::General);
    EXPECT_EQ(registry.detectNumberFormatType(""), xlsxcsv::core::NumberFormatType::General);
    
    // Test Date formats
    EXPECT_EQ(registry.detectNumberFormatType("yyyy-mm-dd"), xlsxcsv::core::NumberFormatType::Date);
    EXPECT_EQ(registry.detectNumberFormatType("mm/dd/yyyy"), xlsxcsv::core::NumberFormatType::Date);
    EXPECT_EQ(registry.detectNumberFormatType("d-mmm-yy"), xlsxcsv::core::NumberFormatType::Date);
    
    // Test Time formats
    EXPECT_EQ(registry.detectNumberFormatType("h:mm:ss"), xlsxcsv::core::NumberFormatType::Time);
    EXPECT_EQ(registry.detectNumberFormatType("hh:mm AM/PM"), xlsxcsv::core::NumberFormatType::Time);
    
    // Test DateTime formats
    EXPECT_EQ(registry.detectNumberFormatType("mm/dd/yyyy h:mm"), xlsxcsv::core::NumberFormatType::DateTime);
    EXPECT_EQ(registry.detectNumberFormatType("yyyy-mm-dd hh:mm:ss"), xlsxcsv::core::NumberFormatType::DateTime);
    
    // Test Percentage
    EXPECT_EQ(registry.detectNumberFormatType("0%"), xlsxcsv::core::NumberFormatType::Percentage);
    EXPECT_EQ(registry.detectNumberFormatType("0.00%"), xlsxcsv::core::NumberFormatType::Percentage);
    
    // Test Currency
    EXPECT_EQ(registry.detectNumberFormatType("$#,##0.00"), xlsxcsv::core::NumberFormatType::Currency);
    EXPECT_EQ(registry.detectNumberFormatType("¤#,##0.00"), xlsxcsv::core::NumberFormatType::Currency);
    
    // Test Scientific
    EXPECT_EQ(registry.detectNumberFormatType("0.00E+00"), xlsxcsv::core::NumberFormatType::Scientific);
    EXPECT_EQ(registry.detectNumberFormatType("0.0e-00"), xlsxcsv::core::NumberFormatType::Scientific);
    
    // Test Fraction
    EXPECT_EQ(registry.detectNumberFormatType("# ?/?"), xlsxcsv::core::NumberFormatType::Fraction);
    EXPECT_EQ(registry.detectNumberFormatType("# ?\x3f/?\x3f"), xlsxcsv::core::NumberFormatType::Fraction);
    
    // Test Text
    EXPECT_EQ(registry.detectNumberFormatType("@"), xlsxcsv::core::NumberFormatType::Text);
    
    // Test Decimal
    EXPECT_EQ(registry.detectNumberFormatType("0.00"), xlsxcsv::core::NumberFormatType::Decimal);
    EXPECT_EQ(registry.detectNumberFormatType("#,##0.00"), xlsxcsv::core::NumberFormatType::Decimal);
    
    // Test Integer
    EXPECT_EQ(registry.detectNumberFormatType("0"), xlsxcsv::core::NumberFormatType::Integer);
    EXPECT_EQ(registry.detectNumberFormatType("#,##0"), xlsxcsv::core::NumberFormatType::Integer);
}

TEST_F(StylesRegistryTest, DateTimeFormatDetection) {
    if (!fs::exists(dateFormatsXlsxPath)) {
        GTEST_SKIP() << "Date formats test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(dateFormatsXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    
    // Test built-in date/time formats
    EXPECT_TRUE(registry.isDateTimeFormat(14));  // mm-dd-yy
    EXPECT_TRUE(registry.isDateTimeFormat(18));  // h:mm AM/PM
    EXPECT_TRUE(registry.isDateTimeFormat(22));  // m/d/yy h:mm
    
    // Test custom date/time formats
    EXPECT_TRUE(registry.isDateTimeFormat(170)); // yyyy-mm-dd hh:mm:ss
    EXPECT_TRUE(registry.isDateTimeFormat(171)); // [$-en-US]mmmm d, yyyy
    EXPECT_TRUE(registry.isDateTimeFormat(172)); // [h]:mm:ss
    
    // Test non-date formats
    EXPECT_FALSE(registry.isDateTimeFormat(0));  // General
    EXPECT_FALSE(registry.isDateTimeFormat(1));  // 0
    EXPECT_FALSE(registry.isDateTimeFormat(9));  // 0%
    
    // Test format code detection
    EXPECT_TRUE(registry.isDateTimeFormat("yyyy-mm-dd"));
    EXPECT_TRUE(registry.isDateTimeFormat("h:mm:ss"));
    EXPECT_TRUE(registry.isDateTimeFormat("mm/dd/yyyy h:mm"));
    EXPECT_FALSE(registry.isDateTimeFormat("0.00"));
    EXPECT_FALSE(registry.isDateTimeFormat("General"));
}

TEST_F(StylesRegistryTest, ComplexStylesParsing) {
    if (!fs::exists(complexStylesXlsxPath)) {
        GTEST_SKIP() << "Complex styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(complexStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    
    EXPECT_TRUE(registry.isOpen());
    EXPECT_EQ(registry.getStyleCount(), 6); // 6 cellXfs
    EXPECT_EQ(registry.getNumberFormatCount(), 5); // 5 custom numFmts
    
    // Test complex number formats
    auto dateFormat = registry.getNumberFormat(164);
    ASSERT_TRUE(dateFormat.has_value());
    EXPECT_EQ(dateFormat->type, xlsxcsv::core::NumberFormatType::Date);
    
    auto timeFormat = registry.getNumberFormat(165);
    ASSERT_TRUE(timeFormat.has_value());
    EXPECT_EQ(timeFormat->type, xlsxcsv::core::NumberFormatType::Time);
    
    auto datetimeFormat = registry.getNumberFormat(166);
    ASSERT_TRUE(datetimeFormat.has_value());
    EXPECT_EQ(datetimeFormat->type, xlsxcsv::core::NumberFormatType::DateTime);
    
    auto currencyFormat = registry.getNumberFormat(167);
    ASSERT_TRUE(currencyFormat.has_value());
    EXPECT_EQ(currencyFormat->type, xlsxcsv::core::NumberFormatType::Currency);
    
    auto scientificFormat = registry.getNumberFormat(168);
    ASSERT_TRUE(scientificFormat.has_value());
    EXPECT_EQ(scientificFormat->type, xlsxcsv::core::NumberFormatType::Scientific);
    
    // Test various font styles
    auto style1 = registry.getCellStyle(1);
    ASSERT_TRUE(style1.has_value());
    EXPECT_EQ(style1->font.name, "Times New Roman");
    EXPECT_EQ(style1->font.size, 14.0);
    EXPECT_TRUE(style1->font.bold);
    EXPECT_EQ(style1->font.color, "FF800080");
    
    auto style2 = registry.getCellStyle(2);
    ASSERT_TRUE(style2.has_value());
    EXPECT_EQ(style2->font.name, "Courier New");
    EXPECT_EQ(style2->font.size, 10.0);
    EXPECT_TRUE(style2->font.underline);
}

TEST_F(StylesRegistryTest, CloseRegistry) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    registry.parse(package);
    EXPECT_TRUE(registry.isOpen());
    
    registry.close();
    EXPECT_FALSE(registry.isOpen());
    EXPECT_EQ(registry.getStyleCount(), 0);
    EXPECT_EQ(registry.getNumberFormatCount(), 0);
    
    // Test that methods return nullopt after close
    auto style = registry.getCellStyle(0);
    EXPECT_FALSE(style.has_value());
    
    auto format = registry.getNumberFormat(0);
    EXPECT_FALSE(format.has_value());
}

TEST_F(StylesRegistryTest, MoveConstruction) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry1;
    registry1.parse(package);
    EXPECT_TRUE(registry1.isOpen());
    
    xlsxcsv::core::StylesRegistry registry2 = std::move(registry1);
    EXPECT_TRUE(registry2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(StylesRegistryTest, MoveAssignment) {
    if (!fs::exists(basicStylesXlsxPath)) {
        GTEST_SKIP() << "Basic styles test XLSX file could not be created";
    }
    
    xlsxcsv::core::OpcPackage package;
    package.open(basicStylesXlsxPath.string());
    
    xlsxcsv::core::StylesRegistry registry1;
    registry1.parse(package);
    EXPECT_TRUE(registry1.isOpen());
    
    xlsxcsv::core::StylesRegistry registry2;
    registry2 = std::move(registry1);
    EXPECT_TRUE(registry2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(StylesRegistryTest, MissingStylesFile) {
    // Create a simple XLSX without styles.xml
    fs::path noStylesPath = testDir / "no_styles.xlsx";
    
    std::string zipCmd = "cd \"" + testDir.string() + "\" && zip -q \"" + 
                        noStylesPath.filename().string() + "\" ";
    
    // Create minimal XLSX structure without styles.xml
    fs::create_directories(testDir / "xl");
    fs::create_directories(testDir / "_rels");
    
    std::ofstream(testDir / "[Content_Types].xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>)";

    std::ofstream(testDir / "_rels" / ".rels") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>)";

    std::ofstream(testDir / "xl" / "workbook.xml") << 
R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
    <sheets>
        <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
    </sheets>
</workbook>)";

    // Create ZIP
    zipCmd += "[Content_Types].xml _rels xl";
    system(zipCmd.c_str());
    
    // Clean up temp files
    fs::remove_all(testDir / "[Content_Types].xml");
    fs::remove_all(testDir / "_rels");
    fs::remove_all(testDir / "xl");
    
    // Test that missing styles.xml throws error
    xlsxcsv::core::OpcPackage package;
    package.open(noStylesPath.string());
    
    xlsxcsv::core::StylesRegistry registry;
    EXPECT_THROW(registry.parse(package), xlsxcsv::core::XlsxError);
}