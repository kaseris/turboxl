#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"

using namespace xlsxcsv::core;

class Phase5FunctionalityTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

// Test CellCoordinate functionality
TEST_F(Phase5FunctionalityTest, CellCoordinateParsing) {
    // Test basic coordinates
    auto coord = CellCoordinate::fromReference("A1");
    ASSERT_TRUE(coord.has_value());
    EXPECT_EQ(coord->row, 1);
    EXPECT_EQ(coord->column, 1);
    EXPECT_EQ(coord->toReference(), "A1");
    
    // Test multi-letter columns
    coord = CellCoordinate::fromReference("AA1");
    ASSERT_TRUE(coord.has_value());
    EXPECT_EQ(coord->row, 1);
    EXPECT_EQ(coord->column, 27); // AA = 27
    EXPECT_EQ(coord->toReference(), "AA1");
    
    // Test larger coordinates
    coord = CellCoordinate::fromReference("BC42");
    ASSERT_TRUE(coord.has_value());
    EXPECT_EQ(coord->row, 42);
    EXPECT_EQ(coord->column, 55); // BC = 55
    EXPECT_EQ(coord->toReference(), "BC42");
    
    // Test invalid coordinates
    EXPECT_FALSE(CellCoordinate::fromReference("").has_value());
    EXPECT_FALSE(CellCoordinate::fromReference("1A").has_value());
    EXPECT_FALSE(CellCoordinate::fromReference("A0").has_value());
    EXPECT_FALSE(CellCoordinate::fromReference("A").has_value());
    EXPECT_FALSE(CellCoordinate::fromReference("1").has_value());
}

TEST_F(Phase5FunctionalityTest, CellCoordinateConversion) {
    // Test column conversion edge cases
    EXPECT_EQ(CellCoordinate::fromReference("Z1")->column, 26);
    EXPECT_EQ(CellCoordinate::fromReference("AA1")->column, 27);
    EXPECT_EQ(CellCoordinate::fromReference("AB1")->column, 28);
    EXPECT_EQ(CellCoordinate::fromReference("AZ1")->column, 52);
    EXPECT_EQ(CellCoordinate::fromReference("BA1")->column, 53);
    
    // Test reverse conversion
    CellCoordinate coord;
    coord.row = 1;
    coord.column = 26;
    EXPECT_EQ(coord.toReference(), "Z1");
    
    coord.column = 27;
    EXPECT_EQ(coord.toReference(), "AA1");
    
    coord.column = 52;
    EXPECT_EQ(coord.toReference(), "AZ1");
}

TEST_F(Phase5FunctionalityTest, CellDataHelperMethods) {
    CellData cell;
    
    // Test empty cell
    EXPECT_TRUE(cell.isEmpty());
    EXPECT_FALSE(cell.isBoolean());
    EXPECT_FALSE(cell.isNumber());
    EXPECT_FALSE(cell.isString());
    
    // Test boolean cell
    cell.value = true;
    cell.type = CellType::Boolean;
    EXPECT_FALSE(cell.isEmpty());
    EXPECT_TRUE(cell.isBoolean());
    EXPECT_TRUE(cell.getBoolean());
    
    // Test numeric cell
    cell.value = 42.5;
    cell.type = CellType::Number;
    EXPECT_TRUE(cell.isNumber());
    EXPECT_DOUBLE_EQ(cell.getNumber(), 42.5);
    
    // Test string cell
    cell.value = std::string("Hello");
    cell.type = CellType::String;
    EXPECT_TRUE(cell.isString());
    EXPECT_EQ(cell.getString(), "Hello");
    
    // Test shared string index
    cell.value = 5;
    cell.type = CellType::SharedString;
    EXPECT_TRUE(cell.isSharedStringIndex());
    EXPECT_EQ(cell.getSharedStringIndex(), 5);
}

TEST_F(Phase5FunctionalityTest, RowDataFunctionality) {
    RowData row;
    row.rowNumber = 3;
    
    // Add cells
    CellData cell1;
    cell1.coordinate.row = 3;
    cell1.coordinate.column = 1; // A3
    cell1.value = std::string("First");
    
    CellData cell2;
    cell2.coordinate.row = 3;
    cell2.coordinate.column = 3; // C3
    cell2.value = 42.0;
    
    row.cells.push_back(cell1);
    row.cells.push_back(cell2);
    
    // Test finding cells
    const CellData* foundCell = row.findCell(1);
    ASSERT_NE(foundCell, nullptr);
    EXPECT_EQ(foundCell->getString(), "First");
    
    foundCell = row.findCell(3);
    ASSERT_NE(foundCell, nullptr);
    EXPECT_DOUBLE_EQ(foundCell->getNumber(), 42.0);
    
    // Test missing cell
    foundCell = row.findCell(2);
    EXPECT_EQ(foundCell, nullptr);
}

TEST_F(Phase5FunctionalityTest, CsvRowCollectorBasic) {
    CsvRowCollector collector;
    
    // Test empty result
    EXPECT_EQ(collector.getRowCount(), 0);
    EXPECT_EQ(collector.getCsvString(), "");
    EXPECT_TRUE(collector.getErrors().empty());
    
    // Create a simple row
    RowData row;
    row.rowNumber = 1;
    
    CellData cell1;
    cell1.coordinate.row = 1;
    cell1.coordinate.column = 1; // A1
    cell1.value = std::string("Hello");
    cell1.type = CellType::String;
    
    CellData cell2;
    cell2.coordinate.row = 1;
    cell2.coordinate.column = 2; // B1
    cell2.value = 42.0;
    cell2.type = CellType::Number;
    
    row.cells.push_back(cell1);
    row.cells.push_back(cell2);
    
    // Handle the row
    collector.handleRow(row);
    
    // Check results
    EXPECT_EQ(collector.getRowCount(), 1);
    std::string csvResult = collector.getCsvString();
    EXPECT_EQ(csvResult, "Hello,42\n");
}

TEST_F(Phase5FunctionalityTest, CsvRowCollectorEscaping) {
    CsvRowCollector collector;
    
    RowData row;
    row.rowNumber = 1;
    
    // Cell with comma (needs quoting)
    CellData cell1;
    cell1.coordinate.row = 1;
    cell1.coordinate.column = 1;
    cell1.value = std::string("Hello, World");
    cell1.type = CellType::String;
    
    // Cell with quotes (needs escaping)
    CellData cell2;
    cell2.coordinate.row = 1;
    cell2.coordinate.column = 2;
    cell2.value = std::string("Say \"Hello\"");
    cell2.type = CellType::String;
    
    row.cells.push_back(cell1);
    row.cells.push_back(cell2);
    
    collector.handleRow(row);
    
    std::string csvResult = collector.getCsvString();
    EXPECT_EQ(csvResult, "\"Hello, World\",\"Say \"\"Hello\"\"\"\n");
}

TEST_F(Phase5FunctionalityTest, CsvRowCollectorSparseData) {
    CsvRowCollector collector;
    
    RowData row;
    row.rowNumber = 1;
    
    // Add cells in columns 1 and 4 (A1 and D1), skipping B1 and C1
    CellData cell1;
    cell1.coordinate.row = 1;
    cell1.coordinate.column = 1; // A1
    cell1.value = std::string("First");
    cell1.type = CellType::String;
    
    CellData cell4;
    cell4.coordinate.row = 1;
    cell4.coordinate.column = 4; // D1
    cell4.value = std::string("Fourth");
    cell4.type = CellType::String;
    
    row.cells.push_back(cell1);
    row.cells.push_back(cell4);
    
    collector.handleRow(row);
    
    std::string csvResult = collector.getCsvString();
    EXPECT_EQ(csvResult, "First,,,Fourth\n");
}

TEST_F(Phase5FunctionalityTest, CsvRowCollectorErrorHandling) {
    CsvRowCollector collector;
    
    // Test error handling
    collector.handleError("Test error message");
    
    EXPECT_EQ(collector.getErrors().size(), 1);
    EXPECT_EQ(collector.getErrors()[0], "Test error message");
    
    // Multiple errors
    collector.handleError("Second error");
    EXPECT_EQ(collector.getErrors().size(), 2);
    EXPECT_EQ(collector.getErrors()[1], "Second error");
}