#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include <fstream>
#include <filesystem>

namespace fs = std::filesystem;

class ZipReaderTest : public ::testing::Test {
protected:
    void SetUp() override {
        // Create test directory
        testDir = fs::temp_directory_path() / "turboxl_test";
        fs::create_directories(testDir);
        
        // Create a simple test ZIP file
        createTestZipFile();
    }
    
    void TearDown() override {
        // Clean up test files
        if (fs::exists(testDir)) {
            fs::remove_all(testDir);
        }
    }
    
    void createTestZipFile() {
        testZipPath = testDir / "test.zip";
        
        // Create a simple test file to zip
        auto textFile = testDir / "test.txt";
        std::ofstream file(textFile);
        file << "Hello, World!\nThis is a test file.";
        file.close();
        
        // Use system zip command to create test zip
        std::string cmd = "cd " + testDir.string() + " && zip -q test.zip test.txt";
        system(cmd.c_str());
    }
    
    fs::path testDir;
    fs::path testZipPath;
};

TEST_F(ZipReaderTest, DefaultConstruction) {
    xlsxcsv::core::ZipReader reader;
    EXPECT_FALSE(reader.isOpen());
    
    auto limits = reader.getSecurityLimits();
    EXPECT_EQ(limits.maxEntries, 10000u);
    EXPECT_EQ(limits.maxEntrySize, 256u * 1024 * 1024);
    EXPECT_EQ(limits.maxTotalUncompressed, 2ULL * 1024 * 1024 * 1024);
}

TEST_F(ZipReaderTest, CustomSecurityLimits) {
    xlsxcsv::core::ZipSecurityLimits limits;
    limits.maxEntries = 1000;
    limits.maxEntrySize = 1024 * 1024;
    limits.maxTotalUncompressed = 100 * 1024 * 1024;
    
    xlsxcsv::core::ZipReader reader(limits);
    auto actualLimits = reader.getSecurityLimits();
    
    EXPECT_EQ(actualLimits.maxEntries, 1000u);
    EXPECT_EQ(actualLimits.maxEntrySize, 1024u * 1024);
    EXPECT_EQ(actualLimits.maxTotalUncompressed, 100u * 1024 * 1024);
}

TEST_F(ZipReaderTest, OpenValidZipFile) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    EXPECT_NO_THROW(reader.open(testZipPath.string()));
    EXPECT_TRUE(reader.isOpen());
}

TEST_F(ZipReaderTest, OpenNonExistentFile) {
    xlsxcsv::core::ZipReader reader;
    EXPECT_THROW(reader.open("nonexistent.zip"), xlsxcsv::core::XlsxError);
    EXPECT_FALSE(reader.isOpen());
}

TEST_F(ZipReaderTest, OpenInvalidZipFile) {
    // Create an invalid ZIP file
    auto invalidZip = testDir / "invalid.zip";
    std::ofstream file(invalidZip);
    file << "This is not a ZIP file";
    file.close();
    
    xlsxcsv::core::ZipReader reader;
    EXPECT_THROW(reader.open(invalidZip.string()), xlsxcsv::core::XlsxError);
    EXPECT_FALSE(reader.isOpen());
}

TEST_F(ZipReaderTest, ListEntries) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    
    auto entries = reader.listEntries();
    EXPECT_FALSE(entries.empty());
    
    bool foundTestFile = false;
    for (const auto& entry : entries) {
        if (entry.path == "test.txt") {
            foundTestFile = true;
            EXPECT_GT(entry.compressedSize, 0u);
            EXPECT_GT(entry.uncompressedSize, 0u);
            EXPECT_FALSE(entry.isEncrypted);
            break;
        }
    }
    EXPECT_TRUE(foundTestFile);
}

TEST_F(ZipReaderTest, HasEntry) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    
    EXPECT_TRUE(reader.hasEntry("test.txt"));
    EXPECT_FALSE(reader.hasEntry("nonexistent.txt"));
}

TEST_F(ZipReaderTest, ReadEntry) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    
    auto data = reader.readEntry("test.txt");
    EXPECT_FALSE(data.empty());
    
    std::string content(data.begin(), data.end());
    EXPECT_EQ(content, "Hello, World!\nThis is a test file.");
}

TEST_F(ZipReaderTest, ReadEntryAsString) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    
    auto content = reader.readEntryAsString("test.txt");
    EXPECT_EQ(content, "Hello, World!\nThis is a test file.");
}

TEST_F(ZipReaderTest, ReadNonExistentEntry) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    
    EXPECT_THROW(reader.readEntry("nonexistent.txt"), xlsxcsv::core::XlsxError);
    EXPECT_THROW(reader.readEntryAsString("nonexistent.txt"), xlsxcsv::core::XlsxError);
}

TEST_F(ZipReaderTest, CloseFile) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader;
    reader.open(testZipPath.string());
    EXPECT_TRUE(reader.isOpen());
    
    reader.close();
    EXPECT_FALSE(reader.isOpen());
    
    // Operations should fail after close
    EXPECT_THROW(reader.listEntries(), xlsxcsv::core::XlsxError);
    EXPECT_THROW(reader.hasEntry("test.txt"), xlsxcsv::core::XlsxError);
    EXPECT_THROW(reader.readEntry("test.txt"), xlsxcsv::core::XlsxError);
}

TEST_F(ZipReaderTest, MoveConstruction) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader1;
    reader1.open(testZipPath.string());
    EXPECT_TRUE(reader1.isOpen());
    
    xlsxcsv::core::ZipReader reader2 = std::move(reader1);
    EXPECT_TRUE(reader2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

TEST_F(ZipReaderTest, MoveAssignment) {
    if (!fs::exists(testZipPath)) {
        GTEST_SKIP() << "Test ZIP file could not be created";
    }
    
    xlsxcsv::core::ZipReader reader1;
    reader1.open(testZipPath.string());
    EXPECT_TRUE(reader1.isOpen());
    
    xlsxcsv::core::ZipReader reader2;
    reader2 = std::move(reader1);
    EXPECT_TRUE(reader2.isOpen());
    // Note: Moved-from object is in valid but unspecified state, don't test its state
}

// Security tests
TEST_F(ZipReaderTest, PathTraversalPrevention) {
    // This test would require creating a ZIP with malicious paths
    // For now, we test that the API exists and can be called
    xlsxcsv::core::ZipReader reader;
    
    if (fs::exists(testZipPath)) {
        reader.open(testZipPath.string());
        // Normal path should work
        EXPECT_TRUE(reader.hasEntry("test.txt"));
        // Absolute path should be rejected or normalized
        EXPECT_FALSE(reader.hasEntry("/etc/passwd"));
        // Path traversal should be rejected or normalized  
        EXPECT_FALSE(reader.hasEntry("../../../etc/passwd"));
    }
}

TEST_F(ZipReaderTest, EncryptionDetection) {
    // This is a placeholder for encryption detection tests
    // In a real implementation, we would create encrypted ZIP files
    // and verify they are properly detected and rejected
    xlsxcsv::core::ZipReader reader;
    
    if (fs::exists(testZipPath)) {
        reader.open(testZipPath.string());
        auto entries = reader.listEntries();
        
        for (const auto& entry : entries) {
            // Test files should not be encrypted
            EXPECT_FALSE(entry.isEncrypted);
        }
    }
}
