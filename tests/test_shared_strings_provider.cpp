#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"

class SharedStringsProviderTest : public ::testing::Test {
protected:
    void SetUp() override {}
    void TearDown() override {}
};

TEST_F(SharedStringsProviderTest, BasicConstruction) {
    xlsxcsv::core::SharedStringsProvider provider;
    EXPECT_FALSE(provider.isOpen());
    EXPECT_EQ(provider.getStringCount(), 0);
    EXPECT_FALSE(provider.hasStrings());
}

TEST_F(SharedStringsProviderTest, CustomConfiguration) {
    xlsxcsv::core::SharedStringsConfig config;
    config.mode = xlsxcsv::core::SharedStringsMode::InMemory;
    config.memoryThreshold = 1024;
    config.maxStringLength = 100;
    config.flattenRichText = false;
    
    xlsxcsv::core::SharedStringsProvider provider(config);
    
    EXPECT_EQ(provider.getConfig().mode, xlsxcsv::core::SharedStringsMode::InMemory);
    EXPECT_EQ(provider.getConfig().memoryThreshold, 1024);
    EXPECT_EQ(provider.getConfig().maxStringLength, 100);
    EXPECT_FALSE(provider.getConfig().flattenRichText);
}

TEST_F(SharedStringsProviderTest, ErrorHandling) {
    xlsxcsv::core::SharedStringsProvider provider;
    
    // Test error on invalid index when not open
    EXPECT_THROW(provider.getString(0), xlsxcsv::core::XlsxError);
    
    // Test optional access returns nullopt when not open
    auto invalidStr = provider.tryGetString(0);
    EXPECT_FALSE(invalidStr.has_value());
}

TEST_F(SharedStringsProviderTest, MoveSemantics) {
    xlsxcsv::core::SharedStringsProvider provider1;
    
    // Test move constructor
    auto provider2 = std::move(provider1);
    EXPECT_FALSE(provider2.isOpen());
    EXPECT_EQ(provider2.getStringCount(), 0);
    
    // Test move assignment
    xlsxcsv::core::SharedStringsProvider provider3;
    provider3 = std::move(provider2);
    EXPECT_FALSE(provider3.isOpen());
    EXPECT_EQ(provider3.getStringCount(), 0);
}

TEST_F(SharedStringsProviderTest, ConfigurationPersistence) {
    xlsxcsv::core::SharedStringsConfig config;
    config.mode = xlsxcsv::core::SharedStringsMode::External;
    config.memoryThreshold = 1000;
    config.maxStringLength = 500;
    config.flattenRichText = false;
    
    xlsxcsv::core::SharedStringsProvider provider(config);
    
    const auto& storedConfig = provider.getConfig();
    EXPECT_EQ(storedConfig.mode, xlsxcsv::core::SharedStringsMode::External);
    EXPECT_EQ(storedConfig.memoryThreshold, 1000);
    EXPECT_EQ(storedConfig.maxStringLength, 500);
    EXPECT_FALSE(storedConfig.flattenRichText);
}
