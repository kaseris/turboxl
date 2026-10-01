#include <gtest/gtest.h>
#include "xlsxcsv/core.hpp"
#include "fixture_helpers.hpp"
#include <filesystem>
#include <string>
#include <vector>

class SharedStringsProviderTest : public ::testing::Test {
protected:
    void SetUp() override {
        testDir = std::filesystem::temp_directory_path() / "turboxl_shared_strings_test";
        std::filesystem::create_directories(testDir);
    }
    void TearDown() override { std::filesystem::remove_all(testDir); }

    // Parses the same sharedStrings.xml through the in-memory fast scan and
    // the libxml2 reader (External storage) and returns both string lists.
    struct Both {
        std::vector<std::string> fast;
        std::vector<std::string> reference;
    };

    Both parseBoth(const std::string& name, const std::string& xml,
                   size_t maxStringLength = 32767) {
        const auto archive = writePackageWithPart(testDir, name, "xl/sharedStrings.xml", xml);
        xlsxcsv::core::OpcPackage package;
        package.open(archive.string());
        Both both;
        for (const bool fast : {true, false}) {
            xlsxcsv::core::SharedStringsConfig config;
            config.mode = fast ? xlsxcsv::core::SharedStringsMode::InMemory
                               : xlsxcsv::core::SharedStringsMode::External;
            config.maxStringLength = maxStringLength;
            xlsxcsv::core::SharedStringsProvider provider(config);
            provider.parse(package);
            auto& output = fast ? both.fast : both.reference;
            for (size_t i = 0; i < provider.getStringCount(); ++i) {
                output.push_back(provider.getString(i));
            }
        }
        return both;
    }

    std::filesystem::path testDir;
};

TEST_F(SharedStringsProviderTest, FastScanMatchesLibxml2Reader) {
    const auto both = parseBoth("plain_and_rich", R"(<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="9">
  <si><t>plain</t></si>
  <si><t xml:space="preserve">  padded  </t></si>
  <si><r><rPr><b/></rPr><t>bold </t></r><r><t>and plain</t></r></si>
  <si><t>R&amp;D &lt;tag&gt; &quot;q&quot; &apos;a&apos; &#233; &#x1F600;</t></si>
  <si/>
  <si><t/></si>
  <si><t><![CDATA[raw <cdata> & text]]></t></si>
  <si><t>base</t><rPh sb="0" eb="1"><t>phonetic</t></rPh><phoneticPr fontId="1"/></si>
  <si><t>line one
line two</t></si>
</sst>)");
    ASSERT_EQ(both.fast.size(), 9u);
    EXPECT_EQ(both.fast, both.reference);
    EXPECT_EQ(both.fast[0], "plain");
    EXPECT_EQ(both.fast[1], "  padded  ");
    EXPECT_EQ(both.fast[2], "bold and plain");
    EXPECT_EQ(both.fast[4], "");
    EXPECT_EQ(both.fast[6], "raw <cdata> & text");
}

TEST_F(SharedStringsProviderTest, FastScanDefersToLibxml2ForUnsupportedInput) {
    // Carriage returns are normalized by libxml2 but not by the scan.
    auto both = parseBoth("crlf", std::string(
        "<?xml version=\"1.0\"?><sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
        "<si><t>a\r\nb</t></si></sst>"));
    EXPECT_EQ(both.fast, both.reference);
    ASSERT_EQ(both.fast.size(), 1u);
    EXPECT_EQ(both.fast[0], "a\nb");

    // Custom entities and extension lists are outside the scan's subset.
    both = parseBoth("dtd", R"(<?xml version="1.0"?>
<!DOCTYPE sst [<!ENTITY co "Acme">]>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>&co; Ltd</t></si></sst>)");
    EXPECT_EQ(both.fast, both.reference);
    ASSERT_EQ(both.fast.size(), 1u);
    EXPECT_EQ(both.fast[0], "Acme Ltd");

    both = parseBoth("extlst", R"(<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>x</t></si><extLst><ext uri="u"><si/></ext></extLst><si><t>y</t></si></sst>)");
    EXPECT_EQ(both.fast, both.reference);
}

TEST_F(SharedStringsProviderTest, FastScanAppliesMaximumStringLength) {
    const auto both = parseBoth("truncate", R"(<?xml version="1.0"?>
<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><si><t>abcdefghij</t></si><si><t>abc</t></si></sst>)", 5);
    EXPECT_EQ(both.fast, both.reference);
    ASSERT_EQ(both.fast.size(), 2u);
    EXPECT_EQ(both.fast[0], "abcde");
    EXPECT_EQ(both.fast[1], "abc");
}

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

TEST_F(SharedStringsProviderTest, InMemoryLookupReturnsStableView) {
    xlsxcsv::core::OpcPackage package;
    package.open(INTEGRATION_XLSX);
    xlsxcsv::core::SharedStringsConfig config;
    config.mode = xlsxcsv::core::SharedStringsMode::InMemory;
    xlsxcsv::core::SharedStringsProvider provider(config);
    provider.parse(package);
    const auto view = provider.tryGetStringView(0);
    ASSERT_TRUE(view.has_value());
    EXPECT_EQ(*view, "caf\xc3\xa9, \"quoted\"");
    EXPECT_EQ(provider.tryGetString(0), std::optional<std::string>(std::string(*view)));
}

TEST_F(SharedStringsProviderTest, ExternalStorageDoesNotExposeView) {
    xlsxcsv::core::OpcPackage package;
    package.open(INTEGRATION_XLSX);
    xlsxcsv::core::SharedStringsConfig config;
    config.mode = xlsxcsv::core::SharedStringsMode::External;
    xlsxcsv::core::SharedStringsProvider provider(config);
    provider.parse(package);
    EXPECT_FALSE(provider.tryGetStringView(0).has_value());
    EXPECT_EQ(provider.getString(0), "caf\xc3\xa9, \"quoted\"");
}

TEST_F(SharedStringsProviderTest, ExternalStorageClosesCleanly) {
    xlsxcsv::core::OpcPackage package;
    package.open(INTEGRATION_XLSX);
    xlsxcsv::core::SharedStringsConfig config;
    config.mode = xlsxcsv::core::SharedStringsMode::External;
    xlsxcsv::core::SharedStringsProvider provider(config);
    provider.parse(package);

    EXPECT_NO_THROW(provider.close());
    EXPECT_FALSE(provider.isOpen());
}
