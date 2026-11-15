// The published passthrough list of an .xlsb names only the parts the reader kept without modelling them; the
// writer's retained originals of modelled parts (styles, a lifted PivotTable group) stay off it.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "xlsb_twin_compare.h"
#include "zip_repack_helpers.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

namespace tc = test;

const char kBook[] = "/excel/win/pivot_days_end/pivot_days_end.xlsb";
const char kDefinition[] = "xl/pivotCache/pivotCacheDefinition1.bin";
const char kDefinitionRels[] = "xl/pivotCache/_rels/pivotCacheDefinition1.bin.rels";

struct Loaded {
  fm_workbook_t* wb = nullptr;
  ~Loaded() { fm_workbook_destroy(wb); }
};

std::set<std::string> Listed(const fm_workbook_t* wb) {
  std::size_t count = 0;
  EXPECT_EQ(fm_workbook_get_passthrough_count(wb, &count), 0);
  std::set<std::string> out;
  for (std::size_t i = 0; i < count; ++i) {
    const char* path = nullptr;
    EXPECT_EQ(fm_workbook_get_passthrough_at(wb, i, &path), 0);
    out.insert(path);
  }
  const char* past = nullptr;
  EXPECT_EQ(fm_workbook_get_passthrough_at(wb, count, &past), FM_STATUS_INDEX_OUT_OF_RANGE);
  return out;
}

bool Has(const fm_workbook_t* wb, const std::string& path) {
  int32_t present = -1;
  EXPECT_EQ(fm_workbook_has_passthrough(wb, path.c_str(), &present), 0);
  return present == 1;
}

std::size_t PivotCount(const fm_workbook_t* wb) {
  std::size_t sheets = 0;
  EXPECT_EQ(fm_workbook_get_sheet_count(wb, &sheets), 0);
  std::size_t total = 0;
  for (std::size_t s = 0; s < sheets; ++s) {
    std::size_t n = 0;
    EXPECT_EQ(fm_sheet_get_pivot_count(wb, s, &n), 0);
    total += n;
  }
  return total;
}

std::set<std::string> PivotParts(const std::set<std::string>& listed) {
  std::set<std::string> out;
  for (const std::string& path : listed) {
    if (path.find("pivot") != std::string::npos) {
      out.insert(path);
    }
  }
  return out;
}

TEST(XlsbPassthroughListing, ModelledPartsAreNotListed) {
  const tc::Bytes source = tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + kBook);
  Loaded book;
  ASSERT_EQ(fm_workbook_load(source.data(), source.size(), nullptr, &book.wb), 0) << fm_get_last_error_message();
  ASSERT_EQ(PivotCount(book.wb), 3U);
  const std::set<std::string> listed = Listed(book.wb);
  EXPECT_EQ(listed.count("xl/styles.bin"), 0U);
  EXPECT_FALSE(Has(book.wb, "xl/styles.bin"));
  EXPECT_TRUE(PivotParts(listed).empty());
  EXPECT_FALSE(Has(book.wb, kDefinition));
}

// One group the reader leaves as stored beside two it models: only that group's parts and rels are listed.
TEST(XlsbPassthroughListing, OnlyTheUnmodelledGroupIsListed) {
  const tc::Bytes source = tc::ReadBytes(std::string(FORMULON_FIXTURES_DIR) + kBook);
  const std::optional<tc::Bytes> definition = tc::ReadPart(source, kDefinition);
  const std::optional<tc::Bytes> rels_bytes = tc::ReadPart(source, kDefinitionRels);
  ASSERT_TRUE(definition.has_value() && rels_bytes.has_value());
  // Records related under rId2 instead of rId1 keep the group out of the model.
  const tc::Bytes wide1{'r', 0, 'I', 0, 'd', 0, '1', 0};
  tc::Bytes renamed_definition = *definition;
  const auto at = std::search(renamed_definition.begin(), renamed_definition.end(), wide1.begin(), wide1.end());
  ASSERT_NE(at, renamed_definition.end());
  *(at + 6) = '2';
  std::string rels(rels_bytes->begin(), rels_bytes->end());
  const std::size_t id = rels.find("Id=\"rId1\"");
  ASSERT_NE(id, std::string::npos);
  rels.replace(id, 10U, "Id=\"rId2\"");
  const tc::Bytes mixed = RepackWithEntries(
      source, {{kDefinition, renamed_definition}, {kDefinitionRels, tc::Bytes(rels.begin(), rels.end())}});

  // The table drawing from that cache names its definition in its own rels part.
  std::string table;
  for (const std::string& name : tc::ListParts(mixed)) {
    if (name.rfind("xl/pivotTables/_rels/", 0) != 0) {
      continue;
    }
    const tc::Bytes bytes = tc::ReadPart(mixed, name).value_or(tc::Bytes{});
    if (std::string(bytes.begin(), bytes.end()).find("pivotCacheDefinition1.bin") != std::string::npos) {
      table = "xl/pivotTables/" + name.substr(std::string("xl/pivotTables/_rels/").size());
      table.resize(table.size() - std::string(".rels").size());
    }
  }
  ASSERT_FALSE(table.empty());

  Loaded book;
  ASSERT_EQ(fm_workbook_load(mixed.data(), mixed.size(), nullptr, &book.wb), 0) << fm_get_last_error_message();
  ASSERT_EQ(PivotCount(book.wb), 2U);
  const std::set<std::string> listed = Listed(book.wb);
  const std::string table_rels =
      "xl/pivotTables/_rels/" + table.substr(std::string("xl/pivotTables/").size()) + ".rels";
  EXPECT_EQ(PivotParts(listed), (std::set<std::string>{table, table_rels, kDefinition, kDefinitionRels,
                                                       "xl/pivotCache/pivotCacheRecords1.bin"}));
  EXPECT_TRUE(Has(book.wb, kDefinition));
  EXPECT_FALSE(Has(book.wb, "xl/pivotCache/pivotCacheDefinition2.bin"));
  EXPECT_FALSE(Has(book.wb, "xl/styles.bin"));
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
