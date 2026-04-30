// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for the MS-XLSB function-id mapping table.

#include <cstdint>
#include <string_view>

#include "gtest/gtest.h"
#include "io/xlsb/func_id_table.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

TEST(XlsbFuncId, TableIsSortedById) {
  for (std::size_t i = 1; i < kXlsbFuncEntryCount; ++i) {
    EXPECT_LT(kXlsbFuncEntries[i - 1].id, kXlsbFuncEntries[i].id) << "row " << i << " out of order";
  }
}

TEST(XlsbFuncId, KnownIdsResolveToExpectedNames) {
  struct Pair {
    std::uint16_t id;
    const char* name;
  };
  // Spot-checks across the classic id range. These ids are documented
  // in [MS-XLS] §2.5.198.17 (Cetab) and [MS-XLSB] §2.5.97.74.
  const Pair kCases[] = {
      {1, "IF"},      {4, "SUM"},           {7, "MAX"},        {10, "NA"},        {32, "LEN"},
      {64, "MATCH"},  {100, "CHOOSE"},      {102, "VLOOKUP"},  {148, "INDIRECT"}, {169, "COUNTA"},
      {221, "TODAY"}, {336, "CONCATENATE"}, {344, "SUBTOTAL"}, {346, "COUNTIF"},
  };
  for (const Pair& tc : kCases) {
    const XlsbFuncEntry* e = lookup_func_by_id(tc.id);
    ASSERT_NE(e, nullptr) << "id " << tc.id << " missing from table";
    EXPECT_STREQ(e->name, tc.name) << "id " << tc.id;
  }
}

TEST(XlsbFuncId, NameLookupRoundTripsIds) {
  // Each name we look up by id should also resolve back to the same id
  // by name (case-insensitive).
  const std::uint16_t kIds[] = {1, 4, 32, 64, 100, 102, 148, 169, 336, 346};
  for (std::uint16_t id : kIds) {
    const XlsbFuncEntry* by_id = lookup_func_by_id(id);
    ASSERT_NE(by_id, nullptr) << "id " << id;
    const XlsbFuncEntry* by_name = lookup_func_by_name(by_id->name);
    ASSERT_NE(by_name, nullptr) << "name " << by_id->name;
    EXPECT_EQ(by_name->id, id);
  }

  // Lower-case entry must round-trip via the case-insensitive scan.
  const XlsbFuncEntry* sum_lower = lookup_func_by_name("sum");
  ASSERT_NE(sum_lower, nullptr);
  EXPECT_EQ(sum_lower->id, 4);
}

TEST(XlsbFuncId, UnknownIdReturnsNull) {
  // 0xFFFE / 0xFFFF are sentinels we never emit; lookup must report
  // "unknown".
  EXPECT_EQ(lookup_func_by_id(0xFFFEU), nullptr);
  EXPECT_EQ(lookup_func_by_id(0xFFFFU), nullptr);
  // A gap inside the range — id 53 is reserved (LINEST family at 49-52
  // followed by 56 PV in our table) and not in the mapping.
  EXPECT_EQ(lookup_func_by_id(53U), nullptr);
}

TEST(XlsbFuncId, UnknownNameReturnsNull) {
  EXPECT_EQ(lookup_func_by_name(""), nullptr);
  EXPECT_EQ(lookup_func_by_name("DEFINITELY_NOT_AN_EXCEL_FUNCTION"), nullptr);
}

TEST(XlsbFuncId, ArityMetadataMatchesExpectations) {
  // IF: fixed 2..3 args.
  const XlsbFuncEntry* if_e = lookup_func_by_id(1);
  ASSERT_NE(if_e, nullptr);
  EXPECT_EQ(if_e->arg_min, 2);
  EXPECT_EQ(if_e->arg_max, 3);
  EXPECT_FALSE(if_e->variadic);

  // SUM: variadic, 0 to "any".
  const XlsbFuncEntry* sum_e = lookup_func_by_id(4);
  ASSERT_NE(sum_e, nullptr);
  EXPECT_EQ(sum_e->arg_min, 0);
  EXPECT_TRUE(sum_e->variadic);

  // VLOOKUP: 3..4, marked variadic because the optional arg makes it
  // PtgFuncVar in the wire format.
  const XlsbFuncEntry* vlookup_e = lookup_func_by_id(102);
  ASSERT_NE(vlookup_e, nullptr);
  EXPECT_EQ(vlookup_e->arg_min, 3);
  EXPECT_EQ(vlookup_e->arg_max, 4);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
