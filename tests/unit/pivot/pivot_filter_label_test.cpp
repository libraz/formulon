// Label, caption, and manual item filters for pivot fields.

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot_evaluator_fixtures.h"
#include "pivot_filter_test_helpers.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using filter_test::build_shared_north_south_cache;
using filter_test::decoded_filter;
using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;
using test::row_index;

// Same three-column shape as `build_basic_cache()`, but the amounts are
// deliberately non-integral / past the 64-bit integer range so a record's
// label goes through the general numeric rendering path rather than the
// integral one.
//
//   Region  Product  Amount
//   ------  -------  ------
//   North   Widget      1.5
//   North   Gadget      2
//   South   Widget     1e20
//   South   Gadget      4
PivotCache build_non_integral_amount_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "Widget", 1.5);
  add("North", "Gadget", 2.0);
  add("South", "Widget", 1e20);
  add("South", "Gadget", 4.0);
  return cache;
}

// Index into `build_shared_region_cache()`'s Region `shared_items` of the
// entry with no value, and of the text entry deliberately spelled like the
// default blank placeholder.
constexpr std::uint32_t kSharedRegionBlank = 1U;
constexpr std::uint32_t kSharedRegionPlaceholderText = 2U;

// Same three-column shape as `build_basic_cache()`, but Region is a *shared*
// field the way a discrete axis column arrives from OOXML: each record stores
// an index into `shared_items`, one of which has no value at all.
//
// The third shared item is a genuine text value spelled exactly like the
// default blank placeholder, so a filter that identified the empty item by its
// rendered label would catch this row too.
//
//   Region     Product  Amount
//   ---------  -------  ------
//   North      Widget    100
//   <no value> Widget     10
//   "(blank)"  Widget      7
PivotCache build_shared_region_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "North"));
  region.shared_items.push_back(Value::blank());
  region.shared_items.push_back(owned_text(cache, std::string{PivotLayoutOptions{}.blank_item_label}));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](std::uint32_t region_index, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), owned_text(cache, product), Value::number(amount)};
    rec.cell_is_index = {true, false, false};
    cache.mutable_records().push_back(std::move(rec));
  };

  add(0U, "Widget", 100.0);
  add(kSharedRegionBlank, "Widget", 10.0);
  add(kSharedRegionPlaceholderText, "Widget", 7.0);
  return cache;
}

// ---------------------------------------------------------------------------
// 8d. PivotFilter (LabelContains / LabelBeginsWith / ValueTop10 / ValueGreaterThan)
// ---------------------------------------------------------------------------

TEST(PivotEvaluator, LabelContainsFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Drop records whose Region label contains "outh" (i.e. South).
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelContains;
  f.value = std::string("outh");
  // PivotFilter pre-aggregation reject: every match drops the record.
  // Wait — LabelContains with payload "outh" *passes* records whose label
  // contains it. To drop South we instead use LabelBeginsWith("North")
  // below; LabelContains here keeps only South.
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 500.0);  // 200 + 300
}

TEST(PivotEvaluator, LabelBeginsWithFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Region";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("Nor");
  table.mutable_active_filters().push_back(f);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionEqualFilterDropsRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_active_filters().push_back(decoded_filter(0, FilterType::LabelEqual, std::string("North")));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
}

TEST(PivotEvaluator, AuthoredCaptionFilterIgnoresCase) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_active_filters().push_back(decoded_filter(0, FilterType::LabelEqual, std::string("nOrTh")));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

TEST(PivotEvaluator, AuthoredCaptionNotContainsFilterDropsMatchingRecords) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_active_filters().push_back(decoded_filter(0, FilterType::LabelNotContains, std::string("outh")));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "North");
}

// The 13 source items of the Mac Excel 365 ja-JP caption-filter probe. Record
// `i` carries the amount `2^i`, so the grand total is the bit mask of the
// records that survive a filter.
constexpr std::array<const char*, 13> kProbeBLabels = {{
    "\xC3\x84pfel",  // 0 Äpfel
    "Stra\xC3\x9F"
    "e",                                                             // 1 Straße
    "STRASSE",                                                       // 2 STRASSE
    "Hello",                                                         // 3 Hello
    "\xEF\xBC\xA4\xEF\xBC\xA5\xEF\xBC\xA6",                          // 4 ＤＥＦ
    "Apfel",                                                         // 5 Apfel
    "xyz\xC3\xA9w",                                                  // 6 xyzéw
    "\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3\xE5\x95\x86\xE4\xBA\x8B",  // 7 ＡＢＣ商事
    "abc",                                                           // 8 abc
    "\xEF\xBD\xB1\xEF\xBD\xB2\xEF\xBD\xB3\xE5\xA0\x82",              // 9 ｱｲｳ堂
    "\xE3\x82\xAB\xE3\x82\xAD\xE3\x82\xAF",                          // 10 カキク
    "\xE3\x82\xAC\xE3\x82\xAE",                                      // 11 ガギ
    "\xCE\x91\xCE\x92\xCE\x93",                                      // 12 ΑΒΓ
}};

PivotCache build_probe_label_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (std::size_t i = 0; i < kProbeBLabels.size(); ++i) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, kProbeBLabels[i]));
    rec.cells.push_back(owned_text(cache, "P"));
    rec.cells.push_back(Value::number(static_cast<double>(1U << i)));
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

// Bit mask of the probe records `table` keeps under mac-365-ja_JP.
unsigned surviving_probe_mask(const PivotTable& table) {
  const PivotCache cache = build_probe_label_cache();
  auto r_or = evaluate(table, cache, pivot_layout_options_for(mac_365_ja_jp_profile()));
  EXPECT_TRUE(static_cast<bool>(r_or));
  if (!r_or || !r_or.value().grand_total.is_number()) {
    return 0U;
  }
  return static_cast<unsigned>(r_or.value().grand_total.as_number());
}

// Probe B rows: begins / ends / contains match the autoFilter pattern against
// the whole label one code point at a time under `=` (case, width and kana
// type fold; diacritics count; no expansion or voicing composition).
TEST(PivotEvaluatorCaptionFilter, MatchesTheExcelMacJaJpCaptionFilters) {
  // `pattern` is the autoFilter criterion Excel saved; the filter's operand is
  // that pattern without the `*` its type affixes.
  struct Case {
    FilterType type;
    const char* pattern;
    unsigned mask;
  };
  const Case cases[] = {
      {FilterType::LabelEqual, "abc\xE5\x95\x86\xE4\xBA\x8B", 0x80U},                  // abc商事
      {FilterType::LabelEqual, "Strasse", 0x6U},                                       // Strasse
      {FilterType::LabelEqual, "\xC3\xA4pfel", 0x1U},                                  // äpfel
      {FilterType::LabelBeginsWith, "hEL*", 0x8U},                                     // hEL*
      {FilterType::LabelBeginsWith, "\xC3\xA4*", 0x1U},                                // ä*
      {FilterType::LabelBeginsWith, "\xC3\x84*", 0x1U},                                // Ä*
      {FilterType::LabelBeginsWith, "a*", 0x1A0U},                                     // a*
      {FilterType::LabelBeginsWith, "abc*", 0x180U},                                   // abc*
      {FilterType::LabelBeginsWith, "\xEF\xBC\xA1\xEF\xBC\xA2\xEF\xBC\xA3*", 0x180U},  // ＡＢＣ*
      {FilterType::LabelBeginsWith, "\xE3\x82\xA2\xE3\x82\xA4*", 0x200U},              // アイ*
      {FilterType::LabelBeginsWith, "\xE3\x81\x82\xE3\x81\x84*", 0x200U},              // あい*
      {FilterType::LabelBeginsWith, "\xE3\x81\x8B\xE3\x81\x8D*", 0x400U},              // かき*
      {FilterType::LabelBeginsWith, "\xEF\xBD\xB6\xEF\xBD\xB7*", 0x400U},              // ｶｷ*
      {FilterType::LabelBeginsWith, "\xEF\xBD\xB6\xEF\xBE\x9E*", 0x0U},                // ｶﾞ*
      {FilterType::LabelBeginsWith, "Strass*", 0x4U},                                  // Strass*
      {FilterType::LabelBeginsWith, "STRA\xC3\x9F*", 0x2U},                            // STRAß*
      {FilterType::LabelContains, "*LL*", 0x8U},                                       // *LL*
      {FilterType::LabelContains, "*\xC3\x89*", 0x40U},                                // *É*
      {FilterType::LabelContains, "*e*", 0x3FU},                                       // *e*
      {FilterType::LabelContains, "*\xCE\xB2*", 0x1000U},                              // *β*
      {FilterType::LabelContains, "*ss*", 0x4U},                                       // *ss*
      {FilterType::LabelContains, "*\xC3\x9F*", 0x2U},                                 // *ß*
      {FilterType::LabelContains, "*\xE3\x81\x84\xE3\x81\x86*", 0x200U},               // *いう*
      {FilterType::LabelEndsWith, "*ef", 0x10U},                                       // *ef
      {FilterType::LabelEndsWith, "*SSE", 0x4U},                                       // *SSE
      {FilterType::LabelEndsWith,
       "*\xC3\x9F"
       "e",
       0x2U},                                                            // *ße
      {FilterType::LabelEndsWith, "*\xE3\x81\x8D\xE3\x81\x8F", 0x400U},  // *きく
      {FilterType::LabelNotContains, "*\xC3\xA4*", 0x1FFEU},             // *ä*
      {FilterType::LabelNotBeginsWith, "abc*", 0x1E7FU},                 // abc*
      {FilterType::LabelContains, "***", 0x1FFFU},                       // ***
      {FilterType::LabelBeginsWith, "?b*", 0x180U},                      // ?b*
      {FilterType::LabelBeginsWith, "Hel*", 0x8U},                       // Hel*
      {FilterType::LabelContains, "*~**", 0x0U},                         // *~**
      {FilterType::LabelBeginsWith, "~?b*", 0x0U},                       // ~?b*
  };
  for (const Case& c : cases) {
    std::string operand(c.pattern);
    if (c.type != FilterType::LabelEqual) {
      const bool begins = c.type == FilterType::LabelBeginsWith || c.type == FilterType::LabelNotBeginsWith;
      const bool ends = c.type == FilterType::LabelEndsWith;
      operand = operand.substr(begins ? 0U : 1U, operand.size() - (ends || begins ? 1U : 2U));
    }
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    table.mutable_active_filters().push_back(decoded_filter(0, c.type, operand));
    EXPECT_EQ(surviving_probe_mask(table), c.mask) << c.pattern;
  }
}

// `PivotFilter` builds the same pattern from its value, whose own `*` / `?` /
// `~` stay live.
TEST(PivotEvaluatorCaptionFilter, LabelFilterUsesTheCaptionMatcher) {
  struct Case {
    FilterType type;
    const char* value;
    unsigned mask;
  };
  const Case cases[] = {
      {FilterType::LabelBeginsWith, "\xC3\xA4", 0x1U},                    // ä
      {FilterType::LabelBeginsWith, "\xE3\x81\x82\xE3\x81\x84", 0x200U},  // あい
      {FilterType::LabelContains, "\xCE\xB2", 0x1000U},                   // β
      {FilterType::LabelContains, "ss", 0x4U},
      {FilterType::LabelBeginsWith, "?b", 0x180U},
      {FilterType::LabelContains, "*", 0x1FFFU},
      {FilterType::LabelContains, "~*", 0x0U},
  };
  for (const Case& c : cases) {
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    PivotFilter f;
    f.axis = PivotAxis::Row;
    f.field_name = "Region";
    f.type = c.type;
    f.value = std::string(c.value);
    table.mutable_active_filters().push_back(std::move(f));
    EXPECT_EQ(surviving_probe_mask(table), c.mask) << c.value;
  }
}

// The Apple / Apricot / Banana / Cherry / *star items of the Mac Excel 365
// ja-JP caption ordering probes, with Excel's amounts.
PivotCache build_fruit_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  const std::pair<const char*, double> items[] = {
      {"Apple", 100.0}, {"Apricot", 300.0}, {"Banana", 400.0}, {"Cherry", 200.0}, {"*star", 50.0}};
  for (const auto& [label, amount] : items) {
    PivotCacheRecord rec;
    rec.cells = {owned_text(cache, label), owned_text(cache, "P"), Value::number(amount)};
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

// Grand total of the fruit items `filter` keeps under mac-365-ja_JP; 0 when
// it keeps none.
double surviving_fruit_total(PivotFilter filter) {
  const PivotCache cache = build_fruit_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_active_filters().push_back(std::move(filter));
  auto r_or = evaluate(table, cache, pivot_layout_options_for(mac_365_ja_jp_profile()));
  EXPECT_TRUE(static_cast<bool>(r_or));
  if (!r_or || !r_or.value().grand_total.is_number()) {
    return 0.0;
  }
  return r_or.value().grand_total.as_number();
}

// A wildcard operand orders by the pattern match: a matching item counts as
// equal and any other as greater.
TEST(PivotEvaluatorCaptionFilter, WildcardOperandsOrderByTheMatch) {
  struct Case {
    FilterType type;
    const char* operand;
    double total;
  };
  const Case cases[] = {
      {FilterType::LabelGreaterThan, "B*", 650.0},
      {FilterType::LabelLessThan, "B*", 0.0},
      {FilterType::LabelLessThanOrEqual, "B*", 400.0},
      {FilterType::LabelGreaterThanOrEqual, "B*", 1050.0},
      {FilterType::LabelGreaterThan, "Z*", 1050.0},
      {FilterType::LabelGreaterThan, "A?ple", 950.0},
      {FilterType::LabelGreaterThan, "~*star", 1000.0},
      {FilterType::LabelGreaterThan, "B", 600.0},
      {FilterType::LabelGreaterThan, "*", 0.0},
      {FilterType::LabelGreaterThan, "Banana", 200.0},
      {FilterType::LabelLessThan, "Banana", 450.0},
      {FilterType::LabelLessThanOrEqual, "Banana", 850.0},
      {FilterType::LabelEqual, "A*", 400.0},
      {FilterType::LabelEqual, "~*star", 50.0},
      {FilterType::LabelEqual, "*star", 50.0},
      {FilterType::LabelEqual, "\xEF\xBC\xA1pple", 100.0},
      {FilterType::LabelNotEqual, "Apple", 950.0},
      {FilterType::LabelContains, "*", 1050.0},
  };
  for (const Case& c : cases) {
    EXPECT_DOUBLE_EQ(surviving_fruit_total(decoded_filter(0, c.type, std::string(c.operand))), c.total) << c.operand;
  }
  EXPECT_DOUBLE_EQ(
      surviving_fruit_total(decoded_filter(0, FilterType::LabelBetween, std::string("Apricot"), std::string("Cherry"))),
      900.0);
  EXPECT_DOUBLE_EQ(surviving_fruit_total(
                       decoded_filter(0, FilterType::LabelNotBetween, std::string("Apricot"), std::string("Cherry"))),
                   150.0);
}

// Numeric labels compare as numbers against a numeric operand.
TEST(PivotEvaluatorCaptionFilter, NumericLabelsCompareAsNumbers) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  const std::pair<double, double> items[] = {{1.0, 100.0}, {2.0, 300.0}, {3.0, 400.0}, {4.0, 200.0}, {10.0, 50.0}};
  for (const auto& [label, amount] : items) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(label), owned_text(cache, "P"), Value::number(amount)};
    cache.mutable_records().push_back(std::move(rec));
  }
  const auto total = [&](PivotFilter filter) {
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    table.mutable_active_filters().push_back(std::move(filter));
    auto r_or = evaluate(table, cache, pivot_layout_options_for(mac_365_ja_jp_profile()));
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or && r_or.value().grand_total.is_number() ? r_or.value().grand_total.as_number() : 0.0;
  };
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelGreaterThan, std::string("3"))), 250.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelBetween, std::string("2"), std::string("4"))), 900.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelEqual, std::string("10"))), 50.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelLessThan, std::string("3"))), 400.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelGreaterThanOrEqual, std::string("10"))), 50.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelBeginsWith, std::string("1"))), 150.0);
}

// A numeric label orders before a text operand that does not read as a number;
// "between" with such a bound matches nothing, and "not between" keeps the labels
// below the low bound or above the high one.
TEST(PivotEvaluatorCaptionFilter, NumericLabelsOrderBeforeTextOperands) {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  const std::pair<double, double> items[] = {{1.0, 100.0}, {2.0, 300.0}, {3.0, 400.0}, {4.0, 200.0}, {10.0, 50.0}};
  for (const auto& [label, amount] : items) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(label), owned_text(cache, "P"), Value::number(amount)};
    cache.mutable_records().push_back(std::move(rec));
  }
  const auto total = [&](PivotFilter filter) {
    PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
    table.mutable_active_filters().push_back(std::move(filter));
    auto r_or = evaluate(table, cache, pivot_layout_options_for(mac_365_ja_jp_profile()));
    EXPECT_TRUE(static_cast<bool>(r_or));
    return r_or && r_or.value().grand_total.is_number() ? r_or.value().grand_total.as_number() : 0.0;
  };
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelGreaterThan, std::string("3"))), 250.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelGreaterThan, std::string("abc"))), 0.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelGreaterThanOrEqual, std::string("abc"))), 0.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelLessThan, std::string("abc"))), 1050.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelLessThanOrEqual, std::string("abc"))), 1050.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelBetween, std::string("2"), std::string("abc"))), 0.0);
  EXPECT_DOUBLE_EQ(total(decoded_filter(0, FilterType::LabelNotBetween, std::string("2"), std::string("abc"))), 100.0);
}

TEST(PivotEvaluator, AuthoredCaptionBetweenFilterKeepsTheInclusiveRange) {
  PivotCache cache = build_basic_cache();
  // Row on Product so the range spans "Gadget" / "Widget".
  PivotTable table = build_sum_amount_table(/*row=*/{1}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_active_filters().push_back(
      decoded_filter(1, FilterType::LabelBetween, std::string("A"), std::string("M")));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  ASSERT_EQ(r_or.value().rows.size(), 1U);
  EXPECT_EQ(r_or.value().rows[0].label, "Gadget");
  EXPECT_DOUBLE_EQ(r_or.value().values[0][0][0].as_number(), 350.0);  // 50 + 300
}

TEST(PivotEvaluator, AuthoredCaptionFilterWithOutOfRangeFieldIsInert) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_active_filters().push_back(decoded_filter(99, FilterType::LabelEqual, std::string("North")));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  EXPECT_EQ(r_or.value().rows.size(), 2U);
}

TEST(PivotEvaluator, ManualFilterHidesItem) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Hide all records with Region == "South" via the Region field's items.
  table.mutable_fields()[0].items = {
      PivotItem{"North", /*visible=*/true},
      PivotItem{"South", /*visible=*/false},
  };

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only North survives.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);
}

TEST(PivotEvaluator, ManualFilterHidesTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_blank;
  hidden_blank.visible = false;
  hidden_blank.has_cache_index = true;
  hidden_blank.cache_index = kSharedRegionBlank;
  table.mutable_fields()[0].items.push_back(hidden_blank);
  // The load path leaves the item unnamed, because the value it binds to
  // renders to nothing. Running it here pins that the filter works on the
  // shape a workbook actually arrives in.
  resolve_pivot_names(table, cache);
  ASSERT_TRUE(table.fields()[0].items[0].name.empty());

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  const PivotLayoutOptions defaults;
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  // The text value spelled like the placeholder is untouched; only the group
  // with no value is gone, and its 10 is out of every aggregate.
  const std::size_t placeholder_text = row_index(r, defaults.blank_item_label);
  ASSERT_NE(placeholder_text, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[placeholder_text][0][0].as_number(), 7.0);
}

TEST(PivotEvaluator, HidingPlaceholderSpelledTextKeepsTheBlankItem) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_text;
  hidden_text.visible = false;
  hidden_text.has_cache_index = true;
  hidden_text.cache_index = kSharedRegionPlaceholderText;
  table.mutable_fields()[0].items.push_back(hidden_text);
  resolve_pivot_names(table, cache);
  const PivotLayoutOptions defaults;
  ASSERT_EQ(table.fields()[0].items[0].name, defaults.blank_item_label);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_DOUBLE_EQ(r.values[row_index(r, "North")][0][0].as_number(), 100.0);
  const std::size_t blank_leaf = row_index(r, defaults.blank_item_label);
  ASSERT_NE(blank_leaf, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[blank_leaf][0][0].as_number(), 10.0);
}

TEST(PivotEvaluator, UnnamedItemBoundByCacheIndexHidesItsValue) {
  PivotCache cache = build_shared_north_south_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  PivotItem hidden_south;
  hidden_south.visible = false;
  hidden_south.has_cache_index = true;
  hidden_south.cache_index = 1U;
  table.mutable_fields()[0].items.push_back(hidden_south);
  ASSERT_TRUE(table.fields()[0].items[0].name.empty());

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 175.0);
}

TEST(PivotItemLabel, NameWinsThenBindingThenEmpty) {
  const PivotCache cache = build_shared_north_south_cache();
  PivotItem named;
  named.name = "Renamed";
  named.cache_index = 1U;
  EXPECT_EQ(pivot_item_label(cache, 0U, named), "Renamed");
  PivotItem bound;
  bound.has_cache_index = true;
  bound.cache_index = 1U;
  EXPECT_EQ(pivot_item_label(cache, 0U, bound), "South");
  PivotItem dangling;
  dangling.has_cache_index = true;
  dangling.cache_index = 9U;
  EXPECT_EQ(pivot_item_label(cache, 0U, dangling), "");
  EXPECT_EQ(pivot_item_label(cache, 7U, bound), "");
}

TEST(PivotEvaluator, ManualFilterMatchesNumericDisplayLabel) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // The numeric Amount field is rendered as its integral display label.
  // Hiding 300 removes only South/Gadget without allocating a label string
  // per record while the filter is evaluated.
  table.mutable_fields()[2].items = {PivotItem{"300", /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 175.0);
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
}

TEST(PivotEvaluator, ManualFilterHidesNonIntegralNumericItem) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1.5)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t north = row_index(r, "North");
  ASSERT_NE(north, static_cast<std::size_t>(-1));
  // North keeps only the 2.0 record; the 1.5 one is gone from the aggregate.
  EXPECT_DOUBLE_EQ(r.values[north][0][0].as_number(), 2.0);
}

TEST(PivotEvaluator, ManualFilterHidesNumericItemBeyondIntegerRange) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  table.mutable_fields()[2].items = {PivotItem{display_string(Value::number(1e20)), /*visible=*/false}};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 4.0);
}

TEST(PivotEvaluator, LabelContainsFilterMatchesNumericItemLabel) {
  PivotCache cache = build_non_integral_amount_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  // A caption or date filter acts only on a field placed on an axis.
  table.set_field_axis(2, PivotAxis::Page);
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Amount";
  f.type = FilterType::LabelContains;
  f.value = display_string(Value::number(1e20));  // "1E+20"
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 1e20);
}

TEST(PivotEvaluator, AnItemListThatHidesNothingDoesNotChangeTheResult) {
  PivotCache cache = build_basic_cache();
  PivotTable unfiltered = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  unfiltered.set_grand_totals(/*rows=*/false, /*cols=*/false);
  auto baseline_or = evaluate(unfiltered, cache);
  ASSERT_TRUE(static_cast<bool>(baseline_or)) << baseline_or.error().message;

  PivotTable listed = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  listed.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // Far more items than the field has values, none of them hidden.
  for (std::size_t i = 0; i < 500; ++i) {
    listed.mutable_fields()[0].items.push_back(PivotItem{"item_" + std::to_string(i), /*visible=*/true});
  }
  listed.mutable_fields()[0].items.push_back(PivotItem{"North", /*visible=*/true});
  listed.mutable_fields()[0].items.push_back(PivotItem{"South", /*visible=*/true});
  auto listed_or = evaluate(listed, cache);
  ASSERT_TRUE(static_cast<bool>(listed_or)) << listed_or.error().message;

  const PivotResult& baseline = baseline_or.value();
  const PivotResult& with_list = listed_or.value();
  ASSERT_EQ(with_list.rows.size(), baseline.rows.size());
  for (std::size_t i = 0; i < baseline.rows.size(); ++i) {
    EXPECT_EQ(with_list.rows[i].label, baseline.rows[i].label);
    EXPECT_DOUBLE_EQ(with_list.values[i][0][0].as_number(), baseline.values[i][0][0].as_number());
  }
}

TEST(PivotEvaluator, OneFieldHidesTheBlankAndANamedItemTogether) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem hidden_blank;
  hidden_blank.visible = false;
  hidden_blank.has_cache_index = true;
  hidden_blank.cache_index = kSharedRegionBlank;
  table.mutable_fields()[0].items.push_back(hidden_blank);
  table.mutable_fields()[0].items.push_back(PivotItem{"North", /*visible=*/false});
  resolve_pivot_names(table, cache);

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the text value spelled like the placeholder survives.
  const PivotLayoutOptions defaults;
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, defaults.blank_item_label);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 7.0);
}

TEST(PivotEvaluator, AnUnlabelledHiddenItemBoundToAValueHidesThatValue) {
  PivotCache cache = build_shared_region_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  PivotItem unlabelled;
  unlabelled.visible = false;
  unlabelled.has_cache_index = true;
  unlabelled.cache_index = 0U;  // "North", a value that renders to a label.
  table.mutable_fields()[0].items.push_back(std::move(unlabelled));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // North is gone; the blank group and the text value spelled like the
  // placeholder both survive. They draw the same label, so they are counted
  // through the total rather than looked up by it.
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(row_index(r, "North"), static_cast<std::size_t>(-1));
  double total = 0.0;
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    total += r.values[i][0][0].as_number();
  }
  EXPECT_DOUBLE_EQ(total, 10.0 + 7.0);
}

TEST(PivotEvaluator, ManualFilterMatchesANumericLabelAmongManyItems) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);
  // 2,000 visible numeric items the cache has no record for, plus the one that
  // matches. Only the last one may prune anything.
  for (std::size_t i = 1000; i < 3000; ++i) {
    table.mutable_fields()[2].items.push_back(PivotItem{std::to_string(i), /*visible=*/true});
  }
  table.mutable_fields()[2].items.push_back(PivotItem{"300", /*visible=*/false});

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  const std::size_t south = row_index(r, "South");
  ASSERT_NE(south, static_cast<std::size_t>(-1));
  // South loses its 300 record and keeps the 200 one.
  EXPECT_DOUBLE_EQ(r.values[south][0][0].as_number(), 200.0);
}

TEST(PivotEvaluator, ManualFilterCostDoesNotFollowTheItemListLength) {
  constexpr std::size_t kRecords = 10000;

  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Serial", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (std::size_t i = 0; i < kRecords; ++i) {
    PivotCacheRecord rec;
    // Non-integral so the label goes through the general numeric rendering.
    rec.cells = {Value::number(static_cast<double>(i) + 0.25), Value::number(1.0)};
    cache.mutable_records().push_back(std::move(rec));
  }

  // Every item is hidden and none of them names a value any record carries, so
  // the item list prunes nothing however long it is.
  const auto build = [](std::size_t item_count) {
    PivotTable table;
    table.set_pivot_cache_id(1);
    PivotField serial_field;
    serial_field.source_name = "Serial";
    serial_field.axis = PivotAxis::Row;
    for (std::size_t i = 0; i < item_count; ++i) {
      serial_field.items.push_back(
          PivotItem{display_string(Value::number(-static_cast<double>(i) - 0.5)), /*visible=*/false});
    }
    table.mutable_fields().push_back(std::move(serial_field));
    PivotField amount_field;
    amount_field.source_name = "Amount";
    amount_field.axis = PivotAxis::Value;
    table.mutable_fields().push_back(std::move(amount_field));
    PivotDataField count;
    count.name = "Count of Amount";
    count.field_index = 1;
    count.aggregation = Aggregation::Count;
    table.mutable_data_fields().push_back(std::move(count));
    table.mutable_row_field_order() = {0};
    table.set_grand_totals(/*rows=*/false, /*cols=*/false);
    return table;
  };
  const PivotTable short_list = build(20);
  const PivotTable long_list = build(20000);

  const auto timed = [&](const PivotTable& table) {
    const auto started = std::chrono::steady_clock::now();
    auto r_or = evaluate(table, cache);
    const double seconds = std::chrono::duration<double>(std::chrono::steady_clock::now() - started).count();
    EXPECT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
    EXPECT_EQ(r_or.value().rows.size(), kRecords);
    return seconds;
  };
  const double short_seconds = timed(short_list);
  const double long_seconds = timed(long_list);

  // The allowance is deliberately loose — a factor of ten plus half a second —
  // because it is separating a constant from a thousandfold, not measuring the
  // machine.
  EXPECT_LT(long_seconds, short_seconds * 10.0 + 0.5)
      << "short list " << short_seconds << "s, long list " << long_seconds << "s";
}

TEST(PivotEvaluator, LabelFilterResolvesFieldByCustomName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  table.mutable_fields()[0].custom_name = "Area";

  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Area";  // The display name, not the source name.
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("N");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the North records survive, so both the axis and the aggregate move.
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_EQ(r.values.size(), 1U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 175.0);  // 100 + 50 + 25
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 175.0);
}

TEST(PivotEvaluator, LabelFilterResolvesFieldByDataFieldName) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{});
  // A caption or date filter acts only on a field placed on an axis.
  table.set_field_axis(2, PivotAxis::Page);

  // "Sum of Amount" is the data field's display name for the Amount field,
  // so the filter applies to Amount's own values.
  PivotFilter f;
  f.axis = PivotAxis::Row;
  f.field_name = "Sum of Amount";
  f.type = FilterType::LabelBeginsWith;
  f.value = std::string("1");
  table.mutable_active_filters().push_back(std::move(f));

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();

  // Only the Amount=100 record starts with "1".
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_EQ(r.rows[0].label, "North");
  ASSERT_TRUE(r.grand_total.is_number());
  EXPECT_DOUBLE_EQ(r.grand_total.as_number(), 100.0);
}

}  // namespace
}  // namespace formulon::pivot
