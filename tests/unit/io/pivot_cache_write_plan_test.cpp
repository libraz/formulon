//
// Unit tests for `build_pivot_cache_write_plan` and the plan-driven output of
// `write_pivot_cache_definition` / `write_pivot_cache_records`.

#include <cstdint>
#include <string>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_cache_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::io {
namespace {

Value MakeText(pivot::PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.mutable_text_storage().back());
}

void AddField(pivot::PivotCache& cache, const char* name) {
  pivot::PivotCacheField field;
  field.name = name;
  cache.mutable_fields().push_back(std::move(field));
}

void AddRecord(pivot::PivotCache& cache, std::vector<Value> cells) {
  pivot::PivotCacheRecord record;
  record.cells = std::move(cells);
  cache.mutable_records().push_back(std::move(record));
}

pivot::PivotTable TableWithAxes(const std::vector<pivot::PivotAxis>& axes) {
  pivot::PivotTable table;
  for (const pivot::PivotAxis axis : axes) {
    pivot::PivotField field;
    field.axis = axis;
    table.mutable_fields().push_back(std::move(field));
  }
  return table;
}

bool Contains(const std::string& haystack, const char* needle) {
  return haystack.find(needle) != std::string::npos;
}

TEST(PivotCacheWritePlan, TextAndBlankFieldIsDerivedInFirstAppearanceOrder) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Region");
  AddRecord(cache, {MakeText(cache, "West")});
  AddRecord(cache, {MakeText(cache, "East")});
  AddRecord(cache, {Value::blank()});
  AddRecord(cache, {MakeText(cache, "West")});

  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  ASSERT_EQ(plan.fields.size(), 1U);
  const PivotCacheFieldWritePlan& f = plan.fields[0];
  ASSERT_TRUE(f.derived);
  ASSERT_EQ(f.items.size(), 3U);
  EXPECT_EQ(f.items[0].as_text(), "West");
  EXPECT_EQ(f.items[1].as_text(), "East");
  EXPECT_TRUE(f.items[2].is_blank());
  EXPECT_EQ(f.record_index, (std::vector<std::uint32_t>{0, 1, 2, 0}));
  // Ascending: East, West, then blank last.
  EXPECT_EQ(f.asc_order, (std::vector<std::uint32_t>{1, 0, 2}));

  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(
      Contains(def, "<sharedItems containsBlank=\"1\" count=\"3\"><s v=\"West\"/><s v=\"East\"/><m/></sharedItems>"))
      << def;
  const std::string rec = write_pivot_cache_records(cache, &plan);
  EXPECT_TRUE(Contains(rec, "<r><x v=\"0\"/></r><r><x v=\"1\"/></r><r><x v=\"2\"/></r><r><x v=\"0\"/></r>")) << rec;
}

TEST(PivotCacheWritePlan, TextOnlyFieldCarriesJustTheCount) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Product");
  AddRecord(cache, {MakeText(cache, "Pen")});
  AddRecord(cache, {MakeText(cache, "Ink")});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(def, "<sharedItems count=\"2\"><s v=\"Pen\"/><s v=\"Ink\"/></sharedItems>")) << def;
}

TEST(PivotCacheWritePlan, CaseVariantsCollapseUnderTheProfile) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Region");
  AddRecord(cache, {MakeText(cache, "East")});
  AddRecord(cache, {MakeText(cache, "EAST")});
  AddRecord(cache, {MakeText(cache, "West")});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  ASSERT_TRUE(plan.fields[0].derived);
  ASSERT_EQ(plan.fields[0].items.size(), 2U);
  EXPECT_EQ(plan.fields[0].items[0].as_text(), "East");
  EXPECT_EQ(plan.fields[0].record_index, (std::vector<std::uint32_t>{0, 0, 1}));
}

TEST(PivotCacheWritePlan, NumericOffAxisFieldKeepsThePlaceholder) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Amt");
  AddRecord(cache, {Value::number(30)});
  AddRecord(cache, {Value::number(100)});
  const pivot::PivotTable table = TableWithAxes({pivot::PivotAxis::Value});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {&table});
  EXPECT_FALSE(plan.fields[0].derived);
  EXPECT_EQ(write_pivot_cache_definition(cache, &plan), write_pivot_cache_definition(cache));
  EXPECT_EQ(write_pivot_cache_records(cache, &plan), write_pivot_cache_records(cache));
  EXPECT_TRUE(Contains(write_pivot_cache_definition(cache, &plan), "<sharedItems containsNumber=\"1\"/>"));
}

TEST(PivotCacheWritePlan, NumericFieldOnAnAxisBecomesDiscrete) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Qty");
  AddRecord(cache, {Value::number(5)});
  AddRecord(cache, {Value::number(3)});
  AddRecord(cache, {Value::number(5)});
  const pivot::PivotTable table = TableWithAxes({pivot::PivotAxis::Row});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {&table});
  ASSERT_TRUE(plan.fields[0].derived);
  EXPECT_EQ(plan.fields[0].asc_order, (std::vector<std::uint32_t>{1, 0}));
  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(def,
                       "<sharedItems containsSemiMixedTypes=\"0\" containsString=\"0\" containsNumber=\"1\" "
                       "containsInteger=\"1\" minValue=\"3\" maxValue=\"5\" count=\"2\"><n v=\"5\"/><n v=\"3\"/>"))
      << def;
}

TEST(PivotCacheWritePlan, NumbersWithBlankAndFractionsHint) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "X");
  AddRecord(cache, {Value::number(1.5)});
  AddRecord(cache, {Value::blank()});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(def,
                       "<sharedItems containsString=\"0\" containsBlank=\"1\" containsNumber=\"1\" minValue=\"1.5\" "
                       "maxValue=\"1.5\" count=\"2\">"))
      << def;
}

TEST(PivotCacheWritePlan, MixedNumberAndTextHint) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "X");
  AddRecord(cache, {Value::number(1)});
  AddRecord(cache, {MakeText(cache, "a")});
  AddRecord(cache, {Value::blank()});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(
      Contains(def, "<sharedItems containsBlank=\"1\" containsMixedTypes=\"1\" containsNumber=\"1\" count=\"3\">"))
      << def;
}

TEST(PivotCacheWritePlan, BoolFieldEmitsItemsOnly) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Flag");
  AddRecord(cache, {Value::boolean(true)});
  AddRecord(cache, {Value::boolean(false)});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  const std::string def = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(def, "<sharedItems count=\"2\"><b v=\"1\"/><b v=\"0\"/></sharedItems>")) << def;
}

TEST(PivotCacheWritePlan, FileReadFieldsAreNotDerived) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  // Shared items, captured hints and a captured grouping each keep the field as read.
  pivot::PivotCacheField shared;
  shared.name = "Shared";
  shared.shared_items.push_back(MakeText(cache, "a"));
  cache.mutable_fields().push_back(std::move(shared));
  pivot::PivotCacheField hinted;
  hinted.name = "Hinted";
  hinted.shared_items_hints.present = true;
  cache.mutable_fields().push_back(std::move(hinted));
  pivot::PivotCacheField grouped;
  grouped.name = "Grouped";
  grouped.field_group_xml = "<fieldGroup base=\"0\"/>";
  cache.mutable_fields().push_back(std::move(grouped));
  AddRecord(cache, {Value::number(0), MakeText(cache, "x"), MakeText(cache, "y")});

  const pivot::PivotTable table = TableWithAxes({pivot::PivotAxis::Row, pivot::PivotAxis::Row, pivot::PivotAxis::Row});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {&table});
  ASSERT_EQ(plan.fields.size(), 3U);
  for (const PivotCacheFieldWritePlan& f : plan.fields) {
    EXPECT_FALSE(f.derived);
  }
  EXPECT_EQ(write_pivot_cache_definition(cache, &plan), write_pivot_cache_definition(cache));
}

TEST(PivotCacheWritePlan, NullPlanKeepsTheOldBytesAndDerivedOutputReads) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Region");
  AddRecord(cache, {MakeText(cache, "East")});
  AddRecord(cache, {Value::blank()});
  const std::string old_def = write_pivot_cache_definition(cache);
  EXPECT_TRUE(Contains(old_def, "<sharedItems containsNumber=\"1\"/>"));
  EXPECT_TRUE(Contains(write_pivot_cache_records(cache), "<r><s v=\"East\"/></r><r><m/></r>"));

  // The derived form reads back as a shared field resolving to the same values.
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  const std::string def = write_pivot_cache_definition(cache, &plan);
  const std::string rec = write_pivot_cache_records(cache, &plan);
  auto parsed = read_pivot_cache_definition(std::vector<std::uint8_t>(def.begin(), def.end()));
  ASSERT_TRUE(static_cast<bool>(parsed)) << parsed.error().message;
  auto status = read_pivot_cache_records(std::vector<std::uint8_t>(rec.begin(), rec.end()), parsed.value());
  ASSERT_TRUE(static_cast<bool>(status));
  ASSERT_EQ(parsed.value().fields()[0].shared_items.size(), 2U);
  EXPECT_EQ(parsed.value().fields()[0].shared_items[0].as_text(), "East");
  EXPECT_TRUE(parsed.value().fields()[0].shared_items[1].is_blank());
}

}  // namespace
}  // namespace formulon::io
