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
  grouped.field_group = pivot::FieldGroup{};
  grouped.field_group->base = 0U;
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

// ---------------------------------------------------------------------------
// refreshOnLoad and date hints
// ---------------------------------------------------------------------------

// A cache as read from a file: captured root attributes, one of them an
// authored `refreshOnLoad="0"`.
pivot::PivotCache LoadedCache() {
  pivot::PivotCache cache;
  cache.mutable_passthrough_attrs().emplace_back("refreshedBy", "x");
  cache.mutable_passthrough_attrs().emplace_back("refreshOnLoad", "0");
  pivot::PivotCacheField field;
  field.name = "Amt";
  field.shared_items_hints.present = true;
  field.shared_items_hints.has_contains_number = true;
  field.shared_items_hints.contains_number = true;
  cache.mutable_fields().push_back(std::move(field));
  AddRecord(cache, {Value::number(1)});
  return cache;
}

TEST(PivotCacheWritePlan, UntouchedLoadedCacheKeepsItsRefreshAttribute) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  const pivot::PivotCache cache = LoadedCache();
  const pivot::PivotTable table = TableWithAxes({pivot::PivotAxis::Value});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {&table});
  EXPECT_FALSE(plan.refresh_on_load);
  const std::string xml = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(xml, " refreshOnLoad=\"0\"")) << xml;
  EXPECT_FALSE(Contains(xml, "refreshOnLoad=\"1\"")) << xml;
}

TEST(PivotCacheWritePlan, EditedTableAsksExcelToRebuildALoadedCache) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  const pivot::PivotCache cache = LoadedCache();
  const pivot::PivotTable untouched = TableWithAxes({pivot::PivotAxis::Value});
  pivot::PivotTable edited = TableWithAxes({pivot::PivotAxis::Value});
  edited.mark_layout_dirty();
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {&untouched, &edited});
  EXPECT_TRUE(plan.refresh_on_load);
  const std::string xml = write_pivot_cache_definition(cache, &plan);
  // The authored value is replaced, not duplicated.
  EXPECT_TRUE(Contains(xml, " refreshOnLoad=\"1\"")) << xml;
  EXPECT_FALSE(Contains(xml, "refreshOnLoad=\"0\"")) << xml;
  EXPECT_EQ(xml.find("refreshOnLoad"), xml.rfind("refreshOnLoad")) << xml;
  EXPECT_TRUE(Contains(xml, " refreshedBy=\"x\"")) << xml;
}

TEST(PivotCacheWritePlan, InMemoryCacheRefreshesOnLoad) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache;
  AddField(cache, "Amt");
  AddRecord(cache, {Value::number(1)});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  EXPECT_TRUE(plan.refresh_on_load);
  EXPECT_TRUE(Contains(write_pivot_cache_definition(cache, &plan), " refreshOnLoad=\"1\""));
}

// An in-memory date field grouped by months: serials 45306.75 (2024-01-15
// 18:00) and 45495 (2024-07-22).
pivot::PivotCache InMemoryMonthGroupedCache() {
  pivot::PivotCache cache;
  AddField(cache, "Date");
  cache.mutable_fields()[0].field_group = pivot::FieldGroup{};
  cache.mutable_fields()[0].field_group->par = 1U;
  pivot::PivotCacheField months;
  months.name = "Months";
  months.num_fmt_id = 0U;
  months.is_database_field = false;
  months.field_group = pivot::FieldGroup{};
  months.field_group->base = 0U;
  months.field_group->kind = pivot::GroupKind::Range;
  months.field_group->range.by = pivot::GroupBy::Months;
  months.field_group->range.is_date = true;
  cache.mutable_fields().push_back(std::move(months));
  AddRecord(cache, {Value::number(45495.0), Value::blank()});
  AddRecord(cache, {Value::number(45306.75), Value::blank()});
  return cache;
}

TEST(PivotCacheWritePlan, InMemoryDateGroupBaseGetsDateHints) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  const pivot::PivotCache cache = InMemoryMonthGroupedCache();
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  ASSERT_TRUE(plan.fields[0].hints.present);
  EXPECT_FALSE(plan.fields[0].derived);
  EXPECT_FALSE(plan.fields[1].hints.present);
  const std::string xml = write_pivot_cache_definition(cache, &plan);
  EXPECT_TRUE(Contains(xml,
                       "<cacheField name=\"Date\"><sharedItems containsSemiMixedTypes=\"0\" containsNonDate=\"0\" "
                       "containsDate=\"1\" containsString=\"0\" minDate=\"2024-01-15T00:00:00\" "
                       "maxDate=\"2024-07-23T00:00:00\"/><fieldGroup par=\"1\"/></cacheField>"))
      << xml;
  // The group's bounds repeat the base's, as Excel writes them.
  EXPECT_TRUE(Contains(xml,
                       "<cacheField name=\"Months\" numFmtId=\"0\" databaseField=\"0\"><fieldGroup base=\"0\">"
                       "<rangePr groupBy=\"months\" startDate=\"2024-01-15T00:00:00\" "
                       "endDate=\"2024-07-23T00:00:00\"/></fieldGroup>"))
      << xml;
}

TEST(PivotCacheWritePlan, InMemoryDateGroupBaseRecordsAreWrittenAsDates) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  const pivot::PivotCache cache = InMemoryMonthGroupedCache();
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  EXPECT_TRUE(plan.fields[0].dates);
  const std::string records = write_pivot_cache_records(cache, &plan);
  EXPECT_TRUE(Contains(records, "<r><d v=\"2024-07-22T00:00:00\"/></r><r><d v=\"2024-01-15T18:00:00\"/></r>"))
      << records;

  // The records read back as the same serials.
  const std::string definition = write_pivot_cache_definition(cache, &plan);
  auto def_or = read_pivot_cache_definition(std::vector<std::uint8_t>(definition.begin(), definition.end()));
  ASSERT_TRUE(static_cast<bool>(def_or)) << def_or.error().message;
  ASSERT_TRUE(static_cast<bool>(
      read_pivot_cache_records(std::vector<std::uint8_t>(records.begin(), records.end()), def_or.value())));
  EXPECT_EQ(def_or.value().records()[1].cells[0].as_number(), 45306.75);
}

TEST(PivotCacheWritePlan, ExplicitGroupBoundsWinOverTheBaseHints) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache = InMemoryMonthGroupedCache();
  cache.mutable_fields()[1].field_group->range.start_date = "2024-02-01T00:00:00";
  cache.mutable_fields()[1].field_group->range.auto_start = false;
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  EXPECT_TRUE(Contains(write_pivot_cache_definition(cache, &plan),
                       "<rangePr autoStart=\"0\" groupBy=\"months\" startDate=\"2024-02-01T00:00:00\" "
                       "endDate=\"2024-07-23T00:00:00\"/>"));
}

TEST(PivotCacheWritePlan, DateHintsNeedEveryRecordToBeADate) {
  const eval::ScopedEvalProfile scope(default_excel_profile());
  pivot::PivotCache cache = InMemoryMonthGroupedCache();
  AddRecord(cache, {MakeText(cache, "n/a"), Value::blank()});
  const PivotCacheWritePlan plan = build_pivot_cache_write_plan(cache, {});
  EXPECT_FALSE(plan.fields[0].hints.present);
  EXPECT_TRUE(Contains(write_pivot_cache_definition(cache, &plan),
                       "<cacheField name=\"Date\"><sharedItems "
                       "containsNumber=\"1\"/>"));
}

}  // namespace
}  // namespace formulon::io
