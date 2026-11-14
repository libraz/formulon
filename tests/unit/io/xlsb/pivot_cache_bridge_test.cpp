#include "io/xlsb/pivot_cache_bridge.h"

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/pivot_cache_reader.h"
#include "io/xlsb/pivot_cache_definition_records.h"
#include "io/xlsb/pivot_calc_item_records.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot_twin_corpus.h"
#include "utils/resource_budget.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using ItemEntries = std::vector<pivot_cache_definition::ItemEntry>;

/// One cache of a corpus package: its two binary parts and the xlsx twin's parts of the same number.
struct CacheCase {
  std::string label;
  std::vector<std::uint8_t> definition;
  std::vector<std::uint8_t> records;
  std::vector<std::uint8_t> xlsx_definition;
  std::vector<std::uint8_t> xlsx_records;
  bool date1904 = false;
};

std::vector<std::uint8_t> ZipEntry(const std::vector<std::uint8_t>& package, const std::string& name) {
  ZipReader zip;
  if (!zip.open(SpanOf(package))) {
    return {};
  }
  auto bytes = zip.read_entry(name);
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

/// `<case>.xlsx` for `<case>.xlsb` and for a second save `<case>.2.xlsb`.
std::string TwinOf(const std::string& xlsb_path) {
  std::string stem = xlsb_path.substr(0, xlsb_path.size() - 5U);
  if (stem.size() > 2U && stem.compare(stem.size() - 2U, 2U, ".2") == 0) {
    stem.resize(stem.size() - 2U);
  }
  return stem + ".xlsx";
}

/// The cache-item twins of xlsb_parity/pivot, both Excel saves of each.
std::vector<std::string> ItemKindPaths() {
  const std::string root = std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/pivot/";
  std::vector<std::string> paths;
  for (const char* stem : {"item_bare_m", "item_mixed_d", "item_b_e", "inline_datetime"}) {
    for (const char* save : {".xlsb", ".2.xlsb"}) {
      paths.push_back(root + stem + save);
    }
  }
  for (const char* stem : {"calc_items", "calc_items2", "d1904_inline", "d1904_items", "run_threshold"}) {
    paths.push_back(root + stem + ".xlsb");
  }
  return paths;
}

std::vector<CacheCase> Cases(const std::string& xlsb_path) {
  const std::string kDefinitionPrefix = "xl/pivotCache/pivotCacheDefinition";
  const std::string xlsx_path = TwinOf(xlsb_path);
  const std::vector<std::uint8_t> xlsx = ReadFixtureFile(xlsx_path);
  const std::vector<std::uint8_t> workbook = ZipEntry(xlsx, "xl/workbook.xml");
  const std::string workbook_xml(workbook.begin(), workbook.end());
  const bool date1904 = workbook_xml.find("date1904=\"1\"") != std::string::npos ||
                        workbook_xml.find("date1904=\"true\"") != std::string::npos;
  const auto records = pivot_twins::Parts(xlsb_path, "xl/pivotCache/pivotCacheRecords");
  std::vector<CacheCase> out;
  for (const auto& [name, bytes] : pivot_twins::Parts(xlsb_path, kDefinitionPrefix)) {
    const std::string number = name.substr(kDefinitionPrefix.size(), name.size() - kDefinitionPrefix.size() - 4U);
    CacheCase c;
    c.label = xlsb_path + " " + name;
    c.definition = bytes;
    c.date1904 = date1904;
    for (const auto& [records_name, records_bytes] : records) {
      if (records_name == "xl/pivotCache/pivotCacheRecords" + number + ".bin") {
        c.records = records_bytes;
      }
    }
    c.xlsx_definition = ZipEntry(xlsx, kDefinitionPrefix + number + ".xml");
    c.xlsx_records = ZipEntry(xlsx, "xl/pivotCache/pivotCacheRecords" + number + ".xml");
    EXPECT_FALSE(c.records.empty()) << c.label << " has no records part";
    EXPECT_FALSE(c.xlsx_definition.empty()) << c.label << " has no xlsx twin";
    out.push_back(std::move(c));
  }
  return out;
}

std::string Fixture(const std::string& stem) {
  std::vector<std::string> paths = pivot_twins::CorpusPaths();
  for (const std::string& path : ItemKindPaths()) {
    paths.push_back(path);
  }
  for (const std::string& path : paths) {
    const std::string suffix = "/" + stem + ".xlsb";
    if (path.size() >= suffix.size() && path.compare(path.size() - suffix.size(), suffix.size(), suffix) == 0) {
      return path;
    }
  }
  ADD_FAILURE() << "no corpus package " << stem;
  return {};
}

pivot::PivotCache FromXlsb(const CacheCase& c) {
  auto cache = pivot_cache_from_xlsb(SpanOf(c.definition), SpanOf(c.records), c.date1904);
  EXPECT_TRUE(cache.has_value()) << c.label << ": " << (cache ? "" : cache.error().message);
  return cache ? std::move(cache.value()) : pivot::PivotCache{};
}

pivot::PivotCache FirstCache(const std::string& stem, std::size_t index = 0) {
  const std::vector<CacheCase> cases = Cases(Fixture(stem));
  if (index >= cases.size()) {
    ADD_FAILURE() << stem << " has no cache " << index;
    return pivot::PivotCache{};
  }
  return FromXlsb(cases[index]);
}

std::string RecordsRelId(const std::vector<std::uint8_t>& definition) {
  auto doc = decode_pivot_cache_definition(SpanOf(definition));
  return doc ? doc.value().records_rel_id : std::string();
}

void ExpectSameHints(const pivot::SharedItemsHints& a, const pivot::SharedItemsHints& b, const std::string& where) {
  EXPECT_EQ(a.contains_number, b.contains_number) << where;
  EXPECT_EQ(a.contains_integer, b.contains_integer) << where;
  EXPECT_EQ(a.contains_date, b.contains_date) << where;
  EXPECT_EQ(a.contains_string, b.contains_string) << where;
  EXPECT_EQ(a.contains_blank, b.contains_blank) << where;
  EXPECT_EQ(a.contains_semi_mixed, b.contains_semi_mixed) << where;
  EXPECT_EQ(a.contains_non_date, b.contains_non_date) << where;
  EXPECT_EQ(a.contains_mixed_types, b.contains_mixed_types) << where;
  EXPECT_EQ(a.has_contains_number, b.has_contains_number) << where;
  EXPECT_EQ(a.has_contains_integer, b.has_contains_integer) << where;
  EXPECT_EQ(a.has_contains_date, b.has_contains_date) << where;
  EXPECT_EQ(a.has_contains_string, b.has_contains_string) << where;
  EXPECT_EQ(a.has_contains_blank, b.has_contains_blank) << where;
  EXPECT_EQ(a.has_contains_semi_mixed, b.has_contains_semi_mixed) << where;
  EXPECT_EQ(a.has_contains_non_date, b.has_contains_non_date) << where;
  EXPECT_EQ(a.has_contains_mixed_types, b.has_contains_mixed_types) << where;
  EXPECT_EQ(a.has_min_value, b.has_min_value) << where;
  EXPECT_EQ(a.has_max_value, b.has_max_value) << where;
  EXPECT_EQ(a.has_min_date, b.has_min_date) << where;
  EXPECT_EQ(a.has_max_date, b.has_max_date) << where;
  EXPECT_EQ(a.min_value, b.min_value) << where;
  EXPECT_EQ(a.max_value, b.max_value) << where;
  EXPECT_EQ(a.min_date, b.min_date) << where;
  EXPECT_EQ(a.max_date, b.max_date) << where;
  EXPECT_EQ(a.present, b.present) << where;
  EXPECT_EQ(a.extra_attrs, b.extra_attrs) << where;
}

void ExpectSameGroup(const pivot::FieldGroup& a, const pivot::FieldGroup& b, const std::string& where) {
  EXPECT_EQ(a.par, b.par) << where;
  EXPECT_EQ(a.base, b.base) << where;
  EXPECT_EQ(a.kind, b.kind) << where;
  EXPECT_EQ(a.range.by, b.range.by) << where;
  EXPECT_EQ(a.range.is_date, b.range.is_date) << where;
  EXPECT_EQ(a.range.start_date, b.range.start_date) << where;
  EXPECT_EQ(a.range.end_date, b.range.end_date) << where;
  EXPECT_EQ(a.range.start_num, b.range.start_num) << where;
  EXPECT_EQ(a.range.end_num, b.range.end_num) << where;
  EXPECT_EQ(a.range.auto_start, b.range.auto_start) << where;
  EXPECT_EQ(a.range.auto_end, b.range.auto_end) << where;
  EXPECT_EQ(a.range.interval, b.range.interval) << where;
  EXPECT_EQ(a.discrete_map, b.discrete_map) << where;
  EXPECT_EQ(a.group_items, b.group_items) << where;
  EXPECT_EQ(a.group_item_is_date, b.group_item_is_date) << where;
  EXPECT_EQ(a.extra_attrs, b.extra_attrs) << where;
}

void ExpectSameArea(const pivot::PivotArea& a, const pivot::PivotArea& b, const std::string& where) {
  EXPECT_EQ(a.cache_index, b.cache_index) << where;
  EXPECT_EQ(a.outline, b.outline) << where;
  EXPECT_EQ(a.field_position, b.field_position) << where;
  EXPECT_EQ(a.extra_attrs, b.extra_attrs) << where;
  EXPECT_EQ(a.extra_children_xml, b.extra_children_xml) << where;
  ASSERT_EQ(a.references.size(), b.references.size()) << where;
  for (std::size_t i = 0; i < a.references.size(); ++i) {
    EXPECT_EQ(a.references[i].field, b.references[i].field) << where;
    EXPECT_EQ(a.references[i].items, b.references[i].items) << where;
    EXPECT_EQ(a.references[i].extra_attrs, b.references[i].extra_attrs) << where;
    EXPECT_EQ(a.references[i].extra_children_xml, b.references[i].extra_children_xml) << where;
  }
}

/// Every member a cache part carries, `cache_id` (workbook part) aside.
void ExpectSameCache(const pivot::PivotCache& a, const pivot::PivotCache& b, const std::string& label) {
  EXPECT_EQ(a.passthrough_attrs(), b.passthrough_attrs()) << label;
  EXPECT_EQ(a.save_data(), b.save_data()) << label;
  EXPECT_EQ(a.missing_items_limit(), b.missing_items_limit()) << label;
  EXPECT_EQ(a.refresh_on_load_attr(), b.refresh_on_load_attr()) << label;
  EXPECT_EQ(a.cache_source_attrs(), b.cache_source_attrs()) << label;
  EXPECT_EQ(a.cache_source_children_xml(), b.cache_source_children_xml()) << label;
  EXPECT_EQ(a.worksheet_source().present, b.worksheet_source().present) << label;
  EXPECT_EQ(a.worksheet_source().ref, b.worksheet_source().ref) << label;
  EXPECT_EQ(a.worksheet_source().sheet, b.worksheet_source().sheet) << label;
  EXPECT_EQ(a.worksheet_source().name, b.worksheet_source().name) << label;
  EXPECT_EQ(a.trailing_children_xml(), b.trailing_children_xml()) << label;
  EXPECT_EQ(a.trailing_children_after_calculated_xml(), b.trailing_children_after_calculated_xml()) << label;

  ASSERT_EQ(a.fields().size(), b.fields().size()) << label;
  for (std::size_t i = 0; i < a.fields().size(); ++i) {
    const pivot::PivotCacheField& fa = a.fields()[i];
    const pivot::PivotCacheField& fb = b.fields()[i];
    const std::string where = label + " field " + std::to_string(i);
    EXPECT_EQ(fa.name, fb.name) << where;
    EXPECT_EQ(fa.shared_items, fb.shared_items) << where;
    EXPECT_EQ(fa.shared_item_is_date, fb.shared_item_is_date) << where;
    EXPECT_EQ(fa.shared_item_is_calculated, fb.shared_item_is_calculated) << where;
    EXPECT_EQ(fa.shared_item_is_missing, fb.shared_item_is_missing) << where;
    ExpectSameHints(fa.shared_items_hints, fb.shared_items_hints, where);
    EXPECT_EQ(fa.is_database_field, fb.is_database_field) << where;
    EXPECT_EQ(fa.num_fmt_id, fb.num_fmt_id) << where;
    ASSERT_EQ(fa.field_group.has_value(), fb.field_group.has_value()) << where;
    if (fa.field_group) {
      ExpectSameGroup(*fa.field_group, *fb.field_group, where);
    }
    EXPECT_EQ(fa.formula, fb.formula) << where;
    EXPECT_EQ(fa.extra_attrs, fb.extra_attrs) << where;
    EXPECT_EQ(fa.trailing_children_xml, fb.trailing_children_xml) << where;
  }

  ASSERT_EQ(a.calculated_items().size(), b.calculated_items().size()) << label;
  for (std::size_t i = 0; i < a.calculated_items().size(); ++i) {
    const pivot::PivotCalculatedItem& ia = a.calculated_items()[i];
    const pivot::PivotCalculatedItem& ib = b.calculated_items()[i];
    const std::string where = label + " calculated item " + std::to_string(i);
    EXPECT_EQ(ia.cache_field_index, ib.cache_field_index) << where;
    EXPECT_EQ(ia.formula, ib.formula) << where;
    EXPECT_EQ(ia.has_area, ib.has_area) << where;
    EXPECT_EQ(ia.has_field_attr, ib.has_field_attr) << where;
    EXPECT_EQ(ia.extra_children_xml, ib.extra_children_xml) << where;
    ExpectSameArea(ia.area, ib.area, where);
  }

  ASSERT_EQ(a.records().size(), b.records().size()) << label;
  for (std::size_t i = 0; i < a.records().size(); ++i) {
    const std::string where = label + " record " + std::to_string(i);
    EXPECT_EQ(a.records()[i].cells, b.records()[i].cells) << where;
    EXPECT_EQ(a.records()[i].cell_is_index, b.records()[i].cell_is_index) << where;
    EXPECT_EQ(a.records()[i].cell_is_date, b.records()[i].cell_is_date) << where;
  }
}

TEST(XlsbPivotCacheBridge, CorpusReadsWritesBackAndMatchesTheXlsxTwin) {
  std::size_t caches = 0;
  std::vector<std::string> paths = pivot_twins::CorpusPaths();
  for (const std::string& path : ItemKindPaths()) {
    paths.push_back(path);
  }
  for (const std::string& path : paths) {
    for (const CacheCase& c : Cases(path)) {
      ++caches;
      auto cache = pivot_cache_from_xlsb(SpanOf(c.definition), SpanOf(c.records), c.date1904);
      ASSERT_TRUE(cache.has_value()) << c.label << ": " << cache.error().message;

      auto parts = pivot_cache_to_xlsb(cache.value(), c.date1904, RecordsRelId(c.definition));
      ASSERT_TRUE(parts.has_value()) << c.label << ": " << parts.error().message;
      EXPECT_EQ(parts.value().definition, c.definition) << c.label;
      EXPECT_EQ(parts.value().records, c.records) << c.label;

      auto twin = read_pivot_cache_definition(c.xlsx_definition, c.date1904);
      ASSERT_TRUE(twin.has_value()) << c.label << ": " << twin.error().message;
      ASSERT_TRUE(
          read_pivot_cache_records(c.xlsx_records, twin.value(), kMaxPivotCacheRecordCells, c.date1904).has_value())
          << c.label;
      ExpectSameCache(cache.value(), twin.value(), c.label);
    }
  }
  EXPECT_GT(caches, 80U);
}

TEST(XlsbPivotCacheBridge, InlineDateRecordCellIsASerial) {
  const pivot::PivotCache cache = FirstCache("B01_base");
  ASSERT_FALSE(cache.records().empty());
  const pivot::PivotCacheRecord& first = cache.records().front();
  ASSERT_EQ(first.cells.size(), 5U);
  EXPECT_EQ(first.cells[2], Value::number(45296.0));  // 2024-01-05
  ASSERT_EQ(first.cell_is_date.size(), 5U);
  EXPECT_TRUE(first.cell_is_date[2]);
  EXPECT_FALSE(first.cell_is_index[2]);
  EXPECT_EQ(first.cells[1], Value::text("Apple"));
}

TEST(XlsbPivotCacheBridge, WorksheetSourceIsKept) {
  const pivot::PivotCache cache = FirstCache("B01_base");
  EXPECT_TRUE(cache.worksheet_source().present);
  EXPECT_EQ(cache.worksheet_source().ref, "A1:E13");
  EXPECT_EQ(cache.worksheet_source().sheet, "Sheet1");
  const std::vector<std::pair<std::string, std::string>> attrs{{"type", "worksheet"}};
  EXPECT_EQ(cache.cache_source_attrs(), attrs);
}

TEST(XlsbPivotCacheBridge, NumberItemsAndUnusedMissingItem) {
  const pivot::PivotCache cache = FirstCache("A5_C9__sum");
  ASSERT_EQ(cache.fields().size(), 5U);
  const pivot::PivotCacheField& region = cache.fields()[0];
  ASSERT_EQ(region.shared_items.size(), 5U);
  EXPECT_EQ(region.shared_items[3], Value::blank());
  EXPECT_EQ(region.shared_item_is_missing, (std::vector<bool>{false, false, true, true, true}));
  EXPECT_TRUE(region.shared_items_hints.contains_blank);
  const pivot::PivotCacheField& q = cache.fields()[3];
  EXPECT_EQ(q.shared_items, (std::vector<Value>{Value::number(1.0), Value::number(2.0)}));
  EXPECT_EQ(q.shared_items_hints.min_value, "1");
  EXPECT_EQ(q.shared_items_hints.max_value, "2");
}

TEST(XlsbPivotCacheBridge, DateRunUsesExcelCalendar) {
  const pivot::PivotCache cache = FirstCache("pivot_days_end");
  ASSERT_EQ(cache.fields().size(), 3U);
  const pivot::PivotCacheField& d = cache.fields()[0];
  ASSERT_EQ(d.shared_items.size(), 71U);
  EXPECT_EQ(d.shared_items[0], Value::number(0.0));    // 1900-01-00
  EXPECT_EQ(d.shared_items[60], Value::number(60.0));  // 1900-02-29
  EXPECT_EQ(d.shared_items[70], Value::number(70.0));
  EXPECT_EQ(d.shared_item_is_date, std::vector<bool>(71U, true));
  EXPECT_EQ(d.shared_items_hints.min_date, "1899-12-30T00:00:00");
  EXPECT_EQ(d.shared_items_hints.max_date, "1900-03-11T00:00:00");
}

TEST(XlsbPivotCacheBridge, DaysGroupIsRead) {
  const pivot::PivotCache cache = FirstCache("pivot_days_end");
  ASSERT_EQ(cache.fields().size(), 3U);
  ASSERT_TRUE(cache.fields()[0].field_group.has_value());
  EXPECT_EQ(cache.fields()[0].field_group->par, std::optional<std::uint32_t>(2U));
  EXPECT_EQ(cache.fields()[0].field_group->kind, pivot::GroupKind::None);
  const pivot::PivotCacheField& days = cache.fields()[2];
  EXPECT_FALSE(days.is_database_field);
  ASSERT_TRUE(days.field_group.has_value());
  const pivot::FieldGroup& g = *days.field_group;
  EXPECT_EQ(g.base, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(g.kind, pivot::GroupKind::Range);
  EXPECT_EQ(g.range.by, pivot::GroupBy::Days);
  EXPECT_EQ(g.range.start_date, "1899-12-30T00:00:00");
  EXPECT_EQ(g.range.end_date, "1900-02-28T00:00:00");
  EXPECT_TRUE(g.range.auto_start);
  EXPECT_FALSE(g.range.auto_end);
  EXPECT_EQ(g.range.interval, 7.0);
  ASSERT_EQ(g.group_items.size(), 11U);
  EXPECT_EQ(g.group_items.front(), Value::text("<1900/1/0"));
  ASSERT_FALSE(cache.records().empty());
  EXPECT_EQ(cache.records().front().cells[2], Value::blank());
}

TEST(XlsbPivotCacheBridge, CalculatedItemAndItsFlaggedSharedItem) {
  const pivot::PivotCache cache = FirstCache("C13_calc_item");
  const pivot::PivotCacheField& region = cache.fields()[0];
  ASSERT_EQ(region.shared_items.size(), 4U);
  EXPECT_EQ(region.shared_items[3], Value::text("EW"));
  EXPECT_EQ(region.shared_item_is_calculated, (std::vector<bool>{false, false, false, true}));
  ASSERT_EQ(cache.calculated_items().size(), 1U);
  const pivot::PivotCalculatedItem& item = cache.calculated_items().front();
  EXPECT_EQ(item.formula, "Region[East]+Region[West]");
  EXPECT_EQ(item.cache_field_index, 0U);
  EXPECT_FALSE(item.has_field_attr);
  ASSERT_TRUE(item.has_area);
  EXPECT_EQ(item.area.cache_index, std::optional<bool>(true));
  ASSERT_EQ(item.area.references.size(), 1U);
  EXPECT_EQ(item.area.references[0].field, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(item.area.references[0].items, std::vector<std::uint32_t>{3U});
}

TEST(XlsbPivotCacheBridge, BareMissingItem) {
  const pivot::PivotCache cache = FirstCache("item_bare_m");
  const pivot::PivotCacheField& region = cache.fields()[0];
  ASSERT_EQ(region.shared_items.size(), 4U);
  EXPECT_EQ(region.shared_items[3], Value::blank());
  EXPECT_EQ(region.shared_item_is_missing, std::vector<bool>(4U, false));
  EXPECT_TRUE(region.shared_items_hints.contains_blank);
}

TEST(XlsbPivotCacheBridge, LoneDateAndStringItemsOfAMixedField) {
  const pivot::PivotCache cache = FirstCache("item_mixed_d");
  const pivot::PivotCacheField& date = cache.fields()[2];
  ASSERT_EQ(date.shared_items.size(), 12U);
  EXPECT_EQ(date.shared_items[0], Value::number(45296.0));  // 2024-01-05
  EXPECT_EQ(date.shared_items[2], Value::text("n/a"));
  EXPECT_EQ(date.shared_items[6], Value::text("pending"));
  EXPECT_EQ(date.shared_item_is_date,
            (std::vector<bool>{true, true, false, true, true, true, false, true, true, true, true, true}));
  const pivot::SharedItemsHints& h = date.shared_items_hints;
  EXPECT_TRUE(h.has_contains_mixed_types && h.contains_mixed_types);
  EXPECT_TRUE(h.has_contains_date && h.contains_date);
  EXPECT_FALSE(h.has_contains_string);
  EXPECT_EQ(h.min_date, "2024-01-05T00:00:00");
}

TEST(XlsbPivotCacheBridge, BooleanAndErrorItems) {
  const pivot::PivotCache cache = FirstCache("item_b_e");
  const pivot::PivotCacheField& product = cache.fields()[1];
  ASSERT_EQ(product.shared_items.size(), 7U);
  EXPECT_EQ(product.shared_items[0], Value::boolean(true));
  EXPECT_EQ(product.shared_items[1], Value::boolean(false));
  EXPECT_EQ(product.shared_items[2], Value::error(ErrorCode::NA));
  EXPECT_EQ(product.shared_items[3], Value::error(ErrorCode::Div0));
  EXPECT_EQ(product.shared_items[4], Value::text("Widget"));
}

TEST(XlsbPivotCacheBridge, InlineDateTimeRecordCell) {
  const pivot::PivotCache cache = FirstCache("inline_datetime");
  ASSERT_FALSE(cache.records().empty());
  const double expected = 45296.0 + (6.0 * 3600.0 + 14.0 * 60.0 + 22.0) / 86400.0;  // 2024-01-05T06:14:22
  EXPECT_EQ(cache.records().front().cells[2], Value::number(expected));
  EXPECT_TRUE(cache.records().front().cell_is_date[2]);
  EXPECT_EQ(cache.fields()[2].shared_items_hints.min_date, "2024-01-05T06:14:22");
  EXPECT_EQ(cache.fields()[2].num_fmt_id, std::optional<std::uint32_t>(176U));
}

/// The rgce of the first calculated item of a written definition part.
std::vector<std::uint8_t> WrittenFormula(const pivot::PivotCache& cache) {
  auto parts = pivot_cache_to_xlsb(cache, false, "rId1");
  EXPECT_TRUE(parts.has_value());
  if (!parts) {
    return {};
  }
  auto doc = decode_pivot_cache_definition(SpanOf(parts.value().definition));
  EXPECT_TRUE(doc.has_value());
  if (!doc) {
    return {};
  }
  std::size_t pos = 0;
  auto items = pivot_calc::decode_calc_items(doc.value().calculated_items, &pos);
  EXPECT_TRUE(items.has_value());
  return items && !items.value().empty() ? items.value().front().rgce : std::vector<std::uint8_t>{};
}

void ExpectNotWritable(const pivot::PivotCache& cache) {
  auto parts = pivot_cache_to_xlsb(cache, false, "rId1");
  ASSERT_FALSE(parts.has_value());
  EXPECT_EQ(parts.error().code, FormulonErrorCode::kIoFeatureNotWritable);
}

/// The rgce of each calculated item of an Excel-saved definition part.
std::vector<std::vector<std::uint8_t>> ExcelFormulas(const std::string& stem) {
  const std::vector<CacheCase> cases = Cases(Fixture(stem));
  std::vector<std::vector<std::uint8_t>> out;
  if (cases.empty()) {
    ADD_FAILURE() << stem;
    return out;
  }
  auto doc = decode_pivot_cache_definition(SpanOf(cases.front().definition));
  EXPECT_TRUE(doc.has_value());
  if (!doc) {
    return out;
  }
  std::size_t pos = 0;
  auto items = pivot_calc::decode_calc_items(doc.value().calculated_items, &pos);
  EXPECT_TRUE(items.has_value());
  if (items) {
    for (const pivot_calc::CalcItem& item : items.value()) {
      out.push_back(item.rgce);
    }
  }
  return out;
}

void ExpectFormulasCompileToExcelsTokens(const std::string& stem, const std::vector<std::string>& formulas) {
  const std::vector<std::vector<std::uint8_t>> excel = ExcelFormulas(stem);
  ASSERT_EQ(excel.size(), formulas.size()) << stem;
  pivot::PivotCache cache = FirstCache("C13_calc_item");
  ASSERT_EQ(cache.calculated_items().size(), 1U);
  for (std::size_t i = 0; i < excel.size(); ++i) {
    cache.mutable_calculated_items()[0].formula = formulas[i];
    EXPECT_EQ(WrittenFormula(cache), excel[i]) << formulas[i];
  }
}

TEST(XlsbPivotCacheBridge, CalculatedItemFormulasCompileToExcelsTokens) {
  ExpectFormulasCompileToExcelsTokens(
      "calc_items", {"-Region[East]", "(Region[East]+Region[West])*2", "Region[East]*3", "Region[West]/2",
                     "Region[North]^2", "Region[East]&Region[West]", "Region[East]>Region[West]"});
  ExpectFormulasCompileToExcelsTokens(
      "calc_items2", {"Region[East]-Region[West]", "Region[East]<Region[West]", "Region[East]=Region[West]",
                      "Region[East]<=Region[West]", "Region[East]>=Region[West]", "Region[East]<>Region[West]"});
}

TEST(XlsbPivotCacheBridge, RefusesUnmeasuredFormulaOperators) {
  pivot::PivotCache cache = FirstCache("C13_calc_item");
  for (const char* formula : {"+Region[East]", "Region[East]%", "Region[East] Region[West]",
                              "Region[East],Region[West]", "Region[East]:Region[West]", "Region[East]==Region[West]"}) {
    cache.mutable_calculated_items()[0].formula = formula;
    SCOPED_TRACE(formula);
    ExpectNotWritable(cache);
  }
}

TEST(XlsbPivotCacheBridge, RefusesFormulaTextWithoutATokenSpelling) {
  pivot::PivotCache cache = FirstCache("C13_calc_item");
  const std::string deep = std::string(300, '(') + "Region[East]" + std::string(300, ')');
  for (const std::string& formula :
       {std::string("Region[East]*2.5"), std::string("Region[East]*02"), std::string("Region[East]*70000"),
        std::string("Region[East] + Region[West]"), deep}) {
    cache.mutable_calculated_items()[0].formula = formula;
    SCOPED_TRACE(formula.substr(0, 40));
    ExpectNotWritable(cache);
  }
}

TEST(XlsbPivotCacheBridge, RefusesANamedRangeSource) {
  pivot::PivotCache cache = FirstCache("B01_base");
  cache.mutable_worksheet_source().name = "Data";
  ExpectNotWritable(cache);
}

TEST(XlsbPivotCacheBridge, RefusesASharedDateItemWithAFractionOfASecond) {
  pivot::PivotCache cache = FirstCache("pivot_days_end");
  pivot::PivotCacheField& d = cache.mutable_fields()[0];
  d.shared_items[70] = Value::number(70.5);  // 1900-03-10T12:00:00 has a spelling.
  ASSERT_TRUE(pivot_cache_to_xlsb(cache, false, "rId1").has_value());
  d.shared_items[70] = Value::number(70.5 + 0.5 / 86400.0);
  ExpectNotWritable(cache);
}

void AppendSharedItem(pivot::PivotCache& cache, std::size_t field_index, const Value& value) {
  pivot::PivotCacheField& field = cache.mutable_fields()[field_index];
  field.shared_items.push_back(value);
  field.shared_item_is_date.resize(field.shared_items.size(), false);
  field.shared_item_is_calculated.resize(field.shared_items.size(), false);
  field.shared_item_is_missing.resize(field.shared_items.size(), false);
}

/// The item table of field `field_index` of a written definition part.
ItemEntries WrittenItems(const pivot::PivotCache& cache, std::size_t field_index) {
  auto parts = pivot_cache_to_xlsb(cache, false, "rId1");
  EXPECT_TRUE(parts.has_value());
  if (!parts) {
    return {};
  }
  auto doc = decode_pivot_cache_definition(SpanOf(parts.value().definition));
  EXPECT_TRUE(doc.has_value());
  if (!doc || !doc.value().fields[field_index].items) {
    return {};
  }
  return doc.value().fields[field_index].items->entries;
}

TEST(XlsbPivotCacheBridge, FourItemsInARowAreOneRun) {
  pivot::PivotCache strings = FirstCache("B01_base");
  strings.mutable_text_storage().emplace_back("South");
  AppendSharedItem(strings, 0, Value::text(strings.text_storage().back()));
  const ItemEntries four = WrittenItems(strings, 0);
  ASSERT_EQ(four.size(), 1U);
  EXPECT_TRUE(four[0].is_run);

  pivot::PivotCache numbers = FirstCache("A5_C9__sum");
  AppendSharedItem(numbers, 3, Value::number(3.0));
  const ItemEntries three = WrittenItems(numbers, 3);
  ASSERT_EQ(three.size(), 3U);
  EXPECT_FALSE(three[0].is_run);
}

TEST(XlsbPivotCacheBridge, Date1904TimedRunItem) {
  const pivot::PivotCache cache = FirstCache("d1904_items");
  const pivot::PivotCacheField& date = cache.fields()[2];
  ASSERT_EQ(date.shared_items.size(), 12U);
  // 2024-02-05T03:07:11 as the xlsx reader computes it in the 1904 system.
  const double expected = (45327.0 + (3.0 * 3600.0 + 7.0 * 60.0 + 11.0) / 86400.0) - 1462.0;
  EXPECT_EQ(date.shared_items[1], Value::number(expected));
  EXPECT_TRUE(date.shared_item_is_date[1]);
}

TEST(XlsbPivotCacheBridge, RefusesRootAttributesThatWouldNotReadBack) {
  pivot::PivotCache without_uid = FirstCache("B01_base");
  auto& attrs = without_uid.mutable_passthrough_attrs();
  ASSERT_EQ(attrs.back().first, "xr:uid");
  attrs.pop_back();
  ExpectNotWritable(without_uid);

  pivot::PivotCache short_date = FirstCache("B01_base");
  for (auto& [name, value] : short_date.mutable_passthrough_attrs()) {
    if (name == "refreshedDate") {
      value = "46336.50";  // the same double, not its 17-digit spelling
    }
  }
  ExpectNotWritable(short_date);
}

TEST(XlsbPivotCacheBridge, RefusesACacheWithoutRefreshMetadata) {
  pivot::PivotCache cache = FirstCache("B01_base");
  cache.mutable_passthrough_attrs().clear();
  ExpectNotWritable(cache);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
