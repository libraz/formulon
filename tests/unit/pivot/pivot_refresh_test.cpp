#include "pivot/pivot_refresh.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "merge_range.h"
#include "pivot/field_grouping.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/shared_items_hints.h"
#include "sheet.h"
#include "styles.h"
#include "table.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using pivot::FieldRemap;
using pivot::PivotAxis;
using pivot::PivotCache;
using pivot::PivotCacheField;
using pivot::PivotCacheRecord;
using pivot::PivotDataField;
using pivot::PivotField;
using pivot::PivotItem;
using pivot::PivotTable;
using pivot::RefreshContext;
using pivot::RefreshPlan;
using pivot::RefreshSource;
using pivot::SharedItemsHints;

constexpr std::size_t kDataSheet = 0U;
constexpr std::size_t kPivotSheet = 1U;
constexpr std::uint32_t kCacheId = 1U;
constexpr std::uint32_t kRegion = 0U;
constexpr std::uint32_t kAmt = 2U;

// 2026-01-01 and 2026-03-01 as 1900-system serials.
constexpr double kJan1 = 46023.0;
constexpr double kMar1 = 46082.0;

std::optional<std::uint32_t> At(std::uint32_t v) {
  return v;
}

PivotItem Item(std::uint32_t x, std::string name = {}) {
  PivotItem item;
  item.has_cache_index = true;
  item.cache_index = x;
  item.name = std::move(name);
  return item;
}

PivotField Field(std::string name, PivotAxis axis = PivotAxis::None) {
  PivotField field;
  field.source_name = std::move(name);
  field.axis = axis;
  return field;
}

bool HasAttr(const PivotItem& item, const std::string& name) {
  for (const auto& attr : item.passthrough_attrs) {
    if (attr.first == name && attr.second == "1") {
      return true;
    }
  }
  return false;
}

std::vector<std::string> ItemTexts(const PivotCacheField& field) {
  std::vector<std::string> out;
  for (const Value& v : field.shared_items) {
    out.emplace_back(v.is_text() ? std::string(v.as_text()) : std::string("#"));
  }
  return out;
}

// --- Source and pure-build fixtures ----------------------------------------

RefreshSource Source(std::vector<std::string> headers, const std::vector<std::vector<Value>>& rows,
                     const std::vector<std::vector<bool>>& dates = {}) {
  RefreshSource source;
  source.columns = static_cast<std::uint32_t>(headers.size());
  source.rows = static_cast<std::uint32_t>(rows.size());
  source.headers = std::move(headers);
  source.num_fmt_ids.assign(source.columns, 0U);
  for (std::size_t r = 0; r < rows.size(); ++r) {
    for (std::size_t c = 0; c < source.columns; ++c) {
      source.cells.push_back(c < rows[r].size() ? rows[r][c] : Value::blank());
      source.cell_is_date.push_back(r < dates.size() && c < dates[r].size() && dates[r][c]);
    }
  }
  return source;
}

RefreshContext Context(std::vector<bool> axis_fields, ExcelProfile profile = mac_365_ja_jp_profile()) {
  RefreshContext context;
  context.axis_fields = std::move(axis_fields);
  context.blank_header_word = pivot::blank_header_field_word(profile);
  context.profile = profile;
  return context;
}

PivotCache RegionCache(std::vector<Value> items) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  cache.mutable_fields().push_back(PivotCacheField{"Region", std::move(items)});
  return cache;
}

RefreshPlan Build(const PivotCache& old, const RefreshSource& source, const RefreshContext& context) {
  auto plan = pivot::build_refreshed_cache(old, source, context);
  EXPECT_TRUE(static_cast<bool>(plan)) << (plan ? std::string() : plan.error().message);
  return std::move(plan.value());
}

// --- Workbook fixture --------------------------------------------------------

void SetText(Workbook& wb, std::uint32_t row, std::uint32_t col, const char* text) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kDataSheet, row, col, text)));
}

void SetNumber(Workbook& wb, std::uint32_t row, std::uint32_t col, double v) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(kDataSheet, row, col, Value::number(v))));
}

// The cache the base source reads to: Region and Product shared, Amt inline.
PivotCache BaseCache(const std::string& sheet_name) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  cache.mutable_fields().push_back(
      PivotCacheField{"Region", {Value::text("North"), Value::text("South"), Value::text("East")}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {Value::text("A"), Value::text("B")}});
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  const double rows[][3] = {{0, 0, 100}, {1, 1, 200}, {2, 0, 300}, {0, 1, 400}, {1, 0, 500}};
  for (const auto& row : rows) {
    PivotCacheRecord record;
    record.cells = {Value::number(row[0]), Value::number(row[1]), Value::number(row[2])};
    record.cell_is_index = {true, true, false};
    cache.mutable_records().push_back(std::move(record));
  }
  pivot::WorksheetSource& source = cache.mutable_worksheet_source();
  source.present = true;
  source.ref = "A1:C6";
  source.sheet = sheet_name;
  return cache;
}

std::unique_ptr<PivotTable> RegionPivot(const char* name, std::uint32_t row, std::uint32_t col) {
  auto table = std::make_unique<PivotTable>();
  table->set_name(name);
  table->set_pivot_cache_id(kCacheId);
  table->mutable_fields().push_back(Field("Region", PivotAxis::Row));
  table->mutable_fields().push_back(Field("Product"));
  table->mutable_fields().push_back(Field("Amt", PivotAxis::Value));
  for (std::uint32_t x = 0; x < 3U; ++x) {
    table->mutable_fields()[kRegion].items.push_back(Item(x));
  }
  table->mutable_row_field_order() = {kRegion};
  PivotDataField sum;
  sum.name = "Sum of Amt";
  sum.field_index = kAmt;
  table->mutable_data_fields().push_back(std::move(sum));
  table->set_anchor(row, col, 1U, 1U);
  return table;
}

// Region / Product / Amt on the first sheet, one or two Region pivots on `Pivot`.
Workbook Book(int pivots = 1, ExcelProfile profile = mac_365_en_us_profile()) {
  Workbook wb = Workbook::create_for_profile(profile);
  const char* regions[] = {"North", "South", "East", "North", "South"};
  const char* products[] = {"A", "B", "A", "B", "A"};
  SetText(wb, 0U, 0U, "Region");
  SetText(wb, 0U, 1U, "Product");
  SetText(wb, 0U, 2U, "Amt");
  for (std::uint32_t r = 0; r < 5U; ++r) {
    SetText(wb, r + 1U, 0U, regions[r]);
    SetText(wb, r + 1U, 1U, products[r]);
    SetNumber(wb, r + 1U, 2U, 100.0 * (r + 1U));
  }
  wb.add_pivot_cache(std::make_unique<PivotCache>(BaseCache(wb.sheet(kDataSheet).name())));
  EXPECT_TRUE(static_cast<bool>(wb.add_sheet_entered("Pivot")));
  EXPECT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, RegionPivot("PT1", 2U, 0U))));
  EXPECT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 0U)));
  if (pivots > 1) {
    EXPECT_TRUE(static_cast<bool>(wb.add_pivot(kPivotSheet, RegionPivot("PT2", 2U, 6U))));
    EXPECT_TRUE(static_cast<bool>(wb.project_pivot(kPivotSheet, 1U)));
  }
  return wb;
}

PivotCache& CacheOf(Workbook& wb) {
  return *wb.mutable_pivot_caches().front();
}

PivotTable& PivotAt(Workbook& wb, std::size_t index = 0U) {
  return *wb.sheet(kPivotSheet).mutable_pivot_tables()[index];
}

Expected<void, Error> Refresh(Workbook& wb) {
  return wb.refresh_pivot_cache(kCacheId, eval::default_registry());
}

// What a refusal must leave as it was: the cache object, its fields and
// records, the pivots' definitions and the cells of their reports.
struct Snapshot {
  const PivotCache* cache = nullptr;
  std::vector<std::string> field_names;
  std::size_t records = 0;
  std::vector<std::size_t> item_counts;
  std::vector<std::vector<std::string>> report;

  static Snapshot Of(Workbook& wb) {
    Snapshot s;
    s.cache = wb.find_pivot_cache(kCacheId);
    for (const PivotCacheField& field : s.cache->fields()) {
      s.field_names.push_back(field.name);
    }
    s.records = s.cache->records().size();
    for (const auto& table : wb.sheet(kPivotSheet).pivot_tables()) {
      s.item_counts.push_back(table->fields()[kRegion].items.size());
    }
    const Sheet& sheet = wb.sheet(kPivotSheet);
    for (std::uint32_t r = 0; r < 16U; ++r) {
      std::vector<std::string> row;
      for (std::uint32_t c = 0; c < 12U; ++c) {
        const Value v = sheet.resolve_cell_value(r, c);
        row.push_back(v.is_text() ? std::string(v.as_text())
                                  : (v.is_number() ? std::to_string(static_cast<long long>(v.as_number())) : ""));
      }
      s.report.push_back(std::move(row));
    }
    return s;
  }

  bool operator==(const Snapshot& o) const {
    return cache == o.cache && field_names == o.field_names && records == o.records && item_counts == o.item_counts &&
           report == o.report;
  }
};

void ExpectRefused(Workbook& wb, FormulonErrorCode code) {
  const Snapshot before = Snapshot::Of(wb);
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_FALSE(static_cast<bool>(refreshed));
  EXPECT_EQ(refreshed.error().code, code) << refreshed.error().message;
  EXPECT_TRUE(Snapshot::Of(wb) == before) << "a refused refresh changed the workbook";
}

bool ReportShows(Workbook& wb, std::size_t pivot_index, double value) {
  const PivotTable& table = PivotAt(wb, pivot_index);
  if (!table.projected_extent().has_value()) {
    return false;
  }
  const pivot::PivotRect& r = *table.projected_extent();
  for (std::uint32_t row = r.row0; row <= r.row1; ++row) {
    for (std::uint32_t col = r.col0; col <= r.col1; ++col) {
      const Value v = wb.sheet(kPivotSheet).resolve_cell_value(row, col);
      if (v.is_number() && v.as_number() == value) {
        return true;
      }
    }
  }
  return false;
}

// ===========================================================================
// Rules: the measured header, typing, item and reconciliation rules
// ===========================================================================

// Saves `wb` as .xlsx and reads it back.
Workbook SaveAndReload(const Workbook& wb) {
  auto bytes = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(bytes));
  auto read = io::read_ooxml(io::ByteSpan{bytes.value().data(), bytes.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return std::move(read.value().workbook);
}

TEST(PivotRefresh, SaveHandFilledCacheAsksExcelToRefreshOnLoad) {
  Workbook wb = Book();
  Workbook saved = SaveAndReload(wb);
  EXPECT_TRUE(saved.pivot_caches().front()->refresh_on_load());
  // The built table still saves the lines of its projected report.
  EXPECT_NE(saved.sheet(kPivotSheet).pivot_tables().front()->raw_passthrough_after_row_fields().find("<rowItems"),
            std::string::npos);
}

TEST(PivotRefresh, SaveRefreshedCacheNeedsNoRefreshOnLoad) {
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  ASSERT_TRUE(CacheOf(wb).records_from_source());
  Workbook saved = SaveAndReload(wb);
  EXPECT_FALSE(saved.pivot_caches().front()->refresh_on_load());
  EXPECT_NE(saved.sheet(kPivotSheet).pivot_tables().front()->raw_passthrough_after_row_fields().find("<rowItems"),
            std::string::npos);
}

TEST(PivotRefresh, AFieldLeftWithoutSharedItemsListsNoItems) {
  Workbook wb = Book();
  PivotTable& table = PivotAt(wb);
  // Product sits on no axis, so a refresh drops its shared items.
  table.mutable_fields()[1].items = {Item(0U), Item(1U)};
  pivot::PivotMarkerItem marker;
  marker.type = "default";
  marker.after_items = 2U;
  table.mutable_fields()[1].marker_items.push_back(marker);
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  ASSERT_TRUE(CacheOf(wb).fields()[1].shared_items.empty());
  EXPECT_TRUE(PivotAt(wb).fields()[1].items.empty());
  EXPECT_TRUE(PivotAt(wb).fields()[1].marker_items.empty());
}

TEST(PivotRefresh, EditingARefreshedCacheMakesItHandFilled) {
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  CacheOf(wb).mutable_records().front().cells[2] = Value::number(1.0);
  EXPECT_FALSE(CacheOf(wb).records_from_source());
}

TEST(PivotRefresh, RulesBlankHeaderTakesLocaleWordAfterLiterals) {
  EXPECT_EQ(pivot::refreshed_field_names({"", "フィールド1"}, "フィールド"),
            (std::vector<std::string>{"フィールド12", "フィールド1"}));
  EXPECT_EQ(pivot::refreshed_field_names({"Region", ""}, "Field"), (std::vector<std::string>{"Region", "Field2"}));
}

TEST(PivotRefresh, RulesDuplicateHeadersTakeSmallestFreeSuffix) {
  const auto names = [](std::vector<std::string> headers) { return pivot::refreshed_field_names(headers, "Field"); };
  EXPECT_EQ(names({"Dup", "Dup", "dup"}), (std::vector<std::string>{"Dup", "Dup2", "dup3"}));
  EXPECT_EQ(names({"Dup", "Dup", "Dup2"}), (std::vector<std::string>{"Dup", "Dup2", "Dup22"}));
  EXPECT_EQ(names({"Amt", "Amt", "Amt"}), (std::vector<std::string>{"Amt", "Amt2", "Amt3"}));
  EXPECT_EQ(names({"Amt", "Amt "}), (std::vector<std::string>{"Amt", "Amt "}));
}

TEST(PivotRefresh, RulesHeaderFoldIgnoresCaseWidthAndKanaButNotSpaces) {
  EXPECT_TRUE(pivot::same_header("A", "Ａ"));
  EXPECT_TRUE(pivot::same_header("a", "Ａ"));
  EXPECT_TRUE(pivot::same_header("ア", "ｱ"));
  EXPECT_TRUE(pivot::same_header("ア", "あ"));
  EXPECT_TRUE(pivot::same_header("Dup", "dup"));
  EXPECT_FALSE(pivot::same_header("Amt", "Amt "));
}

TEST(PivotRefresh, RulesBlankHeaderTakesTheLocaleFieldWordAndColumn) {
  const std::vector<std::pair<ExcelLocale, std::string>> expected = {
      {ExcelLocale::kJaJP, "フィールド2"}, {ExcelLocale::kEnUS, "Field2"}, {ExcelLocale::kDeDE, "Feld2"},
      {ExcelLocale::kFrFR, "Champ2"},      {ExcelLocale::kEsES, "Campo2"}, {ExcelLocale::kEsMX, "Campo2"},
      {ExcelLocale::kItIT, "Campo2"},      {ExcelLocale::kPtBR, "Campo2"}, {ExcelLocale::kZhCN, "字段2"},
      {ExcelLocale::kZhTW, "欄位2"},       {ExcelLocale::kKoKR, "필드2"},  {ExcelLocale::kThTH, "เขตข้อมูล2"},
      {ExcelLocale::kRuRU, "Поле2"},       {ExcelLocale::kNlNL, "Veld2"},
  };
  for (const auto& [locale, name] : expected) {
    for (const ExcelHost host : {ExcelHost::kMac365, ExcelHost::kWin365}) {
      const auto names = pivot::refreshed_field_names({"Region", ""}, pivot::blank_header_field_word({host, locale}));
      EXPECT_EQ(names, (std::vector<std::string>{"Region", name})) << static_cast<int>(locale);
    }
  }
}

TEST(PivotRefresh, RulesItemFoldIsCaseOnlyAndKeepsFirstSpelling) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const PivotCache old = RegionCache({Value::text("North")});
  const RefreshSource source = Source({"Region"}, {{Value::text("NORTH")},
                                                   {Value::text("North")},
                                                   {Value::text("ｎｏｒｔｈ")},
                                                   {Value::number(1)},
                                                   {Value::text("1")}});
  const RefreshPlan plan = Build(old, source, Context({true}));
  const PivotCacheField& field = plan.cache.fields()[0];
  ASSERT_EQ(field.shared_items.size(), 4U);
  EXPECT_EQ(field.shared_items[0].as_text(), "NORTH");
  EXPECT_EQ(field.shared_items[1].as_text(), "ｎｏｒｔｈ");
  EXPECT_TRUE(field.shared_items[2].is_number());
  EXPECT_EQ(field.shared_items[3].as_text(), "1");
  EXPECT_EQ(plan.remap.items[0], (std::vector<std::optional<std::uint32_t>>{At(0U)}));
  EXPECT_EQ(plan.cache.records()[1].cells[0].as_number(), 0.0);
}

TEST(PivotRefresh, RulesDateAndSameSerialNumberAreDistinctItems) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const RefreshSource source = Source({"Region"}, {{Value::number(kJan1)}, {Value::number(kJan1)}}, {{true}, {false}});
  const RefreshPlan plan = Build(RegionCache({}), source, Context({true}));
  const PivotCacheField& field = plan.cache.fields()[0];
  ASSERT_EQ(field.shared_items.size(), 2U);
  EXPECT_EQ(field.shared_item_is_date, (std::vector<bool>{true, false}));
}

TEST(PivotRefresh, RulesMissingItemsFollowPresentOnesInOldOrder) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const PivotCache old = RegionCache({Value::text("North"), Value::text("South"), Value::text("East")});
  const RefreshPlan plan =
      Build(old, Source({"Region"}, {{Value::text("South")}, {Value::text("West")}}), Context({true}));
  const PivotCacheField& field = plan.cache.fields()[0];
  EXPECT_EQ(ItemTexts(field), (std::vector<std::string>{"South", "West", "North", "East"}));
  EXPECT_EQ(field.shared_item_is_missing, (std::vector<bool>{false, false, true, true}));
  EXPECT_EQ(plan.remap.items[0], (std::vector<std::optional<std::uint32_t>>{At(2U), At(0U), At(3U)}));
  EXPECT_EQ(plan.cache.records().size(), 2U);
}

TEST(PivotRefresh, RulesMissingItemsLimitZeroPurgesThem) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old = RegionCache({Value::text("North"), Value::text("South")});
  old.set_missing_items_limit(0U);
  const RefreshPlan plan = Build(old, Source({"Region"}, {{Value::text("South")}}), Context({true}));
  EXPECT_EQ(ItemTexts(plan.cache.fields()[0]), (std::vector<std::string>{"South"}));
  EXPECT_EQ(plan.remap.items[0], (std::vector<std::optional<std::uint32_t>>{std::nullopt, At(0U)}));
  EXPECT_EQ(plan.cache.missing_items_limit(), std::optional<std::uint32_t>(0U));
}

TEST(PivotRefresh, RulesCalculatedItemsComeLastAndKeepWhatTheyName) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old = RegionCache({Value::text("North"), Value::text("South"), Value::text("NS")});
  old.mutable_fields()[0].shared_item_is_calculated = {false, false, true};
  old.set_missing_items_limit(0U);
  const RefreshPlan plan = Build(old, Source({"Region"}, {{Value::text("South")}}), Context({true}));
  const PivotCacheField& field = plan.cache.fields()[0];
  EXPECT_EQ(ItemTexts(field), (std::vector<std::string>{"South", "North", "NS"}));
  EXPECT_EQ(field.shared_item_is_missing, (std::vector<bool>{false, true, false}));
  EXPECT_EQ(field.shared_item_is_calculated, (std::vector<bool>{false, false, true}));
}

TEST(PivotRefresh, RulesOnlyAxisFieldsKeepSharedItems) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old = RegionCache({Value::text("North")});
  old.mutable_fields().push_back(PivotCacheField{"Product", {Value::text("A")}});
  const RefreshPlan plan =
      Build(old, Source({"Region", "Product"}, {{Value::text("North"), Value::text("A")}}), Context({true, false}));
  EXPECT_EQ(plan.cache.fields()[0].shared_items.size(), 1U);
  EXPECT_TRUE(plan.cache.fields()[1].shared_items.empty());
  EXPECT_TRUE(plan.cache.records()[0].cell_is_index[0]);
  EXPECT_FALSE(plan.cache.records()[0].cell_is_index[1]);
  EXPECT_EQ(plan.cache.records()[0].cells[1].as_text(), "A");
  EXPECT_EQ(plan.remap.items[1], (std::vector<std::optional<std::uint32_t>>{std::nullopt}));
}

TEST(PivotRefresh, RulesFieldsMatchByHeaderFoldAndFollowSourceOrder) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old = RegionCache({Value::text("North")});
  old.mutable_fields().push_back(PivotCacheField{"Product", {}});
  old.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  const RefreshSource source =
      Source({"Amt", "ＲＥＧＩＯＮ", "Prod"}, {{Value::number(1), Value::text("North"), Value::text("A")}});
  const RefreshPlan plan = Build(old, source, Context({true, false, false}));
  EXPECT_EQ(plan.remap.fields, (std::vector<std::optional<std::uint32_t>>{At(1U), std::nullopt, At(0U)}));
  EXPECT_EQ(plan.remap.names, (std::vector<std::string>{"Amt", "ＲＥＧＩＯＮ", "Prod"}));
}

TEST(PivotRefresh, RulesNonIntegerNumericGroupIsNotImplemented) {
  PivotCache old = RegionCache({});
  PivotCacheField group;
  group.name = "Amt2";
  group.is_database_field = false;
  group.field_group = pivot::FieldGroup{};
  group.field_group->base = 0U;
  group.field_group->kind = pivot::GroupKind::Range;
  group.field_group->range.start_num = 0.0;
  group.field_group->range.end_num = 400.0;
  group.field_group->range.interval = 100.0;
  old.mutable_fields().push_back(std::move(group));
  const auto plan = pivot::build_refreshed_cache(old, Source({"Region"}, {{Value::number(5.5)}}), Context({true}));
  ASSERT_FALSE(static_cast<bool>(plan));
  EXPECT_EQ(plan.error().code, FormulonErrorCode::kNotImplemented);
}

TEST(PivotRefresh, RulesRefreshReadsCurrentValuesAndReprojects) {
  Workbook wb = Book();
  SetNumber(wb, 1U, 2U, 1000.0);  // North 100 -> 1000
  SetText(wb, 3U, 0U, "West");    // East -> West
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  const PivotCache& cache = *wb.find_pivot_cache(kCacheId);
  EXPECT_EQ(ItemTexts(cache.fields()[kRegion]), (std::vector<std::string>{"North", "South", "West", "East"}));
  EXPECT_EQ(cache.fields()[kRegion].shared_item_is_missing, (std::vector<bool>{false, false, false, true}));
  EXPECT_EQ(cache.records()[0].cells[kAmt].as_number(), 1000.0);
  EXPECT_FALSE(PivotAt(wb).projection_pending());
  EXPECT_TRUE(ReportShows(wb, 0U, 1400.0));  // North: 1000 + 400
  const std::vector<PivotItem>& items = PivotAt(wb).fields()[kRegion].items;
  ASSERT_EQ(items.size(), 4U);
  EXPECT_TRUE(HasAttr(items[2], "m"));  // East, kept in place as missing
  EXPECT_EQ(items[3].cache_index, 2U);  // West, appended
}

TEST(PivotRefresh, RulesHeaderDisplayTextNamesTheField) {
  Workbook wb = Book();
  SetNumber(wb, 0U, 2U, 123.0);
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  const PivotCache& cache = *wb.find_pivot_cache(kCacheId);
  EXPECT_EQ(cache.fields()[kAmt].name, "123");
  // The renamed column is a new field, and the data field on Amt went with it.
  EXPECT_TRUE(PivotAt(wb).data_fields().empty());
  EXPECT_EQ(PivotAt(wb).fields()[kAmt].axis, PivotAxis::None);
}

TEST(PivotRefresh, RulesDateTypingAndNumFmtFollowTheCellFormat) {
  Workbook wb = Book();
  CellXf xf = wb.styles().cell_xfs.empty() ? CellXf{} : wb.styles().cell_xfs.front();
  xf.num_fmt_id = 14U;
  xf.apply_number_format = true;
  const std::uint32_t date_xf = intern_cell_xf(wb.mutable_styles(), xf);
  for (std::uint32_t r = 1; r <= 5U; ++r) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_xf_index(kDataSheet, r, 2U, date_xf)));
  }
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  const PivotCache& cache = *wb.find_pivot_cache(kCacheId);
  const PivotCacheField& amt = cache.fields()[kAmt];
  EXPECT_EQ(amt.num_fmt_id, std::optional<std::uint32_t>(14U));
  EXPECT_TRUE(amt.shared_items_hints.contains_date);
  EXPECT_TRUE(cache.records()[0].cell_is_date[kAmt]);
  EXPECT_EQ(cache.fields()[kRegion].num_fmt_id, std::optional<std::uint32_t>(0U));
}

// ===========================================================================
// Remap: every index- or name-keyed piece of a table definition
// ===========================================================================

PivotCache NamedCache(const std::vector<std::string>& names) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  for (const std::string& name : names) {
    cache.mutable_fields().push_back(PivotCacheField{name, {}});
  }
  return cache;
}

TEST(PivotRefresh, RemapFieldListPlacesNewFieldsUnplaced) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Row), Field("B"), Field("C", PivotAxis::Value)};
  FieldRemap remap;
  remap.fields = {At(1U), std::nullopt, At(0U)};
  remap.names = {"C", "A", "D"};
  pivot::remap_table(table, NamedCache({"C", "A", "D"}), remap);
  ASSERT_EQ(table.fields().size(), 3U);
  EXPECT_EQ(table.fields()[0].source_name, "C");
  EXPECT_EQ(table.fields()[0].axis, PivotAxis::Value);
  EXPECT_EQ(table.fields()[1].axis, PivotAxis::Row);
  EXPECT_EQ(table.fields()[2].source_name, "D");
  EXPECT_EQ(table.fields()[2].axis, PivotAxis::None);
}

TEST(PivotRefresh, RemapAxisOrdersAndValuesPosition) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Row), Field("B", PivotAxis::Row), Field("C", PivotAxis::Col)};
  table.mutable_row_field_order() = {0U, 1U};
  table.set_row_values_position(2U);
  table.mutable_col_field_order() = {2U};
  FieldRemap remap;
  remap.fields = {std::nullopt, At(1U), At(0U)};
  pivot::remap_table(table, NamedCache({"C", "B"}), remap);
  EXPECT_EQ(table.row_field_order(), (std::vector<std::uint32_t>{1U}));
  EXPECT_EQ(table.row_values_position(), std::optional<std::size_t>(1U));
  EXPECT_EQ(table.col_field_order(), (std::vector<std::uint32_t>{0U}));
}

TEST(PivotRefresh, RemapPageFieldAndItsSelectedItem) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A"), Field("B")};
  table.set_field_axis(1U, PivotAxis::Page);
  table.mutable_fields()[1].items = {Item(0U), Item(1U)};
  table.mutable_page_fields()[0].item_index = 1U;
  PivotCache cache = NamedCache({"B", "A"});
  cache.mutable_fields()[0].shared_items = {Value::text("y")};
  FieldRemap remap;
  remap.fields = {At(1U), At(0U)};
  remap.items = {{}, {std::nullopt, At(0U)}};
  pivot::remap_table(table, cache, remap);
  ASSERT_EQ(table.page_fields().size(), 1U);
  EXPECT_EQ(table.page_fields()[0].field_index, 0U);
  EXPECT_EQ(table.page_fields()[0].item_index, std::optional<std::uint32_t>(0U));
}

TEST(PivotRefresh, RemapDataFieldsAndDroppedFieldAggregates) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Value), Field("B", PivotAxis::Value)};
  PivotDataField on_a;
  on_a.name = "Sum of A";
  on_a.field_index = 0U;
  PivotDataField on_b;
  on_b.name = "Sum of B";
  on_b.field_index = 1U;
  table.mutable_data_fields() = {on_a, on_b};
  FieldRemap remap;
  remap.fields = {std::nullopt, At(0U)};
  pivot::remap_table(table, NamedCache({"B"}), remap);
  ASSERT_EQ(table.data_fields().size(), 1U);
  EXPECT_EQ(table.data_fields()[0].name, "Sum of B");
  EXPECT_EQ(table.data_fields()[0].field_index, 0U);
}

TEST(PivotRefresh, RemapShowAsBaseFieldAndItem) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("V", PivotAxis::Value), Field("R", PivotAxis::Row)};
  table.mutable_fields()[1].items = {Item(0U), Item(1U)};
  PivotDataField diff;
  diff.name = "Diff";
  diff.field_index = 0U;
  diff.show_as = pivot::ShowValuesAs::DifferenceFrom;
  diff.show_as_base_field = 1U;
  diff.show_as_base_item = 1U;
  PivotDataField prev = diff;
  prev.name = "Prev";
  prev.show_as_base_item = pivot::kShowAsBasePrev;
  table.mutable_data_fields() = {diff, prev};
  PivotCache cache = NamedCache({"R", "V"});
  cache.mutable_fields()[0].shared_items = {Value::text("y")};
  FieldRemap remap;
  remap.fields = {At(1U), At(0U)};
  remap.items = {{}, {std::nullopt, At(0U)}};
  pivot::remap_table(table, cache, remap);
  EXPECT_EQ(table.data_fields()[0].field_index, 1U);
  EXPECT_EQ(table.data_fields()[0].show_as_base_field, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(table.data_fields()[0].show_as_base_item, std::optional<std::uint32_t>(0U));
  EXPECT_EQ(table.data_fields()[1].show_as_base_item, std::optional<std::uint32_t>(pivot::kShowAsBasePrev));
}

TEST(PivotRefresh, RemapFilterSourceDropsWhenItsFieldMoves) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Row), Field("B", PivotAxis::Row), Field("C")};
  pivot::PivotFilter stays;
  stays.field_name = "A";
  stays.type = pivot::FilterType::LabelContains;
  stays.source = pivot::PivotFilterSource{"<filter fld=\"0\"/>", 0U, 0U, 1U};
  pivot::PivotFilter moves = stays;
  moves.field_name = "B";
  moves.source = pivot::PivotFilterSource{"<filter fld=\"1\"/>", 1U, 0U, 2U};
  table.mutable_active_filters() = {stays, moves};
  FieldRemap remap;
  remap.fields = {At(0U), At(2U), At(1U)};
  pivot::remap_table(table, NamedCache({"A", "C", "B"}), remap);
  ASSERT_EQ(table.active_filters().size(), 2U);
  EXPECT_TRUE(table.active_filters()[0].source.has_value());
  EXPECT_FALSE(table.active_filters()[1].source.has_value());
}

TEST(PivotRefresh, RemapFilterFieldNameFollowsTheNewSpelling) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("region", PivotAxis::Row)};
  pivot::PivotFilter filter;
  filter.field_name = "region";
  filter.type = pivot::FilterType::LabelContains;
  table.mutable_active_filters() = {filter};
  FieldRemap remap;
  remap.fields = {At(0U)};
  remap.names = {"Region"};
  pivot::remap_table(table, NamedCache({"Region"}), remap);
  EXPECT_EQ(table.active_filters()[0].field_name, "Region");
}

TEST(PivotRefresh, RemapSortByFieldFollowsTheNewSpelling) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("R", PivotAxis::Row), Field("amt", PivotAxis::Value)};
  table.mutable_fields()[0].sort.by_field = "amt";
  FieldRemap remap;
  remap.fields = {At(0U), At(1U)};
  remap.names = {"R", "Amt"};
  pivot::remap_table(table, NamedCache({"R", "Amt"}), remap);
  EXPECT_EQ(table.fields()[0].sort.by_field, "Amt");
}

TEST(PivotRefresh, RemapSourceNameTakesNewSpellingAndCustomNameStays) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  PivotField field = Field("region", PivotAxis::Row);
  field.custom_name = "Reg";
  table.mutable_fields() = {field};
  FieldRemap remap;
  remap.fields = {At(0U)};
  remap.names = {"Region"};
  pivot::remap_table(table, NamedCache({"Region"}), remap);
  EXPECT_EQ(table.fields()[0].source_name, "Region");
  EXPECT_EQ(table.fields()[0].custom_name, "Reg");
}

// Old items East(2), North(0) "N!", South(1) hidden; South and West present,
// North and East kept missing.
PivotCache RegionAfterRefresh() {
  PivotCache cache =
      RegionCache({Value::text("South"), Value::text("West"), Value::text("North"), Value::text("East")});
  cache.mutable_fields()[0].shared_item_is_missing = {false, false, true, true};
  return cache;
}

PivotTable RegionBeforeRefresh() {
  PivotTable table;
  table.mutable_fields() = {Field("Region", PivotAxis::Row)};
  table.mutable_row_field_order() = {0U};
  PivotItem north = Item(0U, "North");
  north.caption = "N!";
  PivotItem south = Item(1U, "South");
  south.visible = false;
  table.mutable_fields()[0].items = {Item(2U, "East"), north, south};
  return table;
}

FieldRemap RegionItemRemap() {
  FieldRemap remap;
  remap.fields = {At(0U)};
  remap.items = {{At(2U), At(0U), At(3U)}};
  return remap;
}

TEST(PivotRefresh, RemapPivotItemsKeepOrderMarkMissingAndAppendNewHidden) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = RegionBeforeRefresh();
  pivot::remap_table(table, RegionAfterRefresh(), RegionItemRemap());
  const std::vector<PivotItem>& items = table.fields()[0].items;
  ASSERT_EQ(items.size(), 4U);
  EXPECT_EQ(items[0].cache_index, 3U);
  EXPECT_TRUE(HasAttr(items[0], "m"));
  EXPECT_EQ(items[1].cache_index, 2U);
  EXPECT_EQ(items[1].caption, "N!");
  EXPECT_TRUE(HasAttr(items[1], "m"));
  EXPECT_EQ(items[2].cache_index, 0U);
  EXPECT_FALSE(items[2].visible);
  EXPECT_FALSE(HasAttr(items[2], "m"));
  EXPECT_EQ(items[3].cache_index, 1U);  // West
  EXPECT_FALSE(items[3].visible);       // the field still hides South
}

TEST(PivotRefresh, RemapSortedFieldIsResortedWithMissingItems) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table = RegionBeforeRefresh();
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[0].sort.stated = true;
  pivot::remap_table(table, RegionAfterRefresh(), RegionItemRemap());
  std::vector<std::uint32_t> order;
  for (const PivotItem& item : table.fields()[0].items) {
    order.push_back(item.cache_index);
  }
  // West, South, North, East.
  EXPECT_EQ(order, (std::vector<std::uint32_t>{1U, 0U, 2U, 3U}));
}

TEST(PivotRefresh, RemapDroppedFieldLeavesAxesPagesAndFilters) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Row), Field("B")};
  table.set_field_axis(1U, PivotAxis::Page);
  table.mutable_row_field_order() = {0U};
  pivot::PivotFilter filter;
  filter.field_name = "B";
  filter.type = pivot::FilterType::LabelContains;
  table.mutable_active_filters() = {filter};
  FieldRemap remap;
  remap.fields = {At(0U), std::nullopt};
  pivot::remap_table(table, NamedCache({"A"}), remap);
  EXPECT_EQ(table.fields().size(), 1U);
  EXPECT_TRUE(table.page_fields().empty());
  EXPECT_TRUE(table.active_filters().empty());
}

TEST(PivotRefresh, RemapCalculatedItemFieldAndAreaReferences) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old = RegionCache({Value::text("North"), Value::text("South"), Value::text("NS")});
  old.mutable_fields()[0].shared_item_is_calculated = {false, false, true};
  old.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  pivot::PivotCalculatedItem entry;
  entry.cache_field_index = 0U;
  entry.formula = "Region[North]+Region[South]";
  entry.area.field_position = 0U;
  pivot::PivotAreaReference reference;
  reference.field = 0U;
  reference.items = {2U};
  entry.area.references.push_back(reference);
  old.mutable_calculated_items().push_back(entry);
  const RefreshPlan plan =
      Build(old, Source({"Amt", "Region"}, {{Value::number(1), Value::text("South")}}), Context({true, false}));
  ASSERT_EQ(plan.cache.calculated_items().size(), 1U);
  const pivot::PivotCalculatedItem& moved = plan.cache.calculated_items()[0];
  EXPECT_EQ(moved.cache_field_index, 1U);
  ASSERT_EQ(moved.area.references.size(), 1U);
  EXPECT_EQ(moved.area.references[0].field, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(moved.area.references[0].items, (std::vector<std::uint32_t>{2U}));  // South, North(u), NS(f)
}

TEST(PivotRefresh, RemapPivotAreaFieldPosition) {
  PivotCache cache = RegionCache({Value::text("North")});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  pivot::PivotCalculatedItem entry;
  entry.cache_field_index = 1U;
  entry.area.field_position = 0U;
  cache.mutable_calculated_items().push_back(entry);
  PivotTable table;
  table.mutable_fields() = {Field("Region", PivotAxis::Row), Field("Product", PivotAxis::Row)};
  table.mutable_row_field_order() = {0U, 1U};
  pivot::place_calculated_item_areas(cache, table);
  EXPECT_EQ(cache.calculated_items()[0].area.field_position, std::optional<std::uint32_t>(1U));
}

TEST(PivotRefresh, RemapFieldGroupParAndBase) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old;
  old.set_cache_id(kCacheId);
  PivotCacheField date{"Date", {}};
  date.field_group = pivot::FieldGroup{};
  date.field_group->par = 2U;
  old.mutable_fields().push_back(std::move(date));
  old.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  PivotCacheField months;
  months.name = "Months";
  months.is_database_field = false;
  months.field_group = pivot::FieldGroup{};
  months.field_group->base = 0U;
  months.field_group->kind = pivot::GroupKind::Range;
  months.field_group->range.by = pivot::GroupBy::Months;
  months.field_group->range.is_date = true;
  months.field_group->range.start_date = "2026-01-01T00:00:00";
  months.field_group->range.end_date = "2026-03-02T00:00:00";
  for (const char* label : {"<2026/1/1", "1月", ">2026/3/2"}) {
    months.field_group->group_items.push_back(Value::text(label));
  }
  old.mutable_fields().push_back(std::move(months));
  const RefreshPlan plan = Build(
      old,
      Source({"Amt", "Date"}, {{Value::number(1), Value::number(kJan1)}, {Value::number(2), Value::number(kMar1)}},
             {{false, true}, {false, true}}),
      Context({false, false, true}));
  ASSERT_EQ(plan.cache.fields().size(), 3U);
  EXPECT_EQ(plan.cache.fields()[1].field_group->par, std::optional<std::uint32_t>(2U));
  EXPECT_EQ(plan.cache.fields()[2].field_group->base, std::optional<std::uint32_t>(1U));
  EXPECT_EQ(plan.remap.fields, (std::vector<std::optional<std::uint32_t>>{At(1U), At(0U), At(2U)}));
}

TEST(PivotRefresh, RemapDateGroupBoundsAndBoundaryLabelsFollowTheData) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old;
  old.set_cache_id(kCacheId);
  old.mutable_fields().push_back(PivotCacheField{"Date", {}});
  PivotCacheField months;
  months.name = "Months";
  months.is_database_field = false;
  months.field_group = pivot::FieldGroup{};
  months.field_group->base = 0U;
  months.field_group->kind = pivot::GroupKind::Range;
  months.field_group->range.by = pivot::GroupBy::Months;
  months.field_group->range.is_date = true;
  months.field_group->range.start_date = "2025-01-01T00:00:00";
  months.field_group->range.end_date = "2025-01-02T00:00:00";
  for (const char* label : {"<2025/1/1", "1月", ">2025/1/2"}) {
    months.field_group->group_items.push_back(Value::text(label));
  }
  old.mutable_fields().push_back(std::move(months));
  const RefreshPlan plan =
      Build(old, Source({"Date"}, {{Value::number(kJan1)}, {Value::number(kMar1)}}, {{true}, {true}}),
            Context({false, true}));
  const pivot::FieldGroup& group = *plan.cache.fields()[1].field_group;
  EXPECT_EQ(group.range.start_date, "2026-01-01T00:00:00");
  EXPECT_EQ(group.range.end_date, "2026-03-02T00:00:00");
  EXPECT_EQ(group.group_items.front().as_text(), "<2026/1/1");
  EXPECT_EQ(group.group_items.back().as_text(), ">2026/3/2");
}

TEST(PivotRefresh, RemapDiscreteGroupMapGivesNewItemsTheirOwnGroup) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotCache old;
  old.set_cache_id(kCacheId);
  old.mutable_fields().push_back(PivotCacheField{"Item", {Value::text("a"), Value::text("b")}});
  PivotCacheField grouped;
  grouped.name = "Item2";
  grouped.is_database_field = false;
  grouped.field_group = pivot::FieldGroup{};
  grouped.field_group->base = 0U;
  grouped.field_group->kind = pivot::GroupKind::Discrete;
  grouped.field_group->discrete_map = {0U, 0U};
  grouped.field_group->group_items = {Value::text("G1")};
  old.mutable_fields().push_back(std::move(grouped));
  const RefreshPlan plan =
      Build(old, Source({"Item"}, {{Value::text("b")}, {Value::text("c")}}), Context({false, true}));
  EXPECT_EQ(ItemTexts(plan.cache.fields()[0]), (std::vector<std::string>{"b", "c", "a"}));
  const pivot::FieldGroup& group = *plan.cache.fields()[1].field_group;
  EXPECT_EQ(group.discrete_map, (std::vector<std::uint32_t>{0U, 1U, 0U}));
  ASSERT_EQ(group.group_items.size(), 2U);
  EXPECT_EQ(group.group_items[1].as_text(), "c");
  EXPECT_EQ(plan.remap.items[1], (std::vector<std::optional<std::uint32_t>>{At(0U)}));
}

TEST(PivotRefresh, RemapRemovalOnlyKeepsItemsAsTheyAre) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  PivotTable table;
  table.mutable_fields() = {Field("A", PivotAxis::Row), Field("Calc"), Field("B", PivotAxis::Col)};
  table.mutable_fields()[0].items = {Item(1U), Item(0U)};
  table.mutable_row_field_order() = {0U};
  table.mutable_col_field_order() = {2U};
  PivotCache cache = NamedCache({"A", "B"});
  cache.mutable_fields()[0].shared_items = {Value::text("x"), Value::text("y"), Value::text("z")};
  FieldRemap remap;
  remap.fields = {At(0U), std::nullopt, At(1U)};
  pivot::remap_table(table, cache, remap);
  ASSERT_EQ(table.fields()[0].items.size(), 2U);
  EXPECT_EQ(table.fields()[0].items[0].cache_index, 1U);
  EXPECT_EQ(table.col_field_order(), (std::vector<std::uint32_t>{1U}));
}

// ===========================================================================
// Source: every source that does not resolve refuses with nothing changed
// ===========================================================================

TEST(PivotRefresh, SourceUnknownCacheIsNotFound) {
  Workbook wb = Book();
  const Expected<void, Error> refreshed = wb.refresh_pivot_cache(99U, eval::default_registry());
  ASSERT_FALSE(static_cast<bool>(refreshed));
  EXPECT_EQ(refreshed.error().code, FormulonErrorCode::kNotFound);
}

TEST(PivotRefresh, SourceAbsentIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().present = false;
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceInAnotherWorkbookIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().sheet = "[Other.xlsx]Data";
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceOfAnotherKindIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_cache_source_attrs() = {{"type", "external"}, {"connectionId", "1"}};
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceOnMissingSheetIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().sheet = "Gone";
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceNamingNothingIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().name = "NoSuchName";
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceNameDefinedAsRefErrorIsRefused) {
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Src", "#REF!")));
  CacheOf(wb).mutable_worksheet_source().name = "Src";
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceNameWithComputedBodyIsRefused) {
  Workbook wb = Book();
  const std::string sheet = wb.sheet(kDataSheet).name();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Src", "OFFSET(" + sheet + "!$A$1,0,0,6,3)")));
  CacheOf(wb).mutable_worksheet_source().name = "Src";
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
}

TEST(PivotRefresh, SourceRangeDoesNotGrow) {
  Workbook wb = Book();
  SetText(wb, 6U, 0U, "West");
  SetText(wb, 6U, 1U, "A");
  SetNumber(wb, 6U, 2U, 700.0);
  ASSERT_TRUE(static_cast<bool>(Refresh(wb)));
  EXPECT_EQ(wb.find_pivot_cache(kCacheId)->records().size(), 5U);
}

TEST(PivotRefresh, SourceDefinedNameFollowsItsCurrentDefinition) {
  Workbook wb = Book();
  SetText(wb, 6U, 0U, "West");
  SetText(wb, 6U, 1U, "A");
  SetNumber(wb, 6U, 2U, 700.0);
  const std::string sheet = wb.sheet(kDataSheet).name();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Src", sheet + "!$A$1:$C$7")));
  pivot::WorksheetSource& source = CacheOf(wb).mutable_worksheet_source();
  source.ref.clear();
  source.name = "Src";
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  EXPECT_EQ(wb.find_pivot_cache(kCacheId)->records().size(), 6U);
}

TEST(PivotRefresh, SourceTableUsesItsCurrentRange) {
  Workbook wb = Book();
  TableMetadata table;
  table.name = "Tbl_Src";
  table.display_name = "Tbl_Src";
  table.ref = "A1:C6";
  table.sheet_index = kDataSheet;
  for (const char* name : {"Region", "Product", "Amt"}) {
    TableColumn column;
    column.name = name;
    table.columns.push_back(std::move(column));
  }
  ASSERT_TRUE(static_cast<bool>(wb.add_table(std::move(table))));
  pivot::WorksheetSource& source = CacheOf(wb).mutable_worksheet_source();
  source.ref.clear();
  source.name = "Tbl_Src";
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  EXPECT_EQ(wb.find_pivot_cache(kCacheId)->records().size(), 5U);
}

// ===========================================================================
// Limit
// ===========================================================================

TEST(PivotRefresh, LimitNominalWholeColumnSourceIsRefused) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().ref = "A:XFD";
  ExpectRefused(wb, FormulonErrorCode::kSecurityResourceLimit);
}

TEST(PivotRefresh, LimitSourceWithinTheCeilingRefreshes) {
  Workbook wb = Book();
  CacheOf(wb).mutable_worksheet_source().ref = "A1:C6";
  EXPECT_TRUE(static_cast<bool>(Refresh(wb)));
  static_assert(5U * 3U < kMaxPivotCacheRecordCells, "the base source is under the ceiling");
}

// ===========================================================================
// Rollback: a report that cannot grow refuses the whole refresh
// ===========================================================================

TEST(PivotRefresh, RollbackSecondPivotBlockedLeavesBothUnchanged) {
  Workbook wb = Book(/*pivots=*/2);
  ASSERT_TRUE(PivotAt(wb, 1U).projected_extent().has_value());
  const pivot::PivotRect second = *PivotAt(wb, 1U).projected_extent();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, second.row1 + 1U, second.col0, "keep me")));
  SetText(wb, 3U, 0U, "West");  // a fourth region: each report grows one row
  SetText(wb, 1U, 0U, "Central");
  ExpectRefused(wb, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(wb.sheet(kPivotSheet).resolve_cell_value(second.row1 + 1U, second.col0).as_text(), "keep me");
}

TEST(PivotRefresh, RollbackGrowthNamesTheBlockingCell) {
  Workbook wb = Book();
  const pivot::PivotRect first = *PivotAt(wb, 0U).projected_extent();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(kPivotSheet, first.row1 + 1U, first.col0, "x")));
  SetText(wb, 4U, 0U, "West");  // the second North becomes a fourth region
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_FALSE(static_cast<bool>(refreshed));
  EXPECT_NE(refreshed.error().message.find("The pivot report would overwrite non-empty cell Pivot!A"),
            std::string::npos)
      << refreshed.error().message;
}

TEST(PivotRefresh, RollbackPlannedReportsOverlappingEachOtherAreRefused) {
  Workbook wb = Book(/*pivots=*/2);
  std::vector<Workbook::PivotGrowth> growth(2U);
  growth[0].sheet = kPivotSheet;
  growth[0].pivot = 0U;
  growth[0].extent = MergeRange{40U, 0U, 45U, 3U};
  growth[1].sheet = kPivotSheet;
  growth[1].pivot = 1U;
  growth[1].extent = MergeRange{43U, 2U, 50U, 5U};
  const Expected<void, Error> checked = wb.check_pivot_growth(growth);
  ASSERT_FALSE(static_cast<bool>(checked));
  EXPECT_EQ(checked.error().code, FormulonErrorCode::kPreconditionFailed);
  growth[1].extent = MergeRange{46U, 2U, 50U, 5U};
  EXPECT_TRUE(static_cast<bool>(wb.check_pivot_growth(growth)));
}

TEST(PivotRefresh, RollbackPlannedReportOverAnotherPivotIsRefused) {
  Workbook wb = Book(/*pivots=*/2);
  const pivot::PivotRect second = *PivotAt(wb, 1U).projected_extent();
  std::vector<Workbook::PivotGrowth> growth(1U);
  growth[0].sheet = kPivotSheet;
  growth[0].pivot = 0U;
  growth[0].extent = MergeRange{second.row0, 0U, second.row1, second.col0};
  EXPECT_FALSE(static_cast<bool>(wb.check_pivot_growth(growth)));
}

// ===========================================================================
// Hints: the <sharedItems> flags and bounds
// ===========================================================================

SharedItemsHints Hints(const std::vector<Value>& values, const std::vector<bool>& dates = {}) {
  return pivot::derive_shared_items_hints(values, dates, /*date1904=*/false);
}

TEST(PivotRefresh, HintsBoolsAndErrorsCountAsStrings) {
  const SharedItemsHints only_bools = Hints({Value::boolean(true), Value::boolean(false)});
  EXPECT_TRUE(only_bools.present);
  EXPECT_FALSE(only_bools.has_contains_string);
  EXPECT_FALSE(only_bools.has_contains_mixed_types);
  const SharedItemsHints mixed = Hints({Value::error(ErrorCode::Div0), Value::number(2)});
  EXPECT_TRUE(mixed.has_contains_mixed_types && mixed.contains_mixed_types);
  EXPECT_TRUE(mixed.has_contains_number && mixed.contains_number);
  EXPECT_FALSE(mixed.has_contains_string);
  EXPECT_EQ(mixed.min_value, "2");
}

TEST(PivotRefresh, HintsAllBlankField) {
  const SharedItemsHints h = Hints({Value::blank(), Value::blank()});
  EXPECT_TRUE(h.has_contains_non_date && !h.contains_non_date);
  EXPECT_TRUE(h.has_contains_string && !h.contains_string);
  EXPECT_TRUE(h.has_contains_blank && h.contains_blank);
  EXPECT_FALSE(h.has_contains_semi_mixed);
}

TEST(PivotRefresh, HintsLongTextFromTheTwoHundredFiftySixthUnit) {
  const std::string at_limit(255U, 'x');
  const std::string over(256U, 'x');
  EXPECT_TRUE(Hints({Value::text(at_limit)}).extra_attrs.empty());
  const SharedItemsHints h = Hints({Value::text(over)});
  ASSERT_EQ(h.extra_attrs.size(), 1U);
  EXPECT_EQ(h.extra_attrs[0].first, "longText");
}

TEST(PivotRefresh, HintsMaxDateIsTheNextMidnightUnlessTimed) {
  SharedItemsHints h = Hints({Value::number(kJan1), Value::number(kMar1)}, {true, true});
  EXPECT_EQ(h.min_date, "2026-01-01T00:00:00");
  EXPECT_EQ(h.max_date, "2026-03-02T00:00:00");
  h = Hints({Value::number(kJan1 + 4.25)}, {true});
  EXPECT_EQ(h.max_date, "2026-01-05T06:00:00");
}

TEST(PivotRefresh, HintsMixedDateSlotTextReinterpretsTheNumber) {
  EXPECT_EQ(pivot::excel_mixed_date_slot_text(5.0), "1899-12-31T00:21:04");
  EXPECT_EQ(pivot::excel_mixed_date_slot_text(0.1), "1900-01-06T12:06:03");
  EXPECT_EQ(pivot::excel_mixed_date_slot_text(-5.0), "1899-12-31T00:23:12");
  EXPECT_EQ(pivot::excel_mixed_date_slot_text(1.0 / 3.0), "1899-12-30T00:00:00");
  EXPECT_EQ(pivot::excel_mixed_date_slot_text(1e300), "1899-12-30T00:00:00");
}

TEST(PivotRefresh, HintsMixedDateSlotsFollowRecordOrder) {
  std::vector<Value> values = {Value::number(kJan1), Value::number(5)};
  std::vector<bool> dates = {true, false};
  SharedItemsHints h = Hints(values, dates);
  EXPECT_FALSE(h.has_contains_number);
  EXPECT_FALSE(h.has_min_date);
  pivot::apply_mixed_date_slots(values, dates, false, h);
  EXPECT_EQ(h.min_date, "2026-01-01T00:00:00");
  EXPECT_EQ(h.max_date, "1899-12-31T00:21:04");

  values = {Value::number(5), Value::number(kJan1)};
  dates = {false, true};
  h = Hints(values, dates);
  pivot::apply_mixed_date_slots(values, dates, false, h);
  EXPECT_EQ(h.min_date, "1899-12-31T00:21:04");
  EXPECT_EQ(h.max_date, "2026-01-02T00:00:00");
}

TEST(PivotRefresh, HintsEmptyValuesAreNotPresent) {
  EXPECT_FALSE(Hints({}).present);
}

// ===========================================================================
// Numeric groups: automatic bounds regenerate, fixed ones keep their items
// ===========================================================================

std::vector<std::string> GroupTexts(const PivotCacheField& field) {
  std::vector<std::string> out;
  for (const Value& v : field.field_group->group_items) {
    out.emplace_back(v.is_text() ? std::string(v.as_text()) : std::string("#"));
  }
  return out;
}

// Amt (inline) and its numeric group Amt2 over [0..400] by 100.
PivotCache AmountGroupCache(bool fixed, std::vector<const char*> labels) {
  PivotCache old;
  old.set_cache_id(kCacheId);
  old.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  old.mutable_fields()[0].field_group = pivot::FieldGroup{};
  old.mutable_fields()[0].field_group->par = 1U;
  PivotCacheField group;
  group.name = "Amt2";
  group.is_database_field = false;
  group.field_group = pivot::FieldGroup{};
  group.field_group->base = 0U;
  group.field_group->kind = pivot::GroupKind::Range;
  group.field_group->range.start_num = 0.0;
  group.field_group->range.end_num = 400.0;
  group.field_group->range.interval = 100.0;
  group.field_group->range.auto_start = !fixed;
  group.field_group->range.auto_end = !fixed;
  for (const char* label : labels) {
    group.field_group->group_items.push_back(Value::text(label));
  }
  old.mutable_fields().push_back(std::move(group));
  return old;
}

TEST(PivotRefresh, RulesNumericGroupAutoBoundsRegenerateItems) {
  const PivotCache old = AmountGroupCache(false, {"<0", "0-99", "100-199", "200-299", "300-399", "400-499", ">500"});
  const RefreshPlan plan = Build(
      old, Source({"Amt"}, {{Value::number(20)}, {Value::number(300)}, {Value::number(140)}}), Context({false, true}));
  const pivot::FieldGroup& group = *plan.cache.fields()[1].field_group;
  EXPECT_EQ(group.range.start_num, 20.0);
  EXPECT_EQ(group.range.end_num, 300.0);
  EXPECT_TRUE(group.range.auto_start);
  EXPECT_TRUE(group.range.auto_end);
  EXPECT_EQ(GroupTexts(plan.cache.fields()[1]),
            (std::vector<std::string>{"<20", "20-119", "120-219", "220-319", ">320"}));
  // No label survives, so every old item goes and the pivot field takes the new ones in order.
  EXPECT_EQ(plan.remap.items[1], (std::vector<std::optional<std::uint32_t>>(7U, std::nullopt)));
}

TEST(PivotRefresh, RulesNumericGroupKeepsItemsWhoseLabelsStay) {
  const PivotCache old = AmountGroupCache(false, {"<0", "0-99", "100-199", "200-299", "300-399", "400-499", ">500"});
  const RefreshPlan plan =
      Build(old, Source({"Amt"}, {{Value::number(0)}, {Value::number(550)}}), Context({false, true}));
  EXPECT_EQ(GroupTexts(plan.cache.fields()[1]),
            (std::vector<std::string>{"<0", "0-99", "100-199", "200-299", "300-399", "400-499", "500-599", ">600"}));
  EXPECT_EQ(plan.remap.items[1],
            (std::vector<std::optional<std::uint32_t>>{At(0U), At(1U), At(2U), At(3U), At(4U), At(5U), std::nullopt}));
}

TEST(PivotRefresh, RulesNumericGroupFixedBoundsKeepLabelsVerbatim) {
  // Excel keeps an authored label as it is.
  const PivotCache old = AmountGroupCache(true, {"<0", "0-99", "100-199", "200-299", "300-400", ">400"});
  RefreshPlan plan = Build(old, Source({"Amt"}, {{Value::number(-10)}, {Value::number(150)}, {Value::number(999)}}),
                           Context({false, true}));
  const pivot::FieldGroup& group = *plan.cache.fields()[1].field_group;
  EXPECT_EQ(group.range.start_num, 0.0);
  EXPECT_EQ(group.range.end_num, 400.0);
  EXPECT_EQ(GroupTexts(plan.cache.fields()[1]),
            (std::vector<std::string>{"<0", "0-99", "100-199", "200-299", "300-400", ">400"}));
  pivot::materialize_group_fields(plan.cache, /*date1904=*/false);
  std::vector<double> numbers;
  for (const PivotCacheRecord& record : plan.cache.records()) {
    numbers.push_back(record.cells[1].as_number());
  }
  EXPECT_EQ(numbers, (std::vector<double>{0, 2, 5}));
}

TEST(PivotRefresh, RulesNumericGroupFixedNonIntegerBoundIsNotImplemented) {
  PivotCache old = AmountGroupCache(true, {"<0", "0-99", ">100"});
  old.mutable_fields()[1].field_group->range.end_num = 99.5;
  const auto plan = pivot::build_refreshed_cache(old, Source({"Amt"}, {{Value::number(5)}}), Context({false, true}));
  ASSERT_FALSE(static_cast<bool>(plan));
  EXPECT_EQ(plan.error().code, FormulonErrorCode::kNotImplemented);
}

TEST(PivotRefresh, NumericGroupRefreshFollowsTheSourceOnEveryPivot) {
  Workbook wb = Book(2);
  {
    const eval::ScopedEvalProfile scope(wb.excel_profile());
    ASSERT_TRUE(
        static_cast<bool>(pivot::set_numeric_group(wb, kPivotSheet, 0U, kAmt, std::nullopt, std::nullopt, 100.0)));
  }
  EXPECT_EQ(GroupTexts(CacheOf(wb).fields()[3]),
            (std::vector<std::string>{"<100", "100-199", "200-299", "300-399", "400-499", "500-599", ">600"}));
  SetNumber(wb, 1U, 2U, 1000.0);  // 100 -> 1000: the data now spans 200..1000
  const Expected<void, Error> refreshed = Refresh(wb);
  ASSERT_TRUE(static_cast<bool>(refreshed)) << refreshed.error().message;
  const PivotCache& cache = *wb.find_pivot_cache(kCacheId);
  ASSERT_EQ(cache.fields().size(), 4U);
  EXPECT_EQ(cache.fields()[3].field_group->range.start_num, 200.0);
  EXPECT_EQ(cache.fields()[3].field_group->range.end_num, 1000.0);
  EXPECT_EQ(cache.fields()[3].field_group->group_items.size(), 11U);
  EXPECT_EQ(cache.records()[0].cells[3].as_number(), 9.0);  // 1000 in 1000-1099
  for (std::size_t p = 0; p < 2U; ++p) {
    const std::vector<PivotItem>& items = PivotAt(wb, p).fields()[3].items;
    ASSERT_EQ(items.size(), 11U);
    for (std::uint32_t x = 0; x < 11U; ++x) {
      EXPECT_EQ(items[x].cache_index, x);
    }
  }
}

// ===========================================================================
// Inline fold and sharing a field that joins an axis
// ===========================================================================

// The source column measured in Excel: pear, Apple, (blank), Pear, apple, Fig,
// (blank), Apple, Fig, PEAR, fig, Apple.
std::vector<Value> MeasuredProducts() {
  return {Value::text("pear"),  Value::text("Apple"), Value::blank(),     Value::text("Pear"),
          Value::text("apple"), Value::text("Fig"),   Value::blank(),     Value::text("Apple"),
          Value::text("Fig"),   Value::text("PEAR"),  Value::text("fig"), Value::text("Apple")};
}

// A cache whose one field, Product, holds `values` inline, one record each,
// with the hints a read or refreshed field carries; `dates` marks the cells
// spelled as dates.
PivotCache InlineCache(const std::vector<Value>& values, const std::vector<bool>& dates = {}) {
  PivotCache cache;
  cache.set_cache_id(kCacheId);
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields()[0].shared_items_hints.present = true;
  for (std::size_t r = 0; r < values.size(); ++r) {
    PivotCacheRecord record;
    record.cells = {values[r]};
    record.cell_is_index = {false};
    record.cell_is_date = {r < dates.size() && dates[r]};
    cache.mutable_records().push_back(std::move(record));
  }
  return cache;
}

std::vector<double> RecordIndices(const PivotCache& cache, std::size_t field) {
  std::vector<double> out;
  for (const PivotCacheRecord& record : cache.records()) {
    EXPECT_TRUE(record.cell_is_index[field]);
    out.push_back(record.cells[field].as_number());
  }
  return out;
}

TEST(PivotRefresh, RulesInlineTextFoldsToTheFirstSpelling) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  std::vector<std::vector<Value>> rows;
  for (const Value& v : MeasuredProducts()) {
    rows.push_back({v});
  }
  const RefreshPlan plan = Build(RegionCache({}), Source({"Product"}, rows), Context({false}));
  ASSERT_TRUE(plan.cache.fields()[0].shared_items.empty());
  std::vector<std::string> cells;
  for (const PivotCacheRecord& record : plan.cache.records()) {
    EXPECT_FALSE(record.cell_is_index[0]);
    cells.emplace_back(record.cells[0].is_text() ? std::string(record.cells[0].as_text()) : std::string("#"));
  }
  EXPECT_EQ(cells, (std::vector<std::string>{"pear", "Apple", "#", "pear", "Apple", "Fig", "#", "Apple", "Fig", "pear",
                                             "Fig", "Apple"}));
}

TEST(PivotRefresh, ShareKeepsFirstSpellingAndPutsTheBlankInRecordOrder) {
  PivotCache cache = InlineCache(MeasuredProducts());
  ASSERT_TRUE(pivot::share_field_items(cache, 0U));
  const PivotCacheField& field = cache.fields()[0];
  ASSERT_EQ(field.shared_items.size(), 4U);
  EXPECT_EQ(field.shared_items[0].as_text(), "pear");
  EXPECT_EQ(field.shared_items[1].as_text(), "Apple");
  EXPECT_TRUE(field.shared_items[2].is_blank());
  EXPECT_EQ(field.shared_items[3].as_text(), "Fig");
  EXPECT_EQ(RecordIndices(cache, 0U), (std::vector<double>{0, 1, 2, 0, 1, 3, 2, 1, 3, 0, 3, 1}));
  EXPECT_EQ(field.shared_item_is_date, (std::vector<bool>(4U, false)));
  EXPECT_EQ(field.shared_item_is_missing, (std::vector<bool>(4U, false)));
  EXPECT_EQ(field.shared_item_is_calculated, (std::vector<bool>(4U, false)));
}

TEST(PivotRefresh, ShareKeepsNumbersAndTextApart) {
  PivotCache cache = InlineCache({Value::number(30), Value::number(10), Value::number(30), Value::text("10"),
                                  Value::number(-0.0), Value::number(0)});
  ASSERT_TRUE(pivot::share_field_items(cache, 0U));
  const PivotCacheField& field = cache.fields()[0];
  ASSERT_EQ(field.shared_items.size(), 4U);
  EXPECT_EQ(field.shared_items[0].as_number(), 30.0);
  EXPECT_EQ(field.shared_items[1].as_number(), 10.0);
  EXPECT_EQ(field.shared_items[2].as_text(), "10");
  EXPECT_EQ(field.shared_items[3].as_number(), 0.0);
  EXPECT_EQ(RecordIndices(cache, 0U), (std::vector<double>{0, 1, 0, 2, 3, 3}));
}

TEST(PivotRefresh, ShareKeepsTheDateMarkPerItem) {
  PivotCache cache =
      InlineCache({Value::number(kJan1), Value::number(kJan1), Value::number(kMar1), Value::number(kJan1)},
                  {true, false, true, true});
  ASSERT_TRUE(pivot::share_field_items(cache, 0U));
  const PivotCacheField& field = cache.fields()[0];
  ASSERT_EQ(field.shared_items.size(), 3U);
  EXPECT_EQ(field.shared_items[0].as_number(), kJan1);
  EXPECT_EQ(field.shared_items[1].as_number(), kJan1);
  EXPECT_EQ(field.shared_items[2].as_number(), kMar1);
  EXPECT_EQ(field.shared_item_is_date, (std::vector<bool>{true, false, true}));
  EXPECT_EQ(RecordIndices(cache, 0U), (std::vector<double>{0, 1, 2, 0}));
  for (const PivotCacheRecord& record : cache.records()) {
    EXPECT_FALSE(record.cell_is_date[0]);
  }
}

TEST(PivotRefresh, ShareLeavesASharedFieldAsItIs) {
  PivotCache cache = RegionCache({Value::text("North"), Value::text("South")});
  for (const double x : {1.0, 0.0}) {
    PivotCacheRecord record;
    record.cells = {Value::number(x)};
    record.cell_is_index = {true};
    cache.mutable_records().push_back(std::move(record));
  }
  EXPECT_FALSE(pivot::share_field_items(cache, 0U));
  EXPECT_EQ(ItemTexts(cache.fields()[0]), (std::vector<std::string>{"North", "South"}));
  EXPECT_EQ(RecordIndices(cache, 0U), (std::vector<double>{1, 0}));
}

TEST(PivotRefresh, ShareRefusesFieldsItDoesNotOwn) {
  PivotCache empty = InlineCache({});
  EXPECT_FALSE(pivot::share_field_items(empty, 0U));
  // Filled in memory: the writer derives its shared items at save.
  PivotCache hand_filled = InlineCache({Value::text("A")});
  hand_filled.mutable_fields()[0].shared_items_hints.present = false;
  EXPECT_FALSE(pivot::share_field_items(hand_filled, 0U));
  EXPECT_FALSE(hand_filled.records()[0].cell_is_index[0]);
  PivotCache cache = InlineCache({Value::text("A")});
  cache.mutable_fields().push_back(PivotCacheField{"Calc", {}});
  cache.mutable_fields().back().formula = "Product*2";
  cache.mutable_fields().back().is_database_field = false;
  EXPECT_FALSE(pivot::share_field_items(cache, 1U));
  EXPECT_FALSE(pivot::share_field_items(cache, 2U));
  EXPECT_TRUE(cache.fields()[1].shared_items.empty());
}

TEST(PivotRefresh, SharePinsTheOtherFieldsAndKeepsTheCacheRefreshed) {
  // Region is shared with its encoding inferred (no per-cell flags), Product inline.
  PivotCache cache = RegionCache({Value::text("North"), Value::text("South")});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields()[1].shared_items_hints.present = true;
  const double regions_in[] = {1.0, 0.0, 1.0};
  const char* const products[] = {"B", "A", "b"};
  for (std::size_t r = 0; r < 3U; ++r) {
    PivotCacheRecord record;
    record.cells = {Value::number(regions_in[r]), Value::text(products[r])};
    cache.mutable_records().push_back(std::move(record));
  }
  cache.mark_records_from_source();
  ASSERT_TRUE(pivot::share_field_items(cache, 1U));
  EXPECT_TRUE(cache.records_from_source());
  EXPECT_EQ(ItemTexts(cache.fields()[1]), (std::vector<std::string>{"B", "A"}));
  EXPECT_EQ(RecordIndices(cache, 1U), (std::vector<double>{0, 1, 0}));
  std::vector<std::string> regions;
  for (const PivotCacheRecord& record : cache.records()) {
    regions.emplace_back(pivot::cell_value(cache, record, 0U).as_text());
  }
  EXPECT_EQ(regions, (std::vector<std::string>{"South", "North", "South"}));
}

TEST(PivotRefresh, ShareSurvivesALaterWriteToTheSameCell) {
  PivotCache cache = InlineCache({Value::blank()});
  ASSERT_TRUE(cache.set_record_text(0U, 0U, "First"));
  ASSERT_TRUE(pivot::share_field_items(cache, 0U));
  ASSERT_TRUE(cache.set_record_text(0U, 0U, "Second"));
  EXPECT_EQ(ItemTexts(cache.fields()[0]), (std::vector<std::string>{"First"}));
}

}  // namespace
}  // namespace formulon
