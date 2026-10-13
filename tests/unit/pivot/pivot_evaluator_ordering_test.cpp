// Ordering and comparator behavior for `formulon::pivot::evaluate`.

#include <array>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/groupby_pivotby/common.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "pivot_evaluator_fixtures.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using test::build_basic_cache;
using test::build_sum_amount_table;
using test::owned_text;

TEST(PivotComparatorParity, BoolVsTextMatchesGroupByOrder) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const Value text_val = Value::text("zebra");
  const Value bool_val = Value::boolean(false);

  // Pivot: Text sorts before Bool.
  EXPECT_TRUE(value_less(text_val, bool_val));
  EXPECT_FALSE(value_less(bool_val, text_val));

  // GROUPBY / SORT: same ordering (negative => first argument sorts first).
  EXPECT_LT(eval::cmp_value_asc(text_val, bool_val), 0);
  EXPECT_GT(eval::cmp_value_asc(bool_val, text_val), 0);
}

TEST(PivotComparatorParity, NumberBeforeTextBeforeBool) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const Value num = Value::number(1.0);
  const Value text_val = Value::text("a");
  const Value bool_val = Value::boolean(true);

  EXPECT_TRUE(value_less(num, text_val));
  EXPECT_TRUE(value_less(text_val, bool_val));

  EXPECT_LT(eval::cmp_value_asc(num, text_val), 0);
  EXPECT_LT(eval::cmp_value_asc(text_val, bool_val), 0);
}

TEST(PivotComparatorParity, OrdersNonFiniteNumbersAfterFiniteValues) {
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const Value negative_infinity = Value::number(-std::numeric_limits<double>::infinity());
  const Value finite = Value::number(42.0);
  const Value positive_infinity = Value::number(std::numeric_limits<double>::infinity());
  const Value nan = Value::number(std::numeric_limits<double>::quiet_NaN());

  EXPECT_TRUE(value_less(negative_infinity, finite));
  EXPECT_TRUE(value_less(finite, positive_infinity));
  EXPECT_TRUE(value_less(positive_infinity, nan));
  EXPECT_FALSE(value_less(nan, positive_infinity));
  EXPECT_FALSE(value_less(nan, nan));

  std::map<Value, int, ValueLess> values;
  values.emplace(negative_infinity, 1);
  values.emplace(finite, 2);
  values.emplace(positive_infinity, 3);
  values.emplace(nan, 4);
  EXPECT_EQ(values.size(), 4U);
  EXPECT_EQ(values.at(negative_infinity), 1);
  EXPECT_EQ(values.at(finite), 2);
  EXPECT_EQ(values.at(positive_infinity), 3);
  EXPECT_EQ(values.at(nan), 4);
}

TEST(PivotEvaluator, FieldDescendingSortReversesHierarchyAndValues) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].sort.ascending = false;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.cols.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  EXPECT_EQ(r.cols[0].label, "Widget");
  EXPECT_EQ(r.cols[1].label, "Gadget");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);  // South/Widget
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 300.0);  // South/Gadget
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 125.0);  // North/Widget
  EXPECT_DOUBLE_EQ(r.values[1][1][0].as_number(), 50.0);   // North/Gadget
}

TEST(PivotEvaluator, FieldSortByValueFieldReordersAxisAndValues) {
  PivotCache cache = build_basic_cache();
  PivotTable table = build_sum_amount_table(/*row=*/{0}, /*col=*/{1});
  table.mutable_fields()[0].sort.by_field = "Amount";
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].sort.by_field = "Sum of Amount";
  table.mutable_fields()[1].sort.ascending = true;

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  ASSERT_EQ(r.cols.size(), 2U);
  // Row totals: South=500, North=175, so descending by Amount puts South first.
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  // Column totals: Widget=325, Gadget=350, so ascending by Sum of Amount puts Widget first.
  EXPECT_EQ(r.cols[0].label, "Widget");
  EXPECT_EQ(r.cols[1].label, "Gadget");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);  // South/Widget
  EXPECT_DOUBLE_EQ(r.values[0][1][0].as_number(), 300.0);  // South/Gadget
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 125.0);  // North/Widget
  EXPECT_DOUBLE_EQ(r.values[1][1][0].as_number(), 50.0);   // North/Gadget
}

TEST(PivotEvaluator, ManualSortOrdersByItemDocumentPosition) {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(owned_text(cache, "South"));
  region.shared_items.push_back(owned_text(cache, "North"));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](std::uint32_t region_index, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), Value::number(amount)};
    rec.cell_is_index = {true, false};
    cache.mutable_records().push_back(std::move(rec));
  };
  add(0U, 200.0);  // South
  add(1U, 100.0);  // North

  PivotTable table;
  table.set_pivot_cache_id(1);
  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  region_f.sort.manual = true;
  region_f.items.push_back(PivotItem{"South", true, /*has_cache_index=*/true, /*cache_index=*/0U});
  region_f.items.push_back(PivotItem{"North", true, /*has_cache_index=*/true, /*cache_index=*/1U});
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(amount_f));

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.mutable_row_field_order() = {0};

  auto r_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(r_or)) << r_or.error().message;
  const PivotResult& r = r_or.value();
  ASSERT_EQ(r.rows.size(), 2U);
  EXPECT_EQ(r.rows[0].label, "South");
  EXPECT_EQ(r.rows[1].label, "North");
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 200.0);
  EXPECT_DOUBLE_EQ(r.values[1][0][0].as_number(), 100.0);
}

TEST(PivotEvaluator, SortedSubtotalWalkPreservesTypedDuplicateDisplayLabels) {
  PivotCache cache;
  cache.set_cache_id(17);
  cache.mutable_fields().push_back(PivotCacheField{"Group", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Kind", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Item", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* group, Value kind, double amount) {
    PivotCacheRecord record;
    record.cells.push_back(owned_text(cache, group));
    record.cells.push_back(std::move(kind));
    record.cells.push_back(owned_text(cache, "Item"));
    record.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(record));
  };
  // Number(1) and Text("1") render to the same label, but remain distinct
  // hierarchy owners. The same is repeated under two sorted groups.
  add("S", Value::number(3.0), 10.0);
  add("S", owned_text(cache, "3"), 20.0);
  add("D", Value::number(1.0), 30.0);
  add("D", owned_text(cache, "1"), 40.0);

  PivotTable table;
  table.set_pivot_cache_id(17);
  for (const char* name : {"Group", "Kind", "Item"}) {
    PivotField field;
    field.source_name = name;
    field.axis = PivotAxis::Row;
    table.mutable_fields().push_back(std::move(field));
  }
  PivotField amount_field;
  amount_field.source_name = "Amount";
  amount_field.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(amount_field));
  table.mutable_row_field_order() = {0, 1, 2};
  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.mutable_fields()[0].sort.ascending = false;
  table.mutable_fields()[1].subtotal_fns = {SubtotalFn::Sum, SubtotalFn::Average};
  table.mutable_fields()[0].subtotal_top = true;
  table.set_grand_totals(/*rows=*/true, /*cols=*/true);
  table.set_anchor(0, 0, 1, 1);

  auto result_or = evaluate(table, cache);
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const PivotResult& result = result_or.value();
  ASSERT_EQ(result.rows.size(), 2U);
  ASSERT_EQ(result.rows[0].label, "S");
  ASSERT_EQ(result.rows[1].label, "D");
  ASSERT_EQ(result.row_subtotals.size(), 10U);

  PivotLayoutOptions options;
  options.row_labels_label = "Row Labels";
  auto cells_or = layout(table, result, options);
  ASSERT_TRUE(static_cast<bool>(cells_or)) << cells_or.error().message << " " << cells_or.error().context;
  const PivotCells& cells = cells_or.value();

  std::vector<std::pair<PivotCellKind, double>> cells_in_render_order;
  for (const PivotCell& cell : cells.cells) {
    if (cell.col != cells.left + 1U || !cell.value.is_number()) {
      continue;
    }
    if (cell.kind == PivotCellKind::Data || cell.kind == PivotCellKind::RowSubtotal ||
        cell.kind == PivotCellKind::GrandTotal) {
      cells_in_render_order.emplace_back(cell.kind, cell.value.as_number());
    }
  }
  EXPECT_EQ(cells_in_render_order, (std::vector<std::pair<PivotCellKind, double>>{
                                       {PivotCellKind::RowSubtotal, 30.0},  // S subtotal.
                                       {PivotCellKind::Data, 10.0},         // S / Number(3) detail.
                                       {PivotCellKind::RowSubtotal, 10.0},  // S / Number(3) SUM.
                                       {PivotCellKind::RowSubtotal, 10.0},  // S / Number(3) AVERAGE.
                                       {PivotCellKind::Data, 20.0},         // S / Text("3") detail.
                                       {PivotCellKind::RowSubtotal, 20.0},  // S / Text("3") SUM.
                                       {PivotCellKind::RowSubtotal, 20.0},  // S / Text("3") AVERAGE.
                                       {PivotCellKind::RowSubtotal, 70.0},  // D subtotal.
                                       {PivotCellKind::Data, 30.0},         // D / Number(1) detail.
                                       {PivotCellKind::RowSubtotal, 30.0},  // D / Number(1) SUM.
                                       {PivotCellKind::RowSubtotal, 30.0},  // D / Number(1) AVERAGE.
                                       {PivotCellKind::Data, 40.0},         // D / Text("1") detail.
                                       {PivotCellKind::RowSubtotal, 40.0},  // D / Text("1") SUM.
                                       {PivotCellKind::RowSubtotal, 40.0},  // D / Text("1") AVERAGE.
                                       {PivotCellKind::GrandTotal, 100.0},
                                   }));
  EXPECT_EQ(cells.cols, 2U);
  EXPECT_GT(cells.rows, result.row_subtotals.size());
}

TEST(PivotComparatorParity, MergesWhatGroupByMerges) {
  // U+FF71 and U+30A2 are one GROUPBY group on Mac ja_JP
  // (tools/collation/data/fixtures/eq_ja_JP.tsv).
  const eval::ScopedEvalProfile scope(mac_365_ja_jp_profile());
  const Value full = Value::text("\xE3\x82\xA2");
  const Value half = Value::text("\xEF\xBD\xB1");
  EXPECT_FALSE(value_less(full, half));
  EXPECT_FALSE(value_less(half, full));
  EXPECT_EQ(eval::cmp_value_asc(full, half), 0);
}

// One row field over `labels`, one record per label with amount 10 * (i + 1).
PivotCache build_label_cache(const std::vector<std::string>& labels) {
  PivotCache cache;
  cache.set_cache_id(5);
  cache.mutable_fields().push_back(PivotCacheField{"Item", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  for (std::size_t i = 0; i < labels.size(); ++i) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, labels[i]));
    rec.cells.push_back(Value::number(10.0 * static_cast<double>(i + 1U)));
    cache.mutable_records().push_back(std::move(rec));
  }
  return cache;
}

PivotTable build_label_table() {
  PivotTable table;
  table.set_pivot_cache_id(5);
  PivotField item_f;
  item_f.source_name = "Item";
  item_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;
  table.mutable_fields().push_back(std::move(item_f));
  table.mutable_fields().push_back(std::move(amount_f));
  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 1;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  table.mutable_row_field_order() = {0};
  return table;
}

PivotResult evaluate_labels(const std::vector<std::string>& labels, ExcelProfile profile,
                            const PivotTable& table = build_label_table()) {
  const PivotCache cache = build_label_cache(labels);
  auto r_or = evaluate(table, cache, pivot_layout_options_for(profile));
  EXPECT_TRUE(static_cast<bool>(r_or));
  return r_or ? std::move(r_or.value()) : PivotResult{};
}

ExcelProfile mac(ExcelLocale locale) {
  return ExcelProfile{ExcelHost::kMac365, locale};
}

// Item counts per locale, from the `groupby` column of
// tools/collation/data/fixtures/eq_{ja_JP,en_US,zh_CN}.tsv; pivot items merge
// at the same strength as GROUPBY keys.
TEST(PivotEvaluatorCollation, MacItemsMergeAsMeasured) {
  struct ItemCase {
    const char* a;
    const char* b;
    std::size_t ja_jp;
    std::size_t en_us;
    std::size_t zh_cn;
  };
  const std::array<ItemCase, 6> cases = {{
      {"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC", 1U, 1U, 1U},              // U+FF76 U+FF9E | U+30AC
      {"\xE3\x82\xA2", "\xE3\x81\x82", 1U, 1U, 2U},                          // U+30A2 | U+3042
      {"\xE3\x81\x8B\xE3\x83\xBC", "\xE3\x81\x8B\xE3\x81\x82", 1U, 2U, 2U},  // U+304B U+30FC | U+304B U+3042
      {"\xE3\x82\x9D", "\xE3\x83\xBD", 1U, 2U, 1U},                          // U+309D | U+30FD
      {"\xC3\xA6", "ae", 2U, 2U, 1U},                                        // U+00E6 | U+0061 U+0065
      {"\xC3\xA4", "a", 2U, 2U, 2U},                                         // U+00E4 | U+0061
  }};
  for (const ItemCase& c : cases) {
    const std::vector<std::string> labels{c.a, c.b};
    EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kJaJP)).rows.size(), c.ja_jp) << c.a << " | " << c.b;
    EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kEnUS)).rows.size(), c.en_us) << c.a << " | " << c.b;
    EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kZhCN)).rows.size(), c.zh_cn) << c.a << " | " << c.b;
  }
}

TEST(PivotEvaluatorCollation, MergedItemsAggregateTogether) {
  const PivotResult r = evaluate_labels({"\xEF\xBD\xB1", "\xE3\x82\xA2"}, mac(ExcelLocale::kJaJP));
  ASSERT_EQ(r.rows.size(), 1U);
  EXPECT_DOUBLE_EQ(r.values[0][0][0].as_number(), 30.0);
}

// Distinct items order as Mac Excel's SORT orders the same pair
// (colfn_sort_choon / colfn_sort_thai_prefix goldens).
TEST(PivotEvaluatorCollation, MacItemOrderFollowsDynamicCollation) {
  const std::string ka_choon = "\xE3\x81\x8B\xE3\x83\xBC";
  const std::string ka_a = "\xE3\x81\x8B\xE3\x81\x82";
  const PivotResult en = evaluate_labels({ka_a, ka_choon}, mac(ExcelLocale::kEnUS));
  ASSERT_EQ(en.rows.size(), 2U);
  EXPECT_EQ(en.rows[0].label, ka_choon);
  const PivotResult zh = evaluate_labels({ka_choon, ka_a}, mac(ExcelLocale::kZhCN));
  ASSERT_EQ(zh.rows.size(), 2U);
  EXPECT_EQ(zh.rows[0].label, ka_a);

  const std::string thai_e_ko = "\xE0\xB9\x80\xE0\xB8\x81";
  const std::string thai_ko_aa = "\xE0\xB8\x81\xE0\xB8\xB2";
  const PivotResult ja = evaluate_labels({thai_e_ko, thai_ko_aa}, mac(ExcelLocale::kJaJP));
  ASSERT_EQ(ja.rows.size(), 2U);
  EXPECT_EQ(ja.rows[0].label, thai_ko_aa);
  const PivotResult zh_thai = evaluate_labels({thai_ko_aa, thai_e_ko}, mac(ExcelLocale::kZhCN));
  ASSERT_EQ(zh_thai.rows.size(), 2U);
  EXPECT_EQ(zh_thai.rows[0].label, thai_e_ko);
}

// Windows keeps the byte model (win-365-ja_JP unique_case_insensitive_text,
// xlookup_exact_voiced_compose, xlookup_exact_hira_eq_kata).
TEST(PivotEvaluatorCollation, WindowsJaJpKeepsTheByteModel) {
  const ExcelProfile win = win_365_ja_jp_profile();
  EXPECT_EQ(evaluate_labels({"A", "a"}, win).rows.size(), 1U);
  EXPECT_EQ(evaluate_labels({"\xEF\xBD\xB6\xEF\xBE\x9E", "\xE3\x82\xAC"}, win).rows.size(), 1U);
  EXPECT_EQ(evaluate_labels({"\xE3\x81\x82", "\xE3\x82\xA2"}, win).rows.size(), 2U);
}

PivotTable caption_filtered_table(CaptionPredicate predicate, std::string value) {
  PivotTable table = build_label_table();
  AuthoredCaptionFilter f;
  f.field_index = 0;
  f.predicate = predicate;
  f.value = std::move(value);
  table.mutable_authored_caption_filters().push_back(std::move(f));
  return table;
}

// A caption filter compares at the strength that merged the items.
TEST(PivotEvaluatorCollation, CaptionEqualUsesItemIdentity) {
  const std::vector<std::string> labels{"\xE3\x82\xA2", "\xE3\x82\xA4"};                          // ア, イ
  const PivotTable equal_hira = caption_filtered_table(CaptionPredicate::Equal, "\xE3\x81\x82");  // あ
  const PivotResult mac_r = evaluate_labels(labels, mac(ExcelLocale::kJaJP), equal_hira);
  ASSERT_EQ(mac_r.rows.size(), 1U);
  EXPECT_EQ(mac_r.rows[0].label, "\xE3\x82\xA2");
  EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kZhCN), equal_hira).rows.size(), 0U);
  EXPECT_EQ(evaluate_labels(labels, win_365_ja_jp_profile(), equal_hira).rows.size(), 0U);

  const PivotTable not_equal_hira = caption_filtered_table(CaptionPredicate::NotEqual, "\xE3\x81\x82");
  const PivotResult mac_ne = evaluate_labels(labels, mac(ExcelLocale::kJaJP), not_equal_hira);
  ASSERT_EQ(mac_ne.rows.size(), 1U);
  EXPECT_EQ(mac_ne.rows[0].label, "\xE3\x82\xA4");
}

TEST(PivotEvaluatorCollation, CaptionOrderingUsesItemOrder) {
  // かー sorts after かあ in zh_CN and before it in en_US.
  const std::vector<std::string> labels{"\xE3\x81\x8B\xE3\x83\xBC"};
  const PivotTable greater = caption_filtered_table(CaptionPredicate::GreaterThan, "\xE3\x81\x8B\xE3\x81\x82");
  EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kZhCN), greater).rows.size(), 1U);
  EXPECT_EQ(evaluate_labels(labels, mac(ExcelLocale::kEnUS), greater).rows.size(), 0U);
}

}  // namespace
}  // namespace formulon::pivot
