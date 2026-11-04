//
// Calculated fields and items: the model operations on a cache and its
// tables, and the cells the evaluator gives them. The data is the
// Region / Product / Amt / Qty set the Excel fixtures under
// tests/fixtures/excel/pivot_calc/ were built from.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/eval_profile_scope.h"
#include "eval/pivot_formula.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "pivot/calculated.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"

namespace formulon::pivot {
namespace {

constexpr std::uint32_t kRegion = 0;
constexpr std::uint32_t kProduct = 1;
constexpr std::uint32_t kAmt = 2;
constexpr std::uint32_t kQty = 3;

Value Owned(PivotCache& cache, const char* text) {
  cache.mutable_text_storage().emplace_back(text);
  return Value::text(cache.text_storage().back());
}

/// Region (North, South, East) / Product (A, B) as shared items, Amt and Qty
/// inline:
///   North A 120 4, North B 90 3, South A 300 5, South B 50 2, East A 70 6.
PivotCache BuildCache() {
  PivotCache cache;
  cache.set_cache_id(7);
  PivotCacheField region("Region", {});
  region.shared_items = {Owned(cache, "North"), Owned(cache, "South"), Owned(cache, "East")};
  PivotCacheField product("Product", {});
  product.shared_items = {Owned(cache, "A"), Owned(cache, "B")};
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(std::move(product));
  cache.mutable_fields().push_back(PivotCacheField{"Amt", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Qty", {}});
  const auto add = [&](double region_index, double product_index, double amt, double qty) {
    PivotCacheRecord record;
    record.cells = {Value::number(region_index), Value::number(product_index), Value::number(amt), Value::number(qty)};
    cache.mutable_records().push_back(std::move(record));
  };
  add(0, 0, 120, 4);
  add(0, 1, 90, 3);
  add(1, 0, 300, 5);
  add(1, 1, 50, 2);
  add(2, 0, 70, 6);
  return cache;
}

PivotDataField DataField(const char* name, std::uint32_t field, Aggregation aggregation = Aggregation::Sum) {
  PivotDataField data_field;
  data_field.name = name;
  data_field.field_index = field;
  data_field.aggregation = aggregation;
  return data_field;
}

/// Region on rows in the order East, North, South (manual, as Excel lists
/// them), Product unplaced or on columns, Sum of Amt and Sum of Qty.
PivotTable BuildTable(bool product_on_columns = false) {
  PivotTable table;
  table.set_pivot_cache_id(7);
  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  region.sort.manual = true;
  for (const std::uint32_t x : {2U, 0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    region.items.push_back(std::move(item));
  }
  PivotField product;
  product.source_name = "Product";
  product.axis = product_on_columns ? PivotAxis::Col : PivotAxis::None;
  product.sort.manual = true;
  for (const std::uint32_t x : {0U, 1U}) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = x;
    product.items.push_back(std::move(item));
  }
  PivotField amt;
  amt.source_name = "Amt";
  amt.axis = PivotAxis::Value;
  PivotField qty;
  qty.source_name = "Qty";
  qty.axis = PivotAxis::Value;
  table.mutable_fields() = {std::move(region), std::move(product), std::move(amt), std::move(qty)};
  table.mutable_row_field_order() = {kRegion};
  if (product_on_columns) {
    table.mutable_col_field_order() = {kProduct};
  }
  table.mutable_data_fields() = {DataField("Sum of Amt", kAmt), DataField("Sum of Qty", kQty)};
  return table;
}

PivotResult Evaluate(const PivotTable& table, const PivotCache& cache, bool with_host = true) {
  eval::PivotFormulaEvaluator host;
  PivotFilterEnv env;
  env.formulas = with_host ? &host : nullptr;
  auto result_or = evaluate(table, cache, PivotLayoutOptions{}, env);
  EXPECT_TRUE(static_cast<bool>(result_or)) << (result_or ? "" : result_or.error().message);
  return result_or ? std::move(result_or.value()) : PivotResult{};
}

std::vector<std::string> RowLabels(const PivotResult& result) {
  std::vector<std::string> labels;
  for (const AxisHierarchyNode& node : result.rows) {
    labels.push_back(node.label);
  }
  return labels;
}

std::vector<std::string> ColLabels(const PivotResult& result) {
  std::vector<std::string> labels;
  for (const AxisHierarchyNode& node : result.cols) {
    labels.push_back(node.label);
  }
  return labels;
}

/// Slot `slot` of every row leaf in column 0, then the grand total.
std::vector<Value> Column(const PivotResult& result, std::size_t slot) {
  std::vector<Value> out;
  for (const auto& row : result.values) {
    out.push_back(row[0][slot]);
  }
  out.push_back(result.grand_totals.at(slot));
  return out;
}

void ExpectNumbers(const std::vector<Value>& actual, const std::vector<double>& expected) {
  ASSERT_EQ(actual.size(), expected.size());
  for (std::size_t i = 0; i < expected.size(); ++i) {
    ASSERT_TRUE(actual[i].is_number()) << "position " << i;
    EXPECT_DOUBLE_EQ(actual[i].as_number(), expected[i]) << "position " << i;
  }
}

class PivotCalculated : public ::testing::Test {
 protected:
  const eval::ScopedEvalProfile scope_{mac_365_ja_jp_profile()};
  eval::PivotFormulaEvaluator host_;
  PivotCache cache_ = BuildCache();
};

TEST_F(PivotCalculated, AddedFieldAppliesItsFormulaToTheSumsOfEveryCellAndTotal) {
  PivotTable table = BuildTable();
  auto added = add_calculated_field(cache_, {&table}, "Price", "=Amt/Qty", host_);
  ASSERT_TRUE(static_cast<bool>(added)) << added.error().message;
  EXPECT_EQ(added.value(), 4U);
  const PivotCacheField& field = cache_.fields()[4];
  EXPECT_EQ(field.formula.value_or(""), "Amt/Qty");
  EXPECT_FALSE(field.is_database_field);
  EXPECT_TRUE(field.shared_items.empty());

  table.mutable_data_fields().push_back(DataField("合計 / Price", 4));
  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South"}));
  ExpectNumbers(Column(result, 2), {70.0 / 6.0, 30, 50, 31.5});
  ExpectNumbers(Column(result, 0), {70, 210, 350, 630});
  EXPECT_EQ(result.values[0][0].size(), 3U) << "the hidden Sum slots are dropped";
}

TEST_F(PivotCalculated, FieldValuesIgnoreTheDataFieldFunction) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  table.mutable_data_fields().push_back(DataField("個数 / Price", 4, Aggregation::Count));
  ExpectNumbers(Column(Evaluate(table, cache_), 2), {70.0 / 6.0, 30, 50, 31.5});
}

TEST_F(PivotCalculated, FieldMayNameAnotherCalculatedField) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price2", "Price*2", host_)));
  table.mutable_data_fields().push_back(DataField("合計 / Price2", 5));
  ExpectNumbers(Column(Evaluate(table, cache_), 2), {2.0 * (70.0 / 6.0), 60, 100, 63});
}

TEST_F(PivotCalculated, FieldErrorReachesEveryCell) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Bad", "Amt/0", host_)));
  table.mutable_data_fields().push_back(DataField("合計 / Bad", 4));
  for (const Value& value : Column(Evaluate(table, cache_), 2)) {
    ASSERT_TRUE(value.is_error());
    EXPECT_EQ(value.as_error(), ErrorCode::Div0);
  }
}

TEST_F(PivotCalculated, WithoutAHostCalculatedCellsAreBlank) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  table.mutable_data_fields().push_back(DataField("合計 / Price", 4));
  const PivotResult result = Evaluate(table, cache_, /*with_host=*/false);
  for (const Value& value : Column(result, 2)) {
    EXPECT_TRUE(value.is_blank());
  }
  ExpectNumbers(Column(result, 0), {70, 210, 350, 630});
}

TEST_F(PivotCalculated, AddedFieldIsAnUnplacedFieldInEveryTable) {
  PivotTable first = BuildTable();
  PivotTable second = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&first, &second}, "Price", "Amt/Qty", host_)));
  for (const PivotTable* table : {&first, &second}) {
    ASSERT_EQ(table->fields().size(), 5U);
    const PivotField& field = table->fields()[4];
    EXPECT_EQ(field.source_name, "Price");
    EXPECT_EQ(field.axis, PivotAxis::None);
    EXPECT_FALSE(field.show_all);
    EXPECT_FALSE(field.default_subtotal);
    EXPECT_EQ(field.passthrough_attrs, (std::vector<std::pair<std::string, std::string>>{
                                           {"dragToRow", "0"}, {"dragToCol", "0"}, {"dragToPage", "0"}}));
  }
}

TEST_F(PivotCalculated, SyncExtendsTablesToTheCacheFieldCount) {
  PivotTable table = BuildTable();
  cache_.mutable_fields().push_back(PivotCacheField{"Extra", {}});
  sync_table_fields(cache_, {&table});
  ASSERT_EQ(table.fields().size(), 5U);
  EXPECT_EQ(table.fields()[4].source_name, "Extra");
  EXPECT_EQ(table.fields()[4].axis, PivotAxis::None);
  EXPECT_TRUE(table.fields()[4].default_subtotal);
  EXPECT_TRUE(table.fields()[4].passthrough_attrs.empty());
}

TEST_F(PivotCalculated, FieldNamesAndFormulasAreChecked) {
  PivotTable table = BuildTable();
  for (const char* name : {"", "All", "blank", "a[b", "a]", "region"}) {
    auto added = add_calculated_field(cache_, {&table}, name, "Amt", host_);
    ASSERT_FALSE(static_cast<bool>(added)) << name;
    EXPECT_EQ(added.error().code, FormulonErrorCode::kInvalidArgument) << name;
  }
  auto bad_formula = add_calculated_field(cache_, {&table}, "Price", "Nope*2", host_);
  ASSERT_FALSE(static_cast<bool>(bad_formula));
  EXPECT_EQ(bad_formula.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(cache_.fields().size(), 4U);
  EXPECT_EQ(table.fields().size(), 4U);
}

TEST_F(PivotCalculated, SetFormulaRejectsACycle) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price2", "Price*2", host_)));
  auto cycle = set_calculated_field_formula(cache_, 4, "Price2+1", host_);
  ASSERT_FALSE(static_cast<bool>(cycle));
  EXPECT_EQ(cycle.error().code, FormulonErrorCode::kInvalidArgument);
  EXPECT_EQ(cache_.fields()[4].formula.value_or(""), "Amt/Qty");

  ASSERT_TRUE(static_cast<bool>(set_calculated_field_formula(cache_, 4, "=Amt*2", host_)));
  EXPECT_EQ(cache_.fields()[4].formula.value_or(""), "Amt*2");
  EXPECT_FALSE(static_cast<bool>(set_calculated_field_formula(cache_, kAmt, "Qty", host_)));
}

TEST_F(PivotCalculated, RemovingAFieldRenumbersTheFieldsAfterIt) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Double", "Amt*2", host_)));
  table.mutable_data_fields().push_back(DataField("合計 / Price", 4));
  table.mutable_data_fields().push_back(DataField("合計 / Double", 5));

  ASSERT_TRUE(static_cast<bool>(remove_calculated_field(cache_, {&table}, 4, host_)));
  ASSERT_EQ(cache_.fields().size(), 5U);
  EXPECT_EQ(cache_.fields()[4].name, "Double");
  ASSERT_EQ(table.fields().size(), 5U);
  EXPECT_EQ(table.fields()[4].source_name, "Double");
  ASSERT_EQ(table.data_fields().size(), 3U);
  EXPECT_EQ(table.data_fields()[2].name, "合計 / Double");
  EXPECT_EQ(table.data_fields()[2].field_index, 4U);
  ExpectNumbers(Column(Evaluate(table, cache_), 2), {140, 420, 700, 1260});
}

TEST_F(PivotCalculated, AFieldAnotherFormulaNamesCannotBeRemoved) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price2", "Price*2", host_)));
  auto removed = remove_calculated_field(cache_, {&table}, 4, host_);
  ASSERT_FALSE(static_cast<bool>(removed));
  EXPECT_EQ(removed.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_EQ(cache_.fields().size(), 6U);
  EXPECT_FALSE(static_cast<bool>(remove_calculated_field(cache_, {&table}, kAmt, host_)));
}

TEST_F(PivotCalculated, ItemSumsTheNamedItemsAndTheTotalsIncludeIt) {
  PivotTable table = BuildTable();
  auto added = add_calculated_item(cache_, {&table}, kRegion, "X", "Region[North]+Region[South]", host_);
  ASSERT_TRUE(static_cast<bool>(added)) << added.error().message;
  EXPECT_EQ(added.value(), 0U);

  const PivotCacheField& region = cache_.fields()[kRegion];
  ASSERT_EQ(region.shared_items.size(), 4U);
  EXPECT_EQ(region.shared_items[3].as_text(), "X");
  EXPECT_EQ(region.shared_item_is_calculated, (std::vector<bool>{false, false, false, true}));
  const PivotCalculatedItem& entry = cache_.calculated_items()[0];
  EXPECT_EQ(entry.cache_field_index, kRegion);
  EXPECT_EQ(entry.formula, "Region[North]+Region[South]");
  EXPECT_EQ(entry.area.cache_index, std::optional<bool>(true));
  EXPECT_EQ(entry.area.outline, std::optional<bool>(false));
  EXPECT_EQ(entry.area.field_position, std::optional<std::uint32_t>(0U));
  ASSERT_EQ(entry.area.references.size(), 1U);
  EXPECT_EQ(entry.area.references[0].field, std::optional<std::uint32_t>(kRegion));
  EXPECT_EQ(entry.area.references[0].items, (std::vector<std::uint32_t>{3U}));
  const PivotItem& item = table.fields()[kRegion].items.back();
  EXPECT_TRUE(item.calculated);
  EXPECT_EQ(item.cache_index, 3U);

  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South", "X"}));
  ExpectNumbers(Column(result, 0), {70, 210, 350, 560, 1190});
  ExpectNumbers(Column(result, 1), {6, 7, 7, 14, 34});
}

TEST_F(PivotCalculated, ItemReadsTheDataFieldFunction) {
  PivotTable table = BuildTable();
  table.mutable_data_fields()[0] = DataField("個数 / Amt", kAmt, Aggregation::Count);
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  ExpectNumbers(Column(Evaluate(table, cache_), 0), {1, 2, 2, 4, 9});
}

TEST_F(PivotCalculated, ItemTotalsFoldCellsWithTheDataFieldFunction) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  table.mutable_data_fields()[0] = DataField("最大 / Amt", kAmt, Aggregation::Max);
  ExpectNumbers(Column(Evaluate(table, cache_), 0), {70, 120, 300, 420, 420});
  table.mutable_data_fields()[0] = DataField("最小 / Amt", kAmt, Aggregation::Min);
  ExpectNumbers(Column(Evaluate(table, cache_), 0), {70, 90, 50, 140, 50});
}

TEST_F(PivotCalculated, PositionalReferencesCountDisplayedItems) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "Prev", "Region[-1]", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "Next", "Region[+1]", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "Far", "Region[-9]", host_)));
  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South", "Prev", "Next", "Far"}));
  const std::vector<Value> amt = Column(result, 0);
  ASSERT_EQ(amt.size(), 7U);
  EXPECT_DOUBLE_EQ(amt[3].as_number(), 350);
  for (const std::size_t i : {4U, 5U, 6U}) {
    ASSERT_TRUE(amt[i].is_error()) << i;
    EXPECT_EQ(amt[i].as_error(), ErrorCode::Ref) << i;
  }
}

TEST_F(PivotCalculated, ItemsSortByNameInASortedField) {
  PivotTable table = BuildTable();
  table.mutable_fields()[kRegion].sort.manual = false;
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "M1", "North+South", host_)));
  EXPECT_EQ(RowLabels(Evaluate(table, cache_)), (std::vector<std::string>{"East", "M1", "North", "South"}));
}

TEST_F(PivotCalculated, LaterItemWinsAtAnIntersectionAndMoveSwapsTheWinner) {
  PivotTable table = BuildTable(/*product_on_columns=*/true);
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "RX", "North+South", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kProduct, "PY", "Product[A]+100", host_)));
  {
    const PivotResult result = Evaluate(table, cache_);
    EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South", "RX"}));
    EXPECT_EQ(ColLabels(result), (std::vector<std::string>{"A", "B", "PY"}));
    EXPECT_TRUE(result.values[0][1][0].is_blank()) << "East has no B";
    EXPECT_DOUBLE_EQ(result.values[0][2][0].as_number(), 170);
    EXPECT_DOUBLE_EQ(result.values[3][2][0].as_number(), 520);
    EXPECT_DOUBLE_EQ(result.values[3][2][1].as_number(), 109);
    EXPECT_DOUBLE_EQ(result.row_leaf_totals[3][0].as_number(), 1080);
    EXPECT_DOUBLE_EQ(result.col_leaf_totals[2][0].as_number(), 1310);
    EXPECT_DOUBLE_EQ(result.grand_totals[0].as_number(), 2500);
  }
  ASSERT_TRUE(static_cast<bool>(move_calculated_item(cache_, 1, 0)));
  EXPECT_EQ(cache_.calculated_items()[0].formula, "Product[A]+100");
  {
    const PivotResult result = Evaluate(table, cache_);
    EXPECT_DOUBLE_EQ(result.values[3][2][0].as_number(), 620);
    EXPECT_DOUBLE_EQ(result.values[3][2][1].as_number(), 209);
    EXPECT_DOUBLE_EQ(result.row_leaf_totals[3][0].as_number(), 1180);
    EXPECT_DOUBLE_EQ(result.grand_totals[0].as_number(), 2600);
  }
  EXPECT_FALSE(static_cast<bool>(move_calculated_item(cache_, 0, 2)));
}

TEST_F(PivotCalculated, CalculatedFieldOnACalculatedItemRowSumsTheItemCells) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  table.mutable_data_fields().push_back(DataField("合計 / Price", 4));
  ExpectNumbers(Column(Evaluate(table, cache_), 2), {70.0 / 6.0, 30, 50, 40, 35});
}

TEST_F(PivotCalculated, AveragingFunctionsExcludeCalculatedItems) {
  PivotTable table = BuildTable();
  table.mutable_data_fields()[0].aggregation = Aggregation::Average;
  auto rejected = add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_);
  ASSERT_FALSE(static_cast<bool>(rejected));
  EXPECT_EQ(rejected.error().code, FormulonErrorCode::kPreconditionFailed);
  EXPECT_TRUE(cache_.calculated_items().empty());

  table.mutable_data_fields()[0].aggregation = Aggregation::Max;
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  for (const Aggregation aggregation :
       {Aggregation::Average, Aggregation::StdDev, Aggregation::StdDevP, Aggregation::Var, Aggregation::VarP}) {
    auto checked = check_aggregation_with_calculated_items(cache_, aggregation);
    ASSERT_FALSE(static_cast<bool>(checked));
    EXPECT_EQ(checked.error().code, FormulonErrorCode::kPreconditionFailed);
  }
  for (const Aggregation aggregation : {Aggregation::Sum, Aggregation::Count, Aggregation::Max, Aggregation::Min,
                                        Aggregation::Product, Aggregation::CountNumbers}) {
    EXPECT_TRUE(static_cast<bool>(check_aggregation_with_calculated_items(cache_, aggregation)));
  }
}

TEST_F(PivotCalculated, ItemTargetsAndNamesAreChecked) {
  PivotTable table = BuildTable();
  auto off_axis = add_calculated_item(cache_, {&table}, kProduct, "Y", "A", host_);
  ASSERT_FALSE(static_cast<bool>(off_axis));
  EXPECT_EQ(off_axis.error().code, FormulonErrorCode::kPreconditionFailed);

  ASSERT_TRUE(static_cast<bool>(add_calculated_field(cache_, {&table}, "Price", "Amt/Qty", host_)));
  table.mutable_fields()[4].axis = PivotAxis::Row;
  auto on_field = add_calculated_item(cache_, {&table}, 4, "Y", "1", host_);
  ASSERT_FALSE(static_cast<bool>(on_field));
  EXPECT_EQ(on_field.error().code, FormulonErrorCode::kPreconditionFailed);
  table.mutable_fields()[4].axis = PivotAxis::None;

  cache_.mutable_fields()[kRegion].field_group = FieldGroup{};
  auto grouped = add_calculated_item(cache_, {&table}, kRegion, "Y", "North", host_);
  ASSERT_FALSE(static_cast<bool>(grouped));
  EXPECT_EQ(grouped.error().code, FormulonErrorCode::kPreconditionFailed);
  cache_.mutable_fields()[kRegion].field_group.reset();

  for (const char* name : {"", "All", "a[1]", "north"}) {
    auto added = add_calculated_item(cache_, {&table}, kRegion, name, "North", host_);
    ASSERT_FALSE(static_cast<bool>(added)) << name;
    EXPECT_EQ(added.error().code, FormulonErrorCode::kInvalidArgument) << name;
  }
  for (const char* formula : {"Product[A]", "North+Amt"}) {
    auto added = add_calculated_item(cache_, {&table}, kRegion, "Y", formula, host_);
    ASSERT_FALSE(static_cast<bool>(added)) << formula;
    EXPECT_EQ(added.error().code, FormulonErrorCode::kInvalidArgument) << formula;
  }
  EXPECT_TRUE(cache_.calculated_items().empty());
  EXPECT_EQ(cache_.fields()[kRegion].shared_items.size(), 3U);
}

TEST_F(PivotCalculated, SetItemFormulaRejectsACycle) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "Y", "X*2", host_)));
  auto cycle = set_calculated_item_formula(cache_, 0, "Y+1", host_);
  ASSERT_FALSE(static_cast<bool>(cycle));
  EXPECT_EQ(cycle.error().code, FormulonErrorCode::kInvalidArgument);
  ASSERT_TRUE(static_cast<bool>(set_calculated_item_formula(cache_, 0, "=North*3", host_)));
  EXPECT_EQ(cache_.calculated_items()[0].formula, "North*3");
  ExpectNumbers(Column(Evaluate(table, cache_), 0), {70, 210, 350, 630, 1260, 2520});
}

TEST_F(PivotCalculated, RemovingAnItemRenumbersTheSharedItemsAfterIt) {
  PivotTable table = BuildTable();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "Y", "East*2", host_)));
  auto used = add_calculated_item(cache_, {&table}, kRegion, "Z", "X+1", host_);
  ASSERT_TRUE(static_cast<bool>(used));
  auto blocked = remove_calculated_item(cache_, {&table}, 0, host_);
  ASSERT_FALSE(static_cast<bool>(blocked));
  EXPECT_EQ(blocked.error().code, FormulonErrorCode::kPreconditionFailed);
  ASSERT_TRUE(static_cast<bool>(remove_calculated_item(cache_, {&table}, 2, host_)));

  ASSERT_TRUE(static_cast<bool>(remove_calculated_item(cache_, {&table}, 0, host_)));
  const PivotCacheField& region = cache_.fields()[kRegion];
  ASSERT_EQ(region.shared_items.size(), 4U);
  EXPECT_EQ(region.shared_items[3].as_text(), "Y");
  EXPECT_EQ(region.shared_item_is_calculated, (std::vector<bool>{false, false, false, true}));
  ASSERT_EQ(cache_.calculated_items().size(), 1U);
  EXPECT_EQ(cache_.calculated_items()[0].area.references[0].items, (std::vector<std::uint32_t>{3U}));
  const std::vector<PivotItem>& items = table.fields()[kRegion].items;
  ASSERT_EQ(items.size(), 4U);
  EXPECT_EQ(items[3].cache_index, 3U);
  EXPECT_TRUE(items[3].calculated);

  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South", "Y"}));
  ExpectNumbers(Column(result, 0), {70, 210, 350, 140, 770});
}

TEST_F(PivotCalculated, ItemOnAFieldWithoutItemsLandsAfterTheImplicitOnes) {
  PivotTable table = BuildTable();
  table.mutable_fields()[kRegion].items.clear();
  ASSERT_TRUE(static_cast<bool>(add_calculated_item(cache_, {&table}, kRegion, "X", "North+South", host_)));
  const std::vector<PivotItem>& items = table.fields()[kRegion].items;
  ASSERT_EQ(items.size(), 4U);
  for (std::uint32_t k = 0; k < 4U; ++k) {
    EXPECT_EQ(items[k].cache_index, k);
    EXPECT_EQ(items[k].calculated, k == 3U);
  }
  EXPECT_EQ(table.fields()[kProduct].items.size(), 2U);
}

TEST_F(PivotCalculated, ShowAllDrawsItemsOtherFieldsFilteredEmpty) {
  PivotTable table = BuildTable(/*product_on_columns=*/true);
  table.mutable_fields()[kProduct].items[0].visible = false;  // Hide A: East keeps no record.
  EXPECT_EQ(RowLabels(Evaluate(table, cache_)), (std::vector<std::string>{"North", "South"}));

  table.mutable_fields()[kRegion].show_all = true;
  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East", "North", "South"}));
  EXPECT_EQ(ColLabels(result), (std::vector<std::string>{"B"}));
  EXPECT_TRUE(result.values[0][0][0].is_blank());
  EXPECT_TRUE(result.row_leaf_totals[0][0].is_blank());
  EXPECT_DOUBLE_EQ(result.grand_totals[0].as_number(), 140);

  // An item hidden on the field itself stays hidden.
  table.mutable_fields()[kRegion].items[0].visible = false;
  EXPECT_EQ(RowLabels(Evaluate(table, cache_)), (std::vector<std::string>{"North", "South"}));
}

TEST_F(PivotCalculated, ShowAllOnColumnsDrawsAnEmptyColumn) {
  PivotTable table = BuildTable(/*product_on_columns=*/true);
  table.mutable_fields()[kRegion].items[1].visible = false;  // North
  table.mutable_fields()[kRegion].items[2].visible = false;  // South
  EXPECT_EQ(ColLabels(Evaluate(table, cache_)), (std::vector<std::string>{"A"}));

  table.mutable_fields()[kProduct].show_all = true;
  const PivotResult result = Evaluate(table, cache_);
  EXPECT_EQ(RowLabels(result), (std::vector<std::string>{"East"}));
  EXPECT_EQ(ColLabels(result), (std::vector<std::string>{"A", "B"}));
  EXPECT_DOUBLE_EQ(result.values[0][0][0].as_number(), 70);
  EXPECT_TRUE(result.values[0][1][0].is_blank());
  EXPECT_TRUE(result.col_leaf_totals[1][0].is_blank());
}

}  // namespace
}  // namespace formulon::pivot
