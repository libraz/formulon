//
// The rich detail a recalc gives an error cell: a measured `#CALC!` /
// `#VALUE!` origin, an error that arrived from another cell, or nothing for a
// plain error. Also the side record's lifecycle across writes, row edits and
// sheet copies.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_types.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/pivot_formula.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/xlsb/ptg.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_formula_host.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Shape = CellErrorDetail::Shape;

constexpr std::uint32_t kC = 2U;
constexpr std::uint32_t kE = 4U;
constexpr std::uint32_t kH = 7U;
constexpr std::uint32_t kX = 23U;

// H1:H3 = 1, 2, 3 and I1:I3 = 10, 20, 30: every empty FILTER below asks for values over 100.
Workbook Book() {
  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 3U; ++r) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, kH, Value::number(r + 1.0))));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0U, r, kH + 1U, Value::number(10.0 * (r + 1.0)))));
  }
  return wb;
}

void Set(Workbook& wb, std::uint32_t row, std::uint32_t col, const std::string& formula) {
  const Expected<void, Error> set = wb.set_cell_formula(0U, row, col, formula);
  ASSERT_TRUE(static_cast<bool>(set)) << formula;
}

void Recalc(Workbook& wb) {
  const Expected<void, Error> done = wb.ensure_calculated(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(done)) << done.error().message;
}

ErrorCode ErrorAt(const Workbook& wb, std::uint32_t row, std::uint32_t col) {
  const Value v = wb.sheet(0U).resolve_cell_value(row, col);
  EXPECT_TRUE(v.is_error()) << "row " << row << " col " << col;
  return v.is_error() ? v.as_error() : ErrorCode::Null;
}

void ExpectDetail(const Workbook& wb, std::uint32_t row, std::uint32_t col, ErrorCode code, Shape shape,
                  std::uint32_t sub_type = 0U, std::uint32_t iftab = 0U) {
  SCOPED_TRACE("row " + std::to_string(row) + " col " + std::to_string(col));
  EXPECT_EQ(ErrorAt(wb, row, col), code);
  const std::optional<CellErrorDetail> detail = wb.sheet(0U).error_detail_at(row, col);
  ASSERT_TRUE(detail.has_value());
  EXPECT_EQ(detail->code, code);
  EXPECT_EQ(detail->shape, shape);
  if (shape != Shape::kPropagated) {
    EXPECT_EQ(detail->sub_type, sub_type);
  }
  if (shape == Shape::kIftab) {
    EXPECT_EQ(detail->iftab, iftab);
  }
}

void ExpectPlain(const Workbook& wb, std::uint32_t row, std::uint32_t col, ErrorCode code) {
  SCOPED_TRACE("row " + std::to_string(row) + " col " + std::to_string(col));
  EXPECT_EQ(ErrorAt(wb, row, col), code);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(row, col).has_value());
}

constexpr const char* kEmptyFilter = "=FILTER(H1:H3,H1:H3>100)";

// ---------------------------------------------------------------------------
// Classification
// ---------------------------------------------------------------------------

TEST(ErrorDetail, PropagatedThroughAReference) {
  Workbook wb = Book();
  Set(wb, 0U, kC, kEmptyFilter);
  Set(wb, 0U, kE, "=C1");
  Set(wb, 1U, kE, "=C1&\"x\"");
  // C10 is a spill blocked by C11; its detail is the footprint, its readers' is propagated.
  Set(wb, 9U, kC, "=SEQUENCE(3)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_text(0U, 10U, kC, "x")));
  Set(wb, 9U, kE, "=C10");
  Set(wb, 10U, kE, "=LET(x,C10,x)");
  Set(wb, 11U, kE, "=IFERROR(C10,1)+C10");
  Recalc(wb);

  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  ExpectDetail(wb, 0U, kE, ErrorCode::Calc, Shape::kPropagated);
  ExpectDetail(wb, 1U, kE, ErrorCode::Calc, Shape::kPropagated);
  EXPECT_EQ(ErrorAt(wb, 9U, kC), ErrorCode::Spill);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(9U, kC).has_value());
  EXPECT_TRUE(wb.sheet(0U).blocked_spill_footprint_at(9U, kC).has_value());
  ExpectDetail(wb, 9U, kE, ErrorCode::Spill, Shape::kPropagated);
  ExpectDetail(wb, 10U, kE, ErrorCode::Spill, Shape::kPropagated);
  ExpectDetail(wb, 11U, kE, ErrorCode::Spill, Shape::kPropagated);
}

TEST(ErrorDetail, EmptyFilterIsCalcSubTypeThree) {
  Workbook wb = Book();
  const char* formulas[] = {
      kEmptyFilter,
      "=TAKE(H1:H3,0)",
      "=DROP(H1:H3,3)",
      "=TEXTSPLIT(\",\",\",\",,TRUE)",
      // An origin passing through the same formula keeps its form.
      "=FILTER(H1:H3,H1:H3>100)&\"x\"",
      "=ROWS(FILTER(H1:H3,H1:H3>100))",
      "=UNIQUE(FILTER(H1:H3,H1:H3>100))",
  };
  std::uint32_t row = 0;
  for (const char* formula : formulas) {
    Set(wb, row, kC, formula);
    row += 2U;
  }
  Recalc(wb);
  for (std::uint32_t r = 0; r < row; r += 2U) {
    SCOPED_TRACE(formulas[r / 2U]);
    ExpectDetail(wb, r, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  }
}

TEST(ErrorDetail, NestedArraysAndUncalledLambdaHaveTheirSubTypes) {
  Workbook wb = Book();
  Set(wb, 0U, kC, "=BYROW(H1:I2,LAMBDA(r,r))");
  Set(wb, 2U, kC, "=MAP(H1:H2,LAMBDA(x,SEQUENCE(2)))");
  Set(wb, 4U, kC, "=SCAN(0,H1:H2,LAMBDA(a,x,SEQUENCE(2)))");
  Set(wb, 6U, kC, "=MAKEARRAY(2,2,LAMBDA(r,c,SEQUENCE(2)))");
  Set(wb, 8U, kC, "=LAMBDA(x,x)");
  Set(wb, 10U, kC, "=BYCOL(H1:I2,LAMBDA(c,c))");
  Recalc(wb);
  ExpectDetail(wb, 10U, kC, ErrorCode::Calc, Shape::kSubType, 2U);
  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 2U);
  ExpectDetail(wb, 2U, kC, ErrorCode::Calc, Shape::kSubType, 1U);
  ExpectDetail(wb, 4U, kC, ErrorCode::Calc, Shape::kSubType, 1U);
  ExpectDetail(wb, 6U, kC, ErrorCode::Calc, Shape::kSubType, 1U);
  ExpectDetail(wb, 8U, kC, ErrorCode::Calc, Shape::kSubType, 24U);
}

TEST(ErrorDetail, CommittedSpillAnchorKeepsItsElementsOrigin) {
  Workbook wb = Book();
  // Every element is an empty FILTER; the spill commits and its anchor holds the first.
  Set(wb, 0U, kE, "=BYROW(H1:H3,LAMBDA(r,FILTER(r,r>100)))");
  Set(wb, 0U, kC, "=SEQUENCE(3)/{1;0;1}");
  Recalc(wb);
  ExpectDetail(wb, 0U, kE, ErrorCode::Calc, Shape::kSubType, 3U);
  EXPECT_TRUE(wb.sheet(0U).error_detail_at(0U, kE)->elements);
  EXPECT_EQ(ErrorAt(wb, 2U, kE), ErrorCode::Calc);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(2U, kE).has_value());
  // Classic error elements carry no origin.
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(0U, kC).as_number(), 1.0);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(0U, kC).has_value());
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(1U, kC).has_value());
}

TEST(ErrorDetail, ElementsOfANonErrorAnchorCarryTheOrigin) {
  Workbook wb = Book();
  // The first element is a number; the others are empty FILTERs.
  Set(wb, 0U, kE, "=BYROW(H1:H3,LAMBDA(r,IF(r>1,FILTER(r,r>100),r)))");
  Recalc(wb);
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(0U, kE).as_number(), 1.0);
  EXPECT_EQ(ErrorAt(wb, 1U, kE), ErrorCode::Calc);
  const std::optional<CellErrorDetail> detail = wb.sheet(0U).error_detail_at(0U, kE);
  ASSERT_TRUE(detail.has_value());
  EXPECT_TRUE(detail->elements);
  EXPECT_EQ(detail->code, ErrorCode::Calc);
  EXPECT_EQ(detail->shape, Shape::kSubType);
  EXPECT_EQ(detail->sub_type, 3U);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(1U, kE).has_value());
}

TEST(ErrorDetail, LeftErrorOperandWins) {
  Workbook wb = Book();
  Set(wb, 4U, kC, "=BYROW(H1:I2,LAMBDA(r,r))");
  Set(wb, 0U, kE, "=C5&FILTER(H1:H3,H1:H3>100)");
  Set(wb, 1U, kE, "=FILTER(H1:H3,H1:H3>100)&C5");
  Set(wb, 2U, kE, "=C5+FILTER(H1:H3,H1:H3>100)+1");
  Recalc(wb);
  ExpectDetail(wb, 4U, kC, ErrorCode::Calc, Shape::kSubType, 2U);
  ExpectDetail(wb, 0U, kE, ErrorCode::Calc, Shape::kPropagated);
  ExpectDetail(wb, 1U, kE, ErrorCode::Calc, Shape::kSubType, 3U);
  ExpectDetail(wb, 2U, kE, ErrorCode::Calc, Shape::kPropagated);
}

TEST(ErrorDetail, RichValueDependentIsPropagatedWithPtgThroughConcat) {
  struct Case {
    const char* formula;
    Shape shape;
    std::uint32_t ptg;
    const char* root_function;
    std::uint8_t root_argpos;
  };
  const Case cases[] = {
      {"=C5&\"x\"", Shape::kPropagatedPtg, 8U, "", 0U},
      {"=C5+1", Shape::kPropagatedPtg, 3U, "", 0U},
      {"=-C5", Shape::kPropagated, 0U, "", 0U},
      {"=ABS(C5)", Shape::kPropagatedIftab, 0U, "ABS", 1U},
      {"=SUM(1,C5)", Shape::kPropagatedIftab, 0U, "SUM", 2U},
      {"=C5", Shape::kPropagated, 0U, "", 0U},
      {"=IF(C5,1)", Shape::kPropagatedIftab, 0U, "IF", 1U},
      {"=IF(1,C5)", Shape::kPropagatedIftab, 0U, "IF", 2U},
      // A called LAMBDA is shaped by its body.
      {"=LAMBDA(x,x+0)(C5)", Shape::kPropagatedPtg, 3U, "", 0U},
      {"=LAMBDA(x,ABS(x))(C5)", Shape::kPropagatedIftab, 0U, "ABS", 1U},
  };
  Workbook wb = Book();
  // C5 holds a rich #VALUE! of its own.
  Set(wb, 4U, kC, "=GROUPBY(H1:H3,I1:I3,1)");
  std::uint32_t row = 0;
  for (const Case& entry : cases) {
    Set(wb, row++, kE, entry.formula);
  }
  Set(wb, row, kE, "=\"a\"+1");
  Recalc(wb);
  ExpectDetail(wb, 4U, kC, ErrorCode::Value, Shape::kSubType, 37U);
  row = 0;
  for (const Case& entry : cases) {
    SCOPED_TRACE(entry.formula);
    const std::uint32_t r = row++;
    EXPECT_EQ(ErrorAt(wb, r, kE), ErrorCode::Value);
    const std::optional<CellErrorDetail> detail = wb.sheet(0U).error_detail_at(r, kE);
    ASSERT_TRUE(detail.has_value());
    EXPECT_EQ(detail->code, ErrorCode::Value);
    EXPECT_EQ(detail->shape, entry.shape);
    // The engine records the operator; the Ptg id is the writer's mapping.
    const std::uint32_t ptg =
        detail->root_binop == 0U ? 0U : io::xlsb::ptg_of_binop(static_cast<parser::BinOp>(detail->root_binop - 1U));
    EXPECT_EQ(ptg, entry.ptg);
    EXPECT_EQ(detail->ptg, 0U);
    EXPECT_EQ(detail->root_function, std::string_view(entry.root_function));
    EXPECT_EQ(detail->root_argpos, entry.root_argpos);
    EXPECT_EQ(detail->iftab, 0U);
  }
  // No precedent holds a rich #VALUE!.
  ExpectPlain(wb, row, kE, ErrorCode::Value);
}

TEST(ErrorDetail, TextCapCarriesItsIftab) {
  Workbook wb = Book();
  Set(wb, 0U, kC, "=CONCAT(REPT(\"a\",20000),REPT(\"a\",20000))");
  Set(wb, 1U, kC, "=TEXTJOIN(\",\",TRUE,REPT(\"a\",20000),REPT(\"a\",20000))");
  Recalc(wb);
  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kIftab, 9U, 606U);
  ExpectDetail(wb, 1U, kC, ErrorCode::Calc, Shape::kIftab, 9U, 605U);
}

TEST(ErrorDetail, GroupbyFunctionArgumentIsValueSubType37) {
  Workbook wb = Book();
  Set(wb, 0U, kC, "=GROUPBY(H1:H3,I1:I3,1)");
  Set(wb, 4U, kC, "=\"a\"+1");
  Set(wb, 5U, kC, "=1/0");
  Recalc(wb);
  ExpectDetail(wb, 0U, kC, ErrorCode::Value, Shape::kSubType, 37U);
  ExpectPlain(wb, 4U, kC, ErrorCode::Value);
  ExpectPlain(wb, 5U, kC, ErrorCode::Div0);
}

TEST(ErrorDetail, IfIsErrorFallbackIsPropagated) {
  Workbook wb = Book();
  Set(wb, 9U, kC, "=BYROW(H1:I2,LAMBDA(r,r))");
  Set(wb, 0U, kE, "=IF(ISERROR(FILTER(H1:H3,H1:H3>100)),C10,0)");
  Set(wb, 1U, kE, "=IFERROR(FILTER(H1:H3,H1:H3>100),C10)");
  // The consumed GROUPBY origin does not label the fallback's own error.
  Set(wb, 2U, kE, "=IFERROR(GROUPBY(H1:H3,I1:I3,1),\"a\"+1)");
  Recalc(wb);
  ExpectDetail(wb, 9U, kC, ErrorCode::Calc, Shape::kSubType, 2U);
  ExpectDetail(wb, 0U, kE, ErrorCode::Calc, Shape::kPropagated);
  ExpectDetail(wb, 1U, kE, ErrorCode::Calc, Shape::kPropagated);
  ExpectPlain(wb, 2U, kE, ErrorCode::Value);
}

TEST(ErrorDetail, UnknownProducerIsPlain) {
  Workbook wb = Book();
  // TOCOL left with nothing is not measured.
  Set(wb, 0U, kC, "=TOCOL({1,2}/0,2)");
  Recalc(wb);
  ExpectPlain(wb, 0U, kC, ErrorCode::Calc);
}

TEST(ErrorDetail, ACellRecomputedToANumberClearsItsDetail) {
  Workbook wb = Book();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kX, Value::number(100.0))));
  Set(wb, 0U, kC, "=SUM(FILTER(H1:H3,H1:H3>X1))");
  Recalc(wb);
  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kX, Value::number(0.0))));
  Recalc(wb);
  EXPECT_EQ(wb.sheet(0U).resolve_cell_value(0U, kC).as_number(), 6.0);
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(0U, kC).has_value());
  EXPECT_FALSE(wb.sheet(0U).has_error_details());
}

TEST(ErrorDetail, AWriteClearsTheDetail) {
  Workbook wb = Book();
  Set(wb, 0U, kC, kEmptyFilter);
  Set(wb, 1U, kC, kEmptyFilter);
  Recalc(wb);
  ASSERT_TRUE(wb.sheet(0U).error_detail_at(0U, kC).has_value());
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, kC, Value::number(1.0))));
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(0U, kC).has_value());
  Set(wb, 1U, kC, "=1");
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(1U, kC).has_value());
  EXPECT_FALSE(wb.sheet(0U).has_error_details());
}

TEST(ErrorDetail, RowInsertAndDeleteRemapTheDetail) {
  Workbook wb = Book();
  Set(wb, 4U, kC, kEmptyFilter);
  Recalc(wb);
  // Rows 4 onward, below H1:H3, so the formula text and its cached error stay.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 3U, 2U)));
  EXPECT_FALSE(wb.sheet(0U).error_detail_at(4U, kC).has_value());
  ExpectDetail(wb, 6U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0U, 3U, 1U)));
  ExpectDetail(wb, 5U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0U, 5U, 1U)));
  EXPECT_FALSE(wb.sheet(0U).has_error_details());
}

TEST(ErrorDetail, CopySheetCopiesTheDetails) {
  Workbook wb = Book();
  Set(wb, 0U, kC, kEmptyFilter);
  Recalc(wb);
  const Expected<std::size_t, Error> copied = wb.copy_sheet(0U, 1U, std::nullopt);
  ASSERT_TRUE(static_cast<bool>(copied)) << copied.error().message;
  const std::optional<CellErrorDetail> detail = wb.sheet(copied.value()).error_detail_at(0U, kC);
  ASSERT_TRUE(detail.has_value());
  EXPECT_EQ(detail->sub_type, 3U);
  EXPECT_TRUE(wb.sheet(copied.value()).has_error_details());
}

// ---------------------------------------------------------------------------
// Evaluations that never commit
// ---------------------------------------------------------------------------

TEST(ErrorDetail, CfAtAnErrorCellKeepsItsDetail) {
  Workbook wb = Book();
  Set(wb, 0U, kC, kEmptyFilter);
  Recalc(wb);
  ASSERT_TRUE(wb.sheet(0U).error_detail_at(0U, kC).has_value());

  Sheet& sheet = wb.sheet(0U);
  Arena arena;
  eval::EvalState state;
  const eval::EvalContext eval_ctx(wb, sheet, state);
  cf::CFRule rule;
  rule.type = cf::RuleType::Expression;
  // A rule whose formula produces and consumes errors of other origins.
  rule.formula1 = "ISERROR(BYROW($H$1:$I$2,LAMBDA(r,r)))";
  cf::CFEvalContext ctx;
  ctx.anchor = CellAddress{0U, kC};
  ctx.target = CellAddress{0U, kC};
  ctx.arena = &arena;
  ctx.registry = &eval::default_registry();
  ctx.eval_ctx = &eval_ctx;
  EXPECT_TRUE(cf::match_rule(rule, sheet.resolve_cell_value(0U, kC), ctx));

  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  Recalc(wb);
  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
}

TEST(ErrorDetail, CalculatedFieldLeavesTheEnclosingOriginAlone) {
  const eval::ScopedEvalProfile profile(mac_365_en_us_profile());
  pivot::PivotCache cache;
  cache.set_cache_id(1U);
  cache.mutable_fields().push_back(pivot::PivotCacheField{"Amt", {}});
  eval::PivotFormulaEvaluator host;
  std::vector<pivot::FormulaRef> refs;
  const Expected<std::uint32_t, Error> id =
      host.compile("=ISERROR(1/0)", pivot::FormulaScope{&cache, std::nullopt}, &refs, nullptr);
  ASSERT_TRUE(static_cast<bool>(id)) << id.error().message;

  const eval::ErrorOriginScope cell_scope;
  eval::note_calc_sub_type(3U);
  const Value v = host.evaluate(id.value(), nullptr, 0U);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
  EXPECT_EQ(eval::error_origin(), eval::ErrorOrigin::kMeasured);
  EXPECT_EQ(eval::noted_error_detail().sub_type, 3U);
}

// ---------------------------------------------------------------------------
// The side record itself
// ---------------------------------------------------------------------------

TEST(ErrorDetail, SideRecordCountGuardsEveryAccessor) {
  Sheet sheet("Sheet1");
  EXPECT_FALSE(sheet.has_error_details());
  EXPECT_FALSE(sheet.error_detail_at(0U, 0U).has_value());
  sheet.clear_error_detail(0U, 0U);
  sheet.set_error_detail(1U, 2U, CellErrorDetail{ErrorCode::Calc, Shape::kSubType, 3U, 0U});
  sheet.set_error_detail(4U, 5U, CellErrorDetail{ErrorCode::Spill, Shape::kPropagated, 0U, 0U});
  EXPECT_TRUE(sheet.has_error_details());
  EXPECT_EQ(sheet.error_details().size(), 2U);
  sheet.clear_error_detail(1U, 2U);
  EXPECT_FALSE(sheet.error_detail_at(1U, 2U).has_value());
  EXPECT_TRUE(sheet.has_error_details());
  sheet.replace_error_details({{CellAddress{7U, 8U}, CellErrorDetail{ErrorCode::Calc, Shape::kPropagated, 0U, 0U}}});
  EXPECT_FALSE(sheet.error_detail_at(4U, 5U).has_value());
  ASSERT_TRUE(sheet.error_detail_at(7U, 8U).has_value());
  sheet.replace_error_details({});
  EXPECT_FALSE(sheet.has_error_details());
}

TEST(ErrorDetail, RestoreRecordsTheLoadedDetail) {
  Workbook wb = Book();
  Set(wb, 0U, kC, kEmptyFilter);
  ASSERT_TRUE(static_cast<bool>(
      wb.restore_error_detail(0U, 0U, kC, ErrorCode::Calc, CellErrorDetail{ErrorCode::Calc, Shape::kSubType, 3U, 0U})));
  ExpectDetail(wb, 0U, kC, ErrorCode::Calc, Shape::kSubType, 3U);
  EXPECT_FALSE(static_cast<bool>(wb.restore_error_detail(5U, 0U, kC, ErrorCode::Calc, CellErrorDetail{})));
}

// The detail lives beside the cell store, never in it.
struct CellLayoutMirror {
  std::string formula_text;
  Value cached_value = Value::blank();
  std::unique_ptr<std::string> cached_text_owned;
  std::vector<PhoneticRun> phonetic_runs;
  std::uint32_t xf_index = 0;
  PhoneticProperties phonetic_props;
  bool dynamic_array : 1;
  bool has_explicit_xf : 1;
  bool show_phonetic : 1;
  bool phonetic_present : 1;
  std::uint16_t array_block_cols = 0;
  std::uint32_t array_block_rows = 0;
};
static_assert(sizeof(Cell) == sizeof(CellLayoutMirror), "Cell gained a member");
static_assert(sizeof(Value) <= 24U, "Value grew");

}  // namespace
}  // namespace formulon
