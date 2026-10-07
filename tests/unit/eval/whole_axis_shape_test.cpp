//
// A whole column / row in an operand, a LET binding or a scalar function
// argument keeps its declared shape: `L!B:B+0` is 1048576x1 with blank cells
// behaving as blanks (0 under arithmetic, "" under `&`, TRUE for ISBLANK).
//
// Sheet L: A1:A5 = "k1","k2","k3",(blank),"k5"; B1:B5 = 10..50; column C is
// empty. Formulas sit on sheet H unless a test says otherwise. Expected values
// are the Mac Excel 365 ja-JP measurements recorded in
// tests/oracle/cases/whole_axis_spill.yaml.

#include <cstddef>
#include <cstdint>
#include <string>

#include "eval/adhoc_eval.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kH = 0U;  // Sheet index of H.
constexpr std::uint32_t kL = 1U;  // Sheet index of L.
constexpr std::uint32_t kHCol = 7U;
constexpr std::uint32_t kLastRow = Sheet::kMaxRows - 1U;

// One grid column of `Value`s; a derived array built densely cannot come in
// under it.
constexpr std::size_t kWholeColumnBytes = static_cast<std::size_t>(Sheet::kMaxRows) * sizeof(Value);

class WholeAxisShape : public ::testing::Test {
 protected:
  WholeAxisShape() : wb_(Workbook::create()) {
    EXPECT_TRUE(static_cast<bool>(wb_.rename_sheet(0U, "H")));
    wb_.add_sheet("L");
    const char* const keys[] = {"k1", "k2", "k3", nullptr, "k5"};
    for (std::uint32_t r = 0; r < 5U; ++r) {
      if (keys[r] != nullptr) {
        EXPECT_TRUE(static_cast<bool>(wb_.set_cell_text(kL, r, 0U, keys[r])));
      }
      EXPECT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, r, 1U, Value::number(10.0 * (r + 1U)))));
    }
  }

  // Stores `formula` at (`row`, `col`) of `sheet` and recalcs.
  void Put(std::uint32_t sheet, std::uint32_t row, std::uint32_t col, const char* formula) {
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_formula(sheet, row, col, formula)));
    ASSERT_TRUE(static_cast<bool>(wb_.recalc(default_registry())));
  }

  Value At(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) {
    return wb_.sheet(sheet).resolve_cell_value(row, col);
  }

  const SpillRegion* Region(std::uint32_t sheet, std::uint32_t row, std::uint32_t col) {
    return wb_.sheet(sheet).spill_region_at_anchor(row, col);
  }

  // Read-only evaluation of `formula` anchored at H1.
  Value Adhoc(const std::string& formula) {
    return evaluate_formula_text_array(wb_, wb_.sheet(kH), 0U, kHCol, formula, arena_, default_registry());
  }

  Workbook wb_;
  Arena arena_;
};

TEST_F(WholeAxisShape, ColumnPlusZeroSpillsDeclaredHeight) {
  Put(kH, 0U, kHCol, "=L!B:B+0");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(region->cols, 1U);
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), 10.0);
  EXPECT_EQ(At(kH, 4U, kHCol).as_number(), 50.0);
  EXPECT_EQ(At(kH, 5U, kHCol).as_number(), 0.0);
  EXPECT_EQ(At(kH, kLastRow, kHCol).as_number(), 0.0);
}

TEST_F(WholeAxisShape, ColumnPlusZeroBelowRowOneIsSpill) {
  Put(kH, 1U, kHCol, "=L!B:B+0");
  const Value v = At(kH, 1U, kHCol);
  ASSERT_TRUE(v.is_error());
  EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  EXPECT_EQ(Region(kH, 1U, kHCol), nullptr);
}

TEST_F(WholeAxisShape, LetBoundColumnSpills) {
  Put(kH, 0U, kHCol, "=LET(x,L!B:B,x)");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), 10.0);
  EXPECT_EQ(At(kH, 5U, kHCol).as_number(), 0.0);
  EXPECT_EQ(At(kH, kLastRow, kHCol).as_number(), 0.0);
}

TEST_F(WholeAxisShape, MultiColumnPlusOneSpillsDeclaredRectangle) {
  Put(kH, 0U, kHCol, "=L!A:C+1");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(region->cols, 3U);
  const Value text_plus_one = At(kH, 0U, kHCol);
  ASSERT_TRUE(text_plus_one.is_error());
  EXPECT_EQ(text_plus_one.as_error(), ErrorCode::Value);
  EXPECT_EQ(At(kH, 0U, kHCol + 1U).as_number(), 11.0);
  EXPECT_EQ(At(kH, 0U, kHCol + 2U).as_number(), 1.0);
  for (std::uint32_t c = 0; c < 3U; ++c) {
    EXPECT_EQ(At(kH, 5U, kHCol + c).as_number(), 1.0);
    EXPECT_EQ(At(kH, kLastRow, kHCol + c).as_number(), 1.0);
  }
}

TEST_F(WholeAxisShape, ColumnPlusShorterColumnArrayIsNaPastIt) {
  Put(kH, 0U, kHCol, "=L!B:B+{1;2;3}");
  ASSERT_NE(Region(kH, 0U, kHCol), nullptr);
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), 11.0);
  EXPECT_EQ(At(kH, 2U, kHCol).as_number(), 33.0);
  for (const std::uint32_t row : {3U, 4U, 5U, kLastRow}) {
    const Value v = At(kH, row, kHCol);
    ASSERT_TRUE(v.is_error()) << row;
    EXPECT_EQ(v.as_error(), ErrorCode::NA) << row;
  }
}

TEST_F(WholeAxisShape, ColumnPlusRowArrayStretchesTheTail) {
  Put(kH, 0U, kHCol, "=L!B:B+{1,2}");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(region->cols, 2U);
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), 11.0);
  EXPECT_EQ(At(kH, 0U, kHCol + 1U).as_number(), 12.0);
  EXPECT_EQ(At(kH, 5U, kHCol).as_number(), 1.0);
  EXPECT_EQ(At(kH, 5U, kHCol + 1U).as_number(), 2.0);
  EXPECT_EQ(At(kH, kLastRow, kHCol).as_number(), 1.0);
  EXPECT_EQ(At(kH, kLastRow, kHCol + 1U).as_number(), 2.0);
}

TEST_F(WholeAxisShape, WholeRowPlusZeroSpillsDeclaredWidth) {
  Put(kH, 0U, 0U, "=L!1:1+0");
  const SpillRegion* region = Region(kH, 0U, 0U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 1U);
  EXPECT_EQ(region->cols, Sheet::kMaxCols);
  const Value text_plus_zero = At(kH, 0U, 0U);
  ASSERT_TRUE(text_plus_zero.is_error());
  EXPECT_EQ(text_plus_zero.as_error(), ErrorCode::Value);
  EXPECT_EQ(At(kH, 0U, 1U).as_number(), 10.0);
  EXPECT_EQ(At(kH, 0U, 2U).as_number(), 0.0);
  EXPECT_EQ(At(kH, 0U, Sheet::kMaxCols - 1U).as_number(), 0.0);
}

TEST_F(WholeAxisShape, ConcatenationReadsBlankAsEmptyText) {
  Put(kH, 0U, kHCol, "=L!A:A&\"x\"");
  ASSERT_NE(Region(kH, 0U, kHCol), nullptr);
  EXPECT_EQ(At(kH, 0U, kHCol).as_text(), "k1x");
  EXPECT_EQ(At(kH, 3U, kHCol).as_text(), "x");
  EXPECT_EQ(At(kH, kLastRow, kHCol).as_text(), "x");
}

TEST_F(WholeAxisShape, ScalarFunctionOverAColumnKeepsItsTail) {
  Put(kH, 0U, kHCol, "=ISBLANK(L!A:A)");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_FALSE(At(kH, 0U, kHCol).as_boolean());
  EXPECT_TRUE(At(kH, 3U, kHCol).as_boolean());
  EXPECT_TRUE(At(kH, kLastRow, kHCol).as_boolean());
}

TEST_F(WholeAxisShape, ColumnTimesTwoSpillsOnItsOwnSheet) {
  Put(kL, 0U, 3U, "=L!B:B*2");
  const SpillRegion* region = Region(kL, 0U, 3U);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, Sheet::kMaxRows);
  EXPECT_EQ(At(kL, 0U, 3U).as_number(), 20.0);
  EXPECT_EQ(At(kL, 4U, 3U).as_number(), 100.0);
  EXPECT_EQ(At(kL, 5U, 3U).as_number(), 0.0);
  EXPECT_EQ(At(kL, kLastRow, 3U).as_number(), 0.0);
}

TEST_F(WholeAxisShape, RowsOfAScalarFunctionOverAColumnIsDeclaredHeight) {
  EXPECT_EQ(Adhoc("=ROWS(ABS(L!B:B))").as_number(), static_cast<double>(Sheet::kMaxRows));
  EXPECT_EQ(Adhoc("=ROWS(L!B:B+0)").as_number(), static_cast<double>(Sheet::kMaxRows));
  EXPECT_EQ(Adhoc("=LET(x,L!B:B,ROWS(x))").as_number(), static_cast<double>(Sheet::kMaxRows));
}

TEST_F(WholeAxisShape, AggregatesOverDerivedColumns) {
  EXPECT_EQ(Adhoc("=SUM(L!B:B*1)").as_number(), 150.0);
  EXPECT_EQ(Adhoc("=SUM((L!A:A=\"k2\")*L!B:B)").as_number(), 20.0);
  EXPECT_EQ(Adhoc("=MAX(L!B:B-1)").as_number(), 49.0);
  EXPECT_EQ(Adhoc("=COUNT(L!B:B+0)").as_number(), static_cast<double>(Sheet::kMaxRows));
  EXPECT_EQ(Adhoc("=AVERAGE(IF(L!B:B>15,L!B:B))").as_number(), 35.0);
  EXPECT_EQ(Adhoc("=SUM(L!B:B*0+1)").as_number(), static_cast<double>(Sheet::kMaxRows));
  EXPECT_EQ(Adhoc("=INDEX(L!B:B+0,3)").as_number(), 30.0);
}

TEST_F(WholeAxisShape, EditingACellBelowTheDataRecalculatesDerivedColumns) {
  Put(kH, 0U, kHCol, "=SUM(L!B:B*1)");
  Put(kH, 0U, kHCol + 2U, "=L!B:B+0");
  ASSERT_EQ(At(kH, 0U, kHCol).as_number(), 150.0);
  ASSERT_EQ(At(kH, 6U, kHCol + 2U).as_number(), 0.0);

  ASSERT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, 6U, 1U, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb_.recalc(default_registry())));
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), 155.0);
  EXPECT_EQ(At(kH, 6U, kHCol + 2U).as_number(), 5.0);
}

TEST_F(WholeAxisShape, ScalarFunctionOverAColumnStaysCompactUntilExpanded) {
  // The compressed form is what keeps a refused spill from costing a column of
  // cells: the footprint is decided on the declared size alone.
  std::size_t bytes = 0;
  {
    Arena arena;
    const Value v =
        evaluate_formula_text_array(wb_, wb_.sheet(kH), 1U, kHCol, "=ISBLANK(L!A:A)", arena, default_registry());
    bytes = arena.bytes_used();
    ASSERT_TRUE(v.is_error());
    EXPECT_EQ(v.as_error(), ErrorCode::Spill);
  }
  EXPECT_LT(bytes, kWholeColumnBytes / 16U);
}

TEST(WholeAxisShapeFirstElement, ReducesAnArrayToItsFirstElementWithoutExpandingIt) {
  Sheet sheet("Sheet1");
  for (std::uint32_t r = 0; r < 5U; ++r) {
    sheet.set_cell_value(r, 1U, Value::number(10.0 * (r + 1U)));
  }
  const auto first = [&](std::string_view formula, std::size_t* bytes) {
    Arena parse_arena;
    Arena eval_arena;
    parser::Parser p(formula, parse_arena);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << formula;
    const EvalContext ctx(sheet);
    const Value v = evaluate_first_element(*root, eval_arena, default_registry(), ctx);
    *bytes = eval_arena.bytes_used();
    return v;
  };
  std::size_t bytes = 0;
  Value v = first("=$B:$B>25", &bytes);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_FALSE(v.as_boolean());
  EXPECT_LT(bytes, kWholeColumnBytes / 16U);

  sheet.set_cell_value(0U, 1U, Value::number(100.0));
  v = first("=$B:$B+0>25", &bytes);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());
  EXPECT_LT(bytes, kWholeColumnBytes / 16U);

  v = first("=$B$1:$B$5>25", &bytes);
  ASSERT_TRUE(v.is_boolean());
  EXPECT_TRUE(v.as_boolean());

  v = first("=7", &bytes);
  ASSERT_TRUE(v.is_number());
  EXPECT_EQ(v.as_number(), 7.0);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
