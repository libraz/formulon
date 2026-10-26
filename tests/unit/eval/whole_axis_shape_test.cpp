//
// A whole column / row in an operand, a LET binding or a scalar function
// argument keeps its declared shape: `L!B:B+0` is 1048576x1 with blank cells
// behaving as blanks (0 under arithmetic, "" under `&`, TRUE for ISBLANK).
//
// Sheet L: A1:A5 = "k1","k2","k3",(blank),"k5"; B1:B5 = 10..50; column C is
// empty. Formulas sit on sheet H unless a test says otherwise. Expected values
// are the Mac Excel 365 ja-JP results recorded in
// tests/oracle/cases/whole_axis_spill.yaml.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <vector>

#include "eval/adhoc_eval.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/range_args.h"
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

TEST_F(WholeAxisShape, ArrayPastExcelsElementBoundMakesTheFormulaZero) {
  // A whole column plus a whole row is the full grid, and 103 whole columns
  // or a 1048576 x 103 generator are 108,003,328 elements: all are past the
  // bound, so the formula holds 0 whatever encloses the array (Mac Excel:
  // SUM(A:CZ+0) and ROWS(A:A+1:1) are 0). A reader of the cell sees that 0.
  const char* const formulas[] = {"=SUM(L!A:A+L!1:1)",       "=ISERROR(L!A:A+L!1:1)",
                                  "=ROWS(L!A:A+L!1:1)",      "=SUM(IFERROR(L!A:A+L!1:1,1))",
                                  "=\"x\"&SUM(L!A:A+L!1:1)", "=SUM(SEQUENCE(1048576,103)*0+1)",
                                  "=SUM(L!A:CY+0)",          "=SUM(-L!A:CY)"};
  for (const char* formula : formulas) {
    Put(kH, 0U, kHCol, formula);
    const Value v = At(kH, 0U, kHCol);
    ASSERT_TRUE(v.is_number()) << formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), 0.0) << formula;
    EXPECT_EQ(Region(kH, 0U, kHCol), nullptr) << formula;
  }
  Put(kH, 1U, kHCol, "=H1+5");
  EXPECT_EQ(At(kH, 1U, kHCol).as_number(), 5.0);
}

TEST_F(WholeAxisShape, RowsAndColumnsSizeAnOperationWithoutBuildingIt) {
  // A SEQUENCE under ROWS / COLUMNS is sized, not built, and so is an
  // operation, up to the one element bound (Mac Excel: ROWS(-A:BB),
  // ROWS(A1:A3277+A1:XFD1), ROWS(A:CX+0) answer; ROWS(A:CY+0) is 0). Both
  // drivers agree.
  const struct {
    const char* formula;
    double expected;
  } cases[] = {{"=ROWS(SEQUENCE(4096,16384))", 4096.0},
               {"=COLUMNS(SEQUENCE(4096,16384))", 16384.0},
               {"=ROWS(L!A:A+L!1:1)", 0.0},
               {"=COLUMNS(L!A:A+L!1:1)", 0.0},
               {"=ROWS(L!A1:A3277+L!1:1)", 3277.0},
               {"=ROWS(SEQUENCE(3277,16384)+0)", 3277.0},
               {"=ROWS(L!A:CX+0)", static_cast<double>(Sheet::kMaxRows)},
               {"=COLUMNS(L!A:CX+0)", 102.0},
               {"=ROWS(L!A:CY+0)", 0.0},
               {"=ROWS(-L!A:CY)", 0.0},
               {"=COLUMNS(L!A:CZ+0)", 0.0},
               {"=ROWS(L!A:AY+0)", static_cast<double>(Sheet::kMaxRows)},
               {"=ROWS(-L!A:BB)", static_cast<double>(Sheet::kMaxRows)},
               {"=ROWS(--L!A:BB)", static_cast<double>(Sheet::kMaxRows)},
               {"=ROWS(L!A:BB+0)", static_cast<double>(Sheet::kMaxRows)},
               {"=COLUMNS(-L!A:BB)", 54.0},
               {"=ROWS(-L!A:A)", static_cast<double>(Sheet::kMaxRows)},
               {"=COLUMNS(L!A1:A3+{1,2})", 2.0}};
  for (const auto& c : cases) {
    Put(kH, 0U, kHCol, c.formula);
    const Value v = At(kH, 0U, kHCol);
    ASSERT_TRUE(v.is_number()) << c.formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
    const Value adhoc = Adhoc(c.formula);
    ASSERT_TRUE(adhoc.is_number()) << c.formula << " " << adhoc.debug_to_string();
    EXPECT_EQ(adhoc.as_number(), c.expected) << c.formula;
  }
}

TEST_F(WholeAxisShape, ScalarErrorAgainstAnArrayFillsTheBroadcast) {
  // A scalar error meeting an array operand is broadcast error per position
  // rather than ending the operation (Mac Excel: ROWS of each is 3), while
  // a scalar error alone stays the error. Both drivers agree.
  const char* const sized[] = {"=ROWS(1/0+{1;2;3})", "=ROWS(IF(TRUE,1/0+{1;2;3}))", "=ROWS(IFERROR(1/0+{1;2;3},0))",
                               "=ROWS(ISERROR(1/0+{1;2;3}))", "=ROWS({1;2;3}+1/0)"};
  for (const char* formula : sized) {
    Put(kH, 0U, kHCol, formula);
    const Value v = At(kH, 0U, kHCol);
    ASSERT_TRUE(v.is_number()) << formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), 3.0) << formula;
    const Value adhoc = Adhoc(formula);
    ASSERT_TRUE(adhoc.is_number()) << formula << " " << adhoc.debug_to_string();
    EXPECT_EQ(adhoc.as_number(), 3.0) << formula;
  }
  Put(kH, 0U, kHCol, "=ROWS(1/0)");
  ASSERT_TRUE(At(kH, 0U, kHCol).is_error());
  EXPECT_EQ(At(kH, 0U, kHCol).as_error(), ErrorCode::Div0);

  Put(kH, 0U, kHCol, "=1/0+{1;2;3}");
  const SpillRegion* region = Region(kH, 0U, kHCol);
  ASSERT_NE(region, nullptr);
  EXPECT_EQ(region->rows, 3U);
  for (std::uint32_t row = 0; row < 3U; ++row) {
    const Value cell = At(kH, row, kHCol);
    ASSERT_TRUE(cell.is_error()) << row;
    EXPECT_EQ(cell.as_error(), ErrorCode::Div0) << row;
  }
}

TEST_F(WholeAxisShape, GeneratorsUnderRowsAbandonPastTheirElementBound) {
  // A generator, or an operation, sized under ROWS / COLUMNS answers up to
  // 107,374,182 elements and abandons the formula past it, whatever its
  // shape (Mac Excel: ROWS(SEQUENCE(1048576,103)) is 0). Both drivers agree.
  const struct {
    const char* formula;
    double expected;
  } cases[] = {{"=ROWS(SEQUENCE(6553,16384))", 6553.0},
               {"=ROWS(SEQUENCE(6554,16384))", 0.0},
               {"=ROWS(SEQUENCE(13107,8192))", 13107.0},
               {"=ROWS(SEQUENCE(13108,8192))", 0.0},
               {"=ROWS(SEQUENCE(107374,1000))", 107374.0},
               {"=COLUMNS(SEQUENCE(107375,1000))", 0.0},
               {"=ROWS(MAKEARRAY(4096,16384,LAMBDA(i,j,1)))", 4096.0},
               {"=ROWS(MAKEARRAY(8193,16384,LAMBDA(i,j,1)))", 0.0},
               {"=ROWS(RANDARRAY(4096,16384))", 4096.0},
               {"=COLUMNS(RANDARRAY(8193,16384))", 0.0},
               {"=ROWS(SEQUENCE(1048576,102))", static_cast<double>(Sheet::kMaxRows)},
               {"=ROWS(SEQUENCE(1048576,103))", 0.0},
               {"=ROWS(L!A1:A3277+L!A1:XFD1)", 3277.0},
               {"=ROWS(L!A1:A6553+L!A1:XFD1)", 6553.0},
               {"=ROWS(L!A1:A6554+L!A1:XFD1)", 0.0},
               {"=ROWS(SEQUENCE(3277,16384)+0)", 3277.0},
               {"=ROWS(SEQUENCE(6553,16384)+0)", 6553.0},
               {"=ROWS(SEQUENCE(6554,16384)+0)", 0.0}};
  for (const auto& c : cases) {
    Put(kH, 0U, kHCol, c.formula);
    const Value v = At(kH, 0U, kHCol);
    ASSERT_TRUE(v.is_number()) << c.formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), c.expected) << c.formula;
    const Value adhoc = Adhoc(c.formula);
    ASSERT_TRUE(adhoc.is_number()) << c.formula << " " << adhoc.debug_to_string();
    EXPECT_EQ(adhoc.as_number(), c.expected) << c.formula;
  }
}

// The reducers that stream a derived whole column / row through the
// dispatcher (COUNT is lazy and not among them).
const char* const kStreamedReducers[] = {"SUM", "AVERAGE", "MIN", "MAX", "PRODUCT", "COUNTA"};

TEST_F(WholeAxisShape, ReducersStreamADerivedWholeAxisPastTheArrayCeiling) {
  // `-L!A:J` is 10.5M elements, within the element bound, so it is evaluated
  // rather than abandoned (Mac Excel: SUM(-A:J) is #VALUE! over the text in
  // column A, 0 over empty columns). Each reducer streams the repeated tail
  // instead of building it. Both drivers agree.
  constexpr double kCells = 10.0 * static_cast<double>(Sheet::kMaxRows);
  const struct {
    const char* formula;
    bool is_error;
    ErrorCode error;
    double number;
  } cases[] = {
      // Text in column A negates to #VALUE!, the first error in row-major order.
      {"=SUM(-L!A:J)", true, ErrorCode::Value, 0.0},
      {"=AVERAGE(-L!A:J)", true, ErrorCode::Value, 0.0},
      {"=MIN(-L!A:J)", true, ErrorCode::Value, 0.0},
      {"=MAX(-L!A:J)", true, ErrorCode::Value, 0.0},
      {"=PRODUCT(-L!A:J)", true, ErrorCode::Value, 0.0},
      {"=COUNTA(-L!A:J)", false, ErrorCode::Value, kCells},
      // Empty columns: every element is 0.
      {"=SUM(-L!C:L)", false, ErrorCode::Value, 0.0},
      {"=SUM(L!C:L+0)", false, ErrorCode::Value, 0.0},
      {"=AVERAGE(-L!C:L)", false, ErrorCode::Value, 0.0},
      {"=MIN(-L!C:L)", false, ErrorCode::Value, 0.0},
      {"=MAX(L!C:L+1)", false, ErrorCode::Value, 1.0},
      {"=PRODUCT(L!C:L+1)", false, ErrorCode::Value, 1.0},
      {"=COUNTA(-L!C:L)", false, ErrorCode::Value, kCells},
      {"=SUM(L!C:L+1)", false, ErrorCode::Value, kCells},
      // An error tail: the first element already is the error.
      {"=SUM(1/L!C:L)", true, ErrorCode::Div0, 0.0},
      {"=MAX(1/L!C:L)", true, ErrorCode::Div0, 0.0},
      {"=COUNTA(1/L!C:L)", false, ErrorCode::Value, kCells},
  };
  for (const auto& c : cases) {
    Put(kH, 0U, kHCol, c.formula);
    for (const Value& v : {At(kH, 0U, kHCol), Adhoc(c.formula)}) {
      if (c.is_error) {
        ASSERT_TRUE(v.is_error()) << c.formula << " " << v.debug_to_string();
        EXPECT_EQ(v.as_error(), c.error) << c.formula;
      } else {
        ASSERT_TRUE(v.is_number()) << c.formula << " " << v.debug_to_string();
        EXPECT_EQ(v.as_number(), c.number) << c.formula;
      }
    }
  }
}

// `name` applied to the dense walk of a derived whole-axis array: the `head`
// computed cells, then `tail` repeated up to `extent`, through the same
// array-element filter the dispatcher applies before the impl.
Value DenseReduce(const char* name, const Value* head, std::size_t head_count, const Value& tail, std::size_t extent,
                  Arena& arena) {
  const FunctionDef* def = default_registry().lookup(name);
  std::vector<Value> values;
  Value err = Value::blank();
  for (std::size_t i = 0; i < extent; ++i) {
    if (!append_array_sourced_value(*def, i < head_count ? head[i] : tail, &values, &err)) {
      return err;
    }
  }
  return def->impl(values.data(), static_cast<std::uint32_t>(values.size()), arena);
}

TEST_F(WholeAxisShape, StreamedReductionMatchesTheDenseWalkBitForBit) {
  // Fractional data with a cancelling pair, so the order of the additions and
  // SUM's last-step snap both show in the result. A column (tail down the
  // rows) and a row (tail across the columns) are each reduced streamed by
  // the dispatcher and densely through the impl; the doubles must be equal.
  const double data[] = {0.1, 0.2, 1e16, 0.3, -1e16, 0.7, 3.3};
  for (std::uint32_t i = 0; i < 7U; ++i) {
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, i, 3U, Value::number(data[i]))));
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, 7U, 4U + i, Value::number(data[i]))));
  }
  const struct {
    const char* whole;  // streamed
    const char* head;   // the populated head plus one tail cell, dense; the
                        // parentheses keep the root-level snap off
    std::size_t extent;
  } shapes[] = {{"L!D:D*1.1+0.3", "=(L!D1:D8*1.1+0.3)", Sheet::kMaxRows},
                {"L!8:8*1.1+0.3", "=(L!A8:L8*1.1+0.3)", Sheet::kMaxCols}};
  for (const auto& shape : shapes) {
    const Value head = Adhoc(shape.head);
    ASSERT_TRUE(head.is_array()) << shape.head << " " << head.debug_to_string();
    const std::size_t head_count = head.as_array_rows() * head.as_array_cols() - 1U;
    const Value* cells = head.as_array_cells();
    for (const char* reducer : kStreamedReducers) {
      const std::string formula = std::string("=") + reducer + "(" + shape.whole + ")";
      const Value streamed = Adhoc(formula);
      const Value dense = DenseReduce(reducer, cells, head_count, cells[head_count], shape.extent, arena_);
      ASSERT_EQ(streamed.kind(), dense.kind()) << formula;
      if (streamed.is_number()) {
        const double a = streamed.as_number();
        const double b = dense.as_number();
        EXPECT_EQ(std::memcmp(&a, &b, sizeof(double)), 0)
            << formula << " streamed=" << streamed.debug_to_string() << " dense=" << dense.debug_to_string();
      }
    }
  }
}

TEST_F(WholeAxisShape, GeneratorsUnderRowsKeepTheirArgumentErrors) {
  // The arguments are validated as the generator validates them, and
  // MAKEARRAY's lambda is checked for two parameters without being called.
  const struct {
    const char* formula;
    ErrorCode expected;
  } cases[] = {{"=ROWS(MAKEARRAY(5,5,LAMBDA(i,1)))", ErrorCode::Value},
               {"=ROWS(MAKEARRAY(8193,16384,LAMBDA(i,1)))", ErrorCode::Value},
               {"=ROWS(MAKEARRAY(0,5,LAMBDA(i,j,1)))", ErrorCode::Value},
               {"=ROWS(RANDARRAY(5,5,10,1))", ErrorCode::Value},
               {"=ROWS(RANDARRAY(5,5,0.5,3,TRUE))", ErrorCode::Value},
               {"=ROWS(RANDARRAY(-1,5))", ErrorCode::Value}};
  for (const auto& c : cases) {
    const Value v = Adhoc(c.formula);
    ASSERT_TRUE(v.is_error()) << c.formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), c.expected) << c.formula;
  }
  // The lambda is never called: one that would fail for every cell still
  // sizes the grid.
  const Value sized = Adhoc("=ROWS(MAKEARRAY(4096,16384,LAMBDA(i,j,1/0)))");
  ASSERT_TRUE(sized.is_number()) << sized.debug_to_string();
  EXPECT_EQ(sized.as_number(), 4096.0);
}

TEST_F(WholeAxisShape, RowsOfASequenceKeepsItsArgumentErrors) {
  const struct {
    const char* formula;
    ErrorCode expected;
  } cases[] = {{"=ROWS(SEQUENCE(0,16384))", ErrorCode::Calc},
               {"=ROWS(SEQUENCE(-1,16384))", ErrorCode::Value},
               {"=ROWS(SEQUENCE(4096,16384,\"x\"))", ErrorCode::Value},
               {"=ROWS(SEQUENCE(1/0,16384))", ErrorCode::Div0}};
  for (const auto& c : cases) {
    const Value v = Adhoc(c.formula);
    ASSERT_TRUE(v.is_error()) << c.formula << " " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), c.expected) << c.formula;
  }
}

TEST_F(WholeAxisShape, ArrayAtExcelsElementBoundStillEvaluates) {
  // 51 whole columns is 53,477,376 elements, inside the bound.
  Put(kH, 0U, kHCol, "=ROWS(L!A:AY+0)");
  EXPECT_EQ(At(kH, 0U, kHCol).as_number(), static_cast<double>(Sheet::kMaxRows));
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
