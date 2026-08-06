//
// End-to-end tests for the ad-hoc, side-effect-free formula evaluation C
// ABI: `fm_workbook_evaluate_formula` / `fm_workbook_evaluate_cf_formula`.
// Exercises anchoring, cross-sheet resolution, defined names, the
// documented array-to-scalar reduction, the self-reference stale-read
// caveat, Excel CF-predicate coercion, and the read-only purity guarantee.

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

// Builds a single-sheet workbook with A1=10, B1=20 (recalculated).
void seed_ab(fm_workbook_t* wb) {
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 0, 10.0), 0);  // A1
  ASSERT_EQ(fm_workbook_set_number(wb, 0, 0, 1, 20.0), 0);  // B1
  ASSERT_EQ(fm_workbook_recalc(wb), 0);
}

}  // namespace

TEST(EvaluateFormula, ScalarArithmeticAgainstCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  seed_ab(wb.handle);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 2, "=A1+B1", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 30.0);
}

TEST(EvaluateFormula, LeadingEqualsOptional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  seed_ab(wb.handle);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 2, "A1*2", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 20.0);
}

// Correction 4: the anchor reaches ROW()/COLUMN(). ROW()/COLUMN() are
// 1-based; the (row, col) arguments are 0-based.
TEST(EvaluateFormula, RowColumnAnchoredAtTarget) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t r{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 4, 2, "=ROW()", &r), 0);
  EXPECT_EQ(r.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(r.u.number, 5.0);

  fm_value_t c{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 4, 2, "=COLUMN()", &c), 0);
  EXPECT_EQ(c.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(c.u.number, 3.0);
}

TEST(EvaluateFormula, CrossSheetReference) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_add_sheet(wb.handle, "Data"), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 1, 0, 0, 99.0), 0);  // Data!A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=Data!A1", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 99.0);
}

TEST(EvaluateFormula, WorkbookScopedDefinedName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_defined_name(wb.handle, "TAX", "0.08"), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=TAX*100", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 8.0);
}

// Correction 1: an array/spill result is reduced to its top-left element.
// This is the pragmatic API shape, documented as an intentional divergence
// (see divergence.yaml: evaluate_formula_array_scalar_reduction). It is NOT
// implicit intersection and NOT spilling.
TEST(EvaluateFormula, ArrayResultReducesToTopLeft) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=SEQUENCE(3)", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 1.0);
}

// Correction 1: SINGLE()/@ perform true intersection inside the evaluated
// formula, independent of the outer scalar reduction.
TEST(EvaluateFormula, SingleOperatorEvaluates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=SINGLE(SEQUENCE(3))", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 1.0);
}

// Correction 2: a self-reference reads the target cell's cached value
// rather than raising #REF! or engaging iterative calc. A1 holds a literal
// 7; evaluating "=A1" anchored at A1's own address returns 7 (stale read),
// NOT an error. This asserts the documented limitation so a future change
// is caught rather than silently drifting.
TEST(EvaluateFormula, SelfReferenceReadsCachedValue) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 7.0), 0);  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=A1", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 7.0);
}

// The evaluation must not mutate the workbook: neither the referenced cell
// nor the anchor cell changes. This is the entire safety argument for the
// const-workbook purity contract.
TEST(EvaluateFormula, IsSideEffectFree) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 10.0), 0);  // A1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t discard{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 1, "=A1+A1", &discard), 0);

  // A1 is unchanged, and the anchor cell B1 was never written.
  fm_value_t a1{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 0, &a1), 0);
  EXPECT_EQ(a1.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(a1.u.number, 10.0);
  fm_value_t b1{};
  ASSERT_EQ(fm_workbook_get_value(wb.handle, 0, 0, 1, &b1), 0);
  EXPECT_EQ(b1.kind, FM_VAL_BLANK);
}

TEST(EvaluateFormula, NullArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  EXPECT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, nullptr, &v),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=1", nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(EvaluateFormula, SheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  EXPECT_EQ(fm_workbook_evaluate_formula(wb.handle, 99, 0, 0, "=1", &v),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

// ---------------------------------------------------------------------------
// Ad-hoc array evaluation (two-step: evaluate + per-cell readback).
// ---------------------------------------------------------------------------

// A column vector: =SEQUENCE(3) is 3 rows x 1 col, values 1,2,3 row-major.
TEST(EvaluateFormulaArray, ColumnVectorPreservesAllCells) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=SEQUENCE(3)", &rows, &cols), 0);
  EXPECT_EQ(rows, 3U);
  EXPECT_EQ(cols, 1U);

  for (uint32_t i = 0; i < 3U; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, i, &v), 0);
    EXPECT_EQ(v.kind, FM_VAL_NUMBER);
    EXPECT_DOUBLE_EQ(v.u.number, static_cast<double>(i + 1U));
  }
}

// A 2x3 matrix: =SEQUENCE(2,3) fills row-major 1..6.
TEST(EvaluateFormulaArray, MatrixDimensionsAndRowMajorOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=SEQUENCE(2,3)", &rows, &cols), 0);
  EXPECT_EQ(rows, 2U);
  EXPECT_EQ(cols, 3U);

  for (uint32_t i = 0; i < 6U; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, i, &v), 0);
    EXPECT_EQ(v.kind, FM_VAL_NUMBER);
    EXPECT_DOUBLE_EQ(v.u.number, static_cast<double>(i + 1U));
  }
}

// A scalar result is reported as a 1x1 array, index 0 carrying the value.
TEST(EvaluateFormulaArray, ScalarReportedAsOneByOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=1+2", &rows, &cols), 0);
  EXPECT_EQ(rows, 1U);
  EXPECT_EQ(cols, 1U);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 0, &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 3.0);
}

// Text cells survive the arena teardown: the stash owns its own bytes, so
// reading them back after the producing call (whose arena is gone) returns
// is safe. Broadcasting concat over a 1x2 SEQUENCE yields {"1x","2x"}.
TEST(EvaluateFormulaArray, TextCellsStayValidAfterEvaluation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=SEQUENCE(1,2)&\"x\"", &rows, &cols), 0);
  EXPECT_EQ(rows, 1U);
  EXPECT_EQ(cols, 2U);

  fm_value_t a{};
  ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 0, &a), 0);
  ASSERT_EQ(a.kind, FM_VAL_TEXT);
  EXPECT_EQ(std::string(a.u.text), "1x");
  fm_value_t b{};
  ASSERT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 1, &b), 0);
  ASSERT_EQ(b.kind, FM_VAL_TEXT);
  EXPECT_EQ(std::string(b.u.text), "2x");
}

// The legacy scalar API keeps its top-left reduction (backward compat): the
// same array formula that the array API preserves must still degrade here.
TEST(EvaluateFormulaArray, ScalarApiStillReducesToTopLeft) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  ASSERT_EQ(fm_workbook_evaluate_formula(wb.handle, 0, 0, 0, "=SEQUENCE(3)", &v), 0);
  EXPECT_EQ(v.kind, FM_VAL_NUMBER);
  EXPECT_DOUBLE_EQ(v.u.number, 1.0);
}

TEST(EvaluateFormulaArray, IndexOutOfRangeRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  ASSERT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=SEQUENCE(3)", &rows, &cols), 0);

  fm_value_t v{};
  EXPECT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 3, &v),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST(EvaluateFormulaArray, NullArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  EXPECT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, nullptr, &rows, &cols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 0, 0, 0, "=1", nullptr, &cols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));

  fm_value_t v{};
  EXPECT_EQ(fm_workbook_evaluate_formula_array_cell(wb.handle, 0, nullptr),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  EXPECT_EQ(fm_workbook_evaluate_formula_array_cell(nullptr, 0, &v),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST(EvaluateFormulaArray, SheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t rows = 0;
  uint32_t cols = 0;
  EXPECT_EQ(fm_workbook_evaluate_formula_array(wb.handle, 99, 0, 0, "=1", &rows, &cols),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

// ---------------------------------------------------------------------------
// Conditional-formula (CF predicate) entry point.
// ---------------------------------------------------------------------------

// Correction 3: CF result coercion. Error / blank / text / numeric-zero
// yield FALSE (rule does not fire); non-zero number and TRUE booleans yield
// TRUE — the generic evaluator would instead propagate the raw error/value.
TEST(EvaluateCfFormula, ExcelPredicateCoercion) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  auto fires = [&](const char* formula) {
    fm_value_t v{};
    EXPECT_EQ(fm_workbook_evaluate_cf_formula(wb.handle, 0, 0, 0, 0, 0, formula, &v), 0);
    EXPECT_EQ(v.kind, FM_VAL_BOOL);
    return v.u.boolean != 0;
  };

  EXPECT_FALSE(fires("=1/0"));        // error -> FALSE, not propagated
  EXPECT_FALSE(fires("=\"\""));       // empty text -> FALSE
  EXPECT_FALSE(fires("=\"hello\""));  // text -> FALSE
  EXPECT_FALSE(fires("=Z99"));        // blank reference -> FALSE
  EXPECT_FALSE(fires("=0"));          // numeric zero -> FALSE
  EXPECT_TRUE(fires("=5"));           // non-zero number -> TRUE
  EXPECT_TRUE(fires("=3>1"));         // TRUE boolean -> TRUE
}

// Correction 4: relative references in a CF rule are written relative to
// the anchor (top-left of the applied range) and shifted to the target
// cell. Rule "=A1=1" anchored at (0,0): at target A1 it stays A1 (=0, does
// not fire); at target A2 it shifts to A2 (=1, fires).
TEST(EvaluateCfFormula, RelativeRefShiftedFromAnchor) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 0.0), 0);  // A1 = 0
  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 0, 1.0), 0);  // A2 = 1
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);

  fm_value_t at_a1{};
  ASSERT_EQ(fm_workbook_evaluate_cf_formula(wb.handle, 0, 0, 0, 0, 0, "=A1=1", &at_a1), 0);
  EXPECT_EQ(at_a1.u.boolean, 0);  // A1 == 1 ? 0 == 1 -> FALSE

  fm_value_t at_a2{};
  ASSERT_EQ(fm_workbook_evaluate_cf_formula(wb.handle, 0, 1, 0, 0, 0, "=A1=1", &at_a2), 0);
  EXPECT_EQ(at_a2.u.boolean, 1);  // shifted to A2 == 1 -> TRUE
}

TEST(EvaluateCfFormula, NullArgumentsRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_value_t v{};
  EXPECT_EQ(fm_workbook_evaluate_cf_formula(wb.handle, 0, 0, 0, 0, 0, nullptr, &v),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}
