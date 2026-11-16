//
// `Ref` construction, the reference operators and dereference: each read
// agrees with the evaluator's own, every grid class reads as specified, and
// reads are reported once to the right observer.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "eval/reference/deref.h"
#include "eval/reference/ops.h"
#include "eval/reference/ref_value.h"
#include "eval/tail_array.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "test_eval_helpers.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint16_t kHome = RefArea::kHomeSheet;

// One note a dynamic-read observer received.
struct Note {
  std::string sheet;
  DeclaredRect rect;
};

void CollectNote(void* user, std::string_view sheet, const DeclaredRect& rect) {
  static_cast<std::vector<Note>*>(user)->push_back(Note{std::string(sheet), rect});
}

parser::Reference Corner(std::uint32_t row, std::uint32_t col, std::string_view sheet = {}) {
  parser::Reference r{};
  r.row = row;
  r.col = col;
  r.sheet = sheet;
  return r;
}

// Sheet s (0-based), row r, column c (both 0-based) hold s * 100 + (r + 1) * 10 + (c + 1).
double CellValue(std::uint32_t s, std::uint32_t r, std::uint32_t c) {
  return static_cast<double>(s * 100U + (r + 1U) * 10U + (c + 1U));
}

class DerefTest : public ::testing::Test {
 protected:
  DerefTest() : wb_(Workbook::create()) {
    wb_.add_sheet("Sheet2");
    wb_.add_sheet("Sheet3");
    for (std::uint32_t s = 0; s < 3U; ++s) {
      for (std::uint32_t r = 0; r < 5U; ++r) {
        for (std::uint32_t c = 0; c < 5U; ++c) {
          wb_.sheet(s).set_cell_value(r, c, Value::number(CellValue(s, r, c)));
        }
      }
    }
    // D1 holds text and E5 is empty, on Sheet1.
    wb_.sheet(0).set_cell_value(0, 3, Value::text("x"));
    wb_.sheet(0).set_cell_value(4, 4, Value::blank());
  }

  EvalContext Ctx() { return test::workbook_context(wb_, wb_.sheet(0), state_); }

  Value Eval(const char* source) {
    parser::Parser p(source, parse_arena_);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << source;
    if (root == nullptr) {
      return Value::error(ErrorCode::Name);
    }
    return evaluate(*root, arena_, default_registry(), Ctx());
  }

  RefArea Area(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2, std::uint16_t sheet = kHome,
               RefOrigin origin = RefOrigin::kSpelled) {
    return make_ref_area(DeclaredRect{r1, r2, c1, c2, false}, sheet, sheet, origin);
  }

  Value AreaRef(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2, std::uint16_t sheet = kHome,
                RefOrigin origin = RefOrigin::kSpelled) {
    const RefArea a = Area(r1, c1, r2, c2, sheet, origin);
    return make_ref(arena_, RefForm::kArea, 0, &a, 1);
  }

  Value Grid(const std::vector<RefArea>& elements, std::uint32_t rows, std::uint32_t cols) {
    return make_grid_ref(arena_, elements.data(), rows, cols);
  }

  // 2 x 2: a one-cell element (A1 = 11), a multi-cell one (A1:B2), an error
  // element (#REF!) and a one-cell element (B2 = 22).
  Value MixedGrid(std::vector<RefArea>* elements) {
    *elements = {Area(0, 0, 0, 0), Area(0, 0, 1, 1), RefArea::error_element(ErrorCode::Ref), Area(1, 1, 1, 1)};
    return Grid(*elements, 2, 2);
  }

  Value Cells(const Value& ref, ReadCount count, std::vector<Value>* out, bool span = true) {
    auto r = deref_cells(ref, span, arena_, default_registry(), Ctx(), count);
    if (!r) {
      return Value::error(r.error());
    }
    *out = r.value().cells;
    return Value::blank();
  }

  static double Sum(const std::vector<Value>& cells) {
    double total = 0.0;
    for (const Value& v : cells) {
      if (v.is_number()) {
        total += v.as_number();
      }
    }
    return total;
  }

  static void ExpectError(const Value& v, ErrorCode code, const char* what) {
    ASSERT_TRUE(v.is_error()) << what << ": " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), code) << what;
  }

  static void ExpectSame(const Value& got, const Value& want, const char* what) {
    if (!want.is_array()) {
      EXPECT_TRUE(got == want) << what << ": " << got.debug_to_string() << " vs " << want.debug_to_string();
      return;
    }
    ASSERT_TRUE(got.is_array()) << what << ": " << got.debug_to_string();
    ASSERT_EQ(got.as_array_rows(), want.as_array_rows()) << what;
    ASSERT_EQ(got.as_array_cols(), want.as_array_cols()) << what;
    for (std::uint32_t i = 0; i < want.as_array_rows() * want.as_array_cols(); ++i) {
      EXPECT_TRUE(got.as_array_cells()[i] == want.as_array_cells()[i]) << what << " cell " << i;
    }
  }

  Workbook wb_;
  EvalState state_;
  Arena parse_arena_;
  Arena arena_;
};

// ---------------------------------------------------------------------------
// Reads agree with the evaluator
// ---------------------------------------------------------------------------

TEST_F(DerefTest, SingleCellReadsAsTheEvaluatorReadsIt) {
  const Value ref = make_area_ref(Corner(1, 1), Corner(1, 1), RefOrigin::kSpelled, arena_, Ctx());
  ASSERT_TRUE(ref.is_ref());
  EXPECT_EQ(ref.ref_form(), RefForm::kArea);
  const Shaped read = deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  EXPECT_EQ(read.tail_array, nullptr);
  ExpectSame(read.value, Eval("B2"), "B2");
  ExpectSame(deref_top_left(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted), Eval("B2"), "top-left");
}

TEST_F(DerefTest, RectangleReadsAsTheEvaluatorReadsIt) {
  const Value ref = make_area_ref(Corner(0, 0), Corner(2, 1), RefOrigin::kSpelled, arena_, Ctx());
  const Shaped read = deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ExpectSame(read.value, Eval("A1:B3"), "A1:B3");
  auto cells = deref_cells(ref, false, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(cells);
  EXPECT_EQ(cells.value().rows, 3U);
  EXPECT_EQ(cells.value().cols, 2U);
  ASSERT_EQ(cells.value().cells.size(), 6U);
  for (std::uint32_t i = 0; i < 6U; ++i) {
    EXPECT_TRUE(cells.value().cells[i] == read.value.as_array_cells()[i]) << i;
  }
}

TEST_F(DerefTest, CrossSheetRectangleReadsAsTheEvaluatorReadsIt) {
  const Value ref = make_area_ref(Corner(0, 0, "Sheet2"), Corner(1, 1), RefOrigin::kSpelled, arena_, Ctx());
  ASSERT_TRUE(ref.is_ref());
  EXPECT_EQ(ref.as_ref().areas[0].sheet_first, 1U);
  const Shaped read = deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ExpectSame(read.value, Eval("Sheet2!A1:B2"), "Sheet2!A1:B2");
  ExpectError(make_area_ref(Corner(0, 0, "Nope"), Corner(1, 1), RefOrigin::kSpelled, arena_, Ctx()), ErrorCode::Ref,
              "unknown sheet");
  ExpectError(make_area_ref(Corner(0, 0, "Sheet2"), Corner(1, 1, "Sheet3"), RefOrigin::kSpelled, arena_, Ctx()),
              ErrorCode::Ref, "two sheets");
}

TEST_F(DerefTest, WholeColumnAndRowReadAtDeclaredSize) {
  parser::Reference col = Corner(0, 0);
  col.is_full_col = true;
  const Value col_ref = make_area_ref(col, col, RefOrigin::kSpelled, arena_, Ctx());
  EXPECT_NE(col_ref.as_ref().areas[0].whole_col, 0U);
  const Shaped col_read = deref_shaped(col_ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_NE(col_read.tail_array, nullptr);
  EXPECT_EQ(col_read.tail_array->rows, Sheet::kMaxRows);
  EXPECT_EQ(col_read.tail_array->cols, 1U);
  std::vector<Value> cells;
  ASSERT_TRUE(Cells(col_ref, ReadCount::kCounted, &cells).is_blank());
  const Value want_col = Eval("SUM(A:A)");
  ASSERT_TRUE(want_col.is_number());
  EXPECT_EQ(Sum(cells), want_col.as_number());

  parser::Reference row = Corner(0, 0);
  row.is_full_row = true;
  const Value row_ref = make_area_ref(row, row, RefOrigin::kSpelled, arena_, Ctx());
  EXPECT_NE(row_ref.as_ref().areas[0].whole_row, 0U);
  const Shaped row_read = deref_shaped(row_ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_NE(row_read.tail_array, nullptr);
  EXPECT_EQ(row_read.tail_array->cols, Sheet::kMaxCols);
  ASSERT_TRUE(Cells(row_ref, ReadCount::kCounted, &cells).is_blank());
  const Value want_row = Eval("SUM(1:1)");
  ASSERT_TRUE(want_row.is_number());
  EXPECT_EQ(Sum(cells), want_row.as_number());
}

TEST_F(DerefTest, SpanReadsEverySheetOnlyWhereAccepted) {
  const Value ref = make_span3d_ref("Sheet1", "Sheet3", Corner(0, 0), Corner(1, 1), RefOrigin::kSpelled, arena_, Ctx());
  ASSERT_TRUE(ref.is_ref());
  EXPECT_EQ(ref.ref_form(), RefForm::kSpan3D);
  std::vector<Value> cells;
  ASSERT_TRUE(Cells(ref, ReadCount::kCounted, &cells).is_blank());
  EXPECT_EQ(cells.size(), 12U);
  const Value want = Eval("SUM(Sheet1:Sheet3!A1:B2)");
  ASSERT_TRUE(want.is_number());
  EXPECT_EQ(Sum(cells), want.as_number());
  ExpectError(Cells(ref, ReadCount::kCounted, &cells, /*span=*/false), ErrorCode::Ref, "span refused");
  const Shaped value_read = deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ExpectSame(value_read.value, Eval("Sheet1:Sheet3!A1:B2"), "span as value");
  ExpectError(make_span3d_ref("Sheet1", "Missing", Corner(0, 0), Corner(0, 0), RefOrigin::kSpelled, arena_, Ctx()),
              ErrorCode::Name, "missing end sheet");
  ExpectError(make_span3d_ref("Missing", "Sheet2", Corner(0, 0), Corner(0, 0), RefOrigin::kSpelled, arena_, Ctx()),
              ErrorCode::Ref, "missing begin sheet");
}

TEST_F(DerefTest, UnionReadsEveryArea) {
  const Value ref = ref_union(AreaRef(0, 0, 1, 0), AreaRef(0, 2, 1, 2), arena_, Ctx());
  ASSERT_TRUE(ref.is_ref());
  EXPECT_EQ(ref.ref_form(), RefForm::kUnion);
  std::vector<Value> cells;
  ASSERT_TRUE(Cells(ref, ReadCount::kCounted, &cells).is_blank());
  EXPECT_EQ(cells.size(), 4U);
  const Value want = Eval("SUM((A1:A2,C1:C2))");
  ASSERT_TRUE(want.is_number());
  EXPECT_EQ(Sum(cells), want.as_number());
  ExpectError(deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted).value, ErrorCode::Value,
              "union as value");
}

TEST_F(DerefTest, DynamicMarkLeavesTheFormIntact) {
  const Value plain = AreaRef(0, 0, 1, 1);
  const Value marked = plain.with_ref_dynamic();
  EXPECT_FALSE(plain.ref_is_dynamic());
  EXPECT_TRUE(marked.ref_is_dynamic());
  EXPECT_EQ(marked.ref_form(), RefForm::kArea);
  EXPECT_EQ(marked.as_ref().areas[0].row_last, 1U);
  EXPECT_FALSE(Value::number(1.0).with_ref_dynamic().ref_is_dynamic());
  const Value external = [&] {
    const RefArea a = Area(0, 0, 0, 0, 0);
    return make_ref(arena_, RefForm::kArea, 1, &a, 1).with_ref_dynamic();
  }();
  EXPECT_TRUE(external.ref_is_dynamic());
  EXPECT_EQ(external.ref_form(), RefForm::kArea);
  EXPECT_EQ(external.as_ref().book, 1U);
}

// ---------------------------------------------------------------------------
// Operators
// ---------------------------------------------------------------------------

TEST_F(DerefTest, EmptyIntersectionIsNull) {
  ExpectError(ref_intersect(AreaRef(0, 0, 1, 0), AreaRef(0, 2, 1, 2), arena_, Ctx()), ErrorCode::Null, "disjoint");
}

TEST_F(DerefTest, IntersectionIsTheOverlap) {
  const Value r = ref_intersect(AreaRef(0, 0, 2, 2), AreaRef(1, 1, 3, 3), arena_, Ctx());
  ASSERT_TRUE(r.is_ref());
  ASSERT_EQ(r.ref_form(), RefForm::kArea);
  const RefArea& a = r.as_ref().areas[0];
  EXPECT_EQ(a.row_first, 1U);
  EXPECT_EQ(a.row_last, 2U);
  EXPECT_EQ(a.col_first, 1U);
  EXPECT_EQ(a.col_last, 2U);
}

TEST_F(DerefTest, UnionOperandIntersectsAreaByArea) {
  const Value u = ref_union(AreaRef(0, 0, 4, 0), AreaRef(0, 2, 4, 2), arena_, Ctx());
  const Value r = ref_intersect(u, AreaRef(1, 0, 2, 2), arena_, Ctx());
  ASSERT_TRUE(r.is_ref());
  ASSERT_EQ(r.ref_form(), RefForm::kUnion);
  ASSERT_EQ(r.as_ref().area_count, 2U);
  EXPECT_EQ(r.as_ref().areas[0].col_first, 0U);
  EXPECT_EQ(r.as_ref().areas[1].col_first, 2U);
  EXPECT_EQ(r.as_ref().areas[1].row_first, 1U);
  EXPECT_EQ(r.as_ref().areas[1].row_last, 2U);
}

TEST_F(DerefTest, CrossSheetOperandsFail) {
  ExpectError(ref_intersect(AreaRef(0, 0, 1, 1, 0), AreaRef(0, 0, 1, 1, 1), arena_, Ctx()), ErrorCode::Ref,
              "intersect across sheets");
  ExpectError(ref_range(AreaRef(0, 0, 0, 0, 0), AreaRef(1, 1, 1, 1, 1), arena_, Ctx()), ErrorCode::Value,
              "range across sheets");
}

TEST_F(DerefTest, RangeIsTheBoundingBoxOfAChain) {
  const Value ab = ref_range(AreaRef(0, 0, 0, 0), AreaRef(2, 2, 2, 2), arena_, Ctx());
  const Value abc = ref_range(ab, AreaRef(4, 4, 4, 4), arena_, Ctx());
  ASSERT_TRUE(abc.is_ref());
  const RefArea& a = abc.as_ref().areas[0];
  EXPECT_EQ(a.row_first, 0U);
  EXPECT_EQ(a.row_last, 4U);
  EXPECT_EQ(a.col_first, 0U);
  EXPECT_EQ(a.col_last, 4U);
  const Value reversed = ref_range(AreaRef(4, 4, 4, 4), AreaRef(0, 0, 0, 0), arena_, Ctx());
  EXPECT_EQ(reversed.as_ref().areas[0].row_first, 0U);
  EXPECT_EQ(reversed.as_ref().areas[0].row_last, 4U);
}

TEST_F(DerefTest, RangeOverWholeColumnsKeepsTheAxisMark) {
  parser::Reference a = Corner(0, 0);
  a.is_full_col = true;
  parser::Reference c = Corner(0, 2);
  c.is_full_col = true;
  const Value r = ref_range(make_area_ref(a, a, RefOrigin::kSpelled, arena_, Ctx()),
                            make_area_ref(c, c, RefOrigin::kSpelled, arena_, Ctx()), arena_, Ctx());
  ASSERT_TRUE(r.is_ref());
  const RefArea& box = r.as_ref().areas[0];
  EXPECT_NE(box.whole_col, 0U);
  EXPECT_EQ(box.whole_row, 0U);
  EXPECT_EQ(box.col_first, 0U);
  EXPECT_EQ(box.col_last, 2U);
}

TEST_F(DerefTest, RangeRejectsNonAreas) {
  const Value u = ref_union(AreaRef(0, 0, 0, 0), AreaRef(2, 2, 2, 2), arena_, Ctx());
  ExpectError(ref_range(u, AreaRef(0, 0, 0, 0), arena_, Ctx()), ErrorCode::Value, "union operand");
  ExpectError(ref_range(Value::number(1.0), AreaRef(0, 0, 0, 0), arena_, Ctx()), ErrorCode::Value, "number operand");
  ExpectError(ref_range(Value::error(ErrorCode::Div0), AreaRef(0, 0, 0, 0), arena_, Ctx()), ErrorCode::Div0,
              "error operand");
}

TEST_F(DerefTest, OriginMergeFollowsTheQualifier) {
  // A bare operand takes a spelled sheet.
  const Value spelled = ref_range(AreaRef(0, 0, 0, 0, kHome, RefOrigin::kImplicit),
                                  AreaRef(1, 1, 1, 1, 1, RefOrigin::kSpelled), arena_, Ctx());
  ASSERT_TRUE(spelled.is_ref());
  EXPECT_EQ(spelled.as_ref().areas[0].sheet_first, 1U);
  EXPECT_EQ(spelled.as_ref().areas[0].origin, static_cast<std::uint16_t>(RefOrigin::kSpelled));
  // A derived sheet does not reach a bare operand unless it is the own sheet.
  ExpectError(ref_range(AreaRef(0, 0, 0, 0, kHome, RefOrigin::kImplicit), AreaRef(1, 1, 1, 1, 1, RefOrigin::kDerived),
                        arena_, Ctx()),
              ErrorCode::Value, "derived foreign sheet");
  const Value own = ref_range(AreaRef(0, 0, 0, 0, kHome, RefOrigin::kImplicit),
                              AreaRef(1, 1, 1, 1, 0, RefOrigin::kDerived), arena_, Ctx());
  ASSERT_TRUE(own.is_ref());
  EXPECT_EQ(own.as_ref().areas[0].sheet_first, 0U);
  EXPECT_EQ(own.as_ref().areas[0].origin, static_cast<std::uint16_t>(RefOrigin::kDerived));
  // Two bare operands stay bare.
  const Value bare = ref_range(AreaRef(0, 0, 0, 0, kHome, RefOrigin::kImplicit),
                               AreaRef(1, 1, 1, 1, kHome, RefOrigin::kImplicit), arena_, Ctx());
  EXPECT_EQ(bare.as_ref().areas[0].sheet_first, kHome);
  EXPECT_EQ(bare.as_ref().areas[0].origin, static_cast<std::uint16_t>(RefOrigin::kImplicit));
}

TEST_F(DerefTest, UnionRules) {
  const Value ab = ref_union(AreaRef(0, 0, 0, 0), AreaRef(2, 2, 2, 2), arena_, Ctx());
  const Value abc = ref_union(ab, AreaRef(4, 4, 4, 4), arena_, Ctx());
  ASSERT_TRUE(abc.is_ref());
  ASSERT_EQ(abc.ref_form(), RefForm::kUnion);
  ASSERT_EQ(abc.as_ref().area_count, 3U);
  EXPECT_EQ(abc.as_ref().areas[0].row_first, 0U);
  EXPECT_EQ(abc.as_ref().areas[2].row_first, 4U);

  const Value span =
      make_span3d_ref("Sheet1", "Sheet2", Corner(0, 0), Corner(0, 0), RefOrigin::kSpelled, arena_, Ctx());
  ExpectError(ref_union(span, AreaRef(0, 0, 0, 0), arena_, Ctx()), ErrorCode::Name, "3-D operand");
  ExpectError(ref_union(AreaRef(0, 0, 0, 0), span, arena_, Ctx()), ErrorCode::Name, "3-D operand on the right");

  const std::vector<RefArea> elements = {Area(0, 0, 0, 0)};
  const Value grid = Grid(elements, 1, 1);
  ExpectError(ref_union(grid, AreaRef(0, 0, 0, 0), arena_, Ctx()), ErrorCode::Value, "grid operand");

  ExpectError(ref_union(AreaRef(0, 0, 0, 0, 0), AreaRef(0, 0, 0, 0, 1), arena_, Ctx()), ErrorCode::Value,
              "two spelled sheets");
  const Value mixed = ref_union(AreaRef(0, 0, 0, 0, 0), AreaRef(0, 0, 0, 0, 1, RefOrigin::kDerived), arena_, Ctx());
  EXPECT_TRUE(mixed.is_ref());
  const Value bare = ref_union(AreaRef(0, 0, 0, 0, kHome, RefOrigin::kImplicit), AreaRef(0, 0, 0, 0, 1), arena_, Ctx());
  EXPECT_TRUE(bare.is_ref());
}

TEST_F(DerefTest, OperatorsPropagateTheDynamicMark) {
  const Value plain = AreaRef(0, 0, 1, 1);
  const Value other = AreaRef(1, 1, 2, 2);
  const Value marked = AreaRef(1, 1, 2, 2).with_ref_dynamic();
  EXPECT_FALSE(ref_range(plain, other, arena_, Ctx()).ref_is_dynamic());
  EXPECT_FALSE(ref_intersect(plain, other, arena_, Ctx()).ref_is_dynamic());
  EXPECT_FALSE(ref_union(plain, other, arena_, Ctx()).ref_is_dynamic());
  EXPECT_TRUE(ref_range(plain, marked, arena_, Ctx()).ref_is_dynamic());
  EXPECT_TRUE(ref_range(marked, plain, arena_, Ctx()).ref_is_dynamic());
  EXPECT_TRUE(ref_intersect(plain, marked, arena_, Ctx()).ref_is_dynamic());
  EXPECT_TRUE(ref_union(marked, plain, arena_, Ctx()).ref_is_dynamic());
}

// ---------------------------------------------------------------------------
// Grid classes
// ---------------------------------------------------------------------------

TEST_F(DerefTest, ElementErrorClass) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  const Value v =
      materialise_grid(grid, GridArg::kElementError, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(v.is_array());
  ASSERT_EQ(v.as_array_rows(), 2U);
  ASSERT_EQ(v.as_array_cols(), 2U);
  const Value* c = v.as_array_cells();
  ExpectError(c[0], ErrorCode::Value, "one-cell");
  ExpectError(c[1], ErrorCode::Value, "multi-cell");
  ExpectError(c[2], ErrorCode::Ref, "error element");
  ExpectError(c[3], ErrorCode::Value, "one-cell");
  // A grid read as a value is the same array.
  const Shaped shaped = deref_shaped(grid, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ExpectSame(shaped.value, v, "default dereference");
}

TEST_F(DerefTest, PerValueClass) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  const Value v =
      materialise_grid(grid, GridArg::kPerValue, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  ASSERT_TRUE(c[0].is_number());
  EXPECT_EQ(c[0].as_number(), CellValue(0, 0, 0));
  ExpectError(c[1], ErrorCode::Value, "multi-cell");
  ExpectError(c[2], ErrorCode::Ref, "error element");
  ASSERT_TRUE(c[3].is_number());
  EXPECT_EQ(c[3].as_number(), CellValue(0, 1, 1));
}

TEST_F(DerefTest, TopLeftClass) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  // Unfiltered: every element's top-left cell, an error element as its error.
  const Value raw =
      materialise_grid(grid, GridArg::kTopLeft, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(raw.is_array());
  ASSERT_EQ(raw.as_array_rows(), 4U);
  ASSERT_EQ(raw.as_array_cols(), 1U);
  EXPECT_EQ(raw.as_array_cells()[0].as_number(), CellValue(0, 0, 0));
  EXPECT_EQ(raw.as_array_cells()[1].as_number(), CellValue(0, 0, 0));
  ExpectError(raw.as_array_cells()[2], ErrorCode::Ref, "error element");
  EXPECT_EQ(raw.as_array_cells()[3].as_number(), CellValue(0, 1, 1));

  // Through SUM's range-sourced filter an error element is the whole result.
  const FunctionDef* sum = default_registry().lookup("SUM");
  ASSERT_NE(sum, nullptr);
  ExpectError(materialise_grid(grid, GridArg::kTopLeft, sum, arena_, default_registry(), Ctx(), ReadCount::kCounted),
              ErrorCode::Ref, "filtered error element");

  // Text and blank cells are skipped as for a range.
  const std::vector<RefArea> cells = {Area(0, 0, 0, 0), Area(0, 3, 0, 3), Area(4, 4, 4, 4), Area(1, 1, 1, 1)};
  const Value filtered = materialise_grid(Grid(cells, 1, 4), GridArg::kTopLeft, sum, arena_, default_registry(), Ctx(),
                                          ReadCount::kCounted);
  ASSERT_TRUE(filtered.is_array());
  ASSERT_EQ(filtered.as_array_rows(), 2U);
  EXPECT_EQ(filtered.as_array_cells()[0].as_number(), CellValue(0, 0, 0));
  EXPECT_EQ(filtered.as_array_cells()[1].as_number(), CellValue(0, 1, 1));

  // Nothing kept: a single blank.
  const std::vector<RefArea> only_text = {Area(0, 3, 0, 3)};
  const Value none = materialise_grid(Grid(only_text, 1, 1), GridArg::kTopLeft, sum, arena_, default_registry(), Ctx(),
                                      ReadCount::kCounted);
  ASSERT_TRUE(none.is_array());
  EXPECT_EQ(none.as_array_rows(), 1U);
  EXPECT_TRUE(none.as_array_cells()[0].is_blank());
}

TEST_F(DerefTest, OpaqueClass) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  const Value v =
      materialise_grid(grid, GridArg::kOpaque, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(v.is_array());
  const Value* c = v.as_array_cells();
  for (const std::size_t i : {0U, 1U, 3U}) {
    EXPECT_TRUE(c[i].is_opaque_element()) << i;
  }
  ExpectError(c[2], ErrorCode::Ref, "error element");
}

TEST_F(DerefTest, WholeResultClasses) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  ExpectError(
      materialise_grid(grid, GridArg::kWholeValue, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted),
      ErrorCode::Value, "whole #VALUE!");
  ExpectError(
      materialise_grid(grid, GridArg::kWholeNA, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted),
      ErrorCode::NA, "whole #N/A");
}

TEST_F(DerefTest, DispatcherClassesAndNonGridsPassThrough) {
  std::vector<RefArea> elements;
  const Value grid = MixedGrid(&elements);
  for (const GridArg cls : {GridArg::kPerRef, GridArg::kKeepRef, GridArg::kReshape, GridArg::kLambdaElement}) {
    const Value v = materialise_grid(grid, cls, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
    ASSERT_TRUE(v.is_ref());
    EXPECT_EQ(v.ref_form(), RefForm::kGrid);
    EXPECT_EQ(v.as_ref().rows, 2U);
    EXPECT_EQ(v.as_ref().cols, 2U);
  }
  const Value area = AreaRef(0, 0, 1, 1);
  const Value v =
      materialise_grid(area, GridArg::kElementError, nullptr, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  ASSERT_TRUE(v.is_ref());
  EXPECT_EQ(v.ref_form(), RefForm::kArea);
}

// ---------------------------------------------------------------------------
// Read reporting
// ---------------------------------------------------------------------------

TEST_F(DerefTest, CountedReadsReachTheObserverAndUncountedOnesDoNot) {
  state_.observe_reads_of(&wb_.sheet(0), 0, 0, 7U);
  const Value ref = AreaRef(0, 0, 1, 1);
  std::vector<Value> cells;
  ASSERT_TRUE(Cells(ref, ReadCount::kUncounted, &cells).is_blank());
  EXPECT_TRUE(state_.observed_reads().empty());
  ASSERT_TRUE(Cells(ref, ReadCount::kCounted, &cells).is_blank());
  ASSERT_EQ(state_.observed_reads().size(), 1U);
  EXPECT_EQ(state_.observed_reads()[0], 7U);

  state_.clear_observed_reads();
  deref_top_left(ref, arena_, default_registry(), Ctx(), ReadCount::kUncounted);
  deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kUncounted);
  EXPECT_TRUE(state_.observed_reads().empty());
  deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted);
  EXPECT_EQ(state_.observed_reads().size(), 1U);
}

TEST_F(DerefTest, OnlyDynamicReferencesReportTheirRectangles) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const Value plain = AreaRef(0, 0, 1, 1);
  deref_cells(plain, false, arena_, default_registry(), ctx, ReadCount::kCounted);
  deref_shaped(plain, arena_, default_registry(), ctx, ReadCount::kCounted);
  EXPECT_TRUE(notes.empty());

  const Value dynamic = plain.with_ref_dynamic();
  deref_cells(dynamic, false, arena_, default_registry(), ctx, ReadCount::kCounted);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].sheet, "");
  EXPECT_EQ(notes[0].rect.row_first, 0U);
  EXPECT_EQ(notes[0].rect.row_last, 1U);
  EXPECT_EQ(notes[0].rect.col_first, 0U);
  EXPECT_EQ(notes[0].rect.col_last, 1U);

  // Uncounted reads are reported to the graph all the same.
  notes.clear();
  deref_cells(dynamic, false, arena_, default_registry(), ctx, ReadCount::kUncounted);
  EXPECT_EQ(notes.size(), 1U);

  // A qualified sheet is reported by name.
  notes.clear();
  deref_cells(AreaRef(0, 0, 0, 0, 1).with_ref_dynamic(), false, arena_, default_registry(), ctx, ReadCount::kCounted);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].sheet, "Sheet2");
}

TEST_F(DerefTest, ElementReadsAreCoalescedPerRectangle) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  // Two elements on one rectangle, one on another, and a multi-cell element
  // that reads only its top-left cell, which is the first rectangle again.
  const std::vector<RefArea> elements = {Area(0, 0, 0, 0), Area(1, 1, 1, 1), Area(0, 0, 0, 0), Area(0, 0, 1, 1)};
  const Value grid = Grid(elements, 2, 2).with_ref_dynamic();
  const Value v =
      materialise_grid(grid, GridArg::kTopLeft, nullptr, arena_, default_registry(), ctx, ReadCount::kCounted);
  ASSERT_TRUE(v.is_array());
  EXPECT_EQ(v.as_array_rows(), 4U);
  ASSERT_EQ(notes.size(), 2U);
  EXPECT_NE(notes[0].rect.row_first, notes[1].rect.row_first);
  for (const Note& n : notes) {
    EXPECT_TRUE(n.rect.single_cell());
  }

  // Classes that read no cell report nothing.
  notes.clear();
  materialise_grid(grid, GridArg::kElementError, nullptr, arena_, default_registry(), ctx, ReadCount::kCounted);
  materialise_grid(grid, GridArg::kOpaque, nullptr, arena_, default_registry(), ctx, ReadCount::kCounted);
  EXPECT_TRUE(notes.empty());

  // PV reads only its one-cell elements: A1 twice and B2 once.
  materialise_grid(grid, GridArg::kPerValue, nullptr, arena_, default_registry(), ctx, ReadCount::kUncounted);
  EXPECT_EQ(notes.size(), 2U);
}

TEST_F(DerefTest, UnionAreasOnOneRectangleAreReportedOnce) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const Value u = ref_union(AreaRef(0, 0, 1, 1), AreaRef(0, 0, 1, 1), arena_, ctx).with_ref_dynamic();
  auto r = deref_cells(u, false, arena_, default_registry(), ctx, ReadCount::kCounted);
  ASSERT_TRUE(r);
  EXPECT_EQ(r.value().cells.size(), 8U);
  EXPECT_EQ(notes.size(), 1U);
}

TEST_F(DerefTest, SpanReportsOneRectanglePerSheet) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const Value span = make_span3d_ref("Sheet1", "Sheet3", Corner(0, 0), Corner(0, 0), RefOrigin::kSpelled, arena_, Ctx())
                         .with_ref_dynamic();
  auto r = deref_cells(span, true, arena_, default_registry(), ctx, ReadCount::kCounted);
  ASSERT_TRUE(r);
  EXPECT_EQ(notes.size(), 3U);
}

// ---------------------------------------------------------------------------
// The walker: references as values, and a `Ref` read in value mode
// ---------------------------------------------------------------------------

class WalkerRefTest : public DerefTest {
 protected:
  const parser::AstNode& Parse(const char* source) {
    parser::Parser p(source, parse_arena_);
    parser::AstNode* root = p.parse();
    EXPECT_NE(root, nullptr) << source;
    return *root;
  }

  Value AsRef(const char* source) { return eval_node_ref(Parse(source), arena_, default_registry(), Ctx()); }

  Shaped ValueMode(const parser::AstNode& node) { return eval_node_shaped(node, arena_, default_registry(), Ctx()); }
};

TEST_F(WalkerRefTest, TokensBecomeReferencesThatReadAsTheEvaluatorReads) {
  struct Case {
    const char* source;
    RefForm form;
    bool compare;  // a whole axis is compared by its form only
  };
  const Case cases[] = {
      {"=A1", RefForm::kArea, true},           {"=B2:C3", RefForm::kArea, true},
      {"=Sheet2!A1:B2", RefForm::kArea, true}, {"=B:B", RefForm::kArea, false},
      {"=2:2", RefForm::kArea, false},         {"=Sheet2:Sheet3!A1:B2", RefForm::kSpan3D, false},
  };
  for (const Case& c : cases) {
    const Value ref = AsRef(c.source);
    ASSERT_TRUE(ref.is_ref()) << c.source << ": " << ref.debug_to_string();
    EXPECT_EQ(ref.ref_form(), c.form) << c.source;
    EXPECT_FALSE(ref.ref_is_dynamic()) << c.source;
    if (c.compare) {
      ExpectSame(densify(deref_shaped(ref, arena_, default_registry(), Ctx(), ReadCount::kCounted), arena_),
                 Eval(c.source), c.source);
    }
  }
}

TEST_F(WalkerRefTest, QualifierSetsTheOrigin) {
  EXPECT_EQ(static_cast<RefOrigin>(AsRef("=A1:B2").as_ref().areas[0].origin), RefOrigin::kImplicit);
  EXPECT_EQ(static_cast<RefOrigin>(AsRef("=Sheet2!A1:B2").as_ref().areas[0].origin), RefOrigin::kSpelled);
}

TEST_F(WalkerRefTest, BadQualifierIsTheEvaluatorsError) {
  ExpectError(AsRef("=Nope!A1"), ErrorCode::Ref, "unknown sheet");
}

TEST_F(WalkerRefTest, OtherNodesAreTheirValues) {
  ExpectSame(AsRef("=1+2"), Value::number(3.0), "operator");
  ExpectSame(AsRef("=SUM(A1:B2)"), Eval("=SUM(A1:B2)"), "call returning a value");
  ExpectSame(AsRef("=A1:B2+0"), Eval("=A1:B2+0"), "array");
  ExpectSame(AsRef("=\"t\""), Value::text("t"), "text literal");
}

TEST_F(WalkerRefTest, LiteralReferenceIsDereferencedInValueMode) {
  const parser::AstNode* area = parser::make_literal(parse_arena_, AreaRef(0, 0, 1, 1));
  ASSERT_NE(area, nullptr);
  ExpectSame(densify(ValueMode(*area), arena_), Eval("=A1:B2"), "rectangle");
  const Value kept = eval_node_ref(*area, arena_, default_registry(), Ctx());
  ASSERT_TRUE(kept.is_ref());
  EXPECT_EQ(kept.ref_form(), RefForm::kArea);

  const parser::AstNode* cell = parser::make_literal(parse_arena_, AreaRef(1, 1, 1, 1));
  ExpectSame(ValueMode(*cell).value, Value::number(CellValue(0, 1, 1)), "one cell");
}

TEST_F(WalkerRefTest, LiteralGridAndUnionFollowTheDefaultRule) {
  std::vector<RefArea> elements;
  const parser::AstNode* grid = parser::make_literal(parse_arena_, MixedGrid(&elements));
  const Value g = densify(ValueMode(*grid), arena_);
  ASSERT_TRUE(g.is_array()) << g.debug_to_string();
  ASSERT_EQ(g.as_array_rows(), 2U);
  ASSERT_EQ(g.as_array_cols(), 2U);
  ExpectError(g.as_array_cells()[0], ErrorCode::Value, "one-cell element");
  ExpectError(g.as_array_cells()[1], ErrorCode::Value, "multi-cell element");
  ExpectError(g.as_array_cells()[2], ErrorCode::Ref, "error element");
  ExpectError(g.as_array_cells()[3], ErrorCode::Value, "one-cell element");

  const RefArea areas[] = {Area(0, 0, 0, 0), Area(2, 2, 2, 2)};
  const parser::AstNode* uni = parser::make_literal(parse_arena_, make_union_ref(arena_, areas, 2));
  ExpectError(ValueMode(*uni).value, ErrorCode::Value, "union");
}

TEST_F(WalkerRefTest, DereferencedLiteralCountsItsReads) {
  std::vector<Note> notes;
  const parser::AstNode* dyn = parser::make_literal(parse_arena_, AreaRef(0, 0, 1, 1).with_ref_dynamic());
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  (void)eval_node_shaped(*dyn, arena_, default_registry(), ctx);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].rect.row_last, 1U);
  notes.clear();
  (void)eval_node_ref(*dyn, arena_, default_registry(), ctx);
  EXPECT_TRUE(notes.empty());
}

// OFFSET and INDIRECT build their reference: a `kDerived` area marked dynamic.
class BuiltRefTest : public WalkerRefTest {
 protected:
  void ExpectBuilt(const char* source, std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2,
                   std::uint16_t sheet = kHome) {
    const Value ref = AsRef(source);
    ASSERT_TRUE(ref.is_ref()) << source << ": " << ref.debug_to_string();
    EXPECT_EQ(ref.ref_form(), RefForm::kArea) << source;
    EXPECT_TRUE(ref.ref_is_dynamic()) << source;
    const RefArea& a = ref.as_ref().areas[0];
    EXPECT_EQ(static_cast<RefOrigin>(a.origin), RefOrigin::kDerived) << source;
    EXPECT_EQ(a.sheet_first, sheet) << source;
    EXPECT_EQ(a.row_first, r1) << source;
    EXPECT_EQ(a.col_first, c1) << source;
    EXPECT_EQ(a.row_last, r2) << source;
    EXPECT_EQ(a.col_last, c2) << source;
  }
};

TEST_F(BuiltRefTest, OffsetAndIndirectReturnDynamicReferences) {
  ExpectBuilt("=OFFSET(A1,1,1,2,2)", 1, 1, 2, 2);
  ExpectBuilt("=OFFSET(B2:C3,-1,-1)", 0, 0, 1, 1);
  ExpectBuilt("=OFFSET(Sheet2!A1,0,0)", 0, 0, 0, 0, 1);
  ExpectBuilt("=OFFSET(OFFSET(A1,1,0),1,0)", 2, 0, 2, 0);
  ExpectBuilt("=INDIRECT(\"B2:C3\")", 1, 1, 2, 2);
  ExpectBuilt("=INDIRECT(\"Sheet3!C4\")", 3, 2, 3, 2, 2);
  ExpectBuilt("=INDIRECT(\"R2C2\",FALSE)", 1, 1, 1, 1);
}

TEST_F(BuiltRefTest, WholeAxisResultsKeepTheAxisMark) {
  const Value col = AsRef("=INDIRECT(\"B:B\")");
  ASSERT_TRUE(col.is_ref());
  EXPECT_EQ(col.as_ref().areas[0].whole_col, 1U);
  const Value row = AsRef("=INDIRECT(\"2:3\")");
  ASSERT_TRUE(row.is_ref());
  EXPECT_EQ(row.as_ref().areas[0].whole_row, 1U);
  const Value shifted = AsRef("=OFFSET(A:A,0,1)");
  ASSERT_TRUE(shifted.is_ref());
  EXPECT_EQ(shifted.as_ref().areas[0].whole_col, 1U);
  EXPECT_EQ(shifted.as_ref().areas[0].col_first, 1U);
}

TEST_F(BuiltRefTest, FailuresAreErrors) {
  ExpectError(AsRef("=INDIRECT(\"Nope!A1\")"), ErrorCode::Ref, "unknown sheet");
  ExpectError(AsRef("=INDIRECT(\"\")"), ErrorCode::Ref, "empty text");
  ExpectError(AsRef("=OFFSET(A1,-1,0)"), ErrorCode::Ref, "off the grid");
  ExpectError(AsRef("=OFFSET(A1,0,0,0)"), ErrorCode::Ref, "zero height");
  ExpectError(AsRef("=OFFSET(A1,1/0,0)"), ErrorCode::Div0, "argument error");
}

TEST_F(BuiltRefTest, ValueModeReadsTheReference) {
  ExpectSame(Eval("=OFFSET(A1,1,1,2,2)"), Eval("=B2:C3"), "rectangle");
  ExpectSame(Eval("=OFFSET(A1,1,1)"), Eval("=B2"), "one cell");
  ExpectSame(Eval("=INDIRECT(\"Sheet2!A1:B2\")"), Eval("=Sheet2!A1:B2"), "qualified rectangle");
  ExpectSame(Eval("=INDIRECT(\"C3\")"), Eval("=C3"), "one cell by text");
  ExpectSame(Eval("=SUM(OFFSET(A1,0,0,2,2))"), Eval("=SUM(A1:B2)"), "range argument");
}

TEST_F(BuiltRefTest, TheReadRecordsTheRectangleAndTheProducerDoesNot) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const parser::AstNode& offset = Parse("=OFFSET(A1,1,1,2,2)");
  (void)eval_node_ref(offset, arena_, default_registry(), ctx);
  EXPECT_TRUE(notes.empty());
  (void)eval_node_shaped(offset, arena_, default_registry(), ctx);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].rect.row_first, 1U);
  EXPECT_EQ(notes[0].rect.row_last, 2U);
  EXPECT_EQ(notes[0].rect.col_first, 1U);
  EXPECT_EQ(notes[0].rect.col_last, 2U);
}

TEST_F(BuiltRefTest, TheResolverRecordsOnTheCallersContext) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const parser::AstNode& indirect = Parse("=INDIRECT(\"Sheet2!B2:C3\")");
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  bool is_range = false;
  ErrorCode err = ErrorCode::Value;
  ASSERT_TRUE(resolve_reference_call(indirect, arena_, default_registry(), ctx, &sheet, &top, &left, &bottom, &right,
                                     &is_range, &err));
  EXPECT_EQ(sheet, "Sheet2");
  EXPECT_EQ(top, 1U);
  EXPECT_EQ(left, 1U);
  EXPECT_EQ(bottom, 2U);
  EXPECT_EQ(right, 2U);
  EXPECT_TRUE(is_range);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].sheet, "Sheet2");

  // A caller resolving only the position strips the callback and records nothing.
  notes.clear();
  ASSERT_TRUE(resolve_reference_call(indirect, arena_, default_registry(), ctx.without_dynamic_read_callback(), &sheet,
                                     &top, &left, &bottom, &right, &is_range, &err));
  EXPECT_TRUE(notes.empty());

  ASSERT_FALSE(resolve_reference_call(Parse("=INDIRECT(\"Nope!A1\")"), arena_, default_registry(), ctx, &sheet, &top,
                                      &left, &bottom, &right, &is_range, &err));
  EXPECT_EQ(err, ErrorCode::Ref);
}

// IF / CHOOSE / IFS / SWITCH with a scalar selector return the selected argument as it is.
class SelectedRefTest : public WalkerRefTest {
 protected:
  void ExpectSelected(const char* source, std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2,
                      std::uint16_t sheet = kHome, bool dynamic = false) {
    const Value ref = AsRef(source);
    ASSERT_TRUE(ref.is_ref()) << source << ": " << ref.debug_to_string();
    EXPECT_EQ(ref.ref_form(), RefForm::kArea) << source;
    EXPECT_EQ(ref.ref_is_dynamic(), dynamic) << source;
    const RefArea& a = ref.as_ref().areas[0];
    EXPECT_EQ(a.sheet_first, sheet) << source;
    EXPECT_EQ(a.row_first, r1) << source;
    EXPECT_EQ(a.col_first, c1) << source;
    EXPECT_EQ(a.row_last, r2) << source;
    EXPECT_EQ(a.col_last, c2) << source;
  }
};

TEST_F(SelectedRefTest, AScalarSelectorReturnsTheSelectedReference) {
  ExpectSelected("=IF(TRUE,B2:C3,A1)", 1, 1, 2, 2);
  ExpectSelected("=IF(FALSE,A1,OFFSET(A1,1,1,2,2))", 1, 1, 2, 2, kHome, /*dynamic=*/true);
  ExpectSelected("=CHOOSE(2,A1,Sheet2!B2)", 1, 1, 1, 1, 1);
  ExpectSelected("=IFS(FALSE,A1,TRUE,C3)", 2, 2, 2, 2);
  ExpectSelected("=SWITCH(2,1,A1,2,B2:B3)", 1, 1, 2, 1);
  ExpectSelected("=SWITCH(9,1,A1,D4)", 3, 3, 3, 3);
  ExpectSelected("=IF(TRUE,IF(FALSE,A1,INDIRECT(\"C3\")))", 2, 2, 2, 2, kHome, /*dynamic=*/true);
}

TEST_F(SelectedRefTest, OtherSelectionsAreValues) {
  const Value number = AsRef("=IF(TRUE,5)");
  ASSERT_TRUE(number.is_number());
  EXPECT_EQ(number.as_number(), 5.0);
  const Value missing_else = AsRef("=IF(FALSE,A1)");
  ASSERT_TRUE(missing_else.is_boolean());
  EXPECT_FALSE(missing_else.as_boolean());
  EXPECT_TRUE(AsRef("=IF({TRUE,FALSE},A1,B1)").is_array());
  EXPECT_TRUE(AsRef("=CHOOSE({1,2},A1,B1)").is_array());
  ExpectError(AsRef("=IFS(FALSE,A1)"), ErrorCode::NA, "no condition holds");
  ExpectError(AsRef("=CHOOSE(3,A1,B1)"), ErrorCode::Value, "index out of range");
}

TEST_F(SelectedRefTest, ThreeDimensionalArmsKeepTheirErrors) {
  ExpectError(AsRef("=IF(TRUE,Sheet2:Sheet3!A1)"), ErrorCode::Value, "IF over a 3-D arm");
  ExpectError(AsRef("=CHOOSE(1,Sheet2:Sheet3!A1)"), ErrorCode::Value, "CHOOSE over a 3-D arm");
  ExpectError(Eval("=IFS(TRUE,Sheet2:Sheet3!A1)"), ErrorCode::Ref, "IFS over a 3-D arm");
  ExpectError(Eval("=SWITCH(1,1,(Sheet2:Sheet3!A1))"), ErrorCode::Name, "grouped 3-D arm");
  ExpectError(AsRef("=(Sheet2:Sheet3!A1)"), ErrorCode::Name, "grouped 3-D span");
}

TEST_F(SelectedRefTest, ValueModeReadsTheSelectedReference) {
  ExpectSame(Eval("=IF(TRUE,B2:C3)"), Eval("=B2:C3"), "IF rectangle");
  ExpectSame(Eval("=CHOOSE(1,C3)"), Eval("=C3"), "CHOOSE one cell");
  ExpectSame(Eval("=SUM(IFS(TRUE,A1:B2))"), Eval("=SUM(A1:B2)"), "IFS range argument");
  ExpectSame(Eval("=SWITCH(1,1,Sheet2!A1:B2)"), Eval("=Sheet2!A1:B2"), "SWITCH qualified rectangle");
}

TEST_F(SelectedRefTest, ASelectedDynamicReferenceIsRecordedByItsRead) {
  std::vector<Note> notes;
  const EvalContext ctx = Ctx().with_dynamic_read_callback(&CollectNote, &notes);
  const parser::AstNode& node = Parse("=IF(TRUE,OFFSET(A1,1,1,2,2))");
  (void)eval_node_ref(node, arena_, default_registry(), ctx);
  EXPECT_TRUE(notes.empty());
  (void)eval_node_shaped(node, arena_, default_registry(), ctx);
  ASSERT_EQ(notes.size(), 1U);
  EXPECT_EQ(notes[0].rect.row_first, 1U);
  EXPECT_EQ(notes[0].rect.col_last, 2U);
}

// INDEX over a reference with scalar indices returns the picked part as a reference.
TEST_F(SelectedRefTest, IndexPicksAReference) {
  ExpectSelected("=INDEX(A1:C3,2,3)", 1, 2, 1, 2);
  ExpectSelected("=INDEX(A1:C3,2,0)", 1, 0, 1, 2);
  ExpectSelected("=INDEX(A1:C3,0,2)", 0, 1, 2, 1);
  ExpectSelected("=INDEX(A1:C3,0,0)", 0, 0, 2, 2);
  ExpectSelected("=INDEX(B2:B9,3)", 3, 1, 3, 1);
  ExpectSelected("=INDEX(B2:E2,3)", 1, 3, 1, 3);
  ExpectSelected("=INDEX(Sheet2!A1:B4,4,2)", 3, 1, 3, 1, 1);
  ExpectSelected("=INDEX((A1:B2,C3:D4),2,1,2)", 3, 2, 3, 2);
  ExpectSelected("=INDEX(INDEX(A1:C3,0,2),3)", 2, 1, 2, 1);
  ExpectSelected("=INDEX(IF(TRUE,B2:C3),1,2)", 1, 2, 1, 2);
  ExpectSelected("=INDEX(OFFSET(A1,1,1,3,3),2,2)", 2, 2, 2, 2, kHome, /*dynamic=*/true);
  const Value col = AsRef("=INDEX(A:B,0,2)");
  ASSERT_TRUE(col.is_ref());
  EXPECT_EQ(col.as_ref().areas[0].whole_col, 1U);
  EXPECT_EQ(col.as_ref().areas[0].col_first, 1U);
  const Value cell = AsRef("=INDEX(A:A,5)");
  ASSERT_TRUE(cell.is_ref());
  EXPECT_EQ(cell.as_ref().areas[0].whole_col, 0U);
  EXPECT_EQ(cell.as_ref().areas[0].row_first, 4U);
}

TEST_F(SelectedRefTest, IndexOverValuesOrArrayIndicesGivesValues) {
  EXPECT_TRUE(AsRef("=INDEX({1,2;3,4},2,1)").is_number());
  EXPECT_TRUE(AsRef("=INDEX(A1:C3,{1;2},1)").is_array());
  EXPECT_TRUE(AsRef("=INDEX(IF(TRUE,{5,6}),2)").is_number());
  ExpectError(AsRef("=INDEX(A1:C3,4,1)"), ErrorCode::Ref, "row past the end");
  ExpectError(AsRef("=INDEX(A1:C3,2)"), ErrorCode::Ref, "2-D reference with a row alone");
  ExpectError(AsRef("=INDEX(OFFSET(A1,-1,0),1)"), ErrorCode::Ref, "source error");
}

TEST_F(SelectedRefTest, ValueModeReadsTheIndexedReference) {
  ExpectSame(Eval("=INDEX(A1:C3,2,3)"), Eval("=C2"), "cell");
  ExpectSame(Eval("=INDEX(A1:C3,0,2)"), Eval("=B1:B3"), "column");
  ExpectSame(Eval("=SUM(INDEX(A1:C3,2,0))"), Eval("=SUM(A2:C2)"), "row in an aggregate");
  ExpectSame(Eval("=INDEX(OFFSET(A1,0,0,3,3),{1;2},2)"), Eval("=B1:B2"), "array indices over a built reference");
  ExpectSame(Eval("=INDEX(IF(TRUE,{5,6}),2)"), Eval("=6"), "value source");
}

// XLOOKUP over a written return_array returns the matched slice as a reference; a scalar
// lookup's fallback is returned as it is.
TEST_F(SelectedRefTest, XlookupReturnsTheMatchedSlice) {
  // A3 holds 31 and B1 holds 12.
  ExpectSelected("=XLOOKUP(31,A1:A5,B1:C5)", 2, 1, 2, 2);
  ExpectSelected("=XLOOKUP(12,A1:C1,A3:C4)", 2, 1, 3, 1);
  ExpectSelected("=XLOOKUP(-1,A1:A5,B1:B5,D4)", 3, 3, 3, 3);
  ExpectSelected("=XLOOKUP(-1,A1:A5,B1:B5,OFFSET(A1,1,1))", 1, 1, 1, 1, kHome, /*dynamic=*/true);
  ExpectSelected("=INDEX(XLOOKUP(31,A1:A5,B1:C5),2)", 2, 2, 2, 2);
}

TEST_F(SelectedRefTest, XlookupOverValuesOrArrayLookupsGivesValues) {
  EXPECT_TRUE(AsRef("=XLOOKUP(2,{1;2},{5;6})").is_number());
  EXPECT_TRUE(AsRef("=XLOOKUP(-1,A1:A5,B1:B5,7)").is_number());
  EXPECT_TRUE(AsRef("=XLOOKUP({1;2},A1:A5,B1:B5)").is_array());
  ExpectError(AsRef("=XLOOKUP(-1,A1:A5,B1:B5)"), ErrorCode::NA, "a miss with no fallback");
}

TEST_F(SelectedRefTest, ValueModeReadsTheMatchedSlice) {
  ExpectSame(Eval("=XLOOKUP(31,A1:A5,B1:C5)"), Eval("=B3:C3"), "matched row");
  ExpectSame(Eval("=XLOOKUP(-1,A1:A5,B1:B5,D4)"), Eval("=D4"), "fallback reference");
}

// TAKE / DROP over a bounded reference and TRIMRANGE over a reference return the kept rectangle.
TEST_F(SelectedRefTest, TakeDropAndTrimReturnTheKeptRectangle) {
  ExpectSelected("=TAKE(A1:C5,2)", 0, 0, 1, 2);
  ExpectSelected("=TAKE(A1:C5,-1,-2)", 4, 1, 4, 2);
  ExpectSelected("=DROP(A1:C5,3,1)", 3, 1, 4, 2);
  ExpectSelected("=DROP(Sheet2!A1:B4,-1)", 0, 0, 2, 1, 1);
  ExpectSelected("=TAKE(OFFSET(A1,0,0,4,4),1,1)", 0, 0, 0, 0, kHome, /*dynamic=*/true);
  ExpectSelected("=INDEX(DROP(A1:C5,1),1,1)", 1, 0, 1, 0);
  // Row 5 and column E hold values but E5 is empty, so A1:E5 trims to itself and D4:F9 to D4:E5.
  ExpectSelected("=TRIMRANGE(A1:E5)", 0, 0, 4, 4);
  ExpectSelected("=TRIMRANGE(D4:F9)", 3, 3, 4, 4);
}

TEST_F(SelectedRefTest, TakeDropOverValuesGiveValues) {
  EXPECT_TRUE(AsRef("=TAKE({1,2;3,4},1)").is_array());
  EXPECT_TRUE(AsRef("=DROP(IF(TRUE,{1,2;3,4}),1)").is_array());
  ExpectError(AsRef("=TAKE(A1:C5,0)"), ErrorCode::Calc, "zero rows");
  ExpectError(AsRef("=DROP(A1:C5,5)"), ErrorCode::Calc, "every row dropped");
  ExpectError(AsRef("=TAKE(OFFSET(A1,-1,0),1)"), ErrorCode::Ref, "source error");
}

TEST_F(SelectedRefTest, ValueModeReadsTheKeptRectangle) {
  ExpectSame(Eval("=TAKE(A1:C5,2)"), Eval("=A1:C2"), "taken rows");
  ExpectSame(Eval("=DROP(A1:C5,3,1)"), Eval("=B4:C5"), "dropped rows and column");
  ExpectSame(Eval("=TAKE(A1:C5,1,1)"), Eval("=A1"), "one cell");
  ExpectSame(Eval("=SUM(DROP(A1:C5,4))"), Eval("=SUM(A5:C5)"), "aggregate");
  ExpectSame(Eval("=TRIMRANGE(D4:F9)"), Eval("=D4:E5"), "trimmed");
}

// TAKE / DROP over a whole column. As a reference argument they name the kept rows of the whole column;
// read as values, the column is still narrowed to its populated rows first, so TAKE(A:A,-1) reads A5 where
// the reference names A1048576.
TEST_F(SelectedRefTest, WholeAxisTakeAndDropKnownReads) {
  struct Case {
    const char* source;
    double value;
  };
  const Case cases[] = {
      {"=ROWS(OFFSET(DROP(A:A,1),0,0))", 1048575.0},
      {"=ROWS(OFFSET(TAKE(A:A,-1),0,0))", 1.0},
      {"=ROW(INDEX(DROP(A:A,1),1))", 2.0},
      {"=ROW(INDEX(TAKE(A:A,-1),1))", 1048576.0},
      {"=ROWS(A1:TAKE(A:A,-1))", 1048576.0},
      {"=ROWS(DROP(A:A,1):B2)", 1048575.0},
      {"=ROWS(TAKE(A:A,-2))", 2.0},
      {"=ROWS(DROP(A:A,1))", 1048575.0},
      {"=SUM(DROP(A:A,1))", 144.0},
      {"=TAKE(A:A,-1)", 51.0},
  };
  for (const Case& c : cases) {
    const Value v = Eval(c.source);
    ASSERT_TRUE(v.is_number()) << c.source << ": " << v.debug_to_string();
    EXPECT_EQ(v.as_number(), c.value) << c.source;
  }
}

}  // namespace
}  // namespace eval
}  // namespace formulon
