//
// Unit tests for the `Ref` variant of `Value`: forms, inline vs out-of-line
// storage, error elements, triviality, and the value-boundary predicate.

#include <cstdint>
#include <type_traits>

#include "eval/eval_state.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace {

RefArea make_area(std::uint32_t r1, std::uint32_t c1, std::uint32_t r2, std::uint32_t c2, std::uint16_t s1 = 0,
                  std::uint16_t s2 = 0) {
  RefArea a{};
  a.row_first = r1;
  a.row_last = r2;
  a.col_first = c1 & 0x3FFFU;
  a.col_last = c2 & 0x3FFFU;
  a.origin = static_cast<std::uint16_t>(RefOrigin::kSpelled);
  a.sheet_first = s1;
  a.sheet_last = s2;
  return a;
}

TEST(ValueRefTest, LayoutAndTriviality) {
  static_assert(sizeof(Value) == 24);
  static_assert(sizeof(RefArea) == 16);
  EXPECT_TRUE(std::is_trivially_copyable_v<Value>);
  EXPECT_TRUE(std::is_trivially_copyable_v<RefArea>);
  EXPECT_TRUE(std::is_trivially_destructible_v<RefValue>);
}

TEST(ValueRefTest, InlineAreaOwnBook) {
  Arena arena;
  const RefArea a = make_area(1, 2, 3, 4, 5, 5);
  const Value v = make_ref(arena, RefForm::kArea, 0, &a, 1);
  ASSERT_TRUE(v.is_ref());
  EXPECT_EQ(ValueKind::Ref, v.kind());
  EXPECT_EQ(RefForm::kArea, v.ref_form());
  const RefView view = v.as_ref();
  EXPECT_EQ(RefForm::kArea, view.form);
  EXPECT_EQ(1U, view.area_count);
  EXPECT_EQ(0U, view.book);
  EXPECT_EQ(1U, view.areas[0].row_first);
  EXPECT_EQ(3U, view.areas[0].row_last);
  EXPECT_EQ(2U, view.areas[0].col_first);
  EXPECT_EQ(4U, view.areas[0].col_last);
  EXPECT_EQ(5U, view.areas[0].sheet_first);
  EXPECT_EQ(0U, arena.bytes_allocated());
}

TEST(ValueRefTest, Span3DDerivedFromSheets) {
  Arena arena;
  const RefArea span = make_area(0, 0, 0, 0, 1, 3);
  EXPECT_EQ(RefForm::kSpan3D, make_ref(arena, RefForm::kArea, 0, &span, 1).ref_form());
  EXPECT_EQ(RefForm::kSpan3D, Value::ref_area(span).ref_form());
  const RefArea degenerate = make_area(0, 0, 0, 0, 2, 2);
  EXPECT_EQ(RefForm::kArea, Value::ref_area(degenerate).ref_form());
  EXPECT_EQ(RefForm::kArea, make_ref(arena, RefForm::kSpan3D, 0, &degenerate, 1).ref_form());
}

TEST(ValueRefTest, ExternalBookAreaIsOutOfLine) {
  Arena arena;
  const RefArea a = make_area(0, 0, 9, 0);
  const Value v = make_ref(arena, RefForm::kArea, 2, &a, 1);
  ASSERT_TRUE(v.is_ref());
  const RefView view = v.as_ref();
  EXPECT_EQ(RefForm::kArea, view.form);
  EXPECT_EQ(2U, view.book);
  EXPECT_EQ(1U, view.area_count);
  EXPECT_EQ(9U, view.areas[0].row_last);
  EXPECT_NE(&a, view.areas);
}

TEST(ValueRefTest, UnionKeepsEveryArea) {
  Arena arena;
  const RefArea areas[] = {make_area(0, 0, 1, 0), make_area(0, 2, 1, 2), make_area(5, 5, 5, 5)};
  const Value v = make_ref(arena, RefForm::kUnion, 0, areas, 3);
  EXPECT_EQ(RefForm::kUnion, v.ref_form());
  const RefView view = v.as_ref();
  ASSERT_EQ(3U, view.area_count);
  EXPECT_EQ(2U, view.areas[1].col_first);
  EXPECT_EQ(5U, view.areas[2].row_first);
  EXPECT_EQ(nullptr, view.element_books);
}

TEST(ValueRefTest, GridWithErrorElementAndBooks) {
  Arena arena;
  const RefArea elems[] = {make_area(0, 0, 0, 0), RefArea::error_element(ErrorCode::Value)};
  const std::uint16_t books[] = {0, 3};
  const Value v = make_ref(arena, RefForm::kGrid, 0, elems, 2, 1, 2, books);
  EXPECT_EQ(RefForm::kGrid, v.ref_form());
  const RefView view = v.as_ref();
  EXPECT_EQ(1U, view.rows);
  EXPECT_EQ(2U, view.cols);
  ASSERT_NE(nullptr, view.element_books);
  EXPECT_EQ(3U, view.element_books[1]);
  EXPECT_FALSE(view.areas[0].is_error_element());
  ASSERT_TRUE(view.areas[1].is_error_element());
  EXPECT_EQ(RefArea::kErrorElement, view.areas[1].sheet_first);
  EXPECT_EQ(ErrorCode::Value, view.areas[1].error_code());
}

TEST(ValueRefTest, ErrorElementRoundTripsEveryCode) {
  for (std::uint16_t c = 0; c <= static_cast<std::uint16_t>(ErrorCode::Unknown); ++c) {
    const RefArea e = RefArea::error_element(static_cast<ErrorCode>(c));
    EXPECT_TRUE(e.is_error_element());
    EXPECT_EQ(static_cast<ErrorCode>(c), e.error_code());
  }
}

TEST(ValueRefTest, OpaqueElementIsBlank) {
  const Value v = Value::blank(BlankGridProjection::kOpaqueElement);
  EXPECT_TRUE(v.is_blank());
  EXPECT_TRUE(v.is_opaque_element());
  EXPECT_FALSE(Value::blank().is_opaque_element());
}

TEST(ValueRefTest, BoundaryPredicateRejectsReferences) {
  Arena arena;
  const RefArea a = make_area(0, 0, 0, 0);
  const RefArea span = make_area(0, 0, 0, 0, 0, 2);
  const RefArea two[] = {a, a};
  EXPECT_TRUE(escapes_boundary(make_ref(arena, RefForm::kArea, 0, &a, 1)));
  EXPECT_TRUE(escapes_boundary(make_ref(arena, RefForm::kArea, 1, &a, 1)));
  EXPECT_TRUE(escapes_boundary(make_ref(arena, RefForm::kSpan3D, 0, &span, 1)));
  EXPECT_TRUE(escapes_boundary(make_ref(arena, RefForm::kUnion, 0, two, 2)));
  EXPECT_TRUE(escapes_boundary(make_ref(arena, RefForm::kGrid, 0, two, 2, 1, 2)));
  EXPECT_TRUE(escapes_boundary(Value::blank(BlankGridProjection::kOpaqueElement)));
}

TEST(ValueRefTest, BoundaryPredicateAcceptsEveryOtherKind) {
  Arena arena;
  EXPECT_FALSE(escapes_boundary(Value::blank()));
  EXPECT_FALSE(escapes_boundary(Value::blank(BlankGridProjection::kReferenceGridZero)));
  EXPECT_FALSE(escapes_boundary(Value::blank(BlankGridProjection::kValueArrayZero)));
  EXPECT_FALSE(escapes_boundary(Value::number(1.0)));
  EXPECT_FALSE(escapes_boundary(Value::boolean(true)));
  EXPECT_FALSE(escapes_boundary(Value::text("x")));
  EXPECT_FALSE(escapes_boundary(Value::error(ErrorCode::NA)));
  const ArrayValue arr{0, 0, nullptr};
  EXPECT_FALSE(escapes_boundary(Value::array(&arr)));
  EXPECT_FALSE(escapes_boundary(Value::lambda(nullptr)));
}

#ifdef NDEBUG
// Release degrades an escaping value to #VALUE! at every boundary instead of
// aborting; Debug asserts there instead.
TEST(ValueRefTest, BoundariesDegradeEscapingValuesToValueError) {
  Arena arena;
  const RefArea a = make_area(0, 0, 0, 0);
  const Value escaping[] = {make_ref(arena, RefForm::kArea, 0, &a, 1),
                            Value::blank(BlankGridProjection::kOpaqueElement)};
  for (const Value& v : escaping) {
    const Value contained = contain_escape(v);
    ASSERT_TRUE(contained.is_error());
    EXPECT_EQ(contained.as_error(), ErrorCode::Value);

    Sheet sheet("Sheet1");
    sheet.set_cell_value(0U, 0U, v);
    const Cell* cell = sheet.cell_at(0U, 0U);
    ASSERT_NE(cell, nullptr);
    EXPECT_TRUE(cell->cached_value.is_error());

    eval::EvalState state;
    state.memoize(&sheet, 1U, 1U, v);
    const Value* memo = state.lookup_memo(&sheet, 1U, 1U);
    ASSERT_NE(memo, nullptr);
    EXPECT_TRUE(memo->is_error());
  }
  EXPECT_TRUE(contain_escape(Value::number(1.0)).is_number());
}
#else
TEST(ValueRefDeathTest, BoundariesAssertOnEscapingValues) {
  Arena arena;
  const RefArea a = make_area(0, 0, 0, 0);
  EXPECT_DEATH(contain_escape(make_ref(arena, RefForm::kArea, 0, &a, 1)), "");
  EXPECT_DEATH(contain_escape(Value::blank(BlankGridProjection::kOpaqueElement)), "");
}
#endif

}  // namespace
}  // namespace formulon
