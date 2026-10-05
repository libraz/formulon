// Sheet range reads over literals, formulas, and spill phantoms.

#include <algorithm>
#include <chrono>
#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace {
TEST(SheetReadRange, MatchesResolveCellValueOverLiteralsAndGaps) {
  Sheet s("Sheet1");
  s.set_cell_value(1U, 1U, Value::number(1.0));
  s.set_cell_value(1U, 3U, Value::number(2.0));
  s.set_cell_value(3U, 2U, Value::text("x"));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 4U, 0U, 4U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 25U);
  EXPECT_TRUE(formula_indices.empty());
  std::size_t index = 0;
  for (std::uint32_t row = 0; row <= 4U; ++row) {
    for (std::uint32_t col = 0; col <= 4U; ++col, ++index) {
      EXPECT_EQ(bulk[index], s.resolve_cell_value(row, col)) << "at (" << row << "," << col << ")";
    }
  }
}
TEST(SheetReadRange, SurfacesSpillPhantomValues) {
  Sheet s("Sheet1");
  std::vector<Value> cells = {Value::number(1.0), Value::number(2.0), Value::number(3.0)};
  ASSERT_TRUE(s.commit_spill(0U, 0U, 3U, 1U, std::move(cells)));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 2U, 0U, 0U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 3U);
  EXPECT_TRUE(formula_indices.empty());
  EXPECT_EQ(bulk[0], Value::number(1.0));
  EXPECT_EQ(bulk[1], Value::number(2.0));
  EXPECT_EQ(bulk[2], Value::number(3.0));
}
TEST(SheetReadRange, ReportsFormulaCoordinatesInsteadOfEvaluatingThem) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(5.0));
  s.set_cell_formula(0U, 1U, "=A1*2");

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 1U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 2U);
  ASSERT_EQ(formula_indices.size(), 1U);
  EXPECT_EQ(formula_indices[0], 1U);
  // The formula slot carries only its cached value; evaluation is the
  // caller's job because it re-enters the sheet.
  EXPECT_TRUE(bulk[1].is_blank());
}
TEST(SheetReadRange, AppendsToTheCallersBufferAndIndexesAbsolutely) {
  Sheet s("Sheet1");
  s.set_cell_formula(0U, 0U, "=1");

  Arena arena;
  std::vector<Value> bulk = {Value::number(99.0)};
  std::vector<std::size_t> formula_indices;
  s.read_range(0U, 0U, 0U, 0U, arena, bulk, formula_indices);

  ASSERT_EQ(bulk.size(), 2U);
  EXPECT_EQ(bulk[0], Value::number(99.0));
  ASSERT_EQ(formula_indices.size(), 1U);
  EXPECT_EQ(formula_indices[0], 1U);
}
TEST(SheetReadRange, ReversedOrOutOfRangeRectangleAppendsNothing) {
  Sheet s("Sheet1");
  s.set_cell_value(0U, 0U, Value::number(1.0));

  Arena arena;
  std::vector<Value> bulk;
  std::vector<std::size_t> formula_indices;
  s.read_range(2U, 1U, 0U, 0U, arena, bulk, formula_indices);
  s.read_range(0U, 0U, 2U, 1U, arena, bulk, formula_indices);
  s.read_range(0U, Sheet::kMaxRows, 0U, 0U, arena, bulk, formula_indices);
  s.read_range(0U, 0U, 0U, Sheet::kMaxCols, arena, bulk, formula_indices);

  EXPECT_TRUE(bulk.empty());
  EXPECT_TRUE(formula_indices.empty());
}

}  // namespace
}  // namespace formulon
