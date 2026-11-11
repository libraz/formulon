// `normalize_column_spans`: loaded `<cols>` / BrtColInfo spans become sorted
// and disjoint, the span that sorts first keeping any shared columns.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "sheet_layout.h"
#include "sheet_layout_edit.h"

namespace formulon {
namespace {

ColumnLayout Span(std::uint32_t first, std::uint32_t last, std::uint32_t xf) {
  ColumnLayout span;
  span.first = first;
  span.last = last;
  span.has_style = true;
  span.style_xf = xf;
  return span;
}

// (first, last, xf) of every span.
std::vector<std::vector<std::uint32_t>> Shape(const SheetLayout& layout) {
  std::vector<std::vector<std::uint32_t>> out;
  for (const ColumnLayout& span : layout.columns) {
    out.push_back({span.first, span.last, span.style_xf});
  }
  return out;
}

struct NormalizeCase {
  const char* name;
  std::vector<ColumnLayout> input;
  std::vector<std::vector<std::uint32_t>> expected;
};

class ColumnSpanNormalize : public ::testing::TestWithParam<NormalizeCase> {};

TEST_P(ColumnSpanNormalize, SortsAndClipsLaterSpans) {
  SheetLayout layout;
  layout.columns = GetParam().input;
  normalize_column_spans(layout);
  EXPECT_EQ(Shape(layout), GetParam().expected);
}

INSTANTIATE_TEST_SUITE_P(
    Cases, ColumnSpanNormalize,
    ::testing::Values(NormalizeCase{"canonical", {Span(0, 1, 1), Span(3, 4, 2)}, {{0, 1, 1}, {3, 4, 2}}},
                      NormalizeCase{"unsorted", {Span(5, 6, 1), Span(0, 1, 2)}, {{0, 1, 2}, {5, 6, 1}}},
                      // The span that starts first keeps the overlap, whatever the file order.
                      NormalizeCase{"overlap", {Span(2, 4, 1), Span(0, 3, 2)}, {{0, 3, 2}, {4, 4, 1}}},
                      // Equal starts: the shorter span sorts first and keeps its columns.
                      NormalizeCase{"same start", {Span(2, 6, 1), Span(2, 3, 2)}, {{2, 3, 2}, {4, 6, 1}}},
                      NormalizeCase{"nested", {Span(0, 9, 1), Span(3, 4, 2)}, {{0, 9, 1}}},
                      // Adjacent equal spans are not merged.
                      NormalizeCase{"adjacent", {Span(2, 3, 1), Span(0, 1, 1)}, {{0, 1, 1}, {2, 3, 1}}},
                      NormalizeCase{"reversed bounds dropped", {Span(4, 2, 1), Span(0, 0, 2)}, {{0, 0, 2}}}),
    [](const ::testing::TestParamInfo<NormalizeCase>& param_info) {
      std::string name = param_info.param.name;
      for (char& c : name) {
        if (c == ' ') {
          c = '_';
        }
      }
      return name;
    });

TEST(ColumnSpanNormalizeBounds, IdenticalSpansAcrossTheGridCollapseToOne) {
  SheetLayout layout;
  layout.columns.assign(16384U, Span(0, 16383, 1));
  normalize_column_spans(layout);
  EXPECT_EQ(Shape(layout), (std::vector<std::vector<std::uint32_t>>{{0, 16383, 1}}));
}

TEST(ColumnSpanNormalizeBounds, KeepsEveryFieldOfTheClippedSpan) {
  SheetLayout layout;
  ColumnLayout wide = Span(1, 5, 3);
  wide.width = 12.5;
  wide.has_width = true;
  wide.hidden = true;
  layout.columns = {wide, Span(0, 2, 1)};
  normalize_column_spans(layout);
  ASSERT_EQ(layout.columns.size(), 2U);
  EXPECT_EQ(layout.columns[1].first, 3U);
  EXPECT_EQ(layout.columns[1].last, 5U);
  EXPECT_TRUE(layout.columns[1].has_width);
  EXPECT_DOUBLE_EQ(layout.columns[1].width, 12.5);
  EXPECT_TRUE(layout.columns[1].hidden);
}

}  // namespace
}  // namespace formulon
