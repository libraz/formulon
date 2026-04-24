// Copyright 2026 libraz. Licensed under the MIT License.
//
// Direct tests for the internal A1 reference parser
// `refs_internal::parse_a1_ref`. End-to-end INDIRECT / OFFSET behaviour
// lives in `builtins_references_test.cpp`; this file focuses on the
// parser's shape recognition, in particular the full-column (`D:D`,
// `$FF:FG`) and full-row (`5:5`, `$12:$23`) extensions that feed
// `ROW(INDIRECT("D:D"))` / `COLUMN(INDIRECT("5:5"))`.

#include "eval/reference_lazy.h"

#include <cstdint>

#include "gtest/gtest.h"
#include "sheet.h"

namespace formulon {
namespace eval {
namespace refs_internal {
namespace {

// ---------------------------------------------------------------------------
// Single-cell and range shapes (regression).
// ---------------------------------------------------------------------------

TEST(ParseA1Ref, SingleCellBasic) {
  const A1Parse p = parse_a1_ref("A1");
  EXPECT_TRUE(p.valid);
  EXPECT_FALSE(p.is_range);
  EXPECT_FALSE(p.is_full_col);
  EXPECT_FALSE(p.is_full_row);
  EXPECT_EQ(p.row, 0U);
  EXPECT_EQ(p.col, 0U);
}

TEST(ParseA1Ref, SingleCellRequiresRow) {
  // Bare column letter is not a valid single-cell ref.
  const A1Parse p = parse_a1_ref("D");
  EXPECT_FALSE(p.valid);
}

TEST(ParseA1Ref, RangeBasic) {
  const A1Parse p = parse_a1_ref("A1:B2");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_range);
  EXPECT_FALSE(p.is_full_col);
  EXPECT_FALSE(p.is_full_row);
  EXPECT_EQ(p.row, 0U);
  EXPECT_EQ(p.col, 0U);
  EXPECT_EQ(p.row2, 1U);
  EXPECT_EQ(p.col2, 1U);
}

TEST(ParseA1Ref, TrailingGarbageIsInvalid) {
  EXPECT_FALSE(parse_a1_ref("A1x").valid);
  EXPECT_FALSE(parse_a1_ref("D:D:D").valid);
}

TEST(ParseA1Ref, MixedShapesAreInvalid) {
  // `D:5` mixes column and row endpoints; neither full-col nor full-row
  // nor single-endpoint accepts it.
  EXPECT_FALSE(parse_a1_ref("D:5").valid);
  EXPECT_FALSE(parse_a1_ref("5:D").valid);
}

// ---------------------------------------------------------------------------
// Full-column shapes.
// ---------------------------------------------------------------------------

TEST(ParseA1Ref, FullColumnSingle) {
  const A1Parse p = parse_a1_ref("D:D");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_FALSE(p.is_full_row);
  EXPECT_TRUE(p.is_range);
  EXPECT_EQ(p.col, 3U);
  EXPECT_EQ(p.col2, 3U);
  EXPECT_EQ(p.row, 0U);
  EXPECT_EQ(p.row2, Sheet::kMaxRows - 1U);
}

TEST(ParseA1Ref, FullColumnSpan) {
  const A1Parse p = parse_a1_ref("A:C");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.col, 0U);
  EXPECT_EQ(p.col2, 2U);
}

TEST(ParseA1Ref, FullColumnMixedAbsolute) {
  // `$FF:FG` -> columns FF (162) and FG (163) 1-based, so 161..162 0-based.
  const A1Parse p = parse_a1_ref("$FF:FG");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.col, 161U);
  EXPECT_EQ(p.col2, 162U);
}

TEST(ParseA1Ref, FullColumnReversedNormalises) {
  // `C:A` -> min=A, max=C after normalisation.
  const A1Parse p = parse_a1_ref("C:A");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.col, 0U);
  EXPECT_EQ(p.col2, 2U);
}

TEST(ParseA1Ref, FullColumnLowercase) {
  // `s:s` -> column S (19) 1-based, 18 0-based (case-insensitive).
  const A1Parse p = parse_a1_ref("s:s");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.col, 18U);
  EXPECT_EQ(p.col2, 18U);
}

TEST(ParseA1Ref, FullColumnXfd) {
  const A1Parse p = parse_a1_ref("XFD:XFD");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.col, Sheet::kMaxCols - 1U);
  EXPECT_EQ(p.col2, Sheet::kMaxCols - 1U);
}

TEST(ParseA1Ref, FullColumnWithSheetQualifier) {
  const A1Parse p = parse_a1_ref("Sheet1!D:D");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_col);
  EXPECT_EQ(p.sheet, "Sheet1");
  EXPECT_EQ(p.col, 3U);
  EXPECT_EQ(p.col2, 3U);
  EXPECT_EQ(p.row, 0U);
  EXPECT_EQ(p.row2, Sheet::kMaxRows - 1U);
}

// ---------------------------------------------------------------------------
// Full-row shapes.
// ---------------------------------------------------------------------------

TEST(ParseA1Ref, FullRowSingle) {
  const A1Parse p = parse_a1_ref("5:5");
  ASSERT_TRUE(p.valid);
  EXPECT_FALSE(p.is_full_col);
  EXPECT_TRUE(p.is_full_row);
  EXPECT_TRUE(p.is_range);
  EXPECT_EQ(p.row, 4U);
  EXPECT_EQ(p.row2, 4U);
  EXPECT_EQ(p.col, 0U);
  EXPECT_EQ(p.col2, Sheet::kMaxCols - 1U);
}

TEST(ParseA1Ref, FullRowSpanWithAbsolute) {
  const A1Parse p = parse_a1_ref("$12:$23");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_row);
  EXPECT_EQ(p.row, 11U);
  EXPECT_EQ(p.row2, 22U);
}

TEST(ParseA1Ref, FullRowLeadingZeros) {
  const A1Parse p = parse_a1_ref("05:05");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_row);
  EXPECT_EQ(p.row, 4U);
  EXPECT_EQ(p.row2, 4U);
}

TEST(ParseA1Ref, FullRowReversedNormalises) {
  const A1Parse p = parse_a1_ref("10:3");
  ASSERT_TRUE(p.valid);
  EXPECT_TRUE(p.is_full_row);
  EXPECT_EQ(p.row, 2U);
  EXPECT_EQ(p.row2, 9U);
}

}  // namespace
}  // namespace refs_internal
}  // namespace eval
}  // namespace formulon
