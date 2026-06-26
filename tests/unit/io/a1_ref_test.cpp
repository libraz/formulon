// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for the shared A1 reference decoder in `io/a1_ref.{h,cpp}`.
// Covers the boundary cases shared by `cell_parser` and `sax_xml_reader`:
// the "A1" / "XFD1048576" extremes, overflow on `XFE`, oversize letters,
// and oversize / zero rows.

#include "io/a1_ref.h"

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "gtest/gtest.h"

namespace formulon {
namespace io {
namespace {

TEST(A1RefParseColumnLetters, SingleLetter) {
  std::size_t pos = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(parse_column_letters("A", &pos, &col));
  EXPECT_EQ(col, 1u);
  EXPECT_EQ(pos, 1u);
  pos = 0;
  EXPECT_TRUE(parse_column_letters("Z", &pos, &col));
  EXPECT_EQ(col, 26u);
}

TEST(A1RefParseColumnLetters, MultiLetterAndExcelMax) {
  std::size_t pos = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(parse_column_letters("AA", &pos, &col));
  EXPECT_EQ(col, 27u);
  pos = 0;
  EXPECT_TRUE(parse_column_letters("XFD", &pos, &col));
  EXPECT_EQ(col, 16384u);  // Excel column ceiling.
}

TEST(A1RefParseColumnLetters, FourLettersRejected) {
  std::size_t pos = 0;
  std::uint32_t col = 0;
  // 4 leading letters is past XFD; the helper must refuse.
  EXPECT_FALSE(parse_column_letters("ABCD", &pos, &col));
}

TEST(A1RefParseColumnLetters, EmptyOrLowercaseRejected) {
  std::size_t pos = 0;
  std::uint32_t col = 0;
  EXPECT_FALSE(parse_column_letters("", &pos, &col));
  pos = 0;
  EXPECT_FALSE(parse_column_letters("a1", &pos, &col));
  pos = 0;
  EXPECT_FALSE(parse_column_letters("123", &pos, &col));
}

TEST(A1RefParseColumnLetters, StopsAtFirstNonLetter) {
  std::size_t pos = 0;
  std::uint32_t col = 0;
  EXPECT_TRUE(parse_column_letters("AB12", &pos, &col));
  EXPECT_EQ(col, 28u);
  EXPECT_EQ(pos, 2u);  // stopped at '1'
}

TEST(A1RefParseUint, Basic) {
  std::size_t pos = 0;
  std::uint32_t v = 0;
  EXPECT_TRUE(parse_uint("0", &pos, &v));
  EXPECT_EQ(v, 0u);
  pos = 0;
  EXPECT_TRUE(parse_uint("42", &pos, &v));
  EXPECT_EQ(v, 42u);
  EXPECT_EQ(pos, 2u);
}

TEST(A1RefParseUint, EmptyOrNonDigitRejected) {
  std::size_t pos = 0;
  std::uint32_t v = 0;
  EXPECT_FALSE(parse_uint("", &pos, &v));
  pos = 0;
  EXPECT_FALSE(parse_uint("abc", &pos, &v));
}

TEST(A1RefParseUint, OverflowRejected) {
  std::size_t pos = 0;
  std::uint32_t v = 0;
  // 2^32 = 4294967296, one beyond the uint32_t ceiling.
  EXPECT_FALSE(parse_uint("4294967296", &pos, &v));
  // Boundary: 2^32 - 1 must succeed.
  pos = 0;
  EXPECT_TRUE(parse_uint("4294967295", &pos, &v));
  EXPECT_EQ(v, 0xFFFFFFFFu);
}

TEST(A1RefParseA1, BasicReferences) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  ASSERT_TRUE(parse_a1_ref("A1", &row, &col));
  EXPECT_EQ(row, 0u);
  EXPECT_EQ(col, 0u);
  ASSERT_TRUE(parse_a1_ref("B2", &row, &col));
  EXPECT_EQ(row, 1u);
  EXPECT_EQ(col, 1u);
  ASSERT_TRUE(parse_a1_ref("AA10", &row, &col));
  EXPECT_EQ(row, 9u);
  EXPECT_EQ(col, 26u);
}

TEST(A1RefParseA1, ExcelMaxAndJustBeyond) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  // Excel's largest legal cell.
  ASSERT_TRUE(parse_a1_ref("XFD1048576", &row, &col));
  EXPECT_EQ(row, 1048575u);
  EXPECT_EQ(col, 16383u);
  // Beyond the column ceiling: 4-letter columns are rejected by
  // `parse_column_letters`.
  EXPECT_FALSE(parse_a1_ref("XFEA1", &row, &col));
}

TEST(A1RefParseA1, RejectsOutOfRangeColumn) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  // "XFE" is a structurally valid 3-letter column whose numeric value is
  // one past XFD (16385 > Sheet::kMaxCols). It must be rejected, not
  // silently stored as a phantom column.
  EXPECT_FALSE(parse_a1_ref("XFE1", &row, &col));
  // "ZZZ" = 18278 columns, well past XFD.
  EXPECT_FALSE(parse_a1_ref("ZZZ1", &row, &col));
}

TEST(A1RefParseA1, RejectsOutOfRangeRow) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  // Row 1,048,577 is one past Excel's last row; it parses as a uint32_t
  // but exceeds Sheet::kMaxRows and must be rejected.
  EXPECT_FALSE(parse_a1_ref("A1048577", &row, &col));
  // A far-out-of-range but still uint32_t-valid row.
  EXPECT_FALSE(parse_a1_ref("A2000000", &row, &col));
}

TEST(A1RefParseA1, RejectsZeroRow) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  // Excel rows are 1-based; row 0 is invalid.
  EXPECT_FALSE(parse_a1_ref("A0", &row, &col));
}

TEST(A1RefParseA1, RejectsTrailingCharacters) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_FALSE(parse_a1_ref("A1B", &row, &col));
  EXPECT_FALSE(parse_a1_ref("A1!", &row, &col));
  EXPECT_FALSE(parse_a1_ref("A1 ", &row, &col));
}

TEST(A1RefParseA1, RejectsMissingColumnOrRow) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  EXPECT_FALSE(parse_a1_ref("", &row, &col));
  EXPECT_FALSE(parse_a1_ref("123", &row, &col));
  EXPECT_FALSE(parse_a1_ref("A", &row, &col));
}

TEST(A1RefParseA1, OverflowRow) {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  // Beyond uint32_t: caught by `parse_uint`.
  EXPECT_FALSE(parse_a1_ref("A4294967296", &row, &col));
}

}  // namespace
}  // namespace io
}  // namespace formulon
