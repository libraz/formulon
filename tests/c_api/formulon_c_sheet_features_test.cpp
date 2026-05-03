// Copyright 2026 libraz. Licensed under the MIT License.
//
// C ABI smoke tests for the sheet UI feature surface (merges,
// hyperlinks, comments). Each test creates a default workbook,
// exercises the new entry points, and asserts the call sequence
// returns kOk and surfaces the expected payload.

#include <cstdint>
#include <cstring>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"

namespace {

class FormulonCApiSheetFeatures : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create(&wb_), 0); }
  void TearDown() override { fm_workbook_destroy(wb_); }
  fm_workbook_t* wb_ = nullptr;
};

TEST_F(FormulonCApiSheetFeatures, MergeAddAndIterate) {
  fm_merge_range m;
  m.first_row = 0;
  m.first_col = 0;
  m.last_row = 1;
  m.last_col = 2;
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, m), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_merge_range got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 0U);
  EXPECT_EQ(got.last_row, 1U);
  EXPECT_EQ(got.last_col, 2U);
}

TEST_F(FormulonCApiSheetFeatures, MergeOutOfRangeIndex) {
  fm_merge_range got{};
  EXPECT_NE(fm_sheet_get_merge_at(wb_, 0, 999, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkAddAndIterate) {
  fm_hyperlink hl{};
  hl.row = 1;
  hl.col = 2;
  hl.target = "https://example.com";
  hl.display = "Click";
  hl.tooltip = "Open site";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, hl), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink out{};
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &out), 0);
  EXPECT_EQ(out.row, 1U);
  EXPECT_EQ(out.col, 2U);
  ASSERT_NE(out.target, nullptr);
  EXPECT_STREQ(out.target, "https://example.com");
  EXPECT_STREQ(out.display, "Click");
  EXPECT_STREQ(out.tooltip, "Open site");
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkAcceptsNullStringFields) {
  fm_hyperlink hl{};
  hl.row = 0;
  hl.col = 0;
  // All string pointers null => empty target/display/etc.
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, hl), 0);
  fm_hyperlink got{};
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_STREQ(got.target, "");
}

TEST_F(FormulonCApiSheetFeatures, CommentSetGetAndRemove) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  fm_comment got{};
  ASSERT_EQ(fm_sheet_get_comment_at(wb_, 0, 1, 1, &got), 0);
  EXPECT_STREQ(got.author, "Alice");
  EXPECT_STREQ(got.text, "Hello");
  // Replace with new text at the same cell.
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Bob", "World"), 0);
  ASSERT_EQ(fm_sheet_get_comment_at(wb_, 0, 1, 1, &got), 0);
  EXPECT_STREQ(got.author, "Bob");
  EXPECT_STREQ(got.text, "World");
  // Empty text removes.
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Bob", ""), 0);
  EXPECT_NE(fm_sheet_get_comment_at(wb_, 0, 1, 1, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, CommentRemoveViaNullText) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, "Alice", "x"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, nullptr, nullptr), 0);
  fm_comment got{};
  EXPECT_NE(fm_sheet_get_comment_at(wb_, 0, 0, 0, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, SheetIndexOutOfRange) {
  fm_merge_range m{};
  EXPECT_NE(fm_sheet_add_merge(wb_, 999, m), 0);
  std::uint32_t count = 0;
  EXPECT_NE(fm_sheet_get_merge_count(wb_, 999, &count), 0);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeByRangeOverlap) {
  fm_merge_range a{0, 0, 1, 1};
  fm_merge_range b{0, 3, 1, 4};
  fm_merge_range c{4, 0, 5, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, c), 0);
  fm_merge_range q{0, 0, 1, 4};
  ASSERT_EQ(fm_sheet_remove_merge(wb_, 0, q), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_merge_range got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 4U);
  EXPECT_EQ(got.first_col, 0U);
  EXPECT_EQ(got.last_row, 5U);
  EXPECT_EQ(got.last_col, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeByRangeNoOverlap) {
  fm_merge_range a{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, a), 0);
  fm_merge_range q{4, 2, 5, 3};
  ASSERT_EQ(fm_sheet_remove_merge(wb_, 0, q), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeAtValid) {
  fm_merge_range a{0, 0, 1, 1};
  fm_merge_range b{4, 4, 5, 5};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_remove_merge_at(wb_, 0, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_merge_range got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 4U);
  EXPECT_EQ(got.first_col, 4U);
  EXPECT_EQ(got.last_row, 5U);
  EXPECT_EQ(got.last_col, 5U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeAtOutOfRange) {
  fm_merge_range a{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, a), 0);
  EXPECT_NE(fm_sheet_remove_merge_at(wb_, 0, 99), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ClearMerges) {
  fm_merge_range a{0, 0, 1, 1};
  fm_merge_range b{2, 2, 3, 3};
  fm_merge_range c{4, 4, 5, 5};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, c), 0);
  ASSERT_EQ(fm_sheet_clear_merges(wb_, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  // Calling clear again is still kOk.
  EXPECT_EQ(fm_sheet_clear_merges(wb_, 0), 0);
}

TEST_F(FormulonCApiSheetFeatures, MergeRemovalSheetIndexOutOfRange) {
  fm_merge_range q{};
  EXPECT_NE(fm_sheet_remove_merge(wb_, 999, q), 0);
  EXPECT_NE(fm_sheet_remove_merge_at(wb_, 999, 0), 0);
  EXPECT_NE(fm_sheet_clear_merges(wb_, 999), 0);
}

}  // namespace
