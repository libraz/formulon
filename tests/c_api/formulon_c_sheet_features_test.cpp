//
// C ABI smoke tests for the sheet UI feature surface (merges,
// hyperlinks, comments). Each test creates a default workbook,
// exercises the new entry points, and asserts the call sequence
// returns kOk and surfaces the expected payload.

#include <cstdint>
#include <cstring>

#include "c_api/formulon_c.h"
#include "gtest/gtest.h"
#include "utils/error.h"

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
  EXPECT_EQ(fm_sheet_get_comment_at(wb_, 0, 1, 1, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound));
}

TEST_F(FormulonCApiSheetFeatures, CommentLookupDistinguishesAbsenceFromInvalidSheet) {
  fm_comment got{};
  EXPECT_EQ(fm_sheet_get_comment_at(wb_, 0, 0, 0, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound));
  EXPECT_EQ(fm_sheet_get_comment_at(wb_, 99, 0, 0, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
}

TEST_F(FormulonCApiSheetFeatures, CommentRemoveViaNullText) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, "Alice", "x"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, nullptr, nullptr), 0);
  fm_comment got{};
  EXPECT_EQ(fm_sheet_get_comment_at(wb_, 0, 0, 0, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound));
}

TEST_F(FormulonCApiSheetFeatures, CommentEnumerateAllIncludingEmptyCell) {
  // Two comments: one anchored on a cell that also carries a value, one
  // anchored on a cell that never had a value written to it. Both must
  // be discoverable via the count/at-index enumerator without already
  // knowing their (row, col).
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 5, 3, "Bob", "Empty cell note"), 0);

  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_comment_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 2U);

  bool saw_alice = false;
  bool saw_bob = false;
  for (std::uint32_t i = 0; i < count; ++i) {
    fm_comment got{};
    ASSERT_EQ(fm_sheet_get_comment_at_index(wb_, 0, i, &got), 0);
    if (got.row == 1U && got.col == 1U) {
      EXPECT_STREQ(got.author, "Alice");
      EXPECT_STREQ(got.text, "Hello");
      saw_alice = true;
    } else if (got.row == 5U && got.col == 3U) {
      EXPECT_STREQ(got.author, "Bob");
      EXPECT_STREQ(got.text, "Empty cell note");
      saw_bob = true;
    }
  }
  EXPECT_TRUE(saw_alice);
  EXPECT_TRUE(saw_bob);
}

TEST_F(FormulonCApiSheetFeatures, CommentEnumerateCountZeroWhenNone) {
  std::uint32_t count = 123;
  ASSERT_EQ(fm_sheet_get_comment_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, CommentEnumerateOutOfRangeIndex) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, "Alice", "x"), 0);
  fm_comment got{};
  EXPECT_NE(fm_sheet_get_comment_at_index(wb_, 0, 999, &got), 0);
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

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkByRowCol) {
  fm_hyperlink a{};
  a.row = 1;
  a.col = 2;
  a.target = "https://a.example";
  fm_hyperlink b{};
  b.row = 3;
  b.col = 4;
  b.target = "https://b.example";
  fm_hyperlink c{};
  c.row = 1;
  c.col = 2;
  c.target = "https://dup.example";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, c), 0);
  // Removes both entries anchored at (1, 2); leaves the (3, 4) entry.
  ASSERT_EQ(fm_sheet_remove_hyperlink(wb_, 0, 1, 2), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink got{};
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.row, 3U);
  EXPECT_EQ(got.col, 4U);
  EXPECT_STREQ(got.target, "https://b.example");
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkByRowColNoMatch) {
  fm_hyperlink a{};
  a.row = 1;
  a.col = 2;
  a.target = "https://a.example";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, a), 0);
  // No hyperlink anchored at (5, 5); call is still kOk.
  ASSERT_EQ(fm_sheet_remove_hyperlink(wb_, 0, 5, 5), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkAtValid) {
  fm_hyperlink a{};
  a.row = 1;
  a.col = 1;
  a.target = "https://a.example";
  fm_hyperlink b{};
  b.row = 2;
  b.col = 2;
  b.target = "https://b.example";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb_, 0, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink got{};
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.row, 2U);
  EXPECT_EQ(got.col, 2U);
  EXPECT_STREQ(got.target, "https://b.example");
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkAtOutOfRange) {
  fm_hyperlink a{};
  a.row = 0;
  a.col = 0;
  a.target = "https://a.example";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, a), 0);
  EXPECT_NE(fm_sheet_remove_hyperlink_at(wb_, 0, 99), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ClearHyperlinks) {
  fm_hyperlink a{};
  a.row = 0;
  a.col = 0;
  a.target = "https://a.example";
  fm_hyperlink b{};
  b.row = 1;
  b.col = 1;
  b.target = "https://b.example";
  fm_hyperlink c{};
  c.row = 2;
  c.col = 2;
  c.target = "https://c.example";
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_add_hyperlink(wb_, 0, c), 0);
  ASSERT_EQ(fm_sheet_clear_hyperlinks(wb_, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  // Calling clear again is still kOk.
  EXPECT_EQ(fm_sheet_clear_hyperlinks(wb_, 0), 0);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkRemovalSheetIndexOutOfRange) {
  EXPECT_NE(fm_sheet_remove_hyperlink(wb_, 999, 0, 0), 0);
  EXPECT_NE(fm_sheet_remove_hyperlink_at(wb_, 999, 0), 0);
  EXPECT_NE(fm_sheet_clear_hyperlinks(wb_, 999), 0);
}

// ---- Data validations ------------------------------------------------------

TEST_F(FormulonCApiSheetFeatures, ValidationCountEmpty) {
  std::uint32_t count = 99U;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddSingleRangeAndIterate) {
  fm_merge_range r{1, 2, 3, 4};
  fm_data_validation v{};
  v.ranges = &r;
  v.range_count = 1;
  v.type = 3;         // list
  v.op = 2;           // equal
  v.error_style = 1;  // warning
  v.allow_blank = 1;
  v.show_input_message = 1;
  v.show_error_message = 0;
  v.formula1 = "$A$1:$A$10";
  v.formula2 = nullptr;
  v.error_title = "Bad";
  v.error_message = "Pick from the list";
  v.prompt_title = "Choose";
  v.prompt_message = "Select a value";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, v), 0);

  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);

  fm_data_validation got{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.range_count, 1U);
  ASSERT_NE(got.ranges, nullptr);
  EXPECT_EQ(got.ranges[0].first_row, 1U);
  EXPECT_EQ(got.ranges[0].first_col, 2U);
  EXPECT_EQ(got.ranges[0].last_row, 3U);
  EXPECT_EQ(got.ranges[0].last_col, 4U);
  EXPECT_EQ(got.type, 3);
  EXPECT_EQ(got.op, 2);
  EXPECT_EQ(got.error_style, 1);
  EXPECT_EQ(got.allow_blank, 1);
  EXPECT_EQ(got.show_input_message, 1);
  EXPECT_EQ(got.show_error_message, 0);
  ASSERT_NE(got.formula1, nullptr);
  EXPECT_STREQ(got.formula1, "$A$1:$A$10");
  ASSERT_NE(got.formula2, nullptr);
  EXPECT_STREQ(got.formula2, "");  // NULL on input surfaces as empty string.
  EXPECT_STREQ(got.error_title, "Bad");
  EXPECT_STREQ(got.error_message, "Pick from the list");
  EXPECT_STREQ(got.prompt_title, "Choose");
  EXPECT_STREQ(got.prompt_message, "Select a value");
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddMultiRangeRoundTrip) {
  fm_merge_range ranges[3] = {{0, 0, 1, 1}, {3, 3, 4, 4}, {7, 7, 9, 9}};
  fm_data_validation v{};
  v.ranges = ranges;
  v.range_count = 3;
  v.type = 2;  // decimal
  v.op = 0;    // between
  v.allow_blank = 0;
  v.formula1 = "0";
  v.formula2 = "100";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, v), 0);

  fm_data_validation got{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  ASSERT_EQ(got.range_count, 3U);
  ASSERT_NE(got.ranges, nullptr);
  EXPECT_EQ(got.ranges[1].first_row, 3U);
  EXPECT_EQ(got.ranges[1].last_col, 4U);
  EXPECT_EQ(got.ranges[2].first_row, 7U);
  EXPECT_EQ(got.ranges[2].last_row, 9U);
  EXPECT_STREQ(got.formula1, "0");
  EXPECT_STREQ(got.formula2, "100");
  EXPECT_EQ(got.allow_blank, 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationCoversEveryTypeAndOpEnum) {
  // Walk every type/op pair so the binding never silently drops a value.
  for (std::uint8_t type = 0; type <= 7; ++type) {
    for (std::uint8_t op = 0; op <= 7; ++op) {
      fm_data_validation v{};
      v.type = type;
      v.op = op;
      v.error_style = static_cast<std::uint8_t>(op % 3);
      v.allow_blank = (type & 1U) != 0U ? 1 : 0;
      ASSERT_EQ(fm_sheet_add_validation(wb_, 0, v), 0);
    }
  }
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 8U * 8U);
  // Probe a couple of entries for fidelity.
  fm_data_validation got{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.type, 0);
  EXPECT_EQ(got.op, 0);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 63, &got), 0);
  EXPECT_EQ(got.type, 7);
  EXPECT_EQ(got.op, 7);
}

TEST_F(FormulonCApiSheetFeatures, ValidationGetAtOutOfRange) {
  fm_data_validation got{};
  EXPECT_NE(fm_sheet_get_validation_at(wb_, 0, 99, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationRemoveAtValid) {
  fm_data_validation a{};
  a.type = 1;
  fm_data_validation b{};
  b.type = 2;
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_remove_validation_at(wb_, 0, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_data_validation got{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.type, 2);  // `b` survived.
}

TEST_F(FormulonCApiSheetFeatures, ValidationRemoveAtOutOfRange) {
  fm_data_validation a{};
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, a), 0);
  EXPECT_NE(fm_sheet_remove_validation_at(wb_, 0, 99), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationClear) {
  fm_data_validation a{};
  fm_data_validation b{};
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, a), 0);
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, b), 0);
  ASSERT_EQ(fm_sheet_clear_validations(wb_, 0), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  // Calling clear again is still kOk.
  EXPECT_EQ(fm_sheet_clear_validations(wb_, 0), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationSheetIndexOutOfRange) {
  fm_data_validation v{};
  std::uint32_t count = 0;
  EXPECT_NE(fm_sheet_get_validation_count(wb_, 999, &count), 0);
  fm_data_validation got{};
  EXPECT_NE(fm_sheet_get_validation_at(wb_, 999, 0, &got), 0);
  EXPECT_NE(fm_sheet_add_validation(wb_, 999, v), 0);
  EXPECT_NE(fm_sheet_remove_validation_at(wb_, 999, 0), 0);
  EXPECT_NE(fm_sheet_clear_validations(wb_, 999), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationBorrowedPointersStableAcrossReads) {
  fm_merge_range r{0, 0, 5, 5};
  fm_data_validation v{};
  v.ranges = &r;
  v.range_count = 1;
  v.type = 7;
  v.formula1 = "=A1>0";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, v), 0);
  fm_data_validation got_a{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got_a), 0);
  // Repeated reads with no intervening mutation must surface the same
  // borrowed pointers (the engine-side `std::string::c_str()` and
  // `std::vector::data()` are stable until the next mutation).
  fm_data_validation got_b{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got_b), 0);
  EXPECT_EQ(got_a.formula1, got_b.formula1);
  EXPECT_EQ(got_a.ranges, got_b.ranges);
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddNullStringFieldsBecomeEmpty) {
  fm_data_validation v{};
  v.type = 1;
  // Every string left as nullptr; getter must surface "" not nullptr.
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, v), 0);
  fm_data_validation got{};
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  ASSERT_NE(got.formula1, nullptr);
  EXPECT_STREQ(got.formula1, "");
  ASSERT_NE(got.formula2, nullptr);
  EXPECT_STREQ(got.formula2, "");
  ASSERT_NE(got.error_title, nullptr);
  EXPECT_STREQ(got.error_title, "");
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddRejectsNullRangesWithCount) {
  fm_data_validation v{};
  v.range_count = 3;
  v.ranges = nullptr;  // contradicts range_count.
  EXPECT_NE(fm_sheet_add_validation(wb_, 0, v), 0);
  std::uint32_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

}  // namespace
