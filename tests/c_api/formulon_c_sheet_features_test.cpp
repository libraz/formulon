//
// C ABI smoke tests for the sheet UI feature surface (merges,
// hyperlinks, comments). Each test creates a default workbook,
// exercises the new entry points, and asserts the call sequence
// returns kOk and surfaces the expected payload.

#include <cstdint>
#include <cstring>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_out_index.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "threaded_comment.h"
#include "utils/error.h"

namespace {

// The one-cell range a validation rule needs.
const fm_cell_rect_t kA1{0, 0, 0, 0};

class FormulonCApiSheetFeatures : public ::testing::Test {
 protected:
  void SetUp() override { ASSERT_EQ(fm_workbook_create_default(&wb_), 0); }
  void TearDown() override { fm_workbook_destroy(wb_); }
  fm_workbook_t* wb_ = nullptr;
};

TEST_F(FormulonCApiSheetFeatures, MergeAddAndIterate) {
  fm_cell_rect_t m;
  m.first_row = 0;
  m.first_col = 0;
  m.last_row = 1;
  m.last_col = 2;
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &m), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_cell_rect_t got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 0U);
  EXPECT_EQ(got.last_row, 1U);
  EXPECT_EQ(got.last_col, 2U);
}

TEST_F(FormulonCApiSheetFeatures, MergeOutOfRangeIndex) {
  fm_cell_rect_t got{};
  EXPECT_NE(fm_sheet_get_merge_at(wb_, 0, 999, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkAddAndIterate) {
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect.first_row = 1;
  hl.rect.first_col = 2;
  hl.rect.last_row = 1;
  hl.rect.last_col = 2;
  hl.target = "https://example.com";
  hl.display = "Click";
  hl.tooltip = "Open site";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 2, &hl), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &out), 0);
  EXPECT_EQ(out.rect.first_row, 1U);
  EXPECT_EQ(out.rect.first_col, 2U);
  EXPECT_EQ(out.rect.last_row, 1U);
  EXPECT_EQ(out.rect.last_col, 2U);
  ASSERT_NE(out.target, nullptr);
  EXPECT_STREQ(out.target, "https://example.com");
  EXPECT_STREQ(out.display, "Click");
  EXPECT_STREQ(out.tooltip, "Open site");
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkRangeAddAndIterate) {
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect.first_row = 7U;
  hl.rect.first_col = 8U;
  hl.rect.last_row = 9U;
  hl.rect.last_col = 10U;
  hl.target = "https://range.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0U, 7U, 8U, &hl), 0);

  fm_hyperlink_t out = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0U, 0U, &out), 0);
  EXPECT_EQ(out.rect.first_row, 7U);
  EXPECT_EQ(out.rect.first_col, 8U);
  EXPECT_EQ(out.rect.last_row, 9U);
  EXPECT_EQ(out.rect.last_col, 10U);
  EXPECT_STREQ(out.target, "https://range.example");
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkRejectsInvertedRectangleAtomically) {
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect.first_row = 9U;
  hl.rect.first_col = 8U;
  hl.rect.last_row = 7U;
  hl.rect.last_col = 10U;
  hl.target = "https://invalid.example";
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0U, 9U, 8U, &hl), invalid);

  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0U, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkNullStringFieldsReadBackNull) {
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect.first_row = 0;
  hl.rect.first_col = 0;
  hl.rect.last_row = 0;
  hl.rect.last_col = 0;
  // NULL and "" both mean absent; an absent string reads NULL.
  hl.display = "";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 0, 0, &hl), 0);
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.target, nullptr);
  EXPECT_EQ(got.location, nullptr);
  EXPECT_EQ(got.display, nullptr);
  EXPECT_EQ(got.tooltip, nullptr);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkSetAtAnchoredCellReplacesTheExistingLink) {
  fm_hyperlink_t first = FM_STRUCT_INIT(fm_hyperlink_t);
  first.rect = fm_cell_rect_t{2U, 3U, 2U, 3U};
  first.target = "https://first.example";
  fm_hyperlink_t second = FM_STRUCT_INIT(fm_hyperlink_t);
  second.rect = fm_cell_rect_t{2U, 3U, 4U, 5U};
  second.target = "https://second.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2U, 3U, &first), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2U, 3U, &second), 0);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  int32_t present = 0;
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb_, 0, 2U, 3U, &present, &got), 0);
  ASSERT_EQ(present, 1);
  EXPECT_STREQ(got.target, "https://second.example");
  EXPECT_EQ(got.rect.last_row, 4U);
  EXPECT_EQ(got.rect.last_col, 5U);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkSetWithNullRemovesTheLinkAnchoredThere) {
  fm_hyperlink_t keep = FM_STRUCT_INIT(fm_hyperlink_t);
  keep.rect = fm_cell_rect_t{0U, 0U, 0U, 0U};
  keep.target = "https://keep.example";
  fm_hyperlink_t drop = FM_STRUCT_INIT(fm_hyperlink_t);
  drop.rect = fm_cell_rect_t{5U, 5U, 5U, 6U};
  drop.target = "https://drop.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 0U, 0U, &keep), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 5U, 5U, &drop), 0);

  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0, 5U, 5U, nullptr), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  int32_t present = 1;
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink(wb_, 0, 5U, 5U, &present, &got), 0);
  EXPECT_EQ(present, 0);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkSetWithNullIsANoOpWhenAbsent) {
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0, 9U, 9U, nullptr), 0);
  std::size_t count = 9;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, HyperlinkRectangleMustBeAnchoredAtTheGivenCell) {
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  fm_hyperlink_t hl = FM_STRUCT_INIT(fm_hyperlink_t);
  hl.rect = fm_cell_rect_t{1U, 1U, 2U, 2U};
  hl.target = "https://anchor.example";
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1U, 2U, &hl), invalid);
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2U, 1U, &hl), invalid);
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2000000U, 1U, &hl), invalid);
  std::size_t count = 9;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 7, 1U, 1U, &hl),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(fm_sheet_set_hyperlink(nullptr, 0, 1U, 1U, &hl),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST_F(FormulonCApiSheetFeatures, CommentSetGetAndRemove) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 1, 1, &present, &got), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(got.author, "Alice");
  EXPECT_STREQ(got.text, "Hello");
  // Replace with new text at the same cell.
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Bob", "World"), 0);
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 1, 1, &present, &got), 0);
  EXPECT_STREQ(got.author, "Bob");
  EXPECT_STREQ(got.text, "World");
  // Removing deletes the note.
  ASSERT_EQ(fm_sheet_remove_comment(wb_, 0, 1, 1), 0);
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 1, 1, &present, &got), 0);
  EXPECT_EQ(present, 0);
}

// Excel keeps a note with no text, so an empty text writes one rather than
// removing the note; the setter never deletes.
TEST_F(FormulonCApiSheetFeatures, CommentEmptyTextWritesAnEmptyNote) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Bob", ""), 0);
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 1, 1, &present, &got), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(got.author, "Bob");
  EXPECT_STREQ(got.text, "");

  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 2, 2, nullptr, nullptr), 0);
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 2, 2, &present, &got), 0);
  EXPECT_EQ(present, 1);
  EXPECT_STREQ(got.author, "");
  EXPECT_STREQ(got.text, "");
}

TEST_F(FormulonCApiSheetFeatures, CommentRejectedOnThreadedCell) {
  formulon::Workbook& wb = wb_->workbook();
  const std::string person = "{11111111-AAAA-4BBB-8CCC-DDDDDDDDDDDD}";
  ASSERT_TRUE(static_cast<bool>(wb.add_person(formulon::Person{person, "Alice", "alice@example.com", "AD"})));
  formulon::ThreadedComment thread;
  thread.id = "{22222222-1111-4222-8333-444444444444}";
  thread.row = 2;
  thread.col = 3;
  thread.person_id = person;
  thread.created = "2026-10-05T10:00:00.00";
  thread.text = "Thread";
  ASSERT_TRUE(static_cast<bool>(wb.add_threaded_comment(0, thread)));
  EXPECT_EQ(fm_sheet_set_comment(wb_, 0, 2, 3, "Bob", "Note"),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists));
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 2, 3, &present, &got), 0);
  EXPECT_EQ(present, 0);
  // A neighbouring cell still takes a note.
  EXPECT_EQ(fm_sheet_set_comment(wb_, 0, 2, 4, "Bob", "Note"), 0);
}

TEST_F(FormulonCApiSheetFeatures, CommentReplaceKeepsRevisionUid) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  wb_->workbook().sheet(0).mutable_comments()[0].uid = "{AAAAAAAA-0000-4000-8000-000000000001}";
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Bob", "World"), 0);
  EXPECT_EQ(wb_->workbook().sheet(0).comments()[0].uid, "{AAAAAAAA-0000-4000-8000-000000000001}");
}

TEST_F(FormulonCApiSheetFeatures, CommentLookupDistinguishesAbsenceFromInvalidSheet) {
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 0, 0, &present, &got), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_get_comment(wb_, 99, 0, 0, &present, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(fm_sheet_get_comment(wb_, 0, 0, 0, nullptr, &got),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
}

TEST_F(FormulonCApiSheetFeatures, CommentRemoveReportsAMissingNote) {
  EXPECT_EQ(fm_sheet_remove_comment(wb_, 0, 0, 0), static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound));
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, "Alice", "x"), 0);
  ASSERT_EQ(fm_sheet_remove_comment(wb_, 0, 0, 0), 0);
  EXPECT_EQ(fm_sheet_remove_comment(wb_, 0, 0, 0), static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound));
  EXPECT_EQ(fm_sheet_remove_comment(wb_, 99, 0, 0),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kSheetIndexOutOfRange));
  EXPECT_EQ(fm_sheet_remove_comment(nullptr, 0, 0, 0),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer));
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_comment(wb_, 0, 0, 0, &present, &got), 0);
  EXPECT_EQ(present, 0);
}

TEST_F(FormulonCApiSheetFeatures, CommentEnumerateAllIncludingEmptyCell) {
  // Two comments: one anchored on a cell that also carries a value, one
  // anchored on a cell that never had a value written to it. Both must
  // be discoverable via the count/at-index enumerator without already
  // knowing their (row, col).
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 1, 1, "Alice", "Hello"), 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 5, 3, "Bob", "Empty cell note"), 0);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_comment_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 2U);

  bool saw_alice = false;
  bool saw_bob = false;
  for (std::uint32_t i = 0; i < count; ++i) {
    fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
    ASSERT_EQ(fm_sheet_get_comment_at(wb_, 0, i, &got), 0);
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
  std::size_t count = 123;
  ASSERT_EQ(fm_sheet_get_comment_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, CommentEnumerateOutOfRangeIndex) {
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 0, 0, "Alice", "x"), 0);
  fm_comment_info_t got = FM_STRUCT_INIT(fm_comment_info_t);
  EXPECT_NE(fm_sheet_get_comment_at(wb_, 0, 999, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, SheetIndexOutOfRange) {
  fm_cell_rect_t m{};
  EXPECT_NE(fm_sheet_add_merge(wb_, 999, &m), 0);
  std::size_t count = 0;
  EXPECT_NE(fm_sheet_get_merge_count(wb_, 999, &count), 0);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeByRangeOverlap) {
  fm_cell_rect_t a{0, 0, 1, 1};
  fm_cell_rect_t b{0, 3, 1, 4};
  fm_cell_rect_t c{4, 0, 5, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &b), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &c), 0);
  fm_cell_rect_t q{0, 0, 1, 4};
  ASSERT_EQ(fm_sheet_remove_merge(wb_, 0, &q), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_cell_rect_t got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 4U);
  EXPECT_EQ(got.first_col, 0U);
  EXPECT_EQ(got.last_row, 5U);
  EXPECT_EQ(got.last_col, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeByRangeNoOverlap) {
  fm_cell_rect_t a{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &a), 0);
  fm_cell_rect_t q{4, 2, 5, 3};
  ASSERT_EQ(fm_sheet_remove_merge(wb_, 0, &q), FM_STATUS_NOT_FOUND);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeAtValid) {
  fm_cell_rect_t a{0, 0, 1, 1};
  fm_cell_rect_t b{4, 4, 5, 5};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &b), 0);
  ASSERT_EQ(fm_sheet_remove_merge_at(wb_, 0, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_cell_rect_t got{};
  ASSERT_EQ(fm_sheet_get_merge_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.first_row, 4U);
  EXPECT_EQ(got.first_col, 4U);
  EXPECT_EQ(got.last_row, 5U);
  EXPECT_EQ(got.last_col, 5U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveMergeAtOutOfRange) {
  fm_cell_rect_t a{0, 0, 1, 1};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &a), 0);
  EXPECT_NE(fm_sheet_remove_merge_at(wb_, 0, 99), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ClearMerges) {
  fm_cell_rect_t a{0, 0, 1, 1};
  fm_cell_rect_t b{2, 2, 3, 3};
  fm_cell_rect_t c{4, 4, 5, 5};
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &a), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &b), 0);
  ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &c), 0);
  ASSERT_EQ(fm_sheet_clear_merges(wb_, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  // Calling clear again is still kOk.
  EXPECT_EQ(fm_sheet_clear_merges(wb_, 0), 0);
}

TEST_F(FormulonCApiSheetFeatures, MergeRemovalSheetIndexOutOfRange) {
  fm_cell_rect_t q{};
  EXPECT_NE(fm_sheet_remove_merge(wb_, 999, &q), 0);
  EXPECT_NE(fm_sheet_remove_merge_at(wb_, 999, 0), 0);
  EXPECT_NE(fm_sheet_clear_merges(wb_, 999), 0);
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkByRowCol) {
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  a.rect.first_row = 1;
  a.rect.first_col = 2;
  a.rect.last_row = 1;
  a.rect.last_col = 2;
  a.target = "https://a.example";
  fm_hyperlink_t b = FM_STRUCT_INIT(fm_hyperlink_t);
  b.rect.first_row = 3;
  b.rect.first_col = 4;
  b.rect.last_row = 3;
  b.rect.last_col = 4;
  b.target = "https://b.example";
  fm_hyperlink_t c = FM_STRUCT_INIT(fm_hyperlink_t);
  c.rect.first_row = 1;
  c.rect.first_col = 2;
  c.rect.last_row = 1;
  c.rect.last_col = 2;
  c.target = "https://dup.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 2, &a), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 3, 4, &b), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 2, &c), 0);
  // Removes the entry anchored at (1, 2); leaves the (3, 4) entry.
  ASSERT_EQ(fm_sheet_remove_hyperlink(wb_, 0, 1, 2), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.rect.first_row, 3U);
  EXPECT_EQ(got.rect.first_col, 4U);
  EXPECT_STREQ(got.target, "https://b.example");
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkByRowColNoMatch) {
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  a.rect.first_row = 1;
  a.rect.first_col = 2;
  a.rect.last_row = 1;
  a.rect.last_col = 2;
  a.target = "https://a.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 2, &a), 0);
  // No hyperlink anchored at (5, 5).
  ASSERT_EQ(fm_sheet_remove_hyperlink(wb_, 0, 5, 5), FM_STATUS_NOT_FOUND);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkAtValid) {
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  a.rect.first_row = 1;
  a.rect.first_col = 1;
  a.rect.last_row = 1;
  a.rect.last_col = 1;
  a.target = "https://a.example";
  fm_hyperlink_t b = FM_STRUCT_INIT(fm_hyperlink_t);
  b.rect.first_row = 2;
  b.rect.first_col = 2;
  b.rect.last_row = 2;
  b.rect.last_col = 2;
  b.target = "https://b.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 1, &a), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2, 2, &b), 0);
  ASSERT_EQ(fm_sheet_remove_hyperlink_at(wb_, 0, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_hyperlink_t got = FM_STRUCT_INIT(fm_hyperlink_t);
  ASSERT_EQ(fm_sheet_get_hyperlink_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.rect.first_row, 2U);
  EXPECT_EQ(got.rect.first_col, 2U);
  EXPECT_STREQ(got.target, "https://b.example");
}

TEST_F(FormulonCApiSheetFeatures, RemoveHyperlinkAtOutOfRange) {
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  a.rect.first_row = 0;
  a.rect.first_col = 0;
  a.rect.last_row = 0;
  a.rect.last_col = 0;
  a.target = "https://a.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 0, 0, &a), 0);
  EXPECT_NE(fm_sheet_remove_hyperlink_at(wb_, 0, 99), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ClearHyperlinks) {
  fm_hyperlink_t a = FM_STRUCT_INIT(fm_hyperlink_t);
  a.rect.first_row = 0;
  a.rect.first_col = 0;
  a.rect.last_row = 0;
  a.rect.last_col = 0;
  a.target = "https://a.example";
  fm_hyperlink_t b = FM_STRUCT_INIT(fm_hyperlink_t);
  b.rect.first_row = 1;
  b.rect.first_col = 1;
  b.rect.last_row = 1;
  b.rect.last_col = 1;
  b.target = "https://b.example";
  fm_hyperlink_t c = FM_STRUCT_INIT(fm_hyperlink_t);
  c.rect.first_row = 2;
  c.rect.first_col = 2;
  c.rect.last_row = 2;
  c.rect.last_col = 2;
  c.target = "https://c.example";
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 0, 0, &a), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 1, 1, &b), 0);
  ASSERT_EQ(fm_sheet_set_hyperlink(wb_, 0, 2, 2, &c), 0);
  ASSERT_EQ(fm_sheet_clear_hyperlinks(wb_, 0), 0);
  std::size_t count = 0;
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
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddSingleRangeAndIterate) {
  fm_cell_rect_t r{1, 2, 3, 4};
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = &r;
  v.range_count = 1;
  v.type = 3;         // list
  v.op = 2;           // equal
  v.error_style = 1;  // warning
  v.allow_blank = FM_TRISTATE_TRUE;
  v.show_input_message = FM_TRISTATE_TRUE;
  v.show_error_message = FM_TRISTATE_FALSE;
  v.formula1 = "$A$1:$A$10";
  v.formula2 = nullptr;
  v.error_title = "Bad";
  v.error_message = "Pick from the list";
  v.prompt_title = "Choose";
  v.prompt_message = "Select a value";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);

  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);

  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
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
  EXPECT_EQ(got.allow_blank, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_input_message, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_error_message, FM_TRISTATE_FALSE);
  ASSERT_NE(got.formula1, nullptr);
  EXPECT_STREQ(got.formula1, "=$A$1:$A$10");
  EXPECT_EQ(got.formula2, nullptr);  // NULL on input reads back NULL.
  EXPECT_STREQ(got.error_title, "Bad");
  EXPECT_STREQ(got.error_message, "Pick from the list");
  EXPECT_STREQ(got.prompt_title, "Choose");
  EXPECT_STREQ(got.prompt_message, "Select a value");
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddMultiRangeRoundTrip) {
  fm_cell_rect_t ranges[3] = {{0, 0, 1, 1}, {3, 3, 4, 4}, {7, 7, 9, 9}};
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = ranges;
  v.range_count = 3;
  v.type = 2;  // decimal
  v.op = 0;    // between
  v.allow_blank = FM_TRISTATE_FALSE;
  v.formula1 = "0";
  v.formula2 = "100";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);

  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  ASSERT_EQ(got.range_count, 3U);
  ASSERT_NE(got.ranges, nullptr);
  EXPECT_EQ(got.ranges[1].first_row, 3U);
  EXPECT_EQ(got.ranges[1].last_col, 4U);
  EXPECT_EQ(got.ranges[2].first_row, 7U);
  EXPECT_EQ(got.ranges[2].last_row, 9U);
  EXPECT_STREQ(got.formula1, "=0");
  EXPECT_STREQ(got.formula2, "=100");
  EXPECT_EQ(got.allow_blank, FM_TRISTATE_FALSE);
}

TEST_F(FormulonCApiSheetFeatures, ValidationCoversEveryTypeAndOpEnum) {
  // Walk every type/op pair so the binding never silently drops a value.
  for (std::uint8_t type = 0; type <= 7; ++type) {
    for (std::uint8_t op = 0; op <= 7; ++op) {
      fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
      v.ranges = &kA1;
      v.range_count = 1;
      v.type = type;
      v.op = op;
      v.error_style = static_cast<std::uint8_t>(op % 3);
      v.allow_blank = (type & 1U) != 0U ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
      ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);
    }
  }
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 8U * 8U);
  // Probe a couple of entries for fidelity.
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.type, 0);
  EXPECT_EQ(got.op, 0);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 63, &got), 0);
  EXPECT_EQ(got.type, 7);
  EXPECT_EQ(got.op, 7);
}

TEST_F(FormulonCApiSheetFeatures, ValidationRejectsEnumOrdinalsPastTheirDomain) {
  // The writer emits an absent attribute for an ordinal it does not
  // recognise, so a mistyped `type` would save as "no validation at all"
  // and the author would only notice on opening the file in Excel.
  const auto expect_rejected = [this](fm_data_validation_t v, const char* what) {
    EXPECT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()),
              static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument))
        << what;
  };

  fm_data_validation_t unknown_type = FM_STRUCT_INIT(fm_data_validation_t);
  unknown_type.type = 8;
  expect_rejected(unknown_type, "type");

  fm_data_validation_t unknown_op = FM_STRUCT_INIT(fm_data_validation_t);
  unknown_op.type = 1;
  unknown_op.op = 8;
  expect_rejected(unknown_op, "op");

  fm_data_validation_t unknown_error_style = FM_STRUCT_INIT(fm_data_validation_t);
  unknown_error_style.type = 1;
  unknown_error_style.error_style = 3;
  expect_rejected(unknown_error_style, "error_style");

  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationGetAtOutOfRange) {
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  EXPECT_NE(fm_sheet_get_validation_at(wb_, 0, 99, &got), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationRemoveByIdValid) {
  fm_data_validation_t a = FM_STRUCT_INIT(fm_data_validation_t);
  a.ranges = &kA1;
  a.range_count = 1;
  a.type = 1;
  fm_data_validation_t b = FM_STRUCT_INIT(fm_data_validation_t);
  b.ranges = &kA1;
  b.range_count = 1;
  b.type = 2;
  const char* added = nullptr;
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &a, &added), 0);
  const std::string first_id = added;
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &b, ignored_id()), 0);
  ASSERT_EQ(fm_sheet_remove_validation(wb_, 0, first_id.c_str()), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  ASSERT_EQ(count, 1U);
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.type, 2);  // `b` survived.
}

TEST_F(FormulonCApiSheetFeatures, ValidationRemoveByUnknownId) {
  fm_data_validation_t a = FM_STRUCT_INIT(fm_data_validation_t);
  a.ranges = &kA1;
  a.range_count = 1;
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &a, ignored_id()), 0);
  EXPECT_EQ(fm_sheet_remove_validation(wb_, 0, "99"), FM_STATUS_NOT_FOUND);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 1U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationClear) {
  fm_data_validation_t a = FM_STRUCT_INIT(fm_data_validation_t);
  a.ranges = &kA1;
  a.range_count = 1;
  fm_data_validation_t b = FM_STRUCT_INIT(fm_data_validation_t);
  b.ranges = &kA1;
  b.range_count = 1;
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &a, ignored_id()), 0);
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &b, ignored_id()), 0);
  ASSERT_EQ(fm_sheet_clear_validations(wb_, 0), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  // Calling clear again is still kOk.
  EXPECT_EQ(fm_sheet_clear_validations(wb_, 0), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationSheetIndexOutOfRange) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  std::size_t count = 0;
  EXPECT_NE(fm_sheet_get_validation_count(wb_, 999, &count), 0);
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  EXPECT_NE(fm_sheet_get_validation_at(wb_, 999, 0, &got), 0);
  EXPECT_NE(fm_sheet_add_validation(wb_, 999, &v, ignored_id()), 0);
  EXPECT_NE(fm_sheet_remove_validation(wb_, 999, "1"), 0);
  EXPECT_NE(fm_sheet_clear_validations(wb_, 999), 0);
}

TEST_F(FormulonCApiSheetFeatures, ValidationBorrowedPointersStableAcrossReads) {
  fm_cell_rect_t r{0, 0, 5, 5};
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = &r;
  v.range_count = 1;
  v.type = 7;
  v.formula1 = "=A1>0";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);
  fm_data_validation_t got_a = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got_a), 0);
  // Repeated reads with no intervening mutation must surface the same
  // model-backed pointers (the engine-side `std::string::c_str()` and
  // `std::vector::data()` are stable until the next mutation). The formulas
  // are read-scratch-backed, so only their text repeats.
  fm_data_validation_t got_b = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got_b), 0);
  EXPECT_EQ(got_a.ranges, got_b.ranges);
  EXPECT_EQ(got_a.error_title, got_b.error_title);
  EXPECT_STREQ(got_b.formula1, "=A1>0");
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddNullStringFieldsReadBackNull) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = &kA1;
  v.range_count = 1;
  v.type = 1;
  // Every string left NULL, or "", is absent and reads back NULL.
  v.prompt_title = "";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_EQ(got.formula1, nullptr);
  EXPECT_EQ(got.formula2, nullptr);
  EXPECT_EQ(got.error_title, nullptr);
  EXPECT_EQ(got.error_message, nullptr);
  EXPECT_EQ(got.prompt_title, nullptr);
  EXPECT_EQ(got.prompt_message, nullptr);
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddRejectsAnEmptyRangeList) {
  // A rule applies to cells, so it needs at least one range, as a CF rule does.
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.type = 1;
  EXPECT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), FM_STATUS_INVALID_ARGUMENT);
  std::size_t count = 9;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, ValidationAddRejectsNullRangesWithCount) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.range_count = 3;
  v.ranges = nullptr;  // contradicts range_count.
  EXPECT_NE(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST_F(FormulonCApiSheetFeatures, MutatorsRejectOutOfGridCoordinatesBeforeMutation) {
  const fm_status_t invalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);

  fm_hyperlink_t hyperlink = FM_STRUCT_INIT(fm_hyperlink_t);
  hyperlink.rect.first_row = 0U;
  hyperlink.rect.first_col = formulon::Sheet::kMaxCols;
  hyperlink.rect.last_row = 0U;
  hyperlink.rect.last_col = formulon::Sheet::kMaxCols;
  EXPECT_EQ(fm_sheet_set_hyperlink(wb_, 0U, 0U, formulon::Sheet::kMaxCols, &hyperlink), invalid);
  std::size_t count = 99U;
  ASSERT_EQ(fm_sheet_get_hyperlink_count(wb_, 0U, &count), 0);
  EXPECT_EQ(count, 0U);

  fm_cell_rect_t merge{0U, 0U, 0U, formulon::Sheet::kMaxCols};
  EXPECT_EQ(fm_sheet_add_merge(wb_, 0U, &merge), invalid);
  ASSERT_EQ(fm_sheet_get_merge_count(wb_, 0U, &count), 0);
  EXPECT_EQ(count, 0U);

  EXPECT_EQ(fm_sheet_set_comment(wb_, 0U, 0U, formulon::Sheet::kMaxCols, "Alice", "invalid"), invalid);
  ASSERT_EQ(fm_sheet_get_comment_count(wb_, 0U, &count), 0);
  EXPECT_EQ(count, 0U);

  const fm_cell_rect_t validation_ranges[2] = {
      {0U, 0U, 0U, 0U},
      {0U, 0U, 0U, formulon::Sheet::kMaxCols},
  };
  fm_data_validation_t validation = FM_STRUCT_INIT(fm_data_validation_t);
  validation.ranges = validation_ranges;
  validation.range_count = 2U;
  EXPECT_EQ(fm_sheet_add_validation(wb_, 0U, &validation, ignored_id()), invalid);
  ASSERT_EQ(fm_sheet_get_validation_count(wb_, 0U, &count), 0);
  EXPECT_EQ(count, 0U);

  // Rejected metadata must not leave the workbook in a state that makes the
  // writer abort. A normal save is the process-survival regression check.
  std::uint8_t* saved = nullptr;
  std::size_t saved_len = 0U;
  ASSERT_EQ(fm_workbook_save(wb_, &saved, &saved_len), 0);
  ASSERT_NE(saved, nullptr);
  EXPECT_GT(saved_len, 0U);
  fm_buffer_destroy(saved);
}

// DV formulas are stored as files carry them: one leading '=' is dropped, so
// a save / load cycle returns the text the setter stored.
TEST_F(FormulonCApiSheetFeatures, ValidationStoresFormulasWithoutTheLeadingEquals) {
  fm_cell_rect_t r{0, 0, 5, 0};
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = &r;
  v.range_count = 1;
  v.type = 1;  // whole
  v.op = 0;    // between
  v.formula1 = "=$B$1";
  v.formula2 = "=$B$1+10";
  ASSERT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()), 0) << fm_get_last_error_message();
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), 0);
  EXPECT_STREQ(got.formula1, "=$B$1");
  EXPECT_STREQ(got.formula2, "=$B$1+10");

  v.formula1 = "==$B$1";
  EXPECT_EQ(fm_sheet_add_validation(wb_, 0, &v, ignored_id()),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kParserUnexpectedToken));

  std::uint8_t* saved = nullptr;
  std::size_t saved_len = 0U;
  ASSERT_EQ(fm_workbook_save(wb_, &saved, &saved_len), 0);
  fm_workbook_t* reloaded = nullptr;
  const fm_status_t loaded = fm_workbook_load(saved, saved_len, &reloaded);
  fm_buffer_destroy(saved);
  ASSERT_EQ(loaded, 0);
  ASSERT_EQ(fm_sheet_get_validation_at(reloaded, 0, 0, &got), 0);
  EXPECT_STREQ(got.formula1, "=$B$1");
  EXPECT_STREQ(got.formula2, "=$B$1+10");
  fm_workbook_destroy(reloaded);
}

}  // namespace

namespace {

TEST_F(FormulonCApiSheetFeatures, ValidationIndexPastTheCountNamesTheParameter) {
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  EXPECT_EQ(fm_sheet_get_validation_at(wb_, 0, 0, &got), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_NE(std::string(fm_get_last_error_context()).find("validation_index=0"), std::string::npos)
      << fm_get_last_error_context();
}

TEST_F(FormulonCApiSheetFeatures, SheetProtectionAbsentStringsReadNull) {
  fm_sheet_protection_t in = FM_STRUCT_INIT(fm_sheet_protection_t);
  in.hash_value = "";
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb_, 0, &in), 0);
  int32_t present = 0;
  fm_sheet_protection_t out = FM_STRUCT_INIT(fm_sheet_protection_t);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb_, 0, &present, &out), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.algorithm_name, nullptr);
  EXPECT_EQ(out.hash_value, nullptr);
  EXPECT_EQ(out.salt_value, nullptr);
  EXPECT_EQ(out.legacy_password_hash, nullptr);

  in.legacy_password_hash = "CC1A";
  ASSERT_EQ(fm_sheet_set_sheet_protection(wb_, 0, &in), 0);
  ASSERT_EQ(fm_sheet_get_sheet_protection(wb_, 0, &present, &out), 0);
  EXPECT_STREQ(out.legacy_password_hash, "CC1A");
}

TEST_F(FormulonCApiSheetFeatures, CommentVisibleWithoutANoteIsAbsentNotAnError) {
  int32_t present = -1;
  int32_t visible = -1;
  ASSERT_EQ(fm_sheet_get_comment_visible(wb_, 0, 2, 2, &present, &visible), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(visible, 0);
  ASSERT_EQ(fm_sheet_set_comment(wb_, 0, 2, 2, "Alice", "note"), 0);
  ASSERT_EQ(fm_sheet_set_comment_visible(wb_, 0, 2, 2, 1), 0);
  ASSERT_EQ(fm_sheet_get_comment_visible(wb_, 0, 2, 2, &present, &visible), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(visible, 1);
}

// A caller buffer smaller than the match count receives `capacity` entries
// and the full count; capacity 0 with a NULL buffer sizes it.
TEST_F(FormulonCApiSheetFeatures, MergesInRangeReportsTheFullCountPastTheCapacity) {
  const fm_cell_rect_t merges[] = {{0, 0, 0, 1}, {2, 0, 2, 1}, {4, 0, 4, 1}};
  for (const fm_cell_rect_t& m : merges) {
    ASSERT_EQ(fm_sheet_add_merge(wb_, 0, &m), 0);
  }
  const fm_cell_rect_t query{0, 0, 9, 9};
  std::size_t count = 0;
  ASSERT_EQ(fm_sheet_get_merges_in_range(wb_, 0, &query, nullptr, 0, &count), 0);
  EXPECT_EQ(count, 3U);
  fm_cell_rect_t out[2]{};
  ASSERT_EQ(fm_sheet_get_merges_in_range(wb_, 0, &query, out, 2, &count), 0);
  EXPECT_EQ(count, 3U);
  EXPECT_EQ(out[0].first_row, 0U);
  EXPECT_EQ(out[1].first_row, 2U);
}

}  // namespace
