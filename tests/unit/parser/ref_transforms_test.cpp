// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Unit tests for `SheetRenameTransform` and the `sheet_name_needs_quoting`
// helper. The walker integration is exercised in `ast_shift_test.cpp`;
// these tests focus on the per-Reference contract.

#include "parser/ref_transforms.h"

#include <optional>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/reference.h"
#include "sheet.h"

namespace formulon {
namespace parser {
namespace {

TEST(SheetNameNeedsQuoting, EmptyRequiresQuotes) {
  EXPECT_TRUE(sheet_name_needs_quoting(""));
}

TEST(SheetNameNeedsQuoting, SimpleAsciiNoQuotes) {
  EXPECT_FALSE(sheet_name_needs_quoting("Sheet1"));
}

TEST(SheetNameNeedsQuoting, UnderscoreAndDotAllowed) {
  EXPECT_FALSE(sheet_name_needs_quoting("My_Sheet.1"));
}

TEST(SheetNameNeedsQuoting, SpaceTriggersQuoting) {
  EXPECT_TRUE(sheet_name_needs_quoting("Sheet 1"));
}

TEST(SheetNameNeedsQuoting, HyphenTriggersQuoting) {
  EXPECT_TRUE(sheet_name_needs_quoting("Sheet-1"));
}

TEST(SheetNameNeedsQuoting, UnicodeBytesTriggerQuoting) {
  // Hiragana sheet name (UTF-8) lives outside the ASCII ident ruleset; we
  // conservatively require quoting.
  EXPECT_TRUE(sheet_name_needs_quoting("\xe3\x82\xb7\xe3\x83\xbc\xe3\x83\x88"));  // "シート"
}

// ---------------------------------------------------------------------------
// SheetRenameTransform
// ---------------------------------------------------------------------------

TEST(SheetRenameTransform, MatchesAndRenames) {
  SheetRenameTransform t("Sheet1", "Renamed");
  Reference r;
  r.sheet = "Sheet1";
  r.col = 0;
  r.row = 0;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->sheet, "Renamed");
  EXPECT_FALSE(out->sheet_quoted);
  EXPECT_EQ(out->col, 0u);
  EXPECT_EQ(out->row, 0u);
}

TEST(SheetRenameTransform, CaseInsensitiveMatch) {
  SheetRenameTransform t("Sheet1", "Renamed");
  Reference r;
  r.sheet = "SHEET1";
  r.col = 1;
  r.row = 1;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->sheet, "Renamed");
}

TEST(SheetRenameTransform, NonMatchingSheetIsPassThrough) {
  SheetRenameTransform t("Sheet1", "Renamed");
  Reference r;
  r.sheet = "Sheet2";
  r.col = 0;
  r.row = 0;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->sheet, "Sheet2");
  EXPECT_EQ(out->sheet_quoted, false);
}

TEST(SheetRenameTransform, EmptySheetIsPassThrough) {
  SheetRenameTransform t("Sheet1", "Renamed");
  Reference r;
  r.col = 0;
  r.row = 0;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_TRUE(out->sheet.empty());
}

TEST(SheetRenameTransform, RenameToQuoted) {
  SheetRenameTransform t("Sheet1", "New Sheet");
  Reference r;
  r.sheet = "Sheet1";
  r.col = 0;
  r.row = 0;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->sheet, "New Sheet");
  EXPECT_TRUE(out->sheet_quoted);
}

TEST(SheetRenameTransform, RenameFromQuotedToBareDropsQuoting) {
  SheetRenameTransform t("My Sheet", "Bare");
  Reference r;
  r.sheet = "My Sheet";
  r.sheet_quoted = true;
  r.col = 0;
  r.row = 0;
  std::optional<Reference> out = t.apply(r);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->sheet, "Bare");
  EXPECT_FALSE(out->sheet_quoted);
}

TEST(SheetRenameTransform, ExternalSheetHookMatches) {
  SheetRenameTransform t("Sheet1", "Renamed");
  std::optional<std::string_view> remapped = t.transform_external_sheet(/*book_id=*/1, "Sheet1");
  ASSERT_TRUE(remapped.has_value());
  EXPECT_EQ(*remapped, "Renamed");
}

TEST(SheetRenameTransform, ExternalSheetHookDoesNotMatch) {
  SheetRenameTransform t("Sheet1", "Renamed");
  std::optional<std::string_view> remapped = t.transform_external_sheet(/*book_id=*/1, "Other");
  EXPECT_FALSE(remapped.has_value());
}

TEST(SheetRenameTransform, RemapSheetHelper) {
  SheetRenameTransform t("Sheet1", "Renamed");
  EXPECT_FALSE(t.remap_sheet("").has_value());
  EXPECT_FALSE(t.remap_sheet("Other").has_value());
  std::optional<std::string_view> match = t.remap_sheet("Sheet1");
  ASSERT_TRUE(match.has_value());
  EXPECT_EQ(*match, "Renamed");
}

namespace {
Reference MakeRef(std::string_view sheet, std::uint32_t row, std::uint32_t col, bool col_abs = false,
                  bool row_abs = false) {
  Reference r;
  r.sheet = sheet;
  r.row = row;
  r.col = col;
  r.col_abs = col_abs;
  r.row_abs = row_abs;
  return r;
}
}  // namespace

TEST(RowColShiftTransform, RowInsertShiftsForward) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/4, /*count=*/2);
  // Row 0 (A1) is below the insert; unchanged.
  std::optional<Reference> below = t.apply(MakeRef("Sheet1", 0, 0));
  ASSERT_TRUE(below.has_value());
  EXPECT_EQ(below->row, 0U);
  // Row 4 (the insert origin) shifts forward by 2.
  std::optional<Reference> at = t.apply(MakeRef("Sheet1", 4, 0));
  ASSERT_TRUE(at.has_value());
  EXPECT_EQ(at->row, 6U);
  // Row 10 also shifts forward.
  std::optional<Reference> past = t.apply(MakeRef("Sheet1", 10, 0));
  ASSERT_TRUE(past.has_value());
  EXPECT_EQ(past->row, 12U);
}

TEST(RowColShiftTransform, RowInsertOverflowCollapsesToRef) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kInsert,
                         /*index=*/0, /*count=*/Sheet::kMaxRows - 1U);
  std::optional<Reference> overflowed = t.apply(MakeRef("Sheet1", 1, 0));
  EXPECT_FALSE(overflowed.has_value());
}

TEST(RowColShiftTransform, RowDeleteCollapsesInsideInterval) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/4, /*count=*/3);
  // Row 5 inside [4,7) collapses.
  EXPECT_FALSE(t.apply(MakeRef("Sheet1", 5, 0)).has_value());
  // Row 7 (just past the deletion) shifts to 4.
  std::optional<Reference> trailing = t.apply(MakeRef("Sheet1", 7, 0));
  ASSERT_TRUE(trailing.has_value());
  EXPECT_EQ(trailing->row, 4U);
  // Row 2 (before the deletion) is unchanged.
  std::optional<Reference> before = t.apply(MakeRef("Sheet1", 2, 0));
  ASSERT_TRUE(before.has_value());
  EXPECT_EQ(before->row, 2U);
}

TEST(RowColShiftTransform, ColInsertSkipsFullRowReferences) {
  RowColShiftTransform t("Sheet1", RowColAxis::kCol, RowColEdit::kInsert, /*index=*/2, /*count=*/3);
  Reference full_row = MakeRef("Sheet1", 5, 0);
  full_row.is_full_row = true;
  std::optional<Reference> result = t.apply(full_row);
  ASSERT_TRUE(result.has_value());
  // Full-row references ignore the column-axis edit entirely.
  EXPECT_EQ(result->col, 0U);
}

TEST(RowColShiftTransform, RowEditPreservesAbsoluteFlags) {
  // Excel shifts absolute references on insert / delete just like
  // relative ones — the dollar marker only matters during fill / paste,
  // not during structural edits. The transform mirrors that.
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/0, /*count=*/2);
  Reference abs = MakeRef("Sheet1", 5, 1, /*col_abs=*/true, /*row_abs=*/true);
  std::optional<Reference> shifted = t.apply(abs);
  ASSERT_TRUE(shifted.has_value());
  EXPECT_EQ(shifted->row, 7U);
  EXPECT_TRUE(shifted->row_abs);
  EXPECT_TRUE(shifted->col_abs);
}

TEST(RowColShiftTransform, OutOfScopeSheetIsUntouched) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/4, /*count=*/2);
  std::optional<Reference> other = t.apply(MakeRef("Other", 5, 0));
  ASSERT_TRUE(other.has_value());
  EXPECT_EQ(other->row, 5U);
}

TEST(RowColShiftTransform, LocalMeansTargetGatesUnqualifiedRefs) {
  // Default: local refs (sheet="") are out of scope unless the caller
  // sets `local_means_target=true`. The flag exists because the per-
  // Reference walker cannot know which sheet owns the formula it is
  // walking; the caller toggles the flag per sheet.
  RowColShiftTransform t_off("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/0, /*count=*/3);
  std::optional<Reference> off = t_off.apply(MakeRef("", 5, 0));
  ASSERT_TRUE(off.has_value());
  EXPECT_EQ(off->row, 5U);

  RowColShiftTransform t_on("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, /*index=*/0, /*count=*/3,
                            /*local_means_target=*/true);
  std::optional<Reference> on = t_on.apply(MakeRef("", 5, 0));
  ASSERT_TRUE(on.has_value());
  EXPECT_EQ(on->row, 8U);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
