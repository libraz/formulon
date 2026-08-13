//
// Unit tests for `SheetRenameTransform` and the `sheet_name_needs_quoting`
// helper. The walker integration is exercised in `ast_shift_test.cpp`;
// these tests focus on the per-Reference contract.

#include "parser/ref_transforms.h"

#include <optional>
#include <string_view>
#include <vector>

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

TEST(SheetNameNeedsQuoting, NumericAndCellReferenceNamesRequireQuotes) {
  EXPECT_TRUE(sheet_name_needs_quoting("2026"));
  EXPECT_TRUE(sheet_name_needs_quoting("S2"));
  EXPECT_FALSE(sheet_name_needs_quoting("S0"));
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

TEST(RowColShiftTransform, RowDeleteShrinksRangeAtFirstEndpoint) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/0, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 0, 0), MakeRef("Sheet1", 2, 0));
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->first.row, 0U);
  EXPECT_EQ(out->second.row, 1U);
}

TEST(RowColShiftTransform, RowDeleteShrinksRangeAtLastEndpoint) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/2, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 0, 0), MakeRef("Sheet1", 2, 0));
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->first.row, 0U);
  EXPECT_EQ(out->second.row, 1U);
}

TEST(RowColShiftTransform, RowDeleteShrinksRangeThroughMiddle) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/1, /*count=*/2);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 0, 0), MakeRef("Sheet1", 4, 0));
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->first.row, 0U);
  EXPECT_EQ(out->second.row, 2U);
}

TEST(RowColShiftTransform, RowDeleteCollapsesSingletonRange) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/2, /*count=*/1);
  EXPECT_FALSE(t.apply_range(MakeRef("Sheet1", 2, 0), MakeRef("Sheet1", 2, 0)).has_value());
}

TEST(RowColShiftTransform, RowDeletePreservesReverseRangeOrientation) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/1, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 3, 0), MakeRef("Sheet1", 0, 0));
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->first.row, 2U);
  EXPECT_EQ(out->second.row, 0U);
}

TEST(RowColShiftTransform, ColDeleteShrinksRangeAtBothEndpoints) {
  RowColShiftTransform t("Sheet1", RowColAxis::kCol, RowColEdit::kDelete, /*index=*/1, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 0, 0), MakeRef("Sheet1", 0, 3));
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->first.col, 0U);
  EXPECT_EQ(out->second.col, 2U);
}

TEST(RowColShiftTransform, ColDeleteCollapsesSingletonRange) {
  RowColShiftTransform t("Sheet1", RowColAxis::kCol, RowColEdit::kDelete, /*index=*/2, /*count=*/1);
  EXPECT_FALSE(t.apply_range(MakeRef("Sheet1", 0, 2), MakeRef("Sheet1", 0, 2)).has_value());
}

TEST(RowColShiftTransform, WholeRowAndColumnRangesUseTheirRelevantAxis) {
  RowColShiftTransform row_delete("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/1, /*count=*/1);
  Reference row_first = MakeRef("Sheet1", 0, 0);
  Reference row_last = MakeRef("Sheet1", 3, 0);
  row_first.is_full_row = true;
  row_last.is_full_row = true;
  const std::optional<std::pair<Reference, Reference>> rows = row_delete.apply_range(row_first, row_last);
  ASSERT_TRUE(rows.has_value());
  EXPECT_EQ(rows->first.row, 0U);
  EXPECT_EQ(rows->second.row, 2U);

  RowColShiftTransform col_delete("Sheet1", RowColAxis::kCol, RowColEdit::kDelete, /*index=*/1, /*count=*/1);
  Reference col_first = MakeRef("Sheet1", 0, 0);
  Reference col_last = MakeRef("Sheet1", 0, 3);
  col_first.is_full_col = true;
  col_last.is_full_col = true;
  const std::optional<std::pair<Reference, Reference>> cols = col_delete.apply_range(col_first, col_last);
  ASSERT_TRUE(cols.has_value());
  EXPECT_EQ(cols->first.col, 0U);
  EXPECT_EQ(cols->second.col, 2U);
}

TEST(RowColShiftTransform, OutOfScopeOrCrossSheetRangeUsesEndpointPolicy) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/1, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out_of_scope =
      t.apply_range(MakeRef("Other", 1, 0), MakeRef("Other", 3, 0));
  ASSERT_TRUE(out_of_scope.has_value());
  EXPECT_EQ(out_of_scope->first.row, 1U);
  EXPECT_EQ(out_of_scope->second.row, 3U);

  const std::optional<std::pair<Reference, Reference>> cross_sheet =
      t.apply_range(MakeRef("Sheet1", 1, 0), MakeRef("Other", 3, 0));
  EXPECT_FALSE(cross_sheet.has_value()) << "an in-scope deleted endpoint still poisons a cross-sheet range";
}

TEST(RowColShiftTransform, RangeDeletePreservesAbsoluteFlags) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, /*index=*/0, /*count=*/1);
  const std::optional<std::pair<Reference, Reference>> out =
      t.apply_range(MakeRef("Sheet1", 0, 0, true, true), MakeRef("Sheet1", 2, 0, true, true));
  ASSERT_TRUE(out.has_value());
  EXPECT_TRUE(out->first.col_abs);
  EXPECT_TRUE(out->first.row_abs);
  EXPECT_TRUE(out->second.col_abs);
  EXPECT_TRUE(out->second.row_abs);
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

// ---------------------------------------------------------------------------
// SheetRemovalTransform
// ---------------------------------------------------------------------------

TEST(SheetRemovalTransform, RemovesOnlyTheQualifiedTarget) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  SheetRemovalTransform t(order, /*removed_index=*/1);

  EXPECT_FALSE(t.apply(MakeRef("Beta", 0, 0)).has_value());
  EXPECT_FALSE(t.apply(MakeRef("bEtA", 0, 0)).has_value());
  EXPECT_TRUE(t.apply(MakeRef("Alpha", 0, 0)).has_value());
  EXPECT_TRUE(t.apply(MakeRef("", 0, 0)).has_value());
}

TEST(SheetRemovalTransform, MiddleSpanKeepsEndpoints) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  SheetRemovalTransform t(order, /*removed_index=*/1);
  const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Alpha", "Gamma");
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->begin, "Alpha");
  EXPECT_EQ(out->end, "Gamma");
}

TEST(SheetRemovalTransform, EndpointSpanMovesInward) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  {
    SheetRemovalTransform t(order, /*removed_index=*/0);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Alpha", "Gamma");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Beta");
    EXPECT_EQ(out->end, "Gamma");
  }
  {
    SheetRemovalTransform t(order, /*removed_index=*/2);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Alpha", "Gamma");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Alpha");
    EXPECT_EQ(out->end, "Beta");
  }
}

TEST(SheetRemovalTransform, ReverseSpanMovesInwardInReverseDirection) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  {
    SheetRemovalTransform t(order, /*removed_index=*/2);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Gamma", "Alpha");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Beta");
    EXPECT_EQ(out->end, "Alpha");
  }
  {
    SheetRemovalTransform t(order, /*removed_index=*/0);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Gamma", "Alpha");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Gamma");
    EXPECT_EQ(out->end, "Beta");
  }
}

TEST(SheetRemovalTransform, DegenerateAndUnresolvedSpansCollapse) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  SheetRemovalTransform remove_beta(order, /*removed_index=*/1);
  EXPECT_FALSE(remove_beta.apply_ref3d_span("Beta", "Beta").has_value());
  const std::optional<RefTransform::Ref3DSheetSpan> unaffected = remove_beta.apply_ref3d_span("Missing", "Gamma");
  ASSERT_TRUE(unaffected.has_value());
  EXPECT_EQ(unaffected->begin, "Missing");
  EXPECT_EQ(unaffected->end, "Gamma");
  EXPECT_FALSE(remove_beta.apply_ref3d_span("Beta", "Missing").has_value());
}

}  // namespace
}  // namespace parser
}  // namespace formulon
