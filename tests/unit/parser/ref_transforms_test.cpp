//
// Unit tests for `SheetRenameTransform` and the `local_sheet_needs_quoting_a1`
// helper. The walker integration is exercised in `ast_shift_test.cpp`;
// these tests focus on the per-Reference contract.

#include "parser/ref_transforms.h"

#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "parser/reference.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "utils/a1_ref.h"

namespace formulon {
namespace parser {
namespace {

TEST(LocalSheetNeedsQuotingA1, EmptyRequiresQuotes) {
  EXPECT_TRUE(local_sheet_needs_quoting_a1(""));
}

TEST(LocalSheetNeedsQuotingA1, SimpleAsciiNoQuotes) {
  EXPECT_FALSE(local_sheet_needs_quoting_a1("Sheet1"));
}

TEST(LocalSheetNeedsQuotingA1, NumericAndCellReferenceNamesRequireQuotes) {
  EXPECT_TRUE(local_sheet_needs_quoting_a1("2026"));
  EXPECT_TRUE(local_sheet_needs_quoting_a1("S2"));
  EXPECT_FALSE(local_sheet_needs_quoting_a1("S0"));
}

TEST(LocalSheetNeedsQuotingA1, DigitLeadingNamesRequireQuotes) {
  // Excel accepts `2024!A1` on entry but always displays such a name quoted.
  EXPECT_TRUE(local_sheet_needs_quoting_a1("3S1"));
  EXPECT_TRUE(local_sheet_needs_quoting_a1("1abc"));
  EXPECT_TRUE(local_sheet_needs_quoting_a1("0"));
}

TEST(LocalSheetNeedsQuotingA1, UnderscoreAndDotAllowed) {
  EXPECT_FALSE(local_sheet_needs_quoting_a1("My_Sheet.1"));
}

TEST(LocalSheetNeedsQuotingA1, SpaceTriggersQuoting) {
  EXPECT_TRUE(local_sheet_needs_quoting_a1("Sheet 1"));
}

TEST(LocalSheetNeedsQuotingA1, HyphenTriggersQuoting) {
  EXPECT_TRUE(local_sheet_needs_quoting_a1("Sheet-1"));
}

TEST(LocalSheetNeedsQuotingA1, NonAsciiLettersStayBare) {
  // Excel displays a katakana sheet name bare; only U+3000 / U+FEFF among
  // non-ASCII characters force quoting.
  EXPECT_FALSE(local_sheet_needs_quoting_a1("\xe3\x82\xb7\xe3\x83\xbc\xe3\x83\x88"));  // "シート"
  EXPECT_TRUE(local_sheet_needs_quoting_a1("\xe3\x82\xb7\xe3\x80\x80"));               // "シ　"
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

TEST(SheetRemovalTransform, RebindsOnlyListedNamesOnTheRemovedSheet) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  const SheetRemovalTransform t(order, /*removed_index=*/1, {"LN"});

  EXPECT_TRUE(t.rebinds_name_to_book("Beta", "LN"));
  EXPECT_TRUE(t.rebinds_name_to_book("bEtA", "ln"));
  EXPECT_FALSE(t.rebinds_name_to_book("Beta", "Lonely"));
  EXPECT_FALSE(t.rebinds_name_to_book("Alpha", "LN"));
  EXPECT_FALSE(SheetRemovalTransform(order, 1).rebinds_name_to_book("Beta", "LN"));
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

// A span that moves inward onto its other endpoint names one sheet, and is
// spelled as that sheet's single reference.
TEST(SheetRemovalTransform, SpanShrunkToOneSheetCollapses) {
  const std::vector<std::string_view> order = {"Alpha", "Beta", "Gamma"};
  {
    SheetRemovalTransform t(order, /*removed_index=*/2);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Beta", "Gamma");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Beta");
    EXPECT_TRUE(out->collapse);
  }
  {
    SheetRemovalTransform t(order, /*removed_index=*/1);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Beta", "Gamma");
    ASSERT_TRUE(out.has_value());
    EXPECT_EQ(out->begin, "Gamma");
    EXPECT_TRUE(out->collapse);
  }
  {
    SheetRemovalTransform t(order, /*removed_index=*/2);
    const std::optional<RefTransform::Ref3DSheetSpan> out = t.apply_ref3d_span("Alpha", "Gamma");
    ASSERT_TRUE(out.has_value());
    EXPECT_FALSE(out->collapse);
  }
}

TEST(SheetRemovalTransform, CollapseFormOfEachStructuralTransform) {
  const std::vector<std::string_view> order = {"Alpha", "Beta"};
  EXPECT_EQ(RowColShiftTransform("Alpha", RowColAxis::kRow, RowColEdit::kDelete, 0, 1).collapse_form(),
            RefCollapse::kKeepSheet);
  EXPECT_EQ(SheetRemovalTransform(order, 1).collapse_form(), RefCollapse::kKeepAddress);
  EXPECT_EQ(SheetCopyTransform("Alpha", "Alpha (2)", {}).collapse_form(), RefCollapse::kKeepAddress);
  EXPECT_EQ(SheetRenameTransform("Alpha", "Beta").collapse_form(), RefCollapse::kBare);
  EXPECT_EQ(IdentityTransform().collapse_form(), RefCollapse::kBare);
}

namespace {

// The Insert / Delete Cells edits on B3:C4: rows 3:4 (0-based 2..3) along the
// row axis with band columns B:C, or columns B:C along the column axis with
// band rows 3:4.
constexpr std::uint32_t kEdgeRow = 2U;
constexpr std::uint32_t kEdgeCol = 1U;
constexpr std::uint32_t kSpan = 2U;
constexpr Band kColumnsBC{1U, 2U};
constexpr Band kRows34{2U, 3U};

RowColShiftTransform Down() {
  return RowColShiftTransform("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, kEdgeRow, kSpan, kColumnsBC);
}
RowColShiftTransform Up() {
  return RowColShiftTransform("Sheet1", RowColAxis::kRow, RowColEdit::kDelete, kEdgeRow, kSpan, kColumnsBC);
}
RowColShiftTransform Right() {
  return RowColShiftTransform("Sheet1", RowColAxis::kCol, RowColEdit::kInsert, kEdgeCol, kSpan, kRows34);
}
RowColShiftTransform Left() {
  return RowColShiftTransform("Sheet1", RowColAxis::kCol, RowColEdit::kDelete, kEdgeCol, kSpan, kRows34);
}

Reference Cell(std::string_view a1) {
  Reference r;
  r.sheet = "Sheet1";
  EXPECT_TRUE(a1::parse_a1_ref(a1, &r.row, &r.col)) << a1;
  return r;
}

// "#REF!" for a collapsed reference, else the A1 text.
std::string Text(const std::optional<Reference>& ref) {
  return ref.has_value() ? a1::encode_a1(ref->row, ref->col) : std::string("#REF!");
}

// Applies `t` to the range `first:last` and spells the result as A1 text.
std::string Range(const RowColShiftTransform& t, std::string_view first, std::string_view last) {
  const std::optional<std::pair<Reference, Reference>> out = t.apply_range(Cell(first), Cell(last));
  if (!out.has_value()) {
    return "#REF!";
  }
  return a1::encode_a1(out->first.row, out->first.col) + ":" + a1::encode_a1(out->second.row, out->second.col);
}

}  // namespace

TEST(RowColShiftTransformBand, ACellInsideTheBandFollowsTheEdit) {
  EXPECT_EQ(Text(Down().apply(Cell("B6"))), "B8");
  EXPECT_EQ(Text(Up().apply(Cell("B6"))), "B4");
  EXPECT_EQ(Text(Right().apply(Cell("D3"))), "F3");
  EXPECT_EQ(Text(Left().apply(Cell("D3"))), "B3");
}

TEST(RowColShiftTransformBand, ACellOutsideTheBandIsUntouched) {
  EXPECT_EQ(Text(Down().apply(Cell("D3"))), "D3");
  EXPECT_EQ(Text(Up().apply(Cell("D3"))), "D3");
  EXPECT_EQ(Text(Down().apply(Cell("A6"))), "A6");
  EXPECT_EQ(Text(Right().apply(Cell("B6"))), "B6");
  EXPECT_EQ(Text(Left().apply(Cell("B6"))), "B6");
  EXPECT_EQ(Text(Left().apply(Cell("D5"))), "D5");
}

TEST(RowColShiftTransformBand, ADeletedCellInsideTheBandIsARefError) {
  EXPECT_EQ(Text(Up().apply(Cell("B3"))), "#REF!");
  EXPECT_EQ(Text(Up().apply(Cell("C4"))), "#REF!");
  EXPECT_EQ(Text(Left().apply(Cell("B3"))), "#REF!");
  EXPECT_EQ(Text(Left().apply(Cell("C4"))), "#REF!");
  EXPECT_EQ(Text(Down().apply(Cell("B3"))), "B5");
  EXPECT_EQ(Text(Right().apply(Cell("C4"))), "E4");
}

TEST(RowColShiftTransformBand, ARangeInsideTheBandMovesGrowsOrShrinks) {
  EXPECT_EQ(Range(Down(), "B6", "C6"), "B8:C8");
  EXPECT_EQ(Range(Up(), "B6", "C6"), "B4:C4");
  EXPECT_EQ(Range(Down(), "B2", "C5"), "B2:C7");
  EXPECT_EQ(Range(Up(), "B2", "C5"), "B2:C3");
  EXPECT_EQ(Range(Down(), "B3", "B10"), "B5:B12");
  EXPECT_EQ(Range(Up(), "B3", "B10"), "B3:B8");
  EXPECT_EQ(Range(Down(), "B1", "C10"), "B1:C12");
  EXPECT_EQ(Range(Up(), "B1", "C10"), "B1:C8");
  EXPECT_EQ(Range(Right(), "D3", "E4"), "F3:G4");
  EXPECT_EQ(Range(Left(), "D3", "E4"), "B3:C4");
  EXPECT_EQ(Range(Right(), "A3", "D3"), "A3:F3");
  EXPECT_EQ(Range(Left(), "A3", "D3"), "A3:B3");
  EXPECT_EQ(Range(Down(), "B3", "C4"), "B5:C6");
  EXPECT_EQ(Range(Right(), "B3", "C4"), "D3:E4");
  EXPECT_EQ(Range(Right(), "A3", "C4"), "A3:E4");
  EXPECT_EQ(Range(Right(), "B3", "D4"), "D3:F4");
  EXPECT_EQ(Range(Left(), "B3", "D4"), "B3:B4");
}

TEST(RowColShiftTransformBand, AWhollyDeletedRangeIsARefError) {
  EXPECT_EQ(Range(Up(), "B3", "C4"), "#REF!");
  EXPECT_EQ(Range(Left(), "B3", "C4"), "#REF!");
}

TEST(RowColShiftTransformBand, ARangeStraddlingTheBandIsUntouched) {
  EXPECT_EQ(Range(Down(), "A6", "C6"), "A6:C6");
  EXPECT_EQ(Range(Up(), "A6", "C6"), "A6:C6");
  EXPECT_EQ(Range(Down(), "D3", "E5"), "D3:E5");
  EXPECT_EQ(Range(Right(), "D3", "E5"), "D3:E5");
  EXPECT_EQ(Range(Left(), "D3", "E5"), "D3:E5");
  // Insert does not grow a straddler across its edge.
  EXPECT_EQ(Range(Down(), "A1", "C10"), "A1:C10");
  EXPECT_EQ(Range(Down(), "A3", "C4"), "A3:C4");
  EXPECT_EQ(Range(Down(), "B3", "D4"), "B3:D4");
}

TEST(RowColShiftTransformBand, WholeColumnsAndRowsAreUntouched) {
  Reference col_first = Cell("B1");
  Reference col_last = Cell("B1048576");
  col_first.is_full_col = true;
  col_last.is_full_col = true;
  Reference row_first = Cell("A3");
  Reference row_last = Cell("XFD3");
  row_first.is_full_row = true;
  row_last.is_full_row = true;
  const RowColShiftTransform edits[] = {Down(), Up(), Right(), Left()};
  for (const RowColShiftTransform& t : edits) {
    const auto cols = t.apply_range(col_first, col_last);
    ASSERT_TRUE(cols.has_value());
    EXPECT_EQ(cols->first.col, 1U);
    EXPECT_EQ(cols->second.col, 1U);
    EXPECT_EQ(cols->first.row, 0U);
    EXPECT_EQ(cols->second.row, Sheet::kMaxRows - 1U);
    const auto rows = t.apply_range(row_first, row_last);
    ASSERT_TRUE(rows.has_value());
    EXPECT_EQ(rows->first.row, 2U);
    EXPECT_EQ(rows->second.row, 2U);
    const auto row_cell = t.apply(row_first);
    ASSERT_TRUE(row_cell.has_value());
    EXPECT_EQ(row_cell->row, 2U);
    const auto col_cell = t.apply(col_first);
    ASSERT_TRUE(col_cell.has_value());
    EXPECT_EQ(col_cell->col, 1U);
  }
}

TEST(RowColShiftTransformBand, DeleteShrinksARangeTheBandTouchesAtOneEnd) {
  EXPECT_EQ(Range(Up(), "A3", "C4"), "A3:A4");
  EXPECT_EQ(Range(Up(), "B3", "D4"), "D3:D4");
  EXPECT_EQ(Range(Left(), "B3", "B10"), "B5:B10");
}

TEST(RowColShiftTransformBand, DeleteLeavesARangeTheBandSitsInsideOf) {
  EXPECT_EQ(Range(Left(), "B2", "C5"), "B2:C5");
  EXPECT_EQ(Range(Up(), "A3", "D3"), "A3:D3");
  EXPECT_EQ(Range(Left(), "B1", "C10"), "B1:C10");
}

TEST(RowColShiftTransformBand, DeleteShrinksARangeTheBandOverhangs) {
  // Band B:C overhangs the right end of A3:B3; that is the one-end case.
  EXPECT_EQ(Range(Up(), "A3", "B3"), "A3:A3");
}

TEST(RowColShiftTransformBand, AnotherSheetIsUntouched) {
  RowColShiftTransform t("Sheet1", RowColAxis::kRow, RowColEdit::kInsert, kEdgeRow, kSpan, kColumnsBC);
  Reference other = Cell("B6");
  other.sheet = "Other";
  const auto out = t.apply(other);
  ASSERT_TRUE(out.has_value());
  EXPECT_EQ(out->row, 5U);
}

}  // namespace
}  // namespace parser
}  // namespace formulon
