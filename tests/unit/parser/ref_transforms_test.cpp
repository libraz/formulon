// Copyright 2026 libraz. Licensed under the MIT License.
//
// Unit tests for `SheetRenameTransform` and the `sheet_name_needs_quoting`
// helper. The walker integration is exercised in `ast_shift_test.cpp`;
// these tests focus on the per-Reference contract.

#include "parser/ref_transforms.h"

#include <optional>
#include <string_view>

#include "gtest/gtest.h"
#include "parser/reference.h"

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

}  // namespace
}  // namespace parser
}  // namespace formulon
