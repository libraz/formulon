// Copyright 2026 libraz. Licensed under the MIT License.
//
// End-to-end tests for the workbook-level structural mutation surface:
// `Workbook::rename_sheet`, `remove_sheet`, `move_sheet`, and
// `set_defined_name`. These exercise the public C++ API directly so the
// validation rules and defined-name rewriter can be observed without
// going through the C ABI.

#include <cstdint>
#include <string>

#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Sentinel sheet index well past `sheet_count()` for any test workbook
// in this file. Used to force the "out-of-range" rejection paths.
constexpr std::uint32_t kOutOfRangeIndex = 99U;

// Convenience factory: workbook with three sheets, named `Alpha`,
// `Beta`, `Gamma` (created via `add_sheet` so the default `Sheet1`
// stays out of the way).
Workbook ThreeSheetWorkbook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  wb.add_sheet("Gamma");
  return wb;
}

TEST(WorkbookSheetOps, RenameUpdatesSheetName) {
  Workbook wb = ThreeSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1, "Charlie")));
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Charlie");
  EXPECT_EQ(wb.sheet(2).name(), "Gamma");
}

TEST(WorkbookSheetOps, RenameRejectsEmptyName) {
  Workbook wb = ThreeSheetWorkbook();
  auto r = wb.rename_sheet(0, "");
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
}

TEST(WorkbookSheetOps, RenameRejectsTooLong) {
  Workbook wb = ThreeSheetWorkbook();
  // 32 chars: one over the limit.
  std::string long_name(32U, 'a');
  auto r = wb.rename_sheet(0, long_name);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
}

TEST(WorkbookSheetOps, RenameRejectsForbiddenCharacters) {
  Workbook wb = ThreeSheetWorkbook();
  for (const char* bad : {"a:b", "a\\b", "a/b", "a?b", "a*b", "a[b", "a]b"}) {
    auto r = wb.rename_sheet(0, bad);
    ASSERT_FALSE(static_cast<bool>(r)) << "expected rejection of " << bad;
    EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
  }
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
}

TEST(WorkbookSheetOps, RenameRejectsCaseInsensitiveCollision) {
  Workbook wb = ThreeSheetWorkbook();
  // `BETA` collides with the existing `Beta`.
  auto r = wb.rename_sheet(0, "BETA");
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
}

TEST(WorkbookSheetOps, RenameAcceptsCaseChangeOfOwnName) {
  Workbook wb = ThreeSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "ALPHA")));
  EXPECT_EQ(wb.sheet(0).name(), "ALPHA");
}

TEST(WorkbookSheetOps, RenameRejectsOutOfRange) {
  Workbook wb = ThreeSheetWorkbook();
  auto r = wb.rename_sheet(kOutOfRangeIndex, "Anything");
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kSheetIndexOutOfRange);
}

TEST(WorkbookSheetOps, RenameUpdatesWorkbookScopedDefinedNames) {
  Workbook wb = ThreeSheetWorkbook();
  // Pre-populate two defined names: one workbook-scoped that mentions
  // `Beta`, one sheet-scoped (must not be touched).
  std::vector<io::DefinedName> names;
  io::DefinedName wb_scoped;
  wb_scoped.name = "TotalArea";
  wb_scoped.formula = "Beta!$A$1:$A$10";
  wb_scoped.local_sheet_id = -1;
  names.push_back(wb_scoped);

  io::DefinedName sheet_scoped;
  sheet_scoped.name = "RegionA";
  sheet_scoped.formula = "Beta!$B$2";
  sheet_scoped.local_sheet_id = 2;
  names.push_back(sheet_scoped);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1, "Banana")));

  ASSERT_EQ(wb.defined_names().size(), 2U);
  EXPECT_EQ(wb.defined_names()[0].formula, "Banana!$A$1:$A$10");
  // Sheet-scoped name keeps its raw text — those refs are bound by index.
  EXPECT_EQ(wb.defined_names()[1].formula, "Beta!$B$2");
}

TEST(WorkbookSheetOps, RenameUpdatesQuotedSheetReferences) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("With Space");
  wb.add_sheet("Other");

  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "Q";
  dn.formula = "'With Space'!$A$1";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Plain")));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "Plain!$A$1");
}

TEST(WorkbookSheetOps, RenameToNameRequiringQuotesAddsQuotes) {
  // Renaming a bare-identifier sheet to a name with whitespace forces
  // the formatter to emit the canonical quoted form.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Beta");

  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "X";
  dn.formula = "Beta!$A$1";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "New Beta")));
  EXPECT_EQ(wb.defined_names()[0].formula, "'New Beta'!$A$1");
}

TEST(WorkbookSheetOps, RenameAcrossExpressionAndCallArgs) {
  // The AST-based rewriter must catch references that sit inside
  // function calls, range endpoints, and arithmetic — not just bare
  // sheet-prefixed identifiers.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Beta");
  wb.add_sheet("Other");

  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "Total";
  dn.formula = "SUM(Beta!$A$1:$A$10)+Beta!B2";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "Banana")));
  // Range endpoints with the same sheet collapse the sheet qualifier on
  // the right-hand side — Excel's canonical form for `Sheet!A1:A10`.
  EXPECT_EQ(wb.defined_names()[0].formula, "SUM(Banana!$A$1:$A$10)+Banana!B2");
}

TEST(WorkbookSheetOps, RenameDoesNotMatchIdentifierSubstrings) {
  // `OtherSheet1` should not match a rename of `Sheet1`.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("OtherSheet1");

  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "X";
  dn.formula = "OtherSheet1!$A$1";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  EXPECT_EQ(wb.defined_names()[0].formula, "OtherSheet1!$A$1");
}

TEST(WorkbookSheetOps, RemoveDropsSheet) {
  Workbook wb = ThreeSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Gamma");
}

TEST(WorkbookSheetOps, RemoveRejectsLastSheet) {
  Workbook wb = Workbook::create();  // single Sheet1
  auto r = wb.remove_sheet(0);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kCannotRemoveLastSheet);
  EXPECT_EQ(wb.sheet_count(), 1U);
}

TEST(WorkbookSheetOps, RemoveRejectsOutOfRange) {
  Workbook wb = ThreeSheetWorkbook();
  auto r = wb.remove_sheet(kOutOfRangeIndex);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kSheetIndexOutOfRange);
}

TEST(WorkbookSheetOps, RemoveDropsWorkbookScopedNamesTargetingRemovedSheet) {
  Workbook wb = ThreeSheetWorkbook();
  std::vector<io::DefinedName> names;
  io::DefinedName keep;
  keep.name = "Keeper";
  keep.formula = "Alpha!$A$1";
  keep.local_sheet_id = -1;
  names.push_back(keep);

  io::DefinedName drop;
  drop.name = "Goner";
  drop.formula = "Beta!$A$1";
  drop.local_sheet_id = -1;
  names.push_back(drop);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));

  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "Keeper");
}

TEST(WorkbookSheetOps, RemoveAdjustsSheetScopedLocalIds) {
  Workbook wb = ThreeSheetWorkbook();
  std::vector<io::DefinedName> names;
  io::DefinedName scoped_alpha;
  scoped_alpha.name = "A";
  scoped_alpha.formula = "$A$1";
  scoped_alpha.local_sheet_id = 0;
  names.push_back(scoped_alpha);

  io::DefinedName scoped_gamma;
  scoped_gamma.name = "G";
  scoped_gamma.formula = "$A$1";
  scoped_gamma.local_sheet_id = 2;
  names.push_back(scoped_gamma);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));  // remove Beta

  ASSERT_EQ(wb.defined_names().size(), 2U);
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, 0);  // Alpha still at 0
  EXPECT_EQ(wb.defined_names()[1].local_sheet_id, 1);  // Gamma shifted down
}

TEST(WorkbookSheetOps, MoveForwardRearrangesSheets) {
  Workbook wb = ThreeSheetWorkbook();
  // Move Alpha (0) to position 2 (the end). Excel UI semantics:
  // `to_index == 2`, not `3`.
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0, 2)));
  EXPECT_EQ(wb.sheet(0).name(), "Beta");
  EXPECT_EQ(wb.sheet(1).name(), "Gamma");
  EXPECT_EQ(wb.sheet(2).name(), "Alpha");
}

TEST(WorkbookSheetOps, MoveBackwardRearrangesSheets) {
  Workbook wb = ThreeSheetWorkbook();
  // Move Gamma (2) to position 0.
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 0)));
  EXPECT_EQ(wb.sheet(0).name(), "Gamma");
  EXPECT_EQ(wb.sheet(1).name(), "Alpha");
  EXPECT_EQ(wb.sheet(2).name(), "Beta");
}

TEST(WorkbookSheetOps, MoveNoOpAcceptsSamePosition) {
  Workbook wb = ThreeSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(1, 1)));
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Beta");
  EXPECT_EQ(wb.sheet(2).name(), "Gamma");
}

TEST(WorkbookSheetOps, MoveRejectsOutOfRange) {
  Workbook wb = ThreeSheetWorkbook();
  auto r1 = wb.move_sheet(kOutOfRangeIndex, 0);
  ASSERT_FALSE(static_cast<bool>(r1));
  EXPECT_EQ(r1.error().code, FormulonErrorCode::kSheetIndexOutOfRange);

  auto r2 = wb.move_sheet(0, kOutOfRangeIndex);
  ASSERT_FALSE(static_cast<bool>(r2));
  EXPECT_EQ(r2.error().code, FormulonErrorCode::kSheetIndexOutOfRange);
}

TEST(WorkbookSheetOps, MoveAdjustsSheetScopedLocalIds) {
  Workbook wb = ThreeSheetWorkbook();
  std::vector<io::DefinedName> names;
  io::DefinedName name_alpha;
  name_alpha.name = "A";
  name_alpha.formula = "$A$1";
  name_alpha.local_sheet_id = 0;
  io::DefinedName name_beta;
  name_beta.name = "B";
  name_beta.formula = "$A$1";
  name_beta.local_sheet_id = 1;
  io::DefinedName name_gamma;
  name_gamma.name = "C";
  name_gamma.formula = "$A$1";
  name_gamma.local_sheet_id = 2;
  names.push_back(name_alpha);
  names.push_back(name_beta);
  names.push_back(name_gamma);
  wb.set_defined_names(std::move(names));

  // Move sheet 0 (Alpha) to position 2: post-removal layout is
  // [Beta, Gamma, Alpha]. Alpha was scope 0 -> now 2; Beta was 1 -> 0;
  // Gamma was 2 -> 1.
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0, 2)));
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, 2);  // A (Alpha)
  EXPECT_EQ(wb.defined_names()[1].local_sheet_id, 0);  // B (Beta)
  EXPECT_EQ(wb.defined_names()[2].local_sheet_id, 1);  // C (Gamma)
}

TEST(WorkbookSheetOps, SetDefinedNameAddsEntry) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Pi", "=3.14159")));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "Pi");
  EXPECT_EQ(wb.defined_names()[0].formula, "=3.14159");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, -1);
}

TEST(WorkbookSheetOps, SetDefinedNameUpdatesExisting) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Pi", "=3.14")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("PI", "=3.14159")));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  // Authored casing of the original entry is preserved.
  EXPECT_EQ(wb.defined_names()[0].name, "Pi");
  EXPECT_EQ(wb.defined_names()[0].formula, "=3.14159");
}

TEST(WorkbookSheetOps, SetDefinedNameEmptyFormulaRemovesEntry) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Pi", "=3.14")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Pi", "")));
  EXPECT_TRUE(wb.defined_names().empty());
}

TEST(WorkbookSheetOps, SetDefinedNameEmptyFormulaOnMissingNameIsNoOp) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Missing", "")));
  EXPECT_TRUE(wb.defined_names().empty());
}

TEST(WorkbookSheetOps, SetDefinedNameRejectsEmptyName) {
  Workbook wb = Workbook::create();
  auto r = wb.set_defined_name("", "=1");
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument);
}

}  // namespace
}  // namespace formulon
