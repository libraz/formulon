// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end tests for the workbook-level structural mutation surface:
// `Workbook::rename_sheet`, `remove_sheet`, `move_sheet`, and
// `set_defined_name`. These exercise the public C++ API directly so the
// validation rules and defined-name rewriter can be observed without
// going through the C ABI.

#include <cstdint>
#include <string>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
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

// Removing a middle sheet renumbers every later sheet's workbook-relative
// index. The dependency graph is keyed by that index, so a cross-sheet
// formula's edge must be re-pointed at the survivor's new position;
// otherwise a later edit to the survivor never dirties the dependent and
// the stale cached value persists.
TEST(WorkbookSheetOps, RemoveMiddleSheetRepointsCrossSheetDependencies) {
  Workbook wb = ThreeSheetWorkbook();                                                // Alpha(0), Beta(1), Gamma(2)
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(2, 0, 0, Value::number(100.0))));  // Gamma!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=Gamma!A1")));         // Alpha!A1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 100.0);

  // Remove Beta (index 1); Gamma slides to index 1.
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  EXPECT_EQ(wb.sheet(1).name(), "Gamma");
  // Recalc after the structural change re-reads Gamma at its new index.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 100.0);

  // Edit Gamma (now index 1); the dependent Alpha!A1 must re-evaluate.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(300.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 300.0)
      << "cross-sheet dependent stale after removing a preceding sheet";
}

// Moving a sheet renumbers the whole [min..max] window. A cross-sheet
// dependency edge that spans the window must survive the renumbering so
// downstream edits still propagate.
TEST(WorkbookSheetOps, MoveSheetRepointsCrossSheetDependencies) {
  Workbook wb = ThreeSheetWorkbook();                                               // Alpha(0), Beta(1), Gamma(2)
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(2, 0, 0, Value::number(10.0))));  // Gamma!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=Gamma!A1")));        // Alpha!A1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 10.0);

  // Move Alpha (0) to the end: layout becomes [Beta, Gamma, Alpha].
  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(0, 2)));
  EXPECT_EQ(wb.sheet(1).name(), "Gamma");
  EXPECT_EQ(wb.sheet(2).name(), "Alpha");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 10.0);

  // Edit Gamma (now index 1); Alpha (now index 2) must re-evaluate.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 42.0)
      << "cross-sheet dependent stale after moving a sheet";
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

// ---------------------------------------------------------------------------
// Row / column insert and delete
// ---------------------------------------------------------------------------

TEST(WorkbookRowColEdits, InsertRowsShiftsCellsAndFormulas) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 5, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 0, "=A1+A6")));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  // Cells at row 0 stayed put; cell at row 5 moved to row 7; the formula
  // cell at row 6 moved to row 8 and its A6 ref shifted to A8.
  const Sheet& s = wb.sheet(0);
  ASSERT_NE(s.cell_at(0, 0), nullptr);
  EXPECT_EQ(s.cell_at(0, 0)->cached_value.as_number(), 10.0);
  EXPECT_EQ(s.cell_at(5, 0), nullptr);
  ASSERT_NE(s.cell_at(7, 0), nullptr);
  EXPECT_EQ(s.cell_at(7, 0)->cached_value.as_number(), 20.0);
  ASSERT_NE(s.cell_at(8, 0), nullptr);
  EXPECT_EQ(s.cell_at(8, 0)->formula_text, "=A1+A8");
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesReferencesInsideInterval) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=A5")));     // A5 lives inside the deletion.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 1, "=A8+A2")));  // A8 trails the deletion; A2 is below.

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/3, /*count=*/3)));

  const Sheet& s = wb.sheet(0);
  // A5 falls inside [3, 6); collapses to #REF!.
  EXPECT_EQ(s.cell_at(0, 0)->formula_text, "=#REF!");
  // A8 -> A5 (shifted up by 3); A2 unchanged.
  EXPECT_EQ(s.cell_at(0, 1)->formula_text, "=A5+A2");
}

TEST(WorkbookRowColEdits, InsertColsShiftsCellsAcrossRow) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 3, Value::number(4.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1, 0, "=A1+D1")));

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/1, /*count=*/2)));

  const Sheet& s = wb.sheet(0);
  EXPECT_EQ(s.cell_at(0, 0)->cached_value.as_number(), 1.0);
  // Slot at (0, 3) exists in the row vector but is a default Cell — the
  // value that used to live there migrated forward.
  ASSERT_NE(s.cell_at(0, 3), nullptr);
  EXPECT_TRUE(s.cell_at(0, 3)->cached_value.is_blank());
  ASSERT_NE(s.cell_at(0, 5), nullptr);
  EXPECT_EQ(s.cell_at(0, 5)->cached_value.as_number(), 4.0);
  // The formula cell sat at (1, 0), which is below the insert origin
  // (col 1) and so does not move. Its references shift: D1 -> F1.
  ASSERT_NE(s.cell_at(1, 0), nullptr);
  EXPECT_EQ(s.cell_at(1, 0)->formula_text, "=A1+F1");
}

TEST(WorkbookRowColEdits, DeleteColsRewritesFormulaAndDropsCell) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 5, "=B1+C1+E1")));

  // Delete cols B..C (col indices 1..2).
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/2)));

  const Sheet& s = wb.sheet(0);
  // Source cell moved from F1 (col 5) to D1 (col 3).
  ASSERT_NE(s.cell_at(0, 3), nullptr);
  // B1 and C1 are inside the deleted span: collapse the entire formula
  // to #REF! because every range / sum endpoint that lands inside the
  // deletion poisons its enclosing expression.
  EXPECT_EQ(s.cell_at(0, 3)->formula_text, "=#REF!+#REF!+C1");
}

TEST(WorkbookRowColEdits, InsertRowsShiftsMergeRanges) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {MergeRange{0, 0, 0, 1}, MergeRange{5, 0, 7, 1}};

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  ASSERT_EQ(s.merges().size(), 2U);
  // First merge sits entirely below the insert; unchanged.
  EXPECT_EQ(s.merges()[0].first_row, 0U);
  EXPECT_EQ(s.merges()[0].last_row, 0U);
  // Second merge shifted forward by 2.
  EXPECT_EQ(s.merges()[1].first_row, 7U);
  EXPECT_EQ(s.merges()[1].last_row, 9U);
}

TEST(WorkbookRowColEdits, DeleteRowsClampsStraddlingMerge) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.mutable_merges() = {MergeRange{2, 0, 8, 1}};

  // Delete rows 4..6.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/4, /*count=*/3)));

  ASSERT_EQ(s.merges().size(), 1U);
  EXPECT_EQ(s.merges()[0].first_row, 2U);
  // Last row was 8 (past the deletion); shifts up by 3 to 5.
  EXPECT_EQ(s.merges()[0].last_row, 5U);
}

TEST(WorkbookRowColEdits, InsertRowsRewritesDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "Region";
  dn.formula = "Sheet1!$A$5:$A$10";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/3, /*count=*/2)));

  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$7:$A$12");
}

TEST(WorkbookRowColEdits, InsertRowsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  // Items: qty (col 1), unit (col 2), subtotal=qty*unit (col 3), tax=sub*0.08 (col 4).
  // Row 0 reserved for headers (left empty).
  // Item rows 1..5.
  const double qty[] = {24, 30, 4, 12, 5};
  const double unit[] = {0.42, 2.5, 4.45, 2.0, 0.75};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 1;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 1, Value::number(qty[i]))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 2, Value::number(unit[i]))));
    // Excel-text uses 1-based addressing; the formula at cell row=i+1 references row=i+1 (1-based).
    const std::string subtotal_formula = "=B" + std::to_string(r + 1) + "*C" + std::to_string(r + 1);
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 3, subtotal_formula)));
    const std::string tax_formula = "=D" + std::to_string(r + 1) + "*0.08";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 4, tax_formula)));
  }
  // Totals at row 7 (1-based row 8). =SUM(D2:D6) covers the entire item range.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 3, "=SUM(D2:D6)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 4, "=SUM(E2:E6)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 8, 3, "=D8+E8")));

  // Initial recalc establishes the baseline.
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  {
    const Sheet& s = wb.sheet(0);
    EXPECT_EQ(s.cell_at(5, 3)->cached_value.as_number(), 5.0 * 0.75);  // eraser subtotal pre-shift
    EXPECT_EQ(s.cell_at(7, 3)->cached_value.as_number(), 24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75);
  }

  // Insert one blank row at row index 1, push every item down by 1.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/1, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  // Items now at rows 2..6, totals at row 8, with-tax at row 9.
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 2;
    ASSERT_NE(s.cell_at(r, 3), nullptr) << "subtotal cell missing at row " << r;
    ASSERT_NE(s.cell_at(r, 4), nullptr) << "tax cell missing at row " << r;
    EXPECT_EQ(s.cell_at(r, 3)->cached_value.as_number(), qty[i] * unit[i])
        << "subtotal mismatch at post-shift row " << r;
    EXPECT_EQ(s.cell_at(r, 4)->cached_value.as_number(), qty[i] * unit[i] * 0.08)
        << "tax mismatch at post-shift row " << r;
  }
  ASSERT_NE(s.cell_at(8, 3), nullptr);
  EXPECT_EQ(s.cell_at(8, 3)->cached_value.as_number(), 24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75)
      << "post-shift SUM mismatch (should NOT be #REF!)";
  ASSERT_NE(s.cell_at(8, 4), nullptr);
  EXPECT_NEAR(s.cell_at(8, 4)->cached_value.as_number(), (24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75) * 0.08,
              1e-9);
  ASSERT_NE(s.cell_at(9, 3), nullptr);
  EXPECT_NEAR(s.cell_at(9, 3)->cached_value.as_number(), (24 * 0.42 + 30 * 2.5 + 4 * 4.45 + 12 * 2.0 + 5 * 0.75) * 1.08,
              1e-9);
}

// Mirror of the row regression on the column axis. Items live in cols B..F,
// SUM is in col H. After insertCols(1, 1) the SUM range shifts to C3..G3
// and every item must keep a fresh value, especially the last shifted col.
TEST(WorkbookRowColEdits, InsertColsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  // Layout (1-based): row 3 has qty/unit per item; row 4 has the
  // subtotal formulas; H4 = SUM(B4:F4).
  const double qty[] = {2, 3, 4, 5, 6};
  const double unit[] = {1.5, 2.5, 3.5, 4.5, 5.5};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t c = i + 1;  // cols 1..5 (B..F in 1-based)
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, c, Value::number(qty[i]))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, c, Value::number(unit[i]))));
    const std::string col_letter(1, static_cast<char>('A' + c));
    const std::string sub_formula = "=" + col_letter + "3*" + col_letter + "4";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4, c, sub_formula)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4, 7, "=SUM(B5:F5)")));

  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  double expected_sum = 0.0;
  for (std::uint32_t i = 0; i < 5; ++i) {
    expected_sum += qty[i] * unit[i];
  }
  EXPECT_EQ(wb.sheet(0).cell_at(4, 7)->cached_value.as_number(), expected_sum);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/1, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t c = i + 2;  // cols 2..6 post-shift
    ASSERT_NE(s.cell_at(4, c), nullptr);
    EXPECT_EQ(s.cell_at(4, c)->cached_value.as_number(), qty[i] * unit[i])
        << "post-shift subtotal mismatch at col " << c;
  }
  ASSERT_NE(s.cell_at(4, 8), nullptr);
  EXPECT_EQ(s.cell_at(4, 8)->cached_value.as_number(), expected_sum) << "post-shift SUM stale or #REF!";
}

// Delete rows must drop dep-graph entries for cells inside the deletion
// band AND re-key entries for cells trailing the band so dependents recalc
// correctly. Without the fix the SUM that spans the deletion would read
// stale (pre-delete) cached values from the now-empty rows.
TEST(WorkbookRowColEdits, DeleteRowsRecomputesShiftedFormulasAndAggregates) {
  Workbook wb = Workbook::create();
  const double qty[] = {2, 3, 4, 5, 6};
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t r = i + 1;
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, r, 1, Value::number(qty[i]))));
    const std::string subtotal_formula = "=B" + std::to_string(r + 1) + "*2";
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, r, 3, subtotal_formula)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 7, 3, "=SUM(D2:D6)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  double expected_pre = 0.0;
  for (std::uint32_t i = 0; i < 5; ++i) {
    expected_pre += qty[i] * 2;
  }
  EXPECT_EQ(wb.sheet(0).cell_at(7, 3)->cached_value.as_number(), expected_pre);

  // Delete rows 2..3 (0-based), removing the second and third items.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/2)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Sheet& s = wb.sheet(0);
  ASSERT_NE(s.cell_at(5, 3), nullptr);  // SUM moved from row 7 → 5
  // SUM range was D2:D6 covering rows 1..5; after deleting rows 2..3
  // the range refs collapse: D2 (row 1, untouched) + D5..D6 (rows 4..5
  // pre-delete, now rows 2..3) survive; rows 2..3 (qty[1], qty[2]) drop.
  const double expected_post = qty[0] * 2 + qty[3] * 2 + qty[4] * 2;
  EXPECT_EQ(s.cell_at(5, 3)->cached_value.as_number(), expected_post)
      << "post-delete SUM mismatch (range refs should collapse, not stale)";
}

TEST(WorkbookRowColEdits, RejectsZeroCount) {
  Workbook wb = Workbook::create();
  auto r = wb.insert_rows(0, /*row=*/0, /*count=*/0);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(WorkbookRowColEdits, RejectsOutOfRangeSheetIndex) {
  Workbook wb = Workbook::create();
  auto r = wb.delete_cols(99, /*col=*/0, /*count=*/1);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidArgument);
}

}  // namespace
}  // namespace formulon
