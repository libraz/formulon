// Structural workbook mutation tests grouped by public surface.
#include "sheet_ops_test_support.h"

namespace formulon {
namespace {
using namespace sheet_ops_test;

TEST(WorkbookSheetOps, SheetMovePreservesRawWorksheetMetadata) {
  Workbook wb = ThreeSheetWorkbook();
  SetMoveSensitiveMetadata(wb.sheet(0), "alpha");
  SetMoveSensitiveMetadata(wb.sheet(1), "beta");
  SetMoveSensitiveMetadata(wb.sheet(2), "gamma");

  // Appending enough sheets guarantees a vector reallocation. Then erase and
  // move exercise move assignment and move construction respectively.
  for (std::uint32_t i = 0; i < 32U; ++i) {
    wb.add_sheet("Extra" + std::to_string(i));
  }
  ExpectMoveSensitiveMetadata(wb.sheet(0), "alpha");
  ExpectMoveSensitiveMetadata(wb.sheet(1), "beta");
  ExpectMoveSensitiveMetadata(wb.sheet(2), "gamma");

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  ExpectMoveSensitiveMetadata(wb.sheet(0), "alpha");
  ExpectMoveSensitiveMetadata(wb.sheet(1), "gamma");

  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(1, 0)));
  ExpectMoveSensitiveMetadata(wb.sheet(0), "gamma");
  ExpectMoveSensitiveMetadata(wb.sheet(1), "alpha");
}

TEST(WorkbookSheetOps, RemoveDropsSheet) {
  Workbook wb = ThreeSheetWorkbook();
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
  EXPECT_EQ(wb.sheet(1).name(), "Gamma");
}

TEST(WorkbookSheetOps, RemoveDropsPivotCacheWhoseWorksheetSourceWasRemoved) {
  Workbook wb = ThreeSheetWorkbook();
  auto removed_source_cache = std::make_unique<pivot::PivotCache>();
  removed_source_cache->set_cache_id(1U);
  removed_source_cache->mutable_worksheet_source() = {true, "$A$1:$B$10", "Beta", ""};
  wb.add_pivot_cache(std::move(removed_source_cache));

  auto surviving_source_cache = std::make_unique<pivot::PivotCache>();
  surviving_source_cache->set_cache_id(2U);
  surviving_source_cache->mutable_worksheet_source() = {true, "$A$1:$B$10", "Gamma", ""};
  wb.add_pivot_cache(std::move(surviving_source_cache));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1U)));  // Beta
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  ASSERT_NE(wb.pivot_caches()[0], nullptr);
  EXPECT_EQ(wb.pivot_caches()[0]->cache_id(), 2U);
  EXPECT_EQ(wb.pivot_caches()[0]->worksheet_source().sheet, "Gamma");
}

TEST(WorkbookSheetOps, SheetOperationsKeepTablesAttachedToTheirOwningSheet) {
  Workbook wb = ThreeSheetWorkbook();
  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.sheet_index = 2;  // Gamma
  table.ref = "A1:B2";
  wb.set_tables({table});

  ASSERT_TRUE(static_cast<bool>(wb.move_sheet(2, 0)));
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.sheet(wb.tables()[0].sheet_index).name(), "Gamma");

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));  // remove Alpha
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.sheet(wb.tables()[0].sheet_index).name(), "Gamma");

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));  // remove Gamma
  EXPECT_TRUE(wb.tables().empty());
}

TEST(WorkbookSheetOps, WriterRejectsTableDetachedFromAllSheets) {
  Workbook wb = Workbook::create();
  TableMetadata table;
  table.id = 1;
  table.name = "Detached";
  table.display_name = "Detached";
  table.sheet_index = 1;
  table.ref = "A1:A2";
  wb.set_tables({table});
  auto written = io::write_ooxml(wb);
  ASSERT_FALSE(static_cast<bool>(written));
  EXPECT_EQ(written.error().code, FormulonErrorCode::kIoWriteFailed);
}

TEST(WorkbookSheetOps, RemoveFreezesReferencingFormulaBeforeSameNameIsReadded) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Source");
  wb.add_sheet("Dependent");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(42))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 0, "=Source!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  ASSERT_EQ(wb.sheet(0).name(), "Dependent");
  ASSERT_NE(wb.sheet(0).cell_at(0, 0), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->formula_text, "=#REF!");

  wb.add_sheet("Source");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(99))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value value = wb.sheet(0).cell_at(0, 0)->cached_value;
  ASSERT_TRUE(value.is_error());
  EXPECT_EQ(value.as_error(), ErrorCode::Ref);
}

TEST(WorkbookSheetOps, AppendReportsTheIndexOfTheSheetItAppended) {
  Workbook wb = Workbook::create_empty();
  EXPECT_EQ(wb.add_sheet("First"), 0U);
  EXPECT_EQ(wb.add_sheet("Second"), 1U);

  auto checked = wb.add_sheet_checked("Third");
  ASSERT_TRUE(static_cast<bool>(checked));
  EXPECT_EQ(checked.value(), 2U);

  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "First");
  EXPECT_EQ(wb.sheet(1).name(), "Second");
  EXPECT_EQ(wb.sheet(2).name(), "Third");
}

TEST(WorkbookSheetOps, AppendedSheetStaysReachableAcrossLaterAppends) {
  Workbook wb = Workbook::create_empty();
  const std::size_t first = wb.add_sheet("First");
  // Enough appends to force the underlying vector to reallocate at least
  // once past whatever capacity the first append reserved.
  for (int i = 0; i < 64; ++i) {
    wb.add_sheet("Filler" + std::to_string(i));
  }
  wb.sheet(first).set_cell_value(0U, 0U, Value::number(7.0));

  EXPECT_EQ(wb.sheet(first).name(), "First");
  const Cell* cell = wb.sheet(first).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->cached_value.as_number(), 7.0);
}

TEST(WorkbookSheetOps, AppendStopsAtTheSheetIdCeiling) {
  Workbook wb = Workbook::create_empty();
  std::string name;
  for (std::size_t i = 0; i < Workbook::kMaxSheets; ++i) {
    name.assign("S");
    name.append(std::to_string(i));
    ASSERT_TRUE(static_cast<bool>(wb.add_sheet_checked(name))) << "rejected at sheet " << i;
  }
  ASSERT_EQ(wb.sheet_count(), Workbook::kMaxSheets);

  auto checked = wb.add_sheet_checked("Overflow");
  ASSERT_FALSE(static_cast<bool>(checked));
  EXPECT_EQ(checked.error().code, FormulonErrorCode::kSheetCountLimitExceeded);

  auto validated = wb.add_sheet_validated("Overflow");
  ASSERT_FALSE(static_cast<bool>(validated));
  EXPECT_EQ(validated.error().code, FormulonErrorCode::kSheetCountLimitExceeded);

  // The name-unchecked overload has no error channel, but it must not grow
  // the workbook past the ceiling either, and the index it reports back must
  // not name a sheet the caller could then mutate by mistake.
  EXPECT_EQ(wb.add_sheet("Overflow"), Workbook::kMaxSheets);
  EXPECT_EQ(wb.sheet_count(), Workbook::kMaxSheets);
  EXPECT_EQ(wb.sheet(Workbook::kMaxSheets - 1U).name(), "S" + std::to_string(Workbook::kMaxSheets - 1U));
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

TEST(WorkbookSheetOps, RemoveFreezesWorkbookScopedNamesTargetingRemovedSheet) {
  Workbook wb = ThreeSheetWorkbook();
  std::vector<DefinedName> names;
  DefinedName keep;
  keep.name = "Keeper";
  keep.formula = "Alpha!$A$1";
  keep.local_sheet_id = -1;
  names.push_back(keep);

  DefinedName frozen;
  frozen.name = "Goner";
  frozen.formula = "Beta!$A$1";
  frozen.local_sheet_id = -1;
  names.push_back(frozen);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));

  ASSERT_EQ(wb.defined_names().size(), 2U);
  EXPECT_EQ(wb.defined_names()[0].name, "Keeper");
  EXPECT_EQ(wb.defined_names()[1].name, "Goner");
  EXPECT_EQ(wb.defined_names()[1].formula, "#REF!");
}

TEST(WorkbookSheetOps, RemoveAdjustsSheetScopedLocalIds) {
  Workbook wb = ThreeSheetWorkbook();
  std::vector<DefinedName> names;
  DefinedName scoped_alpha;
  scoped_alpha.name = "A";
  scoped_alpha.formula = "$A$1";
  scoped_alpha.local_sheet_id = 0;
  names.push_back(scoped_alpha);

  DefinedName scoped_gamma;
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
  std::vector<DefinedName> names;
  DefinedName name_alpha;
  name_alpha.name = "A";
  name_alpha.formula = "$A$1";
  name_alpha.local_sheet_id = 0;
  DefinedName name_beta;
  name_beta.name = "B";
  name_beta.formula = "$A$1";
  name_beta.local_sheet_id = 1;
  DefinedName name_gamma;
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

}  // namespace
}  // namespace formulon
