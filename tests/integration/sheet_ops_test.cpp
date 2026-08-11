//
// End-to-end tests for the workbook-level structural mutation surface:
// `Workbook::rename_sheet`, `remove_sheet`, `move_sheet`, and
// `set_defined_name`. These exercise the public C++ API directly so the
// validation rules and defined-name rewriter can be observed without
// going through the C ABI.

#include <cstdint>
#include <memory>
#include <string>

#include "cf/cf_types.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_writer.h"
#include "io/tables_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
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

void SetMoveSensitiveMetadata(Sheet& sheet, std::string_view suffix) {
  sheet.set_drawing_rel_target("xl/drawings/drawing-" + std::string(suffix) + ".xml");
  sheet.set_auto_filter_xml("<autoFilter ref=\"A1:" + std::string(suffix) + "9\"/>");
  sheet.set_ext_lst_xml("<extLst><ext uri=\"" + std::string(suffix) + "\"/></extLst>");
  sheet.set_root_extra_ns_attrs(" xmlns:x" + std::string(suffix) + "=\"urn:" + std::string(suffix) + "\"");
}

void ExpectMoveSensitiveMetadata(const Sheet& sheet, std::string_view suffix) {
  EXPECT_EQ(sheet.drawing_rel_target(), "xl/drawings/drawing-" + std::string(suffix) + ".xml");
  EXPECT_EQ(sheet.auto_filter_xml(), "<autoFilter ref=\"A1:" + std::string(suffix) + "9\"/>");
  EXPECT_EQ(sheet.ext_lst_xml(), "<extLst><ext uri=\"" + std::string(suffix) + "\"/></extLst>");
  EXPECT_EQ(sheet.root_extra_ns_attrs(), " xmlns:x" + std::string(suffix) + "=\"urn:" + std::string(suffix) + "\"");
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

TEST(WorkbookSheetOps, RenameAccepts31JapaneseCharacters) {
  Workbook wb = ThreeSheetWorkbook();
  // 31 copies of "あ" (U+3042, 3 bytes each = 93 bytes) is exactly at the
  // 31-character limit; a byte-count check would wrongly reject it.
  std::string name;
  for (int i = 0; i < 31; ++i) {
    name += "\xE3\x81\x82";  // "あ"
  }
  auto r = wb.rename_sheet(0, name);
  ASSERT_TRUE(static_cast<bool>(r)) << "31 Japanese characters must be within the limit";
  EXPECT_EQ(wb.sheet(0).name(), name);
}

TEST(WorkbookSheetOps, RenameRejects32JapaneseCharacters) {
  Workbook wb = ThreeSheetWorkbook();
  std::string name;
  for (int i = 0; i < 32; ++i) {
    name += "\xE3\x81\x82";  // "あ"
  }
  auto r = wb.rename_sheet(0, name);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet(0).name(), "Alpha");
}

TEST(WorkbookSheetOps, RenameCountsEmojiAsTwoUnits) {
  Workbook wb = ThreeSheetWorkbook();
  // "😀" (U+1F600) is a supplementary-plane codepoint: two UTF-16 units.
  // 16 emoji = 32 units, one over the limit.
  std::string name;
  for (int i = 0; i < 16; ++i) {
    name += "\xF0\x9F\x98\x80";  // "😀"
  }
  auto r = wb.rename_sheet(0, name);
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
}

TEST(WorkbookSheetOps, AddSheetValidatedAppendsAndReturnsPointer) {
  Workbook wb = ThreeSheetWorkbook();
  auto r = wb.add_sheet_validated("Delta");
  ASSERT_TRUE(static_cast<bool>(r));
  ASSERT_NE(r.value(), nullptr);
  EXPECT_EQ(r.value()->name(), "Delta");
  EXPECT_EQ(wb.sheet_count(), 4U);
}

TEST(WorkbookSheetOps, AddSheetValidatedRejectsDuplicateCaseInsensitively) {
  Workbook wb = ThreeSheetWorkbook();
  auto r = wb.add_sheet_validated("beta");  // collides with "Beta"
  ASSERT_FALSE(static_cast<bool>(r));
  EXPECT_EQ(r.error().code, FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet_count(), 3U);
}

TEST(WorkbookSheetOps, AddSheetValidatedRejectsEmptyForbiddenAndTooLong) {
  Workbook wb = ThreeSheetWorkbook();
  EXPECT_FALSE(static_cast<bool>(wb.add_sheet_validated("")));
  EXPECT_FALSE(static_cast<bool>(wb.add_sheet_validated("a/b")));
  EXPECT_FALSE(static_cast<bool>(wb.add_sheet_validated(std::string(32U, 'a'))));
  EXPECT_EQ(wb.sheet_count(), 3U);
}

TEST(WorkbookSheetOps, AddSheetValidatedAccepts31JapaneseCharacters) {
  Workbook wb = ThreeSheetWorkbook();
  std::string name;
  for (int i = 0; i < 31; ++i) {
    name += "\xE3\x81\x82";  // "あ"
  }
  auto r = wb.add_sheet_validated(name);
  ASSERT_TRUE(static_cast<bool>(r));
  EXPECT_EQ(wb.sheet_count(), 4U);
}

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

TEST(WorkbookSheetOps, RenameUpdatesWorkbookScopedDefinedNames) {
  Workbook wb = ThreeSheetWorkbook();
  // Pre-populate two defined names that both mention `Beta`. A sheet-scoped
  // name remains bound to its local sheet by ordinal, but its formula text
  // may still explicitly target the renamed sheet and must follow the name.
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
  EXPECT_EQ(wb.defined_names()[1].formula, "Banana!$B$2");
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

// Renaming a sheet must rewrite cell formulas that reference it by name,
// not just defined names. A dependent that keeps the stale name resolves
// to #REF!/#NAME? on the next recalc.
TEST(WorkbookSheetOps, RenameRewritesReferencingCellFormulas) {
  Workbook wb = ThreeSheetWorkbook();                                                // Alpha(0), Beta(1), Gamma(2)
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(2, 0, 0, Value::number(100.0))));  // Gamma!A1
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 0, "=Gamma!A1")));         // Alpha!A1
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->cached_value.as_number(), 100.0);

  // Rename Gamma -> Delta. The Alpha!A1 formula text must follow.
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(2, "Delta")));
  ASSERT_NE(wb.sheet(0).cell_at(0, 0), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->formula_text, "=Delta!A1");

  // And it still resolves after recalc (not #REF!/#NAME?).
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value a1 = wb.sheet(0).cell_at(0, 0)->cached_value;
  ASSERT_TRUE(a1.is_number()) << "renamed cross-sheet ref failed to resolve";
  EXPECT_EQ(a1.as_number(), 100.0);
}

TEST(WorkbookSheetOps, RenameRewritesSheetNamedMetadata) {
  Workbook wb = ThreeSheetWorkbook();  // Alpha(0), Beta(1), Gamma(2)
  io::DefinedName local_name;
  local_name.name = "LocalRef";
  local_name.local_sheet_id = 0;
  local_name.formula = "Gamma!$A$1";
  wb.set_defined_names({local_name});

  Hyperlink link;
  link.row = 0;
  link.col = 0;
  link.location = "#Gamma!A1";
  wb.sheet(0).mutable_hyperlinks().push_back(link);
  DataValidation validation;
  validation.ranges.push_back(MergeRange{0, 1, 0, 1});
  validation.formula1 = "Gamma!$A$1";
  validation.formula2 = "Gamma!$B$1";
  wb.sheet(0).mutable_validations().push_back(validation);

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(2, "Delta")));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "Delta!$A$1");
  ASSERT_EQ(wb.sheet(0).hyperlinks().size(), 1U);
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].location, "#Delta!A1");
  ASSERT_EQ(wb.sheet(0).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, "Delta!$A$1");
  EXPECT_EQ(wb.sheet(0).validations()[0].formula2, "Delta!$B$1");
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
  io::TableMetadata table;
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
  io::TableMetadata table;
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
  std::vector<io::DefinedName> names;
  io::DefinedName keep;
  keep.name = "Keeper";
  keep.formula = "Alpha!$A$1";
  keep.local_sheet_id = -1;
  names.push_back(keep);

  io::DefinedName frozen;
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

TEST(WorkbookSheetOps, InsertRowsRekeysAllDependenciesBeforeRegisteringNewCoordinates) {
  Workbook wb = Workbook::create();
  constexpr std::uint32_t kRows = 60U;
  for (std::uint32_t row = 0; row < kRows; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, row, 0U, Value::number(static_cast<double>(row)))));
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, row, 1U, "=A" + std::to_string(row + 1U))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0U, 0U, 1U)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(999.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  for (std::uint32_t row = 0; row < kRows; ++row) {
    const Cell* dependent = wb.sheet(0).cell_at(row + 1U, 1U);
    ASSERT_NE(dependent, nullptr) << "missing formula at row " << row + 2U;
    ASSERT_TRUE(dependent->cached_value.is_number());
    const double expected = (row == 0U) ? 999.0 : static_cast<double>(row);
    EXPECT_DOUBLE_EQ(dependent->cached_value.as_number(), expected) << "stale dependent at row " << row + 2U;
  }
}

TEST(WorkbookSheetOps, ConsecutiveRowInsertsKeepShiftedDependenciesLive) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 2, 1, "=SUM(A1:A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(2, 1)->cached_value.as_number(), 3.0);

  // Inserting inside the range expands it; the original second input moves
  // to A3 and the dependent moves to B4.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(5.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(3, 1), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(3, 1)->formula_text, "=SUM(A1:A3)");
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(3, 1)->cached_value.as_number(), 8.0);

  // A second insert must re-key the already-shifted dependency. Updating
  // the moved original A1 (now A2) must still invalidate the formula B5.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 1)));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(4, 1), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(4, 1)->formula_text, "=SUM(A2:A4)");
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(4, 1)->cached_value.as_number(), 17.0);
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

TEST(WorkbookRowColEdits, InsertRowsShiftsConditionalFormatsLayoutAndBreaks) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{CellAddress{1, 0}, CellAddress{2, 1}});
  cf::CFRule rule;
  rule.formula1 = "A2";
  format.rules.push_back(std::move(rule));
  sheet.mutable_conditional_formats().push_back(std::move(format));
  sheet.mutable_layout().row_overrides.push_back(RowLayout{2, 24.0, true, 1, true});
  sheet.mutable_layout().columns.push_back(ColumnLayout{1, 2, 18.0, true, 2});
  sheet.mutable_print_settings().manual_row_breaks.push_back(ManualBreak{3, 0, 10, true});
  sheet.mutable_print_settings().manual_col_breaks.push_back(ManualBreak{2, 0, 10, true});
  io::TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.sheet_index = 0;
  table.ref = "A2:B3";
  table.columns.push_back(io::TableColumn{1, "Value", {}, {}, "A2"});
  wb.set_tables({table});
  auto pivot = std::make_unique<pivot::PivotTable>();
  pivot->set_anchor(2, 1, 3, 2);
  sheet.add_pivot_table(std::move(pivot));
  ASSERT_TRUE(sheet.commit_spill(4, 4, 2, 1, std::vector<Value>{Value::number(1), Value::number(2)}));

  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 1)));
  const cf::CFCellRange row_shifted = sheet.conditional_formats()[0].sqref[0];
  EXPECT_EQ(row_shifted.first.row, 2U);
  EXPECT_EQ(row_shifted.last.row, 3U);
  ASSERT_TRUE(sheet.conditional_formats()[0].rules[0].formula1.has_value());
  EXPECT_EQ(*sheet.conditional_formats()[0].rules[0].formula1, "A3");
  EXPECT_EQ(sheet.layout().row_overrides[0].row, 3U);
  EXPECT_EQ(sheet.print_settings().manual_row_breaks[0].id, 4U);
  ASSERT_EQ(sheet.pivot_tables().size(), 1U);
  EXPECT_EQ(sheet.pivot_tables()[0]->anchor_row(), 3U);
  EXPECT_EQ(sheet.spill_region_at_anchor(5, 4), nullptr);
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].ref, "A3:B4");
  EXPECT_EQ(wb.tables()[0].columns[0].calculated_column_formula, "A3");
  // Column-bound metadata is not touched by a row edit.
  EXPECT_EQ(sheet.layout().columns[0].first, 1U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks[0].id, 2U);

  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));
  const cf::CFCellRange col_shifted = sheet.conditional_formats()[0].sqref[0];
  EXPECT_EQ(col_shifted.first.col, 1U);
  EXPECT_EQ(col_shifted.last.col, 2U);
  EXPECT_EQ(sheet.layout().columns[0].first, 2U);
  EXPECT_EQ(sheet.layout().columns[0].last, 3U);
  EXPECT_EQ(sheet.print_settings().manual_col_breaks[0].id, 3U);
  EXPECT_EQ(sheet.pivot_tables()[0]->anchor_col(), 2U);
  EXPECT_EQ(wb.tables()[0].ref, "B3:C4");
}

TEST(WorkbookRowColEdits, DeleteRowsRemovesConditionalFormatWhoseRangeIsDeleted) {
  Workbook wb = Workbook::create();
  cf::ConditionalFormat format;
  format.sqref.push_back(cf::CFCellRange{CellAddress{1, 0}, CellAddress{2, 1}});
  cf::CFRule rule;
  rule.formula1 = "A2";
  format.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(format));

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 1, 2)));
  EXPECT_TRUE(wb.sheet(0).conditional_formats().empty());
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

TEST(WorkbookRowColEdits, DeleteRowsShrinksRangeAtFirstLastAndMiddle) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A2)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A3)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A2)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A1:A5)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/2, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(A1:A4)");
  }
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesSingletonAndShrinksWholeRowRange) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A2:A2)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(#REF!)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(1:4)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(9, 0), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(9, 0)->formula_text, "=SUM(1:3)");
  }
}

TEST(WorkbookRowColEdits, DeleteRowsCollapsesFullyDeletedMultiCellRange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 10, 0, "=SUM(A2:A4)")));
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/1, /*count=*/3)));
  ASSERT_NE(wb.sheet(0).cell_at(7, 0), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(7, 0)->formula_text, "=SUM(#REF!)");
}

TEST(WorkbookRowColEdits, DeleteRowsShrunkRangeRecalculatesAfterSurvivorChange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 0, Value::number(40.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 3, "=SUM(A1:A4)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(6, 3)->cached_value.as_number(), 100.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(5, 3), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->formula_text, "=SUM(A1:A3)");
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->cached_value.as_number(), 90.0);

  // The surviving dependency at A2 must still point at the moved formula;
  // changing it and recalculating proves the structural edit rebuilt the
  // dependency graph rather than merely changing formula text.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(50.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(5, 3)->cached_value.as_number(), 110.0);
}

TEST(WorkbookRowColEdits, RowDeletesPreserveThreeDReferenceCoordinates) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:A3)")));

  // Excel 365 16.111.3 keeps the shared 3-D coordinate tail unchanged when
  // a sheet inside the span is edited.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_NE(wb.sheet(2).cell_at(9, 0), nullptr);
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:A3)");

  // Editing the formula owner moves the formula cell, but still does not
  // rewrite the 3-D tail.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(2, /*row=*/0, /*count=*/1)));
  ASSERT_NE(wb.sheet(2).cell_at(8, 0), nullptr);
  EXPECT_EQ(wb.sheet(2).cell_at(8, 0)->formula_text, "=SUM(First:Second!A1:A3)");
}

TEST(WorkbookRowColEdits, ThreeDRowEditsDirtySpanMembersOnly) {
  auto make_workbook = [] {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    wb.add_sheet("Outside");
    for (std::uint32_t row = 0; row < 4U; ++row) {
      EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
    }
    for (std::uint32_t row = 0; row < 3U; ++row) {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:A3)")));
    EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    return wb;
  };

  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:A3)");
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(1, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 56.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 63.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(1, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 36.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_rows(3, /*row=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    EXPECT_EQ(stats.value().cells_evaluated, 0U);
  }
}

TEST(WorkbookRowColEdits, ThreeDDefinedNamesAndNamedLambdasRecalculateAfterRowEdit) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  wb.set_defined_names({
      io::DefinedName{"Inner", "SUM(First:Second!A1:A3)", -1, false, ""},
      io::DefinedName{"Outer", "Inner+0", -1, false, ""},
      io::DefinedName{"F", "LAMBDA(x,Outer+x)", -1, false, ""},
  });
  for (std::uint32_t row = 0; row < 4U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
  }
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=Outer")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 10, 0, "=F(2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 11, 0, "=A10+A11")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(2).cell_at(9, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(9, 0)->cached_value.debug_to_string();
  ASSERT_TRUE(wb.sheet(2).cell_at(10, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(10, 0)->cached_value.debug_to_string();
  ASSERT_TRUE(wb.sheet(2).cell_at(11, 0)->cached_value.is_number())
      << wb.sheet(2).cell_at(11, 0)->cached_value.debug_to_string();
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
  EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 68.0);
  EXPECT_EQ(wb.sheet(2).cell_at(11, 0)->cached_value.as_number(), 134.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  auto stats = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats));
  EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  EXPECT_EQ(wb.sheet(2).cell_at(10, 0)->cached_value.as_number(), 71.0);
  EXPECT_EQ(wb.sheet(2).cell_at(11, 0)->cached_value.as_number(), 140.0);
  EXPECT_EQ(stats.value().cells_evaluated, 3U);
}

TEST(WorkbookRowColEdits, ThreeDOwnerMoveReordersMovedSourceAndDownstream) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 3, 0, "=1+3")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 5, 0, "=SUM(First:Second!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 0, 0, "=First!A6*2")));
  for (std::uint32_t row = 0; row < 3U; ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
  }
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(5, 0)->cached_value.as_number(), 66.0);
  ASSERT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 132.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->formula_text, "=SUM(First:Second!A1:A3)");
  EXPECT_EQ(wb.sheet(0).cell_at(4, 0)->cached_value.as_number(), 69.0);
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->formula_text, "=First!A5*2");
  EXPECT_EQ(wb.sheet(2).cell_at(0, 0)->cached_value.as_number(), 138.0);
}

TEST(WorkbookRowColEdits, ThreeDOwnerDeletedByEditIsDroppedFromRegistry) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("First");
  wb.add_sheet("Second");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(3.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 3, 0, "=SUM(First:Second!A1:A3)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 1, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1, 2, 0, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(3, 0)->cached_value.as_number(), 66.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/3, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Cell* deleted = wb.sheet(0).cell_at(3, 0);
  EXPECT_TRUE(deleted == nullptr || deleted->formula_text.empty());
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

TEST(WorkbookRowColEdits, DeleteColsShrinksRangeAtFirstLastAndMiddle) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:D1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A1:C1)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:D1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/3, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A1:C1)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A1:F1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/2, /*count=*/2)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 8), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 8)->formula_text, "=SUM(A1:D1)");
  }
}

TEST(WorkbookRowColEdits, DeleteColsCollapsesSingletonAndShrinksWholeColumnRange) {
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(B1:B1)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(#REF!)");
  }
  {
    Workbook wb = Workbook::create();
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(A:D)")));
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/1)));
    ASSERT_NE(wb.sheet(0).cell_at(0, 9), nullptr);
    EXPECT_EQ(wb.sheet(0).cell_at(0, 9)->formula_text, "=SUM(A:C)");
  }
}

TEST(WorkbookRowColEdits, DeleteColsCollapsesFullyDeletedMultiCellRange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 10, "=SUM(B1:D1)")));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/1, /*count=*/3)));
  ASSERT_NE(wb.sheet(0).cell_at(0, 7), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 7)->formula_text, "=SUM(#REF!)");
}

TEST(WorkbookRowColEdits, ThreeDColumnEditsDirtySpanMembersOnly) {
  auto make_workbook = [] {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    wb.add_sheet("Outside");
    for (std::uint32_t col = 0; col < 4U; ++col) {
      EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, col, Value::number(static_cast<double>(col + 1U)))));
    }
    for (std::uint32_t col = 0; col < 3U; ++col) {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, 0, col, Value::number(static_cast<double>((col + 1U) * 10U)))));
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A1:C1)")));
    EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    return wb;
  };

  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->formula_text, "=SUM(First:Second!A1:C1)");
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(1, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 56.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 63.0);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(1, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 36.0);
    EXPECT_EQ(stats.value().cells_evaluated, 1U);
  }
  {
    Workbook wb = make_workbook();
    ASSERT_TRUE(static_cast<bool>(wb.insert_cols(3, /*col=*/0, /*count=*/1)));
    auto stats = wb.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(stats));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 66.0);
    EXPECT_EQ(stats.value().cells_evaluated, 0U);
  }
}

TEST(WorkbookRowColEdits, ThreeDWholeAxisEditsRecalculate) {
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    for (std::uint32_t row = 0; row < 4U; ++row) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, row, 0, Value::number(static_cast<double>(row + 1U)))));
    }
    for (std::uint32_t row = 0; row < 3U; ++row) {
      ASSERT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, row, 0, Value::number(static_cast<double>((row + 1U) * 10U)))));
    }
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!A:A)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 70.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, /*row=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  }
  {
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("First");
    wb.add_sheet("Second");
    wb.add_sheet("Summary");
    for (std::uint32_t col = 0; col < 4U; ++col) {
      ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, col, Value::number(static_cast<double>(col + 1U)))));
    }
    for (std::uint32_t col = 0; col < 3U; ++col) {
      ASSERT_TRUE(
          static_cast<bool>(wb.set_cell_value(1, 0, col, Value::number(static_cast<double>((col + 1U) * 10U)))));
    }
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2, 9, 0, "=SUM(First:Second!1:1)")));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    ASSERT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 70.0);

    ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(2).cell_at(9, 0)->cached_value.as_number(), 69.0);
  }
}

TEST(WorkbookRowColEdits, DeleteColsShrunkRangeRecalculatesAfterSurvivorChange) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 2, Value::number(30.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 3, Value::number(40.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 5, 6, "=SUM(A1:D1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_EQ(wb.sheet(0).cell_at(5, 6)->cached_value.as_number(), 100.0);

  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, /*col=*/0, /*count=*/1)));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(5, 5), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->formula_text, "=SUM(A1:C1)");
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->cached_value.as_number(), 90.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 0, 1, Value::number(50.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(5, 5)->cached_value.as_number(), 110.0);
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

// A formula that reaches a shifted range only through a defined name is
// textually unchanged by the shift — a `NameRef` node is the identity case
// for the transform — so the per-formula rewriter never re-registers it. Its
// dep-graph edges were extracted by expanding the *old* definition, which now
// points somewhere else, so without a full re-index the next recalc misses
// the cells the name actually covers.
TEST(WorkbookRowColEdits, InsertRowsReindexesFormulasThatReachRangesViaDefinedName) {
  Workbook wb = Workbook::create();
  std::vector<io::DefinedName> names;
  io::DefinedName dn;
  dn.name = "Region";
  dn.formula = "Sheet1!$A$2:$A$3";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  // A2 = 10, A3 = 20 (0-based rows 1 and 2).
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 1, 0, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2, 0, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0, 2, "=SUM(Region)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(0, 2), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 30.0);

  // Insert a row above the named range: the definition moves to $A$3:$A$4 and
  // the two values move with it. `=SUM(Region)` is unchanged as text.
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, /*row=*/1, /*count=*/1)));
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$3:$A$4");

  // Write to A4 (0-based row 3). That row is inside the *new* definition and
  // outside the old one, so the edit only marks `=SUM(Region)` dirty if the
  // dep graph was re-indexed against the rewritten definition.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0, 3, 0, Value::number(25.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_NE(wb.sheet(0).cell_at(0, 2), nullptr);
  EXPECT_EQ(wb.sheet(0).cell_at(0, 2)->cached_value.as_number(), 35.0)
      << "SUM over the defined name did not follow the shifted definition";
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
