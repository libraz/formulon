// Structural workbook mutation tests grouped by public surface.
#include <functional>

#include "sheet_ops_test_support.h"

namespace formulon {
namespace {
using namespace sheet_ops_test;

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

TEST(WorkbookSheetOps, UnicodeSimpleFoldRejectsDuplicateAndResolvesLookup) {
  Workbook wb = Workbook::create_empty();
  auto added = wb.add_sheet_validated("\xC3\x84");  // Ä
  ASSERT_TRUE(static_cast<bool>(added));
  EXPECT_EQ(wb.sheet_index_by_name("\xC3\xA4"), 0U);  // ä
  EXPECT_EQ(wb.sheet_by_name("\xC3\xA4"), added.value());

  auto duplicate = wb.add_sheet_validated("\xC3\xA4");
  ASSERT_FALSE(static_cast<bool>(duplicate));
  EXPECT_EQ(duplicate.error().code, FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet_count(), 1U);

  auto distinct = wb.add_sheet_validated("\xC3\x96");  // Ö
  ASSERT_TRUE(static_cast<bool>(distinct));
  auto rename_collision = wb.rename_sheet(1U, "\xC3\xA4");  // ä
  ASSERT_FALSE(static_cast<bool>(rename_collision));
  EXPECT_EQ(rename_collision.error().code, FormulonErrorCode::kInvalidSheetName);
  EXPECT_EQ(wb.sheet(0).name(), "\xC3\x84");  // Ä
  EXPECT_EQ(wb.sheet(1).name(), "\xC3\x96");  // Ö
}

TEST(WorkbookSheetOps, UnicodeSimpleFoldSupportsCasingRenameAndDependencies) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Source");
  auto added = wb.add_sheet_validated("\xC3\x84");  // Ä
  ASSERT_TRUE(static_cast<bool>(added));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "='\xC3\xA4'!A1")));  // alternate case: ä
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(0).cell_at(0U, 0U)->cached_value.is_number());
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(0U, 0U)->cached_value.as_number(), 10.0);

  // A casing-only rename preserves the identity used by the dependency edge.
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1U, "\xC3\xA4")));  // ä
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(25.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(0).cell_at(0U, 0U)->cached_value.is_number());
  EXPECT_DOUBLE_EQ(wb.sheet(0).cell_at(0U, 0U)->cached_value.as_number(), 25.0);
}

TEST(WorkbookSheetOps, UnicodeMalformedUtf8IsInvalidSheetName) {
  Workbook wb = Workbook::create();
  const std::string malformed("\xC3", 1U);
  auto added = wb.add_sheet_validated(malformed);
  ASSERT_FALSE(static_cast<bool>(added));
  EXPECT_EQ(added.error().code, FormulonErrorCode::kInvalidSheetName);
  auto renamed = wb.rename_sheet(0U, malformed);
  ASSERT_FALSE(static_cast<bool>(renamed));
  EXPECT_EQ(renamed.error().code, FormulonErrorCode::kInvalidSheetName);
}

TEST(WorkbookSheetOps, RenameUpdatesWorkbookScopedDefinedNames) {
  Workbook wb = ThreeSheetWorkbook();
  // Pre-populate two defined names that both mention `Beta`. A sheet-scoped
  // name remains bound to its local sheet by ordinal, but its formula text
  // may still explicitly target the renamed sheet and must follow the name.
  std::vector<DefinedName> names;
  DefinedName wb_scoped;
  wb_scoped.name = "TotalArea";
  wb_scoped.formula = "Beta!$A$1:$A$10";
  wb_scoped.local_sheet_id = -1;
  names.push_back(wb_scoped);

  DefinedName sheet_scoped;
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

  std::vector<DefinedName> names;
  DefinedName dn;
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

  std::vector<DefinedName> names;
  DefinedName dn;
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

  std::vector<DefinedName> names;
  DefinedName dn;
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

  std::vector<DefinedName> names;
  DefinedName dn;
  dn.name = "X";
  dn.formula = "OtherSheet1!$A$1";
  dn.local_sheet_id = -1;
  names.push_back(dn);
  wb.set_defined_names(std::move(names));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(0, "First")));
  EXPECT_EQ(wb.defined_names()[0].formula, "OtherSheet1!$A$1");
}

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
  DefinedName local_name;
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

  cf::ConditionalFormat conditional_format;
  cf::CFRule rule;
  rule.formula1 = "Gamma!C1";
  rule.formula2 = "Gamma!D1";
  rule.color_scale = cf::ColorScaleSpec{};
  rule.color_scale->thresholds.push_back({cf::CfvoType::Formula, "Gamma!E1", true});
  rule.icon_set = cf::IconSetSpec{};
  rule.icon_set->thresholds.push_back({cf::CfvoType::Formula, "Gamma!F1", true});
  rule.data_bar = cf::DataBarSpec{};
  rule.data_bar->min = {cf::CfvoType::Formula, "Gamma!G1", true};
  rule.data_bar->max = {cf::CfvoType::Formula, "Gamma!H1", true};
  conditional_format.rules.push_back(std::move(rule));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(conditional_format));

  TableMetadata table;
  table.id = 1;
  table.name = "Table1";
  table.display_name = "Table1";
  table.ref = "A1:B2";
  table.sheet_index = 0;
  table.columns.push_back(TableColumn{1, "Value", {}, {}, "Gamma!$A$1"});
  wb.set_tables({std::move(table)});

  auto cache = std::make_unique<pivot::PivotCache>();
  cache->set_cache_id(7U);
  cache->mutable_worksheet_source() = {true, "Gamma!$A$1:$B$2", "gAmMa", ""};
  wb.add_pivot_cache(std::move(cache));

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(2, "Delta")));
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].formula, "Delta!$A$1");
  ASSERT_EQ(wb.sheet(0).hyperlinks().size(), 1U);
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].location, "#Delta!A1");
  ASSERT_EQ(wb.sheet(0).validations().size(), 1U);
  EXPECT_EQ(wb.sheet(0).validations()[0].formula1, "Delta!$A$1");
  EXPECT_EQ(wb.sheet(0).validations()[0].formula2, "Delta!$B$1");
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  ASSERT_EQ(wb.sheet(0).conditional_formats()[0].rules.size(), 1U);
  const cf::CFRule& rewritten_rule = wb.sheet(0).conditional_formats()[0].rules[0];
  ASSERT_TRUE(rewritten_rule.formula1.has_value());
  ASSERT_TRUE(rewritten_rule.formula2.has_value());
  EXPECT_EQ(*rewritten_rule.formula1, "Delta!C1");
  EXPECT_EQ(*rewritten_rule.formula2, "Delta!D1");
  ASSERT_TRUE(rewritten_rule.color_scale.has_value());
  EXPECT_EQ(rewritten_rule.color_scale->thresholds[0].value, "Delta!E1");
  ASSERT_TRUE(rewritten_rule.icon_set.has_value());
  EXPECT_EQ(rewritten_rule.icon_set->thresholds[0].value, "Delta!F1");
  ASSERT_TRUE(rewritten_rule.data_bar.has_value());
  EXPECT_EQ(rewritten_rule.data_bar->min.value, "Delta!G1");
  EXPECT_EQ(rewritten_rule.data_bar->max.value, "Delta!H1");
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].columns[0].calculated_column_formula, "Delta!$A$1");
  ASSERT_EQ(wb.pivot_caches().size(), 1U);
  EXPECT_EQ(wb.pivot_caches()[0]->worksheet_source().sheet, "Delta");
  EXPECT_EQ(wb.pivot_caches()[0]->worksheet_source().ref, "Delta!$A$1:$B$2");
}

namespace {

using AddMissingSheet = std::function<void(Workbook&)>;

void ExerciseMissingSheetRecovery(const AddMissingSheet& add_missing) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Source");
  wb.add_sheet("Summary");

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(0U, 1U, 0U, Value::number(2.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Leaf", "=Missing!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Outer", "=Leaf")));

  // These formulas are entered while Missing does not exist. The evaluator
  // reports #REF!, and the dep extractor intentionally has no target edge to
  // register yet. Adding the sheet must recover every static reference form.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 0U, "=Missing!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 1U, "=SUM(Missing!A1:A2)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 2U, "=SUM(Source:Missing!A1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 3U, "=Outer")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 4U, "=Missing!D1#")));
  // This spill is unrelated to the missing sheet and must remain committed
  // while the recovery pass reparses affected formulas.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(1U, 0U, 5U, "=SEQUENCE(2,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  for (std::uint32_t col = 0U; col < 5U; ++col) {
    const Value value = wb.sheet(1U).resolve_cell_value(0U, col);
    ASSERT_TRUE(value.is_error()) << "column=" << col;
    EXPECT_EQ(value.as_error(), ErrorCode::Ref) << "column=" << col;
  }
  const Value unrelated_before = wb.sheet(1U).resolve_cell_value(1U, 5U);
  ASSERT_TRUE(unrelated_before.is_number());
  EXPECT_DOUBLE_EQ(unrelated_before.as_number(), 2.0);

  add_missing(wb);
  const std::size_t missing = wb.sheet_index_by_name("Missing");
  ASSERT_NE(missing, static_cast<std::size_t>(-1));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(missing, 0U, 0U, Value::number(10.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(missing, 1U, 0U, Value::number(20.0))));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(missing, 0U, 3U, "=SEQUENCE(2,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));

  const Value direct = wb.sheet(1U).resolve_cell_value(0U, 0U);
  const Value range = wb.sheet(1U).resolve_cell_value(0U, 1U);
  const Value span = wb.sheet(1U).resolve_cell_value(0U, 2U);
  const Value alias = wb.sheet(1U).resolve_cell_value(0U, 3U);
  const Value spill_ref = wb.sheet(1U).resolve_cell_value(0U, 4U);
  ASSERT_TRUE(direct.is_number());
  ASSERT_TRUE(range.is_number());
  ASSERT_TRUE(span.is_number());
  ASSERT_TRUE(alias.is_number());
  ASSERT_TRUE(spill_ref.is_number());
  EXPECT_DOUBLE_EQ(direct.as_number(), 10.0);
  EXPECT_DOUBLE_EQ(range.as_number(), 30.0);
  // Source!A1 + Summary!A1 + Missing!A1.
  EXPECT_DOUBLE_EQ(span.as_number(), 21.0);
  EXPECT_DOUBLE_EQ(alias.as_number(), 10.0);
  EXPECT_DOUBLE_EQ(spill_ref.as_number(), 1.0);
  const Value spill_ref_tail = wb.sheet(1U).resolve_cell_value(1U, 4U);
  ASSERT_TRUE(spill_ref_tail.is_number());
  EXPECT_DOUBLE_EQ(spill_ref_tail.as_number(), 2.0);
  const Value unrelated_after = wb.sheet(1U).resolve_cell_value(1U, 5U);
  ASSERT_TRUE(unrelated_after.is_number());
  EXPECT_DOUBLE_EQ(unrelated_after.as_number(), 2.0);

  // The recovered formulas must retain live dependencies after the first
  // successful recalc, including the range and the 3-D endpoint.
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(missing, 0U, 0U, Value::number(15.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  ASSERT_TRUE(wb.sheet(1U).resolve_cell_value(0U, 0U).is_number());
  ASSERT_TRUE(wb.sheet(1U).resolve_cell_value(0U, 1U).is_number());
  ASSERT_TRUE(wb.sheet(1U).resolve_cell_value(0U, 2U).is_number());
  ASSERT_TRUE(wb.sheet(1U).resolve_cell_value(0U, 3U).is_number());
  EXPECT_DOUBLE_EQ(wb.sheet(1U).resolve_cell_value(0U, 0U).as_number(), 15.0);
  EXPECT_DOUBLE_EQ(wb.sheet(1U).resolve_cell_value(0U, 1U).as_number(), 35.0);
  EXPECT_DOUBLE_EQ(wb.sheet(1U).resolve_cell_value(0U, 2U).as_number(), 31.0);
  EXPECT_DOUBLE_EQ(wb.sheet(1U).resolve_cell_value(0U, 3U).as_number(), 15.0);
}

TEST(WorkbookSheetOps, AddSheetRecoversMissingSheetReferences) {
  ExerciseMissingSheetRecovery([](Workbook& wb) { wb.add_sheet("Missing"); });
}

TEST(WorkbookSheetOps, AddSheetRecoversMissingReferenceWithNoGraphNodes) {
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0U, 0U, 0U, "=Missing!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value before = wb.sheet(0U).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(before.is_error());
  EXPECT_EQ(before.as_error(), ErrorCode::Ref);
  ASSERT_EQ(wb.recalc_engine().dep_graph().node_count(), 0U);

  const std::size_t missing = wb.add_sheet("Missing");
  ASSERT_EQ(missing, 1U);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(missing, 0U, 0U, Value::number(12.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  const Value after = wb.sheet(0U).resolve_cell_value(0U, 0U);
  ASSERT_TRUE(after.is_number());
  EXPECT_DOUBLE_EQ(after.as_number(), 12.0);
}

TEST(WorkbookSheetOps, AddSheetCheckedRecoversMissingSheetReferences) {
  ExerciseMissingSheetRecovery([](Workbook& wb) { ASSERT_TRUE(static_cast<bool>(wb.add_sheet_checked("Missing"))); });
}

TEST(WorkbookSheetOps, AddSheetValidatedRecoversMissingSheetReferences) {
  ExerciseMissingSheetRecovery([](Workbook& wb) { ASSERT_TRUE(static_cast<bool>(wb.add_sheet_validated("Missing"))); });
}

TEST(WorkbookSheetOps, RenameToMissingRecoversQualifiedLocalNameAliases) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Source");
  wb.add_sheet("Other");
  wb.add_sheet("Summary");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(7.0))));

  // LocalLeaf is deliberately scoped to Summary. Alias and Alias2 exercise
  // the reverse alias closure through a qualified local NameRef.
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("LocalLeaf", "=Missing!A1", 2)));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Alias", "=Summary!LocalLeaf")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Alias2", "=Alias")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2U, 0U, 0U, "=Missing!A1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2U, 0U, 1U, "=Summary!LocalLeaf")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2U, 0U, 2U, "=Alias")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2U, 0U, 3U, "=Alias2")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(2U, 0U, 5U, "=SEQUENCE(2,1)")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t col = 0U; col < 4U; ++col) {
    const Value value = wb.sheet(2U).resolve_cell_value(0U, col);
    ASSERT_TRUE(value.is_error()) << "column=" << col;
    EXPECT_EQ(value.as_error(), ErrorCode::Ref) << "column=" << col;
  }
  ASSERT_TRUE(wb.sheet(2U).resolve_cell_value(1U, 5U).is_number());
  EXPECT_DOUBLE_EQ(wb.sheet(2U).resolve_cell_value(1U, 5U).as_number(), 2.0);

  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1U, "Missing")));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t col = 0U; col < 4U; ++col) {
    const Value value = wb.sheet(2U).resolve_cell_value(0U, col);
    ASSERT_TRUE(value.is_number()) << "column=" << col;
    EXPECT_DOUBLE_EQ(value.as_number(), 7.0) << "column=" << col;
  }
  ASSERT_TRUE(wb.sheet(2U).resolve_cell_value(1U, 5U).is_number());
  EXPECT_DOUBLE_EQ(wb.sheet(2U).resolve_cell_value(1U, 5U).as_number(), 2.0);

  ASSERT_TRUE(static_cast<bool>(wb.set_cell_value(1U, 0U, 0U, Value::number(11.0))));
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t col = 0U; col < 4U; ++col) {
    const Value value = wb.sheet(2U).resolve_cell_value(0U, col);
    ASSERT_TRUE(value.is_number()) << "column=" << col;
    EXPECT_DOUBLE_EQ(value.as_number(), 11.0) << "column=" << col;
  }
}

}  // namespace

}  // namespace
}  // namespace formulon
