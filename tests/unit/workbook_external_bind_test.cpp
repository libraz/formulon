//
// Every external book a formula in the model names has an external-link
// record: entering one creates the record, re-entering it reuses it, and a
// loaded `[N]` formula reads back in the formula bar's spelling.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/formula_entry.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "parser/ast_format.h"
#include "sheet.h"
#include "sheet_passthrough.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

ExternalLinkRecord LoadedLink(std::uint32_t index, std::string target, std::string absolute_target) {
  ExternalLinkRecord rec;
  rec.index = index;
  rec.rel_id = "rId" + std::to_string(index + 2U);
  rec.part_path = "xl/externalLinks/externalLink" + std::to_string(index) + ".xml";
  rec.body_rel_id = "rId1";
  rec.target = std::move(target);
  rec.absolute_target = std::move(absolute_target);
  if (!rec.absolute_target.empty()) {
    rec.absolute_rel_id = "rId2";
  }
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  rec.book.sheet_names = {"Data"};
  rec.book.sheet_data = {true};
  return rec;
}

Workbook OneSheet() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Host");
  return wb;
}

void SetFormula(Workbook& wb, std::uint32_t row, std::uint32_t col, const std::string& formula) {
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, col, formula))) << formula;
}

// The spelling entry stores for `formula`, before it is bound.
std::string EnteredSpelling(const Workbook& wb, const std::string& formula) {
  eval::EnteredCallSpellings calls;
  std::string stored;
  parser::ParseFailure failure;
  EXPECT_EQ(eval::enter_stored_formula(wb, calls, formula, &stored, &failure), eval::EntryResult::kStored) << formula;
  return stored;
}

TEST(WorkbookExternalBind, EnteringABracketedBookCreatesOneLink) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]Sheet1!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  const ExternalLinkRecord& rec = wb.external_links()[0];
  EXPECT_EQ(rec.index, 1U);
  EXPECT_EQ(rec.kind, ExternalLinkRecord::Kind::kExternalBook);
  EXPECT_TRUE(rec.target_external);
  EXPECT_EQ(rec.target, "Book.xlsx");
  EXPECT_TRUE(rec.part_path.empty());
  EXPECT_TRUE(rec.rel_id.empty());
  ASSERT_EQ(rec.book.sheet_names.size(), 1U);
  EXPECT_EQ(rec.book.sheet_names[0], "Sheet1");
  EXPECT_FALSE(rec.book.sheet_has_data(0));
}

TEST(WorkbookExternalBind, ReEnteringTheSameBookReusesTheLink) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]Sheet1!A1");
  SetFormula(wb, 0, 0, "=[Book.xlsx]Sheet1!A1");
  SetFormula(wb, 1, 0, "=SUM([Book.xlsx]Sheet1!A1:B2)");
  // Book and sheet both match under ASCII case folding.
  SetFormula(wb, 2, 0, "=[book.XLSX]sheet1!C3");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].book.sheet_names.size(), 1U);
}

TEST(WorkbookExternalBind, ANewSheetIsAppendedInFirstSeenOrder) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]Beta!A1+[Book.xlsx]Alpha!A1");
  SetFormula(wb, 1, 0, "=SUM([Book.xlsx]Alpha:Gamma!A1)");
  ASSERT_EQ(wb.external_links().size(), 1U);
  const std::vector<std::string> expected = {"Beta", "Alpha", "Gamma"};
  EXPECT_EQ(wb.external_links()[0].book.sheet_names, expected);
}

// Excel writes a closed source's sheet list in tab order, so a loaded list
// orders a span typed last-to-first.
TEST(WorkbookExternalBind, ALoadedSheetListOrdersAnEnteredSpan) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  links[0].book.sheet_names = {"A", "B", "C"};
  links[0].book.sheet_data = {true, true, true};
  wb.set_external_links(std::move(links));
  EXPECT_EQ(EnteredSpelling(wb, "=SUM([Src.xlsx]C:A!A1)"), "=SUM([Src.xlsx]A:C!A1)");
}

TEST(WorkbookExternalBind, ASpanWithAnEnteredEndKeepsTheTypedSpelling) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  links[0].book.sheet_names = {"A", "B", "C"};
  links[0].book.sheet_data = {true, true, true};
  wb.set_external_links(std::move(links));
  EXPECT_EQ(EnteredSpelling(wb, "=SUM([Src.xlsx]Zz:A!A1)"), "=SUM([Src.xlsx]Zz:A!A1)");
  SetFormula(wb, 0, 0, "=SUM([Src.xlsx]Zz:A!A1)");
  EXPECT_EQ(wb.external_links()[0].book.entered_from, 3U);
}

// The entry boundary is not saved, so a reloaded list built by entry counts as
// loaded and orders a span by its own order.
TEST(WorkbookExternalBind, AReloadedEntryBuiltListOrdersAnEnteredSpan) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Src.xlsx]Y!A1");
  SetFormula(wb, 1, 0, "=[Src.xlsx]X!A1");
  EXPECT_EQ(EnteredSpelling(wb, "=SUM([Src.xlsx]X:Y!A1)"), "=SUM([Src.xlsx]X:Y!A1)");

  const auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  auto loaded = io::read_ooxml(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  Workbook& reloaded = loaded.value().workbook;
  ASSERT_EQ(reloaded.external_links().size(), 1U);
  EXPECT_EQ(reloaded.external_links()[0].book.entered_from, ExternalBook::kNoSheet);
  EXPECT_EQ(EnteredSpelling(reloaded, "=SUM([Src.xlsx]X:Y!A1)"), "=SUM([Src.xlsx]Y:X!A1)");
}

TEST(WorkbookExternalBind, ALinkCreatedByEntryTakesItsWholeListFromEntry) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]Beta!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].book.entered_from, 0U);
}

TEST(WorkbookExternalBind, ANewSheetOnALoadedLinkMarksTheBodyStale) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, 0, "=[Src.xlsx]data!A1");
  EXPECT_FALSE(wb.external_links()[0].body_stale);
  SetFormula(wb, 1, 0, "=[Src.xlsx]Extra!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_TRUE(wb.external_links()[0].body_stale);
  ASSERT_EQ(wb.external_links()[0].book.sheet_names.size(), 2U);
  EXPECT_EQ(wb.external_links()[0].book.sheet_names[1], "Extra");
  EXPECT_FALSE(wb.external_links()[0].book.sheet_has_data(1));
}

TEST(WorkbookExternalBind, UnicodeSheetIdentityDoesNotAppendDuplicate) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1U, "Src.xlsx", ""));
  links[0].book.sheet_names = {"\xC3\x84rea"};
  links[0].book.sheet_data = {true};
  wb.set_external_links(std::move(links));

  SetFormula(wb, 0U, 0U, "=[Src.xlsx]\xC3\xA4rea!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  ASSERT_EQ(wb.external_links()[0].book.sheet_names.size(), 1U);
  EXPECT_EQ(wb.external_links()[0].book.sheet_names[0], "\xC3\x84rea");
}

TEST(WorkbookExternalBind, NewLinksAppendWithoutRenumbering) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, 0, "=[Other.xlsx]S!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[0].index, 1U);
  EXPECT_EQ(wb.external_links()[0].target, "Src.xlsx");
  EXPECT_EQ(wb.external_links()[1].index, 2U);
  EXPECT_EQ(wb.external_links()[1].target, "Other.xlsx");
}

TEST(WorkbookExternalBind, APathMatchesTheAbsoluteTarget) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", "/Users/x/Src.xlsx"));
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, 0, "='/Users/x/[Src.xlsx]Data'!A1");
  SetFormula(wb, 1, 0, "=[Src.xlsx]Data!A1");
  EXPECT_EQ(wb.external_links().size(), 1U);
  // A different directory is a different file.
  SetFormula(wb, 2, 0, "='/Other/[Src.xlsx]Data'!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[1].target, "/Other/Src.xlsx");
}

TEST(WorkbookExternalBind, APathFallsBackToAPathlessRecordOfTheSameFile) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]S!A1");
  SetFormula(wb, 1, 0, "='/Users/x/[Book.xlsx]S'!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  // Excel retargets the merged link at the absolute path.
  EXPECT_EQ(wb.external_links()[0].target, "/Users/x/Book.xlsx");
  EXPECT_EQ(wb.external_links()[0].absolute_target, "/Users/x/Book.xlsx");
}

TEST(WorkbookExternalBind, APathlessBookMatchesTheLowestIndex) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "/a/Src.xlsx", ""));
  links.push_back(LoadedLink(2, "/b/Src.xlsx", ""));
  wb.set_external_links(std::move(links));
  const ExternalLinkRecord* found = wb.find_external_link("", "SRC.xlsx");
  ASSERT_NE(found, nullptr);
  EXPECT_EQ(found->index, 1U);
  ASSERT_NE(wb.find_external_link("/b/", "Src.xlsx"), nullptr);
  EXPECT_EQ(wb.find_external_link("/b/", "Src.xlsx")->index, 2U);
  EXPECT_EQ(wb.find_external_link("", "Nope.xlsx"), nullptr);
}

TEST(WorkbookExternalBind, RetargetingAPathMissingLinkMakesItAPathLink) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Book.xlsx", ""));
  links.back().target_path_missing = true;
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, 0, "='/Users/x/[Book.xlsx]Data'!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].target, "/Users/x/Book.xlsx");
  EXPECT_FALSE(wb.external_links()[0].target_path_missing);
}

TEST(WorkbookExternalBind, APathTargetIsStoredAsTheAbsolutePath) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "='/Users/x/[Book.xlsx]S'!A1");
  SetFormula(wb, 1, 0, "='C:\\a\\b\\[Win.xlsx]S'!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[0].target, "/Users/x/Book.xlsx");
  EXPECT_EQ(wb.external_links()[1].target, "file:///C:\\a\\b\\Win.xlsx");
}

TEST(WorkbookExternalBind, ABookScopeNameCreatesALinkWithNoSheet) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=Book.xlsx!Total");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_TRUE(wb.external_links()[0].book.sheet_names.empty());
}

TEST(WorkbookExternalBind, TheSelfBookNameCreatesNoLink) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[0]!Rate");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Alias", "[0]!Rate")));
  EXPECT_TRUE(wb.external_links().empty());
}

TEST(WorkbookExternalBind, AnUnparseableFormulaCreatesNoLink) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Book.xlsx]S!A1+");
  EXPECT_TRUE(wb.external_links().empty());
}

TEST(WorkbookExternalBind, DefinedNamesBindTheirBooks) {
  Workbook wb = OneSheet();
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("One", "[One.xlsx]S!$A$1")));
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name_scoped("Two", "[Two.xlsx]S!$A$1", 0)));
  wb.set_defined_names({DefinedName{"Three", "[Three.xlsx]S!$A$1", -1, false, ""},
                        DefinedName{"OneAgain", "[One.xlsx]S!$B$1", -1, false, ""}});
  ASSERT_EQ(wb.external_links().size(), 3U);
  EXPECT_EQ(wb.external_links()[0].target, "One.xlsx");
  EXPECT_EQ(wb.external_links()[1].target, "Two.xlsx");
  EXPECT_EQ(wb.external_links()[2].target, "Three.xlsx");
}

TEST(WorkbookExternalBind, TextBindCoversFeatureFormulas) {
  Workbook wb = OneSheet();
  wb.bind_external_books("[Cf.xlsx]S!A1>5");
  wb.bind_external_books("A1>5");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].target, "Cf.xlsx");
}

TEST(WorkbookExternalBind, IngestSpellsALoadedIndexWithItsAbsolutePath) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", "/Users/x/Src.xlsx"));
  links.push_back(LoadedLink(2, "Plain.xlsx", ""));
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=[1]Data!A1", Workbook::FormulaSource::kEntered),
            "='/Users/x/[Src.xlsx]Data'!A1");
  EXPECT_EQ(wb.ingest_stored_formula("=[1]!Total", Workbook::FormulaSource::kEntered), "='/Users/x/Src.xlsx'!Total");
  EXPECT_EQ(wb.ingest_stored_formula("=SUM([2]Data!A1:B2)", Workbook::FormulaSource::kEntered),
            "=SUM([Plain.xlsx]Data!A1:B2)");
  EXPECT_EQ(wb.ingest_stored_formula("=[2]!Total", Workbook::FormulaSource::kEntered), "=Plain.xlsx!Total");
  // Storage prefixes are stripped as before, and `[0]!Name` stays as is.
  EXPECT_EQ(wb.ingest_stored_formula("=_xlfn.XLOOKUP([0]!Key,A:A,B:B)", Workbook::FormulaSource::kEntered),
            "=XLOOKUP([0]!Key,A:A,B:B)");
  EXPECT_EQ(wb.external_links().size(), 2U);
}

TEST(WorkbookExternalBind, IngestReadsABookLevelRefErrorOnTheBookStemSheet) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "ShortBook1.xlsx", ""));
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=COUNTA([1]!#REF!)", Workbook::FormulaSource::kEntered),
            "=COUNTA([ShortBook1.xlsx]ShortBook1!#REF!)");
  const std::vector<std::string> expected{"Data", "ShortBook1"};
  EXPECT_EQ(wb.external_links()[0].book.sheet_names, expected);
  EXPECT_EQ(wb.ingest_stored_formula("=[1]ShortBook1!A1", Workbook::FormulaSource::kEntered),
            "=[ShortBook1.xlsx]ShortBook1!A1");
  EXPECT_EQ(wb.external_links()[0].book.sheet_names, expected);
}

TEST(WorkbookExternalBind, ABookLevelRefErrorSheetIsCutToThirtyOneCharacters) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "LongExternalWorkbookNameForProbe1.xlsx", ""));
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=[1]!#REF!", Workbook::FormulaSource::kEntered),
            "=[LongExternalWorkbookNameForProbe1.xlsx]LongExternalWorkbookNameForProb!#REF!");
  const std::vector<std::string> expected{"Data", "LongExternalWorkbookNameForProb"};
  EXPECT_EQ(wb.external_links()[0].book.sheet_names, expected);
}

TEST(WorkbookExternalBind, IngestOfAnIndexWithNoLinkCreatesOne) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=[3]Data!A1", Workbook::FormulaSource::kEntered), "=[3]Data!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[1].index, 2U);
  EXPECT_EQ(wb.external_links()[1].target, "3");
}

TEST(WorkbookExternalBind, TheIndexerMapsABookToItsLink) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", "/Users/x/Src.xlsx"));
  wb.set_external_links(std::move(links));
  SetFormula(wb, 0, 0, "=[New.xlsx]S!A1");
  const parser::ExternalBookIndexer indexer = wb.external_book_indexer();
  EXPECT_EQ(indexer.index(indexer.ctx, "", "src.xlsx"), 1U);
  EXPECT_EQ(indexer.index(indexer.ctx, "/Users/x/", "Src.xlsx"), 1U);
  EXPECT_EQ(indexer.index(indexer.ctx, "", "New.xlsx"), 2U);
  EXPECT_EQ(indexer.index(indexer.ctx, "", "Absent.xlsx"), 0U);
}

// ---------------------------------------------------------------------------
// Unknown sheet qualifiers bind path-missing links
// ---------------------------------------------------------------------------

// Enters `formula` at Host!(row, col) as a binding does: the entry spelling,
// then the model setter. Returns the stored text.
std::string Enter(Workbook& wb, std::uint32_t row, std::uint32_t col, const std::string& formula) {
  eval::EnteredCallSpellings calls;
  std::string stored;
  parser::ParseFailure failure;
  EXPECT_EQ(eval::enter_cell_formula(wb, calls, 0, row, col, formula, &stored, &failure), eval::EntryResult::kStored)
      << formula;
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, col, stored))) << stored;
  return stored;
}

void ExpectPathMissingLink(const ExternalLinkRecord& rec, const std::string& target,
                           const std::vector<std::string>& sheets) {
  EXPECT_EQ(rec.kind, ExternalLinkRecord::Kind::kExternalBook);
  EXPECT_EQ(rec.target, target);
  EXPECT_TRUE(rec.target_path_missing);
  EXPECT_TRUE(rec.absolute_target.empty());
  EXPECT_EQ(rec.book.sheet_names, sheets);
  for (std::uint32_t i = 0; i < rec.book.sheet_names.size(); ++i) {
    EXPECT_FALSE(rec.book.sheet_has_data(i)) << i;
  }
}

TEST(WorkbookUnknownQualifier, ACellReferenceCreatesAPathMissingLink) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1"), "=[Nope]Nope!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "=SUM(Nope!A1:B2)"), "=SUM([Nope]Nope!A1:B2)");
  EXPECT_EQ(Enter(wb, 2, 0, "=SUM(Nope!A:A,Nope!1:2)"), "=SUM([Nope]Nope!A:A,[Nope]Nope!1:2)");
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->formula_text, "=[Nope]Nope!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  EXPECT_EQ(wb.external_links()[0].index, 1U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
}

TEST(WorkbookUnknownQualifier, AQuotedQualifierKeepsItsQuotes) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "='My Sheet'!A1"), "='[My Sheet]My Sheet'!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "='Nope'!A1"), "=[Nope]Nope!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  ExpectPathMissingLink(wb.external_links()[0], "My%20Sheet", {"My Sheet"});
  ExpectPathMissingLink(wb.external_links()[1], "Nope", {"Nope"});
}

// The rels target percent-encodes a space (and `%`); the book and sheet keep the qualifier verbatim.
TEST(WorkbookUnknownQualifier, TheTargetEncodesASpace) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "='My Sheet'!A1"), "='[My Sheet]My Sheet'!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "='My Sheet'!B2"), "='[My Sheet]My Sheet'!B2");
  EXPECT_EQ(Enter(wb, 2, 0, "='[My Sheet]My Sheet'!C3"), "='[My Sheet]My Sheet'!C3");
  ASSERT_EQ(wb.external_links().size(), 1U);
  ExpectPathMissingLink(wb.external_links()[0], "My%20Sheet", {"My Sheet"});
  const ExternalLinkRecord* found = wb.find_external_link("", "My Sheet");
  ASSERT_NE(found, nullptr);
  EXPECT_EQ(found->index, 1U);

  EXPECT_EQ(Enter(wb, 3, 0, "='a-b'!A1"), "='[a-b]a-b'!A1");
  (void)Enter(wb, 4, 0, "='2024'!A1");
  ASSERT_EQ(wb.external_links().size(), 3U);
  ExpectPathMissingLink(wb.external_links()[1], "a-b", {"a-b"});
  ExpectPathMissingLink(wb.external_links()[2], "2024", {"2024"});
}

TEST(WorkbookUnknownQualifier, DistinctQualifiersGetDistinctLinksInFirstUseOrder) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1+Other!A1"), "=[Nope]Nope!A1+[Other]Other!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "=Third!B2"), "=[Third]Third!B2");
  ASSERT_EQ(wb.external_links().size(), 3U);
  EXPECT_EQ(wb.external_links()[0].target, "Nope");
  EXPECT_EQ(wb.external_links()[1].target, "Other");
  EXPECT_EQ(wb.external_links()[2].target, "Third");
  EXPECT_EQ(wb.external_links()[2].index, 3U);
}

// A qualifier folding to an existing link's book, as sheet names fold, takes
// that link's spelling.
TEST(WorkbookUnknownQualifier, AFoldedQualifierBindsTheExistingLink) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1"), "=[Nope]Nope!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "=nope!B1"), "=[Nope]Nope!B1");
  // Full-width NOPE.
  EXPECT_EQ(Enter(wb, 2, 0, "=\xEF\xBC\xAE\xEF\xBC\xAF\xEF\xBC\xB0\xEF\xBC\xA5!C1"), "=[Nope]Nope!C1");
  EXPECT_EQ(Enter(wb, 3, 0, "=NOPE!LocalName"), "=[Nope]!LocalName");
  EXPECT_EQ(wb.external_links().size(), 1U);
}

TEST(WorkbookUnknownQualifier, AFoldedQualifierInOneFormulaBindsOneLink) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1+nope!B1"), "=[Nope]Nope!A1+[Nope]Nope!B1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
}

TEST(WorkbookUnknownQualifier, ALocalSheetWinsAtEntry) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1"), "=[Nope]Nope!A1");
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Nope")));
  EXPECT_EQ(Enter(wb, 1, 0, "=nope!A1"), "=Nope!A1");
  EXPECT_EQ(Enter(wb, 2, 0, "=Nope!LocalName"), "=Nope!LocalName");
  EXPECT_EQ(wb.external_links().size(), 1U);
}

// Once bound, the link stays the formula's book when a sheet of that name
// appears later, inserted or renamed.
TEST(WorkbookUnknownQualifier, ALaterLocalSheetDoesNotRebind) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!A1"), "=[Nope]Nope!A1");
  EXPECT_EQ(Enter(wb, 1, 0, "=Gone!A1"), "=[Gone]Gone!A1");
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Nope", 0)));
  const std::size_t other = wb.add_sheet("Other");
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(static_cast<std::uint32_t>(other), "Gone")));
  const std::size_t host = wb.sheet_index_by_name("Host");
  EXPECT_EQ(wb.sheet(host).cell_at(0, 0)->formula_text, "=[Nope]Nope!A1");
  EXPECT_EQ(wb.sheet(host).cell_at(1, 0)->formula_text, "=[Gone]Gone!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
  ExpectPathMissingLink(wb.external_links()[1], "Gone", {"Gone"});
}

TEST(WorkbookUnknownQualifier, StoredTextReentersOntoTheSameLink) {
  Workbook wb = OneSheet();
  const std::string stored = Enter(wb, 0, 0, "=Nope!A1+'My Sheet'!B2+Nope!LocalName");
  EXPECT_EQ(stored, "=[Nope]Nope!A1+'[My Sheet]My Sheet'!B2+[Nope]!LocalName");
  EXPECT_EQ(Enter(wb, 1, 0, wb.sheet(0).cell_at(0, 0)->formula_text), stored);
  ASSERT_TRUE(static_cast<bool>(wb.add_sheet_entered("Nope")));
  EXPECT_EQ(Enter(wb, 2, 0, stored), stored);
  EXPECT_EQ(wb.external_links().size(), 2U);
}

TEST(WorkbookUnknownQualifier, TheBookFormEnteredDirectlyCreatesTheLink) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=[Nope]Nope!A1"), "=[Nope]Nope!A1");
  ASSERT_EQ(wb.external_links().size(), 1U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
  EXPECT_EQ(Enter(wb, 1, 0, "=Nope!B1"), "=[Nope]Nope!B1");
  EXPECT_EQ(Enter(wb, 2, 0, "=[Nope]!Rate"), "=[Nope]!Rate");
  EXPECT_EQ(Enter(wb, 3, 0, "='[My Sheet]'!Rate"), "='[My Sheet]'!Rate");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[1].target, "My%20Sheet");
  EXPECT_TRUE(wb.external_links()[1].target_path_missing);
}

TEST(WorkbookUnknownQualifier, AQualifiedNameBecomesABookScopeName) {
  Workbook wb = OneSheet();
  EXPECT_EQ(Enter(wb, 0, 0, "=Nope!LocalName"), "=[Nope]!LocalName");
  EXPECT_EQ(Enter(wb, 1, 0, "='My Sheet'!LocalName"), "='[My Sheet]'!LocalName");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[0].target, "Nope");
  EXPECT_TRUE(wb.external_links()[0].target_path_missing);
  EXPECT_EQ(wb.external_links()[1].target, "My%20Sheet");
  // The link lists the typed qualifier as a sheet and the name with no body at book scope.
  for (std::size_t i = 0; i < 2U; ++i) {
    const ExternalBook& book = wb.external_links()[i].book;
    EXPECT_EQ(book.sheet_names, std::vector<std::string>{i == 0U ? "Nope" : "My Sheet"});
    ASSERT_EQ(book.names.size(), 1U);
    EXPECT_EQ(book.names[0].name, "LocalName");
    EXPECT_EQ(book.names[0].scope_sheet, ExternalBook::kNoSheet);
    EXPECT_FALSE(book.names[0].exists);
  }
  // A call, `#REF!` and a quoted span after the qualifier stay as typed.
  EXPECT_EQ(Enter(wb, 2, 0, "=Nope!#REF!"), "=Nope!#REF!");
  EXPECT_EQ(Enter(wb, 3, 0, "=SUM('Host:Nope'!A1)"), "=SUM('Host:Nope'!A1)");
}

// `X:Y!ref` is a 3-D span only when both ends are local sheets. Otherwise `X`
// is a name and `Y!ref` a qualified reference written quoted, `Y` binding a
// link when it is no local sheet; `X` never binds one.
TEST(WorkbookUnknownQualifier, ASpanWithAnUnknownEndIsANameBeforeAQualifiedReference) {
  Workbook wb = OneSheet();
  wb.add_sheet("Data");
  EXPECT_EQ(Enter(wb, 0, 0, "=SUM(Nope:Data!A1)"), "=SUM(Nope:'Data'!A1)");
  EXPECT_EQ(Enter(wb, 3, 0, "=SUM(Host:Data!A1)"), "=SUM(Host:Data!A1)");
  EXPECT_TRUE(wb.external_links().empty());
  EXPECT_EQ(Enter(wb, 0, 1, "=SUM(Data:Nope!A1)"), "=SUM(Data:'[Nope]Nope'!A1)");
  EXPECT_EQ(Enter(wb, 1, 1, "=SUM(Data:Nope!B2)"), "=SUM(Data:'[Nope]Nope'!B2)");
  EXPECT_EQ(Enter(wb, 2, 1, "=SUM(Nope:Other!A1)"), "=SUM(Nope:'[Other]Other'!A1)");
  ASSERT_EQ(wb.external_links().size(), 2U);
  EXPECT_EQ(wb.external_links()[0].target, "Nope");
  EXPECT_EQ(wb.external_links()[1].target, "Other");
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t row : {0U, 2U}) {
    const Value v = wb.sheet(0).cell_at(row, 1)->cached_value;
    ASSERT_TRUE(v.is_error()) << row;
    EXPECT_EQ(v.as_error(), ErrorCode::Name) << row;
  }
}

// Entered after the source picker is cancelled, an unknown qualifier as an operand of `:` binds whatever the other
// operand is, as on load; a range across two books is #VALUE!.
TEST(WorkbookUnknownQualifier, AnyRangeOperandBindsAtEntry) {
  Workbook wb = OneSheet();
  wb.add_sheet("Data");
  EXPECT_EQ(Enter(wb, 0, 0, "=SUM(A1:Nope!B2)"), "=SUM(A1:'[Nope]Nope'!B2)");
  EXPECT_EQ(Enter(wb, 1, 0, "=SUM(Nope!A1:Data!B2)"), "=SUM([Nope]Nope!A1:'Data'!B2)");
  ASSERT_EQ(wb.external_links().size(), 1U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  for (std::uint32_t row : {0U, 1U}) {
    const Value v = wb.sheet(0).cell_at(row, 0)->cached_value;
    ASSERT_TRUE(v.is_error()) << row;
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << row;
  }
}

TEST(WorkbookUnknownQualifier, ADefinedNameBodyBindsTheSameWay) {
  Workbook wb = OneSheet();
  eval::EnteredCallSpellings calls;
  std::string stored;
  parser::ParseFailure failure;
  ASSERT_EQ(eval::enter_name_body(wb, calls, "=Nope!A1", &stored, &failure), eval::EntryResult::kStored);
  EXPECT_EQ(stored, "=[Nope]Nope!A1");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Far", stored)));
  ASSERT_EQ(wb.external_links().size(), 1U);
  ExpectPathMissingLink(wb.external_links()[0], "Nope", {"Nope"});
  EXPECT_EQ(EnteredSpelling(wb, "=nope!B1>0"), "=[Nope]Nope!B1>0");
}

// ---------------------------------------------------------------------------
// Unknown sheet qualifiers in a loaded book
// ---------------------------------------------------------------------------

// Authored books Mac Excel opened (`<stem>.xlsx`, unknown qualifiers and no link) and Excel's save of each
// (`<stem>_saved.xlsx`). Excel binds every unknown qualifier of a cell formula, a defined name or a table formula on
// open, numbering the links in load order: defined names before cells.

std::vector<std::uint8_t> LoadFixture(const std::string& name) {
  return test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/unknown_qualifier_load/" + name);
}

std::string XmlText(std::string_view escaped) {
  std::string out;
  for (std::size_t i = 0; i < escaped.size(); ++i) {
    bool replaced = false;
    for (const auto& [entity, ch] : {std::pair<std::string_view, char>{"&apos;", '\''},
                                     {"&quot;", '"'},
                                     {"&lt;", '<'},
                                     {"&gt;", '>'},
                                     {"&amp;", '&'}}) {
      if (escaped.substr(i, entity.size()) == entity) {
        out.push_back(ch);
        i += entity.size() - 1U;
        replaced = true;
        break;
      }
    }
    if (!replaced) {
      out.push_back(escaped[i]);
    }
  }
  return out;
}

// The text of every `<tag ...>text</tag>` in `xml`, in document order.
std::vector<std::string> ElementTexts(const std::string& xml, std::string_view tag) {
  std::vector<std::string> texts;
  const std::string open = "<" + std::string(tag);
  const std::string close = "</" + std::string(tag) + ">";
  for (std::size_t at = xml.find(open); at != std::string::npos; at = xml.find(open, at + 1U)) {
    const char next = xml[at + open.size()];
    if (next != '>' && next != ' ') {
      continue;
    }
    const std::size_t body = xml.find('>', at);
    const std::size_t end = xml.find(close, body);
    if (body == std::string::npos || xml[body - 1U] == '/' || end == std::string::npos) {
      continue;
    }
    texts.push_back(XmlText(std::string_view(xml).substr(body + 1U, end - body - 1U)));
  }
  return texts;
}

// What a save stores for the formulas and links: cell formulas, defined names, table formulas, then each link part's
// sheet names in link order.
std::vector<std::string> StoredFormulasAndLinks(const std::vector<std::uint8_t>& package) {
  std::vector<std::string> out;
  std::string xml;
  if (test::extract_part(test::span_of(package), "xl/worksheets/sheet1.xml", &xml)) {
    for (std::string& f : ElementTexts(xml, "f")) {
      out.push_back("f: " + f);
    }
  }
  if (test::extract_part(test::span_of(package), "xl/workbook.xml", &xml)) {
    for (std::string& name : ElementTexts(xml, "definedName")) {
      out.push_back("name: " + name);
    }
  }
  if (test::extract_part(test::span_of(package), "xl/tables/table1.xml", &xml)) {
    for (const char* tag : {"calculatedColumnFormula", "totalsRowFormula"}) {
      for (std::string& f : ElementTexts(xml, tag)) {
        out.push_back(std::string(tag) + ": " + f);
      }
    }
  }
  for (int n = 1;
       test::extract_part(test::span_of(package), "xl/externalLinks/externalLink" + std::to_string(n) + ".xml", &xml);
       ++n) {
    std::string sheets = "link " + std::to_string(n) + ":";
    const std::string key = "<sheetName val=\"";
    for (std::size_t at = xml.find(key); at != std::string::npos; at = xml.find(key, at + 1U)) {
      const std::size_t begin = at + key.size();
      sheets += " " + XmlText(std::string_view(xml).substr(begin, xml.find('"', begin) - begin));
    }
    out.push_back(sheets);
  }
  return out;
}

Workbook LoadAuthored(const std::string& stem) {
  const std::vector<std::uint8_t> bytes = LoadFixture(stem + ".xlsx");
  auto loaded = io::read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << stem;
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

void ExpectSavedAsExcel(const std::string& stem) {
  SCOPED_TRACE(stem);
  const Workbook wb = LoadAuthored(stem);
  const auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  EXPECT_EQ(StoredFormulasAndLinks(saved.value()), StoredFormulasAndLinks(LoadFixture(stem + "_saved.xlsx")));
}

// Mac Excel's saves of a new book where `=SUM(A1:Nope!B2)`, `=SUM(Nope!A1:Data!B2)`, `=SUM(Data:Data2!A1:Nope!B2)`
// and `=SUM(Rng:Nope!B2)` were typed into A1 and the source picker cancelled.
TEST(WorkbookUnknownQualifier, AnEnteredRangeOperandSavesAsExcelDoes) {
  struct Row {
    const char* stem;
    const char* typed;
  };
  for (const Row& row : {Row{"E1", "=SUM(A1:Nope!B2)"}, Row{"E2", "=SUM(Nope!A1:Data!B2)"},
                         Row{"E3", "=SUM(Data:Data2!A1:Nope!B2)"}, Row{"E4", "=SUM(Rng:Nope!B2)"}}) {
    SCOPED_TRACE(row.stem);
    Workbook wb = Workbook::create_empty();
    wb.add_sheet("Sheet1");
    wb.add_sheet("Data");
    if (std::string_view(row.stem) == "E3") {
      wb.add_sheet("Data2");
    }
    if (std::string_view(row.stem) == "E4") {
      ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("Rng", "Sheet1!$C$1:$C$2")));
    }
    Enter(wb, 0, 0, row.typed);
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->cached_value, Value::error(ErrorCode::Value));
    const auto saved = io::write_ooxml(wb);
    ASSERT_TRUE(static_cast<bool>(saved));
    EXPECT_EQ(StoredFormulasAndLinks(saved.value()),
              StoredFormulasAndLinks(test::read_file_bytes(
                  std::string(FORMULON_FIXTURES_DIR) + "/excel/unknown_qualifier_entry/" + row.stem + "_saved.xlsx")));
  }
}

// An unknown qualifier as an operand of `:`, whatever the other operand: `SUM(A1:'[1]Nope'!B2)`,
// `SUM([1]Nope!A1:'Data'!B2)`, `SUM(Data:Data2!A1:'[1]Nope'!B2)`, `SUM(Rng:'[1]Nope'!B2)`, then `[2]Ctl!A1`.
TEST(WorkbookUnknownQualifierLoad, ARangeOperandBindsOnLoad) {
  for (const char* stem : {"Y1", "Y2", "Y3", "Y4"}) {
    ExpectSavedAsExcel(stem);
  }
}

// A range across two books is #VALUE! through the evaluator's general rule for a range whose ends lie apart.
TEST(WorkbookUnknownQualifierLoad, ARangeAcrossTwoBooksIsAValueError) {
  for (const char* stem : {"Y1", "Y2", "Y3", "Y4"}) {
    SCOPED_TRACE(stem);
    Workbook wb = LoadAuthored(stem);
    ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
    const Value v = wb.sheet(0).cell_at(0, 0)->cached_value;
    ASSERT_TRUE(v.is_error());
    EXPECT_EQ(v.as_error(), ErrorCode::Value);
  }
}

// Link order: a defined name's qualifier (workbook or sheet scope) is bound before any cell's, so `Gone` is link 1
// and the cell's `Nope` link 2; a table formula's binds with the cells.
TEST(WorkbookUnknownQualifierLoad, LinkOrderPutsDefinedNamesBeforeCells) {
  for (const char* stem : {"LB1", "LB3", "LB4"}) {
    ExpectSavedAsExcel(stem);
  }
  ExpectSavedAsExcel("LB2");
}

// Model spelling after load is the entry spelling of the bound text.
TEST(WorkbookUnknownQualifierLoad, ALoadedBookReadsBackInTheFormulaBarSpelling) {
  const Workbook wb = LoadAuthored("LB3");
  EXPECT_EQ(wb.sheet(0).cell_at(0, 0)->formula_text, "=[Nope]Nope!A1");
  ASSERT_EQ(wb.external_links().size(), 2U);
  ExpectPathMissingLink(wb.external_links()[0], "Gone", {"Gone"});
  ExpectPathMissingLink(wb.external_links()[1], "Nope", {"Nope"});
}

// ---------------------------------------------------------------------------
// First-use order of sheet-qualified references
// ---------------------------------------------------------------------------

Workbook TwoSheets() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  return wb;
}

XlsbExternSheetEntry Xti(std::uint32_t book, std::string first, std::string last) {
  XlsbExternSheetEntry entry;
  entry.external_book = book;
  entry.first = std::move(first);
  entry.last = std::move(last);
  return entry;
}

void ExpectOrder(const Workbook& wb, const std::vector<XlsbExternSheetEntry>& want) {
  const std::vector<XlsbExternSheetEntry>& got = wb.xti_order();
  ASSERT_EQ(got.size(), want.size());
  for (std::size_t i = 0; i < want.size(); ++i) {
    EXPECT_EQ(got[i].external_book, want[i].external_book) << i;
    EXPECT_EQ(got[i].first, want[i].first) << i;
    EXPECT_EQ(got[i].last, want[i].last) << i;
  }
}

TEST(WorkbookXtiOrder, SelfAndLinkReferencesAreRecordedInEntryOrder) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "=Sheet2!A1");
  SetFormula(wb, 1, 0, "='[S1.xlsx]A'!A1");
  SetFormula(wb, 2, 0, "=SUM(Sheet1:Sheet2!A1)");
  ExpectOrder(wb, {Xti(0, "Sheet2", "Sheet2"), Xti(1, "A", "A"), Xti(0, "Sheet1", "Sheet2")});
}

TEST(WorkbookXtiOrder, CellsAndNamesShareOneSequence) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "='[S2.xlsx]A'!A1");
  ASSERT_TRUE(static_cast<bool>(wb.set_defined_name("N", "='[S1.xlsx]A'!$A$1")));
  ExpectOrder(wb, {Xti(1, "A", "A"), Xti(2, "A", "A")});
}

TEST(WorkbookXtiOrder, LinksAndSelfInterleaveByFirstUse) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "='[S1.xlsx]A'!A1");
  SetFormula(wb, 1, 0, "=SUM(Sheet1:Sheet2!A1)");
  SetFormula(wb, 2, 0, "='[S2.xlsx]A'!A1");
  ExpectOrder(wb, {Xti(1, "A", "A"), Xti(0, "Sheet1", "Sheet2"), Xti(2, "A", "A")});
}

TEST(WorkbookXtiOrder, AReferenceEnteredTwiceIsRecordedOnce) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "=Sheet2!A1");
  SetFormula(wb, 1, 0, "=Sheet2!B7+1");
  SetFormula(wb, 2, 0, "='[S1.xlsx]A'!A1");
  SetFormula(wb, 3, 0, "='[S1.xlsx]A'!C3");
  ExpectOrder(wb, {Xti(0, "Sheet2", "Sheet2"), Xti(1, "A", "A")});
}

TEST(WorkbookXtiOrder, RenamingASheetRewritesItsSelfEntries) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "=Sheet2!A1");
  SetFormula(wb, 1, 0, "='[S1.xlsx]Sheet2'!A1");
  SetFormula(wb, 2, 0, "=SUM(Sheet1:Sheet2!A1)");
  ASSERT_TRUE(static_cast<bool>(wb.rename_sheet(1, "Renamed")));
  // Another book's sheet of the same name is untouched.
  ExpectOrder(wb, {Xti(0, "Renamed", "Renamed"), Xti(1, "Sheet2", "Sheet2"), Xti(0, "Sheet1", "Renamed")});
}

TEST(WorkbookXtiOrder, RemovingASheetDropsTheSelfEntriesNamingIt) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "='[S1.xlsx]Sheet2'!A1");
  SetFormula(wb, 1, 0, "=Sheet2!A1");
  SetFormula(wb, 2, 0, "=SUM(Sheet1:Sheet2!A1)");
  SetFormula(wb, 3, 0, "=Sheet1!B1");
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(1)));
  ExpectOrder(wb, {Xti(1, "Sheet2", "Sheet2"), Xti(0, "Sheet1", "Sheet1")});
}

TEST(WorkbookXtiOrder, RemovingASpanEndpointNarrowsTheSpanInPlace) {
  Workbook wb = Workbook::create_empty();
  for (const char* name : {"Sheet1", "Sheet2", "Sheet3", "Sheet4"}) {
    wb.add_sheet(name);
  }
  SetFormula(wb, 0, 0, "='[S1.xlsx]A'!A1");
  SetFormula(wb, 1, 0, "=SUM(Sheet2:Sheet4!A1)");
  SetFormula(wb, 2, 0, "='[S2.xlsx]A'!A1");
  SetFormula(wb, 3, 0, "=Sheet3!A1");
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(3)));
  ExpectOrder(wb, {Xti(1, "A", "A"), Xti(0, "Sheet2", "Sheet3"), Xti(2, "A", "A"), Xti(0, "Sheet3", "Sheet3")});
}

TEST(WorkbookXtiOrder, SetReplacesTheWholeSequenceAndItsDedup) {
  Workbook wb = TwoSheets();
  SetFormula(wb, 0, 0, "=Sheet2!A1");
  wb.set_xti_order({Xti(1, "A", "A")});
  ExpectOrder(wb, {Xti(1, "A", "A")});
  // The replaced entry is forgotten, the installed one is not re-added.
  SetFormula(wb, 1, 0, "=Sheet2!A2");
  SetFormula(wb, 2, 0, "='[S1.xlsx]A'!A1");
  ExpectOrder(wb, {Xti(1, "A", "A"), Xti(0, "Sheet2", "Sheet2")});
  wb.set_xti_order({});
  EXPECT_TRUE(wb.xti_order().empty());
}

// `Nope!LocalName` in an authored book Mac Excel opened, against Excel's .xlsx save of it: the link lists the qualifier
// as its one sheet and the name with no body, and the cell is #NAME?. PathMissingLinkRecords compares the .xlsb save.
TEST(WorkbookUnknownQualifierLoad, ANameInAnUnlocatedBookSavesExcelsLinkPart) {
  const std::string dir = std::string(FORMULON_FIXTURES_DIR) + "/excel/external_link_path_missing/";
  auto loaded = io::read_ooxml(test::span_of(test::read_file_bytes(dir + "book_name_input.xlsx")));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  ASSERT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->cached_value, Value::error(ErrorCode::Name));

  const auto link_body = [](const std::vector<std::uint8_t>& package) {
    std::string xml;
    if (!test::extract_part(test::span_of(package), "xl/externalLinks/externalLink1.xml", &xml)) {
      return std::string("no link part");
    }
    const std::size_t from = xml.find("<sheetNames>");
    const std::size_t to = xml.find("</externalBook>");
    return from == std::string::npos || to == std::string::npos ? xml : xml.substr(from, to - from);
  };
  auto xlsx = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  EXPECT_EQ(link_body(xlsx.value()), link_body(test::read_file_bytes(dir + "book_name.xlsx")));
}

// A span into a link is stored with its ends in the link's sheet-list order, as Excel saves it and as Excel shows it
// after a reload while the source is closed. Right after entry the formula bar keeps the typed order, since names
// added by entry have no known tab position.
TEST(WorkbookExternalBind, AStoredSpanNamesItsEndsInTheLinkListOrder) {
  Workbook wb = OneSheet();
  SetFormula(wb, 0, 0, "=[Src3.xlsx]Z!A1");
  SetFormula(wb, 1, 0, "=SUM([Src3.xlsx]X:Z!A1)");
  EXPECT_EQ(wb.sheet(0).cell_at(1U, 0U)->formula_text, "=SUM([Src3.xlsx]X:Z!A1)");
  const auto saved = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  EXPECT_EQ(StoredFormulasAndLinks(saved.value()),
            (std::vector<std::string>{"f: [1]Z!A1", "f: SUM([1]Z:X!A1)", "link 1: Z X"}));

  auto loaded = io::read_ooxml(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  Workbook& reloaded = loaded.value().workbook;
  EXPECT_EQ(reloaded.sheet(0).cell_at(1U, 0U)->formula_text, "=SUM([Src3.xlsx]Z:X!A1)");
  // With the source closed, a span over sheets with no cached data is #REF!.
  ASSERT_TRUE(static_cast<bool>(reloaded.recalc_full(eval::default_registry())));
  EXPECT_EQ(reloaded.sheet(0).cell_at(1U, 0U)->cached_value, Value::error(ErrorCode::Ref));
}

// A stored span written against its link's list order reads back in that order.
TEST(WorkbookExternalBind, IngestOrdersASpanByTheLinkList) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src3.xlsx", ""));
  links[0].book.sheet_names = {"Z", "X"};
  links[0].book.sheet_data = {false, false};
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=SUM([1]X:Z!A1)", Workbook::FormulaSource::kEntered), "=SUM([Src3.xlsx]Z:X!A1)");
  EXPECT_EQ(wb.ingest_stored_formula("=SUM([1]Z:X!A1)", Workbook::FormulaSource::kEntered), "=SUM([Src3.xlsx]Z:X!A1)");
  // A span of this workbook is not reordered by a link's list.
  wb.add_sheet("Z");
  wb.add_sheet("X");
  EXPECT_EQ(wb.ingest_stored_formula("=SUM(X:Z!A1)", Workbook::FormulaSource::kEntered), "=SUM(X:Z!A1)");
}

}  // namespace
}  // namespace formulon
