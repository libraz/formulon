//
// Every external book a formula in the model names has an external-link
// record: entering one creates the record, re-entering it reuses it, and a
// loaded `[N]` formula reads back in the formula bar's spelling.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "parser/ast_format.h"
#include "sheet_passthrough.h"
#include "workbook.h"

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

}  // namespace
}  // namespace formulon
