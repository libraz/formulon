//
// Every external book a formula in the model names has an external-link
// record: entering one creates the record, re-entering it reuses it, and a
// loaded `[N]` formula reads back in the formula bar's spelling.

#include <cstddef>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "parser/ast_format.h"
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
  EXPECT_EQ(wb.ingest_stored_formula("=[1]Data!A1"), "='/Users/x/[Src.xlsx]Data'!A1");
  EXPECT_EQ(wb.ingest_stored_formula("=[1]!Total"), "='/Users/x/Src.xlsx'!Total");
  EXPECT_EQ(wb.ingest_stored_formula("=SUM([2]Data!A1:B2)"), "=SUM([Plain.xlsx]Data!A1:B2)");
  EXPECT_EQ(wb.ingest_stored_formula("=[2]!Total"), "=Plain.xlsx!Total");
  // Storage prefixes are stripped as before, and `[0]!Name` stays as is.
  EXPECT_EQ(wb.ingest_stored_formula("=_xlfn.XLOOKUP([0]!Key,A:A,B:B)"), "=XLOOKUP([0]!Key,A:A,B:B)");
  EXPECT_EQ(wb.external_links().size(), 2U);
}

TEST(WorkbookExternalBind, IngestOfAnIndexWithNoLinkCreatesOne) {
  Workbook wb = OneSheet();
  std::vector<ExternalLinkRecord> links;
  links.push_back(LoadedLink(1, "Src.xlsx", ""));
  wb.set_external_links(std::move(links));
  EXPECT_EQ(wb.ingest_stored_formula("=[3]Data!A1"), "=[3]Data!A1");
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

}  // namespace
}  // namespace formulon
