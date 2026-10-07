//
// Evaluation of cross-workbook references written by book name, against
// the cache of the external link the name matches.
//
// The rules under test:
//   * a cached cell reads its value, an uncached one in a cached sheet 0;
//   * a sheet the link lists without cached data reads #REF!, as does a
//     sheet the link does not list;
//   * a name the link does not declare reads #NAME?, and a sheet-local
//     name matches only under its own sheet;
//   * 3-D, whole-column and whole-row forms evaluate like their local
//     counterparts, the latter clipped to the cached extent;
//   * book and sheet names match without regard to ASCII case.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

void Cache(ExternalBook& book, std::uint32_t sheet, std::uint32_t row, std::uint32_t col, double value) {
  ExternalCell cell;
  cell.value = Value::number(value);
  book.cells.emplace(ExternalBook::cell_key(sheet, row, col), std::move(cell));
}

// `Src.xlsx` (absolute `/Users/x/Src.xlsx`) with sheets
//   Data   A1=10 A3=30 B2=5    (cached)
//   Mid    A1=100              (cached)
//   Last   A1=1000             (cached)
//   Empty                      (listed, no cached sheetData)
// and names Total (book scope, Data!A3), Local (local to Data, Data!B2),
// Block (book scope, Data!A1:A3).
ExternalLinkRecord SourceLink() {
  ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId3";
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "Src.xlsx";
  rec.absolute_target = "/Users/x/Src.xlsx";
  rec.absolute_rel_id = "rId2";
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  ExternalBook& book = rec.book;
  book.sheet_names = {"Data", "Mid", "Last", "Empty"};
  book.sheet_data = {true, true, true, false};
  Cache(book, 0, 0, 0, 10.0);
  Cache(book, 0, 2, 0, 30.0);
  Cache(book, 0, 1, 1, 5.0);
  Cache(book, 1, 0, 0, 100.0);
  Cache(book, 2, 0, 0, 1000.0);
  ExternalBookName total;
  total.name = "Total";
  total.sheet = 0;
  total.row = 2;
  total.row_end = 2;
  total.resolvable = true;
  book.names.push_back(total);
  ExternalBookName local;
  local.name = "Local";
  local.scope_sheet = 0;
  local.sheet = 0;
  local.row = 1;
  local.col = 1;
  local.row_end = 1;
  local.col_end = 1;
  local.resolvable = true;
  book.names.push_back(local);
  ExternalBookName block;
  block.name = "Block";
  block.sheet = 0;
  block.row_end = 2;
  block.is_range = true;
  block.resolvable = true;
  book.names.push_back(block);
  return rec;
}

class ExternalRefEval : public ::testing::Test {
 protected:
  void SetUp() override {
    wb_.add_sheet("Host");
    std::vector<ExternalLinkRecord> links;
    links.push_back(SourceLink());
    wb_.set_external_links(std::move(links));
  }

  Value Eval(const std::string& formula) {
    EXPECT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 0, 0, formula))) << formula;
    auto recalc_or = wb_.recalc(eval::default_registry());
    EXPECT_TRUE(static_cast<bool>(recalc_or)) << formula;
    const Cell* cell = wb_.sheet(0).cell_at(0, 0);
    return cell == nullptr ? Value::blank() : cell->cached_value;
  }

  void ExpectNumber(const std::string& formula, double expected) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_number()) << formula << " -> " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), expected) << formula;
  }

  void ExpectError(const std::string& formula, ErrorCode expected) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), expected) << formula;
  }

  Workbook wb_ = Workbook::create_empty();
};

TEST_F(ExternalRefEval, ACachedCellReadsItsValue) {
  ExpectNumber("=[Src.xlsx]Data!A1", 10.0);
  ExpectNumber("=[Src.xlsx]Data!$A$3", 30.0);
  ExpectNumber("='/Users/x/[Src.xlsx]Data'!A1", 10.0);
  ExpectNumber("=SUM([Src.xlsx]Data!A1:A3)", 40.0);
}

TEST_F(ExternalRefEval, AnUncachedCellInACachedSheetReadsZero) {
  ExpectNumber("=[Src.xlsx]Data!Z99", 0.0);
}

TEST_F(ExternalRefEval, ASheetWithoutCachedDataReadsRef) {
  ExpectError("=[Src.xlsx]Empty!A1", ErrorCode::Ref);
}

TEST_F(ExternalRefEval, AnUnlistedSheetReadsRef) {
  ExpectError("=[Src.xlsx]Nowhere!A1", ErrorCode::Ref);
}

TEST_F(ExternalRefEval, ANewlyLinkedBookReadsRef) {
  ExpectError("=[New.xlsx]Data!A1", ErrorCode::Ref);
  EXPECT_EQ(wb_.external_links().size(), 2U);
}

TEST_F(ExternalRefEval, AnAbsentNameReadsName) {
  ExpectError("=Src.xlsx!Nothing", ErrorCode::Name);
  ExpectError("=[Src.xlsx]Data!Nothing", ErrorCode::Name);
}

TEST_F(ExternalRefEval, BookScopeNames) {
  ExpectNumber("=Src.xlsx!Total", 30.0);
  ExpectNumber("='/Users/x/Src.xlsx'!Total", 30.0);
  ExpectNumber("=SUM(Src.xlsx!Block)", 40.0);
  // A sheet-local name is not visible at book scope.
  ExpectError("=Src.xlsx!Local", ErrorCode::Name);
}

TEST_F(ExternalRefEval, SheetLocalNames) {
  ExpectNumber("=[Src.xlsx]Data!Local", 5.0);
  ExpectNumber("=[Src.xlsx]data!local", 5.0);
  // Local to another sheet, and a book-scope name under a sheet.
  ExpectError("=[Src.xlsx]Mid!Local", ErrorCode::Name);
  ExpectError("=[Src.xlsx]Data!Total", ErrorCode::Name);
}

TEST_F(ExternalRefEval, ThreeDSpansTheLinksSheetOrder) {
  ExpectNumber("=SUM([Src.xlsx]Data:Last!A1)", 1110.0);
  ExpectNumber("=SUM('[Src.xlsx]Mid:Last'!A1)", 1100.0);
  ExpectNumber("=SUM([Src.xlsx]Last:Data!A1)", 1110.0);
  ExpectNumber("=SUM([Src.xlsx]Data:Last!A1:A3)", 1140.0);
}

TEST_F(ExternalRefEval, ThreeDOverASheetWithoutDataReadsRef) {
  ExpectError("=SUM([Src.xlsx]Data:Empty!A1)", ErrorCode::Ref);
  ExpectError("=SUM([Src.xlsx]Data:Nowhere!A1)", ErrorCode::Ref);
}

TEST_F(ExternalRefEval, ThreeDInScalarContextMatchesLocalThreeD) {
  // A local `Sheet1:Sheet2!A1` in scalar context is #VALUE!, or #REF! when
  // an endpoint names no sheet.
  ExpectError("=[Src.xlsx]Data:Last!A1", ErrorCode::Value);
  ExpectError("=[Src.xlsx]Data:Nowhere!A1", ErrorCode::Ref);
}

TEST_F(ExternalRefEval, WholeColumnsAndRowsClipToTheCachedExtent) {
  ExpectNumber("=SUM([Src.xlsx]Data!A:A)", 40.0);
  ExpectNumber("=SUM([Src.xlsx]Data!A:B)", 45.0);
  ExpectNumber("=SUM([Src.xlsx]Data!2:2)", 5.0);
  ExpectNumber("=SUM([Src.xlsx]Data!1:3)", 45.0);
  ExpectError("=SUM([Src.xlsx]Empty!A:A)", ErrorCode::Ref);
}

TEST_F(ExternalRefEval, BookAndSheetMatchWithoutCase) {
  ExpectNumber("=[src.XLSX]data!A1", 10.0);
  ExpectNumber("=[SRC.xlsx]DATA!A3", 30.0);
  EXPECT_EQ(wb_.external_links().size(), 1U);
  EXPECT_EQ(wb_.external_links()[0].book.sheet_names.size(), 4U);
}

TEST_F(ExternalRefEval, ADecimalBookIsANameNotALinkIndex) {
  ExpectError("=[1]Data!A1", ErrorCode::Ref);
  ASSERT_EQ(wb_.external_links().size(), 2U);
  EXPECT_EQ(wb_.external_links()[1].target, "1");
}

TEST_F(ExternalRefEval, TheSelfBookNameStillResolvesLocally) {
  ASSERT_TRUE(static_cast<bool>(wb_.set_defined_name("Rate", "7")));
  ExpectNumber("=[0]!Rate", 7.0);
  EXPECT_EQ(wb_.external_links().size(), 1U);
}

}  // namespace
}  // namespace formulon
