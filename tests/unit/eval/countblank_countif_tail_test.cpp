//
// COUNTBLANK / COUNTIF / COUNTIFS over whole-axis references count the
// unpopulated rest of the axis, and the range-taking conditional aggregates
// reject a reference into a supporting workbook as a closed book (#VALUE!).
//
// Local fixture, sheet L: A1:A5 = "k1","k2","k3",(blank),"k5"; B1:B5 = 10..50;
// column C empty. Formulas sit on sheet H. Expected values are Excel 365
// ja-JP measurements.

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "eval/tree_walker.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

void CacheNumber(ExternalBook& book, std::uint32_t row, std::uint32_t col, double value) {
  ExternalCell cell;
  cell.value = Value::number(value);
  book.cells.emplace(ExternalBook::cell_key(0, row, col), std::move(cell));
}

void CacheText(ExternalBook& book, std::uint32_t row, std::uint32_t col, const char* text) {
  ExternalCell cell;
  cell.value = Value::text(text);
  book.cells.emplace(ExternalBook::cell_key(0, row, col), std::move(cell));
}

// `Src.xlsx` sheet S: A1,A2,A3,A5 text (A4 uncached), B1:B5 = 10..50.
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
  book.sheet_names = {"S"};
  book.sheet_data = {true};
  CacheText(book, 0, 0, "k1");
  CacheText(book, 1, 0, "k2");
  CacheText(book, 2, 0, "k3");
  CacheText(book, 4, 0, "k5");
  for (std::uint32_t r = 0; r < 5; ++r) {
    CacheNumber(book, r, 1, 10.0 * (r + 1));
  }
  return rec;
}

class CountBlankCountIfTail : public ::testing::Test {
 protected:
  void SetUp() override {
    wb_.add_sheet("H");
    wb_.add_sheet("L");
    Sheet& l = wb_.sheet(1);
    const char* keys[] = {"k1", "k2", "k3", nullptr, "k5"};
    for (std::uint32_t r = 0; r < 5; ++r) {
      if (keys[r] != nullptr) {
        l.set_cell_value(r, 0, Value::text(keys[r]));
      }
      l.set_cell_value(r, 1, Value::number(10.0 * (r + 1)));
    }
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

  void ExpectValueError(const std::string& formula) {
    const Value v = Eval(formula);
    ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), ErrorCode::Value) << formula;
  }

  Workbook wb_ = Workbook::create_empty();
};

TEST_F(CountBlankCountIfTail, CountBlankCountsTheUnpopulatedAxis) {
  ExpectNumber("=COUNTBLANK(L!C:C)", 1048576.0);
  ExpectNumber("=COUNTBLANK(L!B:B)", 1048571.0);
  ExpectNumber("=COUNTBLANK(L!A1:A5)", 1.0);
  ExpectNumber("=COUNTBLANK(L!A:A)", 1048572.0);
  ExpectNumber("=COUNTBLANK(L!1:1)", 16382.0);
  ExpectNumber("=COUNTBLANK(L!A:B)", 2097143.0);
}

TEST_F(CountBlankCountIfTail, CountIfCountsTheTailWhenABlankMatches) {
  ExpectNumber("=COUNTIF(L!A:A,\"<>k1\")", 1048575.0);
  ExpectNumber("=COUNTIF(L!A:A,\"\")", 1048572.0);
  ExpectNumber("=COUNTIF(L!A:A,\"<>\")", 4.0);
  ExpectNumber("=COUNTIF(L!B:B,\"<20\")", 1.0);
  ExpectNumber("=COUNTIF(L!B:B,\"<>20\")", 1048575.0);
}

TEST_F(CountBlankCountIfTail, CountIfsCountsTheCommonTailWhenEveryCriterionMatchesABlank) {
  ExpectNumber("=COUNTIFS(L!A:A,\"\",L!B:B,\"\")", 1048571.0);
  ExpectNumber("=COUNTIFS(L!A:A,\"<>k1\",L!B:B,\"<>10\")", 1048575.0);
  ExpectNumber("=COUNTIFS(L!A:A,\"k2\",L!B:B,\"\")", 0.0);
}

TEST_F(CountBlankCountIfTail, SumIfAndAverageIfIgnoreTheTail) {
  ExpectNumber("=SUMIF(L!A:A,\"\",L!B:B)", 40.0);
  ExpectNumber("=AVERAGEIF(L!A:A,\"<>k1\",L!B:B)", 35.0);
}

TEST_F(CountBlankCountIfTail, CountBlankOverNonReferenceArgumentsCountsEvaluatedValues) {
  ExpectNumber("=COUNTBLANK(\"\")", 1.0);
  ExpectNumber("=COUNTBLANK(0,\"\",\"x\",\"\")", 2.0);
}

TEST_F(CountBlankCountIfTail, ExternalReferencesAreAClosedBook) {
  ExpectValueError("=COUNTBLANK([Src.xlsx]S!A1:A5)");
  ExpectValueError("=COUNTIF([Src.xlsx]S!A1:A5,\"k2\")");
  ExpectValueError("=COUNTIFS([Src.xlsx]S!A1:A5,\"k2\",[Src.xlsx]S!B1:B5,\">0\")");
  ExpectValueError("=SUMIF([Src.xlsx]S!A:A,\"k2\",[Src.xlsx]S!B:B)");
  ExpectValueError("=SUMIFS([Src.xlsx]S!B:B,[Src.xlsx]S!A:A,\"k2\")");
  ExpectValueError("=AVERAGEIF([Src.xlsx]S!A:A,\"k2\",[Src.xlsx]S!B:B)");
  ExpectValueError("=AVERAGEIFS([Src.xlsx]S!B:B,[Src.xlsx]S!A:A,\"k2\")");
  ExpectValueError("=MAXIFS([Src.xlsx]S!B:B,[Src.xlsx]S!A:A,\"k2\")");
  ExpectValueError("=MINIFS([Src.xlsx]S!B:B,[Src.xlsx]S!A:A,\"k2\")");
}

TEST_F(CountBlankCountIfTail, ALocalRangeWithAnExternalValueRangeIsStillRejected) {
  ExpectValueError("=SUMIF(L!A1:A5,\"k2\",[Src.xlsx]S!B1:B5)");
  ExpectValueError("=SUMIFS(L!B1:B5,L!A1:A5,\"k2\",[Src.xlsx]S!A1:A5,\"k2\")");
}

}  // namespace
}  // namespace formulon
