//
// A whole column held by a defined name or a LET binding keeps its declared
// shape: `MyCol+0` (MyCol = L!$B:$B) is 1048576x1, a range-aware aggregate
// reads only the populated head of `MyCol`, and a derived whole column handed
// to an eager aggregate is expanded. A LET-bound external reference reads the
// external column and is a reference to the range-criteria functions.
//
// Sheet L: A1:A5 = "k1","k2","k3",(blank),"k5"; B1:B5 = 10..50; column C is
// empty. Formulas sit on sheet H. `Src.xlsx` is a closed-book cache of sheet
// S with the same layout (A4 not cached). Expected values are the Mac Excel
// 365 ja-JP measurements recorded in tests/oracle/cases/whole_axis_spill.yaml.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/adhoc_eval.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "external_book.h"
#include "external_link.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

constexpr std::uint32_t kH = 0U;  // Sheet index of H.
constexpr std::uint32_t kL = 1U;  // Sheet index of L.
constexpr std::uint32_t kHCol = 7U;
constexpr std::uint32_t kLastRow = Sheet::kMaxRows - 1U;

// One grid column of `Value`s; a derived array built densely cannot come in
// under it.
constexpr std::size_t kWholeColumnBytes = static_cast<std::size_t>(Sheet::kMaxRows) * sizeof(Value);

ExternalLinkRecord SourceLink() {
  ExternalLinkRecord rec;
  rec.index = 1;
  rec.rel_id = "rId3";
  rec.part_path = "xl/externalLinks/externalLink1.xml";
  rec.body_rel_id = "rId1";
  rec.target = "Src.xlsx";
  rec.kind = ExternalLinkRecord::Kind::kExternalBook;
  ExternalBook& book = rec.book;
  book.sheet_names = {"S"};
  book.sheet_data = {true};
  const char* const keys[] = {"k1", "k2", "k3", nullptr, "k5"};
  for (std::uint32_t row = 0; row < 5U; ++row) {
    if (keys[row] != nullptr) {
      ExternalCell cell;
      cell.value = Value::text("");
      cell.text = keys[row];
      book.cells.emplace(ExternalBook::cell_key(0U, row, 0U), std::move(cell));
    }
    ExternalCell number;
    number.value = Value::number(10.0 * (row + 1U));
    book.cells.emplace(ExternalBook::cell_key(0U, row, 1U), std::move(number));
  }
  return rec;
}

class WholeAxisNames : public ::testing::Test {
 protected:
  WholeAxisNames() : wb_(Workbook::create()) {
    EXPECT_TRUE(static_cast<bool>(wb_.rename_sheet(0U, "H")));
    wb_.add_sheet("L");
    const char* const keys[] = {"k1", "k2", "k3", nullptr, "k5"};
    for (std::uint32_t r = 0; r < 5U; ++r) {
      if (keys[r] != nullptr) {
        EXPECT_TRUE(static_cast<bool>(wb_.set_cell_text(kL, r, 0U, keys[r])));
      }
      EXPECT_TRUE(static_cast<bool>(wb_.set_cell_value(kL, r, 1U, Value::number(10.0 * (r + 1U)))));
    }
    wb_.set_defined_names(
        {DefinedName{"MyCol", "L!$B:$B", -1, false, ""}, DefinedName{"MyCols", "L!$A:$C", -1, false, ""}});
    std::vector<ExternalLinkRecord> links;
    links.push_back(SourceLink());
    wb_.set_external_links(std::move(links));
  }

  // Stores `formula` at (`row`, `col`) of H and recalcs.
  void Put(std::uint32_t row, std::uint32_t col, const char* formula) {
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_formula(kH, row, col, formula)));
    ASSERT_TRUE(static_cast<bool>(wb_.recalc(default_registry())));
  }

  Value At(std::uint32_t row, std::uint32_t col) { return wb_.sheet(kH).resolve_cell_value(row, col); }

  // Read-only evaluation of `formula` anchored at H1, reporting the arena
  // bytes it used.
  Value Adhoc(const std::string& formula, std::size_t* out_bytes = nullptr) {
    Arena arena;
    const Value v = evaluate_formula_text_array(wb_, wb_.sheet(kH), 0U, kHCol, formula, arena, default_registry());
    if (out_bytes != nullptr) {
      *out_bytes = arena.bytes_used();
    }
    return v;
  }

  void ExpectNumber(const std::string& formula, double expected) {
    const Value v = Adhoc(formula);
    ASSERT_TRUE(v.is_number()) << formula << " -> " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), expected) << formula;
  }

  void ExpectError(const std::string& formula, ErrorCode expected) {
    const Value v = Adhoc(formula);
    ASSERT_TRUE(v.is_error()) << formula << " -> " << v.debug_to_string();
    EXPECT_EQ(v.as_error(), expected) << formula;
  }

  // `formula` at H1 spills 1048576x1 with B's values on top and 0 below.
  void ExpectColumnSpill(const char* formula) {
    Put(0U, kHCol, formula);
    const SpillRegion* region = wb_.sheet(kH).spill_region_at_anchor(0U, kHCol);
    ASSERT_NE(region, nullptr) << formula;
    EXPECT_EQ(region->rows, Sheet::kMaxRows) << formula;
    EXPECT_EQ(region->cols, 1U) << formula;
    EXPECT_EQ(At(0U, kHCol).as_number(), 10.0) << formula;
    EXPECT_EQ(At(4U, kHCol).as_number(), 50.0) << formula;
    EXPECT_EQ(At(5U, kHCol).as_number(), 0.0) << formula;
    EXPECT_EQ(At(kLastRow, kHCol).as_number(), 0.0) << formula;
  }

  Workbook wb_;
};

TEST_F(WholeAxisNames, AggregateOverAWholeColumnNameReadsItsCells) {
  ExpectNumber("=SUM(MyCol)", 150.0);
}

TEST_F(WholeAxisNames, AggregateOverAMultiColumnNameReadsItsCells) {
  ExpectNumber("=SUM(MyCols)", 150.0);
}

TEST_F(WholeAxisNames, NamePlusZeroSpillsDeclaredHeight) {
  ExpectColumnSpill("=MyCol+0");
}

TEST_F(WholeAxisNames, NamePlusZeroKeepsItsDeclaredRowCount) {
  ExpectNumber("=ROWS(MyCol+0)", 1048576.0);
}

TEST_F(WholeAxisNames, SumproductOverNamePlusZeroStaysProportionalToTheData) {
  std::size_t bytes = 0;
  const Value v = Adhoc("=SUMPRODUCT(MyCol+0)", &bytes);
  ASSERT_TRUE(v.is_number()) << v.debug_to_string();
  EXPECT_DOUBLE_EQ(v.as_number(), 150.0);
  EXPECT_LT(bytes, std::size_t{1} << 20U);
  EXPECT_LT(bytes, kWholeColumnBytes);
}

TEST_F(WholeAxisNames, EagerAggregatesOverDerivedWholeColumns) {
  ExpectNumber("=SUM(L!B:B*1)", 150.0);
  ExpectNumber("=SUM((L!A:A=\"k2\")*L!B:B)", 20.0);
  ExpectNumber("=SUMPRODUCT((L!A:A=\"k2\")*L!B:B)", 20.0);
  ExpectNumber("=MAX(L!B:B-1)", 49.0);
  ExpectNumber("=COUNT(L!B:B+0)", 1048576.0);
  ExpectNumber("=AVERAGE(IF(L!B:B>15,L!B:B))", 35.0);
}

TEST_F(WholeAxisNames, LetBoundExternalColumnSpillsDeclaredHeight) {
  ExpectColumnSpill("=LET(x,[Src.xlsx]S!B:B,x)");
}

TEST_F(WholeAxisNames, LetBoundExternalRangeIsValueToRangeCriteriaFunctions) {
  ExpectError("=LET(x,[Src.xlsx]S!A1:A5,COUNTIF(x,\"k2\"))", ErrorCode::Value);
  ExpectError("=LET(x,[Src.xlsx]S!A:A,COUNTIF(x,\"k2\"))", ErrorCode::Value);
  ExpectError("=LET(x,[Src.xlsx]S!A1:A5,COUNTBLANK(x))", ErrorCode::Value);
  ExpectError("=LET(a,[Src.xlsx]S!A:A,b,[Src.xlsx]S!B:B,SUMIF(a,\"k2\",b))", ErrorCode::Value);
}

}  // namespace
}  // namespace eval
}  // namespace formulon
