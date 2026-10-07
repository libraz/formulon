//
// Values of local references whose sheet qualifier is written bare but is
// shaped like a cell, a number or an R1C1 reference (`S2!A1`, `2024!A1`,
// `R1C1!A1`), in a workbook where those sheets exist.
//
// Each sheet's A1 holds a distinct power of two, so a sum names exactly the
// sheets it covered; B2 holds ten times that, and every sheet carries a
// sheet-local name `LocalName` over its own A1.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// Workbook order; `Host` (index 0) holds the formulas.
const std::vector<std::string>& SheetNames() {
  static const std::vector<std::string> names = {"Host", "Data", "S2", "2024", "TRUE", "R1C1",
                                                 "R",    "C",    "A1", "XFE1", "Zed"};
  return names;
}

double A1Value(std::size_t sheet_index) {
  return static_cast<double>(1U << (sheet_index - 1U));
}

class LocalSheetQualifierEval : public ::testing::Test {
 protected:
  void SetUp() override {
    std::vector<DefinedName> names;
    for (std::size_t i = 0; i < SheetNames().size(); ++i) {
      Sheet& sheet = wb_.sheet(wb_.add_sheet(SheetNames()[i]));
      if (i == 0) {
        continue;
      }
      sheet.set_cell_value(0, 0, Value::number(A1Value(i)));
      sheet.set_cell_value(1, 1, Value::number(10.0 * A1Value(i)));
      names.push_back(
          DefinedName{"LocalName", "'" + SheetNames()[i] + "'!$A$1", static_cast<std::int32_t>(i), false, ""});
    }
    wb_.set_defined_names(std::move(names));
  }

  void ExpectNumber(const std::string& formula, double expected) {
    ASSERT_TRUE(static_cast<bool>(wb_.set_cell_formula(0, 0, 0, formula))) << formula;
    auto recalc_or = wb_.recalc(eval::default_registry());
    ASSERT_TRUE(static_cast<bool>(recalc_or)) << formula;
    const Value v = wb_.sheet(0).cell_at(0, 0)->cached_value;
    ASSERT_TRUE(v.is_number()) << formula << " -> " << v.debug_to_string();
    EXPECT_DOUBLE_EQ(v.as_number(), expected) << formula;
  }

  Workbook wb_ = Workbook::create_empty();
};

TEST_F(LocalSheetQualifierEval, EveryQualifierShapeReadsItsSheet) {
  for (std::size_t i = 1; i < SheetNames().size(); ++i) {
    const std::string& name = SheetNames()[i];
    if (name == "TRUE" || name == "Data" || name == "Zed") {
      continue;  // Not one of the cell-like shapes; TRUE needs quoting.
    }
    const double a1 = A1Value(i);
    ExpectNumber("=" + name + "!A1", a1);
    ExpectNumber("=" + name + "!$A$1", a1);
    ExpectNumber("=SUM(" + name + "!A:A)", a1);
    ExpectNumber("=SUM(" + name + "!1:1)", a1);
    ExpectNumber("=SUM(" + name + "!A1:B2)", 11.0 * a1);
    ExpectNumber("=" + name + "!LocalName", a1);
    ExpectNumber("=" + name + "! B2", 10.0 * a1);
  }
}

TEST_F(LocalSheetQualifierEval, ThreeDEndpoints) {
  // Data(1) S2(2)
  ExpectNumber("=SUM(Data:S2!A1)", 3.0);
  // 2024(4) TRUE(8) R1C1(16) R(32) C(64) A1(128) XFE1(256) Zed(512)
  ExpectNumber("=SUM(2024:Zed!A1)", 1020.0);
  // Data(1) S2(2) 2024(4) TRUE(8)
  ExpectNumber("=SUM(Data:TRUE!A1)", 15.0);
  ExpectNumber("=SUM(R1C1:XFE1!A1)", 496.0);
}

TEST_F(LocalSheetQualifierEval, QuotedTrueStillReadsItsSheet) {
  ExpectNumber("='TRUE'!A1", 8.0);
}

}  // namespace
}  // namespace formulon
