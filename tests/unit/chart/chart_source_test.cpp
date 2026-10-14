#include "chart/chart_source.h"

#include <string>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "utils/error.h"
#include "workbook.h"

namespace formulon {
namespace {

/// Sheet1 (0), 'My Data' (1), Hid (2); names WbVals, Sheet1!LocVals, DynVals, Joined, WholeB, Const, _xlchart.v1.0.
Workbook Book() {
  Workbook wb = Workbook::create();
  wb.add_sheet("My Data");
  wb.add_sheet("Hid");
  Sheet& s = wb.sheet(0);
  for (std::uint32_t r = 1; r <= 4; ++r) {
    s.set_cell_value(r, 1, Value::number(r));
  }
  std::vector<DefinedName> names;
  names.push_back(DefinedName{"WbVals", "Sheet1!$B$2:$B$5", -1, false, ""});
  names.push_back(DefinedName{"LocVals", "Sheet1!$C$2:$C$5", 0, false, ""});
  names.push_back(DefinedName{"DynVals", "OFFSET(Sheet1!$B$2,0,0,COUNT(Sheet1!$B$2:$B$100),1)", -1, false, ""});
  names.push_back(DefinedName{"Joined", "(Sheet1!$A$1,Sheet1!$C$3:$C$4)", -1, false, ""});
  names.push_back(DefinedName{"WholeB", "Sheet1!$B:$B", -1, false, ""});
  names.push_back(DefinedName{"Const", "42", -1, false, ""});
  names.push_back(DefinedName{"_xlchart.v1.0", "Sheet1!$A$2:$A$5", -1, true, ""});
  names.push_back(DefinedName{"HostOnly", "Sheet1!$D$1", 1, false, ""});
  wb.set_defined_names(std::move(names));
  return wb;
}

std::string Describe(const ResolvedSource& src) {
  std::string out;
  for (const SheetRect& r : src.areas) {
    out += (out.empty() ? "" : ",") + std::to_string(r.sheet) + ":" + std::to_string(r.top) + "," +
           std::to_string(r.left) + "-" + std::to_string(r.bottom) + "," + std::to_string(r.right);
  }
  return out;
}

std::string Resolve(const Workbook& wb, std::size_t host, const std::string& formula) {
  auto got = resolve_chart_source(wb, host, formula, eval::default_registry());
  if (!got) {
    return "error";
  }
  return Describe(got.value());
}

TEST(ChartSource, ResolvesEveryReferenceShapeToRectangles) {
  const Workbook wb = Book();
  struct Case {
    std::size_t host;
    const char* formula;
    const char* expected;
  };
  const std::vector<Case> cases = {
      {0, "Sheet1!$B$2:$B$5", "0:1,1-4,1"},
      {0, "Sheet1!$A$1", "0:0,0-0,0"},
      {0, "'My Data'!$A$2:$B$4", "1:1,0-3,1"},
      {0, "Hid!$B$2", "2:1,1-1,1"},
      {1, "$C$3:$D$4", "1:2,2-3,3"},
      {0, "=Sheet1!$B$2", "0:1,1-1,1"},
      // A parenthesised union keeps its element order.
      {0, "(Sheet1!$B$5,Sheet1!$A$2:$A$3,'My Data'!$C$1)", "0:4,1-4,1,0:1,0-2,0,1:0,2-0,2"},
      // Whole columns and rows keep their full extent.
      {0, "Sheet1!$B:$B", "0:0,1-1048575,1"},
      {0, "Sheet1!$2:$3", "0:1,0-2,16383"},
      {0, "Sheet1!$B:$C", "0:0,1-1048575,2"},
      {0, "[0]!WbVals", "0:1,1-4,1"},
      {0, "Sheet1!LocVals", "0:1,2-4,2"},
      {0, "LocVals", "0:1,2-4,2"},
      {0, "[0]!DynVals", "0:1,1-4,1"},
      {0, "[0]!Joined", "0:0,0-0,0,0:2,2-3,2"},
      {0, "[0]!WholeB", "0:0,1-1048575,1"},
      {0, "_xlchart.v1.0", "0:1,0-4,0"},
      {kChartNoHostSheet, "_xlchart.v1.0", "0:1,0-4,0"},
      {kChartNoHostSheet, "Sheet1!$B$2", "0:1,1-1,1"},
      {1, "HostOnly", "0:0,3-0,3"},
  };
  for (const Case& c : cases) {
    EXPECT_EQ(Resolve(wb, c.host, c.formula), c.expected) << c.formula;
  }
}

TEST(ChartSource, RejectsWhatIsNotAResolvableReference) {
  const Workbook wb = Book();
  struct Case {
    std::size_t host;
    const char* formula;
  };
  const std::vector<Case> cases = {
      {0, "#REF!"},
      {0, "Sheet1!#REF!"},
      {0, "[1]Sheet1!$A$1"},
      {0, "[1]Sheet1!$A$1:$A$4"},
      {0, "Missing!$A$1"},
      {0, "Sheet1!$A$1:"},
      {0, ""},
      {0, "{1,2}"},
      {0, "\"text\""},
      {0, "[0]!Const"},
      {0, "[0]!Undefined"},
      {kChartNoHostSheet, "$A$1"},
      // A sheet-local name is invisible from a chartsheet.
      {kChartNoHostSheet, "LocVals"},
      {0, "HostOnly"},
      {7, "Sheet1!$A$1"},
  };
  for (const Case& c : cases) {
    auto got = resolve_chart_source(wb, c.host, c.formula, eval::default_registry());
    ASSERT_FALSE(static_cast<bool>(got)) << c.formula;
    EXPECT_EQ(got.error().code, FormulonErrorCode::kIoChartInvalidReference) << c.formula;
  }
}

TEST(ChartSource, SpillReferenceCoversTheSpilledRectangle) {
  Workbook wb = Book();
  ASSERT_TRUE(wb.sheet(1).commit_spill(
      5U, 0U, 3U, 2U,
      {Value::number(1), Value::number(2), Value::number(3), Value::number(4), Value::number(5), Value::number(6)}));
  EXPECT_EQ(Resolve(wb, 0, "'My Data'!$A$6#"), "1:5,0-7,1");
  EXPECT_EQ(Resolve(wb, 1, "$A$6#"), "1:5,0-7,1");
  // No spill at the anchor.
  EXPECT_EQ(Resolve(wb, 0, "Sheet1!$H$1#"), "error");
}

}  // namespace
}  // namespace formulon
