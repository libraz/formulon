// Excel-saved LET / LAMBDA fixtures whose parameter and name tokens are
// A1-cell-shaped (`B1`, `A1`, `B2`): the stored text, its display form, the
// recalculated values, and the xlsx / xlsb save paths.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "parser/formula_prefix.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"
#include "xlsb_roundtrip_symmetry_test_helpers.h"

namespace formulon {
namespace {

using xlsb_roundtrip_test_support::FixturePath;
using xlsb_roundtrip_test_support::ReadWireNames;

constexpr std::uint32_t kColE = 4U;

/// One column-E formula cell of a fixture: the canonical stored text, the
/// `<f>` text Excel saved, Excel's FORMULATEXT, and Excel's value.
struct ExpectedRow {
  enum Kind { kNumber, kError };
  std::uint32_t row;  // 1-based
  const char* stored;
  const char* excel_f;
  const char* formula_text;
  Kind kind;
  double number;
  int error_type;  // ERROR.TYPE code when kind == kError
};

const std::vector<ExpectedRow>& LetA1NamesRows() {
  static const std::vector<ExpectedRow> kRows = {
      {3, R"FX(=D3)FX", R"FX(=D3)FX", R"FX(=D3)FX", ExpectedRow::kNumber, 7.0, -1},
      {4, R"FX(=$A$1)FX", R"FX(=$A$1)FX", R"FX(=$A$1)FX", ExpectedRow::kNumber, 100.0, -1},
      {5, R"FX(=D5)FX", R"FX(=D5)FX", R"FX(=D5)FX", ExpectedRow::kNumber, 7.0, -1},
      {6, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+$A$1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+$A$1))FX",
       R"FX(=LET(B1,5,B1+$A$1))FX", ExpectedRow::kNumber, 105.0, -1},
      {7, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+D7))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+D7))FX", R"FX(=LET(B1,5,B1+D7))FX",
       ExpectedRow::kNumber, 12.0, -1},
      {8, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+D8))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+D8))FX", R"FX(=LET(B1,5,B1+D8))FX",
       ExpectedRow::kNumber, 12.0, -1},
      {9, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+$A$1:$A$2))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+$A$1:$A$2))FX",
       R"FX(=LET(B1,5,B1+$A$1:$A$2))FX", ExpectedRow::kError, 0.0, 9},
      {10, R"FX(=LET(_xlpm.B1,5,SUM(_xlpm.B1,$A$1:$A$2)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,SUM(_xlpm.B1,$A$1:$A$2)))FX",
       R"FX(=LET(B1,5,SUM(B1,$A$1:$A$2)))FX", ExpectedRow::kNumber, 125.0, -1},
      {11, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+Sheet1!$A$1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+Sheet1!$A$1))FX",
       R"FX(=LET(B1,5,B1+Sheet1!$A$1))FX", ExpectedRow::kNumber, 105.0, -1},
      {12, R"FX(=LET(_xlpm.B1,5,SUM(_xlpm.B1,$A:$A)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,SUM(_xlpm.B1,$A:$A)))FX",
       R"FX(=LET(B1,5,SUM(B1,$A:$A)))FX", ExpectedRow::kNumber, 125.0, -1},
      {13, R"FX(=LET(_xlpm.B1,5,COUNT($1:$1)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,COUNT($1:$1)))FX",
       R"FX(=LET(B1,5,COUNT($1:$1)))FX", ExpectedRow::kNumber, 1.0, -1},
      {14, R"FX(=_xlnm.B2+LET(_xlpm.B1,5,_xlpm.B1))FX", R"FX(=_xlnm.B2+_xlfn.LET(_xlpm.B1,5,_xlpm.B1))FX",
       R"FX(='B2'+LET(B1,5,B1))FX", ExpectedRow::kError, 0.0, 5},
      {15, R"FX(=LET(_xlpm.B1,5,_xlpm.B1)+_xlnm.B2)FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1)+_xlnm.B2)FX",
       R"FX(=LET(B1,5,B1)+'B2')FX", ExpectedRow::kError, 0.0, 5},
      {16, R"FX(=LET(x,1,x)+LET(_xlpm.B1,5,_xlpm.B1))FX",
       R"FX(=_xlfn.LET(_xlpm.x,1,_xlpm.x)+_xlfn.LET(_xlpm.B1,5,_xlpm.B1))FX", R"FX(=LET(x,1,x)+LET(B1,5,B1))FX",
       ExpectedRow::kNumber, 6.0, -1},
      {17, R"FX(=LET(x,5,x+B2))FX", R"FX(=_xlfn.LET(_xlpm.x,5,_xlpm.x+B2))FX", R"FX(=LET(x,5,x+B2))FX",
       ExpectedRow::kNumber, 5.0, -1},
      {18, R"FX(=LET(_xlpm.B1,100,LOG10(_xlpm.B1)))FX", R"FX(=_xlfn.LET(_xlpm.B1,100,LOG10(_xlpm.B1)))FX",
       R"FX(=LET(B1,100,LOG10(B1)))FX", ExpectedRow::kNumber, 2.0, -1},
      {19, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+XFE1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+XFE1))FX",
       R"FX(=LET(B1,5,B1+XFE1))FX", ExpectedRow::kError, 0.0, 5},
      {21, R"FX(=LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4))FX", R"FX(=_xlfn.LAMBDA(_xlpm.B1,_xlpm.B1+$A$1)(4))FX",
       R"FX(=LAMBDA(B1,B1+$A$1)(4))FX", ExpectedRow::kNumber, 104.0, -1},
      {25, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+A:(A)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+A:(A)))FX",
       R"FX(=LET(B1,5,B1+A:(A)))FX", ExpectedRow::kError, 0.0, 5},
      {27, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+E28))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+E28))FX",
       R"FX(=LET(B1,5,B1+E28))FX", ExpectedRow::kNumber, 8.0, -1},
      {28, R"FX(=LET(_xlpm.B1,5,ROWS($A$1:$A$3)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,ROWS($A$1:$A$3)))FX",
       R"FX(=LET(B1,5,ROWS($A$1:$A$3)))FX", ExpectedRow::kNumber, 3.0, -1},
      {29, R"FX(=LET(_xlpm.B1,5,ROW($A$2)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,ROW($A$2)))FX",
       R"FX(=LET(B1,5,ROW($A$2)))FX", ExpectedRow::kNumber, 2.0, -1},
      {30, R"FX(=LET(_xlpm.B1,5,INDIRECT("A1")+_xlpm.B1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,INDIRECT("A1")+_xlpm.B1))FX",
       R"FX(=LET(B1,5,INDIRECT("A1")+B1))FX", ExpectedRow::kNumber, 105.0, -1},
      {31, R"FX(=LET(_xlpm.B1,5,INDIRECT("R1C1",FALSE)+_xlpm.B1))FX",
       R"FX(=_xlfn.LET(_xlpm.B1,5,INDIRECT("R1C1",FALSE)+_xlpm.B1))FX", R"FX(=LET(B1,5,INDIRECT("R1C1",FALSE)+B1))FX",
       ExpectedRow::kNumber, 105.0, -1},
      {32, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1))FX",
       R"FX(=LET(B1,5,B1+'A1'))FX", ExpectedRow::kError, 0.0, 5},
      {33, R"FX(=LET(_xlpm.A1,5,_xlpm.A1+$A$1))FX", R"FX(=_xlfn.LET(_xlpm.A1,5,_xlpm.A1+$A$1))FX",
       R"FX(=LET(A1,5,A1+$A$1))FX", ExpectedRow::kNumber, 105.0, -1},
      {34, R"FX(=LET(_xlpm.B1,2,_xlpm.B1)*$A$1)FX", R"FX(=_xlfn.LET(_xlpm.B1,2,_xlpm.B1)*$A$1)FX",
       R"FX(=LET(B1,2,B1)*$A$1)FX", ExpectedRow::kNumber, 200.0, -1},
  };
  return kRows;
}

const std::vector<ExpectedRow>& LetA1Names2Rows() {
  static const std::vector<ExpectedRow> kRows = {
      {3, R"FX(=LET(x,5,x+$A$1))FX", R"FX(=_xlfn.LET(_xlpm.x,5,_xlpm.x+$A$1))FX", R"FX(=LET(x,5,x+$A$1))FX",
       ExpectedRow::kNumber, 105.0, -1},
      {4, R"FX(=$A$1+1)FX", R"FX(=$A$1+1)FX", R"FX(=$A$1+1)FX", ExpectedRow::kNumber, 101.0, -1},
      {7, R"FX(=LET(R1C1,3,R1C1))FX", R"FX(=_xlfn.LET(_xlpm.R1C1,3,_xlpm.R1C1))FX", R"FX(=LET(R1C1,3,R1C1))FX",
       ExpectedRow::kNumber, 3.0, -1},
      {8, R"FX(=SUM(R1:R2))FX", R"FX(=SUM(R1:R2))FX", R"FX(=SUM(R1:R2))FX", ExpectedRow::kNumber, 0.0, -1},
      {9, R"FX(=$A$1:$A$2)FX", R"FX(=$A$1:$A$2)FX", R"FX(=$A$1:$A$2)FX", ExpectedRow::kError, 0.0, 9},
      {10, R"FX(=SUM(11:12))FX", R"FX(=SUM(11:12))FX", R"FX(=SUM(11:12))FX", ExpectedRow::kNumber, 1064972.0, -1},
      {11, R"FX(=ROWS(D:D))FX", R"FX(=ROWS(D:D))FX", R"FX(=ROWS(D:D))FX", ExpectedRow::kNumber, 1048576.0, -1},
      {12, R"FX(=COLUMNS(13:13))FX", R"FX(=COLUMNS(13:13))FX", R"FX(=COLUMNS(13:13))FX", ExpectedRow::kNumber, 16384.0,
       -1},
      {13, R"FX(=XEY13)FX", R"FX(=XEY13)FX", R"FX(=XEY13)FX", ExpectedRow::kNumber, 0.0, -1},
      {15, R"FX(=Sheet1!$A$1)FX", R"FX(=Sheet1!$A$1)FX", R"FX(=Sheet1!$A$1)FX", ExpectedRow::kNumber, 100.0, -1},
      {16, R"FX(=$A$1)FX", R"FX(=$A$1)FX", R"FX(=$A$1)FX", ExpectedRow::kNumber, 100.0, -1},
      {17, R"FX(=D17)FX", R"FX(=D17)FX", R"FX(=D17)FX", ExpectedRow::kNumber, 7.0, -1},
      {18, R"FX(=LET(_xlpm.B1,5,Sheet1!_xlnm.B1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,Sheet1!_xlnm.B1))FX",
       R"FX(=LET(B1,5,Sheet1!'B1'))FX", ExpectedRow::kError, 0.0, 5},
      {19, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2))FX",
       R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.A1:_xlnm.B2))FX", R"FX(=LET(B1,5,B1+'A1':'B2'))FX",
       ExpectedRow::kError, 0.0, 5},
      {20, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+A))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+A))FX", R"FX(=LET(B1,5,B1+A))FX",
       ExpectedRow::kError, 0.0, 5},
      {21, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+AB))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+AB))FX",
       R"FX(=LET(B1,5,B1+AB))FX", ExpectedRow::kError, 0.0, 5},
      {22, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+XFD1048577))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+XFD1048577))FX",
       R"FX(=LET(B1,5,B1+XFD1048577))FX", ExpectedRow::kError, 0.0, 5},
      {23, R"FX(=LET(R1C,2,R1C))FX", R"FX(=_xlfn.LET(_xlpm.R1C,2,_xlpm.R1C))FX", R"FX(=LET(R1C,2,R1C))FX",
       ExpectedRow::kNumber, 2.0, -1},
      {26, R"FX(=LET(_xlpm.B1,5,SUM(A:(B))))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,SUM(A:(B))))FX",
       R"FX(=LET(B1,5,SUM(A:(B))))FX", ExpectedRow::kError, 0.0, 5},
      {28, R"FX(=LET(_xlpm.B1,5,SUM($1:$2)))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,SUM($1:$2)))FX",
       R"FX(=LET(B1,5,SUM($1:$2)))FX", ExpectedRow::kNumber, 1120.0, -1},
      {29, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+$A$1:$B$1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+$A$1:$B$1))FX",
       R"FX(=LET(B1,5,B1+$A$1:$B$1))FX", ExpectedRow::kError, 0.0, 9},
      {31, R"FX(=_xlnm.B2)FX", R"FX(=_xlnm.B2)FX", R"FX(='B2')FX", ExpectedRow::kError, 0.0, 5},
      {32, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+_xlnm.B2))FX",
       R"FX(=LET(B1,5,B1+'B2'))FX", ExpectedRow::kError, 0.0, 5},
      {33, R"FX(=LET(_xlpm.B1,5,_xlpm.B1)+B34)FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1)+B34)FX",
       R"FX(=LET(B1,5,B1)+B34)FX", ExpectedRow::kNumber, 5.0, -1},
      {34, R"FX(=LAMBDA(R1C,R1C)(3))FX", R"FX(=_xlfn.LAMBDA(_xlpm.R1C,_xlpm.R1C)(3))FX", R"FX(=LAMBDA(R1C,R1C)(3))FX",
       ExpectedRow::kNumber, 3.0, -1},
      {36, R"FX(=LET(_xlpm.B1,5,_xlpm.B1+B1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1+B1))FX",
       R"FX(=LET(B1,5,B1+B1))FX", ExpectedRow::kNumber, 1005.0, -1},
      {37, R"FX(=LET(_xlpm.B1,5,_xlpm.B1*0+B1))FX", R"FX(=_xlfn.LET(_xlpm.B1,5,_xlpm.B1*0+B1))FX",
       R"FX(=LET(B1,5,B1*0+B1))FX", ExpectedRow::kNumber, 1000.0, -1},
  };
  return kRows;
}

struct FixtureCase {
  const char* base;
  const std::vector<ExpectedRow>& (*rows)();
};

const FixtureCase kCases[] = {
    {"let_a1_names", &LetA1NamesRows},
    {"let_a1_names2", &LetA1Names2Rows},
};

ErrorCode ErrorOfType(int error_type) {
  switch (error_type) {
    case 3:
      return ErrorCode::Value;
    case 5:
      return ErrorCode::Name;
    default:
      return ErrorCode::Spill;  // 9
  }
}

::testing::AssertionResult LoadFixture(const std::string& file, Workbook* out) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(file.c_str()));
  if (bytes.empty()) {
    return ::testing::AssertionFailure() << "fixture bytes empty: " << file;
  }
  if (file.size() > 5U && file.compare(file.size() - 5U, 5U, ".xlsb") == 0) {
    auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
    if (!loaded) {
      return ::testing::AssertionFailure() << "read_xlsb failed: " << loaded.error().message;
    }
    *out = std::move(loaded.value().workbook);
  } else {
    auto loaded = io::read_ooxml(test::span_of(bytes));
    if (!loaded) {
      return ::testing::AssertionFailure() << "read_ooxml failed: " << loaded.error().message;
    }
    *out = std::move(loaded.value().workbook);
  }
  return ::testing::AssertionSuccess();
}

std::string StoredAt(const Workbook& wb, std::uint32_t row) {
  const Cell* cell = wb.sheet(0).cell_at(row - 1U, kColE);
  return cell == nullptr ? std::string("<no cell>") : cell->formula_text;
}

void ExpectStoredText(const Workbook& wb, const std::vector<ExpectedRow>& rows) {
  for (const ExpectedRow& r : rows) {
    SCOPED_TRACE("E" + std::to_string(r.row));
    const std::string stored = StoredAt(wb, r.row);
    EXPECT_EQ(stored, r.stored);
    EXPECT_EQ(parser::display_cell_shaped_names(stored), r.formula_text);
  }
}

void ExpectValues(Workbook& wb, const std::vector<ExpectedRow>& rows) {
  auto recalc = wb.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(recalc)) << recalc.error().message;
  for (const ExpectedRow& r : rows) {
    SCOPED_TRACE("E" + std::to_string(r.row));
    const Value v = wb.sheet(0).resolve_cell_value(r.row - 1U, kColE);
    if (r.kind == ExpectedRow::kError) {
      ASSERT_TRUE(v.is_error());
      EXPECT_EQ(static_cast<int>(v.as_error()), static_cast<int>(ErrorOfType(r.error_type)));
    } else {
      ASSERT_TRUE(v.is_number());
      EXPECT_EQ(v.as_number(), r.number);
    }
  }
}

std::string UnescapeXml(std::string s) {
  static const std::pair<const char*, const char*> kEntities[] = {
      {"&quot;", "\""}, {"&apos;", "'"}, {"&lt;", "<"}, {"&gt;", ">"}, {"&amp;", "&"}};
  for (const auto& [from, to] : kEntities) {
    for (std::size_t pos = s.find(from); pos != std::string::npos; pos = s.find(from, pos + 1U)) {
      s.replace(pos, std::string_view(from).size(), to);
    }
  }
  return s;
}

/// The text of the `<f>` element of cell E<row> in a sheet part, "" if none.
std::string FormulaElement(const std::string& xml, std::uint32_t row) {
  const std::string anchor = "<c r=\"E" + std::to_string(row) + "\"";
  const std::size_t cell = xml.find(anchor);
  if (cell == std::string::npos) {
    return std::string();
  }
  const std::size_t cell_end = xml.find("</c>", cell);
  const std::size_t open = xml.find("<f", cell);
  if (open == std::string::npos || open > cell_end) {
    return std::string();
  }
  const std::size_t body = xml.find('>', open);
  const std::size_t close = xml.find("</f>", body);
  if (body == std::string::npos || close == std::string::npos) {
    return std::string();
  }
  return UnescapeXml(xml.substr(body + 1U, close - body - 1U));
}

std::vector<std::string> FixtureFiles(const FixtureCase& c) {
  return {std::string(c.base) + ".xlsx", std::string(c.base) + ".xlsb"};
}

TEST(LetA1NamesFixture, LoadKeepsCanonicalStoredTextAndDisplay) {
  for (const FixtureCase& c : kCases) {
    for (const std::string& file : FixtureFiles(c)) {
      SCOPED_TRACE(file);
      Workbook wb = Workbook::create_empty();
      ASSERT_TRUE(LoadFixture(file, &wb));
      ExpectStoredText(wb, c.rows());
    }
  }
}

TEST(LetA1NamesFixture, RecalculatedValuesMatchExcel) {
  for (const FixtureCase& c : kCases) {
    for (const std::string& file : FixtureFiles(c)) {
      SCOPED_TRACE(file);
      Workbook wb = Workbook::create_empty();
      ASSERT_TRUE(LoadFixture(file, &wb));
      ExpectValues(wb, c.rows());
    }
  }
}

TEST(LetA1NamesFixture, SavedXlsxFormulaTextMatchesExcel) {
  for (const FixtureCase& c : kCases) {
    for (const std::string& file : FixtureFiles(c)) {
      SCOPED_TRACE(file);
      Workbook wb = Workbook::create_empty();
      ASSERT_TRUE(LoadFixture(file, &wb));
      auto saved = io::write_ooxml(wb);
      ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
      std::string xml;
      ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/worksheets/sheet1.xml", &xml));
      for (const ExpectedRow& r : c.rows()) {
        SCOPED_TRACE("E" + std::to_string(r.row));
        EXPECT_EQ("=" + FormulaElement(xml, r.row), r.excel_f);
      }
    }
  }
}

TEST(LetA1NamesFixture, SavedXlsbRoundTripsTextValuesAndNames) {
  for (const FixtureCase& c : kCases) {
    for (const std::string& file : FixtureFiles(c)) {
      SCOPED_TRACE(file);
      Workbook wb = Workbook::create_empty();
      ASSERT_TRUE(LoadFixture(file, &wb));
      auto saved = io::xlsb::write_xlsb(wb);
      ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;

      std::string workbook_bin;
      ASSERT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/workbook.bin", &workbook_bin));
      const auto names = ReadWireNames(std::vector<std::uint8_t>(workbook_bin.begin(), workbook_bin.end()));
      bool has_cell_name = false;
      for (const auto& n : names) {
        EXPECT_NE(n.name, "_xlnm.B2");
        EXPECT_NE(n.name, "_xlnm.A1");
        EXPECT_EQ(n.name.find("_xlpm._xlpm."), std::string::npos) << n.name;
        has_cell_name = has_cell_name || n.name == "B2";
      }
      EXPECT_TRUE(has_cell_name);

      auto reloaded = io::xlsb::read_xlsb(test::span_of(saved.value()));
      ASSERT_TRUE(static_cast<bool>(reloaded)) << reloaded.error().message;
      Workbook after = std::move(reloaded.value().workbook);
      ExpectStoredText(after, c.rows());
      ExpectValues(after, c.rows());
    }
  }
}

}  // namespace
}  // namespace formulon
