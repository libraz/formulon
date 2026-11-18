//
// The spill operator `X#` over an anchor that does not spill. A formula in X
// is a one-cell region whether or not it spills, so `X#` reads its value and
// measures one cell; a constant in X anchors nothing and reads `#REF!`. The
// saved form is checked in both packages: the bare operator stores as a
// dynamic-array `_xlfn.ANCHORARRAY` block over the result extent, and the
// operator inside ROWS / ISREF as a plain formula.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct AnchorCase {
  // Entered in B1; empty for the typed `#N/A` constant.
  const char* anchor;
  // Entered in D1.
  const char* formula;
  Value expected;
  // Rows of the block D1 stores; 0 when D1 is a plain formula.
  std::uint32_t block_rows;
  // The saved `<c r="D1" ...>...</c>` element.
  const char* saved_d1;
};

std::vector<AnchorCase> Cases() {
  return {
      {"=1+1", "=B1#", Value::number(2.0), 1U,
       R"(<c r="D1" cm="1"><f t="array" ref="D1">_xlfn.ANCHORARRAY(B1)</f><v>2</v></c>)"},
      {"=\"x\"", "=B1#", Value::text("x"), 1U,
       R"(<c r="D1" t="str" cm="1"><f t="array" ref="D1">_xlfn.ANCHORARRAY(B1)</f><v xml:space="preserve">x</v></c>)"},
      {"=1/0", "=B1#", Value::error(ErrorCode::Div0), 1U,
       R"(<c r="D1" t="e" cm="1"><f t="array" ref="D1">_xlfn.ANCHORARRAY(B1)</f><v>#DIV/0!</v></c>)"},
      {"", "=B1#", Value::error(ErrorCode::Ref), 1U,
       R"(<c r="D1" t="e" cm="1"><f t="array" ref="D1">_xlfn.ANCHORARRAY(B1)</f><v>#REF!</v></c>)"},
      {"=NA()", "=ROWS(B1#)", Value::number(1.0), 0U, R"(<c r="D1"><f>ROWS(_xlfn.ANCHORARRAY(B1))</f><v>1</v></c>)"},
      {"=NA()", "=ISREF(B1#)", Value::boolean(true), 0U,
       R"(<c r="D1" t="b"><f>ISREF(_xlfn.ANCHORARRAY(B1))</f><v>1</v></c>)"},
      {"=SEQUENCE(2)/0", "=B1#", Value::error(ErrorCode::Div0), 2U,
       R"(<c r="D1" t="e" cm="1"><f t="array" ref="D1:D2">_xlfn.ANCHORARRAY(B1)</f><v>#DIV/0!</v></c>)"},
      {"=A5", "=B1#", Value::number(0.0), 1U,
       R"(<c r="D1" cm="1"><f t="array" ref="D1">_xlfn.ANCHORARRAY(B1)</f><v>0</v></c>)"},
  };
}

// One sheet per case, recalculated.
Workbook BuildCases(const std::vector<AnchorCase>& cases) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  for (std::size_t i = 1; i < cases.size(); ++i) {
    wb.add_sheet("c" + std::to_string(i + 1U));
  }
  for (std::size_t i = 0; i < cases.size(); ++i) {
    if (std::string_view(cases[i].anchor).empty()) {
      EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(i, 0U, 1U, Value::error(ErrorCode::NA))));
    } else {
      EXPECT_TRUE(
          static_cast<bool>(wb.set_cell_formula(i, 0U, 1U, cases[i].anchor, Workbook::FormulaSource::kEntered)));
    }
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(i, 0U, 3U, cases[i].formula, Workbook::FormulaSource::kEntered)));
  }
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return wb;
}

void ExpectValues(const Workbook& wb, const std::vector<AnchorCase>& cases) {
  for (std::size_t i = 0; i < cases.size(); ++i) {
    const Sheet& s = wb.sheet(i);
    EXPECT_EQ(s.resolve_cell_value(0U, 3U), cases[i].expected) << cases[i].anchor << " " << cases[i].formula;
    for (std::uint32_t row = 1; row < cases[i].block_rows; ++row) {
      EXPECT_EQ(s.resolve_cell_value(row, 3U), cases[i].expected) << cases[i].anchor << " row " << row;
    }
  }
}

TEST(SpillAnchor, FormulaAnchorIsAOneCellRegion) {
  const std::vector<AnchorCase> cases = Cases();
  const Workbook wb = BuildCases(cases);
  ExpectValues(wb, cases);
}

TEST(SpillAnchor, SavedFormMatchesTheResultExtent) {
  const std::vector<AnchorCase> cases = Cases();
  Workbook wb = BuildCases(cases);

  auto saved = wb.save();
  ASSERT_TRUE(static_cast<bool>(saved)) << saved.error().message;
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(saved.value()))));
  for (std::size_t i = 0; i < cases.size(); ++i) {
    auto part = zip.read_entry("xl/worksheets/sheet" + std::to_string(i + 1U) + ".xml");
    ASSERT_TRUE(static_cast<bool>(part));
    const std::string xml(part.value().begin(), part.value().end());
    EXPECT_NE(xml.find(cases[i].saved_d1), std::string::npos) << xml;
  }

  auto xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto xlsb_read = io::xlsb::read_xlsb(SpanOf(xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(xlsb_read)) << xlsb_read.error().message;
  auto xlsx_read = io::read_ooxml(SpanOf(saved.value()));
  ASSERT_TRUE(static_cast<bool>(xlsx_read)) << xlsx_read.error().message;
  for (Workbook* round_tripped : {&xlsx_read.value().workbook, &xlsb_read.value().workbook}) {
    for (std::size_t i = 0; i < cases.size(); ++i) {
      const Cell* d1 = round_tripped->sheet(i).cell_at(0U, 3U);
      ASSERT_NE(d1, nullptr) << i;
      EXPECT_EQ(d1->dynamic_array, cases[i].block_rows != 0U) << cases[i].formula;
    }
    ASSERT_TRUE(static_cast<bool>(round_tripped->recalc(eval::default_registry())));
    ExpectValues(*round_tripped, cases);
  }
}

}  // namespace
}  // namespace formulon
