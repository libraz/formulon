// What a save writes for `<cellXfs>`, against Excel 365 (Mac, ja-JP UI)
// re-saving the same files: which entries survive, in which order, and the
// index every cell, customFormat row and column style names. Every
// expectation is read from Excel's own `_resaved` file next to the authored
// input.

#include <cstdint>
#include <memory>
#include <string>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/styles_writer.h"
#include "io/xlsb/writer.h"
#include "pugixml.hpp"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

using Bytes = std::vector<std::uint8_t>;
using Lines = std::vector<std::string>;

/// A case's authored file `<dir>/<id>.xlsx` and Excel's `<dir>/<id>_resaved.xlsx`.
struct Case {
  const char* dir;
  const char* id;
};

std::string CaseName(const ::testing::TestParamInfo<Case>& info) {
  std::string out;
  for (const char ch : std::string(info.param.id)) {
    out.push_back((ch >= 'A' && ch <= 'Z') || (ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') ? ch : '_');
  }
  return out;
}

/// Cases where Formulon matches Excel.
const std::vector<Case>& MatchingCases() {
  static const std::vector<Case> cases = {
      {"cellxfs_compaction", "cx1_D1_dup_both_used"},
      {"cellxfs_compaction", "cx1_D3_order"},
      {"cellxfs_compaction", "cx1_D4_row_no_custom"},
      {"cellxfs_compaction", "cx1_D5_second_sheet"},
      {"cellxfs_compaction", "cx1_D6_hidden_sheet"},
      {"cellxfs_compaction", "cx1_D7_dup_of_xf0"},
      {"cellxfs_compaction", "cx1_D8_dup_later_used_only"},
      {"cellxfs_compaction", "cx1_D9_col_and_row"},
      {"cellxfs_compaction", "cx1_D11_dup_position"},
      {"cellxfs_compaction", "cx1_D12_lone_lacks_apply"},
      {"cellxfs_compaction", "cx1_D13_dup_position_late_first_use"},
      {"cell_style_save", "cs1_C0_none"},
      {"cell_style_save", "cs1_C1_all_blank"},
      {"cell_style_save", "cs1_C2_all_valued"},
      {"cell_style_save", "cs1_C3_partial"},
      {"cell_style_save", "cs1_C4_row_only"},
      {"cell_style_save", "cs1_C5_col_only"},
  };
  return cases;
}

std::string FixturePath(const Case& c, const std::string& suffix) {
  return std::string(FORMULON_FIXTURES_DIR) + "/" + c.dir + "/" + c.id + suffix;
}

std::unique_ptr<Workbook> LoadXlsxBytes(const Bytes& bytes, const std::string& what) {
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "read_ooxml " << what << ": " << read.error().message;
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

std::unique_ptr<Workbook> LoadAuthored(const Case& c) {
  return LoadXlsxBytes(test::read_file_bytes(FixturePath(c, ".xlsx")), c.id);
}

bool SaveXlsx(const Workbook& wb, Bytes* out) {
  auto saved = write_ooxml(wb);
  if (!saved) {
    ADD_FAILURE() << "write_ooxml: " << saved.error().message;
    return false;
  }
  *out = std::move(saved.value());
  return true;
}

/// Part `name` of the package as text; empty when absent.
std::string PartText(const Bytes& pkg, const std::string& name) {
  std::string text;
  (void)test::extract_part(test::span_of(pkg), name, &text);
  return text;
}

pugi::xml_node LoadRoot(pugi::xml_document& doc, const std::string& xml, const char* root) {
  if (!doc.load_buffer(xml.data(), xml.size())) {
    ADD_FAILURE() << root << " does not parse";
  }
  return doc.child(root);
}

/// Each `<cellXfs>` entry as "numFmtId=N style=<name of its parent cellStyle>".
Lines CellXfLines(const Bytes& pkg) {
  pugi::xml_document doc;
  const pugi::xml_node root = LoadRoot(doc, PartText(pkg, "xl/styles.xml"), "styleSheet");
  Lines out;
  for (const pugi::xml_node xf : root.child("cellXfs").children("xf")) {
    std::string style = "(none)";
    for (const pugi::xml_node cs : root.child("cellStyles").children("cellStyle")) {
      if (cs.attribute("xfId").as_uint() == xf.attribute("xfId").as_uint()) {
        style = cs.attribute("name").value();
        break;
      }
    }
    out.push_back("numFmtId=" + std::string(xf.attribute("numFmtId").value()) + " style=" + style);
  }
  return out;
}

/// Every `s` / `style` the worksheets write, as "<part> cell A1 s=N", "<part> row 5 s=N customFormat=1",
/// "<part> col 3-4 style=N".
Lines SheetRefLines(const Bytes& pkg) {
  Lines out;
  for (const char* part : {"xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"}) {
    const std::string xml = PartText(pkg, part);
    if (xml.empty()) {
      continue;
    }
    pugi::xml_document doc;
    const pugi::xml_node ws = LoadRoot(doc, xml, "worksheet");
    const std::string p = std::string(part).substr(14);
    for (const pugi::xml_node col : ws.child("cols").children("col")) {
      out.push_back(p + " col " + col.attribute("min").value() + "-" + col.attribute("max").value() +
                    " style=" + col.attribute("style").value());
    }
    for (const pugi::xml_node row : ws.child("sheetData").children("row")) {
      if (!row.attribute("s").empty() || row.attribute("customFormat").as_bool()) {
        out.push_back(p + " row " + row.attribute("r").value() + " s=" + row.attribute("s").value() +
                      " customFormat=" + row.attribute("customFormat").value());
      }
      for (const pugi::xml_node c : row.children("c")) {
        out.push_back(p + " cell " + c.attribute("r").value() + " s=" + c.attribute("s").value());
      }
    }
  }
  return out;
}

class CellXfsCompactionGolden : public ::testing::TestWithParam<Case> {
 protected:
  std::string Msg() const { return std::string("case ") + GetParam().id + " (left: Excel, right: Formulon)"; }
};

TEST_P(CellXfsCompactionGolden, CellXfsMatchExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsx(*wb, &ours));
  const Bytes excel = test::read_file_bytes(FixturePath(GetParam(), "_resaved.xlsx"));
  EXPECT_EQ(CellXfLines(excel), CellXfLines(ours)) << Msg();
}

TEST_P(CellXfsCompactionGolden, SheetStyleIndicesMatchExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsx(*wb, &ours));
  const Bytes excel = test::read_file_bytes(FixturePath(GetParam(), "_resaved.xlsx"));
  EXPECT_EQ(SheetRefLines(excel), SheetRefLines(ours)) << Msg();
}

TEST_P(CellXfsCompactionGolden, XlsxSaveIsIdempotent) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  Bytes first;
  ASSERT_TRUE(SaveXlsx(*wb, &first));
  const auto reloaded = LoadXlsxBytes(first, "first save");
  ASSERT_TRUE(reloaded);
  Bytes second;
  ASSERT_TRUE(SaveXlsx(*reloaded, &second));
  EXPECT_EQ(PartText(first, "xl/styles.xml"), PartText(second, "xl/styles.xml")) << Msg();
  EXPECT_EQ(SheetRefLines(first), SheetRefLines(second)) << Msg();
}

TEST_P(CellXfsCompactionGolden, SaveLeavesTheModelUnchanged) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  const std::string before = write_styles(wb->styles());
  Bytes before_sheets;
  ASSERT_TRUE(SaveXlsx(*wb, &before_sheets));
  Bytes saved;
  ASSERT_TRUE(SaveXlsx(*wb, &saved));
  EXPECT_EQ(write_styles(wb->styles()), before) << Msg();
  EXPECT_EQ(SheetRefLines(before_sheets), SheetRefLines(saved)) << Msg();
  auto xlsb = xlsb::write_xlsb(*wb);
  ASSERT_TRUE(xlsb) << xlsb.error().message;
  EXPECT_EQ(write_styles(wb->styles()), before) << Msg() << " after xlsb save";
}

INSTANTIATE_TEST_SUITE_P(Cases, CellXfsCompactionGolden, ::testing::ValuesIn(MatchingCases()), CaseName);

// Excel recomputes the apply flags on save, so two cell xfs that differ only in
// applyNumberFormat collapse there. Formulon writes the flags as stored and
// keeps both (tests/divergence.yaml).
TEST(CellXfsCompactionGoldenApplyFlags, CellXfsDifferingOnlyInApplyFlagsStayApart) {
  for (const char* id : {"cx1_D2_dup_apply_differs", "cx1_D10_dup_first_lacks_apply"}) {
    const Case c{"cellxfs_compaction", id};
    const auto wb = LoadAuthored(c);
    ASSERT_TRUE(wb);
    Bytes ours;
    ASSERT_TRUE(SaveXlsx(*wb, &ours));
    const Bytes excel = test::read_file_bytes(FixturePath(c, "_resaved.xlsx"));
    EXPECT_EQ(CellXfLines(excel), (Lines{"numFmtId=0 style=標準", "numFmtId=3 style=標準"})) << id;
    EXPECT_EQ(CellXfLines(ours), (Lines{"numFmtId=0 style=標準", "numFmtId=3 style=標準", "numFmtId=3 style=標準"}))
        << id;
    Lines row20;
    for (const std::string& line : SheetRefLines(ours)) {
      if (line.find("20 s=") != std::string::npos) {
        row20.push_back(line);
      }
    }
    EXPECT_EQ(row20, (Lines{"sheet1.xml cell A20 s=1", "sheet1.xml cell B20 s=2"})) << id;
  }
}

}  // namespace
}  // namespace formulon::io
