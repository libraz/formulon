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
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
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
      {"cellxfs_compaction", "mg1_M3_in_and_out"},
      {"cellxfs_compaction", "mg1_M5_two_merge_dups"},
      {"cellxfs_compaction", "mg1_M6_merge_dup_of_plain"},
      {"cellxfs_compaction", "mg1_M7_merge_first"},
      {"cellxfs_compaction", "mg1_M8_merge_dup_of_xf0"},
      {"cellxfs_compaction", "mg1_M9_two_merge_variants_order"},
      {"cellxfs_compaction", "mg2_s_merge"},
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

// Record types read from styles.bin and the sheet parts.
constexpr std::uint16_t kRowHdr = 0;
constexpr std::uint16_t kXf = 47;
constexpr std::uint16_t kColInfo = 60;
constexpr std::uint16_t kCellRString = 62;
constexpr std::uint16_t kBeginCellXfs = 617;
constexpr std::uint16_t kEndCellXfs = 618;

template <typename Fn>
void ForEachRecord(const std::string& part, Fn&& fn) {
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(part.data()), part.size()};
  while (cursor.size > 0U) {
    auto rec = xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "record: " << rec.error().message;
      return;
    }
    fn(rec.value().type, rec.value().payload);
  }
}

/// Each BrtXF of the cell table as "numFmt=N parent=P".
Lines XlsbCellXfLines(const Bytes& pkg) {
  Lines out;
  bool in_cell_xfs = false;
  ForEachRecord(PartText(pkg, "xl/styles.bin"), [&](std::uint16_t type, ByteSpan payload) {
    if (type == kBeginCellXfs || type == kEndCellXfs) {
      in_cell_xfs = type == kBeginCellXfs;
    } else if (type == kXf && in_cell_xfs && payload.size >= 4U) {
      out.push_back("numFmt=" + std::to_string(payload.data[2] | (payload.data[3] << 8)) +
                    " parent=" + std::to_string(payload.data[0] | (payload.data[1] << 8)));
    }
  });
  return out;
}

/// Every ixfe the sheet parts write: custom-format rows, cells and column infos.
Lines XlsbSheetRefLines(const Bytes& pkg) {
  Lines out;
  for (const char* part : {"xl/worksheets/sheet1.bin", "xl/worksheets/sheet2.bin"}) {
    const std::string p = std::string(part).substr(14);
    std::uint32_t row = 0;
    ForEachRecord(PartText(pkg, part), [&](std::uint16_t type, ByteSpan payload) {
      if (type == kRowHdr && payload.size >= 12U) {
        row = xlsb::load_u32(payload.data);
        if ((payload.data[11] & 0x40U) != 0U) {
          out.push_back(p + " row " + std::to_string(row + 1U) +
                        " ixfe=" + std::to_string(xlsb::load_u32(payload.data + 4)));
        }
      } else if (((type >= 1U && type <= 11U) || type == kCellRString) && payload.size >= 8U) {
        out.push_back(p + " cell r" + std::to_string(row + 1U) + " c" + std::to_string(xlsb::load_u32(payload.data)) +
                      " ixfe=" + std::to_string(xlsb::load_u32(payload.data + 4) & 0xFFFFFFU));
      } else if (type == kColInfo && payload.size >= 16U) {
        out.push_back(p + " col " + std::to_string(xlsb::load_u32(payload.data)) + "-" +
                      std::to_string(xlsb::load_u32(payload.data + 4)) +
                      " ixfe=" + std::to_string(xlsb::load_u32(payload.data + 12)));
      }
    });
  }
  return out;
}

bool SaveXlsb(const Workbook& wb, Bytes* out) {
  auto saved = xlsb::write_xlsb(wb);
  if (!saved) {
    ADD_FAILURE() << "write_xlsb: " << saved.error().message;
    return false;
  }
  *out = std::move(saved.value());
  return true;
}

std::unique_ptr<Workbook> LoadXlsbBytes(const Bytes& bytes, const std::string& what) {
  auto read = xlsb::read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "read_xlsb " << what << ": " << read.error().message;
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_excel_profile(mac_365_ja_jp_profile());
  return wb;
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

// An .xlsx book saved as .xlsb writes styles.bin from the model, so its cell
// formats compact as in the .xlsx save; Excel's .xlsb of the same book is the
// expectation.
TEST_P(CellXfsCompactionGolden, XlsbFromXlsxMatchesExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  const Bytes excel = test::read_file_bytes(FixturePath(GetParam(), "_resaved.xlsb"));
  EXPECT_EQ(XlsbCellXfLines(excel), XlsbCellXfLines(ours)) << Msg();
  EXPECT_EQ(XlsbSheetRefLines(excel), XlsbSheetRefLines(ours)) << Msg();
}

// Excel's own .xlsb saved unedited keeps its styles.bin and every ixfe as loaded.
TEST_P(CellXfsCompactionGolden, UneditedXlsbKeepsItsStylesPart) {
  const Bytes excel = test::read_file_bytes(FixturePath(GetParam(), "_resaved.xlsb"));
  const auto wb = LoadXlsbBytes(excel, GetParam().id);
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  EXPECT_EQ(PartText(excel, "xl/styles.bin"), PartText(ours, "xl/styles.bin")) << Msg();
  EXPECT_EQ(XlsbSheetRefLines(excel), XlsbSheetRefLines(ours)) << Msg();
}

INSTANTIATE_TEST_SUITE_P(Cases, CellXfsCompactionGolden, ::testing::ValuesIn(MatchingCases()), CaseName);

// A cell xf equal to xf 0 counts as xf 0 in an .xlsx save: a blank cell or a customFormat row carrying only it is not
// written and leaves the used range, and a column keeps its width without the style. B4, B5 (a bare and an `s="0"`
// blank) and B7 (a blank on a distinct xf, kept) are the controls.
const std::vector<Case>& XfCollapseCases() {
  static const std::vector<Case> cases = {
      {"xfcollapse", "xc_B1_blank_cell_dup_xf0"},      {"xfcollapse", "xc_B2_row_custom_dup_xf0"},
      {"xfcollapse", "xc_B3_col_style_dup_xf0"},       {"xfcollapse", "xc_B4_bare_blank_cell"},
      {"xfcollapse", "xc_B5_blank_cell_s0"},           {"xfcollapse", "xc_B6_col_style_dup_xf0_nowidth"},
      {"xfcollapse", "xc_B7_blank_cell_styled_kept"},  {"xfcollapse", "xc_H1_row_run_blank_dup_xf0"},
      {"xfcollapse", "xc_H1b_row_run_values_dup_xf0"}, {"xfcollapse", "xc_H1c_row_run_values_s0"},
      {"xfcollapse", "xc_H2_col_run_blank_dup_xf0"},   {"xfcollapse", "xc_H2b_col_run_values_dup_xf0"},
  };
  return cases;
}

std::string DimensionRef(const Bytes& pkg) {
  pugi::xml_document doc;
  const pugi::xml_node ws = LoadRoot(doc, PartText(pkg, "xl/worksheets/sheet1.xml"), "worksheet");
  return ws.child("dimension").attribute("ref").value();
}

/// Each `<col>` as "min-max width=W" (the style is in SheetRefLines).
Lines ColWidthLines(const Bytes& pkg) {
  pugi::xml_document doc;
  const pugi::xml_node ws = LoadRoot(doc, PartText(pkg, "xl/worksheets/sheet1.xml"), "worksheet");
  Lines out;
  for (const pugi::xml_node col : ws.child("cols").children("col")) {
    out.push_back(std::string(col.attribute("min").value()) + "-" + col.attribute("max").value() +
                  " width=" + col.attribute("width").value());
  }
  return out;
}

class XfCollapseGolden : public ::testing::TestWithParam<Case> {
 protected:
  std::string Msg() const { return std::string("case ") + GetParam().id + " (left: Excel, right: Formulon)"; }
};

TEST_P(XfCollapseGolden, XlsxMatchesExcel) {
  const auto wb = LoadAuthored(GetParam());
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsx(*wb, &ours));
  const Bytes excel = test::read_file_bytes(FixturePath(GetParam(), "_resaved.xlsx"));
  EXPECT_EQ(CellXfLines(excel), CellXfLines(ours)) << Msg();
  EXPECT_EQ(SheetRefLines(excel), SheetRefLines(ours)) << Msg();
  EXPECT_EQ(DimensionRef(excel), DimensionRef(ours)) << Msg();
  EXPECT_EQ(ColWidthLines(excel), ColWidthLines(ours)) << Msg();
}

TEST_P(XfCollapseGolden, XlsxSaveIsIdempotent) {
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
  EXPECT_EQ(DimensionRef(first), DimensionRef(second)) << Msg();
}

INSTANTIATE_TEST_SUITE_P(Cases, XfCollapseGolden, ::testing::ValuesIn(XfCollapseCases()), CaseName);

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

// Excel's .xlsb of a pivot table with merged item labels, renumbered by Excel
// itself on resave, still passes through unedited: its styles.bin is kept.
TEST(CellXfsCompactionGoldenXlsb, UneditedExcelXlsbWithMergedEntriesKeepsItsStylesPart) {
  const Bytes excel = test::read_file_bytes(FixturePath({"cellxfs_compaction", "mg2_B14_merge_items_b"}, ".xlsb"));
  const auto wb = LoadXlsbBytes(excel, "B14");
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  EXPECT_EQ(PartText(excel, "xl/styles.bin"), PartText(ours, "xl/styles.bin"));
  EXPECT_EQ(XlsbSheetRefLines(excel), XlsbSheetRefLines(ours));
}

// A loaded .xlsb whose styles changed writes styles.bin from the model and
// compacts: a cell xf added and never used is dropped again.
TEST(CellXfsCompactionGoldenXlsb, EditedXlsbStylesCompact) {
  const Case c{"cellxfs_compaction", "cx1_D3_order"};
  const Bytes excel = test::read_file_bytes(FixturePath(c, "_resaved.xlsb"));
  const auto wb = LoadXlsbBytes(excel, c.id);
  ASSERT_TRUE(wb);
  CellXf unused;
  unused.num_fmt_id = 10U;
  unused.apply_number_format = true;
  wb->mutable_styles().cell_xfs.push_back(unused);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  EXPECT_EQ(XlsbCellXfLines(excel), XlsbCellXfLines(ours));
  EXPECT_EQ(XlsbSheetRefLines(excel), XlsbSheetRefLines(ours));
}

// A merge over xf-0 cells takes its own entry. The model keeps no blank xf-0
// cell, so the merge's empty second cell B20, which Excel writes with that
// entry, is not written; everything else matches Excel.
TEST(CellXfsCompactionGoldenMergeOnXf0, MergedXf0CellsTakeTheirOwnEntry) {
  const Case c{"cellxfs_compaction", "mg1_M2_merge_on_xf0"};
  const auto wb = LoadAuthored(c);
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsx(*wb, &ours));
  const Bytes excel = test::read_file_bytes(FixturePath(c, "_resaved.xlsx"));
  EXPECT_EQ(CellXfLines(excel), CellXfLines(ours));
  Lines expected;
  for (const std::string& line : SheetRefLines(excel)) {
    if (line != "sheet1.xml cell B20 s=2") {
      expected.push_back(line);
    }
  }
  ASSERT_EQ(expected.size() + 1U, SheetRefLines(excel).size());
  EXPECT_EQ(expected, SheetRefLines(ours));
}

// Pivot tables Excel wrote with merged item labels. Excel also collapses the
// entry that differs from xf 0 only in applyNumberFormat (index 2 in B14, 1 in
// op_merge_tabular); with that entry folded back into 0, Formulon's output is
// Excel's.
TEST(CellXfsCompactionGoldenMergedPivotLabels, MergedCellsTakeTheirOwnEntriesAfterTheOthers) {
  const std::pair<const char*, unsigned> cases[] = {{"mg2_B14_merge_items", 2U}, {"mg2_op_merge_tabular", 1U}};
  for (const auto& [id, apply_only] : cases) {
    const Case c{"cellxfs_compaction", id};
    const auto wb = LoadAuthored(c);
    ASSERT_TRUE(wb);
    Bytes ours;
    ASSERT_TRUE(SaveXlsx(*wb, &ours));
    const Bytes excel = test::read_file_bytes(FixturePath(c, "_resaved.xlsx"));
    Lines xfs = CellXfLines(ours);
    ASSERT_GT(xfs.size(), apply_only) << id;
    EXPECT_EQ(xfs[apply_only], CellXfLines(excel)[0]) << id;
    xfs.erase(xfs.begin() + apply_only);
    EXPECT_EQ(CellXfLines(excel), xfs) << id;
    Lines refs;
    for (std::string line : SheetRefLines(ours)) {
      const std::size_t at = line.find(" s=");
      if (at == std::string::npos) {
        refs.push_back(line);
        continue;
      }
      const std::size_t end = line.find(' ', at + 3U);
      const std::string n = line.substr(at + 3U, end == std::string::npos ? std::string::npos : end - at - 3U);
      unsigned k = 0U;
      for (const char ch : n) {
        k = k * 10U + static_cast<unsigned>(ch - '0');
      }
      if (!n.empty() && k >= apply_only) {
        line.replace(at + 3U, n.size(), k == apply_only ? std::string() : std::to_string(k - 1U));
      }
      refs.push_back(line);
    }
    EXPECT_EQ(SheetRefLines(excel), refs) << id;
  }
}

}  // namespace
}  // namespace formulon::io
