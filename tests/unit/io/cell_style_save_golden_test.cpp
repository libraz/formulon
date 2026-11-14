// What a save writes for cell styles, against Excel 365 (Mac, ja-JP UI)
// re-saving the same files: the <cellStyles> set, names and order, the
// <cellStyleXfs> sequence, the style each cell / row / column resolves to,
// table cell-style attributes, and the xlsb BrtStyle list. Every expectation
// is read from Excel's own `_resaved` file next to the authored input.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <memory>
#include <set>
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
#include "io/xlsb/record_ids.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon::io {
namespace {

using Bytes = std::vector<std::uint8_t>;
using Lines = std::vector<std::string>;

constexpr std::uint16_t kBrtStyleType = 48;
constexpr const char* kSheetPart = "xl/worksheets/sheet1.xml";

// ---------------------------------------------------------------------------
// Cases
// ---------------------------------------------------------------------------

/// Cases with an Excel `_resaved.xlsx` and `_resaved.xlsb`.
const std::vector<std::string>& XlsbCases() {
  static const std::vector<std::string> cases = {
      "cs1_C0_none",
      "cs1_C1_all_blank",
      "cs1_C2_all_valued",
      "cs1_C3_partial",
      "cs1_C4_row_only",
      "cs1_C5_col_only",
      "cs2_X1_xf_order",
      "cs2_X2_table_only",
      "cs2_X3_sort",
      "cs2_X4_hidden_ctl",
      "cs3_Y1_used_ja_name",
      "cs3_Y2_nosuch",
      "cs3_Y3_unused_en_hdr",
      "cs4_R1_ja_name_no_id_unused",
      "cs4_R2_en_name_no_id_unused",
      "cs4_R3_twins_cells",
      "cs4_R4_same_builtin_twice",
      "cs4_R5_ja_name_no_id_used",
      "cs4_R6_builtin_then_custom_cells",
      "cs4_R7_custom_then_builtin_cells",
      "cs4_R8_custom_then_builtin_unused",
  };
  return cases;
}

/// Every case, the xlsx-only ones included.
const std::vector<std::string>& XlsxCases() {
  static const std::vector<std::string> cases = [] {
    std::vector<std::string> all = XlsbCases();
    for (const char* extra :
         {"cs5_N1_custom_hyojun_first", "c12_U1_col_data_used", "c12_U2_tbl_data_used", "c12_U3_tbl_header_used",
          "c12_U4_col_data_style_used_outside", "c12_M1_col_data_used", "c12_M2_tbl_data_used",
          "c12_M3_tbl_header_used", "c12_M4_col_data_style_used_outside", "c12_M5_col_data_unused"}) {
      all.emplace_back(extra);
    }
    return all;
  }();
  return cases;
}

std::string CaseName(const ::testing::TestParamInfo<std::string>& info) {
  return info.param;
}

// ---------------------------------------------------------------------------
// Files, workbooks, packages
// ---------------------------------------------------------------------------

std::string FixturePath(const std::string& file) {
  return std::string(FORMULON_FIXTURES_DIR) + "/cell_style_save/" + file;
}

Bytes ReadFixture(const std::string& file) {
  return test::read_file_bytes(FixturePath(file));
}

/// Loads the xlsx at `file` and gives it `profile`; null (with a failure
/// recorded) when the read fails.
std::unique_ptr<Workbook> LoadXlsx(const std::string& file, ExcelProfile profile = mac_365_ja_jp_profile()) {
  const Bytes bytes = ReadFixture(file);
  auto read = read_ooxml(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "read_ooxml " << file << ": " << read.error().message;
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_excel_profile(profile);
  return wb;
}

std::unique_ptr<Workbook> LoadXlsbBytes(const Bytes& bytes, const std::string& what,
                                        ExcelProfile profile = mac_365_ja_jp_profile()) {
  auto read = xlsb::read_xlsb(ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "read_xlsb " << what << ": " << read.error().message;
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_excel_profile(profile);
  return wb;
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

bool SaveXlsx(const Workbook& wb, Bytes* out) {
  auto saved = write_ooxml(wb);
  if (!saved) {
    ADD_FAILURE() << "write_ooxml: " << saved.error().message;
    return false;
  }
  *out = std::move(saved.value());
  return true;
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

/// Part `name` of the package as text; empty (with a failure recorded) when absent.
std::string PartText(const Bytes& pkg, const std::string& name) {
  std::string text;
  EXPECT_TRUE(test::extract_part(test::span_of(pkg), name, &text));
  return text;
}

Bytes PartBytes(const Bytes& pkg, const std::string& name) {
  const std::string text = PartText(pkg, name);
  return Bytes(text.begin(), text.end());
}

/// Names of the package parts under `prefix` ending in `suffix`, sorted.
Lines PartsMatching(const Bytes& pkg, const std::string& prefix, const std::string& suffix) {
  ZipReader zip;
  auto opened = zip.open(test::span_of(pkg));
  Lines out;
  if (!opened) {
    ADD_FAILURE() << "zip open: " << opened.error().message;
    return out;
  }
  for (const std::string& name : zip.list_entries()) {
    if (name.size() > prefix.size() + suffix.size() && name.compare(0, prefix.size(), prefix) == 0 &&
        name.compare(name.size() - suffix.size(), suffix.size(), suffix) == 0) {
      out.push_back(name);
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

// ---------------------------------------------------------------------------
// xlsx styles.xml model
// ---------------------------------------------------------------------------

struct ParsedStyles {
  struct Entry {
    std::string name;
    bool has_builtin = false;
    std::uint32_t builtin = 0;
    std::uint32_t i_level = 0;
    bool hidden = false;
    bool custom_builtin = false;
    std::uint32_t xf_id = 0;
  };
  std::vector<Entry> cell_styles;
  std::vector<std::uint32_t> style_xf_num_fmt;
  std::vector<std::uint32_t> cell_xf_xf_id;
};

ParsedStyles ParseStyles(const std::string& xml) {
  ParsedStyles out;
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size())) {
    ADD_FAILURE() << "styles.xml does not parse";
    return out;
  }
  const pugi::xml_node root = doc.child("styleSheet");
  for (const pugi::xml_node n : root.child("cellStyles").children("cellStyle")) {
    ParsedStyles::Entry e;
    e.name = n.attribute("name").value();
    e.has_builtin = !n.attribute("builtinId").empty();
    e.builtin = n.attribute("builtinId").as_uint();
    e.i_level = n.attribute("iLevel").as_uint();
    e.hidden = n.attribute("hidden").as_bool();
    e.custom_builtin = n.attribute("customBuiltin").as_bool();
    e.xf_id = n.attribute("xfId").as_uint();
    out.cell_styles.push_back(std::move(e));
  }
  for (const pugi::xml_node n : root.child("cellStyleXfs").children("xf")) {
    out.style_xf_num_fmt.push_back(n.attribute("numFmtId").as_uint());
  }
  for (const pugi::xml_node n : root.child("cellXfs").children("xf")) {
    out.cell_xf_xf_id.push_back(n.attribute("xfId").as_uint());
  }
  return out;
}

/// One `<cellStyle>` as compared with Excel's: everything but `xfId` and `xr:uid`.
std::string EntryText(const ParsedStyles::Entry& e) {
  return "name=" + e.name + " builtinId=" + (e.has_builtin ? std::to_string(e.builtin) : std::string("-")) +
         " iLevel=" + std::to_string(e.i_level) + " hidden=" + (e.hidden ? "1" : "0") +
         " customBuiltin=" + (e.custom_builtin ? "1" : "0");
}

Lines CellStyleLines(const ParsedStyles& styles) {
  Lines out;
  for (const ParsedStyles::Entry& e : styles.cell_styles) {
    out.push_back(EntryText(e));
  }
  return out;
}

/// What cellXf `s` resolves to: its style xf's number format and the
/// cellStyle naming that style xf.
std::string ResolveXf(const ParsedStyles& styles, std::uint32_t s) {
  if (s >= styles.cell_xf_xf_id.size()) {
    return "cellXf " + std::to_string(s) + " out of range";
  }
  const std::uint32_t style_xf = styles.cell_xf_xf_id[s];
  std::string out = "numFmtId=";
  out += style_xf < styles.style_xf_num_fmt.size() ? std::to_string(styles.style_xf_num_fmt[style_xf])
                                                   : std::string("<style xf ") + std::to_string(style_xf) + " missing>";
  for (const ParsedStyles::Entry& e : styles.cell_styles) {
    if (e.xf_id == style_xf) {
      return out + " style=" + e.name + "/" + (e.has_builtin ? std::to_string(e.builtin) : std::string("-"));
    }
  }
  return out + " style=(none)";
}

/// Every non-zero `s` a worksheet references, keyed by "cell A1", "row 5"
/// (with customFormat) and "col 3" (one entry per column of a `<col>` span).
std::map<std::string, std::uint32_t> SheetRefs(const std::string& xml) {
  std::map<std::string, std::uint32_t> out;
  pugi::xml_document doc;
  if (!doc.load_buffer(xml.data(), xml.size())) {
    ADD_FAILURE() << "worksheet does not parse";
    return out;
  }
  const pugi::xml_node ws = doc.child("worksheet");
  for (const pugi::xml_node row : ws.child("sheetData").children("row")) {
    const std::uint32_t row_s = row.attribute("s").as_uint();
    if (row_s != 0U && row.attribute("customFormat").as_bool()) {
      out["row " + std::to_string(row.attribute("r").as_uint())] = row_s;
    }
    for (const pugi::xml_node c : row.children("c")) {
      const std::uint32_t s = c.attribute("s").as_uint();
      if (s != 0U) {
        out[std::string("cell ") + c.attribute("r").value()] = s;
      }
    }
  }
  for (const pugi::xml_node col : ws.child("cols").children("col")) {
    const std::uint32_t style = col.attribute("style").as_uint();
    if (style == 0U) {
      continue;
    }
    const std::uint32_t last = std::min<std::uint32_t>(col.attribute("max").as_uint(), 16384U);
    for (std::uint32_t k = col.attribute("min").as_uint(); k <= last; ++k) {
      out["col " + std::to_string(k)] = style;
    }
  }
  return out;
}

Lines ResolvedRefLines(const std::string& styles_xml, const std::string& sheet_xml) {
  const ParsedStyles styles = ParseStyles(styles_xml);
  Lines out;
  for (const auto& [key, s] : SheetRefs(sheet_xml)) {
    out.push_back(key + " => " + ResolveXf(styles, s));
  }
  return out;
}

// ---------------------------------------------------------------------------
// xlsx table attributes
// ---------------------------------------------------------------------------

std::string StyleAttrs(const pugi::xml_node& n) {
  return std::string("header=[") + n.attribute("headerRowCellStyle").value() + "] data=[" +
         n.attribute("dataCellStyle").value() + "] totals=[" + n.attribute("totalsRowCellStyle").value() + "]";
}

Lines XlsxTableLines(const Bytes& pkg) {
  Lines out;
  for (const std::string& part : PartsMatching(pkg, "xl/tables/table", ".xml")) {
    const std::string xml = PartText(pkg, part);
    pugi::xml_document doc;
    if (!doc.load_buffer(xml.data(), xml.size())) {
      ADD_FAILURE() << part << " does not parse";
      continue;
    }
    const pugi::xml_node table = doc.child("table");
    const std::string name = table.attribute("name").value();
    out.push_back("table " + name + " " + StyleAttrs(table));
    for (const pugi::xml_node col : table.child("tableColumns").children("tableColumn")) {
      out.push_back("column " + name + "/" + col.attribute("name").value() + " " + StyleAttrs(col));
    }
  }
  std::sort(out.begin(), out.end());
  return out;
}

// ---------------------------------------------------------------------------
// xlsb styles.bin / table parts
// ---------------------------------------------------------------------------

std::string Hex(std::uint32_t v) {
  static const char* kDigits = "0123456789abcdef";
  std::string out = "0x";
  for (int shift = 12; shift >= 0; shift -= 4) {
    out += kDigits[(v >> shift) & 0xFU];
  }
  return out;
}

/// BrtStyle records of a styles part as "xf flags builtin iLevel name".
Lines BrtStyleLines(const Bytes& styles_bin) {
  Lines out;
  ByteSpan cursor{styles_bin.data(), styles_bin.size()};
  while (cursor.size > 0U) {
    auto rec = xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << "styles.bin record: " << rec.error().message;
      break;
    }
    if (rec.value().type != kBrtStyleType) {
      continue;
    }
    const ByteSpan payload = rec.value().payload;
    if (payload.size < 8U) {
      out.push_back("<truncated BrtStyle>");
      continue;
    }
    ByteSpan name_cursor{payload.data + 8U, payload.size - 8U};
    auto name = xlsb::read_xlwidestring(name_cursor);
    out.push_back("xf=" + std::to_string(xlsb::load_u32(payload.data)) +
                  " flags=" + Hex(static_cast<std::uint32_t>(payload.data[4] | (payload.data[5] << 8))) +
                  " builtin=" + std::to_string(payload.data[6]) + " iLevel=" + std::to_string(payload.data[7]) +
                  " name=" + (name ? name.value() : std::string("<unreadable>")));
  }
  return out;
}

/// The three cell-style names after the three caption names of a
/// BrtBeginList / BrtBeginListCol payload that starts the names at `offset`.
std::string RecordStyleNames(const ByteSpan payload, std::size_t offset) {
  if (payload.size < offset) {
    return "<truncated>";
  }
  ByteSpan cursor{payload.data + offset, payload.size - offset};
  for (int i = 0; i < 3; ++i) {
    if (!xlsb::read_xlnullablewidestring(cursor)) {
      return "<unreadable>";
    }
  }
  std::string out;
  for (const char* label : {"header", "data", "totals"}) {
    auto style = xlsb::read_xlnullablewidestring(cursor);
    out += std::string(out.empty() ? "" : " ") + label + "=[" + (style ? style.value() : "<unreadable>") + "]";
  }
  return out;
}

Lines XlsbTableLines(const Bytes& pkg) {
  constexpr std::size_t kListNamesOffset = 64;
  constexpr std::size_t kColumnNamesOffset = 24;
  Lines out;
  for (const std::string& part : PartsMatching(pkg, "xl/tables/table", ".bin")) {
    const Bytes bin = PartBytes(pkg, part);
    std::string line;
    ByteSpan cursor{bin.data(), bin.size()};
    while (cursor.size > 0U) {
      auto rec = xlsb::read_record(cursor);
      if (!rec) {
        ADD_FAILURE() << part << " record: " << rec.error().message;
        break;
      }
      if (rec.value().type == xlsb::kBrtBeginList) {
        line += "list " + RecordStyleNames(rec.value().payload, kListNamesOffset) + "; ";
      } else if (rec.value().type == xlsb::kBrtBeginListCol) {
        line += "column " + RecordStyleNames(rec.value().payload, kColumnNamesOffset) + "; ";
      }
    }
    out.push_back(line);
  }
  std::sort(out.begin(), out.end());
  return out;
}

// ---------------------------------------------------------------------------
// Suites
// ---------------------------------------------------------------------------

class GoldenBase : public ::testing::TestWithParam<std::string> {
 protected:
  const std::string& id() const { return GetParam(); }
  std::string Msg() const { return "case " + id() + " (left: Excel, right: Formulon)"; }
};

class CellStyleSaveGolden : public GoldenBase {};
class CellStyleSaveGoldenXlsb : public GoldenBase {};

/// Formulon's xlsx save of the authored file.
bool SaveAuthoredXlsx(const std::string& id, Bytes* out) {
  const auto wb = LoadXlsx(id + ".xlsx");
  return wb && SaveXlsx(*wb, out);
}

TEST_P(CellStyleSaveGolden, CellStylesMatchExcel) {
  Bytes ours;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsx");
  EXPECT_EQ(CellStyleLines(ParseStyles(PartText(excel, "xl/styles.xml"))),
            CellStyleLines(ParseStyles(PartText(ours, "xl/styles.xml"))))
      << Msg();
}

TEST_P(CellStyleSaveGolden, CellStyleXfNumFmtSequenceMatchesExcel) {
  Bytes ours;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsx");
  EXPECT_EQ(ParseStyles(PartText(excel, "xl/styles.xml")).style_xf_num_fmt,
            ParseStyles(PartText(ours, "xl/styles.xml")).style_xf_num_fmt)
      << Msg();
}

TEST_P(CellStyleSaveGolden, WorksheetReferencesResolveLikeExcel) {
  Bytes ours;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsx");
  EXPECT_EQ(ResolvedRefLines(PartText(excel, "xl/styles.xml"), PartText(excel, kSheetPart)),
            ResolvedRefLines(PartText(ours, "xl/styles.xml"), PartText(ours, kSheetPart)))
      << Msg();
}

TEST_P(CellStyleSaveGolden, TableCellStyleAttributesMatchExcel) {
  Bytes ours;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsx");
  EXPECT_EQ(XlsxTableLines(excel), XlsxTableLines(ours)) << Msg();
}

TEST_P(CellStyleSaveGolden, XlsxSaveIsIdempotent) {
  Bytes first;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &first));
  const auto reloaded = LoadXlsxBytes(first, id() + " first save");
  ASSERT_TRUE(reloaded);
  Bytes second;
  ASSERT_TRUE(SaveXlsx(*reloaded, &second));
  EXPECT_EQ(PartText(first, "xl/styles.xml"), PartText(second, "xl/styles.xml")) << Msg();
}

TEST_P(CellStyleSaveGolden, SaveLeavesTheStyleModelUnchanged) {
  const auto wb = LoadXlsx(id() + ".xlsx");
  ASSERT_TRUE(wb);
  const std::string before = write_styles(wb->styles());
  Bytes saved;
  ASSERT_TRUE(SaveXlsx(*wb, &saved));
  EXPECT_EQ(write_styles(wb->styles()), before) << Msg() << " after xlsx save";
  ASSERT_TRUE(SaveXlsb(*wb, &saved));
  EXPECT_EQ(write_styles(wb->styles()), before) << Msg() << " after xlsb save";
}

TEST_P(CellStyleSaveGoldenXlsb, BrtStyleFromAuthoredMatchesExcel) {
  const auto wb = LoadXlsx(id() + ".xlsx");
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsb");
  EXPECT_EQ(BrtStyleLines(PartBytes(excel, "xl/styles.bin")), BrtStyleLines(PartBytes(ours, "xl/styles.bin")))
      << Msg() << " authored xlsx -> xlsb";
}

TEST_P(CellStyleSaveGoldenXlsb, BrtStyleFromExcelResavedXlsxMatchesExcel) {
  const auto wb = LoadXlsx(id() + "_resaved.xlsx");
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsb");
  EXPECT_EQ(BrtStyleLines(PartBytes(excel, "xl/styles.bin")), BrtStyleLines(PartBytes(ours, "xl/styles.bin")))
      << Msg() << " Excel resaved xlsx -> xlsb";
}

// Excel's xlsb is the second save (authored -> xlsx -> reopen -> xlsb); a
// table attribute naming a style the first save dropped is removed by it.
TEST_P(CellStyleSaveGoldenXlsb, TableCellStyleNamesMatchExcel) {
  Bytes first;
  ASSERT_TRUE(SaveAuthoredXlsx(id(), &first));
  const auto wb = LoadXlsxBytes(first, id() + " first save");
  ASSERT_TRUE(wb);
  Bytes ours;
  ASSERT_TRUE(SaveXlsb(*wb, &ours));
  const Bytes excel = ReadFixture(id() + "_resaved.xlsb");
  EXPECT_EQ(XlsbTableLines(excel), XlsbTableLines(ours)) << Msg();
}

TEST_P(CellStyleSaveGoldenXlsb, XlsbSaveIsIdempotent) {
  const auto wb = LoadXlsx(id() + ".xlsx");
  ASSERT_TRUE(wb);
  Bytes first;
  ASSERT_TRUE(SaveXlsb(*wb, &first));
  const auto reloaded = LoadXlsbBytes(first, id() + " first save");
  ASSERT_TRUE(reloaded);
  Bytes second;
  ASSERT_TRUE(SaveXlsb(*reloaded, &second));
  EXPECT_EQ(PartBytes(first, "xl/styles.bin"), PartBytes(second, "xl/styles.bin")) << Msg();
}

INSTANTIATE_TEST_SUITE_P(Cases, CellStyleSaveGolden, ::testing::ValuesIn(XlsxCases()), CaseName);
INSTANTIATE_TEST_SUITE_P(Cases, CellStyleSaveGoldenXlsb, ::testing::ValuesIn(XlsbCases()), CaseName);

// ---------------------------------------------------------------------------
// Merged same-name styles survive a round trip as one
// ---------------------------------------------------------------------------

TEST(CellStyleSaveGoldenRoundTrip, MergedHeading2IsOneStyleWithItsCellsFollowing) {
  constexpr const char* kMerged = "\xE8\xA6\x8B\xE5\x87\xBA\xE3\x81\x97 2";  // 見出し 2
  constexpr std::uint32_t kHeading2 = 17U;
  for (const char* id : {"cs1_C1_all_blank", "cs4_R6_builtin_then_custom_cells", "cs4_R7_custom_then_builtin_cells"}) {
    SCOPED_TRACE(std::string("case ") + id);
    const Bytes authored = ReadFixture(std::string(id) + ".xlsx");
    const ParsedStyles before = ParseStyles(PartText(authored, "xl/styles.xml"));
    std::set<std::uint32_t> merged_xfs;
    for (const ParsedStyles::Entry& e : before.cell_styles) {
      if (e.name == kMerged || (e.has_builtin && e.builtin == kHeading2)) {
        merged_xfs.insert(e.xf_id);
      }
    }
    ASSERT_FALSE(merged_xfs.empty());
    const auto before_refs = SheetRefs(PartText(authored, kSheetPart));

    const auto wb = LoadXlsx(std::string(id) + ".xlsx");
    ASSERT_TRUE(wb);
    Bytes saved;
    ASSERT_TRUE(SaveXlsx(*wb, &saved));
    const ParsedStyles after = ParseStyles(PartText(saved, "xl/styles.xml"));
    std::vector<std::uint32_t> survivors;
    for (const ParsedStyles::Entry& e : after.cell_styles) {
      if (e.name == kMerged) {
        survivors.push_back(e.xf_id);
      }
    }
    ASSERT_EQ(survivors.size(), 1U) << "saved <cellStyles> holds " << survivors.size() << " styles named " << kMerged;
    const auto after_refs = SheetRefs(PartText(saved, kSheetPart));
    std::size_t followed = 0;
    for (const auto& [key, s] : before_refs) {
      if (s >= before.cell_xf_xf_id.size() || merged_xfs.count(before.cell_xf_xf_id[s]) == 0U) {
        continue;
      }
      ++followed;
      ASSERT_EQ(after_refs.count(key), 1U) << key;
      ASSERT_LT(after_refs.at(key), after.cell_xf_xf_id.size()) << key;
      EXPECT_EQ(after.cell_xf_xf_id[after_refs.at(key)], survivors[0])
          << key << ": cellXf must name the surviving style xf";
    }
    EXPECT_GT(followed, 0U) << "no worksheet reference used a merged style";

    const auto reloaded = LoadXlsxBytes(saved, std::string(id) + " saved");
    ASSERT_TRUE(reloaded);
    const StylesTable& model = reloaded->styles();
    std::size_t named = 0;
    std::uint32_t model_xf = 0;
    for (const CellStyleRecord& rec : model.cell_styles) {
      if (rec.name == kMerged) {
        ++named;
        model_xf = rec.xf_id;
      }
    }
    ASSERT_EQ(named, 1U) << "reloaded model holds " << named << " styles named " << kMerged;
    EXPECT_EQ(model_xf, survivors[0]);
    for (const auto& [key, s] : before_refs) {
      if (s >= before.cell_xf_xf_id.size() || merged_xfs.count(before.cell_xf_xf_id[s]) == 0U) {
        continue;
      }
      const std::uint32_t saved_s = after_refs.at(key);
      ASSERT_LT(saved_s, model.cell_xfs.size()) << key;
      EXPECT_EQ(model.cell_xfs[saved_s].xf_id, model_xf) << key << ": reloaded cellXf must name the survivor";
    }
  }
}

// ---------------------------------------------------------------------------
// xlsb styles.bin freshness: a retained part is rewritten when the plan differs
// ---------------------------------------------------------------------------

constexpr const char* kJaHeading1 = "\xE8\xA6\x8B\xE5\x87\xBA\xE3\x81\x97 1";  // 見出し 1

bool HasName(const Lines& styles, const std::string& name) {
  const std::string needle = " name=" + name;
  return std::any_of(styles.begin(), styles.end(), [&](const std::string& line) {
    return line.size() >= needle.size() && line.compare(line.size() - needle.size(), needle.size(), needle) == 0;
  });
}

TEST(CellStyleSaveGoldenXlsbFreshness, ClearingTheOnlyUserOfAStyleRewritesStylesBin) {
  const Bytes retained = ReadFixture("cs1_C3_partial_resaved.xlsb");
  const auto wb = LoadXlsbBytes(retained, "cs1_C3_partial_resaved.xlsb");
  ASSERT_TRUE(wb);

  // Control: untouched, the style a cell still uses is saved.
  Bytes untouched;
  ASSERT_TRUE(SaveXlsb(*wb, &untouched));
  EXPECT_TRUE(HasName(BrtStyleLines(PartBytes(untouched, "xl/styles.bin")), kJaHeading1));

  std::set<std::uint32_t> style_xfs;
  for (const CellStyleRecord& rec : wb->styles().cell_styles) {
    if (rec.name == kJaHeading1) {
      style_xfs.insert(rec.xf_id);
    }
  }
  ASSERT_FALSE(style_xfs.empty());
  std::set<std::uint32_t> cell_xfs;
  for (std::size_t i = 0; i < wb->styles().cell_xfs.size(); ++i) {
    if (style_xfs.count(wb->styles().cell_xfs[i].xf_id) != 0U) {
      cell_xfs.insert(static_cast<std::uint32_t>(i));
    }
  }
  ASSERT_FALSE(cell_xfs.empty());
  std::vector<std::pair<std::uint32_t, std::uint32_t>> users;
  for (std::uint32_t row = 0; row < 100U; ++row) {
    for (std::uint32_t col = 0; col < 20U; ++col) {
      const auto slot = wb->sheet(0).cell_style_at(row, col);
      if (slot && cell_xfs.count(slot->xf) != 0U) {
        users.emplace_back(row, col);
      }
    }
  }
  ASSERT_EQ(users.size(), 1U) << "expected A20 as the only cell using " << kJaHeading1;
  EXPECT_EQ(users[0], (std::pair<std::uint32_t, std::uint32_t>(19U, 0U)));

  const auto cleared_xf = wb->clear_cell_xf_index(0, users[0].first, users[0].second);
  ASSERT_TRUE(static_cast<bool>(cleared_xf)) << cleared_xf.error().message;
  Bytes cleared;
  ASSERT_TRUE(SaveXlsb(*wb, &cleared));
  const Lines styles = BrtStyleLines(PartBytes(cleared, "xl/styles.bin"));
  EXPECT_FALSE(HasName(styles, kJaHeading1)) << "BrtStyle list still names the unused style";
  EXPECT_NE(PartBytes(cleared, "xl/styles.bin"), PartBytes(retained, "xl/styles.bin"))
      << "styles.bin was passed through from the retained part";
}

// The stored part is kept under another profile; a style edit regenerates it
// with the names of that profile's locale.
TEST(CellStyleSaveGoldenXlsbFreshness, SavingUnderAnotherProfileKeepsTheStoredNamesUntilAStyleEdit) {
  const Bytes retained = ReadFixture("cs1_C3_partial_resaved.xlsb");
  auto wb = LoadXlsbBytes(retained, "cs1_C3_partial_resaved.xlsb", mac_365_en_us_profile());
  ASSERT_TRUE(wb);
  Bytes kept;
  ASSERT_TRUE(SaveXlsb(*wb, &kept));
  EXPECT_EQ(PartBytes(kept, "xl/styles.bin"), PartBytes(retained, "xl/styles.bin"));

  wb->mutable_styles().fonts.push_back(wb->styles().fonts.front());
  Bytes saved;
  ASSERT_TRUE(SaveXlsb(*wb, &saved));
  const Lines styles = BrtStyleLines(PartBytes(saved, "xl/styles.bin"));
  for (const char* english : {"Comma", "Heading 1", "Normal"}) {
    EXPECT_TRUE(HasName(styles, english)) << english << " missing from BrtStyle list";
  }
  EXPECT_FALSE(HasName(styles, kJaHeading1));
}

}  // namespace
}  // namespace formulon::io
