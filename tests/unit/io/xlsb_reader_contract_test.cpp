//
// Container-boundary contracts of the XLSB read path.
//
// Each group below pins one rule the reader owes its callers and that no
// literal-expectation fidelity test can express, because the rule is
// about what the reader must NOT do:
//
//   * A `BrtRowHdr`'s `miyRw` is the rendered row height Excel stores for
//     every row; only the `fUnsynced` flag says the height is custom.
//   * A `BrtName` is dropped only when it is one of Excel's storage
//     placeholders. `fHidden` alone is a user-visible property, not a
//     reason to discard the name.
//   * A `BrtExternSheet` entry belonging to another supporting book must
//     not resolve against this workbook's sheet list.
//   * Nothing encoded as MS-XLSB binary records may reach an OOXML
//     package, whichever container the in-memory workbook came from.

#include <algorithm>
#include <array>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string FixturePath(const char* name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
}

// ---------------------------------------------------------------------------
// Minimal synthetic `.xlsb` package builders. Mirrors the record framing in
// `io/xlsb/record.{h,cpp}` (1-/2-byte type, 1..4-byte size, payload) so a
// single record's field layout can be driven directly, which neither the
// writer nor a committed fixture allows.
// ---------------------------------------------------------------------------

void AppendVarInt(std::vector<std::uint8_t>& out, std::uint32_t v, std::size_t max_bytes) {
  for (std::size_t i = 0; i < max_bytes; ++i) {
    const std::uint8_t byte = static_cast<std::uint8_t>(v & 0x7FU);
    v >>= 7;
    if (v == 0) {
      out.push_back(byte);
      return;
    }
    out.push_back(static_cast<std::uint8_t>(byte | 0x80U));
  }
}

void AppendRecord(std::vector<std::uint8_t>& out, std::uint16_t type, const std::vector<std::uint8_t>& payload) {
  AppendVarInt(out, type, /*max_bytes=*/2);
  AppendVarInt(out, static_cast<std::uint32_t>(payload.size()), /*max_bytes=*/4);
  out.insert(out.end(), payload.begin(), payload.end());
}

void AppendU32(std::vector<std::uint8_t>& out, std::uint32_t v) {
  out.push_back(static_cast<std::uint8_t>(v & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 8) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 16) & 0xFFU));
  out.push_back(static_cast<std::uint8_t>((v >> 24) & 0xFFU));
}

void AppendDouble(std::vector<std::uint8_t>& out, double v) {
  std::uint64_t bits = 0;
  std::memcpy(&bits, &v, sizeof(v));
  AppendU32(out, static_cast<std::uint32_t>(bits & 0xFFFFFFFFU));
  AppendU32(out, static_cast<std::uint32_t>((bits >> 32) & 0xFFFFFFFFU));
}

void AppendXLWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  AppendU32(out, static_cast<std::uint32_t>(s.size()));
  for (char c : s) {
    out.push_back(static_cast<std::uint8_t>(c));
    out.push_back(0);
  }
}

void AppendXLNullableWideString(std::vector<std::uint8_t>& out, std::string_view s) {
  if (s.empty()) {
    AppendU32(out, 0xFFFFFFFFU);
    return;
  }
  AppendXLWideString(out, s);
}

std::string ContentTypesXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/workbook.bin\" "
      "ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>"
      "<Override PartName=\"/xl/worksheets/sheet2.bin\" ContentType=\"application/vnd.ms-excel.worksheet\"/>"
      "</Types>");
}

std::string PackageRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.bin\"/>"
      "</Relationships>");
}

std::string WorkbookRelsXml() {
  return std::string(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rIdSheet1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.bin\"/>"
      "<Relationship Id=\"rIdSheet2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet2.bin\"/>"
      "</Relationships>");
}

/// Builds `xl/workbook.bin` with sheets `Alpha` / `Beta` and, when
/// `extern_sheets` is non-empty, one `BrtExternSheet` record carrying
/// those `(iSupBook, itabFirst, itabLast)` triples in order.
std::vector<std::uint8_t> WorkbookBin(const std::vector<std::array<std::uint32_t, 3>>& extern_sheets = {}) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 131, {});  // BrtBeginBook
  AppendRecord(body, 143, {});  // BrtBeginBundleShs
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);  // hsState: visible
    AppendU32(p, 1);  // iTabID
    AppendXLNullableWideString(p, "rIdSheet1");
    AppendXLWideString(p, "Alpha");
    AppendRecord(body, 156, p);  // BrtBundleSh
  }
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendU32(p, 2);
    AppendXLNullableWideString(p, "rIdSheet2");
    AppendXLWideString(p, "Beta");
    AppendRecord(body, 156, p);  // BrtBundleSh
  }
  AppendRecord(body, 144, {});  // BrtEndBundleShs
  if (!extern_sheets.empty()) {
    std::vector<std::uint8_t> p;
    AppendU32(p, static_cast<std::uint32_t>(extern_sheets.size()));
    for (const std::array<std::uint32_t, 3>& entry : extern_sheets) {
      AppendU32(p, entry[0]);  // iSupBook
      AppendU32(p, entry[1]);  // itabFirst
      AppendU32(p, entry[2]);  // itabLast
    }
    AppendRecord(body, 362, p);  // BrtExternSheet
  }
  AppendRecord(body, 132, {});  // BrtEndBook
  return body;
}

/// Builds a layout-only sheet whose `BrtRowHdr` records carry the given
/// `(row, miyRw, flags2)` triples in order.
std::vector<std::uint8_t> SheetBinRowHeaders(const std::vector<std::array<std::uint32_t, 3>>& rows) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData
  for (const std::array<std::uint32_t, 3>& row : rows) {
    std::vector<std::uint8_t> p;
    AppendU32(p, row[0]);  // rw
    AppendU32(p, 0);       // iStyleRef
    p.push_back(static_cast<std::uint8_t>(row[1] & 0xFFU));
    p.push_back(static_cast<std::uint8_t>((row[1] >> 8) & 0xFFU));  // miyRw
    p.push_back(0);                                                 // flags1
    p.push_back(static_cast<std::uint8_t>(row[2] & 0xFFU));         // flags2
    p.push_back(0);                                                 // fPhShow
    AppendU32(p, 0);                                                // ccolspan
    AppendRecord(body, 0, p);                                       // BrtRowHdr
  }
  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

/// Builds a sheet holding one `BrtFmlaNum` at A1: cached numeric result
/// `cached` plus a `CellParsedFormula` whose `rgce` is `rgce`.
std::vector<std::uint8_t> SheetBinFmlaNum(double cached, const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});  // BrtBeginSheet
  AppendRecord(body, 145, {});  // BrtBeginSheetData
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendRecord(body, 0, p);  // BrtRowHdr
  }
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);  // column
    p.push_back(0);   // style[0]
    p.push_back(0);   // style[1]
    p.push_back(0);   // style[2]
    p.push_back(0);   // fPhShow
    AppendDouble(p, cached);
    p.push_back(0);  // grbitFlags lo
    p.push_back(0);  // grbitFlags hi
    AppendU32(p, static_cast<std::uint32_t>(rgce.size()));
    p.insert(p.end(), rgce.begin(), rgce.end());
    AppendU32(p, 0);           // cb: no rgcb
    AppendRecord(body, 9, p);  // BrtFmlaNum
  }
  AppendRecord(body, 146, {});  // BrtEndSheetData
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

struct PartFile {
  std::string path;
  std::vector<std::uint8_t> body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const PartFile& p : parts) {
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, p.path.c_str(), p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_TRUE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  mz_zip_writer_end(&writer);
  return out;
}

std::vector<std::uint8_t> StringToBytes(std::string_view s) {
  return std::vector<std::uint8_t>(s.begin(), s.end());
}

/// Assembles a two-sheet package: `Alpha` carries `sheet1`, `Beta` is
/// empty. `extern_sheets` populates the optional `BrtExternSheet` record.
std::vector<std::uint8_t> BuildPackage(const std::vector<std::uint8_t>& sheet1,
                                       const std::vector<std::array<std::uint32_t, 3>>& extern_sheets = {}) {
  std::vector<PartFile> parts;
  parts.push_back({"[Content_Types].xml", StringToBytes(ContentTypesXml())});
  parts.push_back({"_rels/.rels", StringToBytes(PackageRelsXml())});
  parts.push_back({"xl/_rels/workbook.bin.rels", StringToBytes(WorkbookRelsXml())});
  parts.push_back({"xl/workbook.bin", WorkbookBin(extern_sheets)});
  parts.push_back({"xl/worksheets/sheet1.bin", sheet1});
  parts.push_back({"xl/worksheets/sheet2.bin", SheetBinRowHeaders({})});
  return BuildZip(parts);
}

// ---------------------------------------------------------------------------
// Row-height presence.
// ---------------------------------------------------------------------------

TEST(XlsbRowLayout, OnlyTheUnsyncedFlagMakesAHeightCustom) {
  // Row 0: rendered height stored, no flag -> not an override at all.
  // Row 1: same height with fUnsynced (0x20) -> a custom height of 20pt.
  // Row 2: hidden (0x10) without fUnsynced -> an override, but no height.
  const std::vector<std::uint8_t> archive =
      BuildPackage(SheetBinRowHeaders({{{0U, 400U, 0x00U}}, {{1U, 400U, 0x20U}}, {{2U, 400U, 0x10U}}}));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  const SheetLayout& layout = result.value().workbook.sheet(0).layout();
  ASSERT_EQ(layout.row_overrides.size(), 2U);

  EXPECT_EQ(layout.row_overrides[0].row, 1U);
  EXPECT_TRUE(layout.row_overrides[0].has_height);
  EXPECT_DOUBLE_EQ(layout.row_overrides[0].height, 20.0);

  EXPECT_EQ(layout.row_overrides[1].row, 2U);
  EXPECT_FALSE(layout.row_overrides[1].has_height);
  EXPECT_DOUBLE_EQ(layout.row_overrides[1].height, 0.0);
  EXPECT_TRUE(layout.row_overrides[1].hidden);
}

TEST(XlsbRowLayout, RealWorkbookMatchesItsOoxmlTwin) {
  // The same workbook exported by Excel to both containers. Every row of
  // the binary export carries a rendered `miyRw`; the XML export writes
  // no `ht` at all, so both readers must agree on "no row override".
  const std::vector<std::uint8_t> xlsb_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  const std::vector<std::uint8_t> xlsx_bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsx"));
  ASSERT_FALSE(xlsb_bytes.empty());
  ASSERT_FALSE(xlsx_bytes.empty());

  auto xlsb = io::xlsb::read_xlsb(test::span_of(xlsb_bytes));
  ASSERT_TRUE(static_cast<bool>(xlsb)) << xlsb.error().message;
  auto xlsx = io::read_ooxml(test::span_of(xlsx_bytes));
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;

  const Workbook& b = xlsb.value().workbook;
  const Workbook& x = xlsx.value().workbook;
  ASSERT_EQ(b.sheet_count(), x.sheet_count());
  for (std::size_t s = 0; s < b.sheet_count(); ++s) {
    const std::vector<RowLayout>& rb = b.sheet(s).layout().row_overrides;
    const std::vector<RowLayout>& rx = x.sheet(s).layout().row_overrides;
    ASSERT_EQ(rb.size(), rx.size()) << "sheet " << s;
    for (std::size_t i = 0; i < rb.size(); ++i) {
      EXPECT_EQ(rb[i].row, rx[i].row) << "sheet " << s << " row override " << i;
      EXPECT_EQ(rb[i].has_height, rx[i].has_height) << "sheet " << s << " row override " << i;
      EXPECT_DOUBLE_EQ(rb[i].height, rx[i].height) << "sheet " << s << " row override " << i;
    }
  }
}

TEST(XlsbRowLayout, CustomHeightSurvivesAWriteReadCycle) {
  // The positive control for the flag-only presence rule: a modelled
  // custom height must still be recognised after the writer sets
  // fUnsynced for it.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  RowLayout custom;
  custom.row = 3;
  custom.height = 33.0;
  custom.has_height = true;
  wb.sheet(0).mutable_layout().row_overrides.push_back(custom);

  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  auto reread = io::xlsb::read_xlsb(test::span_of(written.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;

  const std::vector<RowLayout>& rows = reread.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 1U);
  EXPECT_EQ(rows[0].row, 3U);
  EXPECT_TRUE(rows[0].has_height);
  EXPECT_DOUBLE_EQ(rows[0].height, 33.0);
}

// ---------------------------------------------------------------------------
// Defined names.
// ---------------------------------------------------------------------------

TEST(XlsbDefinedNames, HiddenNamesSurviveARoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  std::vector<io::DefinedName> names;
  io::DefinedName visible;
  visible.name = "Vis";
  visible.formula = "Sheet1!$A$1";
  visible.local_sheet_id = -1;
  names.push_back(visible);
  io::DefinedName hidden;
  hidden.name = "Hid";
  hidden.formula = "Sheet1!$B$1";
  hidden.local_sheet_id = -1;
  hidden.hidden = true;
  names.push_back(hidden);
  wb.set_defined_names(std::move(names));

  auto written = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  auto reread = io::xlsb::read_xlsb(test::span_of(written.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;

  const std::vector<io::DefinedName>& out = reread.value().workbook.defined_names();
  ASSERT_EQ(out.size(), 2U);
  EXPECT_EQ(out[0].name, "Vis");
  EXPECT_EQ(out[0].formula, "Sheet1!$A$1");
  EXPECT_FALSE(out[0].hidden);
  EXPECT_EQ(out[1].name, "Hid");
  EXPECT_EQ(out[1].formula, "Sheet1!$B$1");
  EXPECT_TRUE(out[1].hidden);
  EXPECT_EQ(reread.value().undecoded_defined_name_count, 0U);
}

TEST(XlsbDefinedNames, StoragePlaceholdersStayOutOfTheNameTable) {
  // The fixture's workbook part declares six `_xlfn.*` future-function
  // placeholders and one `_xlpm.*` LET parameter alongside the single
  // user-visible name. Only the latter is a defined name.
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto result = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  const std::vector<io::DefinedName>& names = result.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "Rate");
  EXPECT_FALSE(names[0].hidden);
}

// ---------------------------------------------------------------------------
// External-workbook 3-D references.
// ---------------------------------------------------------------------------

/// Encodes `formula` against `sheet_names` and returns its `rgce`.
std::vector<std::uint8_t> EncodeRgce(std::string_view formula, const std::vector<std::string>& sheet_names,
                                     const io::xlsb::SheetRangeTable& sheet_ranges) {
  Arena arena;
  parser::Parser parser(formula, arena);
  parser::AstNode* root = parser.parse();
  EXPECT_NE(root, nullptr);
  EXPECT_TRUE(parser.errors().empty()) << "parse errors for: " << formula;
  if (root == nullptr) {
    return {};
  }
  auto encoded = io::xlsb::encode_ptgs(*root, sheet_names, sheet_ranges, {});
  EXPECT_TRUE(static_cast<bool>(encoded)) << (encoded ? "" : encoded.error().message);
  if (!encoded) {
    return {};
  }
  return std::move(encoded.value().rgce);
}

TEST(XlsbExternalReference, LocalSupBookStillResolves) {
  // Control for the case below: the identical Ptg stream, differing only
  // in the ExternSheet entry's `iSupBook`, decodes normally.
  const std::vector<std::string> sheets = {"Alpha", "Beta"};
  const std::vector<std::uint8_t> rgce = EncodeRgce("Beta!A1*2", sheets, {{1, 1}});
  ASSERT_FALSE(rgce.empty());

  const std::vector<std::uint8_t> archive =
      BuildPackage(SheetBinFmlaNum(14.0, rgce), {{{0U, 1U, 1U}}});  // iSupBook 0 = this workbook
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  const Cell* cell = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=Beta!A1*2");
  EXPECT_EQ(result.value().undecoded_formula_count, 0U);
}

TEST(XlsbExternalReference, ExternalSupBookKeepsTheCachedValueAndReportsTheToken) {
  // The same stream with `iSupBook == 1`: the sheet indices belong to a
  // supporting workbook, so binding them to `Beta` would silently compute
  // on unrelated data. The reader surfaces the token instead and leaves
  // Excel's cached result in place.
  const std::vector<std::string> sheets = {"Alpha", "Beta"};
  const std::vector<std::uint8_t> rgce = EncodeRgce("Beta!A1*2", sheets, {{1, 1}});
  ASSERT_FALSE(rgce.empty());

  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFmlaNum(14.0, rgce), {{{1U, 1U, 1U}}});
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  const Cell* cell = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty()) << "fabricated formula: " << cell->formula_text;
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 14.0);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
}

TEST(XlsbExternalReference, ExternalSupBookAppliesToAreaTokensToo) {
  const std::vector<std::string> sheets = {"Alpha", "Beta"};
  const std::vector<std::uint8_t> rgce = EncodeRgce("SUM(Beta!A1:B2)", sheets, {{1, 1}});
  ASSERT_FALSE(rgce.empty());

  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFmlaNum(3.0, rgce), {{{2U, 1U, 1U}}});
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;

  const Cell* cell = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty()) << "fabricated formula: " << cell->formula_text;
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 3.0);
  EXPECT_EQ(result.value().undecoded_formula_count, 1U);
}

// ---------------------------------------------------------------------------
// Record disposition.
// ---------------------------------------------------------------------------

struct RawRecord {
  std::uint16_t type;
  std::vector<std::uint8_t> payload;
};

/// The `BrtWsProp` payload: three flag bytes, an 8-byte `BrtColor` tab
/// colour, the `rwSync` / `colSync` anchors, and the VBA code name.
std::vector<std::uint8_t> WsPropPayload(std::uint16_t flags, std::uint8_t color_type, std::uint8_t color_index,
                                        std::uint32_t argb, std::string_view code_name) {
  std::vector<std::uint8_t> p;
  p.push_back(static_cast<std::uint8_t>(flags & 0xFFU));
  p.push_back(static_cast<std::uint8_t>((flags >> 8) & 0xFFU));
  p.push_back(0x02U);
  p.push_back(color_type == 0U ? std::uint8_t{0}
                               : static_cast<std::uint8_t>((static_cast<unsigned>(color_type) << 1U) | 0x01U));
  p.push_back(color_index);
  p.push_back(0);  // tint lo
  p.push_back(0);  // tint hi
  p.push_back(static_cast<std::uint8_t>((argb >> 16) & 0xFFU));
  p.push_back(static_cast<std::uint8_t>((argb >> 8) & 0xFFU));
  p.push_back(static_cast<std::uint8_t>(argb & 0xFFU));
  p.push_back(static_cast<std::uint8_t>((argb >> 24) & 0xFFU));
  AppendU32(p, 0xFFFFFFFFU);  // rwSync
  AppendU32(p, 0xFFFFFFFFU);  // colSync
  AppendXLWideString(p, code_name);
  return p;
}

/// The flag word Excel writes, and the one `sheet_writer.cpp` re-derives.
constexpr std::uint16_t kDefaultWsPropFlags = 0x04C9U;

/// `BrtSel`, the per-view selection state: a record Excel writes into
/// every worksheet prefix, and one the model has no field for.
constexpr std::uint16_t kBrtSel = 152;

std::vector<std::uint8_t> SelectionPayload() {
  std::vector<std::uint8_t> p;
  AppendU32(p, 3U);  // pnn: bottom-right pane
  AppendU32(p, 0U);  // rwAct
  AppendU32(p, 0U);  // colAct
  AppendU32(p, 0U);  // irefAct
  AppendU32(p, 0U);  // cref
  return p;
}

/// Builds a worksheet part carrying the frame Excel emits around the cell
/// table -- sheet begin, properties, dimension, the view collection,
/// default formatting, the column collection, the cell table, sheet end --
/// with `prefix_extra` spliced in ahead of the cell table and `tail_extra`
/// after it. The cell table itself is left empty; these tests are about
/// what happens to the records around it.
std::vector<std::uint8_t> SheetBinFramed(const std::vector<std::uint8_t>& ws_prop,
                                         const std::vector<RawRecord>& prefix_extra = {},
                                         const std::vector<RawRecord>& tail_extra = {}) {
  std::vector<std::uint8_t> body;
  AppendRecord(body, 129, {});       // BrtBeginSheet
  AppendRecord(body, 147, ws_prop);  // BrtWsProp
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0);
    AppendU32(p, 0);
    AppendU32(p, 0);
    AppendU32(p, 0);
    AppendRecord(body, 148, p);  // BrtWsDim
  }
  AppendRecord(body, 133, {});  // BrtBeginWsViews
  {
    std::vector<std::uint8_t> p;
    p.push_back(0x1CU);  // flags: grid lines, headings, zeros
    p.push_back(0);
    AppendU32(p, 0);  // view type
    AppendU32(p, 0);  // rwTop
    AppendU32(p, 0);  // colLeft
    p.push_back(0x40U);
    p.push_back(0);
    p.push_back(0);
    p.push_back(0);
    p.push_back(100U);  // wScale
    p.push_back(0);
    AppendRecord(body, 137, p);  // BrtBeginWsView
  }
  AppendRecord(body, 138, {});  // BrtEndWsView
  AppendRecord(body, 134, {});  // BrtEndWsViews
  {
    std::vector<std::uint8_t> p;
    AppendU32(p, 0xFFFFFFFFU);  // dxGCol: absent
    p.push_back(8U);            // cchDefColWidth
    p.push_back(0);
    p.push_back(0x2CU);  // miyDefRwHeight: 300 twips
    p.push_back(0x01U);
    AppendU32(p, 0);
    AppendRecord(body, 485, p);  // BrtWsFmtInfo
  }
  AppendRecord(body, 390, {});  // BrtBeginColInfos
  AppendRecord(body, 391, {});  // BrtEndColInfos
  for (const RawRecord& record : prefix_extra) {
    AppendRecord(body, record.type, record.payload);
  }
  AppendRecord(body, 145, {});  // BrtBeginSheetData
  AppendRecord(body, 146, {});  // BrtEndSheetData
  for (const RawRecord& record : tail_extra) {
    AppendRecord(body, record.type, record.payload);
  }
  AppendRecord(body, 130, {});  // BrtEndSheet
  return body;
}

TEST(XlsbRecordDisposition, TheFrameTheWriterReDerivesIsNotCounted) {
  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFramed(
      WsPropPayload(kDefaultWsPropFlags, /*color_type=*/0U, /*color_index=*/0x40U, /*argb=*/0U, /*code_name=*/"")));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().dropped_record_count, 0U);
}

TEST(XlsbRecordDisposition, APrefixRecordWithNoModelFieldIsCounted) {
  // The worksheet prefix is re-derived whole, so there is no position to
  // splice foreign bytes into. A prefix record the reader does not decode
  // is therefore lost -- and has to say so.
  const std::vector<std::uint8_t> archive = BuildPackage(
      SheetBinFramed(WsPropPayload(kDefaultWsPropFlags, 0U, 0x40U, 0U, ""), {{kBrtSel, SelectionPayload()}}));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().dropped_record_count, 1U);
  EXPECT_EQ(result.value().workbook.sheet_count(), 2U) << "the load itself still succeeds";
}

TEST(XlsbRecordDisposition, TheSameRecordInTheTailIsRetainedInsteadOfCounted) {
  // Past the cell table the writer appends retained bytes verbatim, so the
  // record survives and nothing is reported lost.
  const std::vector<std::uint8_t> archive = BuildPackage(
      SheetBinFramed(WsPropPayload(kDefaultWsPropFlags, 0U, 0x40U, 0U, ""), {}, {{kBrtSel, SelectionPayload()}}));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().dropped_record_count, 0U);
  EXPECT_FALSE(result.value().workbook.sheet(0).xlsb_tail().empty());

  auto written = io::xlsb::write_xlsb(result.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  auto reread = io::xlsb::read_xlsb(test::span_of(written.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  EXPECT_EQ(reread.value().workbook.sheet(0).xlsb_tail().before_merges,
            result.value().workbook.sheet(0).xlsb_tail().before_merges);
}

TEST(XlsbRecordDisposition, RealWorkbookCountsOnlyItsSelectionRecords) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(bytes.empty());
  auto result = io::xlsb::read_xlsb(test::span_of(bytes));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  // Both sheets carry a `BrtSel` and nothing else the reader leaves
  // behind. Pinning the exact number is the point: a record type that
  // starts being dropped becomes visible here.
  EXPECT_EQ(result.value().dropped_record_count, 2U);
}

// ---------------------------------------------------------------------------
// Worksheet properties.
// ---------------------------------------------------------------------------

TEST(XlsbWorksheetProperties, TabColorAndCodeNameReachTheSheetPrFragment) {
  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFramed(
      WsPropPayload(kDefaultWsPropFlags, /*color_type=*/2U, /*color_index=*/0U, 0xFFFF0000U, "AlphaCode")));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().workbook.sheet(0).print_settings().sheet_pr_xml,
            "<sheetPr codeName=\"AlphaCode\"><tabColor rgb=\"FFFF0000\"/></sheetPr>");
  EXPECT_EQ(result.value().dropped_record_count, 0U);
}

TEST(XlsbWorksheetProperties, AnAutomaticTabWithNoCodeNameLeavesNoFragment) {
  // An `.xlsx` sheet with nothing to say writes no `<sheetPr>` at all, so
  // neither may the binary path -- otherwise the same workbook answers
  // differently depending on which container it came from.
  const std::vector<std::uint8_t> archive =
      BuildPackage(SheetBinFramed(WsPropPayload(kDefaultWsPropFlags, 0U, 0x40U, 0U, "")));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_TRUE(result.value().workbook.sheet(0).print_settings().sheet_pr_xml.empty());
}

TEST(XlsbWorksheetProperties, TabColorAndCodeNameSurviveAWriteReadCycle) {
  const std::vector<std::uint8_t> archive = BuildPackage(
      SheetBinFramed(WsPropPayload(kDefaultWsPropFlags, /*color_type=*/3U, /*color_index=*/4U, 0U, "AlphaCode")));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  const std::string original = result.value().workbook.sheet(0).print_settings().sheet_pr_xml;
  EXPECT_EQ(original, "<sheetPr codeName=\"AlphaCode\"><tabColor theme=\"4\"/></sheetPr>");

  auto written = io::xlsb::write_xlsb(result.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written)) << written.error().message;
  auto reread = io::xlsb::read_xlsb(test::span_of(written.value()));
  ASSERT_TRUE(static_cast<bool>(reread)) << reread.error().message;
  EXPECT_EQ(reread.value().workbook.sheet(0).print_settings().sheet_pr_xml, original);
}

TEST(XlsbWorksheetProperties, SheetFlagsWithNoModelFieldAreCounted) {
  // `BrtWsProp` also carries dialog-sheet, fit-to-page and outline-
  // direction bits that no model field holds. Decoding the two members
  // that do fit must not make the rest look preserved.
  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFramed(
      WsPropPayload(/*flags=*/0x0000U, /*color_type=*/2U, /*color_index=*/0U, 0xFF00FF00U, "AlphaCode")));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().dropped_record_count, 1U);
  EXPECT_EQ(result.value().workbook.sheet(0).print_settings().sheet_pr_xml,
            "<sheetPr codeName=\"AlphaCode\"><tabColor rgb=\"FF00FF00\"/></sheetPr>");
}

// ---------------------------------------------------------------------------
// Array-formula shells.
// ---------------------------------------------------------------------------

TEST(XlsbArrayFormula, AShellCellIsNotAnUndecodedFormula) {
  // Excel stores every cell of an array footprint as a lone `PtgExp`
  // naming the anchor; the tokens live in the anchor's `BrtArrFmla`. The
  // shell is the expected encoding, so counting it as an undecoded formula
  // would make a healthy workbook indistinguishable from a lossy load.
  std::vector<std::uint8_t> shell;
  shell.push_back(0x01U);  // PtgExp
  AppendU32(shell, 0U);    // owning anchor row

  const std::vector<std::uint8_t> archive = BuildPackage(SheetBinFmlaNum(42.0, shell));
  auto result = io::xlsb::read_xlsb(test::span_of(archive));
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().undecoded_formula_count, 0U);

  // The cached result is still what the cell shows, exactly as on the
  // genuinely-undecodable path.
  const Cell* cell = result.value().workbook.sheet(0).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 42.0);
}

// ---------------------------------------------------------------------------
// XLSB binary parts must not reach an OOXML package.
// ---------------------------------------------------------------------------

/// Collects every entry name of the package `bytes`.
std::vector<std::string> ArchiveEntries(const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  auto opened = zip.open(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(opened)) << (opened ? "" : opened.error().message);
  if (!opened) {
    return {};
  }
  return zip.list_entries();
}

/// Reads the xlsb fixture and re-serialises it through the OOXML writer.
::testing::AssertionResult SaveFixtureAsOoxml(const char* fixture, std::vector<std::uint8_t>* out) {
  const std::vector<std::uint8_t> bytes = test::read_file_bytes(FixturePath(fixture));
  if (bytes.empty()) {
    return ::testing::AssertionFailure() << "fixture bytes empty: " << fixture;
  }
  Workbook wb = Workbook::create_empty();
  if (std::string_view(fixture).find(".xlsb") != std::string_view::npos) {
    auto loaded = io::xlsb::read_xlsb(test::span_of(bytes));
    if (!loaded) {
      return ::testing::AssertionFailure() << "read_xlsb failed: " << loaded.error().message;
    }
    wb = std::move(loaded.value().workbook);
  } else {
    auto loaded = io::read_ooxml(test::span_of(bytes));
    if (!loaded) {
      return ::testing::AssertionFailure() << "read_ooxml failed: " << loaded.error().message;
    }
    wb = std::move(loaded.value().workbook);
  }
  auto saved = io::write_ooxml(wb);
  if (!saved) {
    return ::testing::AssertionFailure() << "write_ooxml failed: " << saved.error().message;
  }
  *out = std::move(saved.value());
  return ::testing::AssertionSuccess();
}

TEST(XlsbToOoxml, PackageCarriesNoBinaryParts) {
  std::vector<std::uint8_t> saved;
  ASSERT_TRUE(SaveFixtureAsOoxml("xlsb_fidelity_base.xlsb", &saved));

  for (const std::string& entry : ArchiveEntries(saved)) {
    EXPECT_EQ(entry.rfind(".bin"), std::string::npos) << "binary part in OOXML output: " << entry;
  }

  std::string content_types;
  ASSERT_TRUE(test::extract_part(test::span_of(saved), "[Content_Types].xml", &content_types));
  EXPECT_EQ(content_types.find("application/vnd.ms-excel."), std::string::npos) << content_types;
  EXPECT_EQ(content_types.find("Extension=\"bin\""), std::string::npos) << content_types;

  std::string workbook_rels;
  ASSERT_TRUE(test::extract_part(test::span_of(saved), "xl/_rels/workbook.xml.rels", &workbook_rels));
  EXPECT_EQ(workbook_rels.find(".bin"), std::string::npos) << workbook_rels;

  for (const std::string& entry : ArchiveEntries(saved)) {
    if (entry.rfind("xl/worksheets/_rels/", 0) != 0) {
      continue;
    }
    std::string sheet_rels;
    ASSERT_TRUE(test::extract_part(test::span_of(saved), entry, &sheet_rels));
    EXPECT_EQ(sheet_rels.find(".bin"), std::string::npos) << entry << ": " << sheet_rels;
  }
}

TEST(XlsbToOoxml, DroppingBinaryPartsIsScopedToTheOoxmlPipeline) {
  // The XLSB writer keeps the source `xl/styles.bin` verbatim rather than
  // regenerating one, because it carries style detail `StylesTable` does
  // not model. That passthrough is exactly what the OOXML pipeline drops,
  // so the two writers must reach their own emission plans: one shared
  // plan builder would either strip the bytes an `.xlsb` save needs or
  // leak them into an `.xlsx`. Pinning both directions from one source
  // workbook keeps that separation observable.
  const std::vector<std::uint8_t> source = test::read_file_bytes(FixturePath("xlsb_fidelity_base.xlsb"));
  ASSERT_FALSE(source.empty());
  auto loaded = io::xlsb::read_xlsb(test::span_of(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;

  std::string source_styles;
  ASSERT_TRUE(test::extract_part(test::span_of(source), "xl/styles.bin", &source_styles));
  ASSERT_FALSE(source_styles.empty());

  auto as_xlsb = io::xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(as_xlsb)) << as_xlsb.error().message;
  std::string saved_styles;
  ASSERT_TRUE(test::extract_part(test::span_of(as_xlsb.value()), "xl/styles.bin", &saved_styles));
  EXPECT_EQ(saved_styles, source_styles) << "an .xlsb save must retain the original styles bytes";

  auto as_ooxml = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(as_ooxml)) << as_ooxml.error().message;
  const std::vector<std::string> ooxml_entries = ArchiveEntries(as_ooxml.value());
  EXPECT_EQ(std::find(ooxml_entries.begin(), ooxml_entries.end(), "xl/styles.bin"), ooxml_entries.end());
  EXPECT_NE(std::find(ooxml_entries.begin(), ooxml_entries.end(), "xl/styles.xml"), ooxml_entries.end());
}

TEST(XlsbToOoxml, EnvelopeIntroducesNothingTheXlsxOriginWouldNotEmit) {
  // The two fixtures are the same workbook exported by Excel to both
  // containers, so the package envelope produced from the binary source
  // may only be a subset of the one produced from the XML source: the
  // XML source keeps passthrough parts (calcChain, metadata) whose
  // binary counterparts are dropped, but the binary source may not
  // contribute a single declaration of its own.
  std::vector<std::uint8_t> from_xlsb;
  std::vector<std::uint8_t> from_xlsx;
  ASSERT_TRUE(SaveFixtureAsOoxml("xlsb_fidelity_base.xlsb", &from_xlsb));
  ASSERT_TRUE(SaveFixtureAsOoxml("xlsb_fidelity_base.xlsx", &from_xlsx));

  const std::vector<std::string> reference = ArchiveEntries(from_xlsx);
  for (const std::string& entry : ArchiveEntries(from_xlsb)) {
    EXPECT_NE(std::find(reference.begin(), reference.end(), entry), reference.end())
        << "part present only when the source was .xlsb: " << entry;
  }

  std::string from_xlsb_types;
  std::string from_xlsx_types;
  ASSERT_TRUE(test::extract_part(test::span_of(from_xlsb), "[Content_Types].xml", &from_xlsb_types));
  ASSERT_TRUE(test::extract_part(test::span_of(from_xlsx), "[Content_Types].xml", &from_xlsx_types));
  std::size_t pos = 0;
  while ((pos = from_xlsb_types.find("  <", pos)) != std::string::npos) {
    const std::size_t end = from_xlsb_types.find('\n', pos);
    if (end == std::string::npos) {
      break;
    }
    const std::string declaration = from_xlsb_types.substr(pos, end - pos);
    EXPECT_NE(from_xlsx_types.find(declaration), std::string::npos)
        << "content-type declaration present only when the source was .xlsb: " << declaration;
    pos = end;
  }
}

}  // namespace
}  // namespace formulon
