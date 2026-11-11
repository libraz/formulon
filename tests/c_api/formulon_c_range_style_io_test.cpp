// Blank formatted cells through a save and a load: style runs are written as
// one `<c>` / `BrtCellBlank` per column, an explicit default only where it
// overrides a row or column format (bare `<c r="..."/>`), and loaded blank
// cells become runs that the blank-slot rule resolves once the layout is in.

#include <cstddef>
#include <cstdint>
#include <fstream>
#include <iterator>
#include <map>
#include <optional>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "utils/a1_ref.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;
using formulon::io::ByteSpan;
using formulon::io::xlsb::XlsbRecordType;

constexpr std::int64_t kNoFormat = -1;

void MakeWorkbook(WorkbookGuard& wb) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  wb.handle->workbook().mutable_styles().cell_xfs.resize(10U);
}

formulon::Sheet& SheetOf(WorkbookGuard& wb) {
  return wb.handle->workbook().sheet(0);
}

// The coordinate's own xf (stored cell or style run), or `kNoFormat`.
std::int64_t OwnXf(WorkbookGuard& wb, std::uint32_t row, std::uint32_t col) {
  const std::optional<formulon::StyleSlot> slot = SheetOf(wb).cell_style_at(row, col);
  return slot.has_value() ? static_cast<std::int64_t>(slot->xf) : kNoFormat;
}

std::size_t CellCount(WorkbookGuard& wb) {
  std::size_t count = 0U;
  EXPECT_EQ(fm_sheet_get_cell_count(wb.handle, 0, &count), 0);
  return count;
}

Bytes Save(WorkbookGuard& wb, fm_workbook_format_t format) {
  BufferGuard buffer;
  EXPECT_EQ(fm_workbook_save(wb.handle, format, &buffer.data, &buffer.len), 0);
  return Bytes(buffer.data, buffer.data + buffer.len);
}

void Load(const Bytes& bytes, WorkbookGuard& out) {
  ASSERT_EQ(fm_workbook_load(bytes.data(), bytes.size(), &out.handle), 0);
}

Bytes ReadFixture(const char* name) {
  std::ifstream in(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name, std::ios::binary);
  EXPECT_TRUE(in.good()) << name;
  return Bytes(std::istreambuf_iterator<char>(in), std::istreambuf_iterator<char>());
}

Bytes Entry(const Bytes& package, const std::string& name) {
  formulon::io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(ByteSpan{package.data(), package.size()})));
  auto body = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(body)) << name;
  return body ? body.value() : Bytes();
}

std::string SheetXml(const Bytes& package) {
  const Bytes body = Entry(package, "xl/worksheets/sheet1.xml");
  return std::string(body.begin(), body.end());
}

// `package` with sheet1.xml replaced by `sheet_xml`.
Bytes WithSheetXml(const Bytes& package, const std::string& sheet_xml) {
  formulon::io::ZipReader input;
  EXPECT_TRUE(static_cast<bool>(input.open(ByteSpan{package.data(), package.size()})));
  mz_zip_archive writer{};
  EXPECT_EQ(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_TRUE);
  for (const std::string& name : input.list_entries()) {
    Bytes content = input.read_entry(name).value();
    if (name == "xl/worksheets/sheet1.xml") {
      content.assign(sheet_xml.begin(), sheet_xml.end());
    }
    EXPECT_EQ(mz_zip_writer_add_mem(&writer, name.c_str(), content.data(), content.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_TRUE);
  }
  void* archive = nullptr;
  std::size_t archive_size = 0;
  EXPECT_EQ(mz_zip_writer_finalize_heap_archive(&writer, &archive, &archive_size), MZ_TRUE);
  EXPECT_EQ(mz_zip_writer_end(&writer), MZ_TRUE);
  Bytes result(static_cast<const std::uint8_t*>(archive), static_cast<const std::uint8_t*>(archive) + archive_size);
  mz_free(archive);
  return result;
}

// Every `<c ...>` opening tag of the sheet XML, in document order.
std::vector<std::string> CellTags(const std::string& xml) {
  std::vector<std::string> tags;
  for (std::size_t at = xml.find("<c "); at != std::string::npos; at = xml.find("<c ", at + 1U)) {
    tags.push_back(xml.substr(at, xml.find('>', at) - at + 1U));
  }
  return tags;
}

// The `r` of each `<c>` in document order.
std::vector<std::string> CellRefs(const std::string& xml) {
  std::vector<std::string> refs;
  for (const std::string& tag : CellTags(xml)) {
    const std::size_t at = tag.find("r=\"") + 3U;
    refs.push_back(tag.substr(at, tag.find('"', at) - at));
  }
  return refs;
}

std::string TagOf(const std::string& xml, const std::string& ref) {
  for (const std::string& tag : CellTags(xml)) {
    if (tag.find("r=\"" + ref + "\"") != std::string::npos) {
      return tag;
    }
  }
  return std::string();
}

std::string Dimension(const std::string& xml) {
  const std::size_t at = xml.find("<dimension ref=\"");
  if (at == std::string::npos) {
    return std::string();
  }
  const std::size_t from = at + 16U;
  return xml.substr(from, xml.find('"', from) - from);
}

std::uint32_t U32(const std::vector<std::uint8_t>& p, std::size_t at) {
  return static_cast<std::uint32_t>(p[at]) | (static_cast<std::uint32_t>(p[at + 1U]) << 8) |
         (static_cast<std::uint32_t>(p[at + 2U]) << 16) | (static_cast<std::uint32_t>(p[at + 3U]) << 24);
}

// One cell record of the sheet stream.
struct BinCell {
  std::uint32_t row;
  std::uint32_t col;
  std::uint16_t type;
  std::uint32_t xf;
};

// The decoded parts of sheet1.bin these tests look at.
struct SheetBin {
  std::vector<BinCell> cells;
  std::map<std::uint32_t, std::vector<std::pair<std::uint32_t, std::uint32_t>>> spans;  // per row
  std::map<std::uint32_t, std::pair<std::uint32_t, bool>> headers;                      // per row: ixfe, fGhostDirty
  std::optional<std::vector<std::uint32_t>> dim;  // rwFirst, rwLast, colFirst, colLast
};

SheetBin DecodeSheetBin(const Bytes& package) {
  const Bytes body = Entry(package, "xl/worksheets/sheet1.bin");
  SheetBin out;
  ByteSpan cursor{body.data(), body.size()};
  std::uint32_t row = 0;
  bool in_data = false;
  while (cursor.size != 0U) {
    auto rec = formulon::io::xlsb::read_record(cursor);
    if (!rec) {
      ADD_FAILURE() << rec.error().message;
      break;
    }
    const std::uint16_t type = rec.value().type;
    const std::vector<std::uint8_t> p(rec.value().payload.data, rec.value().payload.data + rec.value().payload.size);
    if (type == static_cast<std::uint16_t>(XlsbRecordType::BrtBeginSheetData)) {
      in_data = true;
    } else if (type == static_cast<std::uint16_t>(XlsbRecordType::BrtEndSheetData)) {
      in_data = false;
    } else if (type == static_cast<std::uint16_t>(XlsbRecordType::BrtWsDim)) {
      out.dim = std::vector<std::uint32_t>{U32(p, 0), U32(p, 4), U32(p, 8), U32(p, 12)};
    } else if (in_data && type == static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      row = U32(p, 0);
      out.headers[row] = {U32(p, 4), (p[11] & 0x40U) != 0U};
      const std::uint32_t count = U32(p, 13);
      auto& spans = out.spans[row];
      for (std::uint32_t i = 0; i < count; ++i) {
        spans.emplace_back(U32(p, 17U + 8U * i), U32(p, 21U + 8U * i));
      }
    } else if (in_data && type <= static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaError) && p.size() >= 8U) {
      out.cells.push_back(BinCell{row, U32(p, 0), type, U32(p, 4) & 0xFFFFFFU});
    }
  }
  return out;
}

std::set<std::string> BlankRefs(const SheetBin& bin, std::uint32_t xf) {
  std::set<std::string> refs;
  for (const BinCell& cell : bin.cells) {
    if (cell.type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank) && cell.xf == xf) {
      refs.insert(formulon::a1::encode_a1(cell.row, cell.col));
    }
  }
  return refs;
}

fm_status_t SetLink(WorkbookGuard& wb, fm_cell_rect_t rect) {
  fm_hyperlink_t link = FM_STRUCT_INIT(fm_hyperlink_t);
  link.rect = rect;
  link.target = "https://example.com";
  return fm_sheet_set_hyperlink(wb.handle, 0, &link);
}

// A package whose sheet1 is `sheet_data` (the `<sheetData>` children) under
// `cols` (the `<cols>` element, or empty), with xf 1 in its styles table.
Bytes PackageWith(const std::string& cols, const std::string& sheet_data) {
  WorkbookGuard base;
  MakeWorkbook(base);
  const std::string xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">" +
      cols + "<sheetData>" + sheet_data + "</sheetData></worksheet>";
  return WithSheetXml(Save(base, FM_WORKBOOK_FORMAT_XLSX), xml);
}

}  // namespace

// Excel's own load of the authored file: bare and `s="0"` blanks
// under a column or row format are explicit defaults, the ones with no axis
// format are dropped, and the save writes the same `<c>` set Excel saved.
TEST(FormulonCApiRangeStyleIo, ExcelBlankSlotFileLoadsAndResavesExcelsCellSet) {
  WorkbookGuard loaded;
  Load(ReadFixture("blankslot_authored.xlsx"), loaded);
  EXPECT_EQ(OwnXf(loaded, 4U, 1U), 0);          // B5
  EXPECT_EQ(OwnXf(loaded, 5U, 1U), 0);          // B6
  EXPECT_EQ(OwnXf(loaded, 9U, 3U), 0);          // D10
  EXPECT_EQ(OwnXf(loaded, 9U, 4U), 0);          // E10
  EXPECT_EQ(OwnXf(loaded, 4U, 7U), kNoFormat);  // H5
  EXPECT_EQ(OwnXf(loaded, 5U, 7U), kNoFormat);  // H6
  EXPECT_EQ(OwnXf(loaded, 6U, 1U), kNoFormat);  // B7 inherits
  EXPECT_EQ(CellCount(loaded), 4U);

  const std::string excel_saved = SheetXml(ReadFixture("blankslot_saved.xlsx"));
  const std::string ours = SheetXml(Save(loaded, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_EQ(CellTags(ours), CellTags(excel_saved));
  EXPECT_EQ(Dimension(ours), Dimension(excel_saved));

  // Excel's saved file loads to the same slots and re-saves byte-equal for its `<c>` set.
  WorkbookGuard resaved;
  Load(ReadFixture("blankslot_saved.xlsx"), resaved);
  EXPECT_EQ(CellTags(SheetXml(Save(resaved, FM_WORKBOOK_FORMAT_XLSX))), CellTags(excel_saved));
}

// The same file as Excel saved it to .xlsb: the xf-0 `BrtCellBlank`s block the
// axis format, and the save writes them back.
TEST(FormulonCApiRangeStyleIo, ExcelBlankSlotXlsbLoadsAndResavesItsBlankRecords) {
  WorkbookGuard loaded;
  Load(ReadFixture("blankslot_saved.xlsb"), loaded);
  EXPECT_EQ(OwnXf(loaded, 4U, 1U), 0);  // B5
  EXPECT_EQ(OwnXf(loaded, 9U, 4U), 0);  // E10
  EXPECT_EQ(OwnXf(loaded, 6U, 1U), kNoFormat);
  EXPECT_EQ(CellCount(loaded), 4U);

  const SheetBin bin = DecodeSheetBin(Save(loaded, FM_WORKBOOK_FORMAT_XLSB));
  EXPECT_EQ(BlankRefs(bin, 0U), (std::set<std::string>{"B5", "B6", "D10", "E10"}));
  ASSERT_TRUE(bin.dim.has_value());
  EXPECT_EQ(*bin.dim, (std::vector<std::uint32_t>{4U, 9U, 1U, 4U}));
}

// A run of xf-0 blanks crossing a column-format boundary keeps only the
// covered columns, as Excel's load does; those cover every styled column,
// three of them, so the save writes the row as an xf-0 row format with no cells.
TEST(FormulonCApiRangeStyleIo, PartlyCoveredExplicitDefaultsKeepOnlyTheCoveredColumns) {
  WorkbookGuard loaded;
  Load(ReadFixture("blankslot_partial_authored.xlsx"), loaded);
  EXPECT_EQ(OwnXf(loaded, 2U, 1U), kNoFormat);   // B3
  EXPECT_EQ(OwnXf(loaded, 2U, 3U), 0);           // D3
  EXPECT_EQ(OwnXf(loaded, 2U, 10U), kNoFormat);  // K3

  const std::string xml = SheetXml(Save(loaded, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_TRUE(CellTags(xml).empty()) << xml;
  EXPECT_NE(xml.find("<row r=\"3\" customFormat=\"1\"/>"), std::string::npos) << xml;
  EXPECT_EQ(Dimension(xml), "D3:F3");
  const SheetBin bin = DecodeSheetBin(Save(loaded, FM_WORKBOOK_FORMAT_XLSB));
  EXPECT_TRUE(BlankRefs(bin, 0U).empty());
  ASSERT_EQ(bin.headers.count(2U), 1U);
  EXPECT_EQ(bin.headers.at(2U), std::make_pair(0U, true));
}

// A row format of xf 0 shadows the column formats under it, so a bare blank
// in that row overrides nothing and is dropped, as Excel drops it.
TEST(FormulonCApiRangeStyleIo, ExplicitDefaultUnderARowDefaultFormatIsDropped) {
  WorkbookGuard loaded;
  Load(ReadFixture("blankslot_row_default_authored.xlsx"), loaded);
  EXPECT_EQ(OwnXf(loaded, 2U, 2U), kNoFormat);  // C3
  const std::string xml = SheetXml(Save(loaded, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_TRUE(CellTags(xml).empty()) << xml;
  // With no cell left, the dimension is still the row x column format intersection.
  EXPECT_EQ(Dimension(xml), "C3:D3");
  const SheetBin bin = DecodeSheetBin(Save(loaded, FM_WORKBOOK_FORMAT_XLSB));
  EXPECT_TRUE(bin.cells.empty());
  ASSERT_TRUE(bin.dim.has_value());
  EXPECT_EQ(*bin.dim, (std::vector<std::uint32_t>{2U, 2U, 2U, 3U}));
}

// The writer's own rule when the axis format goes away after the write: the
// explicit default stays in memory but is not saved (in either format).
TEST(FormulonCApiRangeStyleIo, ExplicitDefaultWhoseAxisFormatWasRemovedIsNotSaved) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 10U, 1U), 0);  // B:K
  const fm_cell_rect_t row3{2U, 1U, 2U, 10U};                             // B3:K3
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, &row3, 0U), 0);
  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 1U, 2U), 0);   // B:C
  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 6U, 10U), 0);  // G:K
  EXPECT_EQ(OwnXf(wb, 2U, 1U), 0);                                      // kept in memory
  EXPECT_EQ(OwnXf(wb, 2U, 10U), 0);

  const std::string xml = SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_EQ(CellRefs(xml), (std::vector<std::string>{"D3", "E3", "F3"}));
  EXPECT_EQ(TagOf(xml, "D3"), "<c r=\"D3\"/>");
  EXPECT_EQ(Dimension(xml), "D3:F3");
  const SheetBin bin = DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB));
  EXPECT_EQ(BlankRefs(bin, 0U), (std::set<std::string>{"D3", "E3", "F3"}));

  ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 3U, 5U), 0);  // D:F
  EXPECT_TRUE(CellTags(SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX))).empty());
  EXPECT_TRUE(BlankRefs(DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB)), 0U).empty());

  WorkbookGuard reloaded;
  Load(Save(wb, FM_WORKBOOK_FORMAT_XLSX), reloaded);
  EXPECT_EQ(CellCount(reloaded), 0U);
}

// Success criterion 4: removing a link resets its blanks; with no axis format
// they leave the file, under a column or row format they are bare `<c r/>`.
TEST(FormulonCApiRangeStyleIo, HyperlinkRemoveLeavesBareCellsOnlyUnderAnAxisFormat) {
  const fm_cell_rect_t rect{0U, 0U, 2U, 2U};  // A1:C3
  struct Case {
    const char* name;
    int axis;  // 0 none, 1 column B, 2 row 2
    std::vector<std::string> refs;
  };
  const std::vector<Case> cases = {
      {"no axis", 0, {"A1"}},
      {"column B", 1, {"A1", "B1", "B2", "B3"}},
      {"row 2", 2, {"A1", "A2", "B2", "C2"}},
  };
  for (const Case& c : cases) {
    SCOPED_TRACE(c.name);
    WorkbookGuard wb;
    MakeWorkbook(wb);
    if (c.axis == 1) {
      ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 1U, 1U), 0);
    } else if (c.axis == 2) {
      ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 1U, 1U, 1U), 0);
    }
    ASSERT_EQ(SetLink(wb, rect), 0);
    ASSERT_EQ(fm_sheet_remove_hyperlink(wb.handle, 0, &rect), 0);
    EXPECT_EQ(CellCount(wb), c.refs.size());

    const std::string xml = SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX));
    EXPECT_EQ(CellRefs(xml), c.refs);
    for (std::size_t i = 1; i < c.refs.size(); ++i) {
      EXPECT_EQ(TagOf(xml, c.refs[i]), "<c r=\"" + c.refs[i] + "\"/>");
    }
    EXPECT_NE(TagOf(xml, "A1").find("t=\""), std::string::npos);  // the top-left keeps its text
  }
}

// W1: one record per column in column order, a phantom winning over the run
// column it covers and taking the run's format; the remaining run columns are
// blank cells of the run's format.
TEST(FormulonCApiRangeStyleIo, RowMixingStoredRunAndPhantomColumnsWritesEachColumnOnce) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 0U, "=SEQUENCE(1,3)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const fm_cell_rect_t band{0U, 1U, 0U, 4U};  // B1:E1
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, &band, 1U), 0);

  const std::string xml = SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_EQ(CellRefs(xml), (std::vector<std::string>{"A1", "B1", "C1", "D1", "E1"}));
  EXPECT_NE(TagOf(xml, "B1").find("s=\"1\""), std::string::npos);
  EXPECT_EQ(TagOf(xml, "D1"), "<c r=\"D1\" s=\"1\"/>");
  EXPECT_EQ(Dimension(xml), "A1:E1");

  const SheetBin bin = DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB));
  std::vector<std::uint32_t> cols;
  for (const BinCell& cell : bin.cells) {
    cols.push_back(cell.col);
  }
  EXPECT_EQ(cols, (std::vector<std::uint32_t>{0U, 1U, 2U, 3U, 4U}));
  ASSERT_EQ(bin.cells.size(), 5U);
  EXPECT_NE(bin.cells[1].type, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank));
  EXPECT_EQ(bin.cells[1].xf, 1U);
  EXPECT_EQ(BlankRefs(bin, 1U), (std::set<std::string>{"D1", "E1"}));
}

// A phantom over an explicit-default run under a styled column keeps `s="0"`
// (it reloads as a value cell, where a missing `s` would inherit).
TEST(FormulonCApiRangeStyleIo, PhantomOverAnExplicitDefaultRunKeepsStyleZero) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, 1U, 1U, 1U), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 0U, 1U, 0U), 0);
  ASSERT_EQ(fm_sheet_set_formula(wb.handle, 0, 0U, 0U, "=SEQUENCE(1,2)"), 0);
  ASSERT_EQ(fm_workbook_recalc(wb.handle), 0);
  const std::string xml = SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX));
  EXPECT_NE(TagOf(xml, "B1").find("s=\"0\""), std::string::npos) << TagOf(xml, "B1");
}

// B-4: a run crossing 1,024-column segments 0, 1 and 2 is covered by one
// BrtRowHdr span per segment, clipped to the segment; a run-only row is written.
TEST(FormulonCApiRangeStyleIo, XlsbRowSpansCoverARunInEverySegmentItCrosses) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  const fm_cell_rect_t band{4U, 1000U, 4U, 2100U};
  ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, &band, 2U), 0);
  const SheetBin bin = DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB));
  ASSERT_EQ(bin.spans.count(4U), 1U);
  EXPECT_EQ(bin.spans.at(4U),
            (std::vector<std::pair<std::uint32_t, std::uint32_t>>{{1000U, 1023U}, {1024U, 2047U}, {2048U, 2100U}}));
  EXPECT_EQ(bin.cells.size(), 1101U);
  ASSERT_TRUE(bin.dim.has_value());
  EXPECT_EQ(*bin.dim, (std::vector<std::uint32_t>{4U, 4U, 1000U, 2100U}));

  WorkbookGuard reloaded;
  Load(Save(wb, FM_WORKBOOK_FORMAT_XLSB), reloaded);
  EXPECT_EQ(CellCount(reloaded), 1101U);
  EXPECT_EQ(OwnXf(reloaded, 4U, 2100U), 2);
  EXPECT_EQ(SheetOf(reloaded).cell_at(4U, 1500U), nullptr);
}

// W4: a run-only extremity sets the dimension and BrtWsDim.
TEST(FormulonCApiRangeStyleIo, DimensionIncludesARunOnlyExtremity) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0U, 0U, "a"), 0);
  ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, 9U, 7U, 1U), 0);  // H10
  EXPECT_EQ(Dimension(SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX))), "A1:H10");
  const SheetBin bin = DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB));
  ASSERT_TRUE(bin.dim.has_value());
  EXPECT_EQ(*bin.dim, (std::vector<std::uint32_t>{0U, 9U, 0U, 7U}));
}

// On both read paths, bare and `s="0"` blanks are explicit defaults under a
// column format and under a row format, nothing without one; a styled blank is
// a run of its xf. The DOM path applies row formats after the cells, SAX before.
TEST(FormulonCApiRangeStyleIo, LoadedBlanksResolveTheSameOnTheDomAndSaxPaths) {
  const Bytes package = PackageWith("<cols><col min=\"2\" max=\"2\" style=\"1\"/></cols>",
                                    "<row r=\"1\"><c r=\"A1\"/><c r=\"B1\"/><c r=\"C1\" s=\"2\"/></row>"
                                    "<row r=\"2\"><c r=\"B2\" s=\"0\"/><c r=\"D2\" s=\"0\"/></row>"
                                    "<row r=\"3\" s=\"1\" customFormat=\"1\"><c r=\"D3\"/><c r=\"E3\" s=\"0\"/></row>");
  for (const std::size_t threshold : {static_cast<std::size_t>(-1), std::size_t{0}}) {
    SCOPED_TRACE(threshold == 0U ? "sax" : "dom");
    auto read = formulon::io::internal::ReadOoxmlWithSaxThresholdForTesting(ByteSpan{package.data(), package.size()},
                                                                            threshold);
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    const formulon::Sheet& sheet = read.value().workbook.sheet(0);
    const auto own = [&](std::uint32_t row, std::uint32_t col) {
      const std::optional<formulon::StyleSlot> slot = sheet.cell_style_at(row, col);
      return slot.has_value() ? static_cast<std::int64_t>(slot->xf) : kNoFormat;
    };
    EXPECT_EQ(own(0U, 0U), kNoFormat);  // A1: no axis
    EXPECT_EQ(own(0U, 1U), 0);          // B1: under the column
    EXPECT_EQ(own(0U, 2U), 2);          // C1: its own format
    EXPECT_EQ(own(1U, 1U), 0);          // B2
    EXPECT_EQ(own(1U, 3U), kNoFormat);  // D2
    EXPECT_EQ(own(2U, 3U), 0);          // D3: under the row
    EXPECT_EQ(own(2U, 4U), 0);          // E3
    EXPECT_EQ(sheet.cell_count(), 5U);
    for (const auto& [row, cells] : sheet.rows()) {
      EXPECT_TRUE(cells.has_any()) << "empty row node " << row;
    }
  }
}

// The same rule on the production SAX path, over a sheet part past the
// 256 KiB threshold.
TEST(FormulonCApiRangeStyleIo, LargeSheetLoadsBlanksAsRunsThroughTheStreamingReader) {
  std::string rows = "<row r=\"1\"><c r=\"A1\"/><c r=\"B1\"/><c r=\"C1\" s=\"2\"/></row>";
  for (std::uint32_t r = 2; r <= 20000U; ++r) {
    rows += "<row r=\"" + std::to_string(r) + "\"><c r=\"A" + std::to_string(r) + "\"><v>1</v></c></row>";
  }
  const Bytes package = PackageWith("<cols><col min=\"2\" max=\"2\" style=\"1\"/></cols>", rows);
  ASSERT_GT(SheetXml(package).size(), 256U * 1024U);
  WorkbookGuard loaded;
  Load(package, loaded);
  EXPECT_EQ(OwnXf(loaded, 0U, 0U), kNoFormat);
  EXPECT_EQ(OwnXf(loaded, 0U, 1U), 0);
  EXPECT_EQ(OwnXf(loaded, 0U, 2U), 2);
  EXPECT_EQ(SheetOf(loaded).cell_at(0U, 2U), nullptr);
  EXPECT_EQ(CellCount(loaded), 2U + 19999U);
}

// Overlapping `<cols>` from a non-Excel producer: spans are sorted and the
// span that starts first keeps the shared columns.
TEST(FormulonCApiRangeStyleIo, OverlappingColumnSpansLoadSortedWithTheFirstStartingSpanWinning) {
  const Bytes package =
      PackageWith("<cols><col min=\"3\" max=\"5\" style=\"1\"/><col min=\"1\" max=\"4\" style=\"2\"/></cols>",
                  "<row r=\"1\"><c r=\"D1\"/><c r=\"E1\"/></row>");
  WorkbookGuard loaded;
  Load(package, loaded);
  const std::vector<formulon::ColumnLayout>& columns = SheetOf(loaded).layout().columns;
  ASSERT_EQ(columns.size(), 2U);
  EXPECT_EQ(columns[0].first, 0U);
  EXPECT_EQ(columns[0].last, 3U);
  EXPECT_EQ(columns[0].style_xf, 2U);
  EXPECT_EQ(columns[1].first, 4U);
  EXPECT_EQ(columns[1].last, 4U);
  EXPECT_EQ(columns[1].style_xf, 1U);
  EXPECT_EQ(OwnXf(loaded, 0U, 3U), 0);
  EXPECT_EQ(OwnXf(loaded, 0U, 4U), 0);
}

// A phonetic blank holding an explicit default is content, written bare with
// `ph`; it reloads as a stored cell of xf 0 without the explicit mark, which
// shows the same (a stored cell owns its style).
TEST(FormulonCApiRangeStyleIo, PhoneticBlankWithAnExplicitDefaultReloadsAsAStoredDefaultCell) {
  WorkbookGuard wb;
  MakeWorkbook(wb);
  SheetOf(wb).set_cell_show_phonetic(0U, 0U, true);
  SheetOf(wb).set_cell_xf_index(0U, 0U, 0U);
  const Bytes saved = Save(wb, FM_WORKBOOK_FORMAT_XLSX);
  EXPECT_EQ(TagOf(SheetXml(saved), "A1"), "<c r=\"A1\" ph=\"1\"/>");

  WorkbookGuard reloaded;
  Load(saved, reloaded);
  const std::optional<formulon::StyleSlot> slot = SheetOf(reloaded).cell_style_at(0U, 0U);
  ASSERT_TRUE(slot.has_value());
  EXPECT_EQ(slot->xf, 0U);
  EXPECT_FALSE(slot->explicit_xf);
  ASSERT_NE(SheetOf(reloaded).cell_at(0U, 0U), nullptr);
  EXPECT_TRUE(SheetOf(reloaded).cell_at(0U, 0U)->show_phonetic);
}

namespace {

// One sheet of styled rows and styled columns, and the dimension Excel writes
// for it.
struct AxisDimensionCase {
  const char* name;
  std::vector<std::vector<std::uint32_t>> columns;  // {first, last, xf}
  std::vector<std::vector<std::uint32_t>> rows;     // {row, xf}
  bool text_in_a1;
  const char* dimension;                             // .xlsx `<dimension ref>`
  std::optional<std::vector<std::uint32_t>> ws_dim;  // BrtWsDim rwFirst, rwLast, colFirst, colLast
};

class FormulonCApiAxisDimension : public ::testing::TestWithParam<AxisDimensionCase> {};

}  // namespace

// The dimension also spans every intersection of a customFormat row (any xf,
// 0 included) with a column of non-zero style, with no cell there.
TEST_P(FormulonCApiAxisDimension, CoversStyledRowAndColumnIntersections) {
  const AxisDimensionCase& c = GetParam();
  WorkbookGuard wb;
  MakeWorkbook(wb);
  for (const std::vector<std::uint32_t>& col : c.columns) {
    ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, col[0], col[1], col[2]), 0);
  }
  for (const std::vector<std::uint32_t>& row : c.rows) {
    ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, row[0], row[0], row[1]), 0);
  }
  if (c.text_in_a1) {
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0U, 0U, "a"), 0);
  }
  EXPECT_EQ(Dimension(SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX))), c.dimension);
  EXPECT_EQ(DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB)).dim, c.ws_dim);
}

INSTANTIATE_TEST_SUITE_P(
    Cases, FormulonCApiAxisDimension,
    ::testing::Values(
        AxisDimensionCase{"RowFormatOverStyledColumns",
                          {{2, 3, 1}},
                          {{2, 2}},
                          false,
                          "C3:D3",
                          std::vector<std::uint32_t>{2, 2, 2, 3}},
        AxisDimensionCase{"RowFormatEqualToColumnFormat",
                          {{2, 3, 1}},
                          {{2, 1}},
                          false,
                          "C3:D3",
                          std::vector<std::uint32_t>{2, 2, 2, 3}},
        AxisDimensionCase{"DefaultRowFarBelowTextInA1",
                          {{7, 7, 1}},
                          {{19, 0}},
                          true,
                          "A1:H20",
                          std::vector<std::uint32_t>{0, 19, 0, 7}},
        AxisDimensionCase{
            "TwoDefaultRows", {{2, 3, 1}}, {{2, 0}, {6, 0}}, false, "C3:D7", std::vector<std::uint32_t>{2, 6, 2, 3}},
        AxisDimensionCase{"RowAcrossTwoColumnGroups",
                          {{2, 3, 1}, {10, 10, 2}},
                          {{2, 0}},
                          false,
                          "C3:K3",
                          std::vector<std::uint32_t>{2, 2, 2, 10}},
        AxisDimensionCase{"DefaultRowOverStyledColumns",
                          {{2, 3, 1}},
                          {{2, 0}},
                          false,
                          "C3:D3",
                          std::vector<std::uint32_t>{2, 2, 2, 3}},
        // Styled columns with no written row span row 1.
        AxisDimensionCase{"ColumnsOnly", {{2, 3, 1}}, {}, false, "C1:D1", std::vector<std::uint32_t>{0, 0, 2, 3}},
        // A written row with no styled column spans column A.
        AxisDimensionCase{"RowOnly", {}, {{2, 1}}, false, "A3", std::vector<std::uint32_t>{2, 2, 0, 0}},
        // A column style widens the column span only.
        AxisDimensionCase{
            "ColumnStyleBesideTextInA1", {{7, 7, 1}}, {}, true, "A1:H1", std::vector<std::uint32_t>{0, 0, 0, 7}},
        AxisDimensionCase{"Empty", {}, {}, false, "A1", std::vector<std::uint32_t>{0, 0, 0, 0}}),
    [](const ::testing::TestParamInfo<AxisDimensionCase>& param_info) { return std::string(param_info.param.name); });

namespace {

// Columns used below: D=3, E=4, F=5; row 3 is index 2.
using Setup = void (*)(WorkbookGuard&);

void StyleColumns(WorkbookGuard& wb, std::uint32_t first, std::uint32_t last, std::uint32_t xf = 1U) {
  ASSERT_EQ(fm_sheet_set_column_xf_index(wb.handle, 0, first, last, xf), 0);
}

void Blanks(WorkbookGuard& wb, std::uint32_t row, std::uint32_t first, std::uint32_t last, std::uint32_t xf = 0U) {
  for (std::uint32_t col = first; col <= last; ++col) {
    ASSERT_EQ(fm_sheet_set_cell_xf_index(wb.handle, 0, row, col, xf), 0);
  }
}

// One emit-plan case: the cells and the row-3 element Excel writes in each
// format, and the dimension.
struct EmitPlanCase {
  const char* name;
  Setup setup;
  std::vector<std::string> xlsx_cells;                      // `<c ...>` opening tags in order
  std::string xlsx_row3;                                    // row 3's `<row>` tag, or "" for none
  std::set<std::string> xlsb_blanks;                        // BrtCellBlank refs, any xf
  std::optional<std::pair<std::uint32_t, bool>> xlsb_row3;  // BrtRowHdr ixfe, fGhostDirty
  const char* dimension;
  // False where the save drops state that decided the plan, so the reload
  // plans differently (an explicit default left outside any axis format).
  bool stable = true;
};

class FormulonCApiRowEmitPlan : public ::testing::TestWithParam<EmitPlanCase> {};

std::string RowTag(const std::string& xml, std::uint32_t row1) {
  const std::string open = "<row r=\"" + std::to_string(row1) + "\"";
  const std::size_t at = xml.find(open);
  return at == std::string::npos ? std::string() : xml.substr(at, xml.find('>', at) - at + 1U);
}

std::set<std::string> AllBlankRefs(const SheetBin& bin) {
  std::set<std::string> refs;
  for (const BinCell& cell : bin.cells) {
    if (cell.type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank)) {
      refs.insert(formulon::a1::encode_a1(cell.row, cell.col));
    }
  }
  return refs;
}

}  // namespace

// Step 1 converts a row of explicit defaults covering every styled column (3
// or more) into an xf-0 row format; step 2 omits a run piece whose xf equals
// the xf it would inherit. Both formats, and a save -> load -> save is stable.
TEST_P(FormulonCApiRowEmitPlan, WritesTheRowAsExcelDoes) {
  const EmitPlanCase& c = GetParam();
  WorkbookGuard wb;
  MakeWorkbook(wb);
  c.setup(wb);

  const Bytes xlsx = Save(wb, FM_WORKBOOK_FORMAT_XLSX);
  const std::string xml = SheetXml(xlsx);
  EXPECT_EQ(CellTags(xml), c.xlsx_cells) << xml;
  EXPECT_EQ(RowTag(xml, 3U), c.xlsx_row3) << xml;
  EXPECT_EQ(Dimension(xml), c.dimension);

  const Bytes xlsb = Save(wb, FM_WORKBOOK_FORMAT_XLSB);
  const SheetBin bin = DecodeSheetBin(xlsb);
  EXPECT_EQ(AllBlankRefs(bin), c.xlsb_blanks);
  const auto header = bin.headers.find(2U);
  ASSERT_EQ(header != bin.headers.end(), c.xlsb_row3.has_value());
  if (c.xlsb_row3.has_value()) {
    EXPECT_EQ(header->second, *c.xlsb_row3);
  }

  for (const fm_workbook_format_t format : {FM_WORKBOOK_FORMAT_XLSX, FM_WORKBOOK_FORMAT_XLSB}) {
    if (!c.stable) {
      break;
    }
    const Bytes first = format == FM_WORKBOOK_FORMAT_XLSX ? xlsx : xlsb;
    WorkbookGuard reloaded;
    Load(first, reloaded);
    const Bytes second = Save(reloaded, format);
    const std::string part =
        format == FM_WORKBOOK_FORMAT_XLSX ? "xl/worksheets/sheet1.xml" : "xl/worksheets/sheet1.bin";
    EXPECT_EQ(Entry(second, part), Entry(first, part)) << (format == FM_WORKBOOK_FORMAT_XLSX ? "xlsx" : "xlsb");
  }
}

namespace {

const std::pair<std::uint32_t, bool> kPlainRow{0U, false};
const std::pair<std::uint32_t, bool> kConvertedRow{0U, true};
const char* const kConvertedTag = "<row r=\"3\" customFormat=\"1\"/>";
const char* const kPlainTag = "<row r=\"3\">";

}  // namespace

INSTANTIATE_TEST_SUITE_P(
    Cases, FormulonCApiRowEmitPlan,
    ::testing::Values(EmitPlanCase{"OneOfThreeStyledColumnsCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 3);
                                   },
                                   {"<c r=\"D3\"/>"},
                                   kPlainTag,
                                   {"D3"},
                                   kPlainRow,
                                   "D3:F3"},
                      EmitPlanCase{"TwoOfThreeStyledColumnsCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 4);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3"},
                                   kPlainRow,
                                   "D3:F3"},
                      EmitPlanCase{"TwoStyledColumnsCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 4);
                                     Blanks(wb, 2, 3, 4);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3"},
                                   kPlainRow,
                                   "D3:E3"},
                      EmitPlanCase{"OneStyledColumnCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 3);
                                     Blanks(wb, 2, 3, 3);
                                   },
                                   {"<c r=\"D3\"/>"},
                                   kPlainTag,
                                   {"D3"},
                                   kPlainRow,
                                   "D3"},
                      EmitPlanCase{"ValueInTheRowBlocksConversion",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5);
                                     ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2U, 0U, "x"), 0);
                                   },
                                   {"<c r=\"A3\" t=\"s\">", "<c r=\"D3\"/>", "<c r=\"E3\"/>", "<c r=\"F3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "F3"},
                                   kPlainRow,
                                   "A3:F3"},
                      // An explicit default left behind when its column format was removed
                      // still blocks the conversion, and is itself not written.
                      EmitPlanCase{"LeftoverExplicitDefaultBlocksConversion",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 1, 1);
                                     Blanks(wb, 2, 1, 1);
                                     ASSERT_EQ(fm_sheet_reset_column_xf_index(wb.handle, 0, 1U, 1U), 0);
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>", "<c r=\"F3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "F3"},
                                   kPlainRow,
                                   "D3:F3",
                                   false},
                      EmitPlanCase{"MixedColumnFormatsAllCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 3, 1);
                                     StyleColumns(wb, 4, 4, 2);
                                     StyleColumns(wb, 5, 5, 1);
                                     Blanks(wb, 2, 3, 5);
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:F3"},
                      EmitPlanCase{"NonDefaultBlanksAreWritten",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5, 2);
                                   },
                                   {"<c r=\"D3\" s=\"2\"/>", "<c r=\"E3\" s=\"2\"/>", "<c r=\"F3\" s=\"2\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "F3"},
                                   kPlainRow,
                                   "D3:F3"},
                      EmitPlanCase{"NonDefaultAndDefaultBlanks",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 4, 2);
                                     Blanks(wb, 2, 5, 5);
                                   },
                                   {"<c r=\"D3\" s=\"2\"/>", "<c r=\"E3\" s=\"2\"/>", "<c r=\"F3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "F3"},
                                   kPlainRow,
                                   "D3:F3"},
                      EmitPlanCase{"TwoCoveredRows",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5);
                                     Blanks(wb, 3, 3, 5);
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:F4"},
                      // A styled span to the last column is only partly covered.
                      EmitPlanCase{"StyledSpanToLastColumnPartlyCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 16383);
                                     Blanks(wb, 2, 3, 5);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>", "<c r=\"F3\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "F3"},
                                   kPlainRow,
                                   "D3:XFD3"},
                      EmitPlanCase{"ThreeStyledColumnsCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5);
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:F3"},
                      EmitPlanCase{"FourStyledColumnsCovered",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 6);
                                     Blanks(wb, 2, 3, 6);
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:G3"},
                      EmitPlanCase{"TwoSeparateStyledColumns",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 3);
                                     StyleColumns(wb, 5, 5);
                                     Blanks(wb, 2, 3, 3);
                                     Blanks(wb, 2, 5, 5);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"F3\"/>"},
                                   kPlainTag,
                                   {"D3", "F3"},
                                   kPlainRow,
                                   "D3:F3"},
                      EmitPlanCase{"ThreeSeparateStyledColumns",
                                   [](WorkbookGuard& wb) {
                                     for (const std::uint32_t col : {3U, 5U, 7U}) {
                                       StyleColumns(wb, col, col);
                                       Blanks(wb, 2, col, col);
                                     }
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:H3"},
                      EmitPlanCase{"TwoStyledColumnsTwoRows",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 4);
                                     Blanks(wb, 2, 3, 4);
                                     Blanks(wb, 3, 3, 4);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>", "<c r=\"D4\"/>", "<c r=\"E4\"/>"},
                                   kPlainTag,
                                   {"D3", "E3", "D4", "E4"},
                                   kPlainRow,
                                   "D3:E4"},
                      EmitPlanCase{"RangeDefaultOverStyledColumns",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     const fm_cell_rect_t rect{2U, 3U, 2U, 5U};
                                     ASSERT_EQ(fm_sheet_set_range_xf_index(wb.handle, 0, &rect, 0U), 0);
                                   },
                                   {},
                                   kConvertedTag,
                                   {},
                                   kConvertedRow,
                                   "D3:F3"},
                      // Blanks of the columns' own xf are redundant and are not written; a
                      // row left with nothing is not written either, and the styled
                      // columns span row 1.
                      EmitPlanCase{"BlanksOfTheColumnFormatAreRedundant",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5, 1);
                                   },
                                   {},
                                   "",
                                   {},
                                   std::nullopt,
                                   "D1:F1"},
                      EmitPlanCase{"BlanksUnderANonDefaultRowFormat",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 2U, 2U, 2U), 0);
                                     Blanks(wb, 2, 3, 5);
                                   },
                                   {"<c r=\"D3\"/>", "<c r=\"E3\"/>", "<c r=\"F3\"/>"},
                                   "<row r=\"3\" s=\"2\" customFormat=\"1\">",
                                   {"D3", "E3", "F3"},
                                   std::make_pair(2U, true),
                                   "D3:F3"},
                      EmitPlanCase{"BlanksUnderADefaultRowFormat",
                                   [](WorkbookGuard& wb) {
                                     StyleColumns(wb, 3, 5);
                                     Blanks(wb, 2, 3, 5);
                                     ASSERT_EQ(fm_sheet_set_row_xf_index(wb.handle, 0, 2U, 2U, 0U), 0);
                                   },
                                   {},
                                   "<row r=\"3\" customFormat=\"1\"/>",
                                   {},
                                   kConvertedRow,
                                   "D3:F3"}),
    [](const ::testing::TestParamInfo<EmitPlanCase>& param_info) { return std::string(param_info.param.name); });

namespace {

// One dimension case, with the dimension Excel writes.
struct DimensionCase {
  const char* name;
  Setup setup;
  const char* dimension;
  std::optional<std::vector<std::uint32_t>> ws_dim;
};

class FormulonCApiSheetDimension : public ::testing::TestWithParam<DimensionCase> {};

}  // namespace

// The dimension is (span of written rows) x (span of written cell columns and
// styled columns); width-only columns never count.
TEST_P(FormulonCApiSheetDimension, IsWrittenRowsByCellAndStyledColumns) {
  const DimensionCase& c = GetParam();
  WorkbookGuard wb;
  MakeWorkbook(wb);
  c.setup(wb);
  EXPECT_EQ(Dimension(SheetXml(Save(wb, FM_WORKBOOK_FORMAT_XLSX))), c.dimension);
  EXPECT_EQ(DecodeSheetBin(Save(wb, FM_WORKBOOK_FORMAT_XLSB)).dim, c.ws_dim);
}

INSTANTIATE_TEST_SUITE_P(
    Cases, FormulonCApiSheetDimension,
    ::testing::Values(
        DimensionCase{"TextRowAndStyledColumn",
                      [](WorkbookGuard& wb) {
                        ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 4U, 0U, "x"), 0);
                        StyleColumns(wb, 7, 7);
                      },
                      "A5:H5", std::vector<std::uint32_t>{4, 4, 0, 7}},
        DimensionCase{"HeightOnlyRow",
                      [](WorkbookGuard& wb) { ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 2U, 2U, 30.0), 0); },
                      "A3", std::vector<std::uint32_t>{2, 2, 0, 0}},
        DimensionCase{"TextAndHeightOnlyRow",
                      [](WorkbookGuard& wb) {
                        ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0U, 0U, "x"), 0);
                        ASSERT_EQ(fm_sheet_set_row_height(wb.handle, 0, 4U, 4U, 30.0), 0);
                      },
                      "A1:A5", std::vector<std::uint32_t>{0, 4, 0, 0}},
        // Width-only columns do not count: A1, and BrtWsDim 0..0 / 0..0 as Excel writes.
        DimensionCase{"WidthOnlyColumns",
                      [](WorkbookGuard& wb) { ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 3U, 5U, 20.0), 0); },
                      "A1", std::vector<std::uint32_t>{0, 0, 0, 0}},
        // A fresh sheet still carries BrtWsDim.
        DimensionCase{"FreshSheet", [](WorkbookGuard&) {}, "A1", std::vector<std::uint32_t>{0, 0, 0, 0}},
        DimensionCase{"StyledAndWidthOnlyColumns",
                      [](WorkbookGuard& wb) {
                        StyleColumns(wb, 3, 5);
                        ASSERT_EQ(fm_sheet_set_column_width(wb.handle, 0, 7U, 7U, 20.0), 0);
                      },
                      "D1:F1", std::vector<std::uint32_t>{0, 0, 3, 5}},
        DimensionCase{"TextAndStyledColumnBlank",
                      [](WorkbookGuard& wb) {
                        ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 0U, 0U, "x"), 0);
                        StyleColumns(wb, 7, 7);
                        Blanks(wb, 8, 7, 7);
                      },
                      "A1:H9", std::vector<std::uint32_t>{0, 8, 0, 7}}),
    [](const ::testing::TestParamInfo<DimensionCase>& param_info) { return std::string(param_info.param.name); });
