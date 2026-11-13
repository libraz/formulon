// Tables in XLSB: `xl/tables/tableN.bin` parts are decoded into
// `Workbook::tables()` and generated from it, so a table survives an
// .xlsb resave and an .xlsx -> .xlsb conversion with its style, flags,
// totals row, AutoFilter and formulas. A part the model cannot hold stays
// verbatim.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "passthrough_part.h"
#include "styles.h"
#include "table.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtBeginList = 343;
constexpr std::uint16_t kBrtEndList = 344;
constexpr std::uint16_t kBrtBeginListParts = 660;
constexpr std::uint16_t kBrtListPart = 661;

constexpr const char* kAltTextExt =
    "<extLst><ext uri=\"{504A1905-F514-4f6f-8877-14C23A59335A}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
    "<x14:table altText=\"Sales by region\" altTextSummary=\"Quarterly\"/></ext></extLst>";

TableMetadata AuthoredTable() {
  TableMetadata t;
  t.id = 3;
  t.name = "Sales";
  t.display_name = "Sales";
  t.ref = "B2:D7";
  t.sheet_index = 0;
  t.header_row = true;
  t.totals_row = true;
  t.insert_row = true;
  // CT_Table attribute order, which is how the reader spells them.
  t.root_extra_attrs =
      " comment=\"Q &amp; A\" published=\"1\" headerRowDxfId=\"1\" dataDxfId=\"2\""
      " totalsRowDxfId=\"3\" headerRowBorderDxfId=\"5\" tableBorderDxfId=\"4\" totalsRowBorderDxfId=\"6\""
      " headerRowCellStyle=\"Heading 1\" dataCellStyle=\"Normal\" totalsRowCellStyle=\"Total\"";
  TableColumn region(1, "Region", "Total", "", "");
  region.extra_attrs = " dataDxfId=\"0\" totalsRowDxfId=\"7\"";
  TableColumn amount(2, "Amount", "", "sum", "");
  amount.extra_attrs = " dataCellStyle=\"Currency\"";
  TableColumn tax(3, "Tax", "", "custom", "C3*0.08");
  tax.totals_row_formula = "SUM(D3:D6)";
  t.columns = {region, amount, tax};
  t.style_info.present = true;
  t.style_info.name = "TableStyleMedium9";
  t.style_info.show_last_column = true;
  t.style_info.show_row_stripes = true;
  t.auto_filter_xml.set(auto_filter_from_xml(
      "<autoFilter ref=\"B2:D6\"><filterColumn colId=\"0\" hiddenButton=\"1\"/><filterColumn colId=\"1\">"
      "<customFilters and=\"1\"><customFilter operator=\"greaterThan\" val=\"5\"/>"
      "<customFilter operator=\"lessThan\" val=\"100\"/></customFilters></filterColumn></autoFilter>"));
  t.ext_lst_xml = kAltTextExt;
  return t;
}

/// Adds the cell styles `AuthoredTable` names, as a saved styles part would hold them.
void AddAuthoredCellStyles(Workbook& wb) {
  std::vector<CellStyleRecord>& styles = wb.mutable_styles().cell_styles;
  for (const char* name : {"Normal", "Heading 1", "Total", "Currency"}) {
    if (std::none_of(styles.begin(), styles.end(), [name](const CellStyleRecord& s) { return s.name == name; })) {
      CellStyleRecord style;
      style.name = name;
      styles.push_back(style);
    }
  }
}

Workbook AuthoredBook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Hyperlink link;
  link.row = 0;
  link.col = 0;
  link.last_row = 0;
  link.last_col = 0;
  link.target = "https://example.com/";
  wb.sheet(0).mutable_hyperlinks().push_back(link);
  wb.mutable_tables().push_back(AuthoredTable());
  AddAuthoredCellStyles(wb);
  // One dxf per id `AuthoredTable` names, each used, so a save keeps every id.
  for (std::size_t i = 0; i < 8U; ++i) {
    DifferentialFormat dxf;
    dxf.has_num_fmt = true;
    dxf.num_fmt_id = static_cast<std::uint16_t>(164U + i);
    dxf.num_fmt_code = std::string(i + 1U, '0');
    wb.mutable_styles().dxfs.push_back(dxf);
  }
  return wb;
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb, WriteDiagnostics* diagnostics = nullptr) {
  auto result = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(result)) << (result ? "" : result.error().message);
  if (!result) {
    return {};
  }
  if (diagnostics != nullptr) {
    *diagnostics = result.value().diagnostics;
  }
  return std::move(result.value().bytes);
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

std::vector<std::uint8_t> Entry(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(entry)) << path;
  return entry ? std::move(entry.value()) : std::vector<std::uint8_t>{};
}

std::string EntryText(const std::vector<std::uint8_t>& package, const std::string& path) {
  const std::vector<std::uint8_t> bytes = Entry(package, path);
  return std::string(bytes.begin(), bytes.end());
}

/// `attrs` with every dxf id (`...DxfId="N"` / `dxfId="N"`) replaced by the number format code of the dxf it
/// names in `styles`, so two tables compare by the formats they reference rather than by saved indices.
std::string ByDxfContent(const std::string& attrs, const StylesTable& styles) {
  std::string out;
  std::size_t at = 0;
  while (at < attrs.size()) {
    const std::size_t hit = attrs.find("xfId=\"", at);
    if (hit == std::string::npos || hit == 0U || (attrs[hit - 1] != 'D' && attrs[hit - 1] != 'd')) {
      const std::size_t next = hit == std::string::npos ? attrs.size() : hit + 1U;
      out.append(attrs, at, next - at);
      at = next;
      continue;
    }
    const std::size_t digits = hit + 6U;
    out.append(attrs, at, digits - at);
    const unsigned long id = std::stoul(attrs.substr(digits));
    out += id < styles.dxfs.size() ? "<" + styles.dxfs[id].num_fmt_code + ">" : "<dangling>";
    at = attrs.find('"', digits);
  }
  return out;
}

/// Compares two tables; with both style tables given, dxf ids compare by the dxf they name.
void ExpectSameTable(const TableMetadata& actual, const TableMetadata& expected,
                     const StylesTable* actual_styles = nullptr, const StylesTable* expected_styles = nullptr) {
  const auto attrs = [&](const std::string& text, bool is_actual) {
    return actual_styles == nullptr ? text : ByDxfContent(text, is_actual ? *actual_styles : *expected_styles);
  };
  EXPECT_EQ(actual.id, expected.id);
  EXPECT_EQ(actual.name, expected.name);
  EXPECT_EQ(actual.display_name, expected.display_name);
  EXPECT_EQ(actual.ref, expected.ref);
  EXPECT_EQ(actual.sheet_index, expected.sheet_index);
  EXPECT_EQ(actual.header_row, expected.header_row);
  EXPECT_EQ(actual.totals_row, expected.totals_row);
  EXPECT_EQ(actual.totals_row_shown, expected.totals_row_shown);
  EXPECT_EQ(actual.insert_row, expected.insert_row);
  EXPECT_EQ(actual.insert_row_shift, expected.insert_row_shift);
  EXPECT_EQ(attrs(actual.root_extra_attrs, true), attrs(expected.root_extra_attrs, false));
  ASSERT_EQ(actual.columns.size(), expected.columns.size());
  for (std::size_t i = 0; i < expected.columns.size(); ++i) {
    SCOPED_TRACE(i);
    const TableColumn& a = actual.columns[i];
    const TableColumn& e = expected.columns[i];
    EXPECT_EQ(a.id, e.id);
    EXPECT_EQ(a.name, e.name);
    EXPECT_EQ(a.totals_label, e.totals_label);
    EXPECT_EQ(a.totals_function, e.totals_function);
    EXPECT_EQ(a.calculated_column_formula, e.calculated_column_formula);
    EXPECT_EQ(a.totals_row_formula, e.totals_row_formula);
    EXPECT_EQ(attrs(a.extra_attrs, true), attrs(e.extra_attrs, false));
  }
  EXPECT_EQ(actual.style_info.present, expected.style_info.present);
  EXPECT_EQ(actual.style_info.name, expected.style_info.name);
  EXPECT_EQ(actual.style_info.show_first_column, expected.style_info.show_first_column);
  EXPECT_EQ(actual.style_info.show_last_column, expected.style_info.show_last_column);
  EXPECT_EQ(actual.style_info.show_row_stripes, expected.style_info.show_row_stripes);
  EXPECT_EQ(actual.style_info.show_column_stripes, expected.style_info.show_column_stripes);
  ASSERT_EQ(actual.auto_filter_xml.empty(), expected.auto_filter_xml.empty());
  if (!expected.auto_filter_xml.empty()) {
    EXPECT_EQ(serialize_auto_filter(*actual.auto_filter_xml.get()),
              serialize_auto_filter(*expected.auto_filter_xml.get()));
  }
  EXPECT_EQ(actual.ext_lst_xml, expected.ext_lst_xml);
}

TEST(XlsbTableFidelity, ModelTableRoundTripsThroughXlsb) {
  WriteDiagnostics diagnostics;
  const Workbook source = AuthoredBook();
  const Workbook back = LoadXlsb(SaveXlsb(source, &diagnostics));
  EXPECT_EQ(diagnostics.deferred_feature_count, 0U);
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], AuthoredTable(), &back.styles(), &source.styles());
  // The hyperlink keeps its own relationship next to the table's.
  ASSERT_EQ(back.sheet(0).hyperlinks().size(), 1U);
  EXPECT_EQ(back.sheet(0).hyperlinks()[0].target, "https://example.com/");
  // The table's relationship and BrtListPart are model-owned, not retained.
  for (const UnknownRelationship& rel : back.sheet(0).unknown_relationships()) {
    EXPECT_EQ(rel.type.find("/table"), std::string::npos) << rel.type;
  }
  for (const FramedRecord& rec : split_records(back.sheet(0).xlsb_tail().after_hyperlinks)) {
    EXPECT_NE(rec.type, kBrtBeginListParts);
    EXPECT_NE(rec.type, kBrtListPart);
  }
}

TEST(XlsbTableFidelity, WrittenPartDecodes) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  TableFormulaBytes formulas;
  auto table = read_table_bin(SpanOf(part), 0, &formulas);
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message << " " << table.error().context;
  EXPECT_EQ(table.value().display_name, "Sales");
  EXPECT_FALSE(formulas.calculated.at(2).empty());
  EXPECT_FALSE(formulas.totals.at(2).empty());
}

TEST(XlsbTableFidelity, PackageCarriesThePartItsRelationshipAndContentType) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const std::vector<FramedRecord> recs = split_records(part);
  ASSERT_FALSE(recs.empty());
  EXPECT_EQ(recs.front().type, kBrtBeginList);
  EXPECT_NE(EntryText(package, "[Content_Types].xml")
                .find("PartName=\"/xl/tables/table3.bin\" ContentType=\"application/vnd.ms-excel.table\""),
            std::string::npos);
  const std::string rels = EntryText(package, "xl/worksheets/_rels/sheet1.bin.rels");
  EXPECT_NE(rels.find("relationships/table\" Target=\"../tables/table3.bin\""), std::string::npos) << rels;

  // The sheet names the part through exactly one BrtListPart.
  const std::vector<FramedRecord> sheet = split_records(Entry(package, "xl/worksheets/sheet1.bin"));
  std::size_t parts = 0;
  bool block = false;
  for (const FramedRecord& rec : sheet) {
    block = block || rec.type == kBrtBeginListParts;
    parts += rec.type == kBrtListPart ? 1U : 0U;
  }
  EXPECT_TRUE(block);
  EXPECT_EQ(parts, 1U);
}

TEST(XlsbTableFidelity, XlsbResaveReproducesTheTablePart) {
  const std::vector<std::uint8_t> first = SaveXlsb(AuthoredBook());
  // The styles part is not read back, so the named styles are supplied again for the second save.
  Workbook reloaded = LoadXlsb(first);
  AddAuthoredCellStyles(reloaded);
  const std::vector<std::uint8_t> second = SaveXlsb(reloaded);
  EXPECT_EQ(Entry(second, "xl/tables/table3.bin"), Entry(first, "xl/tables/table3.bin"));
  const Workbook source = AuthoredBook();
  const Workbook again = LoadXlsb(second);
  ExpectSameTable(again.tables().at(0), AuthoredTable(), &again.styles(), &source.styles());
}

// `part` with the first column's header dxf and the second column's header cell style set, after checking that
// every column was written with both slots empty.
std::vector<std::uint8_t> WithColumnHeaderSlotsSet(const std::vector<std::uint8_t>& part) {
  constexpr std::uint16_t kBrtBeginListCol = 347;
  constexpr std::size_t kHeaderDxfOffset = 8;
  constexpr std::size_t kFirstStringOffset = 24;
  constexpr std::uint8_t kNull[4] = {0xFFU, 0xFFU, 0xFFU, 0xFFU};
  // Skips the unique name, caption and totals label to reach the header cell style slot.
  const auto header_style_offset = [](const std::vector<std::uint8_t>& body) {
    std::size_t at = kFirstStringOffset;
    for (int i = 0; i < 3; ++i) {
      const std::uint32_t len =
          static_cast<std::uint32_t>(body.at(at)) | (static_cast<std::uint32_t>(body.at(at + 1)) << 8U) |
          (static_cast<std::uint32_t>(body.at(at + 2)) << 16U) | (static_cast<std::uint32_t>(body.at(at + 3)) << 24U);
      at += 4U + (len == 0xFFFFFFFFU ? 0U : len * 2U);
    }
    return at;
  };
  std::size_t columns = 0;
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(part)) {
    std::vector<std::uint8_t> body(rec.payload.data, rec.payload.data + rec.payload.size);
    if (rec.type == kBrtBeginListCol) {
      ++columns;
      const std::size_t style_at = header_style_offset(body);
      EXPECT_TRUE(std::equal(kNull, kNull + 4, body.begin() + kHeaderDxfOffset)) << "column " << columns;
      EXPECT_TRUE(std::equal(kNull, kNull + 4, body.begin() + static_cast<std::ptrdiff_t>(style_at)))
          << "column " << columns;
      if (columns == 1U) {
        body[kHeaderDxfOffset] = 0x07U;
        body[kHeaderDxfOffset + 1] = body[kHeaderDxfOffset + 2] = body[kHeaderDxfOffset + 3] = 0U;
      } else if (columns == 2U) {
        const std::uint8_t style[] = {0x01U, 0U, 0U, 0U, 'X', 0U};
        body.erase(body.begin() + static_cast<std::ptrdiff_t>(style_at),
                   body.begin() + static_cast<std::ptrdiff_t>(style_at) + 4);
        body.insert(body.begin() + static_cast<std::ptrdiff_t>(style_at), std::begin(style), std::end(style));
      }
    }
    emit_record(patched, rec.type, body);
  }
  EXPECT_EQ(columns, 3U);
  return patched;
}

// Excel refuses a file whose column carries a header dxf or header cell style, so neither BrtBeginListCol slot
// is read into the model, each drop is counted, and a written column always leaves both empty.
TEST(XlsbTableFidelity, AColumnHeaderDxfOrStyleIsDroppedOnReadAndNeverWritten) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> patched = WithColumnHeaderSlotsSet(Entry(package, "xl/tables/table3.bin"));
  TableFormulaBytes formulas;
  ReadDiagnostics diagnostics;
  auto table = read_table_bin(SpanOf(patched), 0, &formulas, &diagnostics);
  ASSERT_TRUE(static_cast<bool>(table)) << table.error().message;
  for (const TableColumn& col : table.value().columns) {
    EXPECT_EQ(col.extra_attrs.find("headerRow"), std::string::npos) << col.extra_attrs;
  }
  const Workbook saved_book = LoadXlsb(package);
  const Workbook source_book = AuthoredBook();
  EXPECT_EQ(ByDxfContent(table.value().columns.at(0).extra_attrs, saved_book.styles()),
            ByDxfContent(" dataDxfId=\"0\" totalsRowDxfId=\"7\"", source_book.styles()));
  EXPECT_EQ(diagnostics.skipped_feature_count, 2U);

  auto loaded = read_xlsb(SpanOf(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_EQ(loaded.value().skipped_feature_count, 2U);
  ASSERT_EQ(loaded.value().workbook.tables().size(), 1U);
  for (const TableColumn& col : loaded.value().workbook.tables()[0].columns) {
    EXPECT_EQ(col.extra_attrs.find("headerRow"), std::string::npos) << col.extra_attrs;
  }
}

TEST(XlsbTableFidelity, AnXlsxColumnHeaderDxfOrStyleIsCountedOnLoad) {
  auto xlsx = write_ooxml(AuthoredBook());
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const std::string path = "xl/tables/table3.xml";
  std::string xml = EntryText(xlsx.value(), path);
  const std::string column = "<tableColumn id=\"";
  const std::size_t first = xml.find(column);
  ASSERT_NE(first, std::string::npos) << xml;
  const std::size_t second = xml.find(column, first + 1U);
  ASSERT_NE(second, std::string::npos) << xml;
  xml.insert(second + 13U, "headerRowCellStyle=\"Heading 1\" ");
  xml.insert(first + 13U, "headerRowDxfId=\"0\" ");
  auto loaded =
      read_ooxml(SpanOf(RepackWithEntries(xlsx.value(), {{path, std::vector<std::uint8_t>(xml.begin(), xml.end())}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_EQ(loaded.value().diagnostics.skipped_feature_count, 2U);
  ASSERT_EQ(loaded.value().workbook.tables().size(), 1U);
  for (const TableColumn& col : loaded.value().workbook.tables()[0].columns) {
    EXPECT_EQ(col.extra_attrs.find("headerRow"), std::string::npos) << col.extra_attrs;
  }
}

TEST(XlsbTableFidelity, XlsxTableSurvivesConversionToXlsb) {
  auto xlsx = write_ooxml(AuthoredBook());
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto loaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& from_xlsx = loaded.value().workbook;
  ASSERT_EQ(from_xlsx.tables().size(), 1U);
  WriteDiagnostics diagnostics;
  const Workbook back = LoadXlsb(SaveXlsb(from_xlsx, &diagnostics));
  EXPECT_EQ(diagnostics.deferred_feature_count, 0U);
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], from_xlsx.tables()[0], &back.styles(), &from_xlsx.styles());
}

TEST(XlsbTableFidelity, ADanglingCellStyleNameIsWrittenEmpty) {
  Workbook wb = AuthoredBook();
  wb.mutable_styles().cell_styles.erase(wb.mutable_styles().cell_styles.begin() + 2);  // "Total"
  wb.mutable_tables()[0].columns[1].extra_attrs = " dataCellStyle=\"Missing\"";
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.tables().size(), 1U);
  const TableMetadata& table = back.tables()[0];
  EXPECT_NE(table.root_extra_attrs.find("headerRowCellStyle=\"Heading 1\""), std::string::npos);
  EXPECT_NE(table.root_extra_attrs.find("dataCellStyle=\"Normal\""), std::string::npos);
  EXPECT_EQ(table.root_extra_attrs.find("totalsRowCellStyle"), std::string::npos) << table.root_extra_attrs;
  EXPECT_EQ(table.columns[1].extra_attrs.find("CellStyle"), std::string::npos) << table.columns[1].extra_attrs;
}

TEST(XlsxTableCellStyles, OnlyNamesOfSavedCellStylesAreWritten) {
  Workbook wb = AuthoredBook();
  wb.mutable_styles().cell_styles.erase(wb.mutable_styles().cell_styles.begin() + 2);  // "Total"
  wb.mutable_tables()[0].columns[1].extra_attrs = " dataCellStyle=\"Currency\"";
  wb.mutable_tables()[0].columns[0].extra_attrs = " dataCellStyle=\"Missing\" dataDxfId=\"0\"";
  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const std::string xml = EntryText(xlsx.value(), "xl/tables/table3.xml");
  EXPECT_NE(xml.find("headerRowCellStyle=\"Heading 1\""), std::string::npos) << xml;
  EXPECT_NE(xml.find(" dataCellStyle=\"Normal\""), std::string::npos) << xml;
  EXPECT_EQ(xml.find("totalsRowCellStyle"), std::string::npos) << xml;
  EXPECT_EQ(xml.find("Missing"), std::string::npos) << xml;
  EXPECT_NE(xml.find("dataDxfId=\"0\""), std::string::npos) << xml;
  EXPECT_NE(xml.find("dataCellStyle=\"Currency\""), std::string::npos) << xml;
}

TEST(XlsbTableFidelity, RevisionUidsTravelAsAcuidBlocks) {
  Workbook wb = AuthoredBook();
  TableMetadata& t = wb.mutable_tables()[0];
  // How Excel 365 spells the uids of a table, its first column and its
  // AutoFilter in `.xlsx`.
  t.root_extra_attrs =
      " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"xr xr3\""
      " xmlns:xr=\"http://schemas.microsoft.com/office/spreadsheetml/2014/revision\""
      " xmlns:xr3=\"http://schemas.microsoft.com/office/spreadsheetml/2016/revision3\""
      " xr:uid=\"{48722CD1-E4B0-409F-A1F2-C5EC8AE2D85A}\"" +
      t.root_extra_attrs;
  t.columns[0].extra_attrs = " xr3:uid=\"{0FB2B22E-7B6A-4D4C-9C1E-0D3B3A51E7F1}\"" + t.columns[0].extra_attrs;
  t.auto_filter_xml.get()->extra_attrs.insert(t.auto_filter_xml.get()->extra_attrs.begin(),
                                              {"xr:uid", "{00000000-0009-0000-0100-000001000000}"});
  const std::vector<std::uint8_t> package = SaveXlsb(wb);
  std::size_t acuids = 0;
  for (const FramedRecord& rec : split_records(Entry(package, "xl/tables/table3.bin"))) {
    acuids += rec.type == 3072U ? 1U : 0U;
  }
  EXPECT_EQ(acuids, 3U);
  Workbook back = LoadXlsb(package);
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], t, &back.styles(), &wb.styles());
  // And the `.xlsx` the uids reach stays well formed.
  AddAuthoredCellStyles(back);
  auto xlsx = write_ooxml(back);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto again = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(again)) << again.error().message;
  ASSERT_EQ(again.value().workbook.tables().size(), 1U);
  const TableMetadata& reread = again.value().workbook.tables()[0];
  EXPECT_EQ(ByDxfContent(reread.root_extra_attrs, again.value().workbook.styles()),
            ByDxfContent(t.root_extra_attrs, wb.styles()));
  EXPECT_NE(reread.columns[0].extra_attrs.find("xr3:uid=\"{0FB2B22E-7B6A-4D4C-9C1E-0D3B3A51E7F1}\""),
            std::string::npos);
}

TEST(XlsbTableFidelity, NeverShownTotalsRowAndNoStyleNameSurvive) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  TableMetadata t;
  t.id = 1;
  t.name = "Prog";
  t.display_name = "Shown";
  t.ref = "A1:B4";
  t.header_row = false;
  t.totals_row_shown = false;
  t.root_extra_attrs = " tableType=\"queryTable\"";
  t.insert_row_shift = true;
  TableColumn a(1, "A", "", "", "");
  a.extra_attrs = " uniqueName=\"A\" queryTableFieldId=\"1\"";
  TableColumn b(2, "B", "", "", "");
  b.extra_attrs = " uniqueName=\"B\" queryTableFieldId=\"2\"";
  t.columns = {a, b};
  t.style_info.present = true;
  t.style_info.show_column_stripes = true;
  wb.mutable_tables().push_back(t);
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], t);
}

/// A table `<sortState>` in the serializer's spelling.
std::string CanonicalSort(const std::string& xml) {
  return serialize_auto_filter(auto_filter_from_xml("<autoFilter ref=\"A1\">" + xml + "</autoFilter>"));
}

TEST(XlsbTableFidelity, SortStateTravelsInsideTheAutoFilter) {
  Workbook wb = AuthoredBook();
  const std::string sort =
      "<sortState ref=\"B3:D6\" caseSensitive=\"1\"><sortCondition descending=\"1\" ref=\"C3:C6\"/>"
      "<sortCondition ref=\"B3:B6\" customList=\"North,South\"/>"
      "<sortCondition sortBy=\"cellColor\" ref=\"D3:D6\" dxfId=\"2\"/></sortState>";
  wb.mutable_tables()[0].sort_state_xml = sort;
  WriteDiagnostics diagnostics;
  const std::vector<std::uint8_t> package = SaveXlsb(wb, &diagnostics);
  EXPECT_EQ(diagnostics.deferred_feature_count, 0U);
  const Workbook back = LoadXlsb(package);
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_EQ(ByDxfContent(CanonicalSort(back.tables()[0].sort_state_xml), back.styles()),
            ByDxfContent(CanonicalSort(sort), wb.styles()));
  // Without an AutoFilter the sort state is its own block after BrtBeginList.
  wb.mutable_tables()[0].auto_filter_xml.reset();
  const Workbook bare = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(bare.tables().size(), 1U);
  EXPECT_EQ(ByDxfContent(CanonicalSort(bare.tables()[0].sort_state_xml), bare.styles()),
            ByDxfContent(CanonicalSort(sort), wb.styles()));
  EXPECT_TRUE(bare.tables()[0].auto_filter_xml.empty());
}

TEST(XlsbTableFidelity, UnrepresentableSettingsAreCountedAsDeferred) {
  Workbook wb = AuthoredBook();
  TableMetadata& t = wb.mutable_tables()[0];
  t.sort_state_xml = "<sortState ref=\"B3:D6\"><sortCondition ref=\"C3:C6\"/></sortState>";
  t.ext_lst_xml = "<extLst><ext uri=\"{0000}\"/></extLst>";
  t.columns[0].trailing_children_xml = "<xmlColumnPr mapId=\"1\" xpath=\"/a\" xmlDataType=\"string\"/>";
  t.auto_filter_xml.get()->columns[1].ext_xml = "<ext uri=\"{1111}\"/>";
  WriteDiagnostics diagnostics;
  const Workbook back = LoadXlsb(SaveXlsb(wb, &diagnostics));
  EXPECT_EQ(diagnostics.deferred_feature_count, 3U);
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_TRUE(back.tables()[0].ext_lst_xml.empty());
  EXPECT_TRUE(back.tables()[0].columns[0].trailing_children_xml.empty());
}

TEST(XlsbTableFidelity, TableOnAMissingSheetFailsTheSave) {
  Workbook wb = AuthoredBook();
  wb.mutable_tables()[0].sheet_index = 4;
  auto result = write_xlsb_with_result(wb);
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoWriteFailed);
}

TEST(XlsbTableFidelity, TablesOnTwoSheetsKeepTheirOwnParts) {
  Workbook wb = AuthoredBook();
  wb.add_sheet("Second");
  TableMetadata other = AuthoredTable();
  other.id = 3;  // duplicate id: the writer mints another
  other.name = "Other";
  other.display_name = "Other";
  other.sheet_index = 1;
  wb.mutable_tables().push_back(other);
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.tables().size(), 2U);
  EXPECT_EQ(back.tables()[0].sheet_index, 0U);
  EXPECT_EQ(back.tables()[1].sheet_index, 1U);
  EXPECT_EQ(back.tables()[1].display_name, "Other");
  EXPECT_NE(back.tables()[0].id, back.tables()[1].id);
}

/// `part` with a future-record block the model has no field for before
/// BrtEndList.
std::vector<std::uint8_t> WithUnknownFutureRecord(const std::vector<std::uint8_t>& part) {
  std::vector<std::uint8_t> patched;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == kBrtEndList) {
      emit_record(patched, 35, std::vector<std::uint8_t>{0x02, 0x0E, 0x00, 0x00});
      emit_record(patched, 5000, std::vector<std::uint8_t>{0x01, 0x01, 0x00});
      emit_record(patched, 36, ByteSpan{});
    }
    append_record(patched, part, rec);
  }
  return patched;
}

TEST(XlsbTableFidelity, PartOutsideTheModelStaysVerbatim) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(part);
  const std::vector<std::uint8_t> source = RepackWithEntries(package, {{"xl/tables/table3.bin", patched}});
  auto read = read_xlsb(SpanOf(source));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  EXPECT_EQ(read.value().skipped_feature_count, 0U);
  const Workbook wb = LoadXlsb(source);
  EXPECT_TRUE(wb.tables().empty());
  // Resaving keeps the part, its relationship and its BrtListPart as read.
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb);
  EXPECT_EQ(Entry(resaved, "xl/tables/table3.bin"), patched);
  const Workbook again = LoadXlsb(resaved);
  EXPECT_TRUE(again.tables().empty());
  bool has_table_rel = false;
  for (const UnknownRelationship& rel : again.sheet(0).unknown_relationships()) {
    has_table_rel = has_table_rel || rel.target == "xl/tables/table3.bin";
  }
  EXPECT_TRUE(has_table_rel);
}

// A part kept verbatim is still written without the column header slots Excel refuses.
TEST(XlsbTableFidelity, AVerbatimPartIsKeptWithItsColumnHeaderSlotsCleared) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(WithColumnHeaderSlotsSet(part));
  auto loaded = read_xlsb(SpanOf(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_TRUE(loaded.value().workbook.tables().empty());
  EXPECT_EQ(loaded.value().skipped_feature_count, 2U);
  EXPECT_EQ(Entry(SaveXlsb(loaded.value().workbook), "xl/tables/table3.bin"), WithUnknownFutureRecord(part));
}

// A part whose records do not frame cannot be cleared, so it is dropped with its relationship and list record.
TEST(XlsbTableFidelity, AnUnframeablePartIsDroppedWithItsRelationship) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  std::vector<std::uint8_t> part = WithColumnHeaderSlotsSet(Entry(package, "xl/tables/table3.bin"));
  part.pop_back();
  auto loaded = read_xlsb(SpanOf(RepackWithEntries(package, {{"xl/tables/table3.bin", part}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  EXPECT_TRUE(wb.tables().empty());
  EXPECT_EQ(loaded.value().skipped_feature_count, 1U);
  for (const PassthroughPart& kept : wb.passthrough_parts()) {
    EXPECT_NE(kept.path, "xl/tables/table3.bin");
  }
  for (const UnknownRelationship& rel : wb.sheet(0).unknown_relationships()) {
    EXPECT_NE(rel.target, "xl/tables/table3.bin");
  }
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(resaved))));
  EXPECT_FALSE(zip.has_entry("xl/tables/table3.bin"));
  for (const FramedRecord& rec : split_records(Entry(resaved, "xl/worksheets/sheet1.bin"))) {
    EXPECT_NE(rec.type, kBrtBeginListParts);
  }
  EXPECT_TRUE(LoadXlsb(resaved).tables().empty());
}

// A column with an unknown totals function makes the part unreadable to the model, but its header slots are still
// located and cleared, so the part is kept rather than dropped.
TEST(XlsbTableFidelity, AnUnreadableColumnStillHasItsHeaderSlotsCleared) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const auto with_unknown_totals_function = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::uint8_t> out;
    for (const FramedRecord& rec : split_records(part)) {
      std::vector<std::uint8_t> body(rec.payload.data, rec.payload.data + rec.payload.size);
      if (rec.type == 347U) {
        body[4] = 0x63U;
      }
      emit_record(out, rec.type, body);
    }
    return out;
  };
  const std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const std::vector<std::uint8_t> patched = with_unknown_totals_function(WithColumnHeaderSlotsSet(part));
  auto loaded = read_xlsb(SpanOf(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_TRUE(loaded.value().workbook.tables().empty());
  EXPECT_EQ(loaded.value().skipped_feature_count, 2U);
  EXPECT_EQ(Entry(SaveXlsb(loaded.value().workbook), "xl/tables/table3.bin"), with_unknown_totals_function(part));
}

// `package` with `added` appended as new entries.
std::vector<std::uint8_t> WithAddedEntries(const std::vector<std::uint8_t>& package,
                                           const std::map<std::string, std::string>& added) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  const auto add = [&writer](const std::string& name, const void* data, std::size_t size) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name.c_str(), data, size, static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << name;
  };
  for (const std::string& name : zip.list_entries()) {
    const std::vector<std::uint8_t> bytes = Entry(package, name);
    add(name, bytes.data(), bytes.size());
  }
  for (const auto& [name, body] : added) {
    add(name, body.data(), body.size());
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  return out;
}

// A dropped table part takes its rels part and the parts only that rels part referenced with it; a kept one keeps
// them.
TEST(XlsbTableFidelity, ADroppedPartTakesItsRelsAndTheirOnlyTargetsWithIt) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  std::string types = EntryText(package, "[Content_Types].xml");
  const std::size_t close = types.rfind("</Types>");
  ASSERT_NE(close, std::string::npos);
  types.insert(close,
               "<Override PartName=\"/xl/queryTables/queryTable1.bin\" "
               "ContentType=\"application/vnd.ms-excel.queryTable\"/>");
  const std::vector<std::uint8_t> with_query = WithAddedEntries(
      RepackWithEntries(package, {{"[Content_Types].xml", std::vector<std::uint8_t>(types.begin(), types.end())}}),
      {{"xl/tables/_rels/table3.bin.rels",
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships "
        "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable\" "
        "Target=\"../queryTables/queryTable1.bin\"/></Relationships>"},
       {"xl/queryTables/queryTable1.bin", std::string("\x00\x00", 2)}});
  const std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const auto kept_paths = [](const Workbook& wb) {
    std::vector<std::string> paths;
    for (const PassthroughPart& kept : wb.passthrough_parts()) {
      paths.push_back(kept.path);
    }
    return paths;
  };
  const auto contains = [](const std::vector<std::string>& paths, const char* path) {
    return std::find(paths.begin(), paths.end(), path) != paths.end();
  };

  const Workbook kept =
      LoadXlsb(RepackWithEntries(with_query, {{"xl/tables/table3.bin", WithUnknownFutureRecord(part)}}));
  EXPECT_TRUE(contains(kept_paths(kept), "xl/tables/table3.bin"));
  EXPECT_TRUE(contains(kept_paths(kept), "xl/queryTables/queryTable1.bin"));

  std::vector<std::uint8_t> truncated = part;
  truncated.pop_back();
  auto dropped = read_xlsb(SpanOf(RepackWithEntries(with_query, {{"xl/tables/table3.bin", truncated}})));
  ASSERT_TRUE(static_cast<bool>(dropped)) << dropped.error().message;
  EXPECT_EQ(dropped.value().skipped_feature_count, 1U);
  const std::vector<std::string> paths = kept_paths(dropped.value().workbook);
  EXPECT_FALSE(contains(paths, "xl/tables/table3.bin"));
  EXPECT_FALSE(contains(paths, "xl/tables/_rels/table3.bin.rels"));
  EXPECT_FALSE(contains(paths, "xl/queryTables/queryTable1.bin"));
}

TEST(XlsbTableFidelity, ModelTableJoinsARetainedListPartsBlock) {
  // A sheet whose source table part stayed verbatim gains a model table.
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  std::vector<std::uint8_t> part = Entry(package, "xl/tables/table3.bin");
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(part);
  Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  TableMetadata added = AuthoredTable();
  added.id = 9;
  added.name = "Added";
  added.display_name = "Added";
  added.ref = "F2:H7";
  added.auto_filter_xml.reset();
  wb.mutable_tables().push_back(added);
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb);
  std::size_t parts = 0;
  std::size_t blocks = 0;
  for (const FramedRecord& rec : split_records(Entry(resaved, "xl/worksheets/sheet1.bin"))) {
    blocks += rec.type == kBrtBeginListParts ? 1U : 0U;
    parts += rec.type == kBrtListPart ? 1U : 0U;
  }
  EXPECT_EQ(blocks, 1U);
  EXPECT_EQ(parts, 2U);
  EXPECT_EQ(Entry(resaved, "xl/tables/table3.bin"), patched);
  EXPECT_FALSE(Entry(resaved, "xl/tables/table9.bin").empty());
}

TEST(XlsbTableFidelity, ModelTableAvoidsTheIdOfAVerbatimPart) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(Entry(package, "xl/tables/table3.bin"));
  Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  TableMetadata added = AuthoredTable();
  added.id = 3;  // the retained part's idList
  added.name = "Added";
  added.display_name = "Added";
  added.ref = "F2:H7";
  added.auto_filter_xml.reset();
  wb.mutable_tables().push_back(added);
  const Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.tables().size(), 1U);
  EXPECT_EQ(back.tables()[0].display_name, "Added");
  EXPECT_NE(back.tables()[0].id, 3U);
}

TEST(XlsbTableFidelity, CopiedSheetTableAvoidsTheIdOfAVerbatimPart) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(Entry(package, "xl/tables/table3.bin"));
  Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  // The sheet holding the verbatim part is not copyable, so the copied table lives on another sheet.
  const std::size_t other = wb.add_sheet("Other");
  TableMetadata added = AuthoredTable();
  added.id = 2;
  added.name = "Added";
  added.display_name = "Added";
  added.ref = "F2:H7";
  added.sheet_index = other;
  added.auto_filter_xml.reset();
  wb.mutable_tables().push_back(added);
  ASSERT_TRUE(wb.copy_sheet(other, other + 1U, std::nullopt).has_value());
  ASSERT_EQ(wb.tables().size(), 2U);
  EXPECT_NE(wb.tables()[1].id, 3U);
  EXPECT_NE(wb.tables()[1].id, wb.tables()[0].id);
}

TEST(XlsbTableFidelity, NewTableAvoidsTheRangeOfAVerbatimPart) {
  const std::vector<std::uint8_t> package = SaveXlsb(AuthoredBook());
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(Entry(package, "xl/tables/table3.bin"));
  Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table3.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  TableMetadata added;
  added.name = "Added";
  added.display_name = "Added";
  added.ref = "D7:E8";
  added.columns = {TableColumn{}, TableColumn{}};
  added.columns[0].name = "a";
  added.columns[1].name = "b";
  auto overlap = wb.add_table(added);
  ASSERT_FALSE(overlap.has_value());
  EXPECT_EQ(overlap.error().message, "table range overlaps another table");
  added.ref = "E7:F8";
  EXPECT_TRUE(wb.add_table(added).has_value());
}

/// Sales B2:D7 with a header row, four data rows and a totals row, plus
/// formulas that reach it through structured references.
Workbook StructuredRefBook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Other");
  TableMetadata t;
  t.id = 4;
  t.name = "Sales";
  t.display_name = "Sales";
  t.ref = "B2:D7";
  t.totals_row = true;
  t.columns = {TableColumn(1, "Region", "Total", "", ""), TableColumn(2, "Amount", "", "sum", ""),
               TableColumn(3, "Tax", "", "custom", "Sales[[#This Row],[Amount]]*0.08")};
  t.columns[2].totals_row_formula = "SUM(Sales[Tax])";
  t.style_info.present = true;
  t.style_info.name = "TableStyleLight1";
  wb.mutable_tables().push_back(t);
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 1, 1, "Region")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 1, 2, "Amount")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 1, 3, "Tax")));
  const char* regions[] = {"North", "South", "East", "West"};
  for (std::uint32_t i = 0; i < 4U; ++i) {
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 2 + i, 1, regions[i])));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_value(0, 2 + i, 2, Value::number(10.0 * (i + 1)))));
    EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 2 + i, 3, "=Sales[[#This Row],[Amount]]*0.08")));
  }
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_text(0, 6, 1, "Total")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 2, "=SUBTOTAL(109,Sales[Amount])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 6, 3, "=SUM(Sales[Tax])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 0, 0, "=SUM(Sales[Amount])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 1, 0, "=ROWS(Sales[#All])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 2, 0, "=SUM(Sales[[#Data],[#Totals],[Amount]])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 3, 0, "=COUNTA(Sales[[#Headers],[Region]:[Tax]])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 4, 0, "=SUM(Sales[[Amount]:[Tax]])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_defined_name("SalesTotal", "SUM(Sales[Amount])")));
  EXPECT_TRUE(static_cast<bool>(wb.set_cell_formula(1, 5, 0, "=SalesTotal")));
  EXPECT_TRUE(static_cast<bool>(wb.recalc(eval::default_registry())));
  return wb;
}

TEST(XlsbTableFidelity, StructuredReferencesSurviveAsPtgList) {
  const Workbook wb = StructuredRefBook();
  WriteDiagnostics diagnostics;
  Workbook back = LoadXlsb(SaveXlsb(wb, &diagnostics));
  EXPECT_EQ(diagnostics.downgraded_formula_count, 0U);
  EXPECT_EQ(diagnostics.deferred_feature_count, 0U);
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], wb.tables()[0]);
  const std::pair<std::size_t, std::pair<std::uint32_t, std::uint32_t>> cells[] = {
      {0, {2, 3}}, {0, {5, 3}}, {0, {6, 2}}, {0, {6, 3}}, {1, {0, 0}},
      {1, {1, 0}}, {1, {2, 0}}, {1, {3, 0}}, {1, {4, 0}}, {1, {5, 0}}};
  ASSERT_EQ(back.defined_names().size(), wb.defined_names().size());
  EXPECT_EQ(back.defined_names().back().formula, wb.defined_names().back().formula);
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  for (const auto& [sheet, rc] : cells) {
    SCOPED_TRACE(std::to_string(sheet) + ":" + std::to_string(rc.first) + "," + std::to_string(rc.second));
    const Cell* before = wb.sheet(sheet).cell_at(rc.first, rc.second);
    const Cell* after = back.sheet(sheet).cell_at(rc.first, rc.second);
    ASSERT_NE(before, nullptr);
    ASSERT_NE(after, nullptr);
    EXPECT_EQ(after->formula_text, before->formula_text);
    const Value expected = wb.sheet(sheet).resolve_cell_value(rc.first, rc.second);
    const Value actual = back.sheet(sheet).resolve_cell_value(rc.first, rc.second);
    ASSERT_TRUE(expected.is_number());
    ASSERT_TRUE(actual.is_number());
    EXPECT_DOUBLE_EQ(actual.as_number(), expected.as_number());
  }
}

/// `part` with the PtgList of its calculated-column formula naming table 99,
/// which no part holds, so the reader cannot model the table.
std::vector<std::uint8_t> WithDanglingCalculatedColumn(const std::vector<std::uint8_t>& part) {
  constexpr std::size_t kListIndexAtPtg = 6;
  std::vector<std::uint8_t> patched = part;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type != 351U) {
      continue;
    }
    // fArray (1) and cce (4) precede the rgce.
    const std::size_t ptg = static_cast<std::size_t>(rec.payload.data - part.data()) + 5U;
    EXPECT_EQ(part[ptg], 0x18);
    patched[ptg + kListIndexAtPtg] = 99;
    patched[ptg + kListIndexAtPtg + 1U] = 0;
  }
  return patched;
}

TEST(XlsbTableFidelity, RetainedTableStructuredRefsSurviveResave) {
  const std::vector<std::uint8_t> package = SaveXlsb(StructuredRefBook());
  const std::vector<std::uint8_t> patched = WithDanglingCalculatedColumn(Entry(package, "xl/tables/table4.bin"));
  const Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table4.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  ASSERT_NE(wb.sheet(1).cell_at(0, 0), nullptr);
  ASSERT_EQ(wb.sheet(1).cell_at(0, 0)->formula_text, "=SUM(Sales[Amount])");
  // Cells and a defined name name the part the model does not hold.
  WriteDiagnostics diagnostics;
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb, &diagnostics);
  ASSERT_FALSE(resaved.empty());
  EXPECT_EQ(diagnostics.downgraded_formula_count, 0U);
  EXPECT_EQ(Entry(resaved, "xl/tables/table4.bin"), patched);
  const Workbook back = LoadXlsb(resaved);
  EXPECT_TRUE(back.tables().empty());
  const std::pair<std::uint32_t, const char*> cells[] = {{0, "=SUM(Sales[Amount])"},
                                                         {1, "=ROWS(Sales[#All])"},
                                                         {2, "=SUM(Sales[[#Data],[#Totals],[Amount]])"},
                                                         {3, "=COUNTA(Sales[[#Headers],[Region]:[Tax]])"},
                                                         {4, "=SUM(Sales[[Amount]:[Tax]])"}};
  for (const auto& [row, formula] : cells) {
    SCOPED_TRACE(row);
    const Cell* cell = back.sheet(1).cell_at(row, 0);
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->formula_text, formula);
  }
  bool has_name = false;
  for (const auto& entry : back.defined_names()) {
    if (entry.name == "SalesTotal") {
      has_name = true;
      EXPECT_EQ(entry.formula, "SUM(Sales[Amount])");
    }
  }
  EXPECT_TRUE(has_name);
}

TEST(XlsbTableFidelity, StructuredRefsIntoAPartRejectedOnReadSurviveResave) {
  const std::vector<std::uint8_t> package = SaveXlsb(StructuredRefBook());
  const std::vector<std::uint8_t> patched = WithUnknownFutureRecord(Entry(package, "xl/tables/table4.bin"));
  const Workbook wb = LoadXlsb(RepackWithEntries(package, {{"xl/tables/table4.bin", patched}}));
  ASSERT_TRUE(wb.tables().empty());
  const auto expect_formulas = [](const Workbook& book) {
    const Cell* cell = book.sheet(1).cell_at(0, 0);
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->formula_text, "=SUM(Sales[Amount])");
    bool has_name = false;
    for (const auto& entry : book.defined_names()) {
      if (entry.name == "SalesTotal") {
        has_name = true;
        EXPECT_EQ(entry.formula, "SUM(Sales[Amount])");
      }
    }
    EXPECT_TRUE(has_name);
  };
  expect_formulas(wb);
  WriteDiagnostics diagnostics;
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb, &diagnostics);
  EXPECT_EQ(diagnostics.downgraded_formula_count, 0U);
  EXPECT_EQ(Entry(resaved, "xl/tables/table4.bin"), patched);
  const Workbook back = LoadXlsb(resaved);
  EXPECT_TRUE(back.tables().empty());
  expect_formulas(back);
}

// A formula referencing a table part dropped on load cannot be decoded: it counts as undecoded, the cell keeps its
// value, and the save neither references the table nor fails.
TEST(XlsbTableFidelity, StructuredRefsIntoADroppedPartCountAsUndecoded) {
  const std::vector<std::uint8_t> package = SaveXlsb(StructuredRefBook());
  // A first column record too short to hold its names: the identity still reads, the header slots do not.
  std::vector<std::uint8_t> part;
  bool shortened = false;
  for (const FramedRecord& rec : split_records(Entry(package, "xl/tables/table4.bin"))) {
    std::vector<std::uint8_t> body(rec.payload.data, rec.payload.data + rec.payload.size);
    if (rec.type == 347U && !shortened) {
      body.resize(20U);
      shortened = true;
    }
    emit_record(part, rec.type, body);
  }
  auto loaded = read_xlsb(SpanOf(RepackWithEntries(package, {{"xl/tables/table4.bin", part}})));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  EXPECT_EQ(loaded.value().skipped_feature_count, 1U);
  EXPECT_GT(loaded.value().undecoded_formula_count, 0U);
  EXPECT_TRUE(wb.tables().empty());
  const Cell* cell = wb.sheet(1).cell_at(0, 0);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  WriteDiagnostics diagnostics;
  const std::vector<std::uint8_t> resaved = SaveXlsb(wb, &diagnostics);
  EXPECT_EQ(diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(resaved))));
  EXPECT_FALSE(zip.has_entry("xl/tables/table4.bin"));
  EXPECT_TRUE(LoadXlsb(resaved).tables().empty());
}

TEST(XlsbTableFidelity, NameUsedOnlyInATableFormulaIsWritten) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  TableMetadata t;
  t.id = 1;
  t.name = "Table1";
  t.display_name = "Table1";
  t.ref = "A1:B4";
  t.columns = {TableColumn(1, "Key", "", "", ""),
               TableColumn(2, "Found", "", "", "XLOOKUP(Table1[[#This Row],[Key]],A1:A4,B1:B4)")};
  t.style_info.present = true;
  wb.mutable_tables().push_back(t);
  WriteDiagnostics diagnostics;
  const std::vector<std::uint8_t> saved = SaveXlsb(wb, &diagnostics);
  EXPECT_EQ(diagnostics.deferred_feature_count, 0U);
  bool has_formula = false;
  for (const FramedRecord& rec : split_records(Entry(saved, "xl/tables/table1.bin"))) {
    has_formula = has_formula || rec.type == 351U;
  }
  EXPECT_TRUE(has_formula);
  const Workbook back = LoadXlsb(saved);
  ASSERT_EQ(back.tables().size(), 1U);
  // Formula text is read back in storage form, as the OOXML reader holds it.
  EXPECT_EQ(back.tables()[0].columns[1].calculated_column_formula,
            "_xlfn.XLOOKUP(Table1[[#This Row],[Key]],A1:A4,B1:B4)");
}

TEST(XlsbTableFidelity, SheetIdsSurviveSaveAndLoad) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("A");
  wb.add_sheet("B");
  wb.sheet(0).set_source_sheet_id(3U);
  wb.sheet(1).set_source_sheet_id(1U);
  Workbook back = LoadXlsb(SaveXlsb(wb));
  ASSERT_EQ(back.sheet_count(), 2U);
  EXPECT_EQ(back.sheet(0).source_sheet_id(), 3U);
  EXPECT_EQ(back.sheet(1).source_sheet_id(), 1U);
  back.add_sheet("C");
  const Workbook again = LoadXlsb(SaveXlsb(back));
  ASSERT_EQ(again.sheet_count(), 3U);
  EXPECT_EQ(again.sheet(2).source_sheet_id(), 4U);
}

TEST(XlsbTableFidelity, CalculatedColumnPtgListHasExcelFlags) {
  // The [MS-XLSB] 3.3.11 example "=Table1[[#This Row],[Price]]*.08" over Table1 (id 1) whose second column is
  // Price, on the first sheet. The flags keep the unused bit clear, as Excel writes them.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  TableMetadata t;
  t.id = 1;
  t.name = "Table1";
  t.display_name = "Table1";
  t.ref = "C4:E7";
  t.columns = {TableColumn(1, "Item", "", "", ""), TableColumn(2, "Price", "", "", ""),
               TableColumn(3, "Sales Tax", "", "", "Table1[[#This Row],[Price]]*0.08")};
  t.style_info.present = true;
  wb.mutable_tables().push_back(t);
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 4, 4, "=Table1[[#This Row],[Price]]*0.08")));
  const std::vector<std::uint8_t> part = Entry(SaveXlsb(wb), "xl/tables/table1.bin");
  std::vector<std::uint8_t> formula;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == 351U) {
      formula.assign(rec.payload.data, rec.payload.data + rec.payload.size);
    }
  }
  const std::vector<std::uint8_t> expected = {0x00,                    // fArray clear
                                              0x18, 0x00, 0x00, 0x00,  // cce
                                              0x18, 0x19, 0x00, 0x00, 0x41, 0x04, 0x01, 0x00, 0x00,
                                              0x00, 0x01, 0x00,                                      // PtgList
                                              0x01, 0x00,                                            //
                                              0x1F, 0x7B, 0x14, 0xAE, 0x47, 0xE1, 0x7A, 0xB4, 0x3F,  // PtgNum 0.08
                                              0x05,                                                  // PtgMul
                                              0x00, 0x00, 0x00, 0x00};                               // cb
  EXPECT_EQ(formula, expected);
}

TEST(XlsbTableFidelity, XlsxStructuredReferencesSurviveConversionToXlsb) {
  auto xlsx = write_ooxml(StructuredRefBook());
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto loaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& from_xlsx = loaded.value().workbook;
  WriteDiagnostics diagnostics;
  const Workbook back = LoadXlsb(SaveXlsb(from_xlsx, &diagnostics));
  EXPECT_EQ(diagnostics.downgraded_formula_count, 0U);
  ASSERT_EQ(back.tables().size(), 1U);
  ExpectSameTable(back.tables()[0], from_xlsx.tables()[0]);
  for (std::uint32_t row = 0; row < 5U; ++row) {
    const Cell* before = from_xlsx.sheet(1).cell_at(row, 0);
    const Cell* after = back.sheet(1).cell_at(row, 0);
    ASSERT_NE(before, nullptr);
    ASSERT_NE(after, nullptr);
    EXPECT_EQ(after->formula_text, before->formula_text);
  }
}

TEST(XlsbTableFidelity, AutoFilterRecordsMatchTheExcelXlsxTwin) {
  // Windows Excel 365 saved the same sheet AutoFilter as .xlsb and .xlsx.
  const std::string dir = std::string(FORMULON_FIXTURES_DIR) + "/excel/retained_children_edit/";
  const std::vector<std::uint8_t> sheet =
      Entry(ReadFixtureFile(dir + "autofilter_0_base.xlsb"), "xl/worksheets/sheet1.bin");
  ByteSpan cursor = SpanOf(sheet);
  std::optional<AutoFilter> decoded;
  while (cursor.size > 0U && !decoded) {
    auto rec = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(rec));
    if (rec.value().type == 161U) {
      auto filter = read_auto_filter_block(cursor, rec.value().payload);
      ASSERT_TRUE(static_cast<bool>(filter)) << filter.error().message;
      decoded = std::move(filter.value());
    }
  }
  ASSERT_TRUE(decoded.has_value());
  auto xlsx = read_ooxml(SpanOf(ReadFixtureFile(dir + "autofilter_0_base.xlsx")));
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  const AutoFilter* expected = xlsx.value().workbook.sheet(0).auto_filter();
  ASSERT_NE(expected, nullptr);
  AutoFilter twin = *expected;
  // Namespace and uid attributes have no record field.
  twin.extra_attrs.clear();
  ASSERT_TRUE(twin.sort.has_value());
  twin.sort->extra_attrs.clear();
  EXPECT_EQ(serialize_auto_filter(*decoded), serialize_auto_filter(twin));
  // And the records written back from the twin decode to the same model.
  std::vector<std::uint8_t> written;
  EXPECT_EQ(emit_auto_filter_block(written, twin), 0U);
  ByteSpan again = SpanOf(written);
  auto begin = read_record(again);
  ASSERT_TRUE(static_cast<bool>(begin));
  auto reread = read_auto_filter_block(again, begin.value().payload);
  ASSERT_TRUE(static_cast<bool>(reread));
  EXPECT_EQ(serialize_auto_filter(reread.value()), serialize_auto_filter(twin));
}

TEST(XlsbTableFidelity, EveryFilterCriterionKindRoundTrips) {
  const char* const kFilters[] = {
      "<autoFilter ref=\"A1:F9\"><filterColumn colId=\"0\"><filters blank=\"1\"><filter val=\"a\"/>"
      "<dateGroupItem year=\"2024\" month=\"3\" dateTimeGrouping=\"month\"/></filters></filterColumn></autoFilter>",
      "<autoFilter ref=\"A1:F9\"><filterColumn colId=\"1\"><top10 top=\"0\" percent=\"1\" val=\"10\" "
      "filterVal=\"4\"/></filterColumn></autoFilter>",
      "<autoFilter ref=\"A1:F9\"><filterColumn colId=\"2\"><dynamicFilter type=\"aboveAverage\" val=\"3.5\"/>"
      "</filterColumn><filterColumn colId=\"3\"><dynamicFilter type=\"thisMonth\" val=\"45000\" "
      "maxVal=\"45030\"/></filterColumn><filterColumn colId=\"4\"><dynamicFilter type=\"Q3\"/></filterColumn>"
      "</autoFilter>",
      "<autoFilter ref=\"A1:F9\"><filterColumn colId=\"4\"><colorFilter dxfId=\"2\" cellColor=\"0\"/>"
      "</filterColumn><filterColumn colId=\"5\" showButton=\"0\"><iconFilter iconSet=\"3Flags\" iconId=\"1\"/>"
      "</filterColumn></autoFilter>",
      "<autoFilter ref=\"A1:F9\"><filterColumn colId=\"0\"><customFilters><customFilter operator=\"notEqual\" "
      "val=\"x*\"/></customFilters></filterColumn></autoFilter>",
  };
  for (const char* xml : kFilters) {
    SCOPED_TRACE(xml);
    const AutoFilter model = auto_filter_from_xml(xml);
    ASSERT_FALSE(model.is_opaque());
    std::vector<std::uint8_t> written;
    EXPECT_EQ(emit_auto_filter_block(written, model), 0U);
    ByteSpan cursor = SpanOf(written);
    auto begin = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(begin));
    auto back = read_auto_filter_block(cursor, begin.value().payload);
    ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
    EXPECT_EQ(serialize_auto_filter(back.value()), serialize_auto_filter(model));
    EXPECT_EQ(cursor.size, 0U);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
