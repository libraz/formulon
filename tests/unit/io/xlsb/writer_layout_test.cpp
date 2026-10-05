// XLSB writer round-trip tests: layout, hyperlinks, and styles.

#include <cstddef>
#include <cstdint>
#include <map>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/styles_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "styles.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
namespace formulon {
namespace io {
namespace xlsb {
namespace {

TEST(XlsbWriter, SheetVisibilitySurvivesRoundTrip) {
  // A hidden sheet must stay hidden across write -> read. The Sheet model
  // tracks visibility via `view().tab_hidden`, which the writer maps to
  // BrtBundleSh hsState and the reader maps back.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Visible");
  wb.add_sheet("Hidden");
  wb.sheet(1).mutable_view().tab_hidden = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Workbook& rt = read_or.value().workbook;
  ASSERT_EQ(rt.sheet_count(), 2U);
  EXPECT_FALSE(rt.sheet(0).view().tab_hidden);
  EXPECT_TRUE(rt.sheet(1).view().tab_hidden);
}

TEST(XlsbWriter, Date1904SurvivesRoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Dates");
  wb.set_date1904(true);

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  EXPECT_TRUE(read_or.value().workbook.date1904());
}

TEST(XlsbWriter, RowAndColumnLayoutSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Layout"));
  sheet.mutable_layout().columns.push_back(ColumnLayout{1U, 3U, 17.25, true, 2U});
  sheet.mutable_layout().row_overrides.push_back(RowLayout{4U, 28.5, true, 3U, true, true});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const SheetLayout& layout = read_or.value().workbook.sheet(0).layout();
  ASSERT_EQ(layout.columns.size(), 1U);
  EXPECT_EQ(layout.columns[0].first, 1U);
  EXPECT_EQ(layout.columns[0].last, 3U);
  EXPECT_DOUBLE_EQ(layout.columns[0].width, 17.25);
  EXPECT_TRUE(layout.columns[0].has_width);
  // BrtColInfo always carries ixfe, so XLSB canonicalizes a width-only
  // aggregate span to effective style 0 on read.
  EXPECT_TRUE(layout.columns[0].has_style);
  EXPECT_EQ(layout.columns[0].style_xf, 0U);
  EXPECT_TRUE(layout.columns[0].hidden);
  EXPECT_EQ(layout.columns[0].outline_level, 2U);
  ASSERT_EQ(layout.row_overrides.size(), 1U);
  EXPECT_EQ(layout.row_overrides[0].row, 4U);
  EXPECT_DOUBLE_EQ(layout.row_overrides[0].height, 28.5);
  EXPECT_TRUE(layout.row_overrides[0].hidden);
  EXPECT_EQ(layout.row_overrides[0].outline_level, 3U);
}

TEST(XlsbWriter, RowStyleFlagCarriesExplicitZeroAndNonZeroStyles) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("RowStyles"));
  RowLayout style_zero;
  style_zero.row = 2U;
  style_zero.has_style = true;
  style_zero.style_xf = 0U;
  RowLayout style_one;
  style_one.row = 3U;
  style_one.has_style = true;
  style_one.style_xf = 1U;
  sheet.mutable_layout().row_overrides = {style_zero, style_one};

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  std::map<std::uint32_t, std::pair<std::uint32_t, std::uint8_t>> raw;
  ByteSpan cursor = SpanOf(sheet_or.value());
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto row_or = read_u32(payload);
    auto style_or = read_u32(payload);
    ASSERT_TRUE(row_or && style_or);
    ASSERT_TRUE(read_u16(payload));
    ASSERT_TRUE(read_u8(payload));
    auto flags_or = read_u8(payload);
    ASSERT_TRUE(read_u8(payload));
    ASSERT_TRUE(flags_or);
    raw[row_or.value()] = {style_or.value(), flags_or.value()};
  }
  ASSERT_EQ(raw.size(), 2U);
  EXPECT_EQ(raw.at(2U).first, 0U);
  EXPECT_NE(raw.at(2U).second & 0x40U, 0U);
  EXPECT_EQ(raw.at(3U).first, 1U);
  EXPECT_NE(raw.at(3U).second & 0x40U, 0U);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const auto& rows = read_or.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 2U);
  EXPECT_TRUE(rows[0].has_style);
  EXPECT_TRUE(rows[1].has_style);
  const auto find_row = [&rows](std::uint32_t row) -> const RowLayout* {
    for (const RowLayout& candidate : rows) {
      if (candidate.row == row) {
        return &candidate;
      }
    }
    return nullptr;
  };
  const RowLayout* loaded_zero = find_row(2U);
  const RowLayout* loaded_one = find_row(3U);
  ASSERT_NE(loaded_zero, nullptr);
  ASSERT_NE(loaded_one, nullptr);
  EXPECT_TRUE(loaded_zero->has_style);
  EXPECT_EQ(loaded_zero->style_xf, 0U);
  EXPECT_TRUE(loaded_one->has_style);
  EXPECT_EQ(loaded_one->style_xf, 1U);
}

TEST(XlsbWriter, MergedRangesSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Merged"));
  sheet.set_cell_value(0U, 0U, Value::text("title"));
  sheet.mutable_merges().push_back(MergeRange{0U, 0U, 1U, 2U});
  sheet.mutable_merges().push_back(MergeRange{4U, 3U, 4U, 5U});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const std::vector<MergeRange>& merges = read_or.value().workbook.sheet(0).merges();
  ASSERT_EQ(merges.size(), 2U);
  EXPECT_EQ(merges[0].first_row, 0U);
  EXPECT_EQ(merges[0].first_col, 0U);
  EXPECT_EQ(merges[0].last_row, 1U);
  EXPECT_EQ(merges[0].last_col, 2U);
  EXPECT_EQ(merges[1].first_row, 4U);
  EXPECT_EQ(merges[1].first_col, 3U);
  EXPECT_EQ(merges[1].last_row, 4U);
  EXPECT_EQ(merges[1].last_col, 5U);
}

TEST(XlsbWriter, SheetViewAndFrozenPanesSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("View"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  SheetView& view = sheet.mutable_view();
  view.zoom_scale = 135U;
  view.freeze_rows = 3U;
  view.freeze_cols = 2U;
  view.show_grid_lines = false;
  view.show_row_col_headers = false;
  view.show_zeros = false;
  view.right_to_left = true;
  view.tab_selected = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan records = SpanOf(sheet_or.value());
  bool saw_pane = false;
  while (records.size != 0U) {
    auto rec_or = read_record(records);
    ASSERT_TRUE(static_cast<bool>(rec_or));
    if (rec_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtPane)) {
      continue;
    }
    saw_pane = true;
    ByteSpan pane = rec_or.value().payload;
    ASSERT_EQ(pane.size, 29U);
    pane.data += 16U;  // two Xnum frozen row/column counts
    pane.size -= 16U;
    auto top_row_or = read_u32(pane);
    auto left_col_or = read_u32(pane);
    auto active_pane_or = read_u32(pane);
    auto flags_or = read_u8(pane);
    ASSERT_TRUE(top_row_or && left_col_or && active_pane_or && flags_or);
    EXPECT_EQ(top_row_or.value(), 3U);
    EXPECT_EQ(left_col_or.value(), 2U);
    EXPECT_EQ(active_pane_or.value(), 0U);
    EXPECT_EQ(flags_or.value(), 0x02U);
  }
  EXPECT_TRUE(saw_pane);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const SheetView& rt = read_or.value().workbook.sheet(0).view();
  EXPECT_EQ(rt.zoom_scale, 135U);
  EXPECT_EQ(rt.freeze_rows, 3U);
  EXPECT_EQ(rt.freeze_cols, 2U);
  EXPECT_FALSE(rt.show_grid_lines);
  EXPECT_FALSE(rt.show_row_col_headers);
  EXPECT_FALSE(rt.show_zeros);
  EXPECT_TRUE(rt.right_to_left);
  EXPECT_TRUE(rt.tab_selected);
}

TEST(XlsbWriter, ReportsDeferredSheetFeatures) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Deferred"));
  Hyperlink hyperlink;
  hyperlink.target = "https://example.com";
  sheet.mutable_hyperlinks().push_back(std::move(hyperlink));
  sheet.mutable_validations().push_back(DataValidation{});
  sheet.set_auto_filter_xml("<autoFilter ref=\"A1:B2\"/>");
  sheet.mutable_view().freeze_rows = 1U;

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  // Auto-filter state remains deferred; hyperlinks and validations are
  // written from the model and therefore no longer inflate this counter.
  EXPECT_EQ(write_or.value().diagnostics.deferred_feature_count, 1U);
}

TEST(XlsbWriter, RejectsInvalidHyperlinkRectangle) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("InvalidHyperlink"));
  Hyperlink inverted;
  inverted.row = 4U;
  inverted.col = 5U;
  inverted.last_row = 3U;
  inverted.last_col = 6U;
  inverted.target = "https://invalid.example";
  sheet.mutable_hyperlinks().push_back(inverted);

  auto inverted_write = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(inverted_write));
  EXPECT_EQ(inverted_write.error().code, FormulonErrorCode::kInvalidArgument);

  sheet.mutable_hyperlinks().clear();
  Hyperlink out_of_grid;
  out_of_grid.row = 0U;
  out_of_grid.col = 0U;
  out_of_grid.last_row = Sheet::kMaxRows;
  out_of_grid.last_col = 0U;
  out_of_grid.target = "https://invalid.example";
  sheet.mutable_hyperlinks().push_back(out_of_grid);
  auto out_of_grid_write = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(out_of_grid_write));
  EXPECT_EQ(out_of_grid_write.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(XlsbWriter, RoundTripsExternalAndInternalHyperlinkRectangles) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Hyperlinks"));
  Hyperlink external;
  external.row = 1U;
  external.col = 2U;
  external.last_row = 3U;
  external.last_col = 4U;
  external.target = "https://external.example/book.xlsx";
  external.location = "#Sheet2!A1";
  external.tooltip = "external";
  external.display = "Open";
  sheet.mutable_hyperlinks().push_back(external);
  Hyperlink internal;
  internal.row = 6U;
  internal.col = 7U;
  internal.last_row = 8U;
  internal.last_col = 9U;
  internal.location = "Sheet1!A1";
  internal.tooltip = "internal";
  internal.display = "Jump";
  sheet.mutable_hyperlinks().push_back(internal);

  auto first_write = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(first_write)) << first_write.error().message << " | " << first_write.error().context;
  auto first_read = read_xlsb(SpanOf(first_write.value()));
  ASSERT_TRUE(static_cast<bool>(first_read)) << first_read.error().message << " | " << first_read.error().context;
  Workbook edited = std::move(first_read.value().workbook);
  ASSERT_EQ(edited.sheet(0).hyperlinks().size(), 2U);
  edited.sheet(0).mutable_hyperlinks()[0].row = 2U;
  edited.sheet(0).mutable_hyperlinks()[0].last_row = 4U;

  auto second_write = write_xlsb(edited);
  ASSERT_TRUE(static_cast<bool>(second_write)) << second_write.error().message << " | " << second_write.error().context;
  auto second_read = read_xlsb(SpanOf(second_write.value()));
  ASSERT_TRUE(static_cast<bool>(second_read)) << second_read.error().message << " | " << second_read.error().context;
  const auto& hyperlinks = second_read.value().workbook.sheet(0).hyperlinks();
  ASSERT_EQ(hyperlinks.size(), 2U);
  EXPECT_EQ(hyperlinks[0].row, 2U);
  EXPECT_EQ(hyperlinks[0].col, 2U);
  EXPECT_EQ(hyperlinks[0].last_row, 4U);
  EXPECT_EQ(hyperlinks[0].last_col, 4U);
  EXPECT_EQ(hyperlinks[0].target, "https://external.example/book.xlsx");
  EXPECT_FALSE(hyperlinks[0].rid.empty());
  EXPECT_EQ(hyperlinks[1].row, 6U);
  EXPECT_EQ(hyperlinks[1].col, 7U);
  EXPECT_EQ(hyperlinks[1].last_row, 8U);
  EXPECT_EQ(hyperlinks[1].last_col, 9U);
  EXPECT_TRUE(hyperlinks[1].rid.empty());
  EXPECT_EQ(hyperlinks[1].location, "Sheet1!A1");
}

TEST(XlsbWriter, ReusesSharedSourceRelationshipIdForMatchingTargets) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("SharedRids"));
  Hyperlink first;
  first.row = 1U;
  first.col = 1U;
  first.last_row = 1U;
  first.last_col = 2U;
  first.target = "https://shared.example/target";
  first.rid = "rIdShared";
  sheet.mutable_hyperlinks().push_back(first);
  Hyperlink second;
  second.row = 3U;
  second.col = 1U;
  second.last_row = 3U;
  second.last_col = 2U;
  second.target = "https://shared.example/target";
  second.rid = "rIdShared";
  sheet.mutable_hyperlinks().push_back(second);

  auto write_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value()))));
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  const std::string id_token = "Id=\"rIdShared\"";
  const std::size_t first_id = rels.find(id_token);
  ASSERT_NE(first_id, std::string::npos) << rels;
  EXPECT_EQ(rels.find(id_token, first_id + id_token.size()), std::string::npos) << rels;

  auto read_or = read_xlsb(SpanOf(write_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const auto& hyperlinks = read_or.value().workbook.sheet(0).hyperlinks();
  ASSERT_EQ(hyperlinks.size(), 2U);
  EXPECT_EQ(hyperlinks[0].rid, "rIdShared");
  EXPECT_EQ(hyperlinks[1].rid, "rIdShared");
  EXPECT_EQ(hyperlinks[0].target, hyperlinks[1].target);
}

TEST(XlsbWriter, GeneratesStylesPartForModelledStyles) {
  // XLSX/native workbooks do not have a raw styles.bin passthrough part.  The
  // XLSB writer must therefore materialise the modelled table and expose it
  // through both the content-type override and workbook relationship.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Styled"));
  sheet.set_cell_value(0U, 0U, Value::number(12.5));

  StylesTable styles;
  FontRecord font;
  font.name = "Aptos";
  font.size = 12.0;
  font.bold = true;
  font.color_argb = 0xFF112233U;
  styles.fonts.push_back(font);
  FillRecord fill;
  fill.pattern = 1U;
  fill.fg_argb = 0xFFFFFF00U;
  styles.fills.push_back(fill);
  styles.borders.push_back(BorderRecord{});
  styles.num_fmt_strings.push_back("0.000");
  styles.num_fmts.push_back(NumFmtRecord{164U, 0U});
  CellXf named;
  named.vertical_align = 2U;
  styles.cell_style_xfs.push_back(named);
  CellXf xf;
  xf.font_index = 0U;
  xf.fill_index = 0U;
  xf.border_index = 0U;
  xf.num_fmt_id = 164U;
  xf.xf_id = 0U;
  xf.apply_number_format = true;
  xf.apply_font = true;
  xf.apply_fill = true;
  styles.cell_xfs.push_back(xf);
  wb.set_styles(std::move(styles));
  sheet.set_cell_xf_index(0U, 0U, 0U);

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/styles.bin"));
  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("relationships/styles"), std::string::npos);
  EXPECT_NE(rels.find("Target=\"styles.bin\""), std::string::npos);
  auto types_or = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(types_or));
  const std::string types(types_or.value().begin(), types_or.value().end());
  EXPECT_NE(types.find("/xl/styles.bin"), std::string::npos);
  EXPECT_NE(types.find("application/vnd.ms-excel.styles"), std::string::npos);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const StylesTable& rt = read_or.value().workbook.styles();
  ASSERT_EQ(rt.num_fmts.size(), 1U);
  EXPECT_EQ(rt.num_fmts[0].id, 164U);
  ASSERT_EQ(rt.num_fmt_strings.size(), 1U);
  EXPECT_EQ(rt.num_fmt_strings[0], "0.000");
  ASSERT_EQ(rt.cell_xfs.size(), 1U);
  EXPECT_EQ(rt.cell_xfs[0].num_fmt_id, 164U);
  EXPECT_EQ(rt.cell_xfs[0].font_index, 0U);
  EXPECT_EQ(rt.cell_xfs[0].fill_index, 0U);
}

TEST(XlsbWriter, StylesColorsUseTheBrtColorLowValidityBit) {
  StylesTable styles;
  FontRecord font;
  font.name = "Aptos";
  font.color_argb = 0xFF112233U;
  styles.fonts.push_back(font);
  const std::vector<std::uint8_t> bytes = write_styles_bin(styles);

  ByteSpan cursor = SpanOf(bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtFont)) {
      continue;
    }
    const ByteSpan payload = record_or.value().payload;
    ASSERT_GE(payload.size, 20U);
    // u16 height/flags/weight/vertAlign + underline/family/charset/reserved
    // precede BrtColor. RGB is XColorType 2 with fValidRGB in bit 0.
    EXPECT_EQ(payload.data[12], 0x05U);
    EXPECT_EQ(payload.data[16], 0x11U);
    EXPECT_EQ(payload.data[17], 0x22U);
    EXPECT_EQ(payload.data[18], 0x33U);
    EXPECT_EQ(payload.data[19], 0xFFU);
    return;
  }
  FAIL() << "missing BrtFont";
}

TEST(XlsbWriter, PreservesRawStylesPartFromExistingXlsb) {
  // Generated styles are for XLSX/native input only.  When an XLSB reader has
  // retained an opaque style part, that byte stream remains authoritative so
  // unmodelled binary formatting extensions cannot be erased on save.
  Workbook wb = Workbook::create_empty();
  wb.sheet(wb.add_sheet("S")).set_cell_value(0U, 0U, Value::number(1.0));
  StylesTable raw_table;
  raw_table.num_fmt_strings.push_back("0.0000");
  raw_table.num_fmts.push_back(NumFmtRecord{164U, 0U});
  const std::vector<std::uint8_t> raw_bytes = write_styles_bin(raw_table);
  PassthroughPart raw_part;
  raw_part.path = "xl/styles.bin";
  raw_part.content_type = "application/vnd.ms-excel.styles";
  raw_part.bytes = raw_bytes;
  wb.set_passthrough_parts({raw_part});

  // Deliberately make the in-memory table differ from the opaque source.  If
  // the writer regenerated it, the byte comparison below would fail.
  StylesTable changed;
  changed.num_fmt_strings.push_back("0.0%");
  changed.num_fmts.push_back(NumFmtRecord{165U, 0U});
  wb.set_styles(std::move(changed));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto styles_or = zip.read_entry("xl/styles.bin");
  ASSERT_TRUE(static_cast<bool>(styles_or));
  EXPECT_EQ(styles_or.value(), raw_bytes);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
