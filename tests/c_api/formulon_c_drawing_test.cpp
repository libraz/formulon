//
// Stable C ABI tests for drawing images: probing, listing next to other
// drawing objects, reading, inserting, removing, and a save and reload.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_image_info) == 12U, "fm_image_info ABI layout changed");
static_assert(sizeof(fm_drawing_object) == (kWasm32 ? 96U : 112U), "fm_drawing_object ABI layout changed");
static_assert(sizeof(fm_image_insert) == (kWasm32 ? 56U : 64U), "fm_image_insert ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr fm_status_t kIoImageUnsupported = static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoImageUnsupported);
constexpr fm_status_t kIoDrawingUnparseable =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kIoDrawingUnparseable);

constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";
constexpr const char* kRelDrawing = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

Bytes Png(std::uint32_t w, std::uint32_t h) {
  Bytes b = {0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 'I', 'H', 'D', 'R'};
  for (std::uint32_t v : {w, h}) {
    for (int shift = 24; shift >= 0; shift -= 8) {
      b.push_back(static_cast<std::uint8_t>(v >> shift));
    }
  }
  b.insert(b.end(), {8, 6, 0, 0, 0, 0x1F, 0x15, 0xC4, 0x89});
  return b;
}

Bytes Jpeg(std::uint16_t w, std::uint16_t h) {
  return {0xFF,
          0xD8,
          0xFF,
          0xC0,
          0x00,
          0x11,
          0x08,
          static_cast<std::uint8_t>(h >> 8),
          static_cast<std::uint8_t>(h),
          static_cast<std::uint8_t>(w >> 8),
          static_cast<std::uint8_t>(w),
          0x03,
          0xFF,
          0xD9};
}

Bytes Gif(std::uint16_t w, std::uint16_t h) {
  return {'G',
          'I',
          'F',
          '8',
          '9',
          'a',
          static_cast<std::uint8_t>(w),
          static_cast<std::uint8_t>(w >> 8),
          static_cast<std::uint8_t>(h),
          static_cast<std::uint8_t>(h >> 8),
          0,
          0,
          0};
}

Bytes Bmp(std::int32_t w, std::int32_t h) {
  Bytes b(26, 0);
  b[0] = 'B';
  b[1] = 'M';
  b[14] = 12;  // BITMAPCOREHEADER
  b[18] = static_cast<std::uint8_t>(w);
  b[19] = static_cast<std::uint8_t>(w >> 8);
  b[20] = static_cast<std::uint8_t>(h);
  b[21] = static_cast<std::uint8_t>(h >> 8);
  b[22] = 1;
  b[24] = 24;
  return b;
}

fm_image_insert Defaults() {
  fm_image_insert opts{};
  opts.anchor_kind = FM_ANCHOR_KIND_ONE_CELL;
  opts.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  return opts;
}

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, &buf.data, &buf.len), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_last_error_message();
}

/// Fills the first sheet's drawing with a chart anchor and a shape anchor.
void AddChartDrawing(fm_workbook_t* wb) {
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">"
      "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>"
      "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>14</xdr:row>"
      "<xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr>"
      "<xdr:cNvPr id=\"2\" name=\"Chart 1\"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm>"
      "<a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/></xdr:xfrm><a:graphic>"
      "<a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
      "<c:chart xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" r:id=\"rId1\"/>"
      "</a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>";
  const std::string rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart\" "
      "Target=\"../charts/chart1.xml\"/></Relationships>";
  const std::string chart = "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"/>";
  formulon::Workbook& model = wb->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart("xl/drawings/_rels/drawing1.xml.rels", "", Bytes(rels.begin(), rels.end())))));
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(formulon::PassthroughPart(
      "xl/charts/chart1.xml", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
      Bytes(chart.begin(), chart.end())))));
  model.sheet(0).set_drawing_rel_target(kDrawingPath);
}

TEST(FormulonCApiDrawing, EnumOrdinals) {
  EXPECT_EQ(FM_IMAGE_FORMAT_UNKNOWN, 0);
  EXPECT_EQ(FM_IMAGE_FORMAT_PNG, 1);
  EXPECT_EQ(FM_IMAGE_FORMAT_JPEG, 2);
  EXPECT_EQ(FM_IMAGE_FORMAT_GIF, 3);
  EXPECT_EQ(FM_IMAGE_FORMAT_BMP, 4);
  EXPECT_EQ(FM_DRAWING_OBJECT_PICTURE, 0);
  EXPECT_EQ(FM_DRAWING_OBJECT_SHAPE, 1);
  EXPECT_EQ(FM_DRAWING_OBJECT_CHART, 2);
  EXPECT_EQ(FM_DRAWING_OBJECT_GROUP, 3);
  EXPECT_EQ(FM_DRAWING_OBJECT_CONNECTOR, 4);
  EXPECT_EQ(FM_DRAWING_OBJECT_GRAPHIC_FRAME, 5);
  EXPECT_EQ(FM_DRAWING_OBJECT_OTHER, 6);
  EXPECT_EQ(FM_ANCHOR_KIND_ONE_CELL, 0);
  EXPECT_EQ(FM_ANCHOR_KIND_TWO_CELL, 1);
  EXPECT_EQ(FM_ANCHOR_KIND_ABSOLUTE, 2);
  EXPECT_EQ(FM_ANCHOR_EDIT_AS_TWO_CELL, 0);
  EXPECT_EQ(FM_ANCHOR_EDIT_AS_ONE_CELL, 1);
  EXPECT_EQ(FM_ANCHOR_EDIT_AS_ABSOLUTE, 2);
}

TEST(FormulonCApiDrawing, ProbeEachFormat) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  struct Case {
    Bytes bytes;
    int32_t format;
    uint32_t w, h;
  };
  const Case cases[] = {{Png(16, 8), FM_IMAGE_FORMAT_PNG, 16, 8},
                        {Jpeg(30, 20), FM_IMAGE_FORMAT_JPEG, 30, 20},
                        {Gif(5, 7), FM_IMAGE_FORMAT_GIF, 5, 7},
                        {Bmp(9, 4), FM_IMAGE_FORMAT_BMP, 9, 4}};
  for (const Case& c : cases) {
    fm_image_info info{};
    ASSERT_EQ(fm_workbook_probe_image(wb.handle, c.bytes.data(), c.bytes.size(), &info), 0) << fm_last_error_message();
    EXPECT_EQ(info.format, c.format);
    EXPECT_EQ(info.px_width, c.w);
    EXPECT_EQ(info.px_height, c.h);
  }
}

TEST(FormulonCApiDrawing, ProbeRejectsUnknownBytesAndNulls) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const Bytes junk = {'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'};
  fm_image_info info{};
  EXPECT_EQ(fm_workbook_probe_image(wb.handle, junk.data(), junk.size(), &info), kIoImageUnsupported);
  EXPECT_EQ(fm_workbook_probe_image(wb.handle, junk.data(), 0, &info), kIoImageUnsupported);
  EXPECT_EQ(fm_workbook_probe_image(wb.handle, nullptr, 0, &info), kBindingNullPointer);
  EXPECT_EQ(fm_workbook_probe_image(wb.handle, junk.data(), junk.size(), nullptr), kBindingNullPointer);
}

TEST(FormulonCApiDrawing, InsertListGetRemoveNextToChart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  AddChartDrawing(wb.handle);
  const Bytes png = Png(16, 8);
  fm_image_insert opts = Defaults();
  opts.name = "Logo";
  opts.descr = "alt";
  opts.row = 3;
  opts.col = 2;
  opts.row_off_emu = 1000;
  opts.col_off_emu = 2000;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), 0) << fm_last_error_message();
  EXPECT_EQ(id, 3U);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_drawing_object_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_drawing_object chart{};
  ASSERT_EQ(fm_sheet_drawing_object_at(wb.handle, 0, 0, &chart), 0);
  EXPECT_EQ(chart.kind, FM_DRAWING_OBJECT_CHART);
  EXPECT_EQ(chart.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(chart.to_col, 6U);
  EXPECT_EQ(chart.to_row, 14U);
  EXPECT_EQ(chart.image_format, FM_IMAGE_FORMAT_UNKNOWN);
  EXPECT_STREQ(chart.name, "Chart 1");
  fm_drawing_object pic{};
  ASSERT_EQ(fm_sheet_drawing_object_at(wb.handle, 0, 1, &pic), 0);
  EXPECT_EQ(pic.object_id, id);
  EXPECT_EQ(pic.kind, FM_DRAWING_OBJECT_PICTURE);
  EXPECT_EQ(pic.anchor_kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(pic.from_row, 3U);
  EXPECT_EQ(pic.from_col, 2U);
  EXPECT_EQ(pic.from_row_off, 1000);
  EXPECT_EQ(pic.from_col_off, 2000);
  EXPECT_EQ(pic.cx, 16 * 9525);
  EXPECT_EQ(pic.cy, 8 * 9525);
  EXPECT_EQ(pic.image_format, FM_IMAGE_FORMAT_PNG);
  EXPECT_STREQ(pic.name, "Logo");
  EXPECT_STREQ(pic.descr, "alt");
  EXPECT_NE(std::string(pic.media_path).find("xl/media/image"), std::string::npos);
  // The previous record's strings live until the next successful read.
  fm_drawing_object again{};
  ASSERT_EQ(fm_sheet_drawing_object_at(wb.handle, 0, 1, &again), 0);
  EXPECT_STREQ(again.name, "Logo");

  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info info{};
  ASSERT_EQ(fm_sheet_get_image(wb.handle, 0, id, &got, &got_len, &info), 0) << fm_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), png);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(info.px_width, 16U);
  EXPECT_EQ(info.px_height, 8U);
  // A chart is not a picture; an unknown id is not found.
  EXPECT_EQ(fm_sheet_get_image(wb.handle, 0, 2, &got, &got_len, &info), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_image(wb.handle, 0, 99, &got, &got_len, &info), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_image(wb.handle, 0, id, nullptr, &got_len, &info), kBindingNullPointer);

  EXPECT_EQ(fm_sheet_remove_image(wb.handle, 0, 2), kInvalidArgument);
  ASSERT_EQ(fm_sheet_remove_image(wb.handle, 0, id), 0) << fm_last_error_message();
  ASSERT_EQ(fm_sheet_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  EXPECT_EQ(fm_sheet_get_image(wb.handle, 0, id, &got, &got_len, &info), kInvalidArgument);
  EXPECT_EQ(fm_sheet_drawing_object_at(wb.handle, 0, 1, &pic), kInvalidArgument);
}

TEST(FormulonCApiDrawing, InsertSaveReloadKeepsBytes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const Bytes png = Png(16, 8);
  const Bytes jpeg = Jpeg(30, 20);
  fm_image_insert opts = Defaults();
  opts.row = 1;
  uint32_t first = 0;
  uint32_t second = 0;
  ASSERT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &first), 0) << fm_last_error_message();
  opts.anchor_kind = FM_ANCHOR_KIND_TWO_CELL;
  opts.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  opts.row = 5;
  ASSERT_EQ(fm_sheet_insert_image(wb.handle, 0, jpeg.data(), jpeg.size(), &opts, &second), 0)
      << fm_last_error_message();
  EXPECT_NE(first, second);

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_drawing_object_count(reloaded.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info info{};
  ASSERT_EQ(fm_sheet_get_image(reloaded.handle, 0, first, &got, &got_len, &info), 0) << fm_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), png);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  ASSERT_EQ(fm_sheet_get_image(reloaded.handle, 0, second, &got, &got_len, &info), 0) << fm_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), jpeg);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_JPEG);
  fm_drawing_object obj{};
  ASSERT_EQ(fm_sheet_drawing_object_at(reloaded.handle, 0, 1, &obj), 0);
  EXPECT_EQ(obj.anchor_kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(obj.from_row, 5U);
}

TEST(FormulonCApiDrawing, InsertRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  const Bytes junk = {1, 2, 3, 4};
  fm_image_insert opts = Defaults();
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, junk.data(), junk.size(), &opts, &id), kIoImageUnsupported);
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, nullptr, 0, &opts, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 7, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts.anchor_kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts.anchor_kind = 3;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts = Defaults();
  opts.edit_as = -1;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts = Defaults();
  opts.width_emu = -1;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiDrawing, XlsbRetainedDrawingRefusesInsert) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"></xdr:wsDr>";
  formulon::Workbook& model = wb.handle->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  model.sheet(0).set_unknown_relationships({formulon::UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  const Bytes png = Png(4, 4);
  fm_image_insert opts = Defaults();
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_insert_image(wb.handle, 0, png.data(), png.size(), &opts, &id), kIoDrawingUnparseable);
  EXPECT_EQ(fm_sheet_remove_image(wb.handle, 0, 2), kIoDrawingUnparseable);
  EXPECT_EQ(model.passthrough_parts().size(), 1U);
}

}  // namespace
