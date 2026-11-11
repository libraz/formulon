//
// Stable C ABI tests for drawing objects: probing, reading and inserting
// pictures next to other objects; moving, reordering, capturing, restoring and
// removing pictures and charts; and a save and reload.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_image_info_t) == 16U, "fm_image_info_t ABI layout changed");
static_assert(sizeof(fm_drawing_object_info_t) == (kWasm32 ? 104U : 112U),
              "fm_drawing_object_info_t ABI layout changed");
static_assert(sizeof(fm_drawing_placement_spec_t) == (kWasm32 ? 88U : 104U),
              "fm_drawing_placement_spec_t ABI layout changed");
static_assert(sizeof(fm_drawing_anchor_t) == 72U, "fm_drawing_anchor_t ABI layout changed");
static_assert(sizeof(fm_drawing_anchor_patch_t) == 144U, "fm_drawing_anchor_patch_t ABI layout changed");
static_assert(sizeof(fm_drawing_patch_t) == (kWasm32 ? 176U : 184U), "fm_drawing_patch_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingUtf8EncodingError =
    static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingUtf8EncodingError);
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

fm_drawing_placement_spec_t Defaults() {
  fm_drawing_placement_spec_t opts = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  opts.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  opts.anchor.edit_as = FM_ANCHOR_EDIT_AS_TWO_CELL;
  return opts;
}

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  BufferGuard buf;
  ASSERT_EQ(fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &buf.data, &buf.len), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(buf.data, buf.len, &out->handle), 0) << fm_get_last_error_message();
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
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_PICTURE, 0);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_SHAPE, 1);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_CHART, 2);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_GROUP, 3);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_CONNECTOR, 4);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_GRAPHIC_FRAME, 5);
  EXPECT_EQ(FM_DRAWING_OBJECT_KIND_OTHER, 6);
  EXPECT_EQ(FM_ANCHOR_KIND_DEFAULT, 0);
  EXPECT_EQ(FM_ANCHOR_KIND_ONE_CELL, 1);
  EXPECT_EQ(FM_ANCHOR_KIND_TWO_CELL, 2);
  EXPECT_EQ(FM_ANCHOR_KIND_ABSOLUTE, 3);
  EXPECT_EQ(FM_ANCHOR_EDIT_AS_DEFAULT, 0);
  // The anchor family: an edit mode takes its anchor kind's value.
  EXPECT_EQ(static_cast<int>(FM_ANCHOR_EDIT_AS_ONE_CELL), static_cast<int>(FM_ANCHOR_KIND_ONE_CELL));
  EXPECT_EQ(static_cast<int>(FM_ANCHOR_EDIT_AS_TWO_CELL), static_cast<int>(FM_ANCHOR_KIND_TWO_CELL));
  EXPECT_EQ(static_cast<int>(FM_ANCHOR_EDIT_AS_ABSOLUTE), static_cast<int>(FM_ANCHOR_KIND_ABSOLUTE));
}

TEST(FormulonCApiDrawing, ProbeEachFormat) {
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
    fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
    ASSERT_EQ(fm_get_image_info(c.bytes.data(), c.bytes.size(), &info), 0) << fm_get_last_error_message();
    EXPECT_EQ(info.format, c.format);
    EXPECT_EQ(info.width_px, c.w);
    EXPECT_EQ(info.height_px, c.h);
  }
}

TEST(FormulonCApiDrawing, ProbeRejectsUnknownBytesAndNulls) {
  const Bytes junk = {'n', 'o', 't', ' ', 'a', 'n', ' ', 'i', 'm', 'a', 'g', 'e'};
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  EXPECT_EQ(fm_get_image_info(junk.data(), junk.size(), &info), kIoImageUnsupported);
  EXPECT_EQ(fm_get_image_info(junk.data(), 0, &info), kIoImageUnsupported);
  EXPECT_EQ(fm_get_image_info(nullptr, 0, &info), kBindingNullPointer);
  EXPECT_EQ(fm_get_image_info(junk.data(), junk.size(), nullptr), kBindingNullPointer);
}

TEST(FormulonCApiDrawing, InsertListGetRemoveNextToChart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);
  const Bytes png = Png(16, 8);
  fm_drawing_placement_spec_t opts = Defaults();
  opts.name = "Logo";
  opts.description = "alt";
  opts.anchor.from_row = 3;
  opts.anchor.from_col = 2;
  opts.anchor.from_row_offset_emu = 1000;
  opts.anchor.from_col_offset_emu = 2000;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(id, 3U);

  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  fm_drawing_object_info_t chart = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 0, &chart), 0);
  EXPECT_EQ(chart.kind, FM_DRAWING_OBJECT_KIND_CHART);
  EXPECT_EQ(chart.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(chart.anchor.to_col, 6U);
  EXPECT_EQ(chart.anchor.to_row, 14U);
  EXPECT_STREQ(chart.name, "Chart 1");
  EXPECT_EQ(chart.media_path, nullptr);  // Only a picture names a media part.
  fm_drawing_object_info_t pic = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 1, &pic), 0);
  EXPECT_EQ(pic.object_id, id);
  EXPECT_EQ(pic.kind, FM_DRAWING_OBJECT_KIND_PICTURE);
  EXPECT_EQ(pic.anchor.kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(pic.anchor.from_row, 3U);
  EXPECT_EQ(pic.anchor.from_col, 2U);
  EXPECT_EQ(pic.anchor.from_row_offset_emu, 1000);
  EXPECT_EQ(pic.anchor.from_col_offset_emu, 2000);
  EXPECT_EQ(pic.anchor.width_emu, 16 * 9525);
  EXPECT_EQ(pic.anchor.height_emu, 8 * 9525);
  EXPECT_STREQ(pic.name, "Logo");
  EXPECT_STREQ(pic.description, "alt");
  EXPECT_NE(std::string(pic.media_path).find("xl/media/image"), std::string::npos);
  // The previous record's strings live until the next successful read.
  fm_drawing_object_info_t again = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 1, &again), 0);
  EXPECT_STREQ(again.name, "Logo");

  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, id, &present, &got, &got_len, &info), 0) << fm_get_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), png);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  EXPECT_EQ(info.width_px, 16U);
  EXPECT_EQ(info.height_px, 8U);
  EXPECT_EQ(present, 1);
  // A chart is not a picture; an unknown id is absent.
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, 2, &present, &got, &got_len, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(got, nullptr);
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, 99, &present, &got, &got_len, &info), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_sheet_get_picture(wb.handle, 0, id, &present, nullptr, &got_len, &info), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_picture(wb.handle, 0, id, nullptr, &got, &got_len, &info), kBindingNullPointer);

  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, id), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 1U);
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, id, &present, &got, &got_len, &info), 0);
  EXPECT_EQ(present, 0);

  // The chart is a drawing object too: removing it drops its part.
  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 2), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
  EXPECT_EQ(formulon::io::ooxml::find_passthrough_part(wb.handle->workbook(), "xl/charts/chart1.xml"), nullptr);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 2), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 1, &pic), FM_STATUS_INDEX_OUT_OF_RANGE);
}

TEST(FormulonCApiDrawing, AddPictureTakesSpecAndReportsPixelSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(30, 12);
  fm_drawing_placement_spec_t spec = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  spec.anchor.kind = FM_ANCHOR_KIND_ONE_CELL;
  spec.anchor.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  spec.anchor.from_row = 1;
  spec.anchor.from_col = 1;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &spec, &id), 0) << fm_get_last_error_message();

  // A NULL name takes the generated default (numbered one below the object id); a zero size takes the pixel size.
  fm_drawing_object_info_t pic = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 0, &pic), 0);
  EXPECT_EQ(pic.kind, FM_DRAWING_OBJECT_KIND_PICTURE);
  EXPECT_EQ(pic.object_id, id);
  EXPECT_EQ(pic.anchor.width_emu, 30 * 9525);
  EXPECT_EQ(pic.anchor.height_emu, 12 * 9525);
  EXPECT_EQ(std::string(pic.name), "Picture " + std::to_string(id - 1U));

  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  ASSERT_EQ(fm_get_image_info(png.data(), png.size(), &info), 0) << fm_get_last_error_message();
  EXPECT_EQ(info.width_px, 30U);
  EXPECT_EQ(info.height_px, 12U);

  const uint8_t* got = nullptr;
  size_t got_len = 0;
  int32_t present = 0;
  fm_image_info_t stored = FM_STRUCT_INIT(fm_image_info_t);
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, id, &present, &got, &got_len, &stored), 0);
  ASSERT_EQ(present, 1);
  EXPECT_EQ(stored.width_px, 30U);
  EXPECT_EQ(stored.height_px, 12U);
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), nullptr, &id), kBindingNullPointer);
}

TEST(FormulonCApiDrawing, InsertSaveReloadKeepsBytes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(16, 8);
  const Bytes jpeg = Jpeg(30, 20);
  fm_drawing_placement_spec_t opts = Defaults();
  opts.anchor.from_row = 1;
  uint32_t first = 0;
  uint32_t second = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &first), 0)
      << fm_get_last_error_message();
  opts.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  opts.anchor.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  opts.anchor.from_row = 5;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, jpeg.data(), jpeg.size(), &opts, &second), 0)
      << fm_get_last_error_message();
  EXPECT_NE(first, second);

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  size_t count = 0;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(reloaded.handle, 0, &count), 0);
  ASSERT_EQ(count, 2U);
  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_picture(reloaded.handle, 0, first, &present, &got, &got_len, &info), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), png);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_PNG);
  ASSERT_EQ(fm_sheet_get_picture(reloaded.handle, 0, second, &present, &got, &got_len, &info), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), jpeg);
  EXPECT_EQ(info.format, FM_IMAGE_FORMAT_JPEG);
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(reloaded.handle, 0, 1, &obj), 0);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.anchor.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(obj.anchor.from_row, 5U);
}

TEST(FormulonCApiDrawing, InsertRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(4, 4);
  const Bytes junk = {1, 2, 3, 4};
  fm_drawing_placement_spec_t opts = Defaults();
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, junk.data(), junk.size(), &opts, &id), kIoImageUnsupported);
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, nullptr, 0, &opts, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), nullptr, &id), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 7, png.data(), png.size(), &opts, &id), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  opts.anchor.kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts.anchor.kind = 4;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts = Defaults();
  opts.anchor.edit_as = -1;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  opts = Defaults();
  opts.anchor.width_emu = -1;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), kInvalidArgument);
  size_t count = 99;
  ASSERT_EQ(fm_sheet_get_drawing_object_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiDrawing, XlsbRetainedDrawingRefusesInsert) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"></xdr:wsDr>";
  formulon::Workbook& model = wb.handle->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  model.sheet(0).set_unknown_relationships({formulon::UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  const Bytes png = Png(4, 4);
  fm_drawing_placement_spec_t opts = Defaults();
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), kIoDrawingUnparseable);
  EXPECT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 2), kIoDrawingUnparseable);
  EXPECT_EQ(model.passthrough_parts().size(), 1U);
}

}  // namespace

namespace {

// A patch that sets every anchor field a one-cell object holds to its current value.
fm_drawing_patch_t PatchOfCurrentOneCell(const fm_drawing_object_info_t& obj) {
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  fm_drawing_anchor_patch_t& a = patch.anchor;
  a.kind_op = FM_PATCH_OP_SET;
  a.kind = obj.anchor.kind;
  a.edit_as_op = FM_PATCH_OP_SET;
  a.edit_as = obj.anchor.edit_as;
  a.from_row_op = FM_PATCH_OP_SET;
  a.from_row = obj.anchor.from_row;
  a.from_col_op = FM_PATCH_OP_SET;
  a.from_col = obj.anchor.from_col;
  a.from_row_offset_emu_op = FM_PATCH_OP_SET;
  a.from_row_offset_emu = obj.anchor.from_row_offset_emu;
  a.from_col_offset_emu_op = FM_PATCH_OP_SET;
  a.from_col_offset_emu = obj.anchor.from_col_offset_emu;
  a.width_emu_op = FM_PATCH_OP_SET;
  a.width_emu = obj.anchor.width_emu;
  a.height_emu_op = FM_PATCH_OP_SET;
  a.height_emu = obj.anchor.height_emu;
  return patch;
}

fm_drawing_object_info_t ReadObject(const fm_workbook_t* wb, std::uint32_t id) {
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object(wb, 0, id, &present, &obj), 0) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return obj;
}

std::vector<std::uint32_t> ObjectIds(const fm_workbook_t* wb) {
  size_t count = 0;
  EXPECT_EQ(fm_sheet_get_drawing_object_count(wb, 0, &count), 0);
  std::vector<std::uint32_t> ids;
  for (size_t i = 0; i < count; ++i) {
    fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
    EXPECT_EQ(fm_sheet_get_drawing_object_at(wb, 0, i, &obj), 0);
    ids.push_back(obj.object_id);
  }
  return ids;
}

Bytes DrawingBytes(fm_workbook_t* wb) {
  const formulon::PassthroughPart* part = formulon::io::ooxml::find_passthrough_part(wb->workbook(), kDrawingPath);
  return part != nullptr ? part->bytes : Bytes();
}

/// Inserts a PNG at row `row` and returns its object id.
std::uint32_t InsertAt(fm_workbook_t* wb, std::uint32_t row, const char* name) {
  const Bytes png = Png(16, 8);
  fm_drawing_placement_spec_t opts = Defaults();
  opts.name = name;
  opts.anchor.from_row = row;
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb, 0, png.data(), png.size(), &opts, &id), 0) << fm_get_last_error_message();
  return id;
}

}  // namespace

TEST(FormulonCApiDrawing, UpdateMovesAndResizesAOneCellPictureKeepingIdAndName) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = InsertAt(wb.handle, 1, "Logo");

  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  fm_drawing_anchor_patch_t& a = patch.anchor;
  a.from_row_op = a.from_col_op = a.from_row_offset_emu_op = a.from_col_offset_emu_op = FM_PATCH_OP_SET;
  a.from_row = 5;
  a.from_col = 3;
  a.from_row_offset_emu = 1000;
  a.from_col_offset_emu = 2000;
  a.width_emu_op = FM_PATCH_OP_SET;
  a.width_emu = 200000;  // the height is kept
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();

  const fm_drawing_object_info_t obj = ReadObject(wb.handle, id);
  EXPECT_EQ(obj.object_id, id);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_ONE_CELL);
  EXPECT_EQ(obj.anchor.from_row, 5U);
  EXPECT_EQ(obj.anchor.from_col, 3U);
  EXPECT_EQ(obj.anchor.from_row_offset_emu, 1000);
  EXPECT_EQ(obj.anchor.from_col_offset_emu, 2000);
  EXPECT_EQ(obj.anchor.width_emu, 200000);
  EXPECT_EQ(obj.anchor.height_emu, 8 * 9525);
  EXPECT_STREQ(obj.name, "Logo");

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  const fm_drawing_object_info_t again = ReadObject(reloaded.handle, id);
  EXPECT_EQ(again.anchor.from_row, 5U);
  EXPECT_EQ(again.anchor.width_emu, 200000);
}

TEST(FormulonCApiDrawing, UpdateConvertsOneCellToTwoCellFromItsSize) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = InsertAt(wb.handle, 1, "Logo");
  const fm_drawing_object_info_t before = ReadObject(wb.handle, id);

  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.kind_op = FM_PATCH_OP_SET;
  patch.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  patch.anchor.edit_as_op = FM_PATCH_OP_SET;
  patch.anchor.edit_as = FM_ANCHOR_EDIT_AS_ONE_CELL;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();

  const fm_drawing_object_info_t obj = ReadObject(wb.handle, id);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.anchor.edit_as, FM_ANCHOR_EDIT_AS_ONE_CELL);
  EXPECT_EQ(obj.anchor.from_row, before.anchor.from_row);
  EXPECT_EQ(obj.anchor.from_col, before.anchor.from_col);
  EXPECT_GE(obj.anchor.to_col, obj.anchor.from_col);
  EXPECT_GT(obj.anchor.width_emu, 0);
  EXPECT_GT(obj.anchor.height_emu, 0);
}

TEST(FormulonCApiDrawing, UpdateMergesAnchorFieldsAndKeepsTheOtherCorners) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);  // a two-cell chart: from (0, 0) to (14, 6)
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.from_row_op = FM_PATCH_OP_SET;
  patch.anchor.from_row = 3;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, 2, &patch), 0) << fm_get_last_error_message();

  const fm_drawing_object_info_t obj = ReadObject(wb.handle, 2);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.anchor.from_row, 3U);
  EXPECT_EQ(obj.anchor.from_col, 0U);
  EXPECT_EQ(obj.anchor.to_row, 14U);
  EXPECT_EQ(obj.anchor.to_col, 6U);
}

TEST(FormulonCApiDrawing, UpdateWithCurrentValuesWritesNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = InsertAt(wb.handle, 2, "Logo");
  const Bytes before = DrawingBytes(wb.handle);
  ASSERT_FALSE(before.empty());
  const fm_drawing_patch_t same = PatchOfCurrentOneCell(ReadObject(wb.handle, id));
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &same), 0) << fm_get_last_error_message();
  EXPECT_EQ(DrawingBytes(wb.handle), before);
}

TEST(FormulonCApiDrawing, UpdateRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = InsertAt(wb.handle, 0, "Logo");
  const Bytes before = DrawingBytes(wb.handle);
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.hidden_op = FM_PATCH_OP_SET;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_drawing(nullptr, 0, id, &patch), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 7, id, &patch), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, 99, &patch), FM_STATUS_NOT_FOUND);

  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.kind_op = FM_PATCH_OP_SET;
  patch.anchor.kind = FM_ANCHOR_KIND_ABSOLUTE;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch.anchor.kind = FM_ANCHOR_KIND_DEFAULT;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch.anchor.kind = 4;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.edit_as_op = FM_PATCH_OP_SET;
  patch.anchor.edit_as = -1;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.width_emu_op = FM_PATCH_OP_SET;
  patch.anchor.width_emu = -1;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  // A to marker needs a two-cell anchor, and a two-cell anchor takes no size.
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.to_row_op = FM_PATCH_OP_SET;
  patch.anchor.to_row = 4;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.kind_op = FM_PATCH_OP_SET;
  patch.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  patch.anchor.width_emu_op = FM_PATCH_OP_SET;
  patch.anchor.width_emu = 100000;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  // Only the description can be cleared.
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.name_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.anchor.from_row_op = FM_PATCH_OP_CLEAR;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.hidden_op = FM_PATCH_OP_SET;
  patch.hidden = 2;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), kInvalidArgument);
  EXPECT_EQ(DrawingBytes(wb.handle), before);
  EXPECT_EQ(ReadObject(wb.handle, id).anchor.from_row, 0U);
}

TEST(FormulonCApiDrawing, UpdateSetsDescriptionAndHiddenOnPictureAndChart) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);  // object 2, the chart
  const std::uint32_t picture = InsertAt(wb.handle, 1, "Logo");
  for (const std::uint32_t id : {picture, 2U}) {
    fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
    patch.description_op = FM_PATCH_OP_SET;
    patch.description = "alt text";
    patch.hidden_op = FM_PATCH_OP_SET;
    patch.hidden = 1;
    ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
    fm_drawing_object_info_t obj = ReadObject(wb.handle, id);
    EXPECT_STREQ(obj.description, "alt text");
    EXPECT_EQ(obj.hidden, 1);

    patch = FM_STRUCT_INIT(fm_drawing_patch_t);
    patch.description_op = FM_PATCH_OP_CLEAR;
    patch.hidden_op = FM_PATCH_OP_SET;
    patch.hidden = 0;
    ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, id, &patch), 0) << fm_get_last_error_message();
    obj = ReadObject(wb.handle, id);
    EXPECT_STREQ(obj.description, "");
    EXPECT_EQ(obj.hidden, 0);
  }
}

TEST(FormulonCApiDrawing, AddPictureHiddenIsReadBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(16, 8);
  fm_drawing_placement_spec_t opts = Defaults();
  opts.hidden = FM_TRISTATE_TRUE;
  uint32_t id = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &id), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadObject(wb.handle, id).hidden, 1);
}

TEST(FormulonCApiDrawing, AddPictureDefaultKindIsOneCellAndTwoCellTakesItsBoxFromTo) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const Bytes png = Png(16, 8);
  fm_drawing_placement_spec_t opts = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  opts.anchor.from_row = 2;
  opts.anchor.from_col = 1;
  uint32_t one = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &one), 0) << fm_get_last_error_message();
  EXPECT_EQ(ReadObject(wb.handle, one).anchor.kind, FM_ANCHOR_KIND_ONE_CELL);

  opts.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  opts.anchor.to_row = 6;
  opts.anchor.to_col = 5;
  uint32_t two = 0;
  ASSERT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &two), 0) << fm_get_last_error_message();
  const fm_drawing_object_info_t obj = ReadObject(wb.handle, two);
  EXPECT_EQ(obj.anchor.kind, FM_ANCHOR_KIND_TWO_CELL);
  EXPECT_EQ(obj.anchor.to_row, 6U);
  EXPECT_EQ(obj.anchor.to_col, 5U);

  // A to marker before the from marker is refused.
  opts.anchor.to_row = 1;
  uint32_t refused = 0;
  EXPECT_EQ(fm_sheet_add_picture(wb.handle, 0, png.data(), png.size(), &opts, &refused), kInvalidArgument);
}

TEST(FormulonCApiDrawing, ZOrderMovesPictureInListOrder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);  // object 2, the chart
  const std::uint32_t first = InsertAt(wb.handle, 1, "A");
  const std::uint32_t second = InsertAt(wb.handle, 2, "B");
  ASSERT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, first, second}));

  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, second, 0), 0) << fm_get_last_error_message();
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{second, 2U, first}));
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, second, 2), 0);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, first, second}));

  const Bytes before = DrawingBytes(wb.handle);
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, second, 2), 0);
  EXPECT_EQ(DrawingBytes(wb.handle), before);

  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, second, 3), FM_STATUS_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, 99, 0), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 9, second, 0), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, first, second}));

  // The chart reorders like a picture.
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, 2, 2), 0) << fm_get_last_error_message();
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{first, second, 2U}));
}

namespace {

fm_status_t Rename(fm_workbook_t* wb, size_t sheet, std::uint32_t id, const char* name) {
  fm_drawing_patch_t patch = FM_STRUCT_INIT(fm_drawing_patch_t);
  patch.name_op = FM_PATCH_OP_SET;
  patch.name = name;
  return fm_sheet_update_drawing(wb, sheet, id, &patch);
}

}  // namespace

TEST(FormulonCApiDrawing, UpdateRenamesPictureAndChartAndSurvivesReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);  // object 2, the chart
  const std::uint32_t picture = InsertAt(wb.handle, 1, "Logo");

  ASSERT_EQ(Rename(wb.handle, 0, picture, "Logo \xE3\x83\xAD\xE3\x82\xB4 \xF0\x9F\x98\x80"), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(Rename(wb.handle, 0, 2, "Sales chart"), 0) << fm_get_last_error_message();
  // A repeated name is allowed.
  ASSERT_EQ(Rename(wb.handle, 0, picture, "Sales chart"), 0) << fm_get_last_error_message();

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(reloaded.handle, 0, 0, &obj), 0);
  EXPECT_STREQ(obj.name, "Sales chart");
  ASSERT_EQ(fm_sheet_get_drawing_object_at(reloaded.handle, 0, 1, &obj), 0);
  EXPECT_STREQ(obj.name, "Sales chart");

  ASSERT_EQ(Rename(wb.handle, 0, picture, "\xE3\x83\xAD\xE3\x82\xB4\xF0\x9F\x98\x80"), 0);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 1, &obj), 0);
  EXPECT_STREQ(obj.name, "\xE3\x83\xAD\xE3\x82\xB4\xF0\x9F\x98\x80");
}

TEST(FormulonCApiDrawing, UpdateNameRejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t id = InsertAt(wb.handle, 1, "Logo");
  const Bytes before = DrawingBytes(wb.handle);

  EXPECT_EQ(Rename(wb.handle, 0, id, nullptr), kInvalidArgument);
  EXPECT_EQ(Rename(nullptr, 0, id, "x"), kBindingNullPointer);
  EXPECT_EQ(Rename(wb.handle, 0, id, ""), kInvalidArgument);
  EXPECT_EQ(Rename(wb.handle, 0, id, "\xC3"), kBindingUtf8EncodingError);
  EXPECT_EQ(Rename(wb.handle, 0, 99, "x"), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(Rename(wb.handle, 9, id, "x"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(Rename(wb.handle, 0, 0xFFFFFFFFU, "x"), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(DrawingBytes(wb.handle), before);
}

TEST(FormulonCApiDrawing, SnapshotRestoreReturnsDeletedPictureToItsPlace) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);
  const std::uint32_t first = InsertAt(wb.handle, 1, "A");
  const std::uint32_t second = InsertAt(wb.handle, 2, "B");

  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, first, &snap.data, &snap.len), 0) << fm_get_last_error_message();
  ASSERT_NE(snap.data, nullptr);
  ASSERT_GT(snap.len, 0U);
  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, first), 0);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, second}));

  uint32_t restored = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(restored, first);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, first, second}));
  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, first, &present, &got, &got_len, &info), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), Png(16, 8));
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 1, &obj), 0);
  EXPECT_EQ(obj.anchor.from_row, 1U);
  EXPECT_STREQ(obj.name, "A");
}

TEST(FormulonCApiDrawing, SnapshotRestoreUndoesMoveAndReorder) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t first = InsertAt(wb.handle, 1, "A");
  const std::uint32_t second = InsertAt(wb.handle, 4, "B");
  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, first, &snap.data, &snap.len), 0);

  fm_drawing_patch_t move = FM_STRUCT_INIT(fm_drawing_patch_t);
  move.anchor.from_row_op = FM_PATCH_OP_SET;
  move.anchor.from_row = 9;
  ASSERT_EQ(fm_sheet_update_drawing(wb.handle, 0, first, &move), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, first, 1), 0);
  ASSERT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{second, first}));

  uint32_t restored = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, &restored), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(restored, first);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{first, second}));
  fm_drawing_object_info_t obj = FM_STRUCT_INIT(fm_drawing_object_info_t);
  ASSERT_EQ(fm_sheet_get_drawing_object_at(wb.handle, 0, 0, &obj), 0);
  EXPECT_EQ(obj.anchor.from_row, 1U);
}

TEST(FormulonCApiDrawing, RestoreWithNewIdAddsACopyOnTop) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::uint32_t first = InsertAt(wb.handle, 1, "A");
  const std::uint32_t second = InsertAt(wb.handle, 2, "B");
  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, first, &snap.data, &snap.len), 0);
  uint32_t copy = 0;
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, FM_DRAWING_RESTORE_NEW_ID, &copy), 0)
      << fm_get_last_error_message();
  EXPECT_NE(copy, first);
  EXPECT_NE(copy, second);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{first, second, copy}));
  const uint8_t* got = nullptr;
  size_t got_len = 0;
  fm_image_info_t info = FM_STRUCT_INIT(fm_image_info_t);
  int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_picture(wb.handle, 0, copy, &present, &got, &got_len, &info), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Bytes(got, got + got_len), Png(16, 8));

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  EXPECT_EQ(ObjectIds(reloaded.handle), (std::vector<std::uint32_t>{first, second, copy}));
}

TEST(FormulonCApiDrawing, SnapshotAndRestoreRejectBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddChartDrawing(wb.handle);
  const std::uint32_t id = InsertAt(wb.handle, 1, "A");
  BufferGuard snap;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, id, &snap.data, &snap.len), 0);

  // A failed snapshot leaves the outputs zeroed.
  BufferGuard none;
  none.len = 7;
  EXPECT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, 99, &none.data, &none.len), FM_STATUS_NOT_FOUND);
  EXPECT_EQ(none.data, nullptr);
  EXPECT_EQ(none.len, 0U);
  EXPECT_EQ(fm_sheet_snapshot_drawing(wb.handle, 4, id, &none.data, &none.len), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, id, nullptr, &none.len), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, id, &none.data, nullptr), kBindingNullPointer);

  uint32_t out = 0;
  const Bytes junk = {1, 2, 3, 4};
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 0, junk.data(), junk.size(), 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len / 2, 0, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 2, &out), kInvalidArgument);
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 9, snap.data, snap.len, 0, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 0, nullptr, 0, 0, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_restore_drawing(wb.handle, 0, snap.data, snap.len, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, id}));

  // A chart captures and restores like a picture.
  BufferGuard chart;
  ASSERT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, 2, &chart.data, &chart.len), 0) << fm_get_last_error_message();
  ASSERT_GT(chart.len, 0U);
  ASSERT_EQ(fm_sheet_remove_drawing(wb.handle, 0, 2), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_restore_drawing(wb.handle, 0, chart.data, chart.len, 0, &out), 0) << fm_get_last_error_message();
  EXPECT_EQ(out, 2U);
  EXPECT_EQ(ObjectIds(wb.handle), (std::vector<std::uint32_t>{2U, id}));
  EXPECT_NE(formulon::io::ooxml::find_passthrough_part(wb.handle->workbook(), "xl/charts/chart1.xml"), nullptr);
}

std::string TwoCellAnchor(const std::string& content) {
  return "<xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row>"
         "<xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</"
         "xdr:row>"
         "<xdr:rowOff>0</xdr:rowOff></xdr:to>" +
         content + "<xdr:clientData/></xdr:twoCellAnchor>";
}

TEST(FormulonCApiDrawing, ShapeWithoutPresetGeometryReadsNull) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string custom =
      "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Freeform\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
      "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm><a:custGeom><a:avLst/>"
      "<a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l=\"0\" t=\"0\" r=\"r\" "
      "b=\"b\"/><a:pathLst/></a:custGeom></xdr:spPr></xdr:sp>";
  const std::string preset =
      "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Rect\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
      "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/></a:xfrm>"
      "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp>";
  const std::string group =
      "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"4\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr>"
      "<a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"100\" cy=\"100\"/><a:chOff x=\"0\" y=\"0\"/>"
      "<a:chExt cx=\"100\" cy=\"100\"/></a:xfrm></xdr:grpSpPr></xdr:grpSp>";
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
      TwoCellAnchor(custom) + TwoCellAnchor(preset) + TwoCellAnchor(group) + "</xdr:wsDr>";
  formulon::Workbook& model = wb.handle->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  model.sheet(0).set_drawing_rel_target(kDrawingPath);

  const auto read = [&](std::uint32_t id) {
    fm_shape_info_t info = FM_STRUCT_INIT(fm_shape_info_t);
    std::int32_t present = 0;
    EXPECT_EQ(fm_sheet_get_shape(wb.handle, 0, id, &present, &info), 0) << fm_get_last_error_message();
    EXPECT_EQ(present, 1) << "object " << id;
    return info;
  };
  const fm_shape_info_t freeform = read(2U);
  EXPECT_EQ(freeform.kind, FM_DRAWING_OBJECT_KIND_SHAPE);
  EXPECT_EQ(freeform.preset_geometry, nullptr);
  const fm_shape_info_t rect = read(3U);
  ASSERT_NE(rect.preset_geometry, nullptr);
  EXPECT_STREQ(rect.preset_geometry, "rect");
  const fm_shape_info_t grouped = read(4U);
  EXPECT_EQ(grouped.kind, FM_DRAWING_OBJECT_KIND_GROUP);
  EXPECT_EQ(grouped.preset_geometry, nullptr);
}

TEST(FormulonCApiDrawing, XlsbRetainedDrawingRefusesImageEdits) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\"></xdr:wsDr>";
  formulon::Workbook& model = wb.handle->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  model.sheet(0).set_unknown_relationships({formulon::UnknownRelationship{"rId1", kRelDrawing, kDrawingPath, false}});
  fm_drawing_patch_t move = FM_STRUCT_INIT(fm_drawing_patch_t);
  move.anchor.from_row_op = FM_PATCH_OP_SET;
  move.anchor.from_row = 1;
  EXPECT_EQ(fm_sheet_update_drawing(wb.handle, 0, 2, &move), kIoDrawingUnparseable);
  EXPECT_EQ(fm_sheet_set_drawing_z_order(wb.handle, 0, 2, 0), kIoDrawingUnparseable);
  BufferGuard snap;
  EXPECT_EQ(fm_sheet_snapshot_drawing(wb.handle, 0, 2, &snap.data, &snap.len), kIoDrawingUnparseable);
  EXPECT_EQ(model.passthrough_parts().size(), 1U);
}
