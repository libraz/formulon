//
// C ABI tests for the geometry fields of a shape: connector glue, preset
// adjust values and the text direction.

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
#include "workbook.h"

namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr const char* kDrawingPath = "xl/drawings/drawing1.xml";

std::string Marker(const char* tag, std::uint32_t row) {
  return std::string("<xdr:") + tag + "><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + std::to_string(row) +
         "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:" + tag + ">";
}

std::string Anchor(std::uint32_t row, const std::string& content) {
  return "<xdr:twoCellAnchor>" + Marker("from", row) + Marker("to", row + 3) + content +
         "<xdr:clientData/></xdr:twoCellAnchor>";
}

std::string Rect(std::uint32_t id) {
  return "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"" + std::to_string(id) + "\" name=\"Rectangle " +
         std::to_string(id) +
         "\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/>"
         "<a:ext cx=\"914400\" cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr>"
         "</xdr:sp>";
}

std::string Connector(std::uint32_t id, const std::string& glue) {
  return "<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"" + std::to_string(id) + "\" name=\"Connector " +
         std::to_string(id) + "\"/><xdr:cNvCxnSpPr>" + glue +
         "</xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"914400\" "
         "cy=\"457200\"/></a:xfrm><a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom></xdr:spPr>"
         "</xdr:cxnSp>";
}

// Rectangles 2 and 3; connector 4 glued at both ends, connector 5 at its start only.
void AddConnectorDrawing(fm_workbook_t* wb) {
  const std::string drawing =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"
      "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
      "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
      Anchor(0, Rect(2)) + Anchor(5, Rect(3)) +
      Anchor(10, Connector(4, "<a:stCxn id=\"2\" idx=\"3\"/><a:endCxn id=\"3\" idx=\"1\"/>")) +
      Anchor(15, Connector(5, "<a:stCxn id=\"2\" idx=\"0\"/>")) + "</xdr:wsDr>";
  formulon::Workbook& model = wb->workbook();
  ASSERT_TRUE(static_cast<bool>(model.add_passthrough_part(
      formulon::PassthroughPart(kDrawingPath, "application/vnd.openxmlformats-officedocument.drawing+xml",
                                Bytes(drawing.begin(), drawing.end())))));
  model.sheet(0).set_drawing_rel_target(kDrawingPath);
}

fm_shape_spec_t ShapeSpec(const char* prst) {
  fm_shape_spec_t s = FM_STRUCT_INIT(fm_shape_spec_t);
  s.placement = FM_STRUCT_INIT(fm_drawing_placement_spec_t);
  s.placement.anchor.kind = FM_ANCHOR_KIND_TWO_CELL;
  s.placement.anchor.from_row = 1;
  s.placement.anchor.from_col = 1;
  s.placement.anchor.to_row = 5;
  s.placement.anchor.to_col = 4;
  s.preset_geometry = prst;
  return s;
}

uint32_t AddShape(fm_workbook_t* wb, const fm_shape_spec_t& spec) {
  uint32_t id = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb, 0, &spec, &id), FM_STATUS_OK) << fm_get_last_error_message();
  return id;
}

fm_shape_info_t ReadShape(const fm_workbook_t* wb, uint32_t id) {
  fm_shape_info_t out = FM_STRUCT_INIT(fm_shape_info_t);
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_shape(wb, 0, id, &present, &out), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(present, 1);
  return out;
}

fm_status_t Patch(fm_workbook_t* wb, uint32_t id, const fm_shape_patch_t& patch) {
  return fm_sheet_update_shape(wb, 0, id, nullptr, &patch);
}

void SaveAndReload(const WorkbookGuard& source, WorkbookGuard* out) {
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(source.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), FM_STATUS_OK)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, nullptr, &out->handle), FM_STATUS_OK)
      << fm_get_last_error_message();
}

// The adjust value named `name`, or -1 when the shape lists none by that name.
int64_t AdjustValue(const fm_shape_info_t& info, const char* name) {
  for (uint32_t i = 0; i < info.adjust_values_count; ++i) {
    if (std::string(info.adjust_values[i].name) == name) {
      return info.adjust_values[i].value;
    }
  }
  return -1;
}

TEST(FormulonCApiShapeGeometry, EnumOrdinals) {
  EXPECT_EQ(FM_SHAPE_TEXT_VERT_ABSENT, -1);
  EXPECT_EQ(FM_SHAPE_TEXT_VERT_HORZ, 0);
  EXPECT_EQ(FM_SHAPE_TEXT_VERT_VERT270, 2);
  EXPECT_EQ(FM_SHAPE_TEXT_VERT_EA_VERT, 3);
  EXPECT_EQ(FM_SHAPE_TEXT_VERT_MONGOLIAN_VERT, 6);
}

// A glued connector reports both ends; a free end and a plain shape report none.
TEST(FormulonCApiShapeGeometry, ConnectorGlueReads) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  AddConnectorDrawing(wb.handle);

  WorkbookGuard again;
  SaveAndReload(wb, &again);
  for (const fm_workbook_t* book :
       {static_cast<const fm_workbook_t*>(wb.handle), static_cast<const fm_workbook_t*>(again.handle)}) {
    const fm_shape_info_t both = ReadShape(book, 4);
    ASSERT_EQ(both.has_start_connection, 1);
    EXPECT_EQ(both.start_connection.object_id, 2U);
    EXPECT_EQ(both.start_connection.site_index, 3U);
    ASSERT_EQ(both.has_end_connection, 1);
    EXPECT_EQ(both.end_connection.object_id, 3U);
    EXPECT_EQ(both.end_connection.site_index, 1U);

    const fm_shape_info_t start_only = ReadShape(book, 5);
    EXPECT_EQ(start_only.has_start_connection, 1);
    EXPECT_EQ(start_only.start_connection.site_index, 0U);
    EXPECT_EQ(start_only.has_end_connection, 0);
    EXPECT_EQ(start_only.end_connection.object_id, 0U);

    const fm_shape_info_t rect = ReadShape(book, 2);
    EXPECT_EQ(rect.has_start_connection, 0);
    EXPECT_EQ(rect.has_end_connection, 0);
  }
}

// Flipping a glued connector unglues both of its ends.
TEST(FormulonCApiShapeGeometry, FlipUngluesAConnector) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  AddConnectorDrawing(wb.handle);
  fm_shape_patch_t flip = FM_STRUCT_INIT(fm_shape_patch_t);
  flip.flip_h_op = FM_PATCH_OP_SET;
  flip.flip_h = 1;
  ASSERT_EQ(Patch(wb.handle, 4, flip), FM_STATUS_OK) << fm_get_last_error_message();
  const fm_shape_info_t out = ReadShape(wb.handle, 4);
  EXPECT_EQ(out.flip_h, 1);
  EXPECT_EQ(out.has_start_connection, 0);
  EXPECT_EQ(out.has_end_connection, 0);
}

// Adjust values are written on add, merged by name on update, and removed by CLEAR.
TEST(FormulonCApiShapeGeometry, AdjustValuesSetMergeAndClear) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const uint32_t plain = AddShape(wb.handle, ShapeSpec("wedgeRectCallout"));
  const fm_shape_info_t none = ReadShape(wb.handle, plain);
  EXPECT_EQ(none.adjust_values_count, 0U);
  EXPECT_EQ(none.adjust_values, nullptr);

  const fm_shape_adjust_t values[] = {{"adj1", -20000}, {"adj2", 60000}};
  fm_shape_spec_t spec = ShapeSpec("wedgeRectCallout");
  spec.adjust_values = values;
  spec.adjust_values_count = 2;
  const uint32_t id = AddShape(wb.handle, spec);

  WorkbookGuard again;
  SaveAndReload(wb, &again);
  fm_shape_info_t out = ReadShape(again.handle, id);
  ASSERT_EQ(out.adjust_values_count, 2U);
  EXPECT_STREQ(out.adjust_values[0].name, "adj1");
  EXPECT_EQ(out.adjust_values[0].value, -20000);
  EXPECT_STREQ(out.adjust_values[1].name, "adj2");
  EXPECT_EQ(out.adjust_values[1].value, 60000);

  const fm_shape_adjust_patch_t edits[] = {{"adj1", FM_PATCH_OP_CLEAR, 0}, {"adj2", FM_PATCH_OP_SET, 10000}};
  fm_shape_patch_t patch = FM_STRUCT_INIT(fm_shape_patch_t);
  patch.adjust_values = edits;
  patch.adjust_values_count = 2;
  ASSERT_EQ(Patch(again.handle, id, patch), FM_STATUS_OK) << fm_get_last_error_message();
  out = ReadShape(again.handle, id);
  EXPECT_EQ(out.adjust_values_count, 1U);
  EXPECT_EQ(AdjustValue(out, "adj1"), -1);
  EXPECT_EQ(AdjustValue(out, "adj2"), 10000);

  // A patch without entries keeps every guide.
  const fm_shape_patch_t keep = FM_STRUCT_INIT(fm_shape_patch_t);
  ASSERT_EQ(Patch(again.handle, id, keep), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(AdjustValue(ReadShape(again.handle, id), "adj2"), 10000);
}

// Malformed adjust-value edits are refused and change nothing.
TEST(FormulonCApiShapeGeometry, AdjustValueRefusals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const fm_shape_adjust_t start[] = {{"adj1", 5000}};
  fm_shape_spec_t spec = ShapeSpec("wedgeRectCallout");
  spec.adjust_values = start;
  spec.adjust_values_count = 1;
  const uint32_t id = AddShape(wb.handle, spec);

  const auto patch_with = [&](const fm_shape_adjust_patch_t* edits, uint32_t count) {
    fm_shape_patch_t p = FM_STRUCT_INIT(fm_shape_patch_t);
    p.adjust_values = edits;
    p.adjust_values_count = count;
    return Patch(wb.handle, id, p);
  };
  const fm_shape_adjust_patch_t keep[] = {{"adj1", FM_PATCH_OP_KEEP, 1}};
  EXPECT_EQ(patch_with(keep, 1), FM_STATUS_INVALID_ARGUMENT);
  const fm_shape_adjust_patch_t bad_op[] = {{"adj1", static_cast<fm_patch_op_t>(3), 1}};
  EXPECT_EQ(patch_with(bad_op, 1), FM_STATUS_INVALID_ARGUMENT);
  const fm_shape_adjust_patch_t unnamed[] = {{nullptr, FM_PATCH_OP_SET, 1}};
  EXPECT_EQ(patch_with(unnamed, 1), FM_STATUS_INVALID_ARGUMENT);
  const fm_shape_adjust_patch_t empty[] = {{"", FM_PATCH_OP_SET, 1}};
  EXPECT_EQ(patch_with(empty, 1), FM_STATUS_INVALID_ARGUMENT);
  const fm_shape_adjust_patch_t repeated[] = {{"adj1", FM_PATCH_OP_SET, 1}, {"adj1", FM_PATCH_OP_CLEAR, 0}};
  const fm_shape_adjust_patch_t undefined[] = {{"adj3", FM_PATCH_OP_SET, 1}};
  EXPECT_EQ(patch_with(repeated, 2), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(patch_with(undefined, 1), FM_STATUS_INVALID_ARGUMENT);
  const fm_shape_adjust_patch_t bad_utf8[] = {{"adj\xC3", FM_PATCH_OP_SET, 1}};
  EXPECT_EQ(patch_with(bad_utf8, 1), FM_STATUS_BINDING_UTF8ENCODING_ERROR);
  EXPECT_EQ(patch_with(nullptr, 1), FM_STATUS_INVALID_ARGUMENT);
  const std::vector<fm_shape_adjust_patch_t> many(65, fm_shape_adjust_patch_t{"adj1", FM_PATCH_OP_SET, 1});
  EXPECT_EQ(patch_with(many.data(), 65), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  EXPECT_EQ(AdjustValue(ReadShape(wb.handle, id), "adj1"), 5000);

  uint32_t ignored = 0;
  fm_shape_spec_t null_array = ShapeSpec("wedgeRectCallout");
  null_array.adjust_values_count = 1;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &null_array, &ignored), FM_STATUS_INVALID_ARGUMENT);
  const std::vector<fm_shape_adjust_t> too_many(65, fm_shape_adjust_t{"adj1", 1});
  fm_shape_spec_t over = ShapeSpec("wedgeRectCallout");
  over.adjust_values = too_many.data();
  over.adjust_values_count = 65;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &over, &ignored), FM_STATUS_SECURITY_RESOURCE_LIMIT);
  const fm_shape_adjust_t twice[] = {{"adj1", 1}, {"adj1", 2}};
  fm_shape_spec_t dup = ShapeSpec("wedgeRectCallout");
  dup.adjust_values = twice;
  dup.adjust_values_count = 2;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &dup, &ignored), FM_STATUS_INVALID_ARGUMENT);
}

// The text direction is unstated by default, set by add and update, and removed by CLEAR.
TEST(FormulonCApiShapeGeometry, TextDirectionSetAndClear) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  fm_shape_spec_t plain = ShapeSpec("rect");
  plain.text = "Text";
  const uint32_t unstated = AddShape(wb.handle, plain);
  EXPECT_EQ(ReadShape(wb.handle, unstated).text_vert, FM_SHAPE_TEXT_VERT_ABSENT);

  fm_shape_spec_t spec = ShapeSpec("rect");
  spec.text = "Text";
  spec.has_text_vert = 1;
  spec.text_vert = FM_SHAPE_TEXT_VERT_VERT270;
  const uint32_t id = AddShape(wb.handle, spec);
  WorkbookGuard again;
  SaveAndReload(wb, &again);
  EXPECT_EQ(ReadShape(again.handle, id).text_vert, FM_SHAPE_TEXT_VERT_VERT270);
  // Horizontal is a stated value, distinct from none.
  fm_shape_patch_t horz = FM_STRUCT_INIT(fm_shape_patch_t);
  horz.text_vert_op = FM_PATCH_OP_SET;
  horz.text_vert = FM_SHAPE_TEXT_VERT_HORZ;
  ASSERT_EQ(Patch(again.handle, id, horz), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(ReadShape(again.handle, id).text_vert, FM_SHAPE_TEXT_VERT_HORZ);

  fm_shape_patch_t clear = FM_STRUCT_INIT(fm_shape_patch_t);
  clear.text_vert_op = FM_PATCH_OP_CLEAR;
  ASSERT_EQ(Patch(again.handle, id, clear), FM_STATUS_OK) << fm_get_last_error_message();
  EXPECT_EQ(ReadShape(again.handle, id).text_vert, FM_SHAPE_TEXT_VERT_ABSENT);
}

// A text direction outside the enumeration, or on a connector, is refused.
TEST(FormulonCApiShapeGeometry, TextDirectionRefusals) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), FM_STATUS_OK);
  const uint32_t id = AddShape(wb.handle, ShapeSpec("rect"));
  for (const int32_t value : {-1, 7}) {
    fm_shape_patch_t p = FM_STRUCT_INIT(fm_shape_patch_t);
    p.text_vert_op = FM_PATCH_OP_SET;
    p.text_vert = value;
    EXPECT_EQ(Patch(wb.handle, id, p), FM_STATUS_INVALID_ARGUMENT) << value;
    fm_shape_spec_t s = ShapeSpec("rect");
    s.has_text_vert = 1;
    s.text_vert = value;
    uint32_t ignored = 0;
    EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &s, &ignored), FM_STATUS_INVALID_ARGUMENT) << value;
  }
  fm_shape_patch_t bad_op = FM_STRUCT_INIT(fm_shape_patch_t);
  bad_op.text_vert_op = static_cast<fm_patch_op_t>(3);
  EXPECT_EQ(Patch(wb.handle, id, bad_op), FM_STATUS_INVALID_ARGUMENT);
  fm_shape_spec_t flag = ShapeSpec("rect");
  flag.has_text_vert = 2;
  uint32_t ignored = 0;
  EXPECT_EQ(fm_sheet_add_shape(wb.handle, 0, &flag, &ignored), FM_STATUS_INVALID_ARGUMENT);

  const uint32_t connector = AddShape(wb.handle, ShapeSpec("straightConnector1"));
  fm_shape_patch_t vert = FM_STRUCT_INIT(fm_shape_patch_t);
  vert.text_vert_op = FM_PATCH_OP_SET;
  vert.text_vert = FM_SHAPE_TEXT_VERT_VERT;
  EXPECT_EQ(Patch(wb.handle, connector, vert), FM_STATUS_INVALID_ARGUMENT);
  EXPECT_EQ(ReadShape(wb.handle, connector).text_vert, FM_SHAPE_TEXT_VERT_ABSENT);
  EXPECT_EQ(ReadShape(wb.handle, id).text_vert, FM_SHAPE_TEXT_VERT_ABSENT);
}

}  // namespace
