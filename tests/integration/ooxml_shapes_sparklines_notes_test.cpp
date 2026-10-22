//
// Integration test: drawing shapes, sparkline groups and legacy note shapes
// that a load -> save -> load cycle must keep, through structural edits and
// through the typed edit APIs.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "drawing/shape_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/sparkline_groups.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  std::string path;
  std::string body;
  /// Empty: covered by the `xml` / `rels` / `vml` defaults.
  std::string content_type;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path.c_str(), p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// One-sheet package; `sheet_rels` are the `<Relationship>` elements of the sheet's rels part.
std::vector<std::uint8_t> BuildPackage(std::string_view sheet_body, std::string_view sheet_rels = {},
                                       const std::vector<PartFile>& extra_parts = {}) {
  std::string content_types =
      std::string(kXmlDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
  for (const PartFile& part : extra_parts) {
    if (!part.content_type.empty()) {
      content_types += "<Override PartName=\"/" + part.path + "\" ContentType=\"" + part.content_type + "\"/>";
    }
  }
  content_types += "</Types>";
  const std::string package_rels =
      std::string(kXmlDecl) +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/></Relationships>";
  const std::string workbook_xml = std::string(kXmlDecl) +
                                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                                   "xmlns:r=\"" +
                                   std::string(kRelNs) + "\">" +
                                   "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>";
  const std::string workbook_rels =
      std::string(kXmlDecl) +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/></Relationships>";
  const std::string worksheet = std::string(kXmlDecl) +
                                "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                                "xmlns:r=\"" +
                                std::string(kRelNs) + "\">" + std::string(sheet_body) + "</worksheet>";
  std::vector<PartFile> parts = {
      {"[Content_Types].xml", content_types, {}},  {"_rels/.rels", package_rels, {}},
      {"xl/workbook.xml", workbook_xml, {}},       {"xl/_rels/workbook.xml.rels", workbook_rels, {}},
      {"xl/worksheets/sheet1.xml", worksheet, {}},
  };
  if (!sheet_rels.empty()) {
    parts.push_back({"xl/worksheets/_rels/sheet1.xml.rels",
                     std::string(kXmlDecl) +
                         "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                         std::string(sheet_rels) + "</Relationships>",
                     {}});
  }
  for (const PartFile& part : extra_parts) {
    parts.push_back(part);
  }
  return BuildZip(parts);
}

Workbook LoadOrDie(const std::vector<std::uint8_t>& bytes) {
  auto load_or = io::read_ooxml(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  return std::move(load_or.value().workbook);
}

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << (save_or ? "" : save_or.error().message);
  return save_or.value();
}

std::string ReadPart(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  auto entry_or = zip.read_entry(name);
  EXPECT_TRUE(static_cast<bool>(entry_or)) << "missing part: " << name;
  if (!entry_or) {
    return {};
  }
  return std::string(reinterpret_cast<const char*>(entry_or.value().data()), entry_or.value().size());
}

bool Has(const std::string& text, std::string_view needle) {
  return text.find(needle) != std::string::npos;
}

// ---------------------------------------------------------------------------
// Drawing shapes
// ---------------------------------------------------------------------------

constexpr std::string_view kDrawingType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";
constexpr std::string_view kDrawingCt = "application/vnd.openxmlformats-officedocument.drawing+xml";

std::string Marker(const char* tag, int col, int row) {
  return std::string("<xdr:") + tag + "><xdr:col>" + std::to_string(col) +
         "</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + std::to_string(row) +
         "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:" + tag + ">";
}

constexpr std::string_view kRect =
    "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"2\" name=\"Rect\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
    "<xdr:spPr><a:xfrm rot=\"900000\"><a:off x=\"100\" y=\"200\"/><a:ext cx=\"300\" cy=\"400\"/></a:xfrm>"
    "<a:prstGeom prst=\"roundRect\"><a:avLst><a:gd name=\"adj\" fmla=\"val 25000\"/></a:avLst></a:prstGeom>"
    "<a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>"
    "<a:ln w=\"19050\"><a:solidFill><a:srgbClr val=\"0000FF\"/></a:solidFill><a:prstDash val=\"dash\"/></a:ln>"
    "<a:effectLst><a:outerShdw blurRad=\"50800\"/></a:effectLst></xdr:spPr>"
    "<xdr:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"/></a:lnRef></xdr:style>"
    "<xdr:txBody><a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/>"
    "<a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-US\" sz=\"1100\" b=\"1\"/><a:t>Hello</a:t></a:r></a:p>"
    "<a:p><a:r><a:rPr lang=\"en-US\"/><a:t>World &amp; more</a:t></a:r></a:p></xdr:txBody></xdr:sp>";

// A shape the API cannot represent: custom geometry, a gradient fill, a theme-coloured outline.
constexpr std::string_view kCustom =
    "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"3\" name=\"Custom\" hidden=\"1\"/><xdr:cNvSpPr "
    "txBox=\"1\"/>"
    "</xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"10\" cy=\"10\"/></a:xfrm>"
    "<a:custGeom><a:avLst/><a:pathLst><a:path w=\"10\" h=\"10\"/></a:pathLst></a:custGeom>"
    "<a:gradFill><a:gsLst><a:gs pos=\"0\"><a:srgbClr val=\"FFFFFF\"/></a:gs></a:gsLst></a:gradFill>"
    "<a:ln><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>";

constexpr std::string_view kConnector =
    "<xdr:cxnSp macro=\"\"><xdr:nvCxnSpPr><xdr:cNvPr id=\"4\" name=\"Line\"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr>"
    "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"10\" cy=\"10\"/></a:xfrm>"
    "<a:prstGeom prst=\"straightConnector1\"><a:avLst/></a:prstGeom>"
    "<a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"00FF00\"/></a:solidFill></a:ln></xdr:spPr></xdr:cxnSp>";

constexpr std::string_view kGroup =
    "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"5\" name=\"Group\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
    "<xdr:grpSpPr><a:xfrm rot=\"60000\"><a:off x=\"0\" y=\"0\"/><a:ext cx=\"20\" cy=\"20\"/>"
    "<a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"20\" cy=\"20\"/></a:xfrm></xdr:grpSpPr>"
    "<xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"6\" name=\"Child\"/><xdr:cNvSpPr/></xdr:nvSpPr>"
    "<xdr:spPr><a:prstGeom prst=\"ellipse\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp>";

std::string Anchored(std::string_view content, int col, int row, int to_col, int to_row) {
  return "<xdr:twoCellAnchor>" + Marker("from", col, row) + Marker("to", to_col, to_row) + std::string(content) +
         "<xdr:clientData/></xdr:twoCellAnchor>";
}

std::string DrawingXml() {
  return std::string(kXmlDecl) +
         "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
         Anchored(kRect, 1, 2, 3, 6) + Anchored(kCustom, 5, 4, 6, 5) + Anchored(kConnector, 0, 8, 2, 9) +
         Anchored(kGroup, 7, 10, 9, 12) + "</xdr:wsDr>";
}

std::vector<std::uint8_t> ShapesPackage() {
  const std::string rels =
      "<Relationship Id=\"rId1\" Type=\"" + std::string(kDrawingType) + "\" Target=\"../drawings/drawing1.xml\"/>";
  return BuildPackage("<sheetData/><drawing r:id=\"rId1\"/>", rels,
                      {{"xl/drawings/drawing1.xml", DrawingXml(), std::string(kDrawingCt)}});
}

TEST(ShapeFidelity, ListReportsEveryModelledField) {
  const Workbook wb = LoadOrDie(ShapesPackage());
  auto shapes = list_shapes(wb, 0);
  ASSERT_TRUE(static_cast<bool>(shapes));
  ASSERT_EQ(shapes.value().size(), 4U);
  const ShapeInfo& rect = shapes.value()[0];
  EXPECT_EQ(rect.object.object_id, 2U);
  EXPECT_EQ(rect.object.kind, DrawingObjectKind::kShape);
  EXPECT_EQ(rect.object.name, "Rect");
  EXPECT_EQ(rect.object.from.row, 2U);
  EXPECT_EQ(rect.object.to.col, 3U);
  EXPECT_EQ(rect.prst, "roundRect");
  EXPECT_EQ(rect.text, "Hello\nWorld & more");
  EXPECT_EQ(rect.fill, ShapePaint::kSolid);
  EXPECT_EQ(rect.fill_rgb, 0xFF0000U);
  EXPECT_EQ(rect.line, ShapePaint::kSolid);
  EXPECT_EQ(rect.line_rgb, 0x0000FFU);
  EXPECT_EQ(rect.line_width_emu, 19050);
  EXPECT_EQ(rect.rotation, 900000);
  EXPECT_FALSE(rect.hidden);
  EXPECT_FALSE(rect.text_box);

  const ShapeInfo& custom = shapes.value()[1];
  EXPECT_TRUE(custom.prst.empty());
  EXPECT_TRUE(custom.hidden);
  EXPECT_TRUE(custom.text_box);
  EXPECT_EQ(custom.fill, ShapePaint::kOther);
  EXPECT_EQ(custom.line, ShapePaint::kOther);

  EXPECT_EQ(shapes.value()[2].object.kind, DrawingObjectKind::kConnector);
  EXPECT_EQ(shapes.value()[2].prst, "straightConnector1");
  EXPECT_EQ(shapes.value()[3].object.kind, DrawingObjectKind::kGroup);
  EXPECT_EQ(shapes.value()[3].rotation, 60000);
}

TEST(ShapeFidelity, UntouchedDrawingSurvivesAResaveVerbatim) {
  const auto saved = SaveOrDie(LoadOrDie(ShapesPackage()));
  EXPECT_EQ(ReadPart(saved, "xl/drawings/drawing1.xml"), DrawingXml());
}

TEST(ShapeFidelity, AnchorsShiftWithRowInsertionAndDeletion) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 3)));
  auto shapes = list_shapes(wb, 0);
  ASSERT_TRUE(static_cast<bool>(shapes));
  ASSERT_EQ(shapes.value().size(), 4U);
  EXPECT_EQ(shapes.value()[0].object.from.row, 5U);
  EXPECT_EQ(shapes.value()[0].object.to.row, 9U);
  EXPECT_EQ(shapes.value()[3].object.from.row, 13U);

  Workbook again = LoadOrDie(SaveOrDie(wb));
  ASSERT_EQ(list_shapes(again, 0).value().size(), 4U);
  // The connector spans rows 11..12 after the insert; deleting both removes it.
  ASSERT_TRUE(static_cast<bool>(again.delete_rows(0, 11, 2)));
  shapes = list_shapes(again, 0);
  ASSERT_EQ(shapes.value().size(), 3U);
  EXPECT_EQ(shapes.value()[0].object.name, "Rect");
  EXPECT_EQ(shapes.value()[1].object.name, "Custom");
  EXPECT_EQ(shapes.value()[2].object.name, "Group");
  // The group sat on row 13, below the deleted band.
  EXPECT_EQ(shapes.value()[2].object.from.row, 11U);
}

TEST(ShapeFidelity, AnchorsShiftWithColumnInsertionAndDeletion) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 2)));
  auto shapes = list_shapes(wb, 0);
  ASSERT_EQ(shapes.value().size(), 4U);
  EXPECT_EQ(shapes.value()[0].object.from.col, 3U);
  EXPECT_EQ(shapes.value()[0].object.to.col, 5U);
  EXPECT_EQ(shapes.value()[3].object.from.col, 9U);

  Workbook again = LoadOrDie(SaveOrDie(wb));
  // The connector spans columns 2..4 after the insert; deleting all three removes it.
  ASSERT_TRUE(static_cast<bool>(again.delete_cols(0, 2, 3)));
  shapes = list_shapes(again, 0);
  ASSERT_EQ(shapes.value().size(), 3U);
  EXPECT_EQ(shapes.value()[2].object.name, "Group");
  EXPECT_EQ(shapes.value()[2].object.from.col, 6U);
}

TEST(ShapeFidelity, UpdateChangesOnlyTheNamedFields) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ShapeSpec spec;
  spec.fields = shape_field::kName | shape_field::kFill | shape_field::kText | shape_field::kRotation;
  spec.name = "Renamed";
  spec.fill = ShapePaint::kSolid;
  spec.fill_rgb = 0x00AA11;
  spec.text = "One\nTwo\nThree";
  spec.rotation = 0;
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 2, spec)));

  const std::string xml = ReadPart(SaveOrDie(wb), "xl/drawings/drawing1.xml");
  EXPECT_TRUE(Has(xml, "name=\"Renamed\"")) << xml;
  EXPECT_TRUE(Has(xml, "<a:srgbClr val=\"00AA11\"/>")) << xml;
  EXPECT_FALSE(Has(xml, "rot=\"900000\"")) << xml;
  // Everything the API does not model stays.
  EXPECT_TRUE(Has(xml, "<a:gd name=\"adj\" fmla=\"val 25000\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<a:prstDash val=\"dash\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<a:outerShdw blurRad=\"50800\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<a:lnRef idx=\"2\">")) << xml;
  EXPECT_TRUE(Has(xml, "<a:pPr algn=\"ctr\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "sz=\"1100\" b=\"1\"")) << xml;
  // The other shapes are untouched.
  EXPECT_TRUE(Has(xml, "<a:custGeom>")) << xml;
  EXPECT_TRUE(Has(xml, "<a:gradFill>")) << xml;

  auto shapes = list_shapes(wb, 0);
  EXPECT_EQ(shapes.value()[0].text, "One\nTwo\nThree");
  EXPECT_EQ(shapes.value()[0].object.name, "Renamed");
  EXPECT_EQ(shapes.value()[0].rotation, 0);
  EXPECT_EQ(shapes.value()[0].line_rgb, 0x0000FFU);
  EXPECT_EQ(shapes.value()[0].prst, "roundRect");
}

TEST(ShapeFidelity, UpdateMovesTheAnchorAndTheTransform) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ShapeSpec spec;
  spec.fields =
      shape_field::kAnchor | shape_field::kPrst | shape_field::kLine | shape_field::kLineWidth | shape_field::kHidden;
  spec.anchor_kind = AnchorKind::kTwoCell;
  spec.from = AnchorPoint{1, 1, 0, 0};
  spec.to = AnchorPoint{4, 3, 0, 0};
  spec.prst = "ellipse";
  spec.line = ShapePaint::kNone;
  spec.line_width_emu = 12700;
  spec.hidden = true;
  ASSERT_TRUE(static_cast<bool>(update_shape(wb, 0, 2, spec)));
  const ShapeInfo rect = list_shapes(wb, 0).value()[0];
  EXPECT_EQ(rect.object.from.row, 1U);
  EXPECT_EQ(rect.object.to.row, 4U);
  EXPECT_EQ(rect.object.to.col, 3U);
  EXPECT_EQ(rect.prst, "ellipse");
  EXPECT_EQ(rect.line, ShapePaint::kNone);
  EXPECT_EQ(rect.line_width_emu, 12700);
  EXPECT_TRUE(rect.hidden);
  // A new preset drops guides that belonged to the old one.
  EXPECT_FALSE(Has(ReadPart(SaveOrDie(wb), "xl/drawings/drawing1.xml"), "val 25000"));

  spec.anchor_kind = AnchorKind::kOneCell;
  EXPECT_FALSE(static_cast<bool>(update_shape(wb, 0, 2, spec)));
  EXPECT_FALSE(static_cast<bool>(update_shape(wb, 0, 99, spec)));
}

TEST(ShapeFidelity, InsertAppendsOnTopAndRemoveDeletes) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ShapeSpec spec;
  spec.fields = shape_field::kAnchor | shape_field::kPrst | shape_field::kText | shape_field::kTextBox |
                shape_field::kFill | shape_field::kLine | shape_field::kLineWidth | shape_field::kName;
  spec.anchor_kind = AnchorKind::kOneCell;
  spec.from = AnchorPoint{20, 4, 1000, 2000};
  spec.cx = 914400;
  spec.cy = 457200;
  spec.prst = "wedgeRectCallout";
  spec.text = "a < b";
  spec.text_box = true;
  spec.fill = ShapePaint::kSolid;
  spec.fill_rgb = 0xFFFF00;
  spec.line = ShapePaint::kSolid;
  spec.line_rgb = 0x123456;
  spec.line_width_emu = 25400;
  spec.name = "Callout";
  auto id = insert_shape(wb, 0, spec);
  ASSERT_TRUE(static_cast<bool>(id));
  EXPECT_EQ(id.value(), 7U);

  Workbook again = LoadOrDie(SaveOrDie(wb));
  auto shapes = list_shapes(again, 0);
  ASSERT_EQ(shapes.value().size(), 5U);
  const ShapeInfo& added = shapes.value()[4];
  EXPECT_EQ(added.object.object_id, 7U);
  EXPECT_EQ(added.object.name, "Callout");
  EXPECT_EQ(added.object.anchor_kind, AnchorKind::kOneCell);
  EXPECT_EQ(added.object.from.row, 20U);
  EXPECT_EQ(added.object.from.col_off, 2000);
  EXPECT_EQ(added.object.cx, 914400);
  EXPECT_EQ(added.prst, "wedgeRectCallout");
  EXPECT_EQ(added.text, "a < b");
  EXPECT_TRUE(added.text_box);
  EXPECT_EQ(added.fill_rgb, 0xFFFF00U);
  EXPECT_EQ(added.line_rgb, 0x123456U);
  EXPECT_EQ(added.line_width_emu, 25400);

  ASSERT_TRUE(static_cast<bool>(remove_shape(again, 0, 5)));
  shapes = list_shapes(again, 0);
  ASSERT_EQ(shapes.value().size(), 4U);
  EXPECT_FALSE(Has(ReadPart(SaveOrDie(again), "xl/drawings/drawing1.xml"), "name=\"Child\""));
  EXPECT_FALSE(static_cast<bool>(remove_shape(again, 0, 5)));
}

TEST(ShapeFidelity, InsertIntoASheetWithoutADrawingCreatesOne) {
  Workbook wb = LoadOrDie(BuildPackage("<sheetData/>"));
  EXPECT_TRUE(list_shapes(wb, 0).value().empty());
  ShapeSpec spec;
  spec.fields = shape_field::kAnchor | shape_field::kPrst | shape_field::kLine;
  spec.anchor_kind = AnchorKind::kAbsolute;
  spec.from = AnchorPoint{0, 0, 6000, 5000};
  spec.cx = 100000;
  spec.cy = 0;
  spec.prst = "line";
  spec.line = ShapePaint::kSolid;
  spec.line_rgb = 0xFF00FF;
  auto id = insert_shape(wb, 0, spec);
  ASSERT_TRUE(static_cast<bool>(id));

  Workbook again = LoadOrDie(SaveOrDie(wb));
  auto shapes = list_shapes(again, 0);
  ASSERT_EQ(shapes.value().size(), 1U);
  EXPECT_EQ(shapes.value()[0].object.kind, DrawingObjectKind::kConnector);
  EXPECT_EQ(shapes.value()[0].object.anchor_kind, AnchorKind::kAbsolute);
  EXPECT_EQ(shapes.value()[0].object.from.col_off, 5000);
  EXPECT_EQ(shapes.value()[0].object.from.row_off, 6000);

  spec.fields = shape_field::kText;
  EXPECT_FALSE(static_cast<bool>(update_shape(again, 0, id.value(), spec)));  // connectors hold no text
}

TEST(ShapeFidelity, InsertRejectsBadInput) {
  Workbook wb = LoadOrDie(ShapesPackage());
  ShapeSpec spec;
  EXPECT_FALSE(static_cast<bool>(insert_shape(wb, 0, spec)));  // no anchor
  spec.fields = shape_field::kAnchor | shape_field::kPrst;
  spec.anchor_kind = AnchorKind::kTwoCell;
  spec.from = AnchorPoint{3, 3, 0, 0};
  spec.to = AnchorPoint{2, 5, 0, 0};
  spec.prst = "rect";
  EXPECT_FALSE(static_cast<bool>(insert_shape(wb, 0, spec)));  // to before from
  spec.to = AnchorPoint{4, 5, 0, 0};
  spec.prst = "rect\"/><x";
  EXPECT_FALSE(static_cast<bool>(insert_shape(wb, 0, spec)));  // not a token
  EXPECT_EQ(list_shapes(wb, 0).value().size(), 4U);
}

// ---------------------------------------------------------------------------
// Sparkline groups
// ---------------------------------------------------------------------------

constexpr std::string_view kSparklineExt =
    "<extLst><ext uri=\"{05C60535-1F16-4fd2-B633-F4F36F0B64E0}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\">"
    "<x14:sparklineGroups xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">"
    "<x14:sparklineGroup type=\"column\" displayEmptyCellsAs=\"span\" markers=\"1\" high=\"1\" low=\"1\" first=\"1\" "
    "last=\"1\" negative=\"1\" displayXAxis=\"1\" displayHidden=\"1\" minAxisType=\"custom\" maxAxisType=\"group\" "
    "manualMin=\"-2.5\" manualMax=\"10\" dateAxis=\"1\" lineWeight=\"1.5\" rightToLeft=\"1\" "
    "xmlns:xr2=\"http://schemas.microsoft.com/office/spreadsheetml/2015/revision2\" "
    "xr2:uid=\"{11111111-2222-3333-4444-555555555555}\">"
    "<x14:colorSeries theme=\"4\" tint=\"-0.499984740745262\"/><x14:colorNegative rgb=\"FFD00000\"/>"
    "<x14:colorAxis rgb=\"FF000000\"/><x14:colorMarkers indexed=\"10\"/><x14:colorFirst auto=\"1\"/>"
    "<x14:colorLast rgb=\"FF0000FF\"/><x14:colorHigh rgb=\"FF00FF00\"/><x14:colorLow theme=\"5\"/>"
    "<xm:f>Sheet1!A1:C1</xm:f><x14:sparklines>"
    "<x14:sparkline><xm:f>Sheet1!A1:C1</xm:f><xm:sqref>E1</xm:sqref></x14:sparkline>"
    "<x14:sparkline><xm:f>Sheet1!A2:C2</xm:f><xm:sqref>E2</xm:sqref></x14:sparkline>"
    "</x14:sparklines></x14:sparklineGroup>"
    "<x14:sparklineGroup><x14:colorSeries rgb=\"FF376092\"/><x14:sparklines>"
    "<x14:sparkline><xm:f>Sheet1!A5:C5</xm:f><xm:sqref>F5</xm:sqref></x14:sparkline>"
    "</x14:sparklines></x14:sparklineGroup>"
    "</x14:sparklineGroups></ext></extLst>";

std::vector<std::uint8_t> SparklinePackage() {
  return BuildPackage("<sheetData/>" + std::string(kSparklineExt));
}

void ExpectFirstGroup(const io::SparklineGroup& g) {
  EXPECT_EQ(g.type, io::SparklineType::kColumn);
  EXPECT_EQ(g.empty_cells_as, io::SparklineEmptyCells::kSpan);
  EXPECT_TRUE(g.markers && g.high && g.low && g.first && g.last && g.negative && g.display_x_axis && g.display_hidden);
  EXPECT_EQ(g.min_axis_type, io::SparklineAxisType::kCustom);
  EXPECT_EQ(g.max_axis_type, io::SparklineAxisType::kGroup);
  EXPECT_TRUE(g.has_manual_min);
  EXPECT_DOUBLE_EQ(g.manual_min, -2.5);
  EXPECT_TRUE(g.has_manual_max);
  EXPECT_DOUBLE_EQ(g.manual_max, 10.0);
  EXPECT_TRUE(g.date_axis);
  EXPECT_EQ(g.date_axis_ref, "Sheet1!A1:C1");
  EXPECT_DOUBLE_EQ(g.line_weight, 1.5);
  EXPECT_EQ(g.color_series.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(g.color_series.theme, 4U);
  EXPECT_NEAR(g.color_series.tint, -0.499984740745262, 1e-12);
  EXPECT_EQ(g.color_negative.rgb, 0xFFD00000U);
  EXPECT_EQ(g.color_axis.rgb, 0xFF000000U);
  EXPECT_EQ(g.color_markers.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(g.color_markers.indexed, 10U);
  EXPECT_EQ(g.color_first.kind, ColorSpec::Kind::kAuto);
  EXPECT_EQ(g.color_last.rgb, 0xFF0000FFU);
  EXPECT_EQ(g.color_high.rgb, 0xFF00FF00U);
  EXPECT_EQ(g.color_low.theme, 5U);
  ASSERT_EQ(g.items.size(), 2U);
  EXPECT_EQ(g.items[1].formula, "Sheet1!A2:C2");
  EXPECT_EQ(g.items[1].location, "E2");
}

TEST(SparklineFidelity, EveryAttributeAndColorSurvivesAResave) {
  const Workbook first = LoadOrDie(SparklinePackage());
  const auto groups = io::read_sparkline_groups(first.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 2U);
  ExpectFirstGroup(groups[0]);
  EXPECT_EQ(groups[1].type, io::SparklineType::kLine);
  EXPECT_DOUBLE_EQ(groups[1].line_weight, 0.75);
  EXPECT_EQ(groups[1].color_series.rgb, 0xFF376092U);

  const auto saved = SaveOrDie(first);
  const std::string sheet = ReadPart(saved, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(sheet, "rightToLeft=\"1\"")) << sheet;
  EXPECT_TRUE(Has(sheet, "xr2:uid=\"{11111111-2222-3333-4444-555555555555}\"")) << sheet;
  const Workbook second = LoadOrDie(saved);
  const auto again = io::read_sparkline_groups(second.sheet(0).ext_lst_xml());
  ASSERT_EQ(again.size(), 2U);
  ExpectFirstGroup(again[0]);
}

TEST(SparklineFidelity, ReferencesShiftWithRowInsertionAndDeletion) {
  Workbook wb = LoadOrDie(SparklinePackage());
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  auto groups = io::read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 2U);
  EXPECT_EQ(groups[0].items[0].formula, "Sheet1!A3:C3");
  EXPECT_EQ(groups[0].items[0].location, "E3");
  EXPECT_EQ(groups[0].items[1].location, "E4");
  EXPECT_EQ(groups[0].date_axis_ref, "Sheet1!A3:C3");
  EXPECT_EQ(groups[1].items[0].location, "F7");

  // Deleting both rows of the first group's sparklines drops the group.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 2, 2)));
  groups = io::read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].items[0].location, "F5");

  Workbook cols = LoadOrDie(SparklinePackage());
  ASSERT_TRUE(static_cast<bool>(cols.insert_cols(0, 0, 1)));
  groups = io::read_sparkline_groups(cols.sheet(0).ext_lst_xml());
  EXPECT_EQ(groups[0].items[0].formula, "Sheet1!B1:D1");
  EXPECT_EQ(groups[0].items[0].location, "F1");
}

TEST(SparklineFidelity, UpdateKeepsSparklinesAndUnmodelledAttributes) {
  Workbook wb = LoadOrDie(SparklinePackage());
  auto groups = io::read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  io::SparklineGroup g = groups[0];
  g.type = io::SparklineType::kStacked;
  g.markers = false;
  g.has_manual_max = false;
  g.color_series = ColorSpec{};
  g.color_high.rgb = 0xFF112233U;
  g.color_axis.kind = ColorSpec::Kind::kTheme;
  g.color_axis.theme = 1;
  g.date_axis_ref.clear();
  g.date_axis = false;
  g.line_weight = 2.25;
  std::string xml = wb.sheet(0).ext_lst_xml();
  ASSERT_TRUE(static_cast<bool>(io::update_sparkline_group(xml, 0, g)));
  wb.sheet(0).set_ext_lst_xml(xml);

  groups = io::read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 2U);
  EXPECT_EQ(groups[0].type, io::SparklineType::kStacked);
  EXPECT_FALSE(groups[0].markers);
  EXPECT_TRUE(groups[0].high);
  EXPECT_FALSE(groups[0].has_manual_max);
  EXPECT_TRUE(groups[0].has_manual_min);
  EXPECT_EQ(groups[0].color_series.kind, ColorSpec::Kind::kNone);
  EXPECT_EQ(groups[0].color_high.rgb, 0xFF112233U);
  EXPECT_EQ(groups[0].color_axis.kind, ColorSpec::Kind::kTheme);
  EXPECT_TRUE(groups[0].date_axis_ref.empty());
  EXPECT_DOUBLE_EQ(groups[0].line_weight, 2.25);
  EXPECT_EQ(groups[0].items.size(), 2U);
  EXPECT_NE(xml.find("rightToLeft=\"1\""), std::string::npos);
  EXPECT_NE(xml.find("{11111111-2222-3333-4444-555555555555}"), std::string::npos);
  EXPECT_EQ(io::read_sparkline_groups(xml)[1].items[0].location, "F5");
}

TEST(SparklineFidelity, AddReplaceAndRemoveGroups) {
  Workbook wb = LoadOrDie(BuildPackage("<sheetData/>"));
  EXPECT_TRUE(io::read_sparkline_groups(wb.sheet(0).ext_lst_xml()).empty());
  io::SparklineGroup g;
  g.type = io::SparklineType::kColumn;
  g.negative = true;
  g.color_series.kind = ColorSpec::Kind::kRgb;
  g.color_series.rgb = 0xFF336699U;
  g.color_negative.kind = ColorSpec::Kind::kTheme;
  g.color_negative.theme = 5;
  g.color_negative.tint = 0.25;
  g.items = {{"Sheet1!A1:D1", "F1"}, {"Sheet1!A2:D2", "F2"}};
  std::string xml = wb.sheet(0).ext_lst_xml();
  ASSERT_TRUE(static_cast<bool>(io::add_sparkline_group(xml, g)));
  wb.sheet(0).set_ext_lst_xml(xml);

  Workbook again = LoadOrDie(SaveOrDie(wb));
  auto groups = io::read_sparkline_groups(again.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].type, io::SparklineType::kColumn);
  EXPECT_TRUE(groups[0].negative);
  EXPECT_EQ(groups[0].color_series.rgb, 0xFF336699U);
  EXPECT_DOUBLE_EQ(groups[0].color_negative.tint, 0.25);
  ASSERT_EQ(groups[0].items.size(), 2U);

  xml = again.sheet(0).ext_lst_xml();
  g.items = {{"Sheet1!A9:D9", "G9"}};
  ASSERT_TRUE(static_cast<bool>(io::add_sparkline_group(xml, g)));
  ASSERT_TRUE(static_cast<bool>(io::set_group_sparklines(xml, 0, {{"Sheet1!B1:B5", "H1"}})));
  groups = io::read_sparkline_groups(xml);
  ASSERT_EQ(groups.size(), 2U);
  ASSERT_EQ(groups[0].items.size(), 1U);
  EXPECT_EQ(groups[0].items[0].formula, "Sheet1!B1:B5");
  EXPECT_EQ(groups[1].items[0].location, "G9");

  EXPECT_FALSE(static_cast<bool>(io::set_group_sparklines(xml, 0, {})));
  EXPECT_FALSE(static_cast<bool>(io::set_group_sparklines(xml, 0, {{"Sheet1!A1:A2", "B1:B2"}})));
  EXPECT_FALSE(static_cast<bool>(io::remove_sparkline_group(xml, 5)));
  ASSERT_TRUE(static_cast<bool>(io::remove_sparkline_group(xml, 0)));
  ASSERT_TRUE(static_cast<bool>(io::remove_sparkline_group(xml, 0)));
  EXPECT_TRUE(xml.empty());
}

// ---------------------------------------------------------------------------
// Legacy note shapes
// ---------------------------------------------------------------------------

constexpr std::string_view kCommentsType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
constexpr std::string_view kVmlType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

std::string RichNoteShape(int id, int row, int col, bool visible, std::string_view anchor, std::string_view fill) {
  return "<v:shape id=\"_x0000_s" + std::to_string(id) +
         "\" type=\"#_x0000_t202\" style='position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:" +
         std::to_string(100 + id) + "pt;height:90pt;z-index:" + std::to_string(id - 1024) +
         ";visibility:" + (visible ? "visible" : "hidden") + "' fillcolor=\"" + std::string(fill) +
         "\" stroked=\"t\" strokecolor=\"#ff0000\" o:insetmode=\"auto\"><v:fill color2=\"" + std::string(fill) +
         "\"/><v:stroke weight=\"2pt\" dashstyle=\"dash\"/><v:shadow on=\"t\" color=\"black\" obscured=\"t\"/>"
         "<v:path o:connecttype=\"none\"/><v:textbox style=\"mso-direction-alt:auto\">"
         "<div style=\"text-align:left\"><font face=\"Arial\" size=\"200\">note " +
         std::to_string(id) +
         "</font></div></v:textbox><x:ClientData ObjectType=\"Note\"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>" +
         std::string(anchor) + "</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>" + std::to_string(row) +
         "</x:Row><x:Column>" + std::to_string(col) + "</x:Column>" + (visible ? "<x:Visible/>" : "") +
         "</x:ClientData></v:shape>";
}

std::vector<std::uint8_t> NotesPackage() {
  const std::string comments =
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<authors><author>Alice</author></authors><commentList>"
      "<comment ref=\"B2\" authorId=\"0\"><text><r><t>one</t></r></text></comment>"
      "<comment ref=\"D4\" authorId=\"0\"><text><r><t>two</t></r></text></comment>"
      "</commentList></comments>";
  const std::string vml =
      "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
      "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" "
      "data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"/>" +
      RichNoteShape(1025, 1, 1, true, "3, 15, 0, 2, 6, 40, 8, 4", "#abcdef") +
      RichNoteShape(1026, 3, 3, false, "5, 15, 2, 10, 8, 20, 9, 4", "#fedcba") + "</xml>";
  const std::string rels = "<Relationship Id=\"rId1\" Type=\"" + std::string(kCommentsType) +
                           "\" Target=\"../comments1.xml\"/><Relationship Id=\"rId2\" Type=\"" + std::string(kVmlType) +
                           "\" Target=\"../drawings/vmlDrawing1.vml\"/>";
  return BuildPackage(
      "<sheetData/><legacyDrawing r:id=\"rId2\"/>", rels,
      {{"xl/comments1.xml", comments, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"},
       {"xl/drawings/vmlDrawing1.vml", vml, {}}});
}

TEST(LegacyNoteShapeFidelity, EachSurvivingNoteKeepsItsOwnShapeWhenTheSetChanges) {
  Workbook wb = LoadOrDie(NotesPackage());
  wb.sheet(0).mutable_comments().emplace_back(8, 8, "Alice", "added");
  const std::string vml = ReadPart(SaveOrDie(wb), "xl/drawings/vmlDrawing1.vml");
  for (const char* kept :
       {"fillcolor=\"#abcdef\"", "fillcolor=\"#fedcba\"", "<v:stroke weight=\"2pt\" dashstyle=\"dash\"/>",
        "stroked=\"t\" strokecolor=\"#ff0000\"", "width:1125pt", "width:1126pt",
        "<font face=\"Arial\" size=\"200\">note 1025</font>", "margin-left:59.25pt", "3, 15, 0, 2, 6, 40, 8, 4",
        "5, 15, 2, 10, 8, 20, 9, 4", "id=\"_x0000_s1025\"", "id=\"_x0000_s1026\""}) {
    EXPECT_TRUE(Has(vml, kept)) << kept << "\n" << vml;
  }
  // The new note gets the default shape and an id that collides with neither original.
  EXPECT_TRUE(Has(vml, "fillcolor=\"#ffffe1\"")) << vml;
  EXPECT_TRUE(Has(vml, "id=\"_x0000_s1027\"")) << vml;

  const Workbook again = LoadOrDie(SaveOrDie(wb));
  ASSERT_EQ(again.sheet(0).comments().size(), 3U);
  EXPECT_TRUE(again.sheet(0).comments()[0].visible);
  EXPECT_FALSE(again.sheet(0).comments()[1].visible);
  EXPECT_TRUE(Has(again.sheet(0).comments()[0].vml_shape, "fillcolor=\"#abcdef\""));
}

TEST(LegacyNoteShapeFidelity, RemovingANoteKeepsTheOtherShapesVerbatim) {
  Workbook wb = LoadOrDie(NotesPackage());
  auto& notes = wb.sheet(0).mutable_comments();
  notes.erase(notes.begin());
  const std::string vml = ReadPart(SaveOrDie(wb), "xl/drawings/vmlDrawing1.vml");
  EXPECT_FALSE(Has(vml, "#abcdef")) << vml;
  EXPECT_TRUE(Has(vml, RichNoteShape(1026, 3, 3, false, "5, 15, 2, 10, 8, 20, 9, 4", "#fedcba"))) << vml;
}

TEST(LegacyNoteShapeFidelity, OnlyAnchorVisibilityAndCellArePatched) {
  Workbook wb = LoadOrDie(NotesPackage());
  auto& notes = wb.sheet(0).mutable_comments();
  notes[0].visible = false;
  notes[1].visible = true;
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));  // B2 -> B4, D4 -> D6
  notes.emplace_back(20, 20, "Alice", "added");
  const std::string vml = ReadPart(SaveOrDie(wb), "xl/drawings/vmlDrawing1.vml");
  // Hidden: the style says so and the marker is gone; visible: the marker is added.
  EXPECT_TRUE(Has(vml, "visibility:hidden' fillcolor=\"#abcdef\"")) << vml;
  EXPECT_TRUE(Has(vml, "visibility:visible' fillcolor=\"#fedcba\"")) << vml;
  EXPECT_TRUE(Has(vml, "<x:Row>3</x:Row><x:Column>1</x:Column></x:ClientData>")) << vml;
  EXPECT_TRUE(Has(vml, "<x:Row>5</x:Row><x:Column>3</x:Column><x:Visible/></x:ClientData>")) << vml;
  // The boxes moved two rows with their notes; their offsets are unchanged.
  EXPECT_TRUE(Has(vml, "3, 15, 2, 2, 6, 40, 10, 4")) << vml;
  EXPECT_TRUE(Has(vml, "5, 15, 4, 10, 8, 20, 11, 4")) << vml;
  EXPECT_TRUE(Has(vml, "stroked=\"t\" strokecolor=\"#ff0000\"")) << vml;
}

}  // namespace
}  // namespace formulon
