//
// Integration test: worksheet view, outline, style and workbook-level
// metadata that a load -> save cycle must keep.
//   * outline gutter depth (`<sheetFormatPr outlineLevelRow/Col>`) derived
//     from the row / column levels, and `collapsed` on `<row>` / `<col>`
//   * every `<sheetView>` / `<selection>` attribute and child
//   * gradient fills and the outline / shadow / condense / extend font toggles
//   * custom icon sets (`<x14:cfIcon>`) of conditional-format rules
//   * `<calcPr>` attributes
//   * `comment` / `hidden` on defined names

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
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

/// One-sheet package. `workbook_extra` lands after `<sheets>`; `styles` is
/// omitted from the package when empty.
std::vector<std::uint8_t> BuildPackage(std::string_view sheet_xml, std::string_view workbook_extra = {},
                                       std::string_view styles = {}) {
  const std::string content_types =
      std::string(kXmlDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>" +
      std::string(styles.empty() ? ""
                                 : "<Override PartName=\"/xl/styles.xml\" "
                                   "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml."
                                   "styles+xml\"/>") +
      "</Types>";
  const std::string package_rels =
      std::string(kXmlDecl) +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/></Relationships>";
  const std::string workbook_xml = std::string(kXmlDecl) +
                                   "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                                   "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                                   "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                                   std::string(workbook_extra) + "</workbook>";
  const std::string workbook_rels =
      std::string(kXmlDecl) +
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>" +
      std::string(styles.empty() ? ""
                                 : "<Relationship Id=\"rId2\" "
                                   "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                                   "styles\" Target=\"styles.xml\"/>") +
      "</Relationships>";
  std::vector<PartFile> parts = {
      {"[Content_Types].xml", content_types},  {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},       {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
  };
  if (!styles.empty()) {
    parts.push_back({"xl/styles.xml", styles});
  }
  return BuildZip(parts);
}

std::string Worksheet(std::string_view body, std::string_view extra_ns = {}) {
  return std::string(kXmlDecl) +
         "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
         "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
         std::string(extra_ns) + ">" + std::string(body) + "</worksheet>";
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

/// Loads `bytes`, saves it and returns the saved package.
std::vector<std::uint8_t> Resave(const std::vector<std::uint8_t>& bytes) {
  return SaveOrDie(LoadOrDie(bytes));
}

bool Has(const std::string& text, std::string_view needle) {
  return text.find(needle) != std::string::npos;
}

// ---------------------------------------------------------------------------
// Outline
// ---------------------------------------------------------------------------

TEST(OutlineFidelity, GutterDepthIsDerivedFromRowAndColumnLevels) {
  Workbook wb = Workbook::create();
  SheetLayout& layout = wb.sheet(0).mutable_layout();
  RowLayout shallow;
  shallow.row = 1;
  shallow.outline_level = 1;
  RowLayout deep;
  deep.row = 2;
  deep.outline_level = 3;
  layout.row_overrides = {shallow, deep};
  ColumnLayout col;
  col.first = 4;
  col.last = 5;
  col.outline_level = 2;
  layout.columns = {col};

  const std::string xml = ReadPart(SaveOrDie(wb), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "outlineLevelRow=\"3\"")) << xml;
  EXPECT_TRUE(Has(xml, "outlineLevelCol=\"2\"")) << xml;
}

TEST(OutlineFidelity, NoLevelsMeansNoGutterAttributes) {
  Workbook wb = Workbook::create();
  const std::string xml = ReadPart(SaveOrDie(wb), "xl/worksheets/sheet1.xml");
  EXPECT_FALSE(Has(xml, "outlineLevelRow")) << xml;
  EXPECT_FALSE(Has(xml, "outlineLevelCol")) << xml;
}

TEST(OutlineFidelity, StaleGutterDepthIsReplacedByTheDerivedOne) {
  const std::string sheet = Worksheet(
      "<sheetFormatPr defaultRowHeight=\"15\" outlineLevelRow=\"5\" outlineLevelCol=\"4\"/>"
      "<sheetData><row r=\"2\" outlineLevel=\"1\"><c r=\"A2\"><v>1</v></c></row></sheetData>");
  const std::string xml = ReadPart(Resave(BuildPackage(sheet)), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "outlineLevelRow=\"1\"")) << xml;
  EXPECT_FALSE(Has(xml, "outlineLevelCol")) << xml;
}

TEST(OutlineFidelity, CollapsedRoundTripsOnRowsAndColumns) {
  const std::string sheet = Worksheet(
      "<sheetFormatPr defaultRowHeight=\"15\" outlineLevelRow=\"1\" outlineLevelCol=\"1\"/>"
      "<cols><col min=\"2\" max=\"3\" width=\"9\" customWidth=\"1\" outlineLevel=\"1\" hidden=\"1\"/>"
      "<col min=\"4\" max=\"4\" width=\"9\" customWidth=\"1\" collapsed=\"1\"/></cols>"
      "<sheetData><row r=\"2\" outlineLevel=\"1\" hidden=\"1\"><c r=\"A2\"><v>1</v></c></row>"
      "<row r=\"3\" collapsed=\"1\"><c r=\"A3\"><v>2</v></c></row></sheetData>");
  const std::vector<std::uint8_t> bytes = BuildPackage(sheet);

  Workbook loaded = LoadOrDie(bytes);
  const SheetLayout& layout = loaded.sheet(0).layout();
  ASSERT_EQ(layout.columns.size(), 2U);
  EXPECT_FALSE(layout.columns[0].collapsed);
  EXPECT_TRUE(layout.columns[1].collapsed);
  ASSERT_EQ(layout.row_overrides.size(), 2U);
  EXPECT_FALSE(layout.row_overrides[0].collapsed);
  EXPECT_TRUE(layout.row_overrides[1].collapsed);

  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<col min=\"4\" max=\"4\"")) << xml;
  EXPECT_TRUE(Has(xml, "collapsed=\"1\"")) << xml;
  Workbook again = LoadOrDie(SaveOrDie(loaded));
  EXPECT_TRUE(again.sheet(0).layout().columns[1].collapsed);
  EXPECT_TRUE(again.sheet(0).layout().row_overrides[1].collapsed);
}

TEST(OutlineFidelity, DefaultRowHeightFlagsSurviveAResave) {
  const std::string sheet = Worksheet(
      "<sheetFormatPr defaultRowHeight=\"18\" customHeight=\"1\" zeroHeight=\"1\"/>"
      "<sheetData><row r=\"2\" hidden=\"0\"><c r=\"A2\"><v>1</v></c></row></sheetData>");
  Workbook loaded = LoadOrDie(BuildPackage(sheet));
  EXPECT_TRUE(loaded.sheet(0).format_defaults().custom_height);
  EXPECT_TRUE(loaded.sheet(0).format_defaults().zero_height);
  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "customHeight=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "zeroHeight=\"1\"")) << xml;
}

// ---------------------------------------------------------------------------
// Sheet view
// ---------------------------------------------------------------------------

TEST(SheetViewFidelity, EveryViewAttributeSurvives) {
  const std::string sheet = Worksheet(
      "<sheetViews><sheetView windowProtection=\"1\" showFormulas=\"1\" showGridLines=\"0\" "
      "showRowColHeaders=\"0\" showZeros=\"0\" rightToLeft=\"1\" tabSelected=\"1\" showRuler=\"0\" "
      "showOutlineSymbols=\"0\" defaultGridColor=\"0\" showWhiteSpace=\"0\" view=\"pageLayout\" "
      "topLeftCell=\"B2\" colorId=\"23\" zoomScale=\"130\" zoomScaleNormal=\"110\" "
      "zoomScaleSheetLayoutView=\"90\" zoomScalePageLayoutView=\"70\" workbookViewId=\"0\">"
      "<selection activeCell=\"C3\" activeCellId=\"1\" sqref=\"A1:B2 C3:D4\"/>"
      "</sheetView></sheetViews><sheetData/>");
  Workbook loaded = LoadOrDie(BuildPackage(sheet));
  const SheetView& view = loaded.sheet(0).view();
  EXPECT_TRUE(view.window_protection);
  EXPECT_TRUE(view.show_formulas);
  EXPECT_FALSE(view.show_ruler);
  EXPECT_FALSE(view.show_outline_symbols);
  EXPECT_FALSE(view.default_grid_color);
  EXPECT_FALSE(view.show_white_space);
  EXPECT_EQ(view.color_id, 23U);
  EXPECT_EQ(view.zoom_scale_normal, 110U);
  EXPECT_EQ(view.zoom_scale_sheet_layout_view, 90U);
  EXPECT_EQ(view.zoom_scale_page_layout_view, 70U);
  ASSERT_EQ(view.selections.size(), 1U);
  EXPECT_EQ(view.selections[0].active_cell_id, 1U);

  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  for (std::string_view expected : {"windowProtection=\"1\"",
                                    "showFormulas=\"1\"",
                                    "showGridLines=\"0\"",
                                    "showRowColHeaders=\"0\"",
                                    "showZeros=\"0\"",
                                    "rightToLeft=\"1\"",
                                    "tabSelected=\"1\"",
                                    "showRuler=\"0\"",
                                    "showOutlineSymbols=\"0\"",
                                    "defaultGridColor=\"0\"",
                                    "showWhiteSpace=\"0\"",
                                    "view=\"pageLayout\"",
                                    "topLeftCell=\"B2\"",
                                    "colorId=\"23\"",
                                    "zoomScale=\"130\"",
                                    "zoomScaleNormal=\"110\"",
                                    "zoomScaleSheetLayoutView=\"90\"",
                                    "zoomScalePageLayoutView=\"70\"",
                                    "activeCellId=\"1\"",
                                    "sqref=\"A1:B2 C3:D4\""}) {
    EXPECT_TRUE(Has(xml, expected)) << expected << "\n" << xml;
  }
}

TEST(SheetViewFidelity, DefaultViewEmitsNoViewElement) {
  Workbook wb = Workbook::create();
  const std::string xml = ReadPart(SaveOrDie(wb), "xl/worksheets/sheet1.xml");
  EXPECT_FALSE(Has(xml, "<sheetViews>")) << xml;
}

TEST(SheetViewFidelity, PivotSelectionAndExtensionListAreKept) {
  const std::string sheet = Worksheet(
      "<sheetViews><sheetView workbookViewId=\"0\">"
      "<selection activeCell=\"B2\" sqref=\"B2\"/>"
      "<pivotSelection pane=\"bottomRight\" showHeader=\"1\" label=\"1\" "
      "axis=\"axisRow\" dimensionPivot=\"1\"/>"
      "<extLst><ext uri=\"{FA7FD5E5-0000-0000-0000-000000000000}\" "
      "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:ignore/></ext></extLst>"
      "</sheetView></sheetViews><sheetData/>");
  const std::string xml = ReadPart(Resave(BuildPackage(sheet)), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<pivotSelection pane=\"bottomRight\"")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:ignore/>")) << xml;
  EXPECT_LT(xml.find("<selection"), xml.find("<pivotSelection")) << xml;
  EXPECT_LT(xml.find("<pivotSelection"), xml.find("<extLst>")) << xml;
}

// ---------------------------------------------------------------------------
// Styles: gradient fills and font toggles
// ---------------------------------------------------------------------------

constexpr std::string_view kStylesWithGradients =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<fonts count=\"2\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font>"
    "<font><outline/><shadow val=\"0\"/><condense/><extend val=\"1\"/><sz val=\"11\"/><name val=\"Arial\"/></font>"
    "</fonts>"
    "<fills count=\"4\"><fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"gray125\"/></fill>"
    "<fill><gradientFill degree=\"90\"><stop position=\"0\"><color rgb=\"FFFF0000\"/></stop>"
    "<stop position=\"1\"><color theme=\"4\" tint=\"0.5\"/></stop></gradientFill></fill>"
    "<fill><gradientFill type=\"path\" left=\"0.5\" right=\"0.5\" top=\"0.25\" bottom=\"0.75\">"
    "<stop position=\"0\"><color indexed=\"10\"/></stop><stop position=\"0.5\"><color rgb=\"FF00FF00\"/></stop>"
    "<stop position=\"1\"><color rgb=\"FF0000FF\"/></stop></gradientFill></fill></fills>"
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
    "<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>"
    "<cellXfs count=\"3\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/>"
    "<xf numFmtId=\"0\" fontId=\"1\" fillId=\"2\" borderId=\"0\" xfId=\"0\" applyFill=\"1\" applyFont=\"1\"/>"
    "<xf numFmtId=\"0\" fontId=\"0\" fillId=\"3\" borderId=\"0\" xfId=\"0\" applyFill=\"1\"/></cellXfs>"
    "<dxfs count=\"1\"><dxf><font><shadow/></font><fill><gradientFill degree=\"45\">"
    "<stop position=\"0\"><color rgb=\"FF112233\"/></stop><stop position=\"1\"><color rgb=\"FF445566\"/></stop>"
    "</gradientFill></fill></dxf></dxfs>"
    "</styleSheet>";

TEST(StyleFidelity, GradientFillsAreModelled) {
  Workbook loaded = LoadOrDie(BuildPackage(Worksheet("<sheetData/>"), {}, kStylesWithGradients));
  const StylesTable& styles = loaded.styles();
  ASSERT_EQ(styles.fills.size(), 4U);
  const FillRecord& linear = styles.fills[2];
  ASSERT_TRUE(linear.has_gradient);
  EXPECT_FALSE(linear.gradient.path);
  EXPECT_EQ(linear.gradient.degree, 90.0);
  ASSERT_EQ(linear.gradient.stops.size(), 2U);
  EXPECT_EQ(linear.gradient.stops[0].position, 0.0);
  EXPECT_EQ(linear.gradient.stops[0].color.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(linear.gradient.stops[0].argb, 0xFFFF0000U);
  EXPECT_EQ(linear.gradient.stops[1].color.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(linear.gradient.stops[1].color.theme, 4U);
  EXPECT_EQ(linear.gradient.stops[1].color.tint, 0.5);

  const FillRecord& path = styles.fills[3];
  ASSERT_TRUE(path.has_gradient);
  EXPECT_TRUE(path.gradient.path);
  EXPECT_EQ(path.gradient.left, 0.5);
  EXPECT_EQ(path.gradient.top, 0.25);
  EXPECT_EQ(path.gradient.bottom, 0.75);
  ASSERT_EQ(path.gradient.stops.size(), 3U);
  EXPECT_EQ(path.gradient.stops[0].color.kind, ColorSpec::Kind::kIndexed);

  ASSERT_EQ(styles.dxfs.size(), 1U);
  ASSERT_TRUE(styles.dxfs[0].has_fill);
  EXPECT_TRUE(styles.dxfs[0].fill.has_gradient);
  EXPECT_EQ(styles.dxfs[0].fill.gradient.degree, 45.0);
}

TEST(StyleFidelity, GradientFillsSurviveResave) {
  const std::string xml =
      ReadPart(Resave(BuildPackage(Worksheet("<sheetData/>"), {}, kStylesWithGradients)), "xl/styles.xml");
  EXPECT_TRUE(Has(xml,
                  "<gradientFill degree=\"90\"><stop position=\"0\"><color rgb=\"FFFF0000\"/></stop>"
                  "<stop position=\"1\"><color theme=\"4\" tint=\"0.5\"/></stop></gradientFill>"))
      << xml;
  EXPECT_TRUE(Has(xml, "<gradientFill type=\"path\" left=\"0.5\" right=\"0.5\" top=\"0.25\" bottom=\"0.75\">")) << xml;
  EXPECT_TRUE(Has(xml, "<stop position=\"0.5\"><color rgb=\"FF00FF00\"/></stop>")) << xml;
  EXPECT_TRUE(Has(xml, "<dxf><font><shadow/></font><fill><gradientFill degree=\"45\">")) << xml;
}

TEST(StyleFidelity, FontTogglesAreModelledAndSaved) {
  Workbook loaded = LoadOrDie(BuildPackage(Worksheet("<sheetData/>"), {}, kStylesWithGradients));
  ASSERT_GE(loaded.styles().fonts.size(), 2U);
  const FontRecord& font = loaded.styles().fonts[1];
  EXPECT_TRUE(font.has_outline);
  EXPECT_TRUE(font.outline);
  EXPECT_TRUE(font.has_shadow);
  EXPECT_FALSE(font.shadow);
  EXPECT_TRUE(font.condense);
  EXPECT_TRUE(font.extend);
  EXPECT_FALSE(loaded.styles().fonts[0].has_outline);
  EXPECT_TRUE(loaded.styles().dxfs[0].font.shadow);

  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/styles.xml");
  EXPECT_TRUE(Has(xml, "<font><condense/><extend/><outline/><shadow val=\"0\"/><sz val=\"11\"/>")) << xml;
}

// ---------------------------------------------------------------------------
// Conditional formatting: custom icon sets
// ---------------------------------------------------------------------------

constexpr std::string_view kX14Ns =
    " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\""
    " xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"";

std::string CustomIconSheet(std::string_view icons) {
  return Worksheet(
      "<sheetData/><conditionalFormatting sqref=\"A1:A5\"><cfRule type=\"iconSet\" priority=\"1\">"
      "<iconSet iconSet=\"3Arrows\"><cfvo type=\"percent\" val=\"0\"/><cfvo type=\"percent\" val=\"33\"/>"
      "<cfvo type=\"percent\" val=\"67\"/></iconSet>"
      "<extLst><ext uri=\"{B025F937-C7B1-47D3-B67F-A62EFF666E3E}\"><x14:id>{AAAAAAAA-0000-0000-0000-000000000001}"
      "</x14:id></ext></extLst></cfRule></conditionalFormatting>"
      "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\"><x14:conditionalFormattings>"
      "<x14:conditionalFormatting><x14:cfRule type=\"iconSet\" priority=\"1\" "
      "id=\"{AAAAAAAA-0000-0000-0000-000000000001}\"><x14:iconSet iconSet=\"3Arrows\" custom=\"1\">"
      "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"percent\"><xm:f>33</xm:f></x14:cfvo>"
      "<x14:cfvo type=\"percent\"><xm:f>67</xm:f></x14:cfvo>" +
          std::string(icons) +
          "</x14:iconSet></x14:cfRule><xm:sqref>A1:A5</xm:sqref></x14:conditionalFormatting>"
          "</x14:conditionalFormattings></ext></extLst>",
      kX14Ns);
}

constexpr std::string_view kThreeIcons =
    "<x14:cfIcon iconSet=\"3Triangles\" iconId=\"0\"/><x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/>"
    "<x14:cfIcon iconSet=\"3Stars\" iconId=\"2\"/>";

cf::IconSetSpec& IconSpec(Workbook& wb) {
  return *wb.sheet(0).mutable_conditional_formats().at(0).rules.at(0).icon_set;
}

TEST(CustomIconFidelity, CfIconsAreReadFromTheLinkedX14Rule) {
  Workbook loaded = LoadOrDie(BuildPackage(CustomIconSheet(kThreeIcons)));
  const cf::IconSetSpec& spec = IconSpec(loaded);
  ASSERT_EQ(spec.custom_icons.size(), 3U);
  EXPECT_EQ(spec.custom_icons[0].icon_set, "3Triangles");
  EXPECT_EQ(spec.custom_icons[0].icon_id, 0);
  EXPECT_EQ(spec.custom_icons[1].icon_set, "NoIcons");
  EXPECT_EQ(spec.custom_icons[2].icon_set, "3Stars");
  EXPECT_EQ(spec.custom_icons[2].icon_id, 2);
}

TEST(CustomIconFidelity, CfIconsSurviveResave) {
  const std::string xml = ReadPart(Resave(BuildPackage(CustomIconSheet(kThreeIcons))), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"3Triangles\" iconId=\"0\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"NoIcons\" iconId=\"0\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"3Stars\" iconId=\"2\"/>")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:id>{AAAAAAAA-0000-0000-0000-000000000001}</x14:id>")) << xml;
}

TEST(CustomIconFidelity, EditedIconsReplaceTheCapturedOnes) {
  Workbook loaded = LoadOrDie(BuildPackage(CustomIconSheet(kThreeIcons)));
  IconSpec(loaded).custom_icons = {{"5Boxes", 4}, {"5Boxes", 0}, {"NoIcons", 0}};
  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"5Boxes\" iconId=\"4\"/>")) << xml;
  EXPECT_FALSE(Has(xml, "3Triangles")) << xml;
}

TEST(CustomIconFidelity, ClearedIconsLeaveAPlainIconSet) {
  Workbook loaded = LoadOrDie(BuildPackage(CustomIconSheet(kThreeIcons)));
  IconSpec(loaded).custom_icons.clear();
  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  EXPECT_FALSE(Has(xml, "x14:cfIcon")) << xml;
  Workbook again = LoadOrDie(SaveOrDie(loaded));
  EXPECT_TRUE(IconSpec(again).custom_icons.empty());
}

TEST(CustomIconFidelity, AnAuthoredRuleGetsAnX14Counterpart) {
  Workbook wb = Workbook::create();
  cf::ConditionalFormat block;
  block.sqref = {cf::CFCellRange{{0, 0}, {4, 0}}};
  cf::CFRule rule;
  rule.type = cf::RuleType::IconSet;
  rule.priority = 1;
  rule.id = "{FC000000-0000-0000-0000-000000000001}";
  cf::IconSetSpec spec;
  spec.name = cf::IconSetName::Three_Arrows;
  spec.thresholds = {{cf::CfvoType::Percent, "33", true}, {cf::CfvoType::Percent, "67", true}};
  spec.custom_icons = {{"3Flags", 0}, {"3Flags", 1}, {"3Flags", 2}};
  rule.icon_set = spec;
  block.rules.push_back(rule);
  wb.sheet(0).mutable_conditional_formats().push_back(block);

  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);
  const std::string xml = ReadPart(bytes, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<x14:iconSet iconSet=\"3Arrows\" custom=\"1\">")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:cfIcon iconSet=\"3Flags\" iconId=\"2\"/>")) << xml;
  Workbook again = LoadOrDie(bytes);
  ASSERT_EQ(IconSpec(again).custom_icons.size(), 3U);
  EXPECT_EQ(IconSpec(again).custom_icons[1].icon_set, "3Flags");
}

// Excel stores 3Stars / 3Triangles / 5Boxes only in the x14 extension. The
// model has no such set, so the rule stays there alone instead of gaining a
// legacy traffic-light twin, and the model reports no set the file lacks.
TEST(CustomIconFidelity, AnX14OnlySetTheModelLacksStaysInTheExtensionAlone) {
  const std::string sheet = Worksheet(
      "<sheetData/><extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\"><x14:conditionalFormattings>"
      "<x14:conditionalFormatting><x14:cfRule type=\"iconSet\" priority=\"1\" "
      "id=\"{AAAAAAAA-0000-0000-0000-000000000002}\"><x14:iconSet iconSet=\"3Stars\">"
      "<x14:cfvo type=\"percent\"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type=\"percent\"><xm:f>33</xm:f></x14:cfvo>"
      "<x14:cfvo type=\"percent\"><xm:f>67</xm:f></x14:cfvo></x14:iconSet></x14:cfRule>"
      "<xm:sqref>A1:A3</xm:sqref></x14:conditionalFormatting></x14:conditionalFormattings></ext></extLst>",
      kX14Ns);
  Workbook loaded = LoadOrDie(BuildPackage(sheet));
  for (const cf::ConditionalFormat& format : loaded.sheet(0).conditional_formats()) {
    for (const cf::CFRule& rule : format.rules) {
      EXPECT_FALSE(rule.icon_set.has_value()) << "the model reports an icon set the file does not name";
    }
  }
  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/worksheets/sheet1.xml");
  std::size_t rules = 0;
  for (std::size_t at = xml.find("type=\"iconSet\""); at != std::string::npos;
       at = xml.find("type=\"iconSet\"", at + 1U)) {
    ++rules;
  }
  EXPECT_EQ(rules, 1U) << xml;
  EXPECT_FALSE(Has(xml, "<conditionalFormatting")) << xml;
  EXPECT_TRUE(Has(xml, "<x14:iconSet iconSet=\"3Stars\">")) << xml;
}

// ---------------------------------------------------------------------------
// calcPr
// ---------------------------------------------------------------------------

TEST(CalcPrFidelity, EveryAttributeSurvivesAResave) {
  const std::string extra =
      "<calcPr calcId=\"191029\" calcMode=\"manual\" fullCalcOnLoad=\"1\" refMode=\"R1C1\" iterate=\"1\" "
      "iterateCount=\"50\" iterateDelta=\"0.01\" fullPrecision=\"0\" calcCompleted=\"0\" calcOnSave=\"0\" "
      "concurrentCalc=\"0\" concurrentManualCount=\"4\" forceFullCalc=\"1\"/>";
  Workbook loaded = LoadOrDie(BuildPackage(Worksheet("<sheetData/>"), extra));
  const CalcOptions& calc = loaded.calc_options();
  EXPECT_FALSE(calc.full_precision);
  EXPECT_FALSE(calc.calc_on_save);
  EXPECT_EQ(calc.ref_mode, RefMode::kR1C1);
  EXPECT_TRUE(calc.has_calc_id);
  EXPECT_EQ(calc.calc_id, 191029U);
  EXPECT_FALSE(calc.concurrent_calc);
  EXPECT_TRUE(calc.has_concurrent_manual_count);
  EXPECT_EQ(calc.concurrent_manual_count, 4U);
  EXPECT_TRUE(calc.force_full_calc);
  EXPECT_TRUE(calc.full_calc_on_load);
  EXPECT_FALSE(calc.calc_completed);

  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/workbook.xml");
  for (std::string_view expected :
       {"calcId=\"191029\"", "calcMode=\"manual\"", "fullCalcOnLoad=\"1\"", "refMode=\"R1C1\"", "iterate=\"1\"",
        "iterateCount=\"50\"", "iterateDelta=\"0.01\"", "fullPrecision=\"0\"", "calcCompleted=\"0\"",
        "calcOnSave=\"0\"", "concurrentCalc=\"0\"", "concurrentManualCount=\"4\"", "forceFullCalc=\"1\""}) {
    EXPECT_TRUE(Has(xml, expected)) << expected << "\n" << xml;
  }
}

TEST(CalcPrFidelity, DefaultsEmitNoElement) {
  Workbook wb = Workbook::create();
  EXPECT_FALSE(Has(ReadPart(SaveOrDie(wb), "xl/workbook.xml"), "<calcPr"));
}

TEST(CalcPrFidelity, TypedSettersReachTheFile) {
  Workbook wb = Workbook::create();
  wb.mutable_calc_options().full_precision = false;
  wb.mutable_calc_options().ref_mode = RefMode::kR1C1;
  const std::string xml = ReadPart(SaveOrDie(wb), "xl/workbook.xml");
  EXPECT_TRUE(Has(xml, "fullPrecision=\"0\"")) << xml;
  EXPECT_TRUE(Has(xml, "refMode=\"R1C1\"")) << xml;
}

// ---------------------------------------------------------------------------
// Defined names
// ---------------------------------------------------------------------------

TEST(WorkbookChildFidelity, UnmodelledChildrenSurviveInSchemaOrder) {
  const std::string extra =
      "<calcPr calcId=\"191029\" iterate=\"1\"/><oleSize ref=\"A1:D10\"/>"
      "<customWorkbookViews><customWorkbookView name=\"Mine\" guid=\"{11111111-2222-3333-4444-555555555555}\" "
      "windowWidth=\"800\" windowHeight=\"600\" activeSheetId=\"1\"/></customWorkbookViews>"
      "<fileRecoveryPr repairLoad=\"1\"/>";
  const std::string xml = ReadPart(Resave(BuildPackage(Worksheet("<sheetData/>"), extra)), "xl/workbook.xml");
  const std::size_t calc = xml.find("<calcPr");
  const std::size_t ole = xml.find("<oleSize ref=\"A1:D10\"/>");
  const std::size_t views = xml.find("<customWorkbookViews>");
  const std::size_t recovery = xml.find("<fileRecoveryPr repairLoad=\"1\"/>");
  ASSERT_NE(calc, std::string::npos) << xml;
  ASSERT_NE(ole, std::string::npos) << xml;
  ASSERT_NE(views, std::string::npos) << xml;
  ASSERT_NE(recovery, std::string::npos) << xml;
  EXPECT_LT(calc, ole) << xml;
  EXPECT_LT(ole, views) << xml;
  EXPECT_LT(views, recovery) << xml;
  EXPECT_LT(recovery, xml.find("</workbook>")) << xml;
}

TEST(DataValidationFidelity, UnmodelledAttributesSurviveAResave) {
  const std::string sheet = Worksheet(
      "<sheetData/><dataValidations count=\"1\" disablePrompts=\"1\" xWindow=\"10\" yWindow=\"20\">"
      "<dataValidation type=\"whole\" imeMode=\"hiragana\" allowBlank=\"1\" sqref=\"A1\">"
      "<formula1>1</formula1></dataValidation></dataValidations>");
  const std::string xml = ReadPart(Resave(BuildPackage(sheet)), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "disablePrompts=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "xWindow=\"10\"")) << xml;
  EXPECT_TRUE(Has(xml, "yWindow=\"20\"")) << xml;
  EXPECT_TRUE(Has(xml, "imeMode=\"hiragana\"")) << xml;
}

TEST(DefinedNameFidelity, UnmodelledAttributesSurviveAResave) {
  const std::string extra =
      "<definedNames><definedName name=\"Macro1\" function=\"1\" vbProcedure=\"1\" shortcutKey=\"q\" "
      "description=\"Runs &amp; checks\">Sheet1!$A$1</definedName></definedNames>";
  const std::string xml = ReadPart(Resave(BuildPackage(Worksheet("<sheetData/>"), extra)), "xl/workbook.xml");
  EXPECT_TRUE(Has(xml, "function=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "vbProcedure=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "shortcutKey=\"q\"")) << xml;
  EXPECT_TRUE(Has(xml, "description=\"Runs &amp; checks\"")) << xml;
}

TEST(DefinedNameFidelity, CommentAndHiddenSurviveAndStayThroughAReplacement) {
  const std::string extra =
      "<definedNames><definedName name=\"Rate\" comment=\"Tax rate\" hidden=\"1\">Sheet1!$A$1</definedName>"
      "</definedNames>";
  Workbook loaded = LoadOrDie(BuildPackage(Worksheet("<sheetData/>"), extra));
  ASSERT_EQ(loaded.defined_names().size(), 1U);
  EXPECT_EQ(loaded.defined_names()[0].comment, "Tax rate");
  EXPECT_TRUE(loaded.defined_names()[0].hidden);

  ASSERT_TRUE(static_cast<bool>(loaded.set_defined_name_scoped("rate", "Sheet1!$B$2", -1)));
  EXPECT_EQ(loaded.defined_names()[0].comment, "Tax rate");
  EXPECT_TRUE(loaded.defined_names()[0].hidden);

  ASSERT_TRUE(static_cast<bool>(loaded.set_defined_name_comment("RATE", -1, "New note")));
  const std::string xml = ReadPart(SaveOrDie(loaded), "xl/workbook.xml");
  EXPECT_TRUE(Has(xml, "comment=\"New note\"")) << xml;
  EXPECT_TRUE(Has(xml, "hidden=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "Sheet1!$B$2")) << xml;
}

}  // namespace
}  // namespace formulon
