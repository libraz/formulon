// `Workbook::create_for_profile` checked against the empty workbooks Mac
// Excel 365 saves in each UI locale. Every expected value below is copied
// from those captures, never from engine output.

#include <cstdint>
#include <iterator>
#include <string>
#include <vector>

#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "pugixml.hpp"
#include "sheet_passthrough.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace {

/// What Excel writes into a new workbook for one locale.
struct Captured {
  const char* profile_id;
  const char* font_name;
  int charset;                ///< 0 when font 0 carries no `<charset>`.
  int phonetic_size;          ///< 0 when there is no phonetic font.
  const char* phonetic_type;  ///< `phoneticPr type`, nullptr when absent or no phoneticPr.
  const char* normal_name;
  double row_height;
  const char* sheet_name;
  double margin_lr;
  double margin_tb;
  double margin_hf;
};

constexpr double kTb787 = 0.78740157499999996;

const Captured kMac[] = {
    {"mac-365-en_US", "Aptos Narrow", 0, 0, nullptr, "Normal", 16, "Sheet1", 0.7, 0.75, 0.3},
    {"mac-365-de_DE", "Aptos Narrow", 0, 0, nullptr, "Standard", 16, "Tabelle1", 0.7, kTb787, 0.3},
    {"mac-365-fr_FR", "Aptos Narrow", 0, 0, nullptr, "Normal", 16, "Feuil1", 0.7, 0.75, 0.3},
    {"mac-365-es_ES", "Aptos Narrow", 0, 0, nullptr, "Normal", 16, "Hoja1", 0.7, 0.75, 0.3},
    {"mac-365-es_MX", "Aptos Narrow", 0, 0, nullptr, "Normal", 16, "Hoja1", 0.7, 0.75, 0.3},
    {"mac-365-it_IT", "Aptos Narrow", 0, 0, nullptr, "Normale", 16, "Foglio1", 0.7, 0.75, 0.3},
    {"mac-365-nl_NL", "Aptos Narrow", 0, 0, nullptr, "Standaard", 16, "Blad1", 0.7, 0.75, 0.3},
    {"mac-365-pt_BR", "Aptos Narrow", 0, 0, nullptr, "Normal", 16, "Planilha1", 0.511811024, kTb787,
     0.31496062000000002},
    {"mac-365-ru_RU", "Aptos Narrow", 204, 0, nullptr, "Обычный", 16, "Лист1", 0.7, 0.75, 0.3},
    {"mac-365-th_TH", "Tahoma", 222, 0, nullptr, "ปกติ", 15, "Sheet1", 0.7, 0.75, 0.3},
    {"mac-365-ja_JP", "游ゴシック", 128, 6, nullptr, "標準", 20, "Sheet1", 0.7, 0.75, 0.3},
    {"mac-365-ko_KR", "맑은 고딕", 129, 8, "noConversion", "표준", 18, "Sheet1", 0.7, 0.75, 0.3},
    {"mac-365-zh_CN", "等线", 134, 9, "noConversion", "常规", 16, "Sheet1", 0.7, 0.75, 0.3},
    {"mac-365-zh_TW", "新細明體", 136, 9, "noConversion", "一般", 15, "工作表1", 0.7, 0.75, 0.3},
};

ExcelProfile ProfileOf(const char* id) {
  ExcelProfile profile;
  EXPECT_TRUE(parse_excel_profile_id(id, &profile)) << id;
  return profile;
}

std::vector<std::uint8_t> Save(const Workbook& wb) {
  auto bytes = wb.save();
  EXPECT_TRUE(static_cast<bool>(bytes));
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>();
}

void LoadPart(const std::vector<std::uint8_t>& pkg, const char* name, pugi::xml_document* doc) {
  std::string xml;
  ASSERT_TRUE(test::extract_part(io::ByteSpan{pkg.data(), pkg.size()}, name, &xml));
  ASSERT_TRUE(test::parse_xml(xml, doc));
}

pugi::xml_node Nth(const pugi::xml_node& parent, const char* name, std::size_t index) {
  std::size_t i = 0;
  for (pugi::xml_node child : parent.children(name)) {
    if (i++ == index) {
      return child;
    }
  }
  return pugi::xml_node();
}

/// Checks the written `styles.xml` against one capture row (`<color>` of the
/// phonetic font is not compared: the writer always emits one).
void ExpectStylesXml(const pugi::xml_document& doc, const Captured& c, double body_size) {
  const pugi::xml_node sheet = doc.child("styleSheet");
  const pugi::xml_node fonts = sheet.child("fonts");
  const std::size_t font_count =
      static_cast<std::size_t>(std::distance(fonts.children("font").begin(), fonts.children("font").end()));
  EXPECT_EQ(font_count, c.phonetic_size != 0 ? 2U : 1U);

  const pugi::xml_node font0 = Nth(fonts, "font", 0);
  EXPECT_DOUBLE_EQ(font0.child("sz").attribute("val").as_double(), body_size);
  EXPECT_STREQ(font0.child("color").attribute("theme").value(), "1");
  EXPECT_STREQ(font0.child("name").attribute("val").value(), c.font_name);
  EXPECT_STREQ(font0.child("family").attribute("val").value(), "2");
  if (c.charset != 0) {
    EXPECT_EQ(font0.child("charset").attribute("val").as_int(), c.charset);
  } else {
    EXPECT_FALSE(font0.child("charset"));
  }
  EXPECT_STREQ(font0.child("scheme").attribute("val").value(), "minor");

  if (c.phonetic_size != 0) {
    const pugi::xml_node font1 = Nth(fonts, "font", 1);
    EXPECT_DOUBLE_EQ(font1.child("sz").attribute("val").as_double(), c.phonetic_size);
    EXPECT_STREQ(font1.child("name").attribute("val").value(), c.font_name);
    EXPECT_STREQ(font1.child("family").attribute("val").value(), "2");
    EXPECT_EQ(font1.child("charset").attribute("val").as_int(), c.charset);
    EXPECT_STREQ(font1.child("scheme").attribute("val").value(), "minor");
  }

  for (const char* section : {"cellStyleXfs", "cellXfs"}) {
    const pugi::xml_node xf = sheet.child(section).child("xf");
    ASSERT_TRUE(xf) << section;
    EXPECT_FALSE(xf.attribute("applyAlignment")) << section;
    if (c.phonetic_size != 0) {
      EXPECT_STREQ(xf.child("alignment").attribute("vertical").value(), "center") << section;
    } else {
      EXPECT_FALSE(xf.child("alignment")) << section;
    }
  }

  const pugi::xml_node style = sheet.child("cellStyles").child("cellStyle");
  EXPECT_STREQ(style.attribute("name").value(), c.normal_name);
  EXPECT_STREQ(style.attribute("builtinId").value(), "0");
}

/// Checks the written `sheet1.xml`; `mac` selects the Mac row
/// height and base column width, which the Windows estimate leaves unset.
void ExpectSheetXml(const pugi::xml_document& doc, const Captured& c, bool mac) {
  const pugi::xml_node ws = doc.child("worksheet");
  const pugi::xml_node fmt = ws.child("sheetFormatPr");
  if (mac) {
    EXPECT_DOUBLE_EQ(fmt.attribute("baseColWidth").as_double(), 10.0);
    EXPECT_DOUBLE_EQ(fmt.attribute("defaultRowHeight").as_double(), c.row_height);
  } else {
    EXPECT_FALSE(fmt.attribute("baseColWidth"));
    EXPECT_FALSE(fmt.attribute("defaultRowHeight"));
  }

  const pugi::xml_node phonetic = ws.child("phoneticPr");
  if (c.phonetic_size != 0) {
    ASSERT_TRUE(phonetic);
    EXPECT_STREQ(phonetic.attribute("fontId").value(), "1");
    if (c.phonetic_type != nullptr) {
      EXPECT_STREQ(phonetic.attribute("type").value(), c.phonetic_type);
    } else {
      EXPECT_FALSE(phonetic.attribute("type"));
    }
    // Schema order: after sheetData, before pageMargins.
    EXPECT_STREQ(phonetic.previous_sibling().name(), "sheetData");
    EXPECT_STREQ(phonetic.next_sibling().name(), "pageMargins");
  } else {
    EXPECT_FALSE(phonetic);
  }

  const pugi::xml_node margins = ws.child("pageMargins");
  ASSERT_TRUE(margins);
  EXPECT_DOUBLE_EQ(margins.attribute("left").as_double(), c.margin_lr);
  EXPECT_DOUBLE_EQ(margins.attribute("right").as_double(), c.margin_lr);
  EXPECT_DOUBLE_EQ(margins.attribute("top").as_double(), c.margin_tb);
  EXPECT_DOUBLE_EQ(margins.attribute("bottom").as_double(), c.margin_tb);
  EXPECT_DOUBLE_EQ(margins.attribute("header").as_double(), c.margin_hf);
  EXPECT_DOUBLE_EQ(margins.attribute("footer").as_double(), c.margin_hf);
}

void ExpectWorkbookXmlSheetName(const pugi::xml_document& doc, const Captured& c) {
  EXPECT_STREQ(doc.child("workbook").child("sheets").child("sheet").attribute("name").value(), c.sheet_name);
}

/// Checks the model a reload of the saved package produces.
void ExpectReloaded(const Workbook& wb, const Captured& c) {
  const StylesTable& styles = wb.styles();
  ASSERT_FALSE(styles.fonts.empty());
  EXPECT_EQ(styles.fonts[0].scheme, 2U);
  EXPECT_EQ(styles.fonts[0].name, c.font_name);
  if (c.phonetic_size != 0) {
    ASSERT_EQ(styles.fonts.size(), 2U);
    EXPECT_EQ(styles.fonts[1].name, c.font_name);
    EXPECT_DOUBLE_EQ(styles.fonts[1].size, c.phonetic_size);
    EXPECT_EQ(styles.fonts[1].charset, c.charset);
    EXPECT_EQ(styles.fonts[1].scheme, 2U);
  } else {
    EXPECT_EQ(styles.fonts.size(), 1U);
  }
  const std::uint8_t vertical = c.phonetic_size != 0 ? 1U : 2U;  // center : bottom
  ASSERT_FALSE(styles.cell_style_xfs.empty());
  ASSERT_FALSE(styles.cell_xfs.empty());
  EXPECT_EQ(styles.cell_style_xfs[0].vertical_align, vertical);
  EXPECT_EQ(styles.cell_xfs[0].vertical_align, vertical);
  ASSERT_FALSE(styles.cell_styles.empty());
  EXPECT_EQ(styles.cell_styles[0].name, c.normal_name);

  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).name(), c.sheet_name);
  bool found_phonetic = false;
  for (const WorksheetRawChild& child : wb.sheet(0).raw_extensions()) {
    if (child.slot == worksheet_child::slot_of("phoneticPr")) {
      found_phonetic = true;
      EXPECT_NE(child.xml.find("fontId=\"1\""), std::string::npos) << child.xml;
    }
  }
  EXPECT_EQ(found_phonetic, c.phonetic_size != 0);
}

void CheckProfile(const Captured& c, const char* profile_id, bool mac) {
  SCOPED_TRACE(profile_id);
  const ExcelProfile profile = ProfileOf(profile_id);
  const Workbook wb = Workbook::create_for_profile(profile);
  EXPECT_TRUE(same_profile(wb.excel_profile(), profile));

  const std::vector<std::uint8_t> pkg = Save(wb);
  ASSERT_FALSE(pkg.empty());
  pugi::xml_document styles;
  pugi::xml_document sheet;
  pugi::xml_document book;
  LoadPart(pkg, "xl/styles.xml", &styles);
  LoadPart(pkg, "xl/worksheets/sheet1.xml", &sheet);
  LoadPart(pkg, "xl/workbook.xml", &book);
  ExpectStylesXml(styles, c, mac ? 12.0 : 11.0);
  ExpectSheetXml(sheet, c, mac);
  ExpectWorkbookXmlSheetName(book, c);

  auto reloaded = io::read_ooxml(io::ByteSpan{pkg.data(), pkg.size()});
  ASSERT_TRUE(static_cast<bool>(reloaded));
  ExpectReloaded(reloaded.value().workbook, c);
}

TEST(NewWorkbookDefaults, MacProfilesMatchExcelCaptures) {
  for (const Captured& c : kMac) {
    CheckProfile(c, c.profile_id, /*mac=*/true);
  }
}

TEST(NewWorkbookDefaults, WindowsEstimateUsesBodySize11AndEngineRowMetrics) {
  CheckProfile(kMac[10], "win-365-ja_JP", /*mac=*/false);
  CheckProfile(kMac[0], "win-365-en_US", /*mac=*/false);
}

TEST(NewWorkbookDefaults, CreateIsUnchanged) {
  const Workbook wb = Workbook::create();
  ASSERT_EQ(wb.styles().fonts.size(), 1U);
  EXPECT_EQ(wb.styles().fonts[0].name, "Calibri");
  EXPECT_DOUBLE_EQ(wb.styles().fonts[0].size, 11.0);
  EXPECT_EQ(wb.styles().fonts[0].scheme, 0U);
  EXPECT_EQ(wb.styles().cell_styles[0].name, "Normal");
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_TRUE(wb.sheet(0).raw_extensions().empty());
  EXPECT_TRUE(wb.sheet(0).print_settings().page_margins_xml.empty());
  EXPECT_TRUE(same_profile(wb.excel_profile(), default_excel_profile()));
}

}  // namespace
}  // namespace formulon
