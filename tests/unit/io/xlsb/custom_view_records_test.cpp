// Custom views encoded for XLSB: each attribute and child of a
// `<customWorkbookView>` / `<customSheetView>` against the record bytes Excel
// writes for it (Mac Excel 16.115, one attribute changed per saved variant
// of tests/fixtures/excel/retained_children_edit/children_0_base.xlsx), and
// the views whose records are not known, which are refused whole.

#include "io/xlsb/custom_view_records.h"

#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr const char* kGuid = "{8F1B2C3D-0000-4000-8000-000000000001}";
constexpr const char* kGuidHex = "3d2c1b8f000000408000000000000001";

std::string Hex(const std::uint8_t* data, std::size_t size) {
  static const char* const kDigits = "0123456789abcdef";
  std::string out;
  for (std::size_t i = 0; i < size; ++i) {
    out.push_back(kDigits[data[i] >> 4U]);
    out.push_back(kDigits[data[i] & 0x0FU]);
  }
  return out;
}

using Records = std::vector<std::pair<std::uint16_t, std::string>>;

/// The records `encode` writes for `xml`, each as its type and payload hex; `ok` reports the encoder's result.
Records Encode(bool (*encode)(const pugi::xml_node&, std::vector<std::uint8_t>&), const std::string& xml, bool* ok) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml.c_str())) << xml;
  std::vector<std::uint8_t> bytes;
  *ok = encode(doc.document_element(), bytes);
  Records out;
  for (const FramedRecord& rec : split_records(bytes)) {
    out.emplace_back(rec.type, Hex(rec.payload.data, rec.payload.size));
  }
  return out;
}

std::string BookView(const std::string& attrs) {
  return std::string("<customWorkbookView name=\"v1\" guid=\"") + kGuid +
         "\" windowWidth=\"1000\" windowHeight=\"800\" activeSheetId=\"1\"" + attrs + "/>";
}

/// The BrtUserBookView payload of `BookView(attrs)`.
std::string BookPayload(const std::string& attrs) {
  bool ok = false;
  const Records records = Encode(encode_book_view, BookView(attrs), &ok);
  EXPECT_TRUE(ok) << attrs;
  return records.size() == 1U && records[0].first == kBrtUserBookView ? records[0].second : std::string();
}

std::string U16At(const std::string& hex, std::size_t offset) {
  return hex.substr(offset * 2U, 4U);
}
std::string U32At(const std::string& hex, std::size_t offset) {
  return hex.substr(offset * 2U, 8U);
}

TEST(CustomViewRecords, BookViewBaseIsExcelsPayload) {
  EXPECT_EQ(BookPayload(""), std::string("00000000e803000000000000200300000100000058020000") + kGuidHex +
                                 "ffffee180000020000007600" + "3100");
}

TEST(CustomViewRecords, BookViewFields) {
  const std::string x = BookPayload(" xWindow=\"100\"");
  EXPECT_EQ(U32At(x, 0), "64000000");
  EXPECT_EQ(U32At(x, 4), "4c040000");
  const std::string y = BookPayload(" yWindow=\"200\"");
  EXPECT_EQ(U32At(y, 8), "c8000000");
  EXPECT_EQ(U32At(y, 12), "e8030000");
  EXPECT_EQ(U32At(BookPayload(" tabRatio=\"300\""), 20), "2c010000");
  EXPECT_EQ(U16At(BookPayload(" mergeInterval=\"7\""), 40), "0700");
  EXPECT_EQ(U16At(BookPayload(" personalView=\"1\""), 44), "0100");
}

TEST(CustomViewRecords, BookViewFlags) {
  const std::vector<std::pair<std::string, std::string>> cases = {
      {" minimized=\"1\"", "ef18"},
      {" showHorizontalScroll=\"0\"", "ec18"},
      {" showVerticalScroll=\"0\"", "ea18"},
      {" showSheetTabs=\"0\"", "e618"},
      {" maximized=\"1\"", "fe18"},
      {" showFormulaBar=\"0\"", "ce18"},
      {" showStatusbar=\"0\"", "ae18"},
      {" showComments=\"commNone\"", "6e18"},
      {" showComments=\"commIndAndComment\"", "6e19"},
      {" showObjects=\"placeholders\"", "ee1a"},
      {" showObjects=\"none\"", "ee1c"},
      {" includeHiddenRowCol=\"0\"", "ee08"},
      {" autoUpdate=\"1\"", "ee38"},
      {" changesSavedWin=\"1\"", "ee58"},
      {" onlySync=\"1\"", "ee98"},
  };
  for (const auto& [attrs, flags] : cases) {
    EXPECT_EQ(U16At(BookPayload(attrs), 42), flags) << attrs;
  }
}

TEST(CustomViewRecords, BookViewsWithoutAKnownRecordAreRefused) {
  for (const std::string attrs :
       {" includePrintSettings=\"0\"", " xWindow=\"-5\"", " tabRatio=\"1001\"", " mergeInterval=\"65535\"",
        " showComments=\"other\"", " minimized=\"yes\"", " xr:uid=\"{00000000-0000-0000-0000-000000000000}\""}) {
    bool ok = true;
    EXPECT_TRUE(Encode(encode_book_view, BookView(attrs), &ok).empty()) << attrs;
    EXPECT_FALSE(ok) << attrs;
  }
  bool ok = true;
  EXPECT_TRUE(Encode(encode_book_view,
                     "<customWorkbookView name=\"v1\" guid=\"x\" windowWidth=\"1\" "
                     "windowHeight=\"1\" activeSheetId=\"1\"/>",
                     &ok)
                  .empty());
  EXPECT_FALSE(ok);
}

constexpr const char* kBaseSelection = "<selection activeCell=\"C5\" sqref=\"C5:D6\"/>";
constexpr const char* kBaseMargins =
    "<pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/>";
constexpr const char* kSel = "030000000400000002000000000000000100000004000000050000000200000003000000";
constexpr const char* kMargins =
    "666666666666e63f666666666666e63f000000000000e83f000000000000e83f333333333333d33f333333333333d33f";

std::string SheetView(const std::string& attrs, const std::string& children) {
  return std::string("<customSheetView guid=\"") + kGuid + "\"" + attrs + ">" + children + "</customSheetView>";
}

Records SheetRecords(const std::string& attrs, const std::string& children) {
  bool ok = false;
  Records records = Encode(encode_sheet_view, SheetView(attrs, children), &ok);
  EXPECT_TRUE(ok) << attrs << children;
  return records;
}

/// The BrtBeginUserShView payload of a view with the base children.
std::string ViewPayload(const std::string& attrs) {
  const Records records = SheetRecords(attrs, std::string(kBaseSelection) + kBaseMargins);
  return records.empty() ? std::string() : records.front().second;
}

std::string ViewHead(const std::string& flags, const std::string& top_left) {
  return std::string(kGuidHex) + "01000000" + "64000000" + "40000000" + flags + top_left;
}

constexpr const char* kTopLeftB3 = "02000000020000000100000001000000";

TEST(CustomViewRecords, SheetViewBaseIsExcelsBlock) {
  const Records expected = {{kBrtBeginUserShView, ViewHead("3c000021", kTopLeftB3)},
                            {152, kSel},
                            {476, kMargins},
                            {477, "1000"},
                            {kBrtEndUserShView, ""}};
  EXPECT_EQ(SheetRecords(" topLeftCell=\"B3\"", std::string(kBaseSelection) + kBaseMargins), expected);
}

TEST(CustomViewRecords, SheetViewFlags) {
  const std::vector<std::pair<std::string, std::string>> cases = {
      {" showPageBreaks=\"1\"", "3d000021"}, {" showFormulas=\"1\"", "3e000021"},
      {" showGridLines=\"0\"", "38000021"},  {" showRowCol=\"0\"", "34000021"},
      {" outlineSymbols=\"0\"", "2c000021"}, {" zeroValues=\"0\"", "1c000021"},
      {" fitToPage=\"1\"", "3c040021"},      {" printArea=\"1\"", "3c080021"},
      {" hiddenRows=\"1\"", "3c000821"},     {" hiddenColumns=\"1\"", "3c002021"},
      {" state=\"hidden\"", "3c004021"},     {" state=\"veryHidden\"", "3c008021"},
      {" filterUnique=\"1\"", "3c000023"},   {" view=\"pageBreakPreview\"", "3c000025"},
      {" view=\"pageLayout\"", "3c000029"},  {" showRuler=\"0\"", "3c000001"},
  };
  for (const auto& [attrs, flags] : cases) {
    EXPECT_EQ(ViewPayload(attrs).substr(56U, 8U), flags) << attrs;
  }
  EXPECT_EQ(ViewPayload(" scale=\"85\"").substr(40U, 8U), "55000000");
  EXPECT_EQ(ViewPayload(" colorId=\"10\"").substr(48U, 8U), "0a000000");
  EXPECT_EQ(ViewPayload(" topLeftCell=\"AA100\"").substr(64U), "63000000630000001a0000001a000000");
  EXPECT_EQ(ViewPayload("").substr(64U), "00000000000000000000000000000000");
}

TEST(CustomViewRecords, SheetViewPrintOptionsGoToTheFlagsAndTheirRecord) {
  const std::vector<std::pair<std::string, std::pair<std::string, std::string>>> cases = {
      {"horizontalCentered", {"7c000021", "1100"}},
      {"verticalCentered", {"bc000021", "1200"}},
      {"headings", {"3c010021", "1400"}},
      {"gridLines", {"3c020021", "1800"}},
  };
  for (const auto& [attr, expected] : cases) {
    const Records records =
        SheetRecords("", std::string(kBaseSelection) + kBaseMargins + "<printOptions " + attr + "=\"1\"/>");
    ASSERT_EQ(records.size(), 5U) << attr;
    EXPECT_EQ(records[0].second.substr(56U, 8U), expected.first) << attr;
    EXPECT_EQ(records[3], (std::pair<std::uint16_t, std::string>{477, expected.second})) << attr;
  }
}

TEST(CustomViewRecords, SheetViewChildren) {
  // No selection writes no BrtSel; no margins writes Excel's defaults.
  Records records = SheetRecords("", "");
  ASSERT_EQ(records.size(), 4U);
  EXPECT_EQ(records[1], (std::pair<std::uint16_t, std::string>{476, kMargins}));
  // Two ranges in one selection.
  records = SheetRecords("", std::string("<selection sqref=\"A1 C5:D6\"/>") + kBaseMargins);
  ASSERT_EQ(records.size(), 5U);
  EXPECT_EQ(records[1].first, 152U);
  EXPECT_EQ(records[1].second.substr(32U, 8U), "02000000");
  // Manual breaks lead the print records.
  records = SheetRecords("", std::string(kBaseMargins) +
                                 "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"5\" max=\"16383\" "
                                 "man=\"1\"/></rowBreaks>");
  ASSERT_EQ(records.size(), 7U);
  EXPECT_EQ(records[1], (std::pair<std::uint16_t, std::string>{392, "0100000001000000"}));
  EXPECT_EQ(records[2], (std::pair<std::uint16_t, std::string>{396, "0500000000000000ff3f00000100000000000000"}));
  EXPECT_EQ(records[3].first, 393U);
  // Header and footer follow the page setup.
  records =
      SheetRecords("", std::string(kBaseMargins) + "<headerFooter><oddHeader>&amp;CTop</oddHeader></headerFooter>");
  ASSERT_EQ(records.size(), 6U);
  EXPECT_EQ(records[3].first, 479U);
  EXPECT_EQ(records[4].first, 480U);
}

TEST(CustomViewRecords, SheetViewPanes) {
  Records records =
      SheetRecords("",
                   "<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"frozen\"/>"
                   "<selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5:D6\"/>" +
                       std::string(kBaseMargins));
  ASSERT_EQ(records.size(), 6U);
  EXPECT_EQ(records[0].second.substr(56U, 8U), "3c000621");
  EXPECT_EQ(records[1].first, 151U);
  EXPECT_EQ(records[1].second.substr(32U), "02000000010000000000000003");
  EXPECT_EQ(records[2].first, 152U);
  EXPECT_EQ(records[2].second.substr(0U, 8U), "00000000");
  records = SheetRecords("",
                         "<pane xSplit=\"1500\" ySplit=\"1200\" topLeftCell=\"C4\" activePane=\"bottomRight\"/>"
                         "<selection pane=\"bottomRight\" activeCell=\"D5\" sqref=\"D5\"/>" +
                             std::string(kBaseMargins));
  ASSERT_EQ(records.size(), 6U);
  EXPECT_EQ(records[1].second.substr(32U), "03000000020000000000000000");
}

TEST(CustomViewRecords, SheetViewsWithoutKnownRecordsAreRefused) {
  const std::vector<std::pair<std::string, std::string>> cases = {
      {" filter=\"1\"", ""},
      {" showAutoFilter=\"1\"", ""},
      {"", "<autoFilter ref=\"A1:E13\"/>"},
      {"", "<extLst/>"},
      {" scale=\"5\"", ""},
      {" colorId=\"65\"", ""},
      {" state=\"gone\"", ""},
      {" topLeftCell=\"B\"", ""},
      {" unknown=\"1\"", ""},
      {"", "<pageSetup r:id=\"rId1\"/>"},
      {"", "<pageSetup paperHeight=\"297mm\"/>"},
      {"", "<printOptions gridLinesSet=\"0\"/>"},
      {"", "<pageMargins left=\"0.7\"/>"},
      {"", "<rowBreaks><brk id=\"5\" max=\"16383\"/></rowBreaks>"},
      {"", "<pane xSplit=\"1\" topLeftCell=\"B1\" activePane=\"topRight\" state=\"frozen\"/>"},
      {"", "<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"frozenSplit\"/>"},
      {"",
       "<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" state=\"frozen\"/>"
       "<selection pane=\"topRight\" activeCell=\"B1\" sqref=\"B1\"/>"
       "<selection pane=\"bottomLeft\" activeCell=\"A3\" sqref=\"A3\"/>"
       "<selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5\"/>"
       "<selection pane=\"bottomRight\" activeCell=\"C6\" sqref=\"C6\"/>"},
      {"", "<selection activeCell=\"C5\" sqref=\"C5\"/><selection activeCell=\"C6\" sqref=\"C6\"/>"},
      {"", "<selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5\"/>"},
  };
  for (const auto& [attrs, children] : cases) {
    bool ok = true;
    EXPECT_TRUE(Encode(encode_sheet_view, SheetView(attrs, children), &ok).empty()) << attrs << children;
    EXPECT_FALSE(ok) << attrs << children;
  }
}

/// `xml` with each element's attributes sorted, so documents that differ only in attribute order compare equal.
std::string Normalized(const pugi::xml_node& node) {
  std::map<std::string, std::string> attrs;
  for (const pugi::xml_attribute& attr : node.attributes()) {
    attrs[attr.name()] = attr.value();
  }
  std::string out = std::string("<") + node.name();
  for (const auto& [name, value] : attrs) {
    out += " " + name + "=" + value;
  }
  out += ">";
  for (const pugi::xml_node& child : node.children()) {
    out += Normalized(child);
  }
  return out + "</>";
}

std::string Normalized(const std::string& xml) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml.c_str())) << xml;
  return Normalized(doc.document_element());
}

/// `decode(encode(xml))` for a sheet view.
std::optional<std::string> SheetRoundTrip(const std::string& xml) {
  pugi::xml_document doc;
  EXPECT_TRUE(doc.load_string(xml.c_str())) << xml;
  std::vector<std::uint8_t> bytes;
  EXPECT_TRUE(encode_sheet_view(doc.document_element(), bytes)) << xml;
  return decode_sheet_view(split_records(bytes), bytes);
}

TEST(CustomViewRecords, BookViewsDecodeBackToTheirXml) {
  for (const std::string attrs : {"",
                                  " xWindow=\"100\"",
                                  " yWindow=\"200\"",
                                  " tabRatio=\"300\"",
                                  " mergeInterval=\"7\"",
                                  " personalView=\"1\"",
                                  " minimized=\"1\"",
                                  " showHorizontalScroll=\"0\"",
                                  " showVerticalScroll=\"0\"",
                                  " showSheetTabs=\"0\"",
                                  " maximized=\"1\"",
                                  " showFormulaBar=\"0\"",
                                  " showStatusbar=\"0\"",
                                  " showComments=\"commNone\"",
                                  " showComments=\"commIndAndComment\"",
                                  " showObjects=\"placeholders\"",
                                  " showObjects=\"none\"",
                                  " includeHiddenRowCol=\"0\"",
                                  " autoUpdate=\"1\"",
                                  " changesSavedWin=\"1\"",
                                  " onlySync=\"1\""}) {
    pugi::xml_document doc;
    ASSERT_TRUE(doc.load_string(BookView(attrs).c_str()));
    std::vector<std::uint8_t> bytes;
    ASSERT_TRUE(encode_book_view(doc.document_element(), bytes)) << attrs;
    const std::vector<FramedRecord> records = split_records(bytes);
    ASSERT_EQ(records.size(), 1U);
    const std::optional<std::string> xml = decode_book_view(records[0].payload);
    ASSERT_TRUE(xml.has_value()) << attrs;
    EXPECT_EQ(Normalized(*xml), Normalized(BookView(attrs))) << attrs;
  }
}

TEST(CustomViewRecords, BookViewDecodesInExcelsAttributeOrder) {
  const std::string flipped = BookView(
      " autoUpdate=\"1\" changesSavedWin=\"1\" onlySync=\"1\" personalView=\"1\" includeHiddenRowCol=\"0\" "
      "maximized=\"1\" showHorizontalScroll=\"0\" showVerticalScroll=\"0\" showSheetTabs=\"0\" "
      "showFormulaBar=\"0\" showStatusbar=\"0\"");
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(flipped.c_str()));
  std::vector<std::uint8_t> bytes;
  ASSERT_TRUE(encode_book_view(doc.document_element(), bytes));
  EXPECT_EQ(decode_book_view(split_records(bytes)[0].payload),
            std::string("<customWorkbookView name=\"v1\" guid=\"") + kGuid +
                "\" autoUpdate=\"1\" changesSavedWin=\"1\" onlySync=\"1\" personalView=\"1\" "
                "includeHiddenRowCol=\"0\" maximized=\"1\" showHorizontalScroll=\"0\" showVerticalScroll=\"0\" "
                "showSheetTabs=\"0\" windowWidth=\"1000\" windowHeight=\"800\" activeSheetId=\"1\" "
                "showFormulaBar=\"0\" showStatusbar=\"0\"/>");
}

TEST(CustomViewRecords, SheetViewsDecodeBackToTheirXml) {
  const std::string base = std::string(kBaseSelection) + kBaseMargins;
  const std::vector<std::pair<std::string, std::string>> cases = {
      {" topLeftCell=\"B3\"", base},
      {" showPageBreaks=\"1\" showFormulas=\"1\" showGridLines=\"0\" showRowCol=\"0\" outlineSymbols=\"0\" "
       "zeroValues=\"0\" fitToPage=\"1\" printArea=\"1\" hiddenRows=\"1\" hiddenColumns=\"1\" "
       "filterUnique=\"1\" showRuler=\"0\"",
       base},
      {" scale=\"85\" colorId=\"10\" state=\"veryHidden\" view=\"pageLayout\" topLeftCell=\"AA100\"", base},
      {" state=\"hidden\" view=\"pageBreakPreview\"", base},
      {"", std::string(kBaseMargins) + "<printOptions horizontalCentered=\"1\" headings=\"1\"/>"},
      {"", std::string(kBaseMargins) + "<pageSetup orientation=\"landscape\" scale=\"80\"/>"},
      {"", std::string("<selection sqref=\"A1 C5:D6\"/>") + kBaseMargins},
      {"", std::string("<pane xSplit=\"1\" ySplit=\"2\" topLeftCell=\"B3\" activePane=\"bottomRight\" "
                       "state=\"frozen\"/><selection pane=\"bottomRight\" activeCell=\"C5\" sqref=\"C5:D6\"/>") +
               kBaseMargins},
      {"", std::string("<pane xSplit=\"1500\" ySplit=\"1200\" topLeftCell=\"C4\" activePane=\"bottomRight\"/>"
                       "<selection pane=\"bottomRight\" activeCell=\"D5\" sqref=\"D5\"/>") +
               kBaseMargins},
  };
  for (const auto& [attrs, children] : cases) {
    const std::optional<std::string> xml = SheetRoundTrip(SheetView(attrs, children));
    ASSERT_TRUE(xml.has_value()) << attrs << children;
    EXPECT_EQ(Normalized(*xml), Normalized(SheetView(attrs, children))) << attrs << children;
  }
}

TEST(CustomViewRecords, SheetViewsWithoutAnXlsxFormAreNotDecoded) {
  // Page breaks and a header have records but no `.xlsx` form here.
  for (const std::string& children :
       {std::string(kBaseMargins) + "<rowBreaks count=\"1\" manualBreakCount=\"1\"><brk id=\"5\" max=\"16383\" "
                                    "man=\"1\"/></rowBreaks>",
        std::string(kBaseMargins) + "<headerFooter><oddHeader>&amp;CTop</oddHeader></headerFooter>"}) {
    EXPECT_FALSE(SheetRoundTrip(SheetView("", children)).has_value()) << children;
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
