//
// What a save writes from `Workbook::location()`: local hyperlink targets
// relative to the book's folder (with the derived display), and the
// regenerated `x15ac:absPath`. Rows follow Mac Excel 365 saves of a book in
// `.../Data/loc_probe3/p q/Book.xlsx`, then Save As to `.../loc_probe4/r/`.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace {

const std::string kData = "/Users/libraz/Library/Containers/com.microsoft.Excel/Data";
const std::string kProbe3 = kData + "/loc_probe3/p q";
const std::string kProbe4 = kData + "/loc_probe4/r";

struct LinkRow {
  std::string input;
  std::string display;
};

// One link per row of column A, in this order.
std::vector<LinkRow> Links() {
  return {
      {"/Users/libraz/Documents/f.xlsx", ""},
      {"/Users/libraz/Documents/a b/f.xlsx", ""},
      {"/Volumes/music/x.xlsx", ""},
      {"/private/tmp/x.xlsx", ""},
      {"/Users/libraz/Library/Containers/com.microsoft.Excel/data/loc_probe3/p q/other.xlsx", ""},
      {kProbe3 + "/sib.xlsx", ""},
      {"/Users/libraz/Documents/g.xlsx", "Mine"},
      {"\\\\server\\share\\f.xlsx", ""},
      {"https://e.x/a b", ""},
  };
}

Workbook BookWithLinks() {
  Workbook wb = Workbook::create();
  const std::vector<LinkRow> links = Links();
  for (std::size_t i = 0; i < links.size(); ++i) {
    Hyperlink link;
    link.row = static_cast<std::uint32_t>(i);
    link.last_row = link.row;
    link.target = links[i].input;
    link.display = links[i].display;
    EXPECT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link))) << links[i].input;
  }
  return wb;
}

std::vector<std::uint8_t> SaveXlsx(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value()) : std::vector<std::uint8_t>{};
}

// The hyperlinks a save leaves on the first sheet, read back as written.
std::vector<Hyperlink> SavedLinks(const Workbook& wb, bool xlsb) {
  std::vector<std::uint8_t> bytes;
  if (xlsb) {
    auto saved = xlsb::write_xlsb(wb);
    EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
    if (!saved) {
      return {};
    }
    auto loaded = xlsb::read_xlsb(test::span_of(saved.value()));
    EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
    return loaded ? loaded.value().workbook.sheet(0).hyperlinks() : std::vector<Hyperlink>{};
  }
  bytes = SaveXlsx(wb);
  auto loaded = read_ooxml(test::span_of(bytes));
  EXPECT_TRUE(static_cast<bool>(loaded)) << (loaded ? "" : loaded.error().message);
  return loaded ? loaded.value().workbook.sheet(0).hyperlinks() : std::vector<Hyperlink>{};
}

struct SavedRow {
  std::string target;
  std::string display;
};

void ExpectSaved(const Workbook& wb, const std::vector<SavedRow>& expected) {
  for (const bool xlsb : {false, true}) {
    const std::vector<Hyperlink> saved = SavedLinks(wb, xlsb);
    ASSERT_EQ(saved.size(), expected.size()) << (xlsb ? "xlsb" : "xlsx");
    for (std::size_t i = 0; i < expected.size(); ++i) {
      EXPECT_EQ(saved[i].target, expected[i].target) << (xlsb ? "xlsb row " : "xlsx row ") << i;
      EXPECT_EQ(saved[i].display, expected[i].display) << (xlsb ? "xlsb row " : "xlsx row ") << i;
    }
  }
}

TEST(BookLocationSave, ModelHoldsAbsoluteLocalTargets) {
  Workbook wb = BookWithLinks();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe3, "Book.xlsx"})));
  const std::vector<Hyperlink>& links = wb.sheet(0).hyperlinks();
  ASSERT_EQ(links.size(), 9U);
  EXPECT_EQ(links[0].target, "file:///Users/libraz/Documents/f.xlsx");
  EXPECT_EQ(links[1].target, "file:///Users/libraz/Documents/a b/f.xlsx");
  EXPECT_EQ(links[5].target, "file://" + kProbe3 + "/sib.xlsx");
  EXPECT_EQ(links[7].target, "file:////server/share/f.xlsx");
  // The writer derives the display; the model keeps none.
  EXPECT_EQ(links[0].display, "");
  EXPECT_EQ(links[6].display, "Mine");
}

TEST(BookLocationSave, WritesSharedPrefixTargetsRelativeWithDerivedDisplay) {
  Workbook wb = BookWithLinks();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe3, "Book.xlsx"})));
  ExpectSaved(wb, {
                      {"../../../../../../Documents/f.xlsx", "../../../../../../Documents/f.xlsx"},
                      {"../../../../../../Documents/a%20b/f.xlsx", "../../../../../../Documents/a b/f.xlsx"},
                      {"file:///Volumes/music/x.xlsx", ""},
                      {"file:///private/tmp/x.xlsx", ""},
                      {"../../../data/loc_probe3/p%20q/other.xlsx", "../../../data/loc_probe3/p q/other.xlsx"},
                      {"sib.xlsx", "sib.xlsx"},
                      {"../../../../../../Documents/g.xlsx", "Mine"},
                      {"file:////server/share/f.xlsx", ""},
                      {"https://e.x/a%20b", ""},
                  });
}

TEST(BookLocationSave, WritesAbsoluteTargetsWithoutAFolder) {
  Workbook wb = BookWithLinks();
  const std::vector<SavedRow> absolute = {
      {"file:///Users/libraz/Documents/f.xlsx", ""},
      {"file:///Users/libraz/Documents/a%20b/f.xlsx", ""},
      {"file:///Volumes/music/x.xlsx", ""},
      {"file:///private/tmp/x.xlsx", ""},
      {"file:///Users/libraz/Library/Containers/com.microsoft.Excel/data/loc_probe3/p%20q/other.xlsx", ""},
      {"file:///Users/libraz/Library/Containers/com.microsoft.Excel/Data/loc_probe3/p%20q/sib.xlsx", ""},
      {"file:///Users/libraz/Documents/g.xlsx", "Mine"},
      {"file:////server/share/f.xlsx", ""},
      {"https://e.x/a%20b", ""},
  };
  ExpectSaved(wb, absolute);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "Book.xlsx"})));
  ExpectSaved(wb, absolute);
}

TEST(BookLocationSave, ANewFolderChangesTheSavedRelativeForm) {
  Workbook wb = BookWithLinks();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe3, "Book.xlsx"})));
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe4, "Book.xlsx"})));
  const std::vector<Hyperlink> saved = SavedLinks(wb, /*xlsb=*/false);
  ASSERT_EQ(saved.size(), 9U);
  EXPECT_EQ(saved[5].target, "../../loc_probe3/p%20q/sib.xlsx");
  EXPECT_EQ(saved[5].display, "../../loc_probe3/p q/sib.xlsx");
  EXPECT_EQ(saved[0].target, "../../../../../../Documents/f.xlsx");
}

TEST(BookLocationSave, RelativeTargetAnchorsOnTheFirstFolder) {
  Workbook wb = Workbook::create();
  Hyperlink link;
  link.target = "docs/a.xlsx";
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link)));
  // Unset: written as entered, with no display.
  std::vector<Hyperlink> saved = SavedLinks(wb, /*xlsb=*/false);
  ASSERT_EQ(saved.size(), 1U);
  EXPECT_EQ(saved[0].target, "docs/a.xlsx");
  EXPECT_EQ(saved[0].display, "");

  const std::string depth = kData + "/hlmodel/addr_depth";
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{depth, "shallow.xlsx"})));
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].target, "file://" + depth + "/docs/a.xlsx");
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{depth + "/x/y/z", "deep.xlsx"})));
  saved = SavedLinks(wb, /*xlsb=*/false);
  ASSERT_EQ(saved.size(), 1U);
  EXPECT_EQ(saved[0].target, "../../../docs/a.xlsx");
  EXPECT_EQ(saved[0].display, "../../../docs/a.xlsx");
}

TEST(BookLocationSave, XlsbDisplayFollowsCallerOrRelativeText) {
  Workbook wb = Workbook::create();
  const char* const displays[] = {"", "Mine"};
  for (std::uint32_t row = 0; row < 2U; ++row) {
    Hyperlink link;
    link.row = row;
    link.last_row = row;
    link.target = "/a/b/c d/f.xlsx";
    link.display = displays[row];
    ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link)));
  }
  Hyperlink outside;
  outside.row = 2U;
  outside.last_row = 2U;
  outside.target = "/z/f.xlsx";
  ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, outside)));
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/a/x"), "Book.xlsb"})));
  const std::vector<Hyperlink> saved = SavedLinks(wb, /*xlsb=*/true);
  ASSERT_EQ(saved.size(), 3U);
  EXPECT_EQ(saved[0].target, "../b/c%20d/f.xlsx");
  EXPECT_EQ(saved[0].display, "../b/c d/f.xlsx");
  EXPECT_EQ(saved[1].target, "../b/c%20d/f.xlsx");
  EXPECT_EQ(saved[1].display, "Mine");
  EXPECT_EQ(saved[2].target, "file:///z/f.xlsx");
  EXPECT_EQ(saved[2].display, "");
  // The model keeps no derived display.
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].display, "");
}

// Spaces become %20 in the rel Target only; non-ASCII stays UTF-8 in both.
TEST(BookLocationSave, SpaceAndNonAsciiPathsSaveRelative) {
  const std::string dir = "/Users/u/Docs/\xE3\x83\x96\xE3\x83\x83\xE3\x82\xAF";  // ブック
  Workbook wb = Workbook::create();
  const std::string targets[] = {
      dir + "/sub dir/\xC3\xA9 f.xlsx",  // é f.xlsx
      "/Users/u/\xC3\x9C x/f.xlsx",      // Ü x
  };
  for (std::uint32_t row = 0; row < 2U; ++row) {
    Hyperlink link;
    link.row = row;
    link.last_row = row;
    link.target = targets[row];
    ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link)));
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{dir, "Book.xlsx"})));
  EXPECT_EQ(wb.sheet(0).hyperlinks()[0].target, "file://" + targets[0]);

  const std::vector<std::uint8_t> bytes = SaveXlsx(wb);
  std::string rels;
  ASSERT_TRUE(test::extract_part(test::span_of(bytes), "xl/worksheets/_rels/sheet1.xml.rels", &rels));
  EXPECT_NE(rels.find("Target=\"sub%20dir/\xC3\xA9%20f.xlsx\""), std::string::npos) << rels;
  EXPECT_NE(rels.find("Target=\"../../\xC3\x9C%20x/f.xlsx\""), std::string::npos) << rels;
  std::string sheet;
  ASSERT_TRUE(test::extract_part(test::span_of(bytes), "xl/worksheets/sheet1.xml", &sheet));
  EXPECT_NE(sheet.find("display=\"sub dir/\xC3\xA9 f.xlsx\""), std::string::npos) << sheet;
  EXPECT_NE(sheet.find("display=\"../../\xC3\x9C x/f.xlsx\""), std::string::npos) << sheet;

  for (const bool xlsb : {false, true}) {
    const std::vector<Hyperlink> saved = SavedLinks(wb, xlsb);
    ASSERT_EQ(saved.size(), 2U);
    EXPECT_EQ(saved[0].target, "sub%20dir/\xC3\xA9%20f.xlsx") << xlsb;
    EXPECT_EQ(saved[0].display, "sub dir/\xC3\xA9 f.xlsx") << xlsb;
    EXPECT_EQ(saved[1].target, "../../\xC3\x9C%20x/f.xlsx") << xlsb;
    EXPECT_EQ(saved[1].display, "../../\xC3\x9C x/f.xlsx") << xlsb;
  }
}

// `%XX` decoded, as a reader of the rel Target does.
std::string PercentDecoded(const std::string& text) {
  std::string out;
  for (std::size_t i = 0; i < text.size(); ++i) {
    if (text[i] == '%' && i + 2 < text.size()) {
      const auto hex = [](char c) { return c <= '9' ? c - '0' : (c | 0x20) - 'a' + 10; };
      out.push_back(static_cast<char>(hex(text[i + 1]) * 16 + hex(text[i + 2])));
      i += 2;
      continue;
    }
    out.push_back(text[i]);
  }
  return out;
}

// Links added with bare POSIX paths to a book in loc_probe6/f1 (Mac Excel
// 365); `%` follows the engine's single-escape rule.
TEST(BookLocationSave, FragmentPercentAndEscapingRows) {
  const std::string f1 = kData + "/loc_probe6/f1";
  const std::string japanese = "\xE6\x97\xA5\xE6\x9C\xAC";  // 日本
  Workbook wb = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{f1, "B.xlsx"})));
  const std::string inputs[] = {
      f1 + "/a%b.xlsx",   f1 + "/a#b.xlsx",   f1 + "/" + japanese + ".xlsx",
      f1 + "/a+b&c.xlsx", f1 + "/a%20b.xlsx", "/private/tmp/" + japanese + " x.xlsx",
  };
  for (std::uint32_t row = 0; row < 6U; ++row) {
    Hyperlink link;
    link.row = row;
    link.last_row = row;
    link.col = 1U;
    link.last_col = 1U;
    link.target = inputs[row];
    ASSERT_TRUE(static_cast<bool>(wb.set_hyperlink(0, link))) << inputs[row];
  }

  // The model and the cell text.
  const std::vector<Hyperlink>& model = wb.sheet(0).hyperlinks();
  ASSERT_EQ(model.size(), 6U);
  const std::string model_targets[] = {
      f1 + "/a%b.xlsx",
      "file://" + f1 + "/a",
      "file://" + f1 + "/" + japanese + ".xlsx",
      "file://" + f1 + "/a+b&c.xlsx",
      f1 + "/a b.xlsx",
      "file:///private/tmp/" + japanese + " x.xlsx",
  };
  for (std::uint32_t row = 0; row < 6U; ++row) {
    EXPECT_EQ(model[row].target, model_targets[row]) << row;
  }
  EXPECT_EQ(model[1].location, "b.xlsx");
  const std::string cell_texts[] = {
      model_targets[0], "file://" + f1 + "/a - b.xlsx", model_targets[2], model_targets[3], model_targets[4],
      model_targets[5],
  };
  for (std::uint32_t row = 0; row < 6U; ++row) {
    const Value shown = wb.sheet(0).resolve_cell_value(row, 1U);
    ASSERT_TRUE(shown.is_text()) << row;
    EXPECT_EQ(std::string(shown.as_text()), cell_texts[row]) << row;
  }

  // The saved rels and attributes, escaped as Excel escapes them.
  const std::vector<std::uint8_t> bytes = SaveXlsx(wb);
  std::string rels;
  ASSERT_TRUE(test::extract_part(test::span_of(bytes), "xl/worksheets/_rels/sheet1.xml.rels", &rels));
  for (const std::string& target :
       {std::string("a%25b.xlsx"), std::string("a"), japanese + ".xlsx", std::string("a+b&amp;c.xlsx"),
        std::string("a%20b.xlsx"), "file:///private/tmp/" + japanese + "%20x.xlsx"}) {
    EXPECT_NE(rels.find("Target=\"" + target + "\""), std::string::npos) << target << "\n" << rels;
  }
  std::string sheet;
  ASSERT_TRUE(test::extract_part(test::span_of(bytes), "xl/worksheets/sheet1.xml", &sheet));
  EXPECT_NE(sheet.find("display=\"a+b&amp;c.xlsx\""), std::string::npos) << sheet;

  const SavedRow expected[] = {
      {"a%25b.xlsx", "a%25b.xlsx"}, {"a", "a - b.xlsx"},        {japanese + ".xlsx", japanese + ".xlsx"},
      {"a+b&c.xlsx", "a+b&c.xlsx"}, {"a%20b.xlsx", "a b.xlsx"}, {"file:///private/tmp/" + japanese + "%20x.xlsx", ""},
  };
  for (const bool xlsb : {false, true}) {
    const std::vector<Hyperlink> saved = SavedLinks(wb, xlsb);
    ASSERT_EQ(saved.size(), 6U) << xlsb;
    for (std::size_t i = 0; i < 6U; ++i) {
      EXPECT_EQ(saved[i].target, expected[i].target) << (xlsb ? "xlsb row " : "xlsx row ") << i;
      EXPECT_EQ(saved[i].display, expected[i].display) << (xlsb ? "xlsb row " : "xlsx row ") << i;
    }
    EXPECT_EQ(saved[1].location, "b.xlsx") << xlsb;
    // One escape level: decoding the saved Target gives the file name back.
    EXPECT_EQ(PercentDecoded(saved[0].target), "a%b.xlsx") << xlsb;
    EXPECT_EQ(PercentDecoded(saved[4].target), "a b.xlsx") << xlsb;
  }
}

std::string SavedWorkbookXml(const Workbook& wb) {
  const std::vector<std::uint8_t> bytes = SaveXlsx(wb);
  std::string xml;
  EXPECT_TRUE(test::extract_part(test::span_of(bytes), "xl/workbook.xml", &xml));
  return xml;
}

std::size_t Count(const std::string& text, const std::string& needle) {
  std::size_t count = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + 1)) {
    ++count;
  }
  return count;
}

TEST(BookLocationSave, AbsPathFollowsTheLocation) {
  Workbook wb = Workbook::create();
  EXPECT_EQ(SavedWorkbookXml(wb).find("absPath"), std::string::npos);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "Book.xlsx"})));
  EXPECT_EQ(SavedWorkbookXml(wb).find("absPath"), std::string::npos);

  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe3, "Book.xlsx"})));
  const std::string xml = SavedWorkbookXml(wb);
  EXPECT_NE(xml.find("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" "
                     "xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"><mc:Choice "
                     "Requires=\"x15\"><x15ac:absPath url=\"" +
                     kProbe3 +
                     "/\" xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/>"
                     "</mc:Choice></mc:AlternateContent>"),
            std::string::npos)
      << xml;

  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/"), "Book.xlsx"})));
  EXPECT_NE(SavedWorkbookXml(wb).find("<x15ac:absPath url=\"/\""), std::string::npos);
  ASSERT_TRUE(static_cast<bool>(wb.set_location(std::nullopt)));
  EXPECT_EQ(SavedWorkbookXml(wb).find("absPath"), std::string::npos);
}

TEST(BookLocationSave, LoadedAbsPathIsRegeneratedNotPassedThrough) {
  const std::vector<std::uint8_t> source =
      test::read_file_bytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/new_workbook_mac_en_us.xlsx");
  ASSERT_FALSE(source.empty());
  std::string source_xml;
  ASSERT_TRUE(test::extract_part(test::span_of(source), "xl/workbook.xml", &source_xml));
  ASSERT_NE(source_xml.find("absPath"), std::string::npos);

  auto loaded = read_ooxml(test::span_of(source));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  Workbook& wb = loaded.value().workbook;
  // A load never sets the location from absPath.
  EXPECT_FALSE(wb.location().has_value());
  EXPECT_EQ(SavedWorkbookXml(wb).find("absPath"), std::string::npos);

  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe4, "Book.xlsx"})));
  const std::string xml = SavedWorkbookXml(wb);
  EXPECT_EQ(Count(xml, "absPath"), 1U) << xml;
  // The root declares x15, so the element is Excel's own byte form.
  EXPECT_NE(xml.find("<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
                     "<mc:Choice Requires=\"x15\"><x15ac:absPath url=\"" +
                     kProbe4 + "/\""),
            std::string::npos)
      << xml;
  // Excel's position: after <workbookPr>, before <bookViews>.
  const std::size_t abs_path = xml.find("absPath");
  EXPECT_LT(xml.find("<workbookPr"), abs_path);
  EXPECT_GT(xml.find("<bookViews"), abs_path);
}

}  // namespace
}  // namespace io
}  // namespace formulon
