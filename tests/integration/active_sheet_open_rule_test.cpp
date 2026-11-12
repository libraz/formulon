// The active sheet a load settles on, measured against Mac Excel 16.115:
// `activeTab` (the first sheet without `<bookViews>`) decides and
// `tabSelected` is never consulted; a hidden active sheet is unhidden, a very
// hidden one passes activation to the first visible sheet, and the active
// sheet ends up as the selected tab.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "miniz.h"
#include "sheet.h"
#include "sheet_view.h"
#include "workbook.h"

namespace formulon {
namespace {

struct PartFile {
  std::string path;
  std::string body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path.c_str(), p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
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

/// One sheet of the package: its `state` attribute (empty for visible) and
/// whether its view is `tabSelected`.
struct SheetSpec {
  std::string_view state;
  bool tab_selected;
};

/// A package of `sheets.size()` sheets, every one carrying `<sheetViews>`;
/// `book_views` lands before `<sheets>` and is left out when empty.
std::vector<std::uint8_t> BuildPackage(const std::vector<SheetSpec>& sheets, std::string_view book_views) {
  std::string overrides;
  std::string sheet_list;
  std::string rels;
  std::vector<PartFile> parts;
  for (std::size_t i = 0; i < sheets.size(); ++i) {
    const std::string n = std::to_string(i + 1U);
    overrides += "<Override PartName=\"/xl/worksheets/sheet" + n +
                 ".xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>";
    sheet_list += "<sheet name=\"Sheet" + n + "\" sheetId=\"" + n + "\"" +
                  (sheets[i].state.empty() ? std::string() : " state=\"" + std::string(sheets[i].state) + "\"") +
                  " r:id=\"rId" + n + "\"/>";
    rels += "<Relationship Id=\"rId" + n +
            "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
            "Target=\"worksheets/sheet" +
            n + ".xml\"/>";
    parts.push_back({"xl/worksheets/sheet" + n + ".xml",
                     std::string(kXmlDecl) +
                         "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                         "<sheetViews><sheetView" +
                         (sheets[i].tab_selected ? " tabSelected=\"1\"" : "") +
                         " workbookViewId=\"0\"/></sheetViews><sheetData/></worksheet>"});
  }
  parts.push_back({"[Content_Types].xml",
                   std::string(kXmlDecl) +
                       "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                       "<Default Extension=\"rels\" "
                       "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                       "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                       "<Override PartName=\"/xl/workbook.xml\" "
                       "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>" +
                       overrides + "</Types>"});
  parts.push_back({"_rels/.rels", std::string(kXmlDecl) +
                                      "<Relationships "
                                      "xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                      "<Relationship Id=\"rId1\" "
                                      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                                      "officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"});
  parts.push_back(
      {"xl/workbook.xml", std::string(kXmlDecl) +
                              "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
                              "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">" +
                              std::string(book_views) + "<sheets>" + sheet_list + "</sheets></workbook>"});
  parts.push_back({"xl/_rels/workbook.xml.rels",
                   std::string(kXmlDecl) +
                       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" + rels +
                       "</Relationships>"});
  return BuildZip(parts);
}

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto load_or = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  return load_or ? std::move(load_or.value().workbook) : Workbook::create_empty();
}

/// Reloads `wb` through an xlsx save.
Workbook ThroughXlsx(const Workbook& wb) {
  auto saved = wb.save();
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? Load(saved.value()) : Workbook::create_empty();
}

/// Reloads `wb` through an xlsb save.
Workbook ThroughXlsb(const Workbook& wb) {
  auto saved = io::xlsb::write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return Workbook::create_empty();
  }
  const std::vector<std::uint8_t>& bytes = saved.value().bytes;
  auto load_or = io::xlsb::read_xlsb(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(load_or)) << (load_or ? "" : load_or.error().message);
  return load_or ? std::move(load_or.value().workbook) : Workbook::create_empty();
}

SheetVisibility Visibility(const Workbook& wb, std::size_t sheet) {
  return wb.sheet(sheet).view().visibility();
}

/// Indices of the sheets whose view is `tabSelected`.
std::vector<std::size_t> Selected(const Workbook& wb) {
  std::vector<std::size_t> out;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).view().tab_selected) {
      out.push_back(i);
    }
  }
  return out;
}

constexpr std::string_view kActiveTab2 = "<bookViews><workbookView activeTab=\"2\"/></bookViews>";

TEST(ActiveSheetOpenRule, WithoutBookViewsTheFirstSheetIsActiveWhateverIsSelected) {
  const Workbook wb = Load(BuildPackage({{"", false}, {"", false}, {"", true}}, {}));
  EXPECT_EQ(wb.active_sheet(), 0U);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{0U}));
}

TEST(ActiveSheetOpenRule, AHiddenFirstSheetWithoutBookViewsIsUnhiddenAndActive) {
  const Workbook wb = Load(BuildPackage({{"hidden", false}, {"", false}, {"", true}}, {}));
  EXPECT_EQ(wb.active_sheet(), 0U);
  EXPECT_EQ(Visibility(wb, 0), SheetVisibility::kVisible);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{0U}));
}

TEST(ActiveSheetOpenRule, OnlyTheActiveHiddenSheetIsUnhidden) {
  const Workbook wb = Load(BuildPackage({{"hidden", false}, {"hidden", false}, {"", false}}, {}));
  EXPECT_EQ(wb.active_sheet(), 0U);
  EXPECT_EQ(Visibility(wb, 0), SheetVisibility::kVisible);
  EXPECT_EQ(Visibility(wb, 1), SheetVisibility::kHidden);
}

TEST(ActiveSheetOpenRule, AVeryHiddenFirstSheetPassesActivationToTheFirstVisibleOne) {
  const Workbook wb = Load(BuildPackage({{"veryHidden", false}, {"", false}, {"", true}}, {}));
  EXPECT_EQ(wb.active_sheet(), 1U);
  EXPECT_EQ(Visibility(wb, 0), SheetVisibility::kVeryHidden);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{1U}));
}

TEST(ActiveSheetOpenRule, AHiddenActiveTabIsUnhidden) {
  const Workbook wb = Load(BuildPackage({{"", false}, {"", false}, {"hidden", true}}, kActiveTab2));
  EXPECT_EQ(wb.active_sheet(), 2U);
  EXPECT_EQ(Visibility(wb, 2), SheetVisibility::kVisible);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{2U}));
}

TEST(ActiveSheetOpenRule, AVeryHiddenActiveTabFallsToTheFirstVisibleSheetNotTheNearest) {
  const Workbook wb = Load(BuildPackage({{"", false}, {"", false}, {"veryHidden", true}}, kActiveTab2));
  EXPECT_EQ(wb.active_sheet(), 0U);
  EXPECT_EQ(Visibility(wb, 2), SheetVisibility::kVeryHidden);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{0U}));
}

TEST(ActiveSheetOpenRule, ASelectedActiveSheetKeepsItsTabGroup) {
  const Workbook wb = Load(BuildPackage({{"", false}, {"", true}, {"", true}}, kActiveTab2));
  EXPECT_EQ(wb.active_sheet(), 2U);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{1U, 2U}));
}

TEST(ActiveSheetOpenRule, TheSettledSheetSurvivesAnXlsxSave) {
  const Workbook wb = ThroughXlsx(Load(BuildPackage({{"veryHidden", false}, {"", false}, {"", true}}, {})));
  EXPECT_EQ(wb.active_sheet(), 1U);
  EXPECT_EQ(Selected(wb), (std::vector<std::size_t>{1U}));
}

TEST(ActiveSheetOpenRule, AnXlsbLoadAppliesTheSameRule) {
  Workbook hidden = Workbook::create();
  hidden.add_sheet("Two");
  hidden.add_sheet("Three");
  hidden.set_active_sheet(2);
  hidden.sheet(2).mutable_view().set_visibility(SheetVisibility::kHidden);
  const Workbook unhidden = ThroughXlsb(hidden);
  EXPECT_EQ(unhidden.active_sheet(), 2U);
  EXPECT_EQ(Visibility(unhidden, 2), SheetVisibility::kVisible);

  Workbook very = Workbook::create();
  very.add_sheet("Two");
  very.add_sheet("Three");
  very.set_active_sheet(2);
  very.sheet(2).mutable_view().set_visibility(SheetVisibility::kVeryHidden);
  const Workbook moved = ThroughXlsb(very);
  EXPECT_EQ(moved.active_sheet(), 0U);
  EXPECT_EQ(Selected(moved), (std::vector<std::size_t>{0U}));
}

}  // namespace
}  // namespace formulon
