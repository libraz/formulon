//
// A loaded `x15ac:absPath` is dropped and regenerated from the location, for
// workbook roots with and without the `mc` / `x15` declarations; a foreign
// `mc:AlternateContent` of the same shape is retained verbatim.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
#include "sheet_passthrough.h"
#include "support/roundtrip_symmetry.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

struct PartFile {
  const char* path;
  std::string body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const PartFile& part : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, part.path, part.body.data(), part.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << part.path;
  }
  void* archive = nullptr;
  std::size_t size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive, &size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive),
                                static_cast<const std::uint8_t*>(archive) + size);
  mz_free(archive);
  return out;
}

constexpr const char* kExcelRootNs =
    " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x15\""
    " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"";

constexpr const char* kExcelAbsPath =
    "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
    "<mc:Choice Requires=\"x15\"><x15ac:absPath url=\"/old/folder/\" "
    "xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/></mc:Choice>"
    "</mc:AlternateContent>";

// A package whose workbook root carries `root_ns` and whose `<workbookPr>`
// is followed by `after_workbook_pr`.
std::vector<std::uint8_t> Package(const std::string& root_ns, const std::string& after_workbook_pr) {
  const std::string workbook =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"" +
      root_ns + "><workbookPr/>" + after_workbook_pr +
      "<bookViews><workbookView/></bookViews><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>"
      "</sheets></workbook>";
  return BuildZip({
      {"[Content_Types].xml",
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
       "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
       "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
       "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
       "<Override PartName=\"/xl/workbook.xml\" "
       "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
       "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
       "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
       "</Types>"},
      {"_rels/.rels",
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
       "<Relationship Id=\"rId1\" "
       "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
       "Target=\"xl/workbook.xml\"/></Relationships>"},
      {"xl/workbook.xml", workbook},
      {"xl/_rels/workbook.xml.rels",
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
       "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
       "<Relationship Id=\"rId1\" "
       "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
       "Target=\"worksheets/sheet1.xml\"/></Relationships>"},
      {"xl/worksheets/sheet1.xml",
       "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
       "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>"},
  });
}

std::string SavedWorkbookXml(const Workbook& wb) {
  auto saved = write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  std::string xml;
  EXPECT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/workbook.xml", &xml));
  return xml;
}

std::size_t Count(const std::string& text, std::string_view needle) {
  std::size_t count = 0;
  for (std::size_t at = text.find(needle); at != std::string::npos; at = text.find(needle, at + 1)) {
    ++count;
  }
  return count;
}

// First element under `node`, depth first, whose local name is `name`.
pugi::xml_node FindLocal(const pugi::xml_node& node, std::string_view name) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (local_name(child) == name) {
      return child;
    }
    if (pugi::xml_node found = FindLocal(child, name)) {
      return found;
    }
  }
  return pugi::xml_node();
}

// The saved absPath parses and every prefix on its path resolves to the
// namespace Excel uses.
void ExpectWellFormedAbsPath(const std::string& xml, const std::string& url) {
  pugi::xml_document doc;
  ASSERT_TRUE(doc.load_string(xml.c_str())) << xml;
  const pugi::xml_node abs_path = FindLocal(doc, "absPath");
  ASSERT_TRUE(abs_path) << xml;
  EXPECT_EQ(std::string(abs_path.attribute("url").value()), url);
  EXPECT_EQ(find_namespace_binding(abs_path, qname_prefix(abs_path.name())), kNsX15ac);
  const pugi::xml_node choice = abs_path.parent();
  EXPECT_EQ(local_name(choice), "Choice");
  EXPECT_EQ(std::string(choice.attribute("Requires").value()), "x15");
  EXPECT_EQ(find_namespace_binding(choice, "x15"), kNsX15);
  EXPECT_EQ(find_namespace_binding(choice, qname_prefix(choice.name())), kNsMarkupCompat);
  const pugi::xml_node alternate = choice.parent();
  EXPECT_EQ(local_name(alternate), "AlternateContent");
  EXPECT_EQ(find_namespace_binding(alternate, qname_prefix(alternate.name())), kNsMarkupCompat);
}

TEST(BookLocationAbsPath, LoadedAbsPathIsDroppedThenRegeneratedForEveryRoot) {
  const char* const roots[] = {
      kExcelRootNs,                                                                 // Excel's root: mc and x15 declared
      " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",  // no x15
      "",                                                                           // neither mc nor x15 on the root
  };
  for (const char* root : roots) {
    const std::vector<std::uint8_t> bytes = Package(root, kExcelAbsPath);
    auto loaded = read_ooxml(test::span_of(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << root << ": " << loaded.error().message;
    Workbook& wb = loaded.value().workbook;
    EXPECT_FALSE(wb.location().has_value()) << root;
    for (const WorksheetRawChild& child : wb.raw_children()) {
      EXPECT_EQ(child.xml.find("absPath"), std::string::npos) << root;
    }
    EXPECT_EQ(SavedWorkbookXml(wb).find("absPath"), std::string::npos) << root;

    ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::string("/new/p q"), "Book.xlsx"})));
    const std::string xml = SavedWorkbookXml(wb);
    EXPECT_EQ(Count(xml, "absPath"), 1U) << root << "\n" << xml;
    ExpectWellFormedAbsPath(xml, "/new/p q/");
  }
}

TEST(BookLocationAbsPath, ForeignSameShapeAlternateContentIsRetained) {
  const std::string foreign[] = {
      // Another Requires.
      "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
      "<mc:Choice Requires=\"x14\"><x15ac:absPath url=\"/k/\" "
      "xmlns:x15ac=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac\"/></mc:Choice>"
      "</mc:AlternateContent>",
      // An absPath in another namespace.
      "<mc:AlternateContent xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
      "<mc:Choice Requires=\"x15\"><o:absPath url=\"/k/\" xmlns:o=\"urn:example:other\"/></mc:Choice>"
      "</mc:AlternateContent>",
  };
  for (const std::string& element : foreign) {
    const std::vector<std::uint8_t> bytes = Package(kExcelRootNs, element);
    auto loaded = read_ooxml(test::span_of(bytes));
    ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
    const std::string xml = SavedWorkbookXml(loaded.value().workbook);
    EXPECT_EQ(Count(xml, "url=\"/k/\""), 1U) << xml;
  }
}

// The BrtACBegin / BrtAbsPath15 / BrtACEnd records of Excel's own saves in
// loc_probe6/f1 and, after Save As, loc_probe6/f2 x (Mac Excel 365).
constexpr const char* kExcelF1AbsPathHex =
    "25060100030f008097109401480000002f00550073006500720073002f006c0069006200720061007a002f004c006900"
    "620072006100720079002f0043006f006e007400610069006e006500720073002f0063006f006d002e006d0069006300"
    "72006f0073006f00660074002e0045007800630065006c002f0044006100740061002f006c006f0063005f0070007200"
    "6f006200650036002f00660031002f002600";
constexpr const char* kExcelF2AbsPathHex =
    "25060100030f0080971098014a0000002f00550073006500720073002f006c0069006200720061007a002f004c006900"
    "620072006100720079002f0043006f006e007400610069006e006500720073002f0063006f006d002e006d0069006300"
    "72006f0073006f00660074002e0045007800630065006c002f0044006100740061002f006c006f0063005f0070007200"
    "6f006200650036002f0066003200200078002f002600";

const std::string kProbe6 = "/Users/libraz/Library/Containers/com.microsoft.Excel/Data/loc_probe6";

std::string FromHex(std::string_view hex) {
  std::string out;
  const auto nibble = [](char c) { return c <= '9' ? c - '0' : c - 'a' + 10; };
  for (std::size_t i = 0; i + 1 < hex.size(); i += 2) {
    out.push_back(static_cast<char>(nibble(hex[i]) * 16 + nibble(hex[i + 1])));
  }
  return out;
}

std::string SavedWorkbookBin(const Workbook& wb) {
  auto saved = xlsb::write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  if (!saved) {
    return {};
  }
  std::string bin;
  EXPECT_TRUE(test::extract_part(test::span_of(saved.value()), "xl/workbook.bin", &bin));
  return bin;
}

// The record types of `bin` in order.
std::vector<std::uint16_t> RecordTypes(const std::string& bin) {
  std::vector<std::uint16_t> types;
  ByteSpan cursor{reinterpret_cast<const std::uint8_t*>(bin.data()), bin.size()};
  while (cursor.size > 0) {
    auto rec = xlsb::read_record(cursor);
    EXPECT_TRUE(static_cast<bool>(rec));
    if (!rec) {
      break;
    }
    types.push_back(rec.value().type);
  }
  return types;
}

TEST(BookLocationAbsPath, XlsbAbsPathFollowsTheLocation) {
  constexpr std::uint16_t kWbProp = 153;
  constexpr std::uint16_t kAbsPath15 = 2071;
  Workbook wb = Workbook::create();
  for (const std::uint16_t type : RecordTypes(SavedWorkbookBin(wb))) {
    EXPECT_NE(type, kAbsPath15);
  }
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{std::nullopt, "B.xlsb"})));
  for (const std::uint16_t type : RecordTypes(SavedWorkbookBin(wb))) {
    EXPECT_NE(type, kAbsPath15);
  }

  // Byte-identical to Excel's block, right after BrtWbProp.
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe6 + "/f1", "B.xlsb"})));
  std::string bin = SavedWorkbookBin(wb);
  EXPECT_NE(bin.find(FromHex(kExcelF1AbsPathHex)), std::string::npos);
  const std::vector<std::uint16_t> types = RecordTypes(bin);
  std::size_t at = 0;
  while (at < types.size() && types[at] != kWbProp) {
    ++at;
  }
  ASSERT_LT(at + 3U, types.size());
  EXPECT_EQ(types[at + 1U], 37U);
  EXPECT_EQ(types[at + 2U], kAbsPath15);
  EXPECT_EQ(types[at + 3U], 38U);

  // Save As to a new folder rewrites it.
  ASSERT_TRUE(static_cast<bool>(wb.set_location(WorkbookLocation{kProbe6 + "/f2 x", "B.xlsb"})));
  bin = SavedWorkbookBin(wb);
  EXPECT_NE(bin.find(FromHex(kExcelF2AbsPathHex)), std::string::npos);
  EXPECT_EQ(bin.find(FromHex(kExcelF1AbsPathHex)), std::string::npos);

  // Read back, the book has no location, and a save writes no absPath.
  auto saved = xlsb::write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(saved));
  auto loaded = xlsb::read_xlsb(test::span_of(saved.value()));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  EXPECT_FALSE(loaded.value().workbook.location().has_value());
  for (const std::uint16_t type : RecordTypes(SavedWorkbookBin(loaded.value().workbook))) {
    EXPECT_NE(type, kAbsPath15);
  }
  EXPECT_EQ(SavedWorkbookXml(loaded.value().workbook).find("absPath"), std::string::npos);
}

}  // namespace
}  // namespace io
}  // namespace formulon
