//
// Integration test: table slicers (`xl/slicers`, `xl/slicerCaches`, the
// workbook and sheet extension lists and the slicer drawing frame) that a
// load -> save -> load cycle must keep, and the typed list / add / update /
// remove edits.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/slicer_edit.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "pugixml.hpp"
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
  std::string content_type;
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

constexpr std::string_view kDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
constexpr std::string_view kRelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
constexpr std::string_view kX14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
constexpr std::string_view kX15 = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main";

std::string Rels(std::string_view body) {
  return std::string(kDecl) + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
         std::string(body) + "</Relationships>";
}

std::string Override(std::string_view path, std::string_view type) {
  return "<Override PartName=\"/" + std::string(path) + "\" ContentType=\"" + std::string(type) + "\"/>";
}

constexpr std::string_view kSlicerCt = "application/vnd.ms-excel.slicer+xml";
constexpr std::string_view kCacheCt = "application/vnd.ms-excel.slicerCache+xml";

std::string SlicerXml(bool pivot) {
  return std::string(kDecl) + "<slicers xmlns=\"" + std::string(kX14) +
         "\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"" +
         std::string(kMain) +
         "\"><slicer name=\"Region\" cache=\"Slicer_Region\" caption=\"Sales Region\" columnCount=\"2\" "
         "style=\"SlicerStyleDark1\" rowHeight=\"241300\"/>" +
         (pivot ? "<slicer name=\"Product\" cache=\"Slicer_Pivot\" caption=\"Product\" rowHeight=\"241300\"/>" : "") +
         "</slicers>";
}

std::string CacheXml() {
  return std::string(kDecl) + "<slicerCacheDefinition xmlns=\"" + std::string(kX14) +
         "\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" mc:Ignorable=\"x\" "
         "xmlns:x=\"" +
         std::string(kMain) +
         "\" name=\"Slicer_Region\" sourceName=\"Region\"><extLst><x:ext "
         "uri=\"{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}\" xmlns:x15=\"" +
         std::string(kX15) +
         "\"><x15:tableSlicerCache tableId=\"1\" column=\"1\" sortOrder=\"descending\" "
         "crossFilter=\"none\"/></x:ext></extLst></slicerCacheDefinition>";
}

std::string TableXml() {
  return std::string(kDecl) + "<table xmlns=\"" + std::string(kMain) +
         "\" id=\"1\" name=\"Sales\" displayName=\"Sales\" ref=\"A1:B4\"><autoFilter ref=\"A1:B4\"/><tableColumns "
         "count=\"2\"><tableColumn id=\"1\" name=\"Region\"/><tableColumn id=\"2\" "
         "name=\"Amount\"/></tableColumns></table>";
}

std::string DrawingXml(int from_row = 1, const char* edit_as = "oneCell", bool extra_shape = false) {
  const std::string row = std::to_string(from_row);
  const std::string to_row = std::to_string(from_row + 10);
  return std::string(kDecl) +
         "<xdr:wsDr xmlns:xdr=\"http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><xdr:twoCellAnchor editAs=\"" +
         std::string(edit_as) +
         "\">"
         "<xdr:from><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" +
         row +
         "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>6</xdr:col><xdr:colOff>0</xdr:colOff>"
         "<xdr:row>" +
         to_row +
         "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><mc:AlternateContent "
         "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"><mc:Choice "
         "xmlns:sle15=\"http://schemas.microsoft.com/office/drawing/2012/slicer\" Requires=\"sle15\">"
         "<xdr:graphicFrame macro=\"\"><xdr:nvGraphicFramePr><xdr:cNvPr id=\"2\" name=\"Region\"/>"
         "<xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"0\" cy=\"0\"/>"
         "</xdr:xfrm><a:graphic><a:graphicData uri=\"http://schemas.microsoft.com/office/drawing/2010/slicer\">"
         "<sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" name=\"Region\"/>"
         "</a:graphicData></a:graphic></xdr:graphicFrame></mc:Choice><mc:Fallback><xdr:sp macro=\"\" textlink=\"\">"
         "<xdr:nvSpPr><xdr:cNvPr id=\"0\" name=\"\"/><xdr:cNvSpPr><a:spLocks noTextEdit=\"1\"/></xdr:cNvSpPr>"
         "</xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"1828800\" cy=\"2438400\"/></a:xfrm>"
         "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr><xdr:txBody><a:bodyPr/><a:lstStyle/><a:p>"
         "<a:r><a:rPr lang=\"en-US\"/><a:t>This shape represents a slicer.</a:t></a:r></a:p></xdr:txBody></xdr:sp>"
         "</mc:Fallback></mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>" +
         std::string(extra_shape ? "<xdr:oneCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>"
                                   "<xdr:row>30</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx=\"10\" "
                                   "cy=\"10\"/><xdr:sp macro=\"\" textlink=\"\"><xdr:nvSpPr><xdr:cNvPr id=\"9\" "
                                   "name=\"Box\"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/>"
                                   "</xdr:oneCellAnchor>"
                                 : "") +
         "</xdr:wsDr>";
}

constexpr std::string_view kRelSlicer = "http://schemas.microsoft.com/office/2007/relationships/slicer";
constexpr std::string_view kRelCache = "http://schemas.microsoft.com/office/2007/relationships/slicerCache";

std::string PivotCacheXml() {
  return std::string(kDecl) + "<slicerCacheDefinition xmlns=\"" + std::string(kX14) +
         "\" name=\"Slicer_Pivot\" sourceName=\"Product\"><pivotTables><pivotTable tabId=\"1\" "
         "name=\"PivotTable1\"/></pivotTables><data><tabular pivotCacheId=\"1\"><items count=\"1\"><i x=\"0\" "
         "s=\"1\"/></items></tabular></data></slicerCacheDefinition>";
}

struct Options {
  bool slicer = true;  ///< False: the table only.
  bool pivot = false;  ///< Adds a slicer over a pivot cache to the same slicer part.
  /// Source rId of the sheet's slicer relationship.
  const char* slicer_rid = "rId3";
  /// `editAs` of the slicer frame, and whether the drawing also holds an unrelated shape.
  const char* edit_as = "oneCell";
  bool extra_shape = false;
};

/// Workbook + sheet with one table and, by default, one table slicer as a desktop spreadsheet writes it.
std::vector<std::uint8_t> SlicerPackage(const Options& opt = {}) {
  std::string content_types =
      std::string(kDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
      Override("xl/workbook.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") +
      Override("xl/worksheets/sheet1.xml",
               "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml") +
      Override("xl/tables/table1.xml", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml");
  std::string workbook_rels =
      "<Relationship Id=\"rId1\" Type=\"" + std::string(kRelNs) + "/worksheet\" Target=\"worksheets/sheet1.xml\"/>";
  std::string sheet_rels =
      "<Relationship Id=\"rId2\" Type=\"" + std::string(kRelNs) + "/table\" Target=\"../tables/table1.xml\"/>";
  std::string defined_names;
  std::string book_ext;
  std::vector<PartFile> parts;
  if (opt.slicer) {
    content_types += Override("xl/drawings/drawing1.xml", "application/vnd.openxmlformats-officedocument.drawing+xml") +
                     Override("xl/slicers/slicer1.xml", kSlicerCt) +
                     Override("xl/slicerCaches/slicerCache1.xml", kCacheCt);
    workbook_rels +=
        "<Relationship Id=\"rId7\" Type=\"" + std::string(kRelCache) + "\" Target=\"slicerCaches/slicerCache1.xml\"/>";
    sheet_rels += "<Relationship Id=\"rId1\" Type=\"" + std::string(kRelNs) +
                  "/drawing\" Target=\"../drawings/drawing1.xml\"/><Relationship Id=\"" + opt.slicer_rid +
                  "\" Type=\"" + std::string(kRelSlicer) + "\" Target=\"../slicers/slicer1.xml\"/>";
    defined_names = "<definedNames><definedName name=\"Slicer_Region\">#N/A</definedName>" +
                    std::string(opt.pivot ? "<definedName name=\"Slicer_Pivot\">#N/A</definedName>" : "") +
                    "</definedNames>";
    book_ext = "<extLst><ext uri=\"{46BE6895-7355-4a93-B00E-2C351335B9C9}\" xmlns:x15=\"" + std::string(kX15) +
               "\"><x15:slicerCaches xmlns:x14=\"" + std::string(kX14) + "\"><x14:slicerCache r:id=\"rId7\"/>" +
               "</x15:slicerCaches></ext>";
    if (opt.pivot) {
      content_types += Override("xl/slicerCaches/slicerCache2.xml", kCacheCt);
      workbook_rels += "<Relationship Id=\"rId8\" Type=\"" + std::string(kRelCache) +
                       "\" Target=\"slicerCaches/slicerCache2.xml\"/>";
      book_ext += "<ext uri=\"{BBE1A952-AA13-448e-AADC-164F8A28A991}\" xmlns:x14=\"" + std::string(kX14) +
                  "\"><x14:slicerCaches><x14:slicerCache r:id=\"rId8\"/></x14:slicerCaches></ext>";
      parts.push_back({"xl/slicerCaches/slicerCache2.xml", PivotCacheXml(), {}});
    }
    book_ext += "</extLst>";
    parts.push_back({"xl/drawings/drawing1.xml", DrawingXml(1, opt.edit_as, opt.extra_shape), {}});
    parts.push_back({"xl/slicers/slicer1.xml", SlicerXml(opt.pivot), {}});
    parts.push_back({"xl/slicerCaches/slicerCache1.xml", CacheXml(), {}});
  }
  content_types += "</Types>";
  // `drawing` precedes `tableParts` in the worksheet schema.
  std::string sheet_body =
      "<sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>Region</t></is></c><c r=\"B1\" "
      "t=\"inlineStr\"><is><t>Amount</t></is></c></row><row r=\"2\"><c r=\"A2\" t=\"inlineStr\"><is><t>N</t></is></c>"
      "<c r=\"B2\"><v>1</v></c></row></sheetData>";
  sheet_body += opt.slicer ? "<drawing r:id=\"rId1\"/>" : "";
  sheet_body += "<tableParts count=\"1\"><tablePart r:id=\"rId2\"/></tableParts>";
  if (opt.slicer) {
    sheet_body += "<extLst><ext uri=\"{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}\" xmlns:x15=\"" + std::string(kX15) +
                  "\"><x14:slicerList xmlns:x14=\"" + std::string(kX14) + "\"><x14:slicer r:id=\"" + opt.slicer_rid +
                  "\"/></x14:slicerList></ext></extLst>";
  }
  const std::string workbook = std::string(kDecl) + "<workbook xmlns=\"" + std::string(kMain) + "\" xmlns:r=\"" +
                               std::string(kRelNs) +
                               "\"><sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                               defined_names + book_ext + "</workbook>";
  const std::string sheet = std::string(kDecl) + "<worksheet xmlns=\"" + std::string(kMain) + "\" xmlns:r=\"" +
                            std::string(kRelNs) + "\">" + sheet_body + "</worksheet>";
  std::vector<PartFile> all = {
      {"[Content_Types].xml", content_types + "", {}},
      {"_rels/.rels",
       Rels("<Relationship Id=\"rId1\" Type=\"" + std::string(kRelNs) +
            "/officeDocument\" Target=\"xl/workbook.xml\"/>"),
       {}},
      {"xl/workbook.xml", workbook, {}},
      {"xl/_rels/workbook.xml.rels", Rels(workbook_rels), {}},
      {"xl/worksheets/sheet1.xml", sheet, {}},
      {"xl/worksheets/_rels/sheet1.xml.rels", Rels(sheet_rels), {}},
      {"xl/tables/table1.xml", TableXml(), {}},
  };
  all.insert(all.end(), parts.begin(), parts.end());
  return BuildZip(all);
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
  if (!entry_or) {
    return {};
  }
  return std::string(reinterpret_cast<const char*>(entry_or.value().data()), entry_or.value().size());
}

bool Has(const std::string& text, std::string_view needle) {
  return text.find(needle) != std::string::npos;
}

bool HasPart(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  io::ZipReader zip;
  return static_cast<bool>(zip.open(SpanOf(bytes))) && zip.has_entry(name);
}

std::string DirOf(const std::string& path) {
  const std::size_t slash = path.rfind('/');
  return slash == std::string::npos ? std::string() : path.substr(0, slash + 1);
}

/// Resolves `target` against the part owning a rels part (`a/_rels/b.xml.rels` -> `a/`).
std::string Resolve(const std::string& rels_path, const std::string& target) {
  if (!target.empty() && target[0] == '/') {
    return target.substr(1);
  }
  std::string dir = DirOf(rels_path);
  dir = dir.size() >= 6 ? dir.substr(0, dir.size() - 6) : std::string();  // drop "_rels/"
  std::vector<std::string> pieces;
  std::string full = dir + target;
  std::size_t at = 0;
  while (at <= full.size()) {
    const std::size_t next = full.find('/', at);
    const std::string piece = full.substr(at, next == std::string::npos ? std::string::npos : next - at);
    if (piece == "..") {
      if (!pieces.empty()) {
        pieces.pop_back();
      }
    } else if (!piece.empty() && piece != ".") {
      pieces.push_back(piece);
    }
    if (next == std::string::npos) {
      break;
    }
    at = next + 1;
  }
  std::string out;
  for (const std::string& piece : pieces) {
    out += (out.empty() ? "" : "/") + piece;
  }
  return out;
}

/// The written package must be closed: every internal relationship target is a part, every part has a content
/// type, every override names a part, and every `r:id` of the workbook and sheet XML names a relationship.
void ExpectSelfConsistent(const std::vector<std::uint8_t>& bytes) {
  io::ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  pugi::xml_document types;
  ASSERT_TRUE(types.load_string(ReadPart(bytes, "[Content_Types].xml").c_str()));
  std::set<std::string> defaults;
  std::set<std::string> overrides;
  for (pugi::xml_node node : types.document_element().children()) {
    if (std::string_view(node.name()) == "Default") {
      defaults.insert(node.attribute("Extension").value());
    } else if (std::string_view(node.name()) == "Override") {
      overrides.insert(std::string(node.attribute("PartName").value()).substr(1));
    }
  }
  for (const std::string& part : overrides) {
    EXPECT_TRUE(zip.has_entry(part)) << "override without a part: " << part;
  }
  for (std::size_t i = 0; i < zip.entry_count(); ++i) {
    const std::string name(zip.entry_name(i));
    if (name == "[Content_Types].xml" || (!name.empty() && name.back() == '/')) {
      continue;
    }
    const std::size_t dot = name.rfind('.');
    EXPECT_TRUE(overrides.count(name) != 0U || (dot != std::string::npos && defaults.count(name.substr(dot + 1)) != 0U))
        << "part without a content type: " << name;
    if (name.size() < 5 || name.compare(name.size() - 5, 5, ".rels") != 0) {
      continue;
    }
    pugi::xml_document rels;
    ASSERT_TRUE(rels.load_string(ReadPart(bytes, name).c_str())) << name;
    for (pugi::xml_node rel : rels.document_element().children()) {
      if (std::string_view(rel.attribute("TargetMode").value()) != "External") {
        EXPECT_TRUE(zip.has_entry(Resolve(name, rel.attribute("Target").value())))
            << name << " names a missing part: " << rel.attribute("Target").value();
      }
    }
  }
  for (const char* owner : {"xl/workbook.xml", "xl/worksheets/sheet1.xml"}) {
    const std::string xml = ReadPart(bytes, owner);
    const std::string rels_path = DirOf(owner) + "_rels/" + std::string(owner).substr(DirOf(owner).size()) + ".rels";
    const std::string rels = ReadPart(bytes, rels_path);
    for (std::size_t at = xml.find(" r:id=\""); at != std::string::npos; at = xml.find(" r:id=\"", at + 1)) {
      const std::size_t begin = at + 7;
      const std::string id = xml.substr(begin, xml.find('"', begin) - begin);
      EXPECT_TRUE(Has(rels, "Id=\"" + id + "\"")) << owner << " names a missing relationship " << id;
    }
  }
}

/// The relationship id of the first relationship in `rels_xml` whose `Type` ends with `type_suffix`.
std::string RidOfType(const std::string& rels_xml, std::string_view type_suffix) {
  pugi::xml_document doc;
  doc.load_string(rels_xml.c_str());
  for (pugi::xml_node rel : doc.document_element().children()) {
    const std::string_view type = rel.attribute("Type").value();
    if (type.size() >= type_suffix.size() && type.substr(type.size() - type_suffix.size()) == type_suffix) {
      return rel.attribute("Id").value();
    }
  }
  return {};
}

SlicerSpec Spec(std::uint32_t fields) {
  SlicerSpec spec;
  spec.fields = fields;
  spec.from = AnchorPoint{3, 4, 0, 0};
  spec.to = AnchorPoint{13, 6, 0, 0};
  return spec;
}

// ---------------------------------------------------------------------------
// Loading and resaving
// ---------------------------------------------------------------------------

TEST(SlicerFidelity, ListReportsEveryModelledField) {
  const Workbook wb = LoadOrDie(SlicerPackage());
  auto slicers = list_slicers(wb, 0);
  ASSERT_TRUE(static_cast<bool>(slicers));
  ASSERT_EQ(slicers.value().size(), 1U);
  const SlicerInfo& s = slicers.value()[0];
  EXPECT_EQ(s.name, "Region");
  EXPECT_EQ(s.caption, "Sales Region");
  EXPECT_EQ(s.cache_name, "Slicer_Region");
  EXPECT_EQ(s.table_name, "Sales");
  EXPECT_EQ(s.column_name, "Region");
  EXPECT_EQ(s.style, "SlicerStyleDark1");
  EXPECT_EQ(s.column_count, 2U);
  EXPECT_EQ(s.row_height, 241300);
  EXPECT_EQ(s.sort_order, SlicerSortOrder::kDescending);
  EXPECT_EQ(s.cross_filter, SlicerCrossFilter::kNone);
  EXPECT_EQ(s.object.object_id, 2U);
  EXPECT_EQ(s.object.kind, DrawingObjectKind::kGraphicFrame);
  EXPECT_EQ(s.object.anchor_kind, AnchorKind::kTwoCell);
  EXPECT_EQ(s.object.from.row, 1U);
  EXPECT_EQ(s.object.from.col, 4U);
  EXPECT_EQ(s.object.to.row, 11U);
  EXPECT_EQ(s.object.to.col, 6U);
}

std::vector<std::uint8_t> ReadExcelFixture(const char* name) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  out.resize(static_cast<std::size_t>(std::ftell(file)));
  std::fseek(file, 0, SEEK_SET);
  if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(file);
  return out;
}

// A table slicer inserted and saved by Excel 365 (ja-JP, Mac).
TEST(SlicerFidelity, ExcelSavedTableSlicerListsAndSurvivesAResave) {
  const Workbook wb = LoadOrDie(ReadExcelFixture("table_slicer.xlsx"));
  const auto check = [](const Workbook& source) {
    auto slicers = list_slicers(source, 0);
    ASSERT_TRUE(static_cast<bool>(slicers));
    ASSERT_EQ(slicers.value().size(), 1U);
    const SlicerInfo& s = slicers.value()[0];
    EXPECT_EQ(s.name, "Product");
    EXPECT_EQ(s.cache_name, "\u30b9\u30e9\u30a4\u30b5\u30fc_Product");
    EXPECT_EQ(s.table_name, "Sales");
    EXPECT_EQ(s.column_name, "Product");
    EXPECT_EQ(s.row_height, 230716);
    EXPECT_EQ(s.object.from.row, 26U);
    EXPECT_EQ(s.object.to.col, 30U);
  };
  check(wb);
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_TRUE(Has(ReadPart(saved, "xl/worksheets/sheet1.xml"), "<x14:slicerList"));
  check(LoadOrDie(saved));
}

TEST(SlicerFidelity, PivotSlicersAreNotListed) {
  const Workbook wb = LoadOrDie(SlicerPackage({true, true}));
  auto slicers = list_slicers(wb, 0);
  ASSERT_TRUE(static_cast<bool>(slicers));
  ASSERT_EQ(slicers.value().size(), 1U);
  EXPECT_EQ(slicers.value()[0].name, "Region");
}

TEST(SlicerFidelity, ResaveKeepsEveryPartVerbatimAndLinked) {
  for (const char* rid : {"rId3", "rId1", "rId9"}) {
    const auto saved = SaveOrDie(LoadOrDie(SlicerPackage({true, true, rid})));
    ExpectSelfConsistent(saved);
    EXPECT_EQ(ReadPart(saved, "xl/slicers/slicer1.xml"), SlicerXml(true));
    EXPECT_EQ(ReadPart(saved, "xl/slicerCaches/slicerCache1.xml"), CacheXml());
    EXPECT_EQ(ReadPart(saved, "xl/slicerCaches/slicerCache2.xml"), PivotCacheXml());
    EXPECT_EQ(ReadPart(saved, "xl/drawings/drawing1.xml"), DrawingXml());
    const std::string book = ReadPart(saved, "xl/workbook.xml");
    EXPECT_TRUE(Has(book, "<definedName name=\"Slicer_Region\">#N/A</definedName>")) << book;
    // The extension entries follow the relationship ids the writer assigned.
    const std::string book_rels = ReadPart(saved, "xl/_rels/workbook.xml.rels");
    EXPECT_TRUE(Has(book_rels, "Target=\"slicerCaches/slicerCache1.xml\""));
    pugi::xml_document rels;
    rels.load_string(book_rels.c_str());
    for (pugi::xml_node rel : rels.document_element().children()) {
      const std::string_view target = rel.attribute("Target").value();
      if (target.find("slicerCache") != std::string_view::npos) {
        EXPECT_TRUE(Has(book, "<x14:slicerCache r:id=\"" + std::string(rel.attribute("Id").value()) + "\"/>")) << book;
      }
    }
    const std::string sheet = ReadPart(saved, "xl/worksheets/sheet1.xml");
    const std::string sheet_rels = ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels");
    EXPECT_TRUE(Has(sheet, "<x14:slicer r:id=\"" + RidOfType(sheet_rels, "/slicer") + "\"/>")) << sheet;
    EXPECT_TRUE(Has(sheet, "<drawing r:id=\"" + RidOfType(sheet_rels, "/drawing") + "\"/>")) << sheet;
    EXPECT_TRUE(Has(sheet, "<tablePart r:id=\"" + RidOfType(sheet_rels, "/table") + "\"/>")) << sheet;
  }
}

TEST(SlicerFidelity, ResaveSurvivesASecondRoundTrip) {
  const auto once = SaveOrDie(LoadOrDie(SlicerPackage({true, true})));
  const auto twice = SaveOrDie(LoadOrDie(once));
  ExpectSelfConsistent(twice);
  EXPECT_EQ(ReadPart(twice, "xl/slicers/slicer1.xml"), SlicerXml(true));
  EXPECT_EQ(ReadPart(twice, "xl/workbook.xml"), ReadPart(once, "xl/workbook.xml"));
}

// ---------------------------------------------------------------------------
// Structural edits and renames
// ---------------------------------------------------------------------------

TEST(SlicerFidelity, FrameShiftsWithRowAndColumnEdits) {
  Workbook wb = LoadOrDie(SlicerPackage());
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 3)));
  auto slicers = list_slicers(wb, 0);
  EXPECT_EQ(slicers.value()[0].object.from.row, 4U);
  EXPECT_EQ(slicers.value()[0].object.to.row, 14U);
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 2)));
  slicers = list_slicers(wb, 0);
  EXPECT_EQ(slicers.value()[0].object.from.col, 6U);
  EXPECT_EQ(slicers.value()[0].object.to.col, 8U);

  Workbook again = LoadOrDie(SaveOrDie(wb));
  ASSERT_TRUE(static_cast<bool>(again.delete_rows(0, 0, 2)));
  ASSERT_TRUE(static_cast<bool>(again.delete_cols(0, 0, 1)));
  slicers = list_slicers(again, 0);
  ASSERT_EQ(slicers.value().size(), 1U);
  EXPECT_EQ(slicers.value()[0].object.from.row, 2U);
  EXPECT_EQ(slicers.value()[0].object.from.col, 5U);
  const auto saved = SaveOrDie(again);
  ExpectSelfConsistent(saved);
  EXPECT_TRUE(Has(ReadPart(saved, "xl/drawings/drawing1.xml"), "<xdr:row>2</xdr:row>"));
}

TEST(SlicerFidelity, TableRenameKeepsTheCacheBinding) {
  Workbook wb = LoadOrDie(SlicerPackage());
  ASSERT_TRUE(static_cast<bool>(wb.rename_table(0, "Orders", "Orders")));
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_EQ(ReadPart(saved, "xl/slicerCaches/slicerCache1.xml"), CacheXml());
  EXPECT_TRUE(Has(ReadPart(saved, "xl/tables/table1.xml"), "id=\"1\" name=\"Orders\""));
  auto slicers = list_slicers(LoadOrDie(saved), 0);
  ASSERT_EQ(slicers.value().size(), 1U);
  EXPECT_EQ(slicers.value()[0].table_name, "Orders");
  EXPECT_EQ(slicers.value()[0].column_name, "Region");
}

TEST(SlicerFidelity, RemovingTheTableRemovesItsSlicersAndCaches) {
  Workbook wb = LoadOrDie(SlicerPackage({true, true}));
  remove_table_slicers(wb, wb.tables()[0].id);
  wb.mutable_tables().clear();
  EXPECT_TRUE(list_slicers(wb, 0).value().empty());
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_FALSE(HasPart(saved, "xl/slicerCaches/slicerCache1.xml"));
  const std::string book = ReadPart(saved, "xl/workbook.xml");
  EXPECT_FALSE(Has(book, "Slicer_Region")) << book;
  // The pivot slicer and its cache are not the table's.
  EXPECT_TRUE(HasPart(saved, "xl/slicerCaches/slicerCache2.xml"));
  EXPECT_TRUE(Has(book, "Slicer_Pivot")) << book;
  const std::string slicers = ReadPart(saved, "xl/slicers/slicer1.xml");
  EXPECT_FALSE(Has(slicers, "Slicer_Region")) << slicers;
  EXPECT_TRUE(Has(slicers, "Slicer_Pivot")) << slicers;
  EXPECT_FALSE(Has(ReadPart(saved, "xl/drawings/drawing1.xml"), "sle:slicer"));
}

TEST(SlicerFidelity, RemovingTheHostSheetRemovesItsSlicers) {
  Workbook wb = LoadOrDie(SlicerPackage());
  wb.add_sheet("Other");
  ASSERT_TRUE(static_cast<bool>(wb.remove_sheet(0)));
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_FALSE(HasPart(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(HasPart(saved, "xl/slicers/slicer1.xml"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/workbook.xml"), "Slicer_Region"));
}

// ---------------------------------------------------------------------------
// Typed edits
// ---------------------------------------------------------------------------

TEST(SlicerEdit, AddCreatesEveryPartOnASheetWithoutADrawing) {
  Workbook wb = LoadOrDie(SlicerPackage({false}));
  ASSERT_TRUE(list_slicers(wb, 0).value().empty());
  SlicerSpec spec = Spec(slicer_field::kAnchor | slicer_field::kCaption | slicer_field::kStyle |
                         slicer_field::kColumnCount | slicer_field::kSortOrder | slicer_field::kCrossFilter);
  spec.caption = "Pick a region";
  spec.style = "SlicerStyleLight2";
  spec.column_count = 3;
  spec.sort_order = SlicerSortOrder::kDescending;
  spec.cross_filter = SlicerCrossFilter::kShowItemsWithNoData;
  auto id = add_slicer(wb, 0, 0, 0, spec);
  ASSERT_TRUE(static_cast<bool>(id)) << id.error().message;
  EXPECT_EQ(id.value(), 2U);

  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  const std::string cache = ReadPart(saved, "xl/slicerCaches/slicerCache1.xml");
  EXPECT_TRUE(Has(cache, "name=\"Slicer_Region\" sourceName=\"Region\"")) << cache;
  EXPECT_TRUE(Has(cache, "tableId=\"1\" column=\"1\" sortOrder=\"descending\" crossFilter=\"showItemsWithNoData\""))
      << cache;
  const std::string slicers = ReadPart(saved, "xl/slicers/slicer1.xml");
  EXPECT_TRUE(Has(slicers, "name=\"Region\" cache=\"Slicer_Region\" caption=\"Pick a region\"")) << slicers;
  EXPECT_TRUE(Has(slicers, "columnCount=\"3\"")) << slicers;
  EXPECT_TRUE(Has(slicers, "style=\"SlicerStyleLight2\"")) << slicers;
  EXPECT_TRUE(Has(slicers, "rowHeight=\"241300\"")) << slicers;
  const std::string drawing = ReadPart(saved, "xl/drawings/drawing1.xml");
  EXPECT_TRUE(Has(drawing, "<mc:AlternateContent")) << drawing;
  EXPECT_TRUE(Has(drawing, "<mc:Fallback>")) << drawing;
  EXPECT_TRUE(Has(drawing,
                  "<sle:slicer xmlns:sle=\"http://schemas.microsoft.com/office/drawing/2010/slicer\" "
                  "name=\"Region\"/>"))
      << drawing;
  EXPECT_TRUE(Has(ReadPart(saved, "[Content_Types].xml"), "application/vnd.ms-excel.slicerCache+xml"));
  EXPECT_TRUE(Has(ReadPart(saved, "[Content_Types].xml"), "application/vnd.ms-excel.slicer+xml"));
  EXPECT_TRUE(Has(ReadPart(saved, "xl/workbook.xml"), "<definedName name=\"Slicer_Region\">#N/A</definedName>"));
  // Excel ignores a table slicer whose worksheet list is not `x14:slicerList`.
  const std::string sheet = ReadPart(saved, "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(sheet, "<x14:slicerList xmlns:x14=")) << sheet;
  EXPECT_FALSE(Has(sheet, "x15:slicerList")) << sheet;

  // It reads back, from the saved package as well.
  const Workbook reloaded = LoadOrDie(saved);
  const auto check = [](const Workbook& source) {
    auto slicers_back = list_slicers(source, 0);
    ASSERT_EQ(slicers_back.value().size(), 1U);
    const SlicerInfo& s = slicers_back.value()[0];
    EXPECT_EQ(s.name, "Region");
    EXPECT_EQ(s.caption, "Pick a region");
    EXPECT_EQ(s.column_count, 3U);
    EXPECT_EQ(s.sort_order, SlicerSortOrder::kDescending);
    EXPECT_EQ(s.cross_filter, SlicerCrossFilter::kShowItemsWithNoData);
    EXPECT_EQ(s.object.from.row, 3U);
    EXPECT_EQ(s.object.to.col, 6U);
  };
  check(wb);
  check(reloaded);
}

TEST(SlicerEdit, AddSecondSlicerSharesThePartAndGetsUniqueNames) {
  Workbook wb = LoadOrDie(SlicerPackage());
  ASSERT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 0, Spec(slicer_field::kAnchor))));
  auto id = add_slicer(wb, 0, 0, 1, Spec(slicer_field::kAnchor | slicer_field::kName));
  ASSERT_FALSE(static_cast<bool>(id));  // an explicit empty name is refused
  id = add_slicer(wb, 0, 0, 1, Spec(slicer_field::kAnchor));
  ASSERT_TRUE(static_cast<bool>(id));
  auto slicers = list_slicers(wb, 0);
  ASSERT_EQ(slicers.value().size(), 3U);
  EXPECT_EQ(slicers.value()[1].name, "Region 1");
  EXPECT_EQ(slicers.value()[1].cache_name, "Slicer_Region1");
  EXPECT_EQ(slicers.value()[2].name, "Amount");
  EXPECT_EQ(slicers.value()[2].column_name, "Amount");
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_FALSE(HasPart(saved, "xl/slicers/slicer2.xml"));
  EXPECT_TRUE(HasPart(saved, "xl/slicerCaches/slicerCache3.xml"));
  EXPECT_EQ(list_slicers(LoadOrDie(saved), 0).value().size(), 3U);
}

TEST(SlicerEdit, AddRefusesBadInput) {
  Workbook wb = LoadOrDie(SlicerPackage());
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 0, Spec(0))));                      // no anchor
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 1, 0, Spec(slicer_field::kAnchor))));  // no such table
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 2, Spec(slicer_field::kAnchor))));  // no such column
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 5, 0, 0, Spec(slicer_field::kAnchor))));  // no such sheet
  SlicerSpec backwards = Spec(slicer_field::kAnchor);
  std::swap(backwards.from, backwards.to);
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 0, backwards)));
  SlicerSpec count = Spec(slicer_field::kAnchor | slicer_field::kColumnCount);
  count.column_count = 0;
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 0, count)));
  SlicerSpec text = Spec(slicer_field::kAnchor | slicer_field::kCaption);
  text.caption = "a\nb";
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 0, text)));
  EXPECT_FALSE(static_cast<bool>(add_slicer(wb, 0, 0, 0, Spec(slicer_field::kAnchor | (1U << 20)))));
  EXPECT_EQ(list_slicers(wb, 0).value().size(), 1U);
}

TEST(SlicerEdit, UpdateChangesOnlyTheNamedFields) {
  Workbook wb = LoadOrDie(SlicerPackage());
  SlicerSpec spec = Spec(slicer_field::kCaption | slicer_field::kStyle | slicer_field::kColumnCount |
                         slicer_field::kAnchor | slicer_field::kSortOrder);
  spec.caption = "Renamed";
  spec.style.clear();
  spec.column_count = 1;
  spec.sort_order = SlicerSortOrder::kAscending;
  ASSERT_TRUE(static_cast<bool>(update_slicer(wb, 0, 0, spec)));
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  const std::string slicers = ReadPart(saved, "xl/slicers/slicer1.xml");
  EXPECT_TRUE(Has(slicers, "caption=\"Renamed\"")) << slicers;
  EXPECT_FALSE(Has(slicers, "style=")) << slicers;
  EXPECT_FALSE(Has(slicers, "columnCount=")) << slicers;
  EXPECT_TRUE(Has(slicers, "rowHeight=\"241300\"")) << slicers;
  const std::string cache = ReadPart(saved, "xl/slicerCaches/slicerCache1.xml");
  EXPECT_FALSE(Has(cache, "sortOrder=")) << cache;
  EXPECT_TRUE(Has(cache, "crossFilter=\"none\"")) << cache;  // not in the mask
  EXPECT_TRUE(Has(cache, "mc:Ignorable=\"x\"")) << cache;    // unmodelled content stays
  const std::string drawing = ReadPart(saved, "xl/drawings/drawing1.xml");
  EXPECT_TRUE(Has(drawing, "<xdr:row>3</xdr:row>")) << drawing;
  EXPECT_TRUE(Has(drawing, "<xdr:row>13</xdr:row>")) << drawing;
  EXPECT_TRUE(Has(drawing, "<xdr:cNvPr id=\"2\" name=\"Region\"/>")) << drawing;
  auto back = list_slicers(LoadOrDie(saved), 0);
  EXPECT_EQ(back.value()[0].caption, "Renamed");
  EXPECT_EQ(back.value()[0].object.object_id, 2U);
  EXPECT_EQ(back.value()[0].object.from.col, 4U);

  EXPECT_FALSE(static_cast<bool>(update_slicer(wb, 0, 1, spec)));
  EXPECT_FALSE(static_cast<bool>(update_slicer(wb, 0, 0, Spec(slicer_field::kName))));
}

TEST(SlicerEdit, RemoveDropsTheFrameAndTheCacheWhenUnshared) {
  Workbook wb = LoadOrDie(SlicerPackage({true, true}));
  ASSERT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 1, Spec(slicer_field::kAnchor))));
  ASSERT_EQ(list_slicers(wb, 0).value().size(), 2U);
  ASSERT_TRUE(static_cast<bool>(remove_slicer(wb, 0, 0)));
  auto left = list_slicers(wb, 0);
  ASSERT_EQ(left.value().size(), 1U);
  EXPECT_EQ(left.value()[0].name, "Amount");
  auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_FALSE(HasPart(saved, "xl/slicerCaches/slicerCache1.xml"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/workbook.xml"), "Slicer_Region\""));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/drawings/drawing1.xml"), "name=\"Region\""));
  EXPECT_TRUE(Has(ReadPart(saved, "xl/slicers/slicer1.xml"), "Slicer_Pivot"));

  // The last table slicer goes; the pivot slicer keeps the slicer part alive.
  ASSERT_TRUE(static_cast<bool>(remove_slicer(wb, 0, 0)));
  EXPECT_TRUE(list_slicers(wb, 0).value().empty());
  saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_TRUE(HasPart(saved, "xl/slicers/slicer1.xml"));
  EXPECT_TRUE(HasPart(saved, "xl/slicerCaches/slicerCache2.xml"));
  EXPECT_FALSE(static_cast<bool>(remove_slicer(wb, 0, 0)));
}

TEST(SlicerEdit, RemovingTheLastSlicerRemovesThePartsAndExtensionEntries) {
  Workbook wb = LoadOrDie(SlicerPackage());
  ASSERT_TRUE(static_cast<bool>(remove_slicer(wb, 0, 0)));
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  for (const char* part : {"xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml"}) {
    EXPECT_FALSE(HasPart(saved, part)) << part;
  }
  EXPECT_FALSE(Has(ReadPart(saved, "xl/worksheets/sheet1.xml"), "extLst"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/workbook.xml"), "extLst"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"), "slicer"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/_rels/workbook.xml.rels"), "slicer"));
}

TEST(SlicerEdit, AddThenRemoveLeavesNoSlicerState) {
  Workbook wb = LoadOrDie(SlicerPackage({false}));
  const auto before = SaveOrDie(wb);
  ASSERT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 0, Spec(slicer_field::kAnchor))));
  ASSERT_TRUE(static_cast<bool>(remove_slicer(wb, 0, 0)));
  const auto after = SaveOrDie(wb);
  ExpectSelfConsistent(after);
  // Every part the add created is gone, the drawing it created included.
  for (const char* part : {"xl/workbook.xml", "xl/worksheets/sheet1.xml", "xl/_rels/workbook.xml.rels",
                           "xl/worksheets/_rels/sheet1.xml.rels"}) {
    EXPECT_EQ(ReadPart(after, part), ReadPart(before, part)) << part;
  }
  for (const char* part : {"xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml", "xl/drawings/drawing1.xml"}) {
    EXPECT_FALSE(HasPart(after, part)) << part;
  }
  EXPECT_FALSE(Has(ReadPart(after, "[Content_Types].xml"), "drawing"));
}

TEST(SlicerEdit, RemovingTheLastSlicerKeepsADrawingThatHoldsOtherObjects) {
  Workbook wb = LoadOrDie(SlicerPackage({true, false, "rId3", "oneCell", true}));
  ASSERT_TRUE(static_cast<bool>(remove_slicer(wb, 0, 0)));
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  const std::string drawing = ReadPart(saved, "xl/drawings/drawing1.xml");
  EXPECT_TRUE(Has(drawing, "name=\"Box\"")) << drawing;
  EXPECT_FALSE(Has(drawing, "sle:slicer")) << drawing;
  EXPECT_FALSE(HasPart(saved, "xl/slicers/slicer1.xml"));
}

// ---------------------------------------------------------------------------
// Frames deleted by structural edits
// ---------------------------------------------------------------------------

TEST(SlicerFidelity, DeletingTheFramesRowsRemovesTheSlicerAndItsCache) {
  Workbook wb = LoadOrDie(SlicerPackage({true, true, "rId3", "twoCell"}));
  ASSERT_TRUE(static_cast<bool>(add_slicer(wb, 0, 0, 1, Spec(slicer_field::kAnchor))));  // frame on rows 3..13
  ASSERT_EQ(list_slicers(wb, 0).value().size(), 2U);
  // The first frame spans rows 1..11 and moves with its cells; deleting them deletes it.
  ASSERT_TRUE(static_cast<bool>(wb.delete_rows(0, 1, 11)));
  auto left = list_slicers(wb, 0);
  ASSERT_EQ(left.value().size(), 1U);
  EXPECT_EQ(left.value()[0].name, "Amount");
  EXPECT_NE(left.value()[0].object.object_id, 0U);
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  EXPECT_FALSE(HasPart(saved, "xl/slicerCaches/slicerCache1.xml"));
  const std::string book = ReadPart(saved, "xl/workbook.xml");
  EXPECT_FALSE(Has(book, "\"Slicer_Region\"")) << book;
  EXPECT_TRUE(Has(book, "Slicer_Amount")) << book;
  EXPECT_TRUE(Has(book, "Slicer_Pivot")) << book;
  const std::string slicers = ReadPart(saved, "xl/slicers/slicer1.xml");
  EXPECT_FALSE(Has(slicers, "name=\"Region\"")) << slicers;
  EXPECT_TRUE(Has(slicers, "Slicer_Pivot")) << slicers;
}

TEST(SlicerFidelity, DeletingTheLastFramesCellsLeavesNoSlicerState) {
  Workbook wb = LoadOrDie(SlicerPackage({true, false, "rId3", "twoCell"}));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 4, 3)));
  EXPECT_TRUE(list_slicers(wb, 0).value().empty());
  const auto saved = SaveOrDie(wb);
  ExpectSelfConsistent(saved);
  for (const char* part : {"xl/slicers/slicer1.xml", "xl/slicerCaches/slicerCache1.xml", "xl/drawings/drawing1.xml"}) {
    EXPECT_FALSE(HasPart(saved, part)) << part;
  }
  EXPECT_FALSE(Has(ReadPart(saved, "xl/workbook.xml"), "Slicer_"));
  EXPECT_FALSE(Has(ReadPart(saved, "xl/worksheets/sheet1.xml"), "extLst"));
}

}  // namespace
}  // namespace formulon
