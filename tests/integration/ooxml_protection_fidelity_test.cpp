//
// Integration test: protection records, extra sheet views, named sheet views,
// legacy notes and header / footer or background pictures that a
// load -> save -> load cycle must keep.

#include <array>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/protection_xml.h"
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
  /// Empty: covered by the `xml` / `rels` / `vml` / `png` defaults.
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

/// One-sheet package. `extra_parts` ride along with their content types;
/// `sheet_rels` are the `<Relationship>` elements of the sheet's rels part.
std::vector<std::uint8_t> BuildPackage(std::string_view sheet_body, std::string_view workbook_extra = {},
                                       std::string_view sheet_rels = {},
                                       const std::vector<PartFile>& extra_parts = {}) {
  std::string content_types =
      std::string(kXmlDecl) +
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Default Extension=\"vml\" ContentType=\"application/vnd.openxmlformats-officedocument.vmlDrawing\"/>"
      "<Default Extension=\"png\" ContentType=\"image/png\"/>"
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
                                   "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>" +
                                   std::string(workbook_extra) + "</workbook>";
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

std::vector<std::uint8_t> Resave(const std::vector<std::uint8_t>& bytes) {
  return SaveOrDie(LoadOrDie(bytes));
}

bool HasPart(const std::vector<std::uint8_t>& bytes, std::string_view name) {
  io::ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes))));
  return zip.has_entry(name);
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
// Workbook protection
// ---------------------------------------------------------------------------

constexpr std::string_view kFullWorkbookProtection =
    "<workbookProtection workbookPassword=\"CF1A\" revisionsPassword=\"8D2B\" lockStructure=\"1\" lockWindows=\"1\" "
    "lockRevision=\"1\" workbookAlgorithmName=\"SHA-512\" workbookHashValue=\"aGFzaA==\" "
    "workbookSaltValue=\"c2FsdA==\" "
    "workbookSpinCount=\"100000\" revisionsAlgorithmName=\"SHA-256\" revisionsHashValue=\"cmV2aGFzaA==\" "
    "revisionsSaltValue=\"cmV2c2FsdA==\" revisionsSpinCount=\"50000\"/>";

void ExpectFullWorkbookProtection(const WorkbookProtection& p) {
  EXPECT_TRUE(p.lock_structure);
  EXPECT_TRUE(p.lock_windows);
  EXPECT_TRUE(p.lock_revision);
  EXPECT_EQ(p.workbook_password, "CF1A");
  EXPECT_EQ(p.workbook_algorithm_name, "SHA-512");
  EXPECT_EQ(p.workbook_hash_value, "aGFzaA==");
  EXPECT_EQ(p.workbook_salt_value, "c2FsdA==");
  EXPECT_EQ(p.workbook_spin_count, 100000U);
  EXPECT_EQ(p.revisions_password, "8D2B");
  EXPECT_EQ(p.revisions_algorithm_name, "SHA-256");
  EXPECT_EQ(p.revisions_hash_value, "cmV2aGFzaA==");
  EXPECT_EQ(p.revisions_salt_value, "cmV2c2FsdA==");
  EXPECT_EQ(p.revisions_spin_count, 50000U);
}

TEST(WorkbookProtectionFidelity, EveryAttributeSurvivesAResave) {
  const auto source = BuildPackage("<sheetData/>", kFullWorkbookProtection);
  const Workbook first = LoadOrDie(source);
  ExpectFullWorkbookProtection(io::parse_workbook_protection(first.workbook_protection_xml()));

  const auto saved = SaveOrDie(first);
  EXPECT_TRUE(Has(ReadPart(saved, "xl/workbook.xml"), "workbookSpinCount=\"100000\""));
  const Workbook second = LoadOrDie(saved);
  ExpectFullWorkbookProtection(io::parse_workbook_protection(second.workbook_protection_xml()));
}

TEST(WorkbookProtectionFidelity, SetterKeepsUnmodelledAttributesAndClearDropsTheElement) {
  Workbook wb = LoadOrDie(
      BuildPackage("<sheetData/>", "<workbookProtection workbookPasswordCharacterSet=\"UTF-8\" lockStructure=\"1\"/>"));
  WorkbookProtection p = io::parse_workbook_protection(wb.workbook_protection_xml());
  p.lock_structure = false;
  p.lock_windows = true;
  p.workbook_algorithm_name = "SHA-512";
  p.workbook_spin_count = 7;
  wb.set_workbook_protection_xml(io::apply_workbook_protection(wb.workbook_protection_xml(), p));
  const Workbook back = LoadOrDie(SaveOrDie(wb));
  const std::string& xml = back.workbook_protection_xml();
  EXPECT_TRUE(Has(xml, "workbookPasswordCharacterSet=\"UTF-8\"")) << xml;
  EXPECT_TRUE(Has(xml, "lockWindows=\"1\"")) << xml;
  EXPECT_FALSE(Has(xml, "lockStructure")) << xml;

  EXPECT_TRUE(io::apply_workbook_protection("<workbookProtection lockStructure=\"1\"/>", WorkbookProtection{}).empty());
}

// ---------------------------------------------------------------------------
// Protected ranges
// ---------------------------------------------------------------------------

constexpr std::string_view kProtectedRanges =
    "<sheetData/><sheetProtection sheet=\"1\"/>"
    "<protectedRanges>"
    "<protectedRange password=\"CC1A\" sqref=\"B2:C3 E5\" name=\"Inputs\"/>"
    "<protectedRange sqref=\"A10:A12\" name=\"Totals\" securityDescriptor=\"O:WDG:WDD:(A;;CC;;;WD)\" "
    "algorithmName=\"SHA-512\" hashValue=\"aA==\" saltValue=\"cw==\" spinCount=\"100000\"/>"
    "</protectedRanges>";

TEST(ProtectedRangesFidelity, EveryAttributeSurvivesAResave) {
  const Workbook first = LoadOrDie(BuildPackage(kProtectedRanges));
  const Workbook second = LoadOrDie(SaveOrDie(first));
  for (const Workbook* wb : {&first, &second}) {
    const auto ranges = io::read_protected_ranges(wb->sheet(0).raw_extensions());
    ASSERT_EQ(ranges.size(), 2U);
    EXPECT_EQ(ranges[0].name, "Inputs");
    EXPECT_EQ(ranges[0].sqref, "B2:C3 E5");
    EXPECT_EQ(ranges[0].password, "CC1A");
    EXPECT_EQ(ranges[1].name, "Totals");
    EXPECT_EQ(ranges[1].security_descriptor, "O:WDG:WDD:(A;;CC;;;WD)");
    EXPECT_EQ(ranges[1].algorithm_name, "SHA-512");
    EXPECT_EQ(ranges[1].hash_value, "aA==");
    EXPECT_EQ(ranges[1].salt_value, "cw==");
    EXPECT_EQ(ranges[1].spin_count, 100000U);
  }
}

TEST(ProtectedRangesFidelity, SqrefShiftsWithRowAndColumnEdits) {
  Workbook wb = LoadOrDie(BuildPackage(kProtectedRanges));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  ASSERT_TRUE(static_cast<bool>(wb.insert_cols(0, 0, 1)));
  const auto ranges = io::read_protected_ranges(LoadOrDie(SaveOrDie(wb)).sheet(0).raw_extensions());
  ASSERT_EQ(ranges.size(), 2U);
  EXPECT_EQ(ranges[0].sqref, "C4:D5 F7");
  EXPECT_EQ(ranges[1].sqref, "B12:B14");
}

TEST(ProtectedRangesFidelity, WriteReplacesAndEmptyListRemoves) {
  Workbook wb = LoadOrDie(BuildPackage(kProtectedRanges));
  WorksheetRawExtensions& raw = wb.sheet(0).mutable_raw_extensions();
  auto ranges = io::read_protected_ranges(raw);
  ranges.pop_back();
  ranges[0].name = "A&B";
  io::write_protected_ranges(raw, ranges);
  auto back = io::read_protected_ranges(LoadOrDie(SaveOrDie(wb)).sheet(0).raw_extensions());
  ASSERT_EQ(back.size(), 1U);
  EXPECT_EQ(back[0].name, "A&B");
  io::write_protected_ranges(raw, {});
  EXPECT_FALSE(Has(ReadPart(SaveOrDie(wb), "xl/worksheets/sheet1.xml"), "protectedRange"));
}

TEST(ProtectedRangesFidelity, ChildElementsSurviveAnEdit) {
  Workbook wb = LoadOrDie(BuildPackage(
      "<sheetData/><sheetProtection sheet=\"1\"/><protectedRanges>"
      "<protectedRange sqref=\"A1\" name=\"Inputs\"><securityDescriptor>O:WD</securityDescriptor></protectedRange>"
      "<protectedRange sqref=\"A3\" name=\"Totals\"/></protectedRanges>"));
  WorksheetRawExtensions& raw = wb.sheet(0).mutable_raw_extensions();
  auto ranges = io::read_protected_ranges(raw);
  ASSERT_EQ(ranges.size(), 2U);
  ranges[0].sqref = "A2";
  io::write_protected_ranges(raw, ranges);
  const std::string xml = ReadPart(SaveOrDie(wb), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "<securityDescriptor>O:WD</securityDescriptor>")) << xml;
  EXPECT_TRUE(Has(xml, "sqref=\"A2\"")) << xml;
}

// ---------------------------------------------------------------------------
// Sheet views
// ---------------------------------------------------------------------------

TEST(SheetViewsFidelity, ViewsAfterTheFirstSurviveAfterIt) {
  const std::string body =
      "<sheetViews><sheetView tabSelected=\"1\" workbookViewId=\"0\"><selection activeCell=\"B2\" sqref=\"B2\"/>"
      "</sheetView><sheetView showGridLines=\"0\" zoomScale=\"150\" workbookViewId=\"1\"><selection activeCell=\"C3\" "
      "sqref=\"C3\"/></sheetView></sheetViews><sheetData/>";
  const auto saved = Resave(BuildPackage(body));
  const std::string xml = ReadPart(saved, "xl/worksheets/sheet1.xml");
  const std::size_t first = xml.find("workbookViewId=\"0\"");
  const std::size_t second = xml.find("workbookViewId=\"1\"");
  ASSERT_NE(first, std::string::npos) << xml;
  ASSERT_NE(second, std::string::npos) << xml;
  EXPECT_LT(first, second);
  EXPECT_TRUE(Has(xml, "zoomScale=\"150\"")) << xml;
  EXPECT_TRUE(Has(xml, "activeCell=\"C3\"")) << xml;

  const Workbook again = LoadOrDie(saved);
  EXPECT_TRUE(again.sheet(0).view().tab_selected);
  EXPECT_TRUE(Has(again.sheet(0).view().extra_views_xml, "workbookViewId=\"1\""));
}

TEST(SheetViewsFidelity, ExtraViewKeepsAnOtherwiseDefaultFirstView) {
  const std::string body =
      "<sheetViews><sheetView workbookViewId=\"0\"/><sheetView workbookViewId=\"1\" zoomScale=\"80\"/></sheetViews>"
      "<sheetData/>";
  const std::string xml = ReadPart(Resave(BuildPackage(body)), "xl/worksheets/sheet1.xml");
  EXPECT_TRUE(Has(xml, "workbookViewId=\"1\"")) << xml;
  EXPECT_TRUE(Has(xml, "zoomScale=\"80\"")) << xml;
}

// ---------------------------------------------------------------------------
// Named sheet views
// ---------------------------------------------------------------------------

constexpr std::string_view kNamedSheetViewType =
    "http://schemas.microsoft.com/office/2019/04/relationships/namedSheetView";
constexpr std::string_view kNamedSheetViewCt = "application/vnd.ms-excel.namedsheetviews+xml";

TEST(NamedSheetViewsFidelity, PartAndRelationshipSurviveAResave) {
  const std::string part =
      "<namedSheetViews xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews\" "
      "xmlns:x=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<namedSheetView name=\"Mine\" id=\"{8F2D3B6A-1111-4222-8333-944455556666}\">"
      "<nsvFilter filterId=\"{AAAAAAAA-1111-4222-8333-944455556666}\" ref=\"A1:B3\" tableId=\"0\">"
      "<columnFilter colId=\"0\"><filter colId=\"0\"><x:filters><x:filter val=\"1\"/></x:filters></filter>"
      "</columnFilter></nsvFilter></namedSheetView></namedSheetViews>";
  const std::string rels = "<Relationship Id=\"rId7\" Type=\"" + std::string(kNamedSheetViewType) +
                           "\" Target=\"../namedSheetViews/namedSheetView1.xml\"/>";
  const auto source = BuildPackage("<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>", {}, rels,
                                   {{"xl/namedSheetViews/namedSheetView1.xml", part, std::string(kNamedSheetViewCt)}});
  const auto once = Resave(source);
  const auto twice = Resave(once);
  for (const auto* saved : {&once, &twice}) {
    ASSERT_TRUE(HasPart(*saved, "xl/namedSheetViews/namedSheetView1.xml"));
    EXPECT_EQ(ReadPart(*saved, "xl/namedSheetViews/namedSheetView1.xml"), part);
    EXPECT_TRUE(Has(ReadPart(*saved, "xl/worksheets/_rels/sheet1.xml.rels"), kNamedSheetViewType));
    EXPECT_TRUE(Has(ReadPart(*saved, "[Content_Types].xml"), kNamedSheetViewCt));
  }
}

TEST(NamedSheetViewsFidelity, PartIsKeptVerbatimAcrossStructuralEdits) {
  const std::string part =
      "<namedSheetViews xmlns=\"http://schemas.microsoft.com/office/spreadsheetml/2019/namedsheetviews\">"
      "<namedSheetView name=\"Mine\" id=\"{8F2D3B6A-1111-4222-8333-944455556666}\">"
      "<nsvFilter filterId=\"{AAAAAAAA-1111-4222-8333-944455556666}\" ref=\"A1:B3\" tableId=\"0\"/>"
      "</namedSheetView></namedSheetViews>";
  const std::string rels = "<Relationship Id=\"rId7\" Type=\"" + std::string(kNamedSheetViewType) +
                           "\" Target=\"../namedSheetViews/namedSheetView1.xml\"/>";
  Workbook wb = LoadOrDie(BuildPackage(
      "<sheetData/>", {}, rels, {{"xl/namedSheetViews/namedSheetView1.xml", part, std::string(kNamedSheetViewCt)}}));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 2)));
  ASSERT_TRUE(static_cast<bool>(wb.delete_cols(0, 0, 1)));
  const auto saved = SaveOrDie(wb);
  EXPECT_EQ(ReadPart(saved, "xl/namedSheetViews/namedSheetView1.xml"), part);
  EXPECT_TRUE(Has(ReadPart(saved, "xl/worksheets/_rels/sheet1.xml.rels"), kNamedSheetViewType));
}

// ---------------------------------------------------------------------------
// Legacy notes
// ---------------------------------------------------------------------------

constexpr std::string_view kCommentsType =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
constexpr std::string_view kVmlType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

std::string NoteShape(int id, int row, int col, bool visible, std::string_view anchor) {
  return "<v:shape id=\"_x0000_s" + std::to_string(id) +
         "\" type=\"#_x0000_t202\" style='position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:180pt;"
         "height:90pt;z-index:" +
         std::to_string(id - 1024) + ";visibility:" + (visible ? "visible" : "hidden") +
         "' fillcolor=\"#ffffe1\" o:insetmode=\"auto\"><v:fill color2=\"#ffffe1\"/><x:ClientData ObjectType=\"Note\">"
         "<x:MoveWithCells/><x:SizeWithCells/><x:Anchor>" +
         std::string(anchor) +
         "</x:Anchor><x:AutoFill>False</x:AutoFill>"
         "<x:Row>" +
         std::to_string(row) + "</x:Row><x:Column>" + std::to_string(col) + "</x:Column>" +
         (visible ? "<x:Visible/>" : "") + "</x:ClientData></v:shape>";
}

std::vector<std::uint8_t> NotesPackage() {
  const std::string comments =
      "<comments xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<authors><author>Alice</author><author>Bob</author><author>Carol</author></authors><commentList>"
      "<comment ref=\"B2\" authorId=\"1\"><text><r><t>shown</t></r></text></comment>"
      "<comment ref=\"D4\" authorId=\"1\"><text><r><t>hidden</t></r></text></comment>"
      "</commentList></comments>";
  const std::string vml =
      "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" "
      "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"><o:shapelayout v:ext=\"edit\"><o:idmap v:ext=\"edit\" "
      "data=\"1\"/></o:shapelayout><v:shapetype id=\"_x0000_t202\" coordsize=\"21600,21600\" o:spt=\"202\"/>" +
      NoteShape(1025, 1, 1, true, "3, 15, 0, 2, 6, 40, 8, 4") +
      NoteShape(1026, 3, 3, false, "5, 15, 2, 10, 8, 20, 9, 4") + "</xml>";
  const std::string rels = "<Relationship Id=\"rId1\" Type=\"" + std::string(kCommentsType) +
                           "\" Target=\"../comments1.xml\"/><Relationship Id=\"rId2\" Type=\"" + std::string(kVmlType) +
                           "\" Target=\"../drawings/vmlDrawing1.vml\"/>";
  return BuildPackage(
      "<sheetData/><legacyDrawing r:id=\"rId2\"/>", {}, rels,
      {{"xl/comments1.xml", comments, "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"},
       {"xl/drawings/vmlDrawing1.vml", vml, {}}});
}

TEST(LegacyNotesFidelity, GeometryVisibilityAndAuthorsAreReadAndKept) {
  const auto source = NotesPackage();
  const Workbook first = LoadOrDie(source);
  const auto& notes = first.sheet(0).comments();
  ASSERT_EQ(notes.size(), 2U);
  EXPECT_TRUE(notes[0].visible);
  EXPECT_FALSE(notes[1].visible);
  ASSERT_TRUE(notes[0].has_box);
  // Relative to the note cell B2 (row 1, col 1).
  EXPECT_EQ(notes[0].box, (std::array<std::int32_t, 8>{2, 15, -1, 2, 5, 40, 7, 4}));
  ASSERT_EQ(first.sheet(0).comment_authors().size(), 3U);

  const auto saved = SaveOrDie(first);
  EXPECT_TRUE(Has(ReadPart(saved, "xl/drawings/vmlDrawing1.vml"), "3, 15, 0, 2, 6, 40, 8, 4"));
  const std::string comments = ReadPart(saved, "xl/comments1.xml");
  EXPECT_TRUE(Has(comments, "<author>Carol</author>")) << comments;
}

TEST(LegacyNotesFidelity, EditingTheNoteSetKeepsEveryExistingBoxAndFlag) {
  Workbook wb = LoadOrDie(NotesPackage());
  wb.sheet(0).mutable_comments().emplace_back(8, 8, "Bob", "new");
  const auto saved = SaveOrDie(wb);
  const std::string vml = ReadPart(saved, "xl/drawings/vmlDrawing1.vml");
  EXPECT_TRUE(Has(vml, "3, 15, 0, 2, 6, 40, 8, 4")) << vml;
  EXPECT_TRUE(Has(vml, "5, 15, 2, 10, 8, 20, 9, 4")) << vml;
  EXPECT_TRUE(Has(vml, "visibility:visible")) << vml;
  // The authors table keeps Carol, who owns no note, and Bob keeps his index.
  const std::string comments = ReadPart(saved, "xl/comments1.xml");
  EXPECT_TRUE(Has(comments, "<author>Carol</author>")) << comments;
  EXPECT_TRUE(Has(comments, "authorId=\"1\"")) << comments;

  const Workbook again = LoadOrDie(saved);
  ASSERT_EQ(again.sheet(0).comments().size(), 3U);
  EXPECT_TRUE(again.sheet(0).comments()[0].visible);
  EXPECT_FALSE(again.sheet(0).comments()[1].visible);
  EXPECT_EQ(again.sheet(0).comments()[1].box, (std::array<std::int32_t, 8>{2, 15, -1, 10, 5, 20, 6, 4}));
}

TEST(LegacyNotesFidelity, ChangingVisibilityRegeneratesTheDrawing) {
  Workbook wb = LoadOrDie(NotesPackage());
  wb.sheet(0).mutable_comments()[1].visible = true;
  wb.sheet(0).mutable_comments()[0].visible = false;
  const Workbook again = LoadOrDie(SaveOrDie(wb));
  EXPECT_FALSE(again.sheet(0).comments()[0].visible);
  EXPECT_TRUE(again.sheet(0).comments()[1].visible);
}

TEST(LegacyNotesFidelity, BoxFollowsTheNoteThroughRowInsertion) {
  Workbook wb = LoadOrDie(NotesPackage());
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 0, 3)));
  const std::string vml = ReadPart(SaveOrDie(wb), "xl/drawings/vmlDrawing1.vml");
  // B2 moved to B5: the box moved three rows down with it.
  EXPECT_TRUE(Has(vml, "3, 15, 3, 2, 6, 40, 11, 4")) << vml;
}

// ---------------------------------------------------------------------------
// Header / footer and background pictures
// ---------------------------------------------------------------------------

constexpr std::string_view kImageType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";

TEST(PictureFidelity, HeaderFooterImageAndBackgroundPictureSurviveAResave) {
  const std::string body =
      "<sheetData/><headerFooter><oddHeader>&amp;L&amp;G</oddHeader></headerFooter>"
      "<legacyDrawingHF r:id=\"rId1\"/><picture r:id=\"rId2\"/>";
  const std::string rels = "<Relationship Id=\"rId1\" Type=\"" + std::string(kVmlType) +
                           "\" Target=\"../drawings/vmlDrawing1.vml\"/><Relationship Id=\"rId2\" Type=\"" +
                           std::string(kImageType) + "\" Target=\"../media/image2.png\"/>";
  const std::string vml =
      "<xml xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:o=\"urn:schemas-microsoft-com:office:office\">"
      "<v:shape id=\"LH\" o:spid=\"_x0000_s2049\"><v:imagedata o:relid=\"rId1\" o:title=\"logo\"/></v:shape></xml>";
  const std::string vml_rels =
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" "
      "Type=\"" +
      std::string(kImageType) + "\" Target=\"../media/image1.png\"/></Relationships>";
  const std::string png1 = "\x89PNG-header-image";
  const std::string png2 = "\x89PNG-background";
  const auto source = BuildPackage(body, {}, rels,
                                   {{"xl/drawings/vmlDrawing1.vml", vml, {}},
                                    {"xl/drawings/_rels/vmlDrawing1.vml.rels", vml_rels, {}},
                                    {"xl/media/image1.png", png1, {}},
                                    {"xl/media/image2.png", png2, {}}});
  const auto once = Resave(source);
  const auto twice = Resave(once);
  for (const auto* saved : {&once, &twice}) {
    const std::string sheet = ReadPart(*saved, "xl/worksheets/sheet1.xml");
    EXPECT_TRUE(Has(sheet, "&amp;L&amp;G")) << sheet;
    EXPECT_TRUE(Has(sheet, "<legacyDrawingHF")) << sheet;
    EXPECT_TRUE(Has(sheet, "<picture")) << sheet;
    EXPECT_EQ(ReadPart(*saved, "xl/drawings/vmlDrawing1.vml"), vml);
    EXPECT_EQ(ReadPart(*saved, "xl/drawings/_rels/vmlDrawing1.vml.rels"), vml_rels);
    EXPECT_EQ(ReadPart(*saved, "xl/media/image1.png"), png1);
    EXPECT_EQ(ReadPart(*saved, "xl/media/image2.png"), png2);
    // Every r:id in the sheet body resolves in the sheet's rels part.
    const std::string saved_rels = ReadPart(*saved, "xl/worksheets/_rels/sheet1.xml.rels");
    for (const char* attr : {"<legacyDrawingHF r:id=\"", "<picture r:id=\""}) {
      const std::size_t at = sheet.find(attr);
      ASSERT_NE(at, std::string::npos) << attr;
      const std::size_t from = at + std::string_view(attr).size();
      const std::string rid = sheet.substr(from, sheet.find('"', from) - from);
      EXPECT_TRUE(Has(saved_rels, "Id=\"" + rid + "\"")) << rid << " in " << saved_rels;
    }
  }
}

}  // namespace
}  // namespace formulon
