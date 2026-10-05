#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "miniz.h"
#include "ooxml_corpus_support.h"

namespace formulon::ooxml_corpus {
namespace {
struct PartFile {
  const char* path;
  std::string_view body;
};

Expected<std::vector<std::uint8_t>, Error> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  if (mz_zip_writer_init_heap(&writer, 0, 4096) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoZipCorrupt, "miniz writer init failed");
  }
  for (const PartFile& p : parts) {
    if (mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                              static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) == MZ_FALSE) {
      mz_zip_writer_end(&writer);
      return make_error(FormulonErrorCode::kIoZipCorrupt, std::string("miniz add failed for ") + p.path);
    }
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size) == MZ_FALSE) {
    mz_zip_writer_end(&writer);
    return make_error(FormulonErrorCode::kIoZipCorrupt, "miniz finalize failed");
  }
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  mz_zip_writer_end(&writer);
  return out;
}

}  // namespace

constexpr std::string_view kPackageRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
    "  <Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
    "Target=\"xl/workbook.xml\"/>\n"
    "</Relationships>\n";

constexpr std::string_view kEmptySheetXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
    "  <sheetData/>\n"
    "</worksheet>\n";

/// Builds a synthetic `.xlsx` package containing a passthrough
/// `xl/theme/theme1.xml` that the reader does not parse but Bundle 2.5
/// promised to round-trip verbatim.
Expected<std::vector<std::uint8_t>, Error> BuildPassthroughThemeArchive() {
  constexpr std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "  <Override PartName=\"/xl/theme/theme1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n"
      "</Types>\n";
  constexpr std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  constexpr std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  // Distinctive payload so the round-trip assertion is meaningful.
  constexpr std::string_view theme_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"CorpusStub\">\n"
      "  <a:themeElements/>\n"
      "</a:theme>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", kPackageRels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", kEmptySheetXml},
      {"xl/theme/theme1.xml", theme_xml},
  });
}

/// Builds a synthetic `.xlsx` package whose Sheet1 references a real
/// `xl/sharedStrings.xml`. Verifies Bundle 2.3 + 2.5 inline-string
/// re-emission preserves the SST text payloads through the writer.
Expected<std::vector<std::uint8_t>, Error> BuildSstArchive() {
  constexpr std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "  <Override PartName=\"/xl/sharedStrings.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>\n"
      "  <Override PartName=\"/xl/styles.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n"
      "</Types>\n";
  constexpr std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  constexpr std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "  <Relationship Id=\"rId2\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" "
      "Target=\"sharedStrings.xml\"/>\n"
      "  <Relationship Id=\"rId3\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
      "Target=\"styles.xml\"/>\n"
      "</Relationships>\n";
  constexpr std::string_view sheet_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData>\n"
      "    <row r=\"1\"><c r=\"A1\" t=\"s\"><v>0</v></c><c r=\"B1\" t=\"s\"><v>1</v></c></row>\n"
      "    <row r=\"2\"><c r=\"A2\" t=\"s\"><v>2</v></c></row>\n"
      "  </sheetData>\n"
      "</worksheet>\n";
  // Three SST entries; A2 reuses index 2 and B1 references index 1.
  constexpr std::string_view sst_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"3\" uniqueCount=\"3\">\n"
      "  <si><t>shared-alpha</t></si>\n"
      "  <si><t>shared-beta</t></si>\n"
      "  <si><t>shared-gamma</t></si>\n"
      "</sst>\n";
  constexpr std::string_view styles_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>\n"
      "</styleSheet>\n";

  return BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", kPackageRels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
      {"xl/sharedStrings.xml", sst_xml},
      {"xl/styles.xml", styles_xml},
  });
}

}  // namespace formulon::ooxml_corpus
