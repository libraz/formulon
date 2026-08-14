//
// Builder for a minimal, hand-assembled OOXML package.
//
// The engine's own writer only ever emits well-formed parts, so a test
// that needs the reader to encounter malformed content -- an unparseable
// `sqref`, an unrecognised workbook content type -- has to assemble the
// archive itself. This header holds that assembly once: the four envelope
// parts every package carries, plus a `[Content_Types].xml` template whose
// workbook content type the caller chooses.
//
// Only the sheet body and the workbook content type vary across callers;
// everything else is fixed so a fixture difference is always the thing the
// test is actually about.

#ifndef FORMULON_TESTS_SUPPORT_OOXML_PACKAGE_FIXTURE_H_
#define FORMULON_TESTS_SUPPORT_OOXML_PACKAGE_FIXTURE_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "miniz.h"

namespace formulon {
namespace test {

/// The content type Excel declares for a plain `.xlsx` workbook part.
inline constexpr std::string_view kXlsxWorkbookContentType =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

/// A `<worksheet>` with an empty `<sheetData>` and nothing else.
inline constexpr std::string_view kEmptySheetXml =
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"><sheetData/></worksheet>";

/// Builds `[Content_Types].xml` declaring the workbook part with
/// `workbook_content_type` and the single worksheet with its canonical one.
inline std::string OoxmlContentTypes(std::string_view workbook_content_type) {
  std::string out =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
      "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
      "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
      "<Override PartName=\"/xl/workbook.xml\" ContentType=\"";
  out.append(workbook_content_type);
  out.append(
      "\"/>"
      "<Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
      "</Types>");
  return out;
}

/// Assembles a one-sheet package around `content_types` and `sheet_xml`.
/// Returns the raw archive bytes, empty on a miniz failure (which is also
/// reported as a non-fatal gtest failure).
inline std::vector<std::uint8_t> BuildOoxmlPackage(std::string_view content_types, std::string_view sheet_xml) {
  constexpr std::string_view kPackageRels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>"
      "</Relationships>";
  constexpr std::string_view kWorkbookXml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
      "</workbook>";
  constexpr std::string_view kWorkbookRels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>"
      "</Relationships>";

  mz_zip_archive writer{};
  if (mz_zip_writer_init_heap(&writer, 0, 4096) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_writer_init_heap failed";
    return {};
  }
  const std::pair<const char*, std::string_view> parts[] = {
      {"[Content_Types].xml", content_types},  {"_rels/.rels", kPackageRels},
      {"xl/workbook.xml", kWorkbookXml},       {"xl/_rels/workbook.xml.rels", kWorkbookRels},
      {"xl/worksheets/sheet1.xml", sheet_xml},
  };
  for (const auto& part : parts) {
    if (mz_zip_writer_add_mem(&writer, part.first, part.second.data(), part.second.size(),
                              static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)) == MZ_FALSE) {
      ADD_FAILURE() << "mz_zip_writer_add_mem failed for " << part.first;
      mz_zip_writer_end(&writer);
      return {};
    }
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size) == MZ_FALSE) {
    ADD_FAILURE() << "mz_zip_writer_finalize_heap_archive failed";
    mz_zip_writer_end(&writer);
    return {};
  }
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  mz_zip_writer_end(&writer);
  return out;
}

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_SUPPORT_OOXML_PACKAGE_FIXTURE_H_
