// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML (.xlsx) package writer — empty-workbook slice. Emits the minimum
// set of parts required for Excel 365 to open the file as an empty
// workbook. The part contents are hard-coded string templates with
// targeted substitutions for sheet names and per-sheet sequence numbers; a
// full builder infrastructure will be added once cells, styles, shared
// strings and defined names arrive.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/04-xlsx-io.md §4.3 (part classification)

#include "io/ooxml_writer.h"

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "miniz.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// XML escaping
// ---------------------------------------------------------------------------

/// Appends `in` to `out` with the five XML-critical characters escaped:
/// ampersand, less-than, greater-than, double-quote, apostrophe. Non-ASCII
/// bytes (UTF-8 continuation bytes etc.) are emitted verbatim so Japanese
/// sheet names round-trip unchanged.
void AppendXmlEscaped(std::string& out, std::string_view in) {
  for (char raw : in) {
    switch (raw) {
      case '&':
        out.append("&amp;");
        break;
      case '<':
        out.append("&lt;");
        break;
      case '>':
        out.append("&gt;");
        break;
      case '"':
        out.append("&quot;");
        break;
      case '\'':
        out.append("&apos;");
        break;
      default:
        out.push_back(raw);
        break;
    }
  }
}

// ---------------------------------------------------------------------------
// Part builders
// ---------------------------------------------------------------------------

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

std::string BuildContentTypes(std::size_t sheet_count) {
  std::string out;
  out.reserve(512 + sheet_count * 128);
  out.append(kXmlDecl);
  out.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n");
  out.append(
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n");
  out.append("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n");
  out.append(
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n");
  for (std::size_t i = 0; i < sheet_count; ++i) {
    out.append("  <Override PartName=\"/xl/worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(
        ".xml\" "
        "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n");
  }
  out.append(
      "  <Override PartName=\"/xl/styles.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>\n");
  out.append("</Types>\n");
  return out;
}

std::string BuildPackageRels() {
  std::string out;
  out.reserve(256);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append(
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n");
  out.append("</Relationships>\n");
  return out;
}

std::string BuildWorkbookXml(const Workbook& wb) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 96);
  out.append(kXmlDecl);
  out.append(
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  out.append("  <sheets>\n");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out.append("    <sheet name=\"");
    AppendXmlEscaped(out, wb.sheet(i).name());
    out.append("\" sheetId=\"");
    out.append(std::to_string(i + 1));
    out.append("\" r:id=\"rId");
    out.append(std::to_string(i + 1));
    out.append("\"/>\n");
  }
  out.append("  </sheets>\n");
  out.append("</workbook>\n");
  return out;
}

std::string BuildWorkbookRels(std::size_t sheet_count) {
  std::string out;
  out.reserve(256 + sheet_count * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  for (std::size_t i = 0; i < sheet_count; ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(i + 1));
    out.append(
        "\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
        "Target=\"worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".xml\"/>\n");
  }
  // Styles relationship follows the worksheet relationships.
  out.append("  <Relationship Id=\"rId");
  out.append(std::to_string(sheet_count + 1));
  out.append(
      "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
      "Target=\"styles.xml\"/>\n");
  out.append("</Relationships>\n");
  return out;
}

std::string BuildWorksheetXml() {
  std::string out;
  out.reserve(192);
  out.append(kXmlDecl);
  out.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  out.append("  <sheetData/>\n");
  out.append("</worksheet>\n");
  return out;
}

std::string BuildStylesXml() {
  std::string out;
  out.reserve(512);
  out.append(kXmlDecl);
  out.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  out.append("  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\n");
  out.append("  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n");
  out.append("  <borders count=\"1\"><border/></borders>\n");
  out.append(
      "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n");
  out.append(
      "  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\n");
  out.append("</styleSheet>\n");
  return out;
}

// ---------------------------------------------------------------------------
// miniz helpers
// ---------------------------------------------------------------------------

/// RAII guard around an initialised `mz_zip_archive` writer. The destructor
/// releases any heap buffer retained by miniz when the writer is abandoned
/// mid-flight (e.g. an `mz_zip_writer_add_mem` call failed and we early-
/// returned an error).
class ZipWriterGuard {
 public:
  ZipWriterGuard() = default;
  ZipWriterGuard(const ZipWriterGuard&) = delete;
  ZipWriterGuard& operator=(const ZipWriterGuard&) = delete;
  ZipWriterGuard(ZipWriterGuard&&) = delete;
  ZipWriterGuard& operator=(ZipWriterGuard&&) = delete;

  ~ZipWriterGuard() {
    if (active_) {
      // Best-effort cleanup; we're already on an error path.
      mz_zip_writer_end(&archive_);
    }
  }

  bool init() {
    if (mz_zip_writer_init_heap(&archive_, /*size_to_reserve_at_beginning=*/0,
                                /*initial_allocation_size=*/8 * 1024) == MZ_FALSE) {
      return false;
    }
    active_ = true;
    return true;
  }

  mz_zip_archive* get() noexcept { return &archive_; }

  /// Releases ownership of the underlying archive to the caller. Subsequent
  /// destruction no longer touches miniz state.
  void release() noexcept { active_ = false; }

 private:
  mz_zip_archive archive_{};
  bool active_ = false;
};

/// Adds a single part to the archive. Returns an `Error` tagged with the part
/// path when miniz refuses the write.
Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body) {
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed", std::move(context));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> write_ooxml(const Workbook& wb) {
  const std::size_t sheet_count = wb.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "workbook has zero sheets", "context=write_ooxml");
  }

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_ooxml");
  }

  // 1. [Content_Types].xml
  {
    auto result = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(sheet_count));
    if (!result) {
      return result.error();
    }
  }

  // 2. _rels/.rels
  {
    auto result = AddPart(writer.get(), "_rels/.rels", BuildPackageRels());
    if (!result) {
      return result.error();
    }
  }

  // 3. xl/workbook.xml
  {
    auto result = AddPart(writer.get(), "xl/workbook.xml", BuildWorkbookXml(wb));
    if (!result) {
      return result.error();
    }
  }

  // 4. xl/_rels/workbook.xml.rels
  {
    auto result = AddPart(writer.get(), "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheet_count));
    if (!result) {
      return result.error();
    }
  }

  // 5. xl/worksheets/sheet<N>.xml
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::string part_path("xl/worksheets/sheet");
    part_path.append(std::to_string(i + 1));
    part_path.append(".xml");
    auto result = AddPart(writer.get(), part_path, BuildWorksheetXml());
    if (!result) {
      return result.error();
    }
  }

  // 6. xl/styles.xml
  {
    auto result = AddPart(writer.get(), "xl/styles.xml", BuildStylesXml());
    if (!result) {
      return result.error();
    }
  }

  // Finalise into a heap buffer, then copy into a std::vector so the caller
  // owns the bytes through normal RAII.
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(writer.get(), &archive_ptr, &archive_size) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_finalize_heap_archive failed",
                      "context=write_ooxml");
  }
  if (mz_zip_writer_end(writer.get()) == MZ_FALSE) {
    // finalize succeeded but end failed — still free the buffer miniz handed
    // us before surfacing the error.
    if (archive_ptr != nullptr) {
      mz_free(archive_ptr);
    }
    writer.release();
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_end failed", "context=write_ooxml");
  }
  writer.release();

  std::vector<std::uint8_t> bytes;
  bytes.resize(archive_size);
  if (archive_size > 0 && archive_ptr != nullptr) {
    std::memcpy(bytes.data(), archive_ptr, archive_size);
  }
  if (archive_ptr != nullptr) {
    mz_free(archive_ptr);
  }
  return bytes;
}

}  // namespace io
}  // namespace formulon
