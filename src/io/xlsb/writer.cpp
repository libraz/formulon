// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the MS-XLSB package writer. See `io/xlsb/writer.h`
// for the contract. The implementation mirrors the structure of
// `io/ooxml_writer.cpp`: build an emission plan (which sheet owns
// which numeric id, which passthrough parts survive collision
// detection), then compose the parts and pipe them through miniz.

#include "io/xlsb/writer.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/passthrough_part.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/sst_writer.h"
#include "io/xml_escape.h"
#include "miniz.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

// The reader accepts both `application/vnd.ms-excel.sheet.binary.macroEnabled.main`
// (used by `.xlsm` and the `.xlsb` corpus xlwings emits on macOS) and
// `application/vnd.ms-excel.sheet.macroEnabled.main` (the alternative
// some non-macro xlsb writers ship). We emit the first form because
// (a) it's what the reader's primary fixture uses and (b) it's the
// content type Excel for Mac actually writes — the other variant is
// accepted for compatibility on input only.
constexpr std::string_view kCtPackageRels = "application/vnd.openxmlformats-package.relationships+xml";
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
constexpr std::string_view kCtWorksheetXlsb = "application/vnd.ms-excel.binIndexWs";
constexpr std::string_view kCtSharedStringsXlsb = "application/vnd.ms-excel.sharedStrings";

constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";

// ---------------------------------------------------------------------------
// Emission plan: where do passthrough parts land, do any collide?
// ---------------------------------------------------------------------------

struct EmissionPlan {
  std::vector<const PassthroughPart*> passthrough_kept;
  bool has_text_cells = false;  // gates emission of xl/sharedStrings.bin
};

std::unordered_set<std::string> BuildGeneratedPathSet(const Workbook& wb, bool emit_sst_part) {
  std::unordered_set<std::string> paths;
  paths.insert("[Content_Types].xml");
  paths.insert("_rels/.rels");
  paths.insert("xl/workbook.bin");
  paths.insert("xl/_rels/workbook.bin.rels");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    paths.insert("xl/worksheets/sheet" + std::to_string(i + 1) + ".bin");
  }
  if (emit_sst_part) {
    paths.insert("xl/sharedStrings.bin");
  }
  return paths;
}

EmissionPlan BuildEmissionPlan(const Workbook& wb, bool sst_present) {
  EmissionPlan plan;
  plan.has_text_cells = sst_present;

  const std::unordered_set<std::string> generated = BuildGeneratedPathSet(wb, sst_present);
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (generated.count(part.path) != 0U) {
      StructuredLog("xlsb.writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("generated_path_wins"))
          .warn();
      continue;
    }
    plan.passthrough_kept.push_back(&part);
  }
  return plan;
}

// ---------------------------------------------------------------------------
// XML part builders: the package envelope is XML even in xlsb.
// ---------------------------------------------------------------------------

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 128 + plan.passthrough_kept.size() * 128);
  out.append(kXmlDecl);
  out.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n");
  out.append("  <Default Extension=\"rels\" ContentType=\"");
  out.append(kCtPackageRels);
  out.append("\"/>\n");
  out.append("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n");
  // Default for `bin` is the workbook content type — matches what the
  // reader test fixture uses and lets passthrough binary parts (e.g.
  // `xl/theme/theme1.xml` is XML, but binary parts like images would
  // ride this Default in a real xlsb).
  out.append("  <Default Extension=\"bin\" ContentType=\"");
  out.append(kCtWorkbookXlsb);
  out.append("\"/>\n");
  out.append("  <Override PartName=\"/xl/workbook.bin\" ContentType=\"");
  out.append(kCtWorkbookXlsb);
  out.append("\"/>\n");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out.append("  <Override PartName=\"/xl/worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".bin\" ContentType=\"");
    out.append(kCtWorksheetXlsb);
    out.append("\"/>\n");
  }
  if (plan.has_text_cells) {
    out.append("  <Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"");
    out.append(kCtSharedStringsXlsb);
    out.append("\"/>\n");
  }
  // Passthrough overrides: only for entries that carried an explicit
  // ContentType in the source archive. Default-typed parts (empty
  // content_type) must NOT appear as Overrides.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part->content_type.empty()) {
      continue;
    }
    out.append("  <Override PartName=\"/");
    AppendXmlEscaped(out, part->path);
    out.append("\" ContentType=\"");
    AppendXmlEscaped(out, part->content_type);
    out.append("\"/>\n");
  }
  out.append("</Types>\n");
  return out;
}

std::string BuildPackageRels() {
  std::string out;
  out.reserve(256);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append("  <Relationship Id=\"rId1\" Type=\"");
  out.append(kRelOfficeDocument);
  out.append("\" Target=\"xl/workbook.bin\"/>\n");
  out.append("</Relationships>\n");
  return out;
}

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst) {
  std::string out;
  out.reserve(256 + sheet_count * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  for (std::size_t i = 0; i < sheet_count; ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(i + 1));
    out.append("\" Type=\"");
    out.append(kRelWorksheet);
    out.append("\" Target=\"worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".bin\"/>\n");
  }
  if (emit_sst) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(sheet_count + 1));
    out.append("\" Type=\"");
    out.append(kRelSharedStrings);
    out.append("\" Target=\"sharedStrings.bin\"/>\n");
  }
  out.append("</Relationships>\n");
  return out;
}

// ---------------------------------------------------------------------------
// Workbook stream (xl/workbook.bin)
// ---------------------------------------------------------------------------

std::vector<std::uint8_t> BuildWorkbookBin(const Workbook& wb) {
  std::vector<std::uint8_t> body;
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBook), ByteSpan{});
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBundleShs), ByteSpan{});
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    // BrtBundleSh ([MS-XLSB] §2.4.304):
    //   hsState    : u32 (0 = visible)
    //   iTabID     : u32 (sheet id; 1-based)
    //   strRelID   : XLNullableWideString
    //   strName    : XLWideString
    std::vector<std::uint8_t> p;
    emit_u32(p, 0U);                                  // hsState (visible)
    emit_u32(p, static_cast<std::uint32_t>(i + 1U));  // iTabID
    const std::string rid = std::string("rId") + std::to_string(i + 1U);
    emit_xlnullablewidestring(p, std::optional<std::string_view>{rid});
    emit_xlwidestring(p, wb.sheet(i).name());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh), p);
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBundleShs), ByteSpan{});

  // Defined names + tables are deferred. Surface a single warning so
  // callers can see what's being dropped without a per-name spam.
  if (!wb.defined_names().empty()) {
    StructuredLog("xlsb.writer.deferred")
        .field("kind", std::string_view("defined_names"))
        .field("count", static_cast<std::int64_t>(wb.defined_names().size()))
        .info();
  }
  if (!wb.tables().empty()) {
    StructuredLog("xlsb.writer.deferred")
        .field("kind", std::string_view("tables"))
        .field("count", static_cast<std::int64_t>(wb.tables().size()))
        .info();
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBook), ByteSpan{});
  return body;
}

// ---------------------------------------------------------------------------
// miniz helpers (mirrors `ooxml_writer.cpp`)
// ---------------------------------------------------------------------------

class ZipWriterGuard {
 public:
  ZipWriterGuard() = default;
  ZipWriterGuard(const ZipWriterGuard&) = delete;
  ZipWriterGuard& operator=(const ZipWriterGuard&) = delete;
  ZipWriterGuard(ZipWriterGuard&&) = delete;
  ZipWriterGuard& operator=(ZipWriterGuard&&) = delete;

  ~ZipWriterGuard() {
    if (active_) {
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

  void release() noexcept { active_ = false; }

 private:
  mz_zip_archive archive_{};
  bool active_ = false;
};

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

Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body) {
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed (binary)",
                      std::move(context));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> write_xlsb(const Workbook& workbook) {
  const std::size_t sheet_count = workbook.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook has zero sheets", "context=write_xlsb");
  }

  // Pre-pass: emit each sheet body so we know whether the SST will be
  // non-empty. We hold the resulting bytes until after we write the
  // envelope so the order of `mz_zip_writer_add_mem` calls matches
  // what the reader expects (it does not, but we keep symmetry with
  // `write_ooxml`).
  SstBuilder sst;
  std::vector<std::vector<std::uint8_t>> sheet_bodies;
  sheet_bodies.reserve(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    auto sheet_body_or = emit_sheet(workbook.sheet(i), sst);
    if (!sheet_body_or) {
      return sheet_body_or.error();
    }
    sheet_bodies.push_back(std::move(sheet_body_or.value()));
  }
  const bool emit_sst_part = !sst.empty();

  const EmissionPlan plan = BuildEmissionPlan(workbook, emit_sst_part);

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_xlsb");
  }

  // 1. [Content_Types].xml
  if (auto r = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(workbook, plan)); !r) {
    return r.error();
  }
  // 2. _rels/.rels
  if (auto r = AddPart(writer.get(), "_rels/.rels", BuildPackageRels()); !r) {
    return r.error();
  }
  // 3. xl/_rels/workbook.bin.rels
  if (auto r = AddPart(writer.get(), "xl/_rels/workbook.bin.rels", BuildWorkbookRels(sheet_count, emit_sst_part)); !r) {
    return r.error();
  }
  // 4. xl/workbook.bin
  {
    const std::vector<std::uint8_t> wb_bytes = BuildWorkbookBin(workbook);
    if (auto r = AddPartBytes(writer.get(), "xl/workbook.bin", wb_bytes); !r) {
      return r.error();
    }
  }
  // 5. xl/worksheets/sheet<N>.bin
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::string path("xl/worksheets/sheet");
    path.append(std::to_string(i + 1));
    path.append(".bin");
    if (auto r = AddPartBytes(writer.get(), path, sheet_bodies[i]); !r) {
      return r.error();
    }
  }
  // 6. xl/sharedStrings.bin (conditional)
  if (emit_sst_part) {
    auto sst_body_or = emit_sst(sst);
    if (!sst_body_or) {
      return sst_body_or.error();
    }
    if (auto r = AddPartBytes(writer.get(), "xl/sharedStrings.bin", sst_body_or.value()); !r) {
      return r.error();
    }
  }
  // 7. Passthrough parts.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (auto r = AddPartBytes(writer.get(), part->path, part->bytes); !r) {
      return r.error();
    }
  }

  // Finalise into a heap buffer, then copy into a std::vector.
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(writer.get(), &archive_ptr, &archive_size) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_finalize_heap_archive failed",
                      "context=write_xlsb");
  }
  if (mz_zip_writer_end(writer.get()) == MZ_FALSE) {
    if (archive_ptr != nullptr) {
      mz_free(archive_ptr);
    }
    writer.release();
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_end failed", "context=write_xlsb");
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

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
