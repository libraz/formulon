// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end detection + passthrough tests for `.xlsm` / `.xltx` /
// `.xltm` workbook variants. Each test builds an in-memory OOXML
// package via miniz with the content type that triggers the variant,
// runs `read_ooxml`, asserts that `Workbook::kind()` was set correctly
// and that any `xl/vbaProject.bin` payload landed in
// `passthrough_parts()`. The macro-enabled cases additionally
// re-emit the workbook via `Workbook::save()` and verify that the VBA
// payload is byte-identical on a second read — the round-trip
// guarantee Excel users expect.
//
// The engine never executes VBA; these tests are purely about
// detection and verbatim passthrough. See `[OPC]` part 1 §10 /
// `[ECMA-376]` for the canonical content-type strings.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/workbook_kind.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  const char* path;
  std::string_view body;
};

/// Materialises `parts` (text bodies) plus optional binary parts into a
/// heap-allocated zip archive byte vector via miniz.
std::vector<std::uint8_t> BuildZipMixed(
    const std::vector<PartFile>& parts,
    const std::vector<std::pair<const char*, std::vector<std::uint8_t>>>& bin_parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  for (const auto& bp : bin_parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, bp.first, bp.second.data(), bp.second.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add (binary) failed for " << bp.first;
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

// ---------------------------------------------------------------------------
// Synthetic VBA payload — 32 distinctive bytes. It does not need to be a
// real OLE compound file: the writer never inspects it, only re-emits.
// ---------------------------------------------------------------------------

std::vector<std::uint8_t> SyntheticVbaPayload() {
  std::vector<std::uint8_t> bytes;
  bytes.reserve(32);
  for (std::uint8_t i = 0; i < 32; ++i) {
    // Mix in a non-trivial pattern so byte-equality assertions are
    // meaningful even if the writer accidentally zeroes or truncates.
    bytes.push_back(static_cast<std::uint8_t>(0xA0u ^ (i * 17u)));
  }
  return bytes;
}

// ---------------------------------------------------------------------------
// Common XML scaffolding — a single empty sheet, no shared strings, no
// styles relationship. Each variant differs only in `[Content_Types].xml`
// (workbook content-type string) and whether `xl/vbaProject.bin` is
// included.
// ---------------------------------------------------------------------------

constexpr std::string_view kPackageRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
    "  <Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
    "Target=\"xl/workbook.xml\"/>\n"
    "</Relationships>\n";

constexpr std::string_view kWorkbookXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
    "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
    "  <sheets>\n"
    "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
    "  </sheets>\n"
    "</workbook>\n";

constexpr std::string_view kWorkbookRels =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
    "  <Relationship Id=\"rId1\" "
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
    "Target=\"worksheets/sheet1.xml\"/>\n"
    "</Relationships>\n";

constexpr std::string_view kSheetXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
    "  <sheetData/>\n"
    "</worksheet>\n";

/// Builds `[Content_Types].xml` for a workbook of `kind`, optionally
/// declaring an Override for `xl/vbaProject.bin` when `with_vba` is
/// true. The engine only first-class-recognises the workbook part's
/// content type; vbaProject is captured generically through the
/// Override-passthrough mechanism, so the Override entry is what
/// triggers preservation.
std::string BuildContentTypesFor(WorkbookKind kind, bool with_vba) {
  std::string out;
  out.reserve(512);
  out.append(
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" ContentType=\"");
  out.append(workbook_kind_content_type(kind));
  out.append(
      "\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n");
  if (with_vba) {
    out.append(
        "  <Override PartName=\"/xl/vbaProject.bin\" "
        "ContentType=\"application/vnd.ms-office.vbaProject\"/>\n");
  }
  out.append("</Types>\n");
  return out;
}

/// Locates the passthrough entry for `xl/vbaProject.bin` in `parts`,
/// or returns nullptr when absent.
const PassthroughPart* FindVba(const std::vector<PassthroughPart>& parts) {
  auto it =
      std::find_if(parts.begin(), parts.end(), [](const PassthroughPart& p) { return p.path == "xl/vbaProject.bin"; });
  return it == parts.end() ? nullptr : &*it;
}

// ---------------------------------------------------------------------------
// .xlsm: detection + VBA passthrough + writer round-trip.
// ---------------------------------------------------------------------------

TEST(OoxmlXlsm, DetectsXlsmKindAndCapturesVba) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXlsm, /*with_vba=*/true);
  const std::vector<std::uint8_t> vba_bytes = SyntheticVbaPayload();

  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {{"xl/vbaProject.bin", vba_bytes}});

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const OoxmlReadResult& result = result_or.value();

  EXPECT_EQ(result.workbook.kind(), WorkbookKind::kXlsm);

  // VBA payload landed on the workbook's passthrough list, byte-identical.
  const PassthroughPart* vba = FindVba(result.workbook.passthrough_parts());
  ASSERT_NE(vba, nullptr) << "vbaProject.bin missing from passthrough list";
  EXPECT_EQ(vba->content_type, "application/vnd.ms-office.vbaProject");
  ASSERT_EQ(vba->bytes.size(), vba_bytes.size());
  EXPECT_TRUE(std::equal(vba->bytes.begin(), vba->bytes.end(), vba_bytes.begin()))
      << "vbaProject.bin bytes diverged on read";
}

TEST(OoxmlXlsm, RoundTripsXlsmThroughWriter) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXlsm, /*with_vba=*/true);
  const std::vector<std::uint8_t> vba_bytes = SyntheticVbaPayload();

  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {{"xl/vbaProject.bin", vba_bytes}});

  auto first_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(first_or)) << "first read_ooxml: " << first_or.error().message;
  const Workbook& first_wb = first_or.value().workbook;
  ASSERT_EQ(first_wb.kind(), WorkbookKind::kXlsm);

  // Re-emit, then read again and verify kind + VBA bytes survive.
  auto save_or = first_wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;

  auto second_or = read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or)) << "second read_ooxml: " << second_or.error().message;
  const Workbook& second_wb = second_or.value().workbook;
  EXPECT_EQ(second_wb.kind(), WorkbookKind::kXlsm);

  const PassthroughPart* vba = FindVba(second_wb.passthrough_parts());
  ASSERT_NE(vba, nullptr) << "vbaProject.bin lost on round-trip";
  ASSERT_EQ(vba->bytes.size(), vba_bytes.size());
  EXPECT_TRUE(std::equal(vba->bytes.begin(), vba->bytes.end(), vba_bytes.begin()))
      << "vbaProject.bin bytes diverged on round-trip";
}

// ---------------------------------------------------------------------------
// .xltx: template (no VBA).
// ---------------------------------------------------------------------------

TEST(OoxmlXlsm, DetectsXltxKindNoVba) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXltx, /*with_vba=*/false);
  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {});

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  EXPECT_EQ(wb.kind(), WorkbookKind::kXltx);
  EXPECT_EQ(FindVba(wb.passthrough_parts()), nullptr);
}

TEST(OoxmlXlsm, RoundTripsXltxThroughWriter) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXltx, /*with_vba=*/false);
  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {});

  auto first_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(first_or));
  ASSERT_EQ(first_or.value().workbook.kind(), WorkbookKind::kXltx);

  auto save_or = first_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or));

  auto second_or = read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or));
  EXPECT_EQ(second_or.value().workbook.kind(), WorkbookKind::kXltx);
}

// ---------------------------------------------------------------------------
// .xltm: macro-enabled template.
// ---------------------------------------------------------------------------

TEST(OoxmlXlsm, DetectsXltmKindAndCapturesVba) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXltm, /*with_vba=*/true);
  const std::vector<std::uint8_t> vba_bytes = SyntheticVbaPayload();

  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {{"xl/vbaProject.bin", vba_bytes}});

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  EXPECT_EQ(wb.kind(), WorkbookKind::kXltm);

  const PassthroughPart* vba = FindVba(wb.passthrough_parts());
  ASSERT_NE(vba, nullptr);
  ASSERT_EQ(vba->bytes.size(), vba_bytes.size());
  EXPECT_TRUE(std::equal(vba->bytes.begin(), vba->bytes.end(), vba_bytes.begin()));
}

TEST(OoxmlXlsm, RoundTripsXltmThroughWriter) {
  const std::string content_types = BuildContentTypesFor(WorkbookKind::kXltm, /*with_vba=*/true);
  const std::vector<std::uint8_t> vba_bytes = SyntheticVbaPayload();

  const std::vector<std::uint8_t> bytes = BuildZipMixed({{"[Content_Types].xml", content_types},
                                                         {"_rels/.rels", kPackageRels},
                                                         {"xl/workbook.xml", kWorkbookXml},
                                                         {"xl/_rels/workbook.xml.rels", kWorkbookRels},
                                                         {"xl/worksheets/sheet1.xml", kSheetXml}},
                                                        {{"xl/vbaProject.bin", vba_bytes}});

  auto first_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(first_or));
  ASSERT_EQ(first_or.value().workbook.kind(), WorkbookKind::kXltm);

  auto save_or = first_or.value().workbook.save();
  ASSERT_TRUE(static_cast<bool>(save_or));

  auto second_or = read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or));
  const Workbook& second_wb = second_or.value().workbook;
  EXPECT_EQ(second_wb.kind(), WorkbookKind::kXltm);

  const PassthroughPart* vba = FindVba(second_wb.passthrough_parts());
  ASSERT_NE(vba, nullptr) << "vbaProject.bin lost on .xltm round-trip";
  ASSERT_EQ(vba->bytes.size(), vba_bytes.size());
  EXPECT_TRUE(std::equal(vba->bytes.begin(), vba->bytes.end(), vba_bytes.begin()));
}

// ---------------------------------------------------------------------------
// Writer regression: emitted [Content_Types].xml uses the kind-specific
// workbook content-type string, not the plain `.xlsx` constant.
// ---------------------------------------------------------------------------

TEST(OoxmlXlsm, WriterEmitsKindSpecificContentType) {
  Workbook wb = Workbook::create();
  wb.set_kind(WorkbookKind::kXlsm);

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;
  const std::vector<std::uint8_t>& bytes = save_or.value();

  // Pull `[Content_Types].xml` directly out of the archive and look for
  // the macro-enabled string.
  mz_zip_archive reader{};
  ASSERT_NE(mz_zip_reader_init_mem(&reader, bytes.data(), bytes.size(), 0), MZ_FALSE);
  const int idx = mz_zip_reader_locate_file(&reader, "[Content_Types].xml", nullptr, 0);
  ASSERT_GE(idx, 0);
  std::size_t extracted_size = 0;
  void* extracted = mz_zip_reader_extract_to_heap(&reader, static_cast<mz_uint>(idx), &extracted_size, 0);
  ASSERT_NE(extracted, nullptr);
  const std::string body(static_cast<const char*>(extracted), extracted_size);
  mz_free(extracted);
  mz_zip_reader_end(&reader);

  EXPECT_NE(body.find("application/vnd.ms-excel.sheet.macroEnabled.main+xml"), std::string::npos)
      << "writer did not emit the .xlsm content-type string";
  EXPECT_EQ(body.find("spreadsheetml.sheet.main+xml"), std::string::npos)
      << "writer leaked the plain .xlsx content-type string for an .xlsm workbook";
}

// ---------------------------------------------------------------------------
// Writer regression: emitting an .xlsm workbook with vbaProject.bin in
// passthrough must NOT trigger the collision-detection code path. The
// generated path set excludes vbaProject, so the part is preserved and
// the ouput archive contains it once and only once.
// ---------------------------------------------------------------------------

TEST(OoxmlXlsm, WriterDoesNotCollideWithVbaPassthrough) {
  Workbook wb = Workbook::create();
  wb.set_kind(WorkbookKind::kXlsm);

  PassthroughPart vba;
  vba.path = "xl/vbaProject.bin";
  vba.content_type = "application/vnd.ms-office.vbaProject";
  vba.bytes = SyntheticVbaPayload();
  std::vector<PassthroughPart> parts;
  parts.push_back(std::move(vba));
  wb.set_passthrough_parts(std::move(parts));

  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save: " << save_or.error().message;

  // Round-trip the produced bytes — the round-trip is the strongest
  // assertion that nothing collided / was double-emitted: if vbaProject
  // had been added twice, miniz would either reject the second add or
  // produce a malformed archive.
  auto second_or = read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or)) << "round-trip after vba passthrough failed";
  const Workbook& second_wb = second_or.value().workbook;
  EXPECT_EQ(second_wb.kind(), WorkbookKind::kXlsm);

  const PassthroughPart* round_tripped = FindVba(second_wb.passthrough_parts());
  ASSERT_NE(round_tripped, nullptr);
  EXPECT_EQ(round_tripped->bytes.size(), 32U);
}

}  // namespace
}  // namespace io
}  // namespace formulon
