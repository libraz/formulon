//
// Unit tests for `formulon::io::read_ooxml`. The reader is paired with
// `Workbook::save()` (the empty-workbook writer) so the test corpus is
// generated in-process; no fixture files are needed for this slice.

#include "io/ooxml_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml/package_validator.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "utils/error.h"
#include "utils/structured_log.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return ByteSpan{bytes.data(), bytes.size()};
}

/// Collects structured-log records for the duration of one test and
/// restores the shipped configuration (no sink, `kOff`) afterwards, so
/// enabling logging here cannot leak into a sibling test.
class StructuredLogCapture {
 public:
  StructuredLogCapture() {
    set_structured_log_sink(&Append, &records_);
    set_structured_log_min_level(StructuredLogLevel::kDebug);
  }
  ~StructuredLogCapture() {
    set_structured_log_sink(nullptr);
    set_structured_log_min_level(StructuredLogLevel::kOff);
  }

  StructuredLogCapture(const StructuredLogCapture&) = delete;
  StructuredLogCapture& operator=(const StructuredLogCapture&) = delete;

  const std::string& records() const { return records_; }

 private:
  static void Append(std::string_view record, void* user_data) { static_cast<std::string*>(user_data)->append(record); }

  std::string records_;
};

std::vector<std::uint8_t> SaveOrDie(const Workbook& wb) {
  auto save_or = wb.save();
  EXPECT_TRUE(static_cast<bool>(save_or)) << "save() failed in test setup";
  return save_or.value();
}

TEST(OoxmlReader, RoundTripsSingleSheetName) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  OoxmlReadResult& result = result_or.value();
  ASSERT_EQ(result.workbook.sheet_count(), 1U);
  EXPECT_EQ(result.workbook.sheet(0).name(), "Sheet1");
}

TEST(OoxmlReader, RoundTripsThreeSheetsInOrder) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  wb.add_sheet("\xE3\x82\xAC\xE3\x83\xB3\xE3\x83\x9E");  // "ガンマ" in UTF-8

  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  OoxmlReadResult& result = result_or.value();
  ASSERT_EQ(result.workbook.sheet_count(), 3U);
  EXPECT_EQ(result.workbook.sheet(0).name(), "Alpha");
  EXPECT_EQ(result.workbook.sheet(1).name(), "Beta");
  EXPECT_EQ(result.workbook.sheet(2).name(), "\xE3\x82\xAC\xE3\x83\xB3\xE3\x83\x9E");
}

// Regression: every text-cell `Value::text` must remain valid for the
// workbook's lifetime, NOT only for the lifetime of the
// `OoxmlReadResult`. Earlier slices stashed the inline-string deque on
// the read result; moving the workbook out of the result and discarding
// the result therefore left every text view dangling. The fix moves
// the deque onto the workbook itself; this test pins that contract.
TEST(OoxmlReader, WorkbookOutlivesReadResult) {
  // Build a workbook with at least one inline-string cell so the
  // text-storage deque is exercised on the read path.
  Workbook src = Workbook::create();
  std::string greeting = "Hello, world!";
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0U, 0U, 0U, Value::text(greeting))));
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  // Move the workbook out of the read result and immediately drop the
  // result. The workbook must still report the cell text correctly.
  Workbook wb = Workbook::create_empty();
  {
    auto result_or = read_ooxml(SpanOf(bytes));
    ASSERT_TRUE(static_cast<bool>(result_or));
    wb = std::move(result_or.value().workbook);
  }  // OoxmlReadResult destructed here.

  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = wb.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "Hello, world!");

  // And `save()` must succeed — the writer walks every cell's
  // `Value::text` view and would crash under the old layout.
  auto save_or = wb.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  EXPECT_FALSE(save_or.value().empty());
}

TEST(OoxmlReader, MultiMillionCellRangeLoadsWithoutMaterialisingItsCells) {
  // The reader routes every formula cell through `Workbook::set_cell_formula`,
  // so a load registers dependencies exactly like an interactive edit does.
  // A two-column, million-row reference is the shape whose per-cell graph
  // would be two million permanently resident nodes; registered as one
  // compact rectangle it contributes none, which is what keeps the loaded
  // workbook's retained heap independent of the referenced area.
  Workbook src = Workbook::create();
  ASSERT_TRUE(static_cast<bool>(src.set_cell_formula(0U, 0U, 3U, "=SUM(A1:B1000000)")));
  const std::vector<std::uint8_t> bytes = SaveOrDie(src);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml failed: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  const Cell* d1 = wb.sheet(0).cell_at(0U, 3U);
  ASSERT_NE(d1, nullptr);
  EXPECT_EQ(d1->formula_text, "=SUM(A1:B1000000)");
  EXPECT_EQ(wb.recalc_engine().dep_graph().node_count(), 0U);
}

TEST(OoxmlReader, JapaneseSheetNameSurvivesRoundTrip) {
  Workbook wb = Workbook::create();
  // "日本シート" in UTF-8: E6 97 A5 E6 9C AC E3 82 B7 E3 83 BC E3 83 88
  wb.sheet(0).set_name("\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.value().workbook.sheet(0).name(), "\xE6\x97\xA5\xE6\x9C\xAC\xE3\x82\xB7\xE3\x83\xBC\xE3\x83\x88");
}

TEST(OoxmlReader, RejectsNonZipBuffer) {
  const std::uint8_t garbage[] = {0u, 0u, 0u, 0u};
  ByteSpan span{garbage, sizeof(garbage)};
  auto result_or = read_ooxml(span);
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipCorrupt);
}

/// Builds an in-memory ZIP that omits a specific entry from the writer's
/// output. Used to verify reader error paths.
std::vector<std::uint8_t> RebuildArchiveWithout(const std::vector<std::uint8_t>& src, std::string_view skip_entry) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    if (std::string_view(name_buf) == skip_entry) {
      continue;
    }
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);

  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithoutWorkbookXml) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "xl/workbook.xml");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoSheetCorrupt);
}

TEST(OoxmlReader, RejectsArchiveWithoutContentTypes) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "[Content_Types].xml");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoContentTypeInvalid);
}

TEST(OoxmlReader, RejectsArchiveWithoutPackageRels) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithout(ok_bytes, "_rels/.rels");

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoRelationshipBroken);
}

TEST(OoxmlReader, UnknownPartsExcludesSheetAndStylesAndSst) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> bytes = SaveOrDie(wb);

  auto result_or = read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or));
  const std::vector<PassthroughPart>& parts = result_or.value().workbook.passthrough_parts();

  // The reader consumes sheet*.xml and xl/styles.xml. The latter is a
  // best-effort validate-only parse but still counts as consumed so the
  // round-trip pipeline does not surface it as unknown. The empty
  // workbook does not emit a sharedStrings part, so the assertion below
  // is just sanity: it must not appear either.
  auto path_eq = [&parts](std::string_view target) {
    return std::find_if(parts.begin(), parts.end(), [target](const PassthroughPart& p) { return p.path == target; });
  };
  EXPECT_EQ(path_eq("xl/styles.xml"), parts.end()) << "styles.xml unexpectedly still in unknown_parts";
  EXPECT_EQ(path_eq("xl/worksheets/sheet1.xml"), parts.end()) << "sheet1.xml unexpectedly still in unknown_parts";
  EXPECT_EQ(path_eq("xl/sharedStrings.xml"), parts.end()) << "sharedStrings.xml unexpectedly in unknown_parts";
}

// ---------------------------------------------------------------------------
// Path-traversal hardening for `ResolveRelativePath`. A hostile rels file
// with a Target like `../../../etc/passwd` must not be allowed to walk
// out of the package root: even though the resulting path never reaches
// the filesystem (the reader looks it up in the in-memory ZIP catalogue),
// allowing the value to flow through to downstream stages — `has_entry`,
// passthrough emission, link resolution — invites confusion-shaped bugs
// at minimum and a real escape if any caller ever maps the path back to
// disk.
// ---------------------------------------------------------------------------

TEST(OoxmlReader, ResolveRelativePathRefusesEscapingTarget) {
  // base_dir = "xl" + target = "../../etc/passwd": one ".." pops "xl",
  // the second ".." escapes the package root and must trigger
  // kIoZipSlip.
  auto result = internal::ResolveRelativePathForTesting("xl", "../../etc/passwd");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
  // Diagnostic context must echo both arguments so the offending input
  // is recoverable from logs.
  EXPECT_NE(result.error().context.find("base_dir=xl"), std::string::npos);
  EXPECT_NE(result.error().context.find("target=../../etc/passwd"), std::string::npos);
}

TEST(OoxmlReader, ResolveRelativePathRefusesPackageAbsoluteTarget) {
  // A package-absolute Target ("/xl/...") bypasses base-dir-relative
  // accounting entirely; we refuse it rather than silently treat it as
  // root-relative.
  auto result = internal::ResolveRelativePathForTesting("xl/_rels", "/etc/passwd");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(OoxmlReader, ResolveRelativePathAllowsValidParentRefs) {
  // A single ".." that stays inside the package is legitimate:
  // xl/worksheets/_rels/sheet1.xml.rels Target="../../theme/theme1.xml"
  // resolves to "xl/theme/theme1.xml".
  auto result = internal::ResolveRelativePathForTesting("xl/worksheets", "../theme/theme1.xml");
  ASSERT_TRUE(static_cast<bool>(result)) << "rejected legitimate parent-relative target";
  EXPECT_EQ(result.value(), "xl/theme/theme1.xml");
}

TEST(OoxmlReader, ResolveRelativePathRefusesEmptyResult) {
  // Pop every directory off the stack but stop short of escaping; the
  // resolver previously returned an empty string here, which downstream
  // callers would happily store into rid->path maps. We now refuse
  // empty results outright.
  auto result = internal::ResolveRelativePathForTesting("xl", "..");
  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kIoZipSlip);
}

/// Builds an `xl/_rels/workbook.xml.rels` body whose worksheet
/// relationship Target attempts to escape the package root. Used to
/// drive the end-to-end rejection test below.
std::vector<std::uint8_t> RebuildArchiveWithMaliciousWorkbookRels(const std::vector<std::uint8_t>& src) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  static constexpr std::string_view kMaliciousRels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
      "<Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"../../etc/passwd\"/>"
      "</Relationships>";

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    const std::string_view name(name_buf);
    if (name == "xl/_rels/workbook.xml.rels") {
      EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, kMaliciousRels.data(), kMaliciousRels.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
                MZ_FALSE);
      continue;
    }
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithEscapingRelsTarget) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithMaliciousWorkbookRels(ok_bytes);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipSlip);
}

// The OfficeDocument target is both an archive key and the base directory
// every downstream rels target resolves against, so it goes through the
// same `is_safe_part_name` gate as the rest of the package rels rather
// than being merely leading-slash-stripped.
TEST(OoxmlReader, ResolveOfficeDocumentPathRejectsTraversalTarget) {
  const auto rels_bytes = [](std::string_view target) {
    std::string xml(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
        "<Relationship Id=\"rId1\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
        "Target=\"");
    xml.append(target);
    xml.append("\"/></Relationships>");
    return std::vector<std::uint8_t>(xml.begin(), xml.end());
  };

  for (const char* hostile :
       {"../../etc/passwd", "/../evil/workbook.xml", "xl/../../out.xml", "xl\\workbook.xml", "C:/Windows/system32"}) {
    auto result_or = ooxml::resolve_office_document_path(rels_bytes(hostile));
    ASSERT_FALSE(static_cast<bool>(result_or)) << hostile;
    EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipSlip) << hostile;
  }

  // Both canonical spellings still resolve to the same archive key.
  for (const char* ok : {"xl/workbook.xml", "/xl/workbook.xml"}) {
    auto result_or = ooxml::resolve_office_document_path(rels_bytes(ok));
    ASSERT_TRUE(static_cast<bool>(result_or)) << ok;
    EXPECT_EQ(result_or.value(), "xl/workbook.xml") << ok;
  }
}

/// Rebuilds a valid archive with one entry's payload replaced. Used to
/// feed the reader worksheet markup the writer would never emit.
std::vector<std::uint8_t> RebuildArchiveReplacing(const std::vector<std::uint8_t>& src, std::string_view entry,
                                                  std::string_view contents) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  bool replaced = false;
  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    if (std::string_view(name_buf) == entry) {
      EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, contents.data(), contents.size(),
                                      static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
                MZ_FALSE);
      replaced = true;
      continue;
    }
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);
  EXPECT_TRUE(replaced) << entry;

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

// The iterative solver stops once the largest change across the cycle
// falls below `iterateDelta`. `strtod` would turn these spellings into a
// NaN or an infinity, and `max_delta < NaN` is false forever: the
// workbook would silently burn its whole iteration budget and report the
// unconverged values. A tolerance outside the shared non-negative-double
// lexical space keeps the engine default instead.
TEST(OoxmlReader, IterateDeltaOutsideTheLexicalSpaceKeepsTheEngineDefault) {
  for (const char* spelling : {"INF", "NaN", "1e999", "-0.001", "abc"}) {
    const std::string workbook_xml =
        std::string(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            "<calcPr iterate=\"1\" iterateCount=\"50\" iterateDelta=\"")
            .append(spelling)
            .append(
                "\"/>"
                "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
                "</workbook>");
    const std::vector<std::uint8_t> mutated =
        RebuildArchiveReplacing(SaveOrDie(Workbook::create()), "xl/workbook.xml", workbook_xml);

    auto loaded_or = read_ooxml(SpanOf(mutated));
    ASSERT_TRUE(static_cast<bool>(loaded_or)) << spelling << ": " << loaded_or.error().message;
    const eval::IterativeOptions& opts = loaded_or.value().workbook.iterative_options();
    // The attributes that do lex are still honoured, so the rejection is
    // scoped to the one value rather than to the whole element.
    EXPECT_TRUE(opts.enabled) << spelling;
    EXPECT_EQ(opts.max_iterations, 50U) << spelling;
    EXPECT_DOUBLE_EQ(opts.max_change, eval::kDefaultMaxChange) << spelling;
  }
}

TEST(OoxmlReader, WellFormedIterateDeltaIsHonoured) {
  constexpr std::string_view kWorkbookXml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
      "<calcPr iterate=\"1\" iterateCount=\"50\" iterateDelta=\"0.25\"/>"
      "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
      "</workbook>";
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(SaveOrDie(Workbook::create()), "xl/workbook.xml", kWorkbookXml);

  auto loaded_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(loaded_or)) << loaded_or.error().message;
  EXPECT_DOUBLE_EQ(loaded_or.value().workbook.iterative_options().max_change, 0.25);
}

// A `frozenSplit` pane must open frozen, and the freeze must survive a
// save: the model does not carry the split position, so the writer
// re-emits the equivalent `state="frozen"` form.
TEST(OoxmlReader, FrozenSplitPaneSurvivesLoadAndSave) {
  constexpr std::string_view kSheetXml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
      "<sheetViews><sheetView workbookViewId=\"0\">"
      "<pane xSplit=\"1\" ySplit=\"1\" topLeftCell=\"B2\" activePane=\"bottomRight\" state=\"frozenSplit\"/>"
      "</sheetView></sheetViews>"
      "<sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c></row></sheetData>"
      "</worksheet>";
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(SaveOrDie(Workbook::create()), "xl/worksheets/sheet1.xml", kSheetXml);

  auto loaded_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(loaded_or)) << loaded_or.error().message;
  EXPECT_EQ(loaded_or.value().workbook.sheet(0).view().freeze_rows, 1U);
  EXPECT_EQ(loaded_or.value().workbook.sheet(0).view().freeze_cols, 1U);

  const std::vector<std::uint8_t> resaved = SaveOrDie(loaded_or.value().workbook);
  auto reloaded_or = read_ooxml(SpanOf(resaved));
  ASSERT_TRUE(static_cast<bool>(reloaded_or)) << reloaded_or.error().message;
  EXPECT_EQ(reloaded_or.value().workbook.sheet(0).view().freeze_rows, 1U);
  EXPECT_EQ(reloaded_or.value().workbook.sheet(0).view().freeze_cols, 1U);
}

/// One sheet carrying a malformed reference in each of the four
/// presentation-overlay kinds, plus one well-formed merge and one live
/// cell so the two tests below can tell "dropped the bad entry" apart
/// from "dropped everything".
constexpr std::string_view kMalformedOverlaysSheetXml =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<sheetData><row r=\"1\"><c r=\"A1\"><v>42</v></c></row></sheetData>"
    "<conditionalFormatting sqref=\"\"><cfRule type=\"expression\" priority=\"1\"/></conditionalFormatting>"
    "<mergeCells count=\"2\"><mergeCell ref=\"\"/><mergeCell ref=\"B1:C1\"/></mergeCells>"
    "<dataValidations count=\"1\"><dataValidation type=\"list\" sqref=\"\"/></dataValidations>"
    "<hyperlinks><hyperlink ref=\"\" location=\"Sheet1!A1\"/></hyperlinks>"
    "</worksheet>";

// Presentation overlays degrade per entry rather than rejecting the
// package: a malformed merge / hyperlink / data-validation reference must
// not cost the caller every cell value in the workbook. This is the same
// disposition the CF reader has always applied to a malformed block.
TEST(OoxmlReader, MalformedPresentationOverlaysDoNotFailTheLoad) {
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(SaveOrDie(Workbook::create()), "xl/worksheets/sheet1.xml", kMalformedOverlaysSheetXml);

  auto loaded_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(loaded_or)) << loaded_or.error().message;
  const Sheet& sheet = loaded_or.value().workbook.sheet(0);
  // The cell data — the part that is not a presentation overlay — is intact.
  ASSERT_NE(sheet.cell_at(0U, 0U), nullptr);
  EXPECT_DOUBLE_EQ(sheet.cell_at(0U, 0U)->cached_value.as_number(), 42.0);
  // Only the malformed entries were dropped; the well-formed merge survived.
  EXPECT_TRUE(sheet.conditional_formats().empty());
  ASSERT_EQ(sheet.merges().size(), 1U);
  EXPECT_EQ(sheet.merges()[0].first_col, 1U);
  EXPECT_TRUE(sheet.validations().empty());
  EXPECT_TRUE(sheet.hyperlinks().empty());
}

// Degrading has to be observable. Dropping an overlay entry is a silent
// data loss unless the reader says so, so the diagnostic is half of the
// contract, not decoration — and because logging ships off, "emitted"
// means "reaches an embedder's sink", which is what this asserts.
TEST(OoxmlReader, MalformedPresentationOverlaysAreDiagnosed) {
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(SaveOrDie(Workbook::create()), "xl/worksheets/sheet1.xml", kMalformedOverlaysSheetXml);

  StructuredLogCapture log;
  auto loaded_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(loaded_or)) << loaded_or.error().message;

  // One record per skipped entry, each naming the part it came from, so a
  // host can tell which overlay it lost.
  const std::string& records = log.records();
  EXPECT_NE(records.find("io.cf.skip"), std::string::npos) << records;
  EXPECT_NE(records.find("\"part\":\"mergeCells\""), std::string::npos) << records;
  EXPECT_NE(records.find("\"part\":\"hyperlinks\""), std::string::npos) << records;
  EXPECT_NE(records.find("\"part\":\"dataValidations\""), std::string::npos) << records;
  // The three sibling readers share one event name; CF keeps its own
  // because it skips a whole block rather than a single entry.
  EXPECT_NE(records.find("io.sheet.overlay.skip"), std::string::npos) << records;
}

TEST(OoxmlReader, IsSafePartNameAcceptsCanonicalNames) {
  EXPECT_TRUE(ooxml::is_safe_part_name("xl/workbook.xml"));
  EXPECT_TRUE(ooxml::is_safe_part_name("xl/media/image1.png"));
  EXPECT_TRUE(ooxml::is_safe_part_name("[Content_Types].xml"));
}

TEST(OoxmlReader, IsSafePartNameRejectsTraversalAndAbsoluteShapes) {
  EXPECT_FALSE(ooxml::is_safe_part_name(""));
  EXPECT_FALSE(ooxml::is_safe_part_name("/etc/passwd"));            // package-absolute
  EXPECT_FALSE(ooxml::is_safe_part_name("../../etc/passwd"));       // parent traversal
  EXPECT_FALSE(ooxml::is_safe_part_name("xl/../../../etc/x"));      // embedded traversal
  EXPECT_FALSE(ooxml::is_safe_part_name("xl\\worksheets\\a.xml"));  // backslash separator
  EXPECT_FALSE(ooxml::is_safe_part_name("C:/Windows/system32"));    // drive-letter colon
  EXPECT_FALSE(ooxml::is_safe_part_name("xl//workbook.xml"));       // empty segment
  EXPECT_FALSE(ooxml::is_safe_part_name("xl/./workbook.xml"));      // dot segment
}

/// Rebuilds a valid archive with one extra archive entry whose name
/// escapes the package root (`../evil.bin`). The reader's Default-typed
/// sweep must refuse to carry it through passthrough.
std::vector<std::uint8_t> RebuildArchiveWithMaliciousExtraEntry(const std::vector<std::uint8_t>& src) {
  mz_zip_archive reader{};
  EXPECT_NE(mz_zip_reader_init_mem(&reader, src.data(), src.size(), 0), MZ_FALSE);
  const mz_uint count = mz_zip_reader_get_num_files(&reader);

  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);

  for (mz_uint i = 0; i < count; ++i) {
    char name_buf[256] = {};
    const mz_uint name_len = mz_zip_reader_get_filename(&reader, i, name_buf, sizeof(name_buf));
    EXPECT_GT(name_len, 0u);
    std::size_t extracted_size = 0;
    void* extracted = mz_zip_reader_extract_to_heap(&reader, i, &extracted_size, 0);
    EXPECT_NE(extracted, nullptr);
    EXPECT_NE(mz_zip_writer_add_mem(&writer, name_buf, extracted, extracted_size,
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE);
    mz_free(extracted);
  }
  mz_zip_reader_end(&reader);

  static constexpr std::string_view kEvil = "malicious";
  EXPECT_NE(mz_zip_writer_add_mem(&writer, "../evil.bin", kEvil.data(), kEvil.size(),
                                  static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
            MZ_FALSE);

  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlReader, RejectsArchiveWithTraversalShapedPartName) {
  Workbook wb = Workbook::create();
  const std::vector<std::uint8_t> ok_bytes = SaveOrDie(wb);
  const std::vector<std::uint8_t> mutated = RebuildArchiveWithMaliciousExtraEntry(ok_bytes);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipSlip);
}

TEST(OoxmlReader, RejectsEncryptedCdfv2ContainerWithEncryptedDiagnostic) {
  // OLE/CDFV2 compound-file signature that wraps a password-protected
  // .xlsx/.xlsb, followed by arbitrary filler. This must surface an
  // "encrypted" diagnostic rather than a generic "corrupt zip".
  std::vector<std::uint8_t> encrypted = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
  encrypted.resize(512, 0x00);

  auto result_or = read_ooxml(SpanOf(encrypted));
  ASSERT_FALSE(static_cast<bool>(result_or));
  EXPECT_EQ(result_or.error().code, FormulonErrorCode::kIoZipEncrypted);
}

/// A worksheet whose CF block names `<dxfs>` entries that are not there:
/// a stale index past the end of the table, and the `-1` a writer emits
/// for "no format", which the attribute's int32 lexical space allows and
/// an unchecked read turns into `0xFFFFFFFF`. The third rule is in range
/// and must be left alone.
constexpr std::string_view kSheetWithUnresolvableDxfIds =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData>"
    "<conditionalFormatting sqref=\"A1:A3\">"
    "<cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" dxfId=\"7\"><formula>1</formula></cfRule>"
    "<cfRule type=\"cellIs\" priority=\"2\" operator=\"lessThan\" dxfId=\"-1\"><formula>0</formula></cfRule>"
    "<cfRule type=\"cellIs\" priority=\"3\" operator=\"equal\" dxfId=\"0\"><formula>5</formula></cfRule>"
    "</conditionalFormatting>"
    "</worksheet>";

TEST(OoxmlReader, UnresolvableCfDxfIdIsDisengagedOnLoad) {
  // A stored style index is supposed to resolve -- that is the promise
  // `NormalizeStyleIndices` keeps for the `<xf>` tables, and the reason a
  // getter may assume it. Reading `dxfId` verbatim broke it for CF alone:
  // a workbook that loaded without error handed out an index its own
  // styles table rejects, so a get / modify / put round trip through any
  // binding could not complete.
  Workbook src = Workbook::create();
  src.mutable_styles().dxfs.resize(1);
  const std::vector<std::uint8_t> base = SaveOrDie(src);
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(base, "xl/worksheets/sheet1.xml", kSheetWithUnresolvableDxfIds);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  ASSERT_EQ(wb.styles().dxfs.size(), 1U);
  ASSERT_EQ(wb.sheet(0).conditional_formats().size(), 1U);
  const std::vector<cf::CFRule>& rules = wb.sheet(0).conditional_formats()[0].rules;
  ASSERT_EQ(rules.size(), 3U);
  EXPECT_FALSE(rules[0].dxf_id.has_value()) << "dxfId=\"7\" against a one-entry table";
  EXPECT_FALSE(rules[1].dxf_id.has_value()) << "dxfId=\"-1\" wraps to 0xFFFFFFFF";
  ASSERT_TRUE(rules[2].dxf_id.has_value()) << "an in-range index must survive untouched";
  EXPECT_EQ(*rules[2].dxf_id, 0U);

  // The invariant itself, stated over the whole loaded workbook: no
  // surface can be handed a `dxf_id` the styles table would reject, which
  // is the only way `fm_styles_get_dxf` can fail on a workbook that
  // loaded.
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const cf::ConditionalFormat& cfmt : wb.sheet(s).conditional_formats()) {
      for (const cf::CFRule& rule : cfmt.rules) {
        if (rule.dxf_id.has_value()) {
          EXPECT_LT(static_cast<std::size_t>(*rule.dxf_id), wb.styles().dxfs.size());
        }
      }
    }
  }
}

TEST(OoxmlReader, DisengagedCfDxfIdIsNotReEmitted) {
  // Normalising on read is what stops the dangling index being written
  // back out, so the defect does not survive a round trip through us.
  Workbook src = Workbook::create();
  src.mutable_styles().dxfs.resize(1);
  const std::vector<std::uint8_t> mutated =
      RebuildArchiveReplacing(SaveOrDie(src), "xl/worksheets/sheet1.xml", kSheetWithUnresolvableDxfIds);

  auto result_or = read_ooxml(SpanOf(mutated));
  ASSERT_TRUE(static_cast<bool>(result_or)) << result_or.error().message;
  const std::vector<std::uint8_t> resaved = SaveOrDie(result_or.value().workbook);

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(resaved))));
  auto entry_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(entry_or)) << entry_or.error().message;
  const std::string sheet_xml(reinterpret_cast<const char*>(entry_or.value().data()), entry_or.value().size());
  EXPECT_EQ(sheet_xml.find("dxfId=\"7\""), std::string::npos) << sheet_xml;
  EXPECT_EQ(sheet_xml.find("dxfId=\"-1\""), std::string::npos) << sheet_xml;
  EXPECT_EQ(sheet_xml.find("dxfId=\"4294967295\""), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("dxfId=\"0\""), std::string::npos) << sheet_xml;
}

TEST(OoxmlReader, EmptyWorkbookFactoryProducesZeroSheets) {
  Workbook wb = Workbook::create_empty();
  EXPECT_EQ(wb.sheet_count(), 0U);
  // save() should reject zero-sheet workbooks (Excel does the same), so
  // the input cannot be round-tripped: the test stops at the factory
  // contract.
  auto save_or = wb.save();
  EXPECT_FALSE(static_cast<bool>(save_or));
}

}  // namespace
}  // namespace io
}  // namespace formulon
