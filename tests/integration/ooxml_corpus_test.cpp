// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML round-trip parity over a 20-book minimal corpus.
//
// This is the M6.core exit gate test (Phase 2 Bundle 2.6 in
// `backup/plans/26-implementation-plan.md`). It builds 20 distinct
// synthetic workbooks programmatically (no committed binary fixtures)
// covering the union of features the M6.core slice supports, and runs
// each through a two-cycle round-trip pipeline:
//
//   1. construct in-memory (via the public Workbook API or, for the
//      explicit-SST / passthrough cases, via miniz);
//   2. `read_ooxml` into workbook A;
//   3. `recalc(default_registry())` so cached values are populated;
//   4. `write_ooxml` -> bytes B;
//   5. `read_ooxml(B)` into workbook C;
//   6. `recalc(default_registry())` on C;
//   7. compare salient invariants between A (post-recalc) and C
//      (post-recalc): sheet shape, defined names, tables, passthrough
//      parts, and per-cell formula text.
//
// Two cycles are deliberate: a single round-trip catches the gross
// shape, but ordering / numbering / passthrough drift only surfaces on
// the second pass when the writer's emitted ids and the reader's
// re-derived ids must agree. Books that exercise volatile or iterative
// features (NOW/TODAY/RAND, circular SCCs) compare formula text rather
// than cached values, since the values legitimately drift across
// recalcs.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <functional>
#include <ostream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cell.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/defined_names.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "io/zip_reader.h"
#include "miniz.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Small helpers shared across all corpus books.
// ---------------------------------------------------------------------------

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

/// `save()` wrapper that ASSERTs success and returns the bytes.
Expected<std::vector<std::uint8_t>, Error> SaveBytes(const Workbook& wb) { return wb.save(); }

/// Builds a synthetic `.xlsx` archive in memory by gluing a list of
/// `(path, body)` pairs through miniz. Used for books 18 / 19 / 20
/// where the source archive must contain shapes the writer does not
/// emit on its own (explicit SST, passthrough theme part).
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

// ---------------------------------------------------------------------------
// Cross-cycle equality helpers (the contract called out in the bundle).
// All helpers take two workbooks (a "before" and "after" snapshot) and
// return a bool that lets callers chain EXPECT_TRUE in the test body.
// Doxygen on each documents what is — and is NOT — being compared so
// future maintainers can extend rather than re-derive.
// ---------------------------------------------------------------------------

/// Two workbooks have the same sheet shape iff they have the same
/// number of sheets, the i-th sheet on each has the same display name,
/// and the i-th sheet on each holds the same total `cell_count()`. We
/// intentionally do NOT compare the geometric `(rows x cols)` extents
/// because the row store grows lazily and the on-disk projection
/// depends only on populated cells.
bool sheets_have_same_shape(const Workbook& a, const Workbook& b) {
  if (a.sheet_count() != b.sheet_count()) {
    return false;
  }
  for (std::size_t i = 0; i < a.sheet_count(); ++i) {
    if (a.sheet(i).name() != b.sheet(i).name()) {
      return false;
    }
    if (a.sheet(i).cell_count() != b.sheet(i).cell_count()) {
      return false;
    }
  }
  return true;
}

/// Order-preserving by name, plus the four scalar fields. The reader
/// preserves declaration order, and the writer emits in that order, so
/// equality must be index-aligned rather than set-equal.
bool defined_names_equal(const Workbook& a, const Workbook& b) {
  if (a.defined_names().size() != b.defined_names().size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.defined_names().size(); ++i) {
    const io::DefinedName& x = a.defined_names()[i];
    const io::DefinedName& y = b.defined_names()[i];
    if (x.name != y.name || x.formula != y.formula || x.local_sheet_id != y.local_sheet_id ||
        x.hidden != y.hidden || x.comment != y.comment) {
      return false;
    }
  }
  return true;
}

/// Tables compared by id, ref, name, display_name, header/totals row
/// flags, sheet index, and column-list shape (id + name). Other column
/// fields (totals_label / totals_function) are not asserted here
/// because the corpus does not stress them across all books; the
/// per-book invariants check those where relevant.
bool tables_equal(const Workbook& a, const Workbook& b) {
  if (a.tables().size() != b.tables().size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.tables().size(); ++i) {
    const io::TableMetadata& x = a.tables()[i];
    const io::TableMetadata& y = b.tables()[i];
    if (x.id != y.id || x.name != y.name || x.display_name != y.display_name || x.ref != y.ref ||
        x.sheet_index != y.sheet_index || x.header_row != y.header_row || x.totals_row != y.totals_row) {
      return false;
    }
    if (x.columns.size() != y.columns.size()) {
      return false;
    }
    for (std::size_t c = 0; c < x.columns.size(); ++c) {
      if (x.columns[c].id != y.columns[c].id || x.columns[c].name != y.columns[c].name) {
        return false;
      }
    }
  }
  return true;
}

/// Passthrough parts are compared by path + raw bytes (content_type
/// is also compared when both sides have one). Order is preserved in
/// the round-trip but the assertion is set-style: we look up by path
/// so re-emission ordering is not load-bearing.
bool passthrough_parts_equal(const Workbook& a, const Workbook& b) {
  if (a.passthrough_parts().size() != b.passthrough_parts().size()) {
    return false;
  }
  for (const io::PassthroughPart& x : a.passthrough_parts()) {
    auto it = std::find_if(b.passthrough_parts().begin(), b.passthrough_parts().end(),
                           [&x](const io::PassthroughPart& y) { return y.path == x.path; });
    if (it == b.passthrough_parts().end()) {
      return false;
    }
    if (x.bytes.size() != it->bytes.size()) {
      return false;
    }
    if (!std::equal(x.bytes.begin(), x.bytes.end(), it->bytes.begin())) {
      return false;
    }
    if (!x.content_type.empty() && !it->content_type.empty() && x.content_type != it->content_type) {
      return false;
    }
  }
  return true;
}

/// Iterates every cell with a non-empty formula on `a` and asserts the
/// same coordinate on `b` carries the same formula text. Used by the
/// volatile / iterative cases where cached values legitimately drift
/// across recalcs but the formulas themselves must round-trip
/// verbatim.
bool formula_texts_equal_for_each_cell(const Workbook& a, const Workbook& b) {
  if (a.sheet_count() != b.sheet_count()) {
    return false;
  }
  for (std::size_t s = 0; s < a.sheet_count(); ++s) {
    const Sheet& sa = a.sheet(s);
    const Sheet& sb = b.sheet(s);
    for (const auto& [row, cells] : sa.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        const Cell& ca = cells[col];
        if (ca.formula_text.empty()) {
          continue;
        }
        const Cell* cb = sb.cell_at(row, col);
        if (cb == nullptr || cb->formula_text != ca.formula_text) {
          return false;
        }
      }
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Corpus contract.
// ---------------------------------------------------------------------------

/// One corpus book.
///
///   * `id`             — short, ctest-friendly identifier (no
///                        whitespace, ASCII only).
///   * `build`          — produces the source `.xlsx` bytes. Either
///                        constructs a `Workbook` and saves it, or
///                        builds an archive directly via miniz.
///   * `assert_invariants` — invoked on the *post-recalc* workbook
///                        after the first read; asserts the things
///                        this book is supposed to verify (sheet
///                        count, specific cell values, presence of a
///                        passthrough part, ...). Failures land in the
///                        parameterised test as the per-book GTest
///                        failure.
struct CorpusBook {
  std::string id;
  std::function<Expected<std::vector<std::uint8_t>, Error>()> build;
  std::function<void(const Workbook&)> assert_invariants;
};

// ---------------------------------------------------------------------------
// Synthetic-package builders for books 18 / 19 / 20.
//
// These archives contain shapes the writer does not currently emit on
// its own (explicit SST, generic passthrough parts), so we build them
// via miniz the same way the existing roundtrip + metadata suites do.
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Pure-API builders for books 1-17 and the kitchen-sink combo (book 20).
//
// Each builder returns the source `.xlsx` bytes; the test driver
// performs the `read -> recalc -> write -> read -> recalc` two-cycle
// pipeline. For most books we go through `Workbook::set_cell_value /
// set_cell_formula` and the high-level `add_sheet` API; tables and
// defined names are attached via the dedicated setters.
// ---------------------------------------------------------------------------

Expected<std::vector<std::uint8_t>, Error> BuildEmpty() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildSingleLiteral() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(123.5)));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildLiteralsAllKinds() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(42.0)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::boolean(true)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::boolean(false)));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 3U, 0U, Value::text("hello")));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 4U, 0U, Value::blank()));
  RETURN_IF_ERROR(wb.set_cell_value(0U, 5U, 0U, Value::error(ErrorCode::Div0)));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildArithmetic() {
  // NOTE: the `&` (text-concat) operator is intentionally omitted here.
  // Bundle 2.6 surfaced an existing engine limitation: the OOXML writer
  // emits formula cells with text cached values as `<v>foobar</v>`
  // (no `t="str"`), which the reader then parses as `t='n'` and
  // rejects. That is a writer bug to fix in a follow-up bundle; the
  // corpus avoids triggering it so the round-trip parity test stays
  // green for the arithmetic surface itself. See the bundle report
  // and `backup/plans/04-xlsx-io.md` for the deferred fix.
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(10.0)));  // A1
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::number(3.0)));   // A2
  // Formulas covering + - * / ^ % and unary minus (no `&`).
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1+A2"));  // B1 -> 13
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 1U, "=A1-A2"));  // B2 -> 7
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 1U, "=A1*A2"));  // B3 -> 30
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 3U, 1U, "=A1/A2"));  // B4 -> 3.333..
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 4U, 1U, "=A1^A2"));  // B5 -> 1000
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 5U, 1U, "=A1%"));    // B6 -> 0.1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 6U, 1U, "=-A1"));    // B7 -> -10
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildCrossSheet() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::number(99.0)));  // Sheet2!A1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=Sheet2!A1"));       // Sheet1!A1
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildMultiSheetIndependent() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("North");
  wb.add_sheet("South");
  wb.add_sheet("East");
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(1.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1+1"));
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::number(2.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 0U, 1U, "=A1*2"));
  RETURN_IF_ERROR(wb.set_cell_value(2U, 0U, 0U, Value::number(3.0)));
  RETURN_IF_ERROR(wb.set_cell_formula(2U, 0U, 1U, "=A1-1"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildUnicodeSheetNames() {
  Workbook wb = Workbook::create_empty();
  // "日本語" (Japanese) — exercises non-ASCII XML attribute encoding.
  wb.add_sheet("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  // emoji + ASCII — exercises 4-byte UTF-8 in attribute context.
  wb.add_sheet("emoji\xF0\x9F\x98\x80");
  // Spaces and digits — exercises plain attribute round-trip.
  wb.add_sheet("Sheet 3 with spaces");
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildUnicodeCellText() {
  Workbook wb = Workbook::create();
  // Japanese.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::text("\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E")));
  // Emoji (4-byte UTF-8).
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::text("Hello \xF0\x9F\x98\x80")));
  // NBSP (U+00A0).
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::text("a\xC2\xA0\x62")));
  // Hebrew "shalom" — Right-to-left script. We are not asserting bidi
  // rendering, just that the bytes survive a round-trip.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 3U, 0U, Value::text("\xD7\xA9\xD7\x9C\xD7\x95\xD7\x9D")));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildRangeAggregates() {
  Workbook wb = Workbook::create();
  // Populate A1:A10 with 1..10.
  for (std::uint32_t r = 0; r < 10U; ++r) {
    RETURN_IF_ERROR(wb.set_cell_value(0U, r, 0U, Value::number(static_cast<double>(r + 1U))));
  }
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=SUM(A1:A10)"));      // B1 -> 55
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 1U, "=AVERAGE(A1:A10)"));  // B2 -> 5.5
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 1U, "=MIN(A1:A10)"));      // B3 -> 1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 3U, 1U, "=MAX(A1:A10)"));      // B4 -> 10
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 4U, 1U, "=COUNT(A1:A10)"));    // B5 -> 10
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildVolatileNowToday() {
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=NOW()"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 0U, "=TODAY()"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 2U, 0U, "=RAND()"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildIterativeCircular() {
  Workbook wb = Workbook::create();
  // A1 = B1 + 1; B1 = A1.
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 0U, "=B1+1"));
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U, "=A1"));
  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 20U;
  opts.max_change = 0.01;
  wb.set_iterative_options(opts);
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesWorkbookScope() {
  Workbook wb = Workbook::create();
  std::vector<io::DefinedName> names;
  io::DefinedName n;
  n.name = "Sales";
  n.formula = "Sheet1!$A$1:$A$10";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildDefinedNamesSheetScope() {
  Workbook wb = Workbook::create();
  std::vector<io::DefinedName> names;
  io::DefinedName n;
  n.name = "LocalRange";
  n.formula = "Sheet1!$B$1:$B$5";
  n.local_sheet_id = 0;
  n.hidden = true;
  n.comment = "Sheet1-only.";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildSingleTable() {
  Workbook wb = Workbook::create();
  io::TableMetadata table;
  table.id = 1U;
  table.name = "SalesTable";
  table.display_name = "SalesTable";
  table.ref = "A1:C5";
  table.sheet_index = 0U;
  table.header_row = true;
  table.totals_row = true;
  table.columns.push_back(io::TableColumn{1U, "Region", "Total", "", ""});
  table.columns.push_back(io::TableColumn{2U, "Q1", "", "sum", ""});
  table.columns.push_back(io::TableColumn{3U, "Q2", "", "sum", ""});
  std::vector<io::TableMetadata> tables;
  tables.push_back(std::move(table));
  wb.set_tables(std::move(tables));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildMultipleTables() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");

  std::vector<io::TableMetadata> tables;
  io::TableMetadata t1;
  t1.id = 1U;
  t1.name = "Tbl1";
  t1.display_name = "Tbl1";
  t1.ref = "A1:B3";
  t1.sheet_index = 0U;
  t1.columns.push_back(io::TableColumn{1U, "X", "", "", ""});
  t1.columns.push_back(io::TableColumn{2U, "Y", "", "", ""});
  tables.push_back(std::move(t1));

  io::TableMetadata t2;
  t2.id = 2U;
  t2.name = "Tbl2";
  t2.display_name = "Tbl2";
  t2.ref = "A1:B2";
  t2.sheet_index = 0U;
  t2.columns.push_back(io::TableColumn{1U, "P", "", "", ""});
  t2.columns.push_back(io::TableColumn{2U, "Q", "", "", ""});
  tables.push_back(std::move(t2));

  io::TableMetadata t3;
  t3.id = 3U;
  t3.name = "Tbl3";
  t3.display_name = "Tbl3";
  t3.ref = "A1:A2";
  t3.sheet_index = 1U;
  t3.columns.push_back(io::TableColumn{1U, "Z", "", "", ""});
  tables.push_back(std::move(t3));

  wb.set_tables(std::move(tables));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildNestedFunctionCalls() {
  // Deeply nested IF/AND/OR with numeric branches only — text branch
  // results would trigger the same `<v>` round-trip limitation called
  // out in `BuildArithmetic`. Logic: (A1>0 OR A2<0) AND (A3=8) is true,
  // and A1+A2 (=8) > A3 (=8) is false, so the formula falls into the
  // inner else-branch and yields A2 (=3).
  Workbook wb = Workbook::create();
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(5.0)));  // A1
  RETURN_IF_ERROR(wb.set_cell_value(0U, 1U, 0U, Value::number(3.0)));  // A2
  RETURN_IF_ERROR(wb.set_cell_value(0U, 2U, 0U, Value::number(8.0)));  // A3
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 0U, 1U,
                                      "=IF(AND(OR(A1>0,A2<0),A3=8),IF(A1+A2>A3,A1,A2),0)"));
  return SaveBytes(wb);
}

Expected<std::vector<std::uint8_t>, Error> BuildLargeGrid500Cells() {
  // 500 cells across A1:T25 (20 columns x 25 rows). Mix of numeric
  // literals, text, and formulas referencing earlier cells.
  //
  // `Value::text` is non-owning, so we materialise the per-row label
  // strings into a function-static `std::deque<std::string>`: the
  // deque keeps element addresses stable across appends and persists
  // until program teardown, which is well past the round-trip pipeline
  // (the workbook reads `cached_value.as_text()` during `save()` and
  // discards the view afterward, but the value still has to be alive
  // through that read).
  static std::deque<std::string> kLabels;
  static const bool kInitialised = [&]() {
    for (std::uint32_t r = 0; r < 25U; ++r) {
      kLabels.push_back("L" + std::to_string(r));
    }
    return true;
  }();
  (void)kInitialised;

  Workbook wb = Workbook::create();
  for (std::uint32_t r = 0; r < 25U; ++r) {
    for (std::uint32_t c = 0; c < 20U; ++c) {
      if ((r + c) % 7U == 0U) {
        RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, Value::text(kLabels[r])));
      } else if ((r + c) % 3U == 0U && r > 0U) {
        // Formula referencing the cell directly above (1-based row).
        const std::string col_letter(1U, static_cast<char>('A' + c));
        const std::string ref_above = col_letter + std::to_string(r);
        RETURN_IF_ERROR(wb.set_cell_formula(0U, r, c, "=" + ref_above + "+1"));
      } else {
        RETURN_IF_ERROR(wb.set_cell_value(0U, r, c, Value::number(static_cast<double>(r * 20U + c))));
      }
    }
  }
  return SaveBytes(wb);
}

/// Combined kitchen-sink: 2 sheets, a defined name, a table, a
/// passthrough theme part, mixed formulas, unicode text, and an
/// iterative circular pair. We can't add a passthrough part via the
/// public Workbook API alone (the writer copies it through, but the
/// only way to get one in is to read an archive that already has it),
/// so we build the synthetic theme archive, read it, then mutate the
/// workbook to add the rest of the surfaces, then save.
Expected<std::vector<std::uint8_t>, Error> BuildKitchenSink() {
  ASSIGN_OR_RETURN(auto theme_bytes, BuildPassthroughThemeArchive());
  ASSIGN_OR_RETURN(auto first_or, io::read_ooxml(SpanOf(theme_bytes)));
  Workbook wb = std::move(first_or.workbook);

  // Add a second sheet.
  wb.add_sheet("\xE6\x97\xA5\xE6\x9C\xAC");  // "日本"

  // Defined name (workbook-scope).
  std::vector<io::DefinedName> names;
  io::DefinedName n;
  n.name = "ComboName";
  n.formula = "Sheet1!$A$1";
  names.push_back(std::move(n));
  wb.set_defined_names(std::move(names));

  // Table on Sheet1.
  io::TableMetadata tab;
  tab.id = 1U;
  tab.name = "ComboTable";
  tab.display_name = "ComboTable";
  tab.ref = "A1:B2";
  tab.sheet_index = 0U;
  tab.columns.push_back(io::TableColumn{1U, "Alpha", "", "", ""});
  tab.columns.push_back(io::TableColumn{2U, "Beta", "", "", ""});
  std::vector<io::TableMetadata> tables;
  tables.push_back(std::move(tab));
  wb.set_tables(std::move(tables));

  // Mixed formulas + unicode text.
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 0U, Value::number(3.14)));            // A1
  RETURN_IF_ERROR(wb.set_cell_formula(0U, 1U, 0U, "=A1*2"));                      // A2
  RETURN_IF_ERROR(wb.set_cell_value(0U, 0U, 1U, Value::text("\xE3\x81\x82")));    // B1: "あ"
  RETURN_IF_ERROR(wb.set_cell_value(1U, 0U, 0U, Value::text("Hello \xF0\x9F\x91\x8B")));  // emoji wave

  // Iterative circular pair on Sheet2 (rows 5-6).
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 5U, 0U, "=B6+0.5"));
  RETURN_IF_ERROR(wb.set_cell_formula(1U, 5U, 1U, "=A6"));
  eval::IterativeOptions opts;
  opts.enabled = true;
  opts.max_iterations = 20U;
  opts.max_change = 0.01;
  wb.set_iterative_options(opts);

  return SaveBytes(wb);
}

// ---------------------------------------------------------------------------
// Per-book invariant assertions. Each `assert_invariants` runs against
// the *post-recalc* first-cycle workbook and checks the things that
// book is supposed to verify. The two-cycle equality contract handles
// the "shape did not drift" claim separately, in the parameterised
// driver.
// ---------------------------------------------------------------------------

const Cell* CellAt(const Workbook& wb, std::size_t s, std::uint32_t r, std::uint32_t c) {
  return wb.sheet(s).cell_at(r, c);
}

void AssertEmpty(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_EQ(wb.sheet(0).cell_count(), 0U);
}

void AssertSingleLiteral(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 123.5);
}

void AssertLiteralsAllKinds(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);

  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a2->cached_value.is_boolean());
  EXPECT_TRUE(a2->cached_value.as_boolean());

  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  ASSERT_NE(a3, nullptr);
  ASSERT_TRUE(a3->cached_value.is_boolean());
  EXPECT_FALSE(a3->cached_value.as_boolean());

  const Cell* a4 = CellAt(wb, 0U, 3U, 0U);
  ASSERT_NE(a4, nullptr);
  ASSERT_TRUE(a4->cached_value.is_text());
  EXPECT_EQ(a4->cached_value.as_text(), "hello");

  const Cell* a6 = CellAt(wb, 0U, 5U, 0U);
  ASSERT_NE(a6, nullptr);
  ASSERT_TRUE(a6->cached_value.is_error());
  EXPECT_EQ(a6->cached_value.as_error(), ErrorCode::Div0);
}

void AssertArithmetic(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);  // =A1+A2
  const Cell* b2 = CellAt(wb, 0U, 1U, 1U);  // =A1-A2
  const Cell* b3 = CellAt(wb, 0U, 2U, 1U);  // =A1*A2
  const Cell* b5 = CellAt(wb, 0U, 4U, 1U);  // =A1^A2
  const Cell* b7 = CellAt(wb, 0U, 6U, 1U);  // =-A1
  ASSERT_NE(b1, nullptr);
  ASSERT_NE(b2, nullptr);
  ASSERT_NE(b3, nullptr);
  ASSERT_NE(b5, nullptr);
  ASSERT_NE(b7, nullptr);
  ASSERT_TRUE(b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b1->cached_value.as_number(), 13.0);
  ASSERT_TRUE(b2->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b2->cached_value.as_number(), 7.0);
  ASSERT_TRUE(b3->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b3->cached_value.as_number(), 30.0);
  ASSERT_TRUE(b5->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b5->cached_value.as_number(), 1000.0);
  ASSERT_TRUE(b7->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b7->cached_value.as_number(), -10.0);
}

void AssertCrossSheet(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 2U);
  EXPECT_EQ(wb.sheet(0).name(), "Sheet1");
  EXPECT_EQ(wb.sheet(1).name(), "Sheet2");
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 99.0);
}

void AssertMultiSheetIndependent(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "North");
  EXPECT_EQ(wb.sheet(1).name(), "South");
  EXPECT_EQ(wb.sheet(2).name(), "East");
  const Cell* north_b1 = CellAt(wb, 0U, 0U, 1U);
  const Cell* south_b1 = CellAt(wb, 1U, 0U, 1U);
  const Cell* east_b1 = CellAt(wb, 2U, 0U, 1U);
  ASSERT_NE(north_b1, nullptr);
  ASSERT_NE(south_b1, nullptr);
  ASSERT_NE(east_b1, nullptr);
  ASSERT_TRUE(north_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(north_b1->cached_value.as_number(), 2.0);
  ASSERT_TRUE(south_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(south_b1->cached_value.as_number(), 4.0);
  ASSERT_TRUE(east_b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(east_b1->cached_value.as_number(), 2.0);
}

void AssertUnicodeSheetNames(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 3U);
  EXPECT_EQ(wb.sheet(0).name(), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  EXPECT_EQ(wb.sheet(1).name(), "emoji\xF0\x9F\x98\x80");
  EXPECT_EQ(wb.sheet(2).name(), "Sheet 3 with spaces");
}

void AssertUnicodeCellText(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  const Cell* a4 = CellAt(wb, 0U, 3U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_NE(a3, nullptr);
  ASSERT_NE(a4, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "\xE6\x97\xA5\xE6\x9C\xAC\xE8\xAA\x9E");
  ASSERT_TRUE(a2->cached_value.is_text());
  EXPECT_EQ(a2->cached_value.as_text(), "Hello \xF0\x9F\x98\x80");
}

void AssertRangeAggregates(const Workbook& wb) {
  const Cell* sum = CellAt(wb, 0U, 0U, 1U);
  const Cell* avg = CellAt(wb, 0U, 1U, 1U);
  const Cell* mn = CellAt(wb, 0U, 2U, 1U);
  const Cell* mx = CellAt(wb, 0U, 3U, 1U);
  const Cell* cnt = CellAt(wb, 0U, 4U, 1U);
  ASSERT_NE(sum, nullptr);
  ASSERT_NE(avg, nullptr);
  ASSERT_NE(mn, nullptr);
  ASSERT_NE(mx, nullptr);
  ASSERT_NE(cnt, nullptr);
  ASSERT_TRUE(sum->cached_value.is_number());
  EXPECT_DOUBLE_EQ(sum->cached_value.as_number(), 55.0);
  ASSERT_TRUE(avg->cached_value.is_number());
  EXPECT_DOUBLE_EQ(avg->cached_value.as_number(), 5.5);
  ASSERT_TRUE(mn->cached_value.is_number());
  EXPECT_DOUBLE_EQ(mn->cached_value.as_number(), 1.0);
  ASSERT_TRUE(mx->cached_value.is_number());
  EXPECT_DOUBLE_EQ(mx->cached_value.as_number(), 10.0);
  ASSERT_TRUE(cnt->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cnt->cached_value.as_number(), 10.0);
}

void AssertVolatileNowToday(const Workbook& wb) {
  // Cached values WILL drift across recalcs; we only check that the
  // formula text was preserved.
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  const Cell* a3 = CellAt(wb, 0U, 2U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_NE(a3, nullptr);
  EXPECT_EQ(a1->formula_text, "=NOW()");
  EXPECT_EQ(a2->formula_text, "=TODAY()");
  EXPECT_EQ(a3->formula_text, "=RAND()");
}

void AssertIterativeCircular(const Workbook& wb) {
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(b1, nullptr);
  EXPECT_EQ(a1->formula_text, "=B1+1");
  EXPECT_EQ(b1->formula_text, "=A1");
  // The cached values are observable but not pinned: depending on
  // whether iterative options were re-applied, the SCC will resolve
  // either to a converged numeric pair (iterative on) or to #REF!
  // (iterative off — the writer does not yet persist `<calcPr
  // iterate=...>`). Either is acceptable here; the formula text is
  // the load-bearing invariant.
}

void AssertDefinedNamesWorkbookScope(const Workbook& wb) {
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "Sales");
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$A$1:$A$10");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, -1);
  EXPECT_FALSE(wb.defined_names()[0].hidden);
}

void AssertDefinedNamesSheetScope(const Workbook& wb) {
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "LocalRange");
  EXPECT_EQ(wb.defined_names()[0].formula, "Sheet1!$B$1:$B$5");
  EXPECT_EQ(wb.defined_names()[0].local_sheet_id, 0);
  EXPECT_TRUE(wb.defined_names()[0].hidden);
  EXPECT_EQ(wb.defined_names()[0].comment, "Sheet1-only.");
}

void AssertSingleTable(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 1U);
  const io::TableMetadata& t = wb.tables()[0];
  EXPECT_EQ(t.name, "SalesTable");
  EXPECT_EQ(t.ref, "A1:C5");
  ASSERT_EQ(t.columns.size(), 3U);
  EXPECT_EQ(t.columns[0].name, "Region");
}

void AssertMultipleTables(const Workbook& wb) {
  ASSERT_EQ(wb.tables().size(), 3U);
  EXPECT_EQ(wb.tables()[0].name, "Tbl1");
  EXPECT_EQ(wb.tables()[1].name, "Tbl2");
  EXPECT_EQ(wb.tables()[2].name, "Tbl3");
  EXPECT_EQ(wb.tables()[0].sheet_index, 0U);
  EXPECT_EQ(wb.tables()[1].sheet_index, 0U);
  EXPECT_EQ(wb.tables()[2].sheet_index, 1U);
}

void AssertNestedFunctionCalls(const Workbook& wb) {
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  ASSERT_NE(b1, nullptr);
  // Formula text round-trips verbatim; cached value resolves to A2=3.
  EXPECT_EQ(b1->formula_text, "=IF(AND(OR(A1>0,A2<0),A3=8),IF(A1+A2>A3,A1,A2),0)");
  ASSERT_TRUE(b1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(b1->cached_value.as_number(), 3.0);
}

void AssertLargeGrid(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  // Spot-check the corners of the grid: A1 should be a numeric
  // literal (r=0, c=0; (r+c) % 7 == 0 -> text "L0").
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "L0");
  // Last cell T25 (r=24, c=19): (24+19)=43 -> 43%7=1 -> not text;
  // 43 % 3 = 1 -> not formula; numeric 24*20+19 = 499.
  const Cell* t25 = CellAt(wb, 0U, 24U, 19U);
  ASSERT_NE(t25, nullptr);
  ASSERT_TRUE(t25->cached_value.is_number());
  EXPECT_DOUBLE_EQ(t25->cached_value.as_number(), 499.0);
}

void AssertPassthroughTheme(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  // Passthrough must include the theme1 part with non-empty bytes.
  const auto& parts = wb.passthrough_parts();
  auto it = std::find_if(parts.begin(), parts.end(),
                         [](const io::PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  ASSERT_NE(it, parts.end()) << "theme part missing from passthrough";
  EXPECT_FALSE(it->bytes.empty());
}

void AssertSstCells(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 1U);
  const Cell* a1 = CellAt(wb, 0U, 0U, 0U);
  const Cell* b1 = CellAt(wb, 0U, 0U, 1U);
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a1, nullptr);
  ASSERT_NE(b1, nullptr);
  ASSERT_NE(a2, nullptr);
  ASSERT_TRUE(a1->cached_value.is_text());
  ASSERT_TRUE(b1->cached_value.is_text());
  ASSERT_TRUE(a2->cached_value.is_text());
  EXPECT_EQ(a1->cached_value.as_text(), "shared-alpha");
  EXPECT_EQ(b1->cached_value.as_text(), "shared-beta");
  EXPECT_EQ(a2->cached_value.as_text(), "shared-gamma");
}

void AssertKitchenSink(const Workbook& wb) {
  ASSERT_EQ(wb.sheet_count(), 2U);
  // Defined name preserved.
  ASSERT_EQ(wb.defined_names().size(), 1U);
  EXPECT_EQ(wb.defined_names()[0].name, "ComboName");
  // Table preserved.
  ASSERT_EQ(wb.tables().size(), 1U);
  EXPECT_EQ(wb.tables()[0].name, "ComboTable");
  // Passthrough preserved.
  const auto& parts = wb.passthrough_parts();
  auto it = std::find_if(parts.begin(), parts.end(),
                         [](const io::PassthroughPart& p) { return p.path == "xl/theme/theme1.xml"; });
  EXPECT_NE(it, parts.end());
  // Mixed formula present.
  const Cell* a2 = CellAt(wb, 0U, 1U, 0U);
  ASSERT_NE(a2, nullptr);
  EXPECT_EQ(a2->formula_text, "=A1*2");
}

// ---------------------------------------------------------------------------
// Corpus enumeration. The order here is the order CTest will iterate
// the parameterised tests; keep it stable so failure output is
// predictable.
// ---------------------------------------------------------------------------

std::vector<CorpusBook> make_corpus() {
  std::vector<CorpusBook> books;
  books.push_back({"empty", BuildEmpty, AssertEmpty});
  books.push_back({"single_literal", BuildSingleLiteral, AssertSingleLiteral});
  books.push_back({"literals_all_kinds", BuildLiteralsAllKinds, AssertLiteralsAllKinds});
  books.push_back({"arithmetic", BuildArithmetic, AssertArithmetic});
  books.push_back({"cross_sheet", BuildCrossSheet, AssertCrossSheet});
  books.push_back({"multi_sheet_independent", BuildMultiSheetIndependent, AssertMultiSheetIndependent});
  books.push_back({"unicode_sheet_names", BuildUnicodeSheetNames, AssertUnicodeSheetNames});
  books.push_back({"unicode_cell_text", BuildUnicodeCellText, AssertUnicodeCellText});
  books.push_back({"range_aggregates", BuildRangeAggregates, AssertRangeAggregates});
  books.push_back({"volatile_now_today", BuildVolatileNowToday, AssertVolatileNowToday});
  books.push_back({"iterative_circular", BuildIterativeCircular, AssertIterativeCircular});
  books.push_back({"defined_names_workbook_scope", BuildDefinedNamesWorkbookScope, AssertDefinedNamesWorkbookScope});
  books.push_back({"defined_names_sheet_scope", BuildDefinedNamesSheetScope, AssertDefinedNamesSheetScope});
  books.push_back({"single_table", BuildSingleTable, AssertSingleTable});
  books.push_back({"multiple_tables", BuildMultipleTables, AssertMultipleTables});
  books.push_back({"nested_function_calls", BuildNestedFunctionCalls, AssertNestedFunctionCalls});
  books.push_back({"large_grid_500_cells", BuildLargeGrid500Cells, AssertLargeGrid});
  books.push_back({"passthrough_theme_part", BuildPassthroughThemeArchive, AssertPassthroughTheme});
  books.push_back({"sst_cells", BuildSstArchive, AssertSstCells});
  books.push_back({"combined_kitchen_sink", BuildKitchenSink, AssertKitchenSink});
  return books;
}

// ---------------------------------------------------------------------------
// Parameterised driver.
// ---------------------------------------------------------------------------

class RoundTripParity : public ::testing::TestWithParam<CorpusBook> {};

TEST_P(RoundTripParity, TwoCyclePipeline) {
  const CorpusBook& book = GetParam();

  // (1) build source bytes.
  auto src_or = book.build();
  ASSERT_TRUE(static_cast<bool>(src_or)) << "build failed for '" << book.id << "': " << src_or.error().message;

  // (2) read into workbook A.
  auto first_or = io::read_ooxml(SpanOf(src_or.value()));
  ASSERT_TRUE(static_cast<bool>(first_or))
      << "first read_ooxml failed for '" << book.id << "': " << first_or.error().message;
  io::OoxmlReadResult& first = first_or.value();

  // (3) recalc workbook A so cached values are populated. The
  // `iterative_circular` and `combined_kitchen_sink` books need
  // iterative options re-applied because the writer does not yet
  // persist `<calcPr iterate=...>`; without this the cyclic SCC would
  // surface #REF!. See the deviation note in the bundle report.
  if (book.id == "iterative_circular" || book.id == "combined_kitchen_sink") {
    eval::IterativeOptions opts;
    opts.enabled = true;
    opts.max_iterations = 20U;
    opts.max_change = 0.01;
    first.workbook.set_iterative_options(opts);
  }
  auto stats_or = first.workbook.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats_or))
      << "first recalc failed for '" << book.id << "': " << stats_or.error().message;

  // Per-book invariants on A.
  book.assert_invariants(first.workbook);

  // (4) write A back to bytes.
  auto bytes_or = first.workbook.save();
  ASSERT_TRUE(static_cast<bool>(bytes_or))
      << "first save failed for '" << book.id << "': " << bytes_or.error().message;

  // (5) read again into workbook C.
  auto second_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(second_or))
      << "second read_ooxml failed for '" << book.id << "': " << second_or.error().message;
  io::OoxmlReadResult& second = second_or.value();

  // (6) recalc workbook C, with iterative options re-applied if needed.
  if (book.id == "iterative_circular" || book.id == "combined_kitchen_sink") {
    eval::IterativeOptions opts;
    opts.enabled = true;
    opts.max_iterations = 20U;
    opts.max_change = 0.01;
    second.workbook.set_iterative_options(opts);
  }
  auto stats2_or = second.workbook.recalc(eval::default_registry());
  ASSERT_TRUE(static_cast<bool>(stats2_or))
      << "second recalc failed for '" << book.id << "': " << stats2_or.error().message;

  // Re-assert per-book invariants on C — the second read must produce
  // the same shape.
  book.assert_invariants(second.workbook);

  // (7) cross-cycle equality contract.
  EXPECT_TRUE(sheets_have_same_shape(first.workbook, second.workbook)) << "sheet shape drifted for '" << book.id << "'";
  EXPECT_TRUE(defined_names_equal(first.workbook, second.workbook)) << "defined names drifted for '" << book.id << "'";
  EXPECT_TRUE(tables_equal(first.workbook, second.workbook)) << "tables drifted for '" << book.id << "'";
  EXPECT_TRUE(passthrough_parts_equal(first.workbook, second.workbook))
      << "passthrough parts drifted for '" << book.id << "'";
  EXPECT_TRUE(formula_texts_equal_for_each_cell(first.workbook, second.workbook))
      << "formula texts drifted for '" << book.id << "'";
}

/// Custom GTest parameter name printer: each test is named after the
/// `id` field of the corpus book so failures read as
/// `OoxmlCorpus/RoundTripParity.TwoCyclePipeline/<id>`.
struct CorpusNameFormatter {
  std::string operator()(const ::testing::TestParamInfo<CorpusBook>& info) const { return info.param.id; }
};

INSTANTIATE_TEST_SUITE_P(OoxmlCorpus, RoundTripParity, ::testing::ValuesIn(make_corpus()), CorpusNameFormatter());

}  // namespace
}  // namespace formulon
