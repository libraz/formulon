#pragma once

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/styles_reader.h"
#include "io/styles_writer.h"
#include "io/zip_reader.h"
#include "workbook.h"

namespace {

// `fm_styles_add_cell_xf` takes this struct BY VALUE, so its size is part of
// the calling convention rather than merely of a buffer: a caller compiled
// against a different definition reads its arguments from the wrong registers
// and stack slots with no diagnosable failure. Pin the layout here so a change
// has to break the build and be acknowledged.
static_assert(sizeof(fm_cell_xf) == 128U, "fm_cell_xf ABI layout changed");
static_assert(offsetof(fm_cell_xf, font_index) == 0U, "fm_cell_xf.font_index offset changed");
static_assert(offsetof(fm_cell_xf, fill_index) == 4U, "fm_cell_xf.fill_index offset changed");
static_assert(offsetof(fm_cell_xf, border_index) == 8U, "fm_cell_xf.border_index offset changed");
static_assert(offsetof(fm_cell_xf, num_fmt_id) == 12U, "fm_cell_xf.num_fmt_id offset changed");
static_assert(offsetof(fm_cell_xf, horizontal_align) == 14U, "fm_cell_xf.horizontal_align offset changed");
static_assert(offsetof(fm_cell_xf, vertical_align) == 15U, "fm_cell_xf.vertical_align offset changed");
static_assert(offsetof(fm_cell_xf, wrap_text) == 16U, "fm_cell_xf.wrap_text offset changed");
static_assert(offsetof(fm_cell_xf, justify_last_line) == 20U, "fm_cell_xf.justify_last_line offset changed");
static_assert(offsetof(fm_cell_xf, xf_id) == 24U, "fm_cell_xf.xf_id offset changed");
static_assert(offsetof(fm_cell_xf, has_alignment) == 28U, "fm_cell_xf.has_alignment offset changed");
static_assert(offsetof(fm_cell_xf, has_text_rotation) == 32U, "fm_cell_xf.has_text_rotation offset changed");
static_assert(offsetof(fm_cell_xf, reading_order) == 68U, "fm_cell_xf.reading_order offset changed");
static_assert(offsetof(fm_cell_xf, has_horizontal_align) == 72U, "fm_cell_xf.has_horizontal_align offset changed");
static_assert(offsetof(fm_cell_xf, has_vertical_align) == 76U, "fm_cell_xf.has_vertical_align offset changed");
static_assert(offsetof(fm_cell_xf, has_wrap_text) == 80U, "fm_cell_xf.has_wrap_text offset changed");
static_assert(offsetof(fm_cell_xf, has_justify_last_line) == 84U, "fm_cell_xf.has_justify_last_line offset changed");
static_assert(offsetof(fm_cell_xf, apply_number_format) == 88U, "fm_cell_xf.apply_number_format offset changed");
static_assert(offsetof(fm_cell_xf, quote_prefix) == 112U, "fm_cell_xf.quote_prefix offset changed");
static_assert(offsetof(fm_cell_xf, has_protection) == 116U, "fm_cell_xf.has_protection offset changed");
static_assert(offsetof(fm_cell_xf, hidden) == 124U, "fm_cell_xf.hidden offset changed");
static_assert(sizeof(fm_dxf_record) == (sizeof(void*) == 4U ? 368U : 376U), "fm_dxf_record ABI layout changed");
static_assert(offsetof(fm_dxf_record, num_fmt_code) == 352U, "fm_dxf_record.num_fmt_code offset changed");
static_assert(offsetof(fm_dxf_record, alignment_xml) == (sizeof(void*) == 4U ? 356U : 360U),
              "fm_dxf_record.alignment_xml offset changed");
static_assert(offsetof(fm_dxf_record, protection_xml) == (sizeof(void*) == 4U ? 360U : 368U),
              "fm_dxf_record.protection_xml offset changed");

// `fm_styles_batch` is fifteen pointer-width slots: five (array, count,
// out-indices) triples whose `size_t` counts are the same width as a pointer on
// both targets. Unlike its neighbours it is reached by no binding -- only the C
// and C++ callers of `fm_styles_add_batch` -- so nothing cross-checks it from
// the other side of the boundary and every offset is pinned here instead. A
// reorder within the struct moves no total size, which is exactly why the
// per-member offsets and not just `sizeof` are recorded.
static_assert(sizeof(fm_styles_batch) == 15U * sizeof(void*), "fm_styles_batch ABI layout changed");
static_assert(alignof(fm_styles_batch) == alignof(void*), "fm_styles_batch ABI alignment changed");
static_assert(offsetof(fm_styles_batch, fonts) == 0U * sizeof(void*), "fm_styles_batch.fonts offset changed");
static_assert(offsetof(fm_styles_batch, font_count) == 1U * sizeof(void*), "fm_styles_batch.font_count offset changed");
static_assert(offsetof(fm_styles_batch, font_indices) == 2U * sizeof(void*),
              "fm_styles_batch.font_indices offset changed");
static_assert(offsetof(fm_styles_batch, fills) == 3U * sizeof(void*), "fm_styles_batch.fills offset changed");
static_assert(offsetof(fm_styles_batch, fill_count) == 4U * sizeof(void*), "fm_styles_batch.fill_count offset changed");
static_assert(offsetof(fm_styles_batch, fill_indices) == 5U * sizeof(void*),
              "fm_styles_batch.fill_indices offset changed");
static_assert(offsetof(fm_styles_batch, borders) == 6U * sizeof(void*), "fm_styles_batch.borders offset changed");
static_assert(offsetof(fm_styles_batch, border_count) == 7U * sizeof(void*),
              "fm_styles_batch.border_count offset changed");
static_assert(offsetof(fm_styles_batch, border_indices) == 8U * sizeof(void*),
              "fm_styles_batch.border_indices offset changed");
static_assert(offsetof(fm_styles_batch, cell_xfs) == 9U * sizeof(void*), "fm_styles_batch.cell_xfs offset changed");
static_assert(offsetof(fm_styles_batch, cell_xf_count) == 10U * sizeof(void*),
              "fm_styles_batch.cell_xf_count offset changed");
static_assert(offsetof(fm_styles_batch, cell_xf_indices) == 11U * sizeof(void*),
              "fm_styles_batch.cell_xf_indices offset changed");
static_assert(offsetof(fm_styles_batch, num_fmt_codes) == 12U * sizeof(void*),
              "fm_styles_batch.num_fmt_codes offset changed");
static_assert(offsetof(fm_styles_batch, num_fmt_count) == 13U * sizeof(void*),
              "fm_styles_batch.num_fmt_count offset changed");
static_assert(offsetof(fm_styles_batch, num_fmt_ids) == 14U * sizeof(void*),
              "fm_styles_batch.num_fmt_ids offset changed");
// The slots are pointer-width because `size_t` is, on every target this ABI
// ships to. Pinned separately so a target where that stops holding fails here
// rather than silently re-laying the struct out.
static_assert(sizeof(size_t) == sizeof(void*), "fm_styles_batch assumes size_t is pointer-width");

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct BufferGuard {
  uint8_t* data = nullptr;
  size_t len = 0;
  ~BufferGuard() { fm_buffer_free(data); }
  BufferGuard() = default;
  BufferGuard(const BufferGuard&) = delete;
  BufferGuard& operator=(const BufferGuard&) = delete;
};

[[maybe_unused]] constexpr char kNumFmtOverrideStyles[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<numFmts count=\"1\"><numFmt numFmtId=\"14\" formatCode=\"yyyy\"/></numFmts>"
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"gray125\"/></fill></fills>"
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>"
    "</styleSheet>";

[[maybe_unused]] constexpr char kNumFmtDuplicateStyles[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<numFmts count=\"2\"><numFmt numFmtId=\"14\" formatCode=\"yyyy\"/>"
    "<numFmt numFmtId=\"14\" formatCode=\"dd/mm/yyyy\"/></numFmts>"
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"gray125\"/></fill></fills>"
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>"
    "</styleSheet>";

[[maybe_unused]] inline void LoadParsedStyles(fm_workbook_t* handle, const char* xml) {
  const std::vector<std::uint8_t> bytes(xml, xml + std::strlen(xml));
  auto parsed = formulon::io::read_styles(bytes);
  ASSERT_TRUE(parsed.has_value());
  handle->workbook().mutable_styles() = std::move(parsed.value());
}

[[maybe_unused]] inline void LoadNumFmtOverrideStyles(fm_workbook_t* handle) {
  LoadParsedStyles(handle, kNumFmtOverrideStyles);
}

[[maybe_unused]] inline void LoadNumFmtDuplicateStyles(fm_workbook_t* handle) {
  LoadParsedStyles(handle, kNumFmtDuplicateStyles);
}

[[maybe_unused]] inline std::string ExtractStylesXml(const BufferGuard& saved) {
  formulon::io::ZipReader zip;
  const auto opened = zip.open(formulon::io::ByteSpan{saved.data, saved.len});
  if (!opened) {
    ADD_FAILURE() << opened.error().message;
    return {};
  }
  auto styles_part = zip.read_entry("xl/styles.xml");
  if (!styles_part) {
    ADD_FAILURE() << styles_part.error().message;
    return {};
  }
  return {styles_part.value().begin(), styles_part.value().end()};
}

[[maybe_unused]] inline fm_font_record MakeArial() {
  fm_font_record r{};
  r.name = "Arial";
  r.size = 12.0;
  r.color_argb = 0xFF112233U;
  r.bold = 1;
  r.italic = 0;
  r.strike = 0;
  r.underline = 0;
  return r;
}

[[maybe_unused]] inline fm_fill_record MakeRedFill() {
  fm_fill_record r{};
  r.pattern = 1;  // solid
  r.fg_argb = 0xFFFF0000U;
  r.bg_argb = 0xFF000000U;
  return r;
}

[[maybe_unused]] inline fm_border_record MakeThinBoxBorder() {
  fm_border_record r{};
  r.left.style = 1;  // thin
  r.left.color_argb = 0xFF000000U;
  r.right = r.left;
  r.top = r.left;
  r.bottom = r.left;
  r.diagonal.style = 0;
  r.diagonal.color_argb = 0;
  r.diagonal_up = 0;
  r.diagonal_down = 0;
  return r;
}

// An Excel-authored `<fonts>` section: the theme colour, `<family>` and
// `<charset>` children are what a template carries and what a record built
// from scratch through the C ABI cannot reproduce unless the ABI exposes
// them.
[[maybe_unused]] constexpr char kExcelAuthoredStyles[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<fonts count=\"3\">"
    "<font><sz val=\"11\"/><color theme=\"1\"/><name val=\"Calibri\"/><family val=\"2\"/>"
    "<charset val=\"128\"/></font>"
    "<font><b/><sz val=\"11\"/><color theme=\"0\" tint=\"-0.25\"/><name val=\"Calibri\"/>"
    "<family val=\"2\"/></font>"
    "<font><vertAlign val=\"superscript\"/><sz val=\"9\"/><color rgb=\"FF112233\"/>"
    "<name val=\"Calibri\"/></font>"
    "</fonts>"
    "<fills count=\"2\">"
    "<fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"solid\"><fgColor theme=\"4\" tint=\"0.5\"/>"
    "<bgColor indexed=\"64\"/></patternFill></fill>"
    "</fills>"
    "<borders count=\"2\">"
    "<border><left/><right/><top/><bottom/><diagonal/></border>"
    "<border><left style=\"thin\"><color theme=\"3\"/></left><right/><top/><bottom/>"
    "<diagonal/></border>"
    "</borders>"
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>"
    "</styleSheet>";

/// Installs `kExcelAuthoredStyles` as the workbook's styles table.
[[maybe_unused]] inline void LoadExcelAuthoredStyles(fm_workbook_t* handle) {
  const std::vector<std::uint8_t> bytes(kExcelAuthoredStyles, kExcelAuthoredStyles + std::strlen(kExcelAuthoredStyles));
  auto parsed = formulon::io::read_styles(bytes);
  ASSERT_TRUE(parsed.has_value());
  handle->workbook().mutable_styles() = std::move(parsed.value());
}

}  // namespace
