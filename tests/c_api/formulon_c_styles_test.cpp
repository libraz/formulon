//
// Stable C ABI tests for the styles surface
// (`fm_cell_get_xf_index`, `fm_cell_set_xf_index`,
// `fm_styles_get_cell_xf`, `fm_styles_get_font`,
// `fm_styles_get_num_fmt_string`).
//
// The test driver is C++ for gtest convenience but everything it
// touches across the boundary is the pure-C surface.

#include <cstdint>
#include <cstring>
#include <string>
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
static_assert(sizeof(fm_cell_xf) == 88U, "fm_cell_xf ABI layout changed");
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
static_assert(sizeof(fm_dxf_record) == (sizeof(void*) == 4U ? 360U : 368U), "fm_dxf_record ABI layout changed");
static_assert(offsetof(fm_dxf_record, num_fmt_code) == 344U, "fm_dxf_record.num_fmt_code offset changed");
static_assert(offsetof(fm_dxf_record, alignment_xml) == (sizeof(void*) == 4U ? 348U : 352U),
              "fm_dxf_record.alignment_xml offset changed");
static_assert(offsetof(fm_dxf_record, protection_xml) == (sizeof(void*) == 4U ? 352U : 360U),
              "fm_dxf_record.protection_xml offset changed");

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

constexpr char kNumFmtOverrideStyles[] =
    "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
    "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
    "<numFmts count=\"1\"><numFmt numFmtId=\"14\" formatCode=\"yyyy\"/></numFmts>"
    "<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>"
    "<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill>"
    "<fill><patternFill patternType=\"gray125\"/></fill></fills>"
    "<borders count=\"1\"><border><left/><right/><top/><bottom/><diagonal/></border></borders>"
    "<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellXfs>"
    "</styleSheet>";

constexpr char kNumFmtDuplicateStyles[] =
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

void LoadParsedStyles(fm_workbook_t* handle, const char* xml) {
  const std::vector<std::uint8_t> bytes(xml, xml + std::strlen(xml));
  auto parsed = formulon::io::read_styles(bytes);
  ASSERT_TRUE(parsed.has_value());
  handle->workbook().mutable_styles() = std::move(parsed.value());
}

void LoadNumFmtOverrideStyles(fm_workbook_t* handle) {
  LoadParsedStyles(handle, kNumFmtOverrideStyles);
}

void LoadNumFmtDuplicateStyles(fm_workbook_t* handle) {
  LoadParsedStyles(handle, kNumFmtDuplicateStyles);
}

std::string ExtractStylesXml(const BufferGuard& saved) {
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

}  // namespace

TEST(FormulonCApiStyles, CellXfIndexDefaultsToZero) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t xf = 99;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 0U);
}

TEST(FormulonCApiStyles, CellXfIndexRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 5, 7, 42), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb.handle, 0, 5, 7, &xf), 0);
  EXPECT_EQ(xf, 42U);
}

TEST(FormulonCApiStyles, SetXfIndexRejectsBadSheet) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // sheet 99 does not exist.
  EXPECT_NE(fm_cell_set_xf_index(wb.handle, 99, 0, 0, 1), 0);
}

TEST(FormulonCApiStyles, NullArgumentRejected) {
  uint32_t xf = 0;
  EXPECT_NE(fm_cell_get_xf_index(nullptr, 0, 0, 0, &xf), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_cell_get_xf_index(wb.handle, 0, 0, 0, nullptr), 0);
}

TEST(FormulonCApiStyles, BuiltinNumFmtResolves) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 0, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "General");
  EXPECT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "mm-dd-yy");
}

TEST(FormulonCApiStyles, CustomNumFmtOverridingBuiltinIdWins) {
  // A file may define a custom <numFmt> whose numFmtId collides with a
  // built-in slot; Excel honours the file's definition. Parse a production-
  // shaped styles.xml and install the resulting table rather than manually
  // constructing the record.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "yyyy");
}

TEST(FormulonCApiStyles, AddNumFmtUsesEffectiveBuiltinMapping) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint16_t override_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &override_id), 0);
  EXPECT_EQ(override_id, 14U);

  uint16_t builtin_code_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_code_id), 0);
  EXPECT_GE(builtin_code_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, builtin_code_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t builtin_code_again = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_code_again), 0);
  EXPECT_EQ(builtin_code_again, builtin_code_id);
}

TEST(FormulonCApiStyles, AddBatchNumFmtUsesEffectiveBuiltinMapping) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  const char* const codes[] = {"General", "yyyy", "mm-dd-yy", "mm-dd-yy"};
  uint16_t ids[] = {0xFFFFU, 0xFFFFU, 0xFFFFU, 0xFFFFU};
  fm_styles_batch batch{};
  batch.num_fmt_codes = codes;
  batch.num_fmt_count = 4U;
  batch.num_fmt_ids = ids;

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(ids[0], 0U);
  EXPECT_EQ(ids[1], 14U);
  EXPECT_GE(ids[2], 164U);
  EXPECT_EQ(ids[3], ids[2]);
  for (size_t i = 0; i < 4U; ++i) {
    const char* resolved = nullptr;
    ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, ids[i], &resolved), 0);
    ASSERT_NE(resolved, nullptr);
    EXPECT_STREQ(resolved, codes[i]);
  }
}

TEST(FormulonCApiStyles, InvalidCustomNumFmtStringIndexDoesNotShadowBuiltin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtOverrideStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 1U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &id), 0);
  EXPECT_EQ(id, 14U);
}

TEST(FormulonCApiStyles, DuplicateNumFmtFirstValidRecordWins) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");

  uint16_t first_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &first_id), 0);
  EXPECT_EQ(first_id, 14U);

  uint16_t shadowed_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "dd/mm/yyyy", &shadowed_id), 0);
  EXPECT_GE(shadowed_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, shadowed_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtInvalidFirstThenValidRecordWins) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  styles.num_fmts[0].format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "dd/mm/yyyy");

  uint16_t valid_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "dd/mm/yyyy", &valid_id), 0);
  EXPECT_EQ(valid_id, 14U);

  uint16_t invalid_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "yyyy", &invalid_id), 0);
  EXPECT_GE(invalid_id, 164U);
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, invalid_id, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "yyyy");
}

TEST(FormulonCApiStyles, DuplicateNumFmtAllInvalidRecordsFallBackToBuiltin) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadNumFmtDuplicateStyles(wb.handle));

  auto& styles = wb.handle->workbook().mutable_styles();
  ASSERT_EQ(styles.num_fmts.size(), 2U);
  const std::uint32_t invalid_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  for (auto& record : styles.num_fmts) {
    record.format_string_index = invalid_index;
  }

  const char* resolved = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, 14U, &resolved), 0);
  ASSERT_NE(resolved, nullptr);
  EXPECT_STREQ(resolved, "mm-dd-yy");

  uint16_t builtin_id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "mm-dd-yy", &builtin_id), 0);
  EXPECT_EQ(builtin_id, 14U);
}

TEST(FormulonCApiStyles, UnknownNumFmtIdRejected) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* s = nullptr;
  // Reserved built-in slot (id 5) without a custom override surfaces
  // an error rather than an empty string.
  EXPECT_NE(fm_styles_get_num_fmt_string(wb.handle, 5, &s), 0);
}

TEST(FormulonCApiStyles, DifferentialFormatGetterExposesCfDxfTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::StylesTable& styles = wb.handle->workbook().mutable_styles();
  formulon::io::DifferentialFormat dxf;
  dxf.has_font = true;
  dxf.font.bold = true;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.has_fill = true;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.has_border = true;
  dxf.border.left.style = 1;
  dxf.border.left.color_argb = 0xFF000000U;
  dxf.has_num_fmt = true;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";
  styles.dxfs.push_back(dxf);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, 0, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.border_engaged, 1);
  EXPECT_EQ(out.border.left.style, 1U);
  EXPECT_EQ(out.border.left.color_argb, 0xFF000000U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");

  EXPECT_NE(fm_styles_get_dxf(wb.handle, 1, &out), 0);
}

TEST(FormulonCApiStyles, AddDxfDedupsAndReadsBack) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font.name = "Arial";
  dxf.font.size = 12.0;
  dxf.font.bold = 1;
  dxf.font.color_argb = 0xFFFF0000U;
  dxf.fill_engaged = 1;
  dxf.fill.pattern = 1;
  dxf.fill.fg_argb = 0xFFFFFF00U;
  dxf.num_fmt_engaged = 1;
  dxf.num_fmt_id = 164;
  dxf.num_fmt_code = "0.00";

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &a), 0);
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &b), 0);
  EXPECT_EQ(a, b);

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  fm_dxf_record out{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, a, &out), 0);
  EXPECT_EQ(out.font_engaged, 1);
  EXPECT_STREQ(out.font.name, "Arial");
  EXPECT_EQ(out.font.bold, 1);
  EXPECT_EQ(out.font.color_argb, 0xFFFF0000U);
  EXPECT_EQ(out.fill_engaged, 1);
  EXPECT_EQ(out.fill.pattern, 1U);
  EXPECT_EQ(out.fill.fg_argb, 0xFFFFFF00U);
  EXPECT_EQ(out.num_fmt_engaged, 1);
  EXPECT_EQ(out.num_fmt_id, 164U);
  ASSERT_NE(out.num_fmt_code, nullptr);
  EXPECT_STREQ(out.num_fmt_code, "0.00");
}

TEST(FormulonCApiStyles, DxfFontRoundTripsVerticalAlignmentThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font.name = "Calibri";
  dxf.font.size = 9.0;
  dxf.font.color_argb = 0xFF112233U;
  dxf.font.vert_align = 1;  // superscript
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &index), 0);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<vertAlign val=\"superscript\"/>"), std::string::npos);

  fm_dxf_record loaded{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, index, &loaded), 0);
  EXPECT_EQ(loaded.font.vert_align, 1U);

  uint32_t count_before = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count_before), 0);
  uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, loaded, &reindex), 0);
  EXPECT_EQ(reindex, index);
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_dxf_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, DxfAlignmentAndProtectionPreserveIdentityThroughOoxml) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  std::string alignment_input = "<alignment horizontal=\"center\" wrapText=\"1\"/>";
  std::string protection_input = "<protection locked=\"0\" hidden=\"1\"/>";
  fm_dxf_record alignment{};
  alignment.alignment_xml = alignment_input.c_str();
  fm_dxf_record protection{};
  protection.protection_xml = protection_input.c_str();

  uint32_t alignment_index = 0xFFFFFFFFU;
  uint32_t protection_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, alignment, &alignment_index), 0);
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, protection, &protection_index), 0);
  EXPECT_NE(alignment_index, protection_index);
  alignment_input = "caller-owned alignment text replaced";
  protection_input = "caller-owned protection text replaced";

  fm_dxf_record got_alignment{};
  fm_dxf_record got_protection{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, alignment_index, &got_alignment), 0);
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, protection_index, &got_protection), 0);
  ASSERT_NE(got_alignment.alignment_xml, nullptr);
  ASSERT_NE(got_alignment.protection_xml, nullptr);
  ASSERT_NE(got_protection.alignment_xml, nullptr);
  ASSERT_NE(got_protection.protection_xml, nullptr);
  EXPECT_STREQ(got_alignment.alignment_xml, "<alignment horizontal=\"center\" wrapText=\"1\"/>");
  EXPECT_STREQ(got_alignment.protection_xml, "");
  EXPECT_STREQ(got_protection.alignment_xml, "");
  EXPECT_STREQ(got_protection.protection_xml, "<protection locked=\"0\" hidden=\"1\"/>");

  uint32_t alignment_again = 0xFFFFFFFFU;
  uint32_t protection_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, got_alignment, &alignment_again), 0);
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, got_protection, &protection_again), 0);
  EXPECT_EQ(alignment_again, alignment_index);
  EXPECT_EQ(protection_again, protection_index);

  const std::string before_save = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(before_save.find("<alignment horizontal=\"center\" wrapText=\"1\"/>"), std::string::npos);
  EXPECT_NE(before_save.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_dxf_record reloaded_alignment{};
  fm_dxf_record reloaded_protection{};
  ASSERT_EQ(fm_styles_get_dxf(reloaded.handle, alignment_index, &reloaded_alignment), 0);
  ASSERT_EQ(fm_styles_get_dxf(reloaded.handle, protection_index, &reloaded_protection), 0);
  EXPECT_STREQ(reloaded_alignment.alignment_xml, "<alignment horizontal=\"center\" wrapText=\"1\"/>");
  EXPECT_STREQ(reloaded_alignment.protection_xml, "");
  EXPECT_STREQ(reloaded_protection.alignment_xml, "");
  EXPECT_STREQ(reloaded_protection.protection_xml, "<protection locked=\"0\" hidden=\"1\"/>");

  uint32_t reloaded_alignment_again = 0xFFFFFFFFU;
  uint32_t reloaded_protection_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(reloaded.handle, reloaded_alignment, &reloaded_alignment_again), 0);
  ASSERT_EQ(fm_styles_add_dxf(reloaded.handle, reloaded_protection, &reloaded_protection_again), 0);
  EXPECT_EQ(reloaded_alignment_again, alignment_index);
  EXPECT_EQ(reloaded_protection_again, protection_index);

  const std::string after_load = formulon::io::write_styles(reloaded.handle->workbook().styles());
  EXPECT_NE(after_load.find("<alignment horizontal=\"center\" wrapText=\"1\"/>"), std::string::npos);
  EXPECT_NE(after_load.find("<protection locked=\"0\" hidden=\"1\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, DxfFontDistinguishesAnExplicitBoldOffFromAnAbsentToggle) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  formulon::io::DifferentialFormat stored;
  stored.has_font = true;
  stored.font.has_bold = true;  // `<b val="0"/>`: switch bold off
  stored.font.bold = false;
  wb.handle->workbook().mutable_styles().dxfs.push_back(stored);

  // A rule that leaves bold alone must not fold onto the switch-it-off one.
  fm_dxf_record leave_bold_alone{};
  leave_bold_alone.font_engaged = 1;
  leave_bold_alone.font.name = "";
  leave_bold_alone.font.size = 11.0;
  leave_bold_alone.font.color_argb = 0xFF000000U;
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, leave_bold_alone, &index), 0);
  EXPECT_NE(index, 0U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<b val=\"0\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, GetCellXfRejectsOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  // A fresh workbook has an empty styles table, so any xf_index is
  // out of range. Calling with index 0 still fails because the table
  // has zero entries.
  EXPECT_NE(fm_styles_get_cell_xf(wb.handle, 0, &xf), 0);
}

TEST(FormulonCApiStyles, CellXfGetterDiagnosticsUseInvokedApi) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  fm_cell_xf cell_xf{};
  expect_prefix(fm_styles_get_cell_xf(wb.handle, 0, &cell_xf), "fm_styles_get_cell_xf:");
  expect_prefix(fm_styles_get_cell_style_xf(wb.handle, 0, &cell_xf), "fm_styles_get_cell_style_xf:");
}

TEST(FormulonCApiStyles, CellXfAdderDiagnosticsUseInvokedApi) {
  auto expect_prefix = [](fm_status_t status, const char* prefix) {
    EXPECT_NE(status, 0);
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(prefix, 0), 0U) << message;
  };

  uint32_t index = 0;
  fm_cell_xf cell_xf{};
  expect_prefix(fm_styles_add_cell_xf(nullptr, cell_xf, &index), "fm_styles_add_cell_xf:");

  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  cell_xf.font_index = 1U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");

  cell_xf = fm_cell_xf{};
  cell_xf.xf_id = 1U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");

  cell_xf = fm_cell_xf{};
  cell_xf.has_text_rotation = 1;
  cell_xf.text_rotation = 181U;
  expect_prefix(fm_styles_add_cell_xf(wb.handle, cell_xf, &index), "fm_styles_add_cell_xf:");
}

TEST(FormulonCApiStyles, CellXfAlignmentEnumsValidateRangesBeforeMutation) {
  const auto expect_invalid = [](fm_status_t status, const char* api) {
    EXPECT_EQ(status, static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument));
    const char* message = fm_last_error_message();
    ASSERT_NE(message, nullptr);
    EXPECT_EQ(std::string(message).rfind(api, 0), 0U) << message;
  };

  // The upper boundary of each ordinal is accepted.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 7;  // distributed
    record.has_horizontal_align = 1;
    record.vertical_align = 4;  // distributed
    record.has_vertical_align = 1;
    uint32_t index = 0;
    ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
  }

  // The presence flag decides whether a value is read at all, so an
  // out-of-range ordinal on an omitted attribute is ignored rather than
  // rejected: the record it describes has no such attribute to be invalid.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.vertical_align = 5;
    uint32_t index = 0xAABBCCDDU;
    ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
    EXPECT_NE(index, 0xAABBCCDDU);
    fm_cell_xf got{};
    ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
    EXPECT_EQ(got.horizontal_align, 0U);
    EXPECT_EQ(got.vertical_align, 2U);
  }

  // Every add shape validates before it creates default roots or changes the
  // output index. Keep each case isolated so the size checks cover all paths.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.has_horizontal_align = 1;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_cell_xf(wb.handle, record, &index), "fm_styles_add_cell_xf:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.vertical_align = 5;
    record.has_vertical_align = 1;
    uint32_t index = 0xAABBCCDDU;
    const std::size_t before = wb.handle->workbook().styles().cell_style_xfs.size();
    expect_invalid(fm_styles_add_cell_style_xf(wb.handle, record, &index), "fm_styles_add_cell_style_xf:");
    EXPECT_EQ(index, 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_style_xfs.size(), before);
  }
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.horizontal_align = 8;
    record.has_horizontal_align = 1;
    uint32_t indices[] = {0xAABBCCDDU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      indices, nullptr, 0U,      nullptr};
    const std::size_t before = wb.handle->workbook().styles().cell_xfs.size();
    expect_invalid(fm_styles_add_batch(wb.handle, &batch), "fm_styles_add_batch:");
    EXPECT_EQ(indices[0], 0xAABBCCDDU);
    EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before);
  }
  // A batch record naming a `<cellStyleXfs>` entry that does not exist is
  // rejected rather than emitted as a dangling `xfId`.
  {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
    fm_cell_xf record{};
    record.xf_id = 3U;
    uint32_t indices[] = {0xAABBCCDDU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      indices, nullptr, 0U,      nullptr};
    expect_invalid(fm_styles_add_batch(wb.handle, &batch), "fm_styles_add_batch:");
    EXPECT_EQ(indices[0], 0xAABBCCDDU);
  }
}

TEST(FormulonCApiStyles, SetThenSaveLoadPreservesXfIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Register the xf records the stamp below names. A `<c s="7">` against
  // a shorter `<cellXfs>` resolves to no style, so the reader falls back
  // to the default rather than handing back an index whose record
  // `fm_styles_get_cell_xf` would then refuse to return.
  wb.handle->workbook().mutable_styles().cell_xfs.resize(8);

  // Stamp xf_index = 7 on cell A1 and save.
  EXPECT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 3.14), 0);
  EXPECT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, 7), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &xf), 0);
  EXPECT_EQ(xf, 7U);
}

// ---- Add-side dedup primitives -------------------------------------

namespace {

fm_font_record MakeArial() {
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

fm_fill_record MakeRedFill() {
  fm_fill_record r{};
  r.pattern = 1;  // solid
  r.fg_argb = 0xFFFF0000U;
  r.bg_argb = 0xFF000000U;
  return r;
}

fm_border_record MakeThinBoxBorder() {
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

}  // namespace

TEST(FormulonCApiStyles, AddFontDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFontDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_font_record arial = MakeArial();
  fm_font_record calibri = MakeArial();
  calibri.name = "Calibri";
  ASSERT_EQ(fm_styles_add_font(wb.handle, arial, &a), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, calibri, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFontGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  // Round-trip: the freshly added font should read back equal.
  fm_font_record out{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &out), 0);
  EXPECT_STREQ(out.name, "Arial");
  EXPECT_DOUBLE_EQ(out.size, 12.0);
  EXPECT_EQ(out.color_argb, 0xFF112233U);
  EXPECT_EQ(out.bold, 1);
}

TEST(FormulonCApiStyles, GetFontRoundTripIsIdentityForSuperscript) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record loaded{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &loaded), 0);
  EXPECT_STREQ(loaded.name, "Arial");
  EXPECT_EQ(loaded.vert_align, 1U);

  uint32_t count_before = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_before), 0);
  uint32_t reindex = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0);
  EXPECT_EQ(reindex, idx);
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count_before);
}

TEST(FormulonCApiStyles, GetEditAddFontPreservesTheUneditedFields) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record superscript = MakeArial();
  superscript.vert_align = 1;
  superscript.has_family = 1;
  superscript.family = 2;
  superscript.has_charset = 1;
  superscript.charset = 128;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, superscript, &idx), 0);

  fm_font_record edited{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, idx, &edited), 0);
  edited.color_argb = 0xFF00FF00U;
  uint32_t recolored = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, edited, &recolored), 0);
  EXPECT_NE(recolored, idx);

  fm_font_record reread{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, recolored, &reread), 0);
  EXPECT_EQ(reread.color_argb, 0xFF00FF00U);
  EXPECT_EQ(reread.vert_align, 1U);
  EXPECT_EQ(reread.has_family, 1);
  EXPECT_EQ(reread.family, 2U);
  EXPECT_EQ(reread.has_charset, 1);
  EXPECT_EQ(reread.charset, 128U);
}

namespace {

// An Excel-authored `<fonts>` section: the theme colour, `<family>` and
// `<charset>` children are what a template carries and what a record built
// from scratch through the C ABI cannot reproduce unless the ABI exposes
// them.
constexpr char kExcelAuthoredStyles[] =
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
void LoadExcelAuthoredStyles(fm_workbook_t* handle) {
  const std::vector<std::uint8_t> bytes(kExcelAuthoredStyles, kExcelAuthoredStyles + std::strlen(kExcelAuthoredStyles));
  auto parsed = formulon::io::read_styles(bytes);
  ASSERT_TRUE(parsed.has_value());
  handle->workbook().mutable_styles() = std::move(parsed.value());
}

}  // namespace

TEST(FormulonCApiStyles, AddFontIsIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t count = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count), 0);
  ASSERT_EQ(count, 3U);
  for (uint32_t i = 0; i < count; ++i) {
    fm_font_record loaded{};
    ASSERT_EQ(fm_styles_get_font(wb.handle, i, &loaded), 0) << "font " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_font(wb.handle, loaded, &reindex), 0) << "font " << i;
    EXPECT_EQ(reindex, i);
  }
  uint32_t count_after = 0;
  ASSERT_EQ(fm_styles_get_font_count(wb.handle, &count_after), 0);
  EXPECT_EQ(count_after, count);
}

TEST(FormulonCApiStyles, AddFillAndBorderAreIdentityAgainstAFileLoadedTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  uint32_t fills = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills), 0);
  ASSERT_EQ(fills, 2U);
  for (uint32_t i = 0; i < fills; ++i) {
    fm_fill_record loaded{};
    ASSERT_EQ(fm_styles_get_fill(wb.handle, i, &loaded), 0) << "fill " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_fill(wb.handle, loaded, &reindex), 0) << "fill " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t borders = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders), 0);
  ASSERT_EQ(borders, 2U);
  for (uint32_t i = 0; i < borders; ++i) {
    fm_border_record loaded{};
    ASSERT_EQ(fm_styles_get_border(wb.handle, i, &loaded), 0) << "border " << i;
    uint32_t reindex = 0xFFFFFFFFU;
    ASSERT_EQ(fm_styles_add_border(wb.handle, loaded, &reindex), 0) << "border " << i;
    EXPECT_EQ(reindex, i);
  }

  uint32_t fills_after = 0;
  uint32_t borders_after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &fills_after), 0);
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &borders_after), 0);
  EXPECT_EQ(fills_after, fills);
  EXPECT_EQ(borders_after, borders);
}

TEST(FormulonCApiStyles, AddFontDoesNotAliasAThemeColourOntoLiteralRgb) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  // Font 0 is `<color theme="1"/>`. Its sibling ARGB is only a compatibility
  // fallback, so folding it together with a caller's literal RGB record
  // would lose the selector and change the template's body text semantics.
  fm_font_record themed{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, 0U, &themed), 0);
  ASSERT_EQ(themed.color.kind, static_cast<uint8_t>(kFmColorTheme));

  fm_font_record literal_rgb = themed;
  literal_rgb.color = fm_color_spec{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, literal_rgb, &index), 0);
  EXPECT_NE(index, 0U);

  // The theme record is untouched and still serialises as a theme colour.
  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<color theme=\"1\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<color theme=\"0\" tint=\"-0.25\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, AddFillAndBorderDistinguishColourSpecifications) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_NO_FATAL_FAILURE(LoadExcelAuthoredStyles(wb.handle));

  fm_fill_record themed_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, 1U, &themed_fill), 0);
  ASSERT_EQ(themed_fill.fg.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_fill_record literal_fill = themed_fill;
  literal_fill.fg = fm_color_spec{};
  literal_fill.fg_argb = themed_fill.fg_argb;
  uint32_t fill_index = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, literal_fill, &fill_index), 0);
  EXPECT_NE(fill_index, 1U);

  fm_border_record themed_border{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, 1U, &themed_border), 0);
  ASSERT_EQ(themed_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_border_record literal_border = themed_border;
  literal_border.left.color = fm_color_spec{};
  uint32_t border_index = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, literal_border, &border_index), 0);
  EXPECT_NE(border_index, 1U);

  const std::string xml = formulon::io::write_styles(wb.handle->workbook().styles());
  EXPECT_NE(xml.find("<fgColor theme=\"4\" tint=\"0.5\"/>"), std::string::npos);
  EXPECT_NE(xml.find("<bgColor indexed=\"64\"/>"), std::string::npos);
}

TEST(FormulonCApiStyles, SelectorColoursRemainAuthoritativeAcrossGetAddAndSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_color_spec theme{};
  theme.kind = static_cast<uint8_t>(kFmColorTheme);
  theme.theme = 3U;
  theme.tint = 0.5;
  fm_color_spec indexed{};
  indexed.kind = static_cast<uint8_t>(kFmColorIndexed);
  indexed.indexed = 9U;
  fm_color_spec automatic{};
  automatic.kind = static_cast<uint8_t>(kFmColorAuto);

  fm_font_record font = MakeArial();
  font.color_argb = 0x01020304U;  // compatibility fallback, not a render result
  font.color = theme;
  uint32_t font_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, font, &font_index), 0);
  fm_font_record got_font{};
  ASSERT_EQ(fm_styles_get_font(wb.handle, font_index, &got_font), 0);
  EXPECT_EQ(got_font.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(got_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(got_font.color.tint, 0.5);
  EXPECT_EQ(got_font.color_argb, 0x01020304U);
  uint32_t font_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_font(wb.handle, got_font, &font_again), 0);
  EXPECT_EQ(font_again, font_index);

  fm_fill_record fill = MakeRedFill();
  fill.fg_argb = 0x05060708U;
  fill.bg_argb = 0x090A0B0CU;
  fill.fg = indexed;
  fill.bg = automatic;
  uint32_t fill_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, fill, &fill_index), 0);
  fm_fill_record got_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, fill_index, &got_fill), 0);
  EXPECT_EQ(got_fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_fill.fg.indexed, 9U);
  EXPECT_EQ(got_fill.bg.kind, static_cast<uint8_t>(kFmColorAuto));
  uint32_t fill_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, got_fill, &fill_again), 0);
  EXPECT_EQ(fill_again, fill_index);

  fm_border_record border = MakeThinBoxBorder();
  border.left.color = theme;
  border.right.color = indexed;
  uint32_t border_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, border, &border_index), 0);
  fm_border_record got_border{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, border_index, &got_border), 0);
  EXPECT_EQ(got_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(got_border.left.color.theme, 3U);
  EXPECT_EQ(got_border.right.color.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_border.right.color.indexed, 9U);
  uint32_t border_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, got_border, &border_again), 0);
  EXPECT_EQ(border_again, border_index);

  fm_dxf_record dxf{};
  dxf.font_engaged = 1;
  dxf.font = font;
  dxf.font.color = automatic;
  dxf.fill_engaged = 1;
  dxf.fill = fill;
  dxf.border_engaged = 1;
  dxf.border = border;
  uint32_t dxf_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, dxf, &dxf_index), 0);
  fm_dxf_record got_dxf{};
  ASSERT_EQ(fm_styles_get_dxf(wb.handle, dxf_index, &got_dxf), 0);
  EXPECT_EQ(got_dxf.font.color.kind, static_cast<uint8_t>(kFmColorAuto));
  EXPECT_EQ(got_dxf.fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(got_dxf.border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  uint32_t dxf_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_dxf(wb.handle, got_dxf, &dxf_again), 0);
  EXPECT_EQ(dxf_again, dxf_index);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);

  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(reloaded.handle, font_index, &reloaded_font), 0);
  EXPECT_EQ(reloaded_font.color.kind, static_cast<uint8_t>(kFmColorTheme));
  EXPECT_EQ(reloaded_font.color.theme, 3U);
  EXPECT_DOUBLE_EQ(reloaded_font.color.tint, 0.5);
  fm_fill_record reloaded_fill{};
  ASSERT_EQ(fm_styles_get_fill(reloaded.handle, fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.fg.kind, static_cast<uint8_t>(kFmColorIndexed));
  EXPECT_EQ(reloaded_fill.fg.indexed, 9U);
  EXPECT_EQ(reloaded_fill.bg.kind, static_cast<uint8_t>(kFmColorAuto));
  fm_border_record reloaded_border{};
  ASSERT_EQ(fm_styles_get_border(reloaded.handle, border_index, &reloaded_border), 0);
  EXPECT_EQ(reloaded_border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
  fm_dxf_record reloaded_dxf{};
  ASSERT_EQ(fm_styles_get_dxf(reloaded.handle, dxf_index, &reloaded_dxf), 0);
  EXPECT_EQ(reloaded_dxf.font.color.kind, static_cast<uint8_t>(kFmColorAuto));
  EXPECT_EQ(reloaded_dxf.border.left.color.kind, static_cast<uint8_t>(kFmColorTheme));
}

TEST(FormulonCApiStyles, AddFontDistinguishesAnExplicitOffToggleFromAnAbsentOne) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_font_record absent = MakeArial();
  absent.bold = 0;
  absent.has_bold = 0;
  fm_font_record explicit_off = absent;
  explicit_off.has_bold = 1;

  uint32_t absent_index = 0;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, absent, &absent_index), 0);
  ASSERT_EQ(fm_styles_add_font(wb.handle, explicit_off, &explicit_index), 0);
  EXPECT_NE(absent_index, explicit_index);
}

TEST(FormulonCApiStyles, AddFillDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddFillDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_fill_record red = MakeRedFill();
  fm_fill_record blue = MakeRedFill();
  blue.fg_argb = 0xFF0000FFU;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, red, &a), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, blue, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddFillGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_fill_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_fill_record out{};
  ASSERT_EQ(fm_styles_get_fill(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.pattern, 1U);
  EXPECT_EQ(out.fg_argb, 0xFFFF0000U);
}

TEST(FormulonCApiStyles, AddBorderDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddBorderDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t a = 0;
  uint32_t b = 0;
  fm_border_record thin = MakeThinBoxBorder();
  fm_border_record dashed = MakeThinBoxBorder();
  dashed.left.style = 3;  // dashed
  ASSERT_EQ(fm_styles_add_border(wb.handle, thin, &a), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, dashed, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddBorderGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &before), 0);
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_border_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 1U);
  EXPECT_EQ(idx, 0U);
  fm_border_record out{};
  ASSERT_EQ(fm_styles_get_border(wb.handle, idx, &out), 0);
  EXPECT_EQ(out.left.style, 1U);
  EXPECT_EQ(out.right.style, 1U);
  EXPECT_EQ(out.diagonal_up, 0);
  EXPECT_EQ(out.diagonal_down, 0);
}

TEST(FormulonCApiStyles, AddNumFmtBuiltinReturnsBuiltinId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0xFFFFU;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "General", &id), 0);
  EXPECT_EQ(id, 0U);
  // Adding a built-in must not create a custom entry.
  uint32_t font_count = 7;  // unrelated, just ensure other tables untouched
  EXPECT_EQ(fm_styles_get_font_count(wb.handle, &font_count), 0);
  EXPECT_EQ(font_count, 0U);
}

TEST(FormulonCApiStyles, AddNumFmtCustomReturnsCustomId) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t id = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &id), 0);
  EXPECT_GE(id, 164U);
  // Resolves through the read-side getter.
  const char* s = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb.handle, id, &s), 0);
  ASSERT_NE(s, nullptr);
  EXPECT_STREQ(s, "\"USD\" #,##0");
}

TEST(FormulonCApiStyles, AddNumFmtCustomDedup) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint16_t a = 0;
  uint16_t b = 0;
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &a), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "\"USD\" #,##0", &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddNumFmtExhaustionReturnsPreconditionWithoutMutation) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::size_t strings_before = styles.num_fmt_strings.size();
  const std::size_t records_before = styles.num_fmts.size();
  uint16_t out = 0xBEEFU;
  EXPECT_EQ(fm_styles_add_num_fmt(wb.handle, "new-format", &out),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(out, 0xBEEFU);
  EXPECT_EQ(styles.num_fmt_strings.size(), strings_before);
  EXPECT_EQ(styles.num_fmts.size(), records_before);
}

TEST(FormulonCApiStyles, AddBatchNumFmtExhaustionIsTransactional) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  auto& styles = wb.handle->workbook().mutable_styles();
  formulon::io::NumFmtRecord max_record;
  max_record.id = 65535U;
  max_record.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size());
  styles.num_fmt_strings.emplace_back("existing");
  styles.num_fmts.push_back(max_record);
  const std::string before_xml = formulon::io::write_styles(styles);
  const char* codes[] = {"first", "second"};
  uint16_t ids[] = {0xAAAAU, 0xBBBBU};
  const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr, 0U,    nullptr, nullptr, 0U,
                              nullptr, nullptr, 0U,      nullptr, codes, 2U,      ids};
  EXPECT_EQ(fm_styles_add_batch(wb.handle, &batch),
            static_cast<fm_status_t>(formulon::FormulonErrorCode::kPreconditionFailed));
  EXPECT_EQ(ids[0], 0xAAAAU);
  EXPECT_EQ(ids[1], 0xBBBBU);
  EXPECT_EQ(formulon::io::write_styles(styles), before_xml);
}

TEST(FormulonCApiStyles, AddCellXfDedupReturnsExistingIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;  // built-in General
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  uint32_t a = 0xFFFFFFFFU;
  uint32_t b = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_EQ(a, b);
}

TEST(FormulonCApiStyles, AddCellXfDistinctReturnsNewIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  xf.horizontal_align = 1;
  xf.vertical_align = 2;
  xf.wrap_text = 1;

  xf.has_wrap_text = 1;

  uint32_t a = 0;
  uint32_t b = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &a), 0);
  // Flip wrap_text — distinct record.
  xf.wrap_text = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &b), 0);
  EXPECT_NE(a, b);
}

TEST(FormulonCApiStyles, AddCellXfGrowsTable) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);
  uint32_t before = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &before), 0);
  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;
  uint32_t idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
  uint32_t after = 0;
  ASSERT_EQ(fm_styles_get_cell_xf_count(wb.handle, &after), 0);
  EXPECT_EQ(before, 0U);
  EXPECT_EQ(after, 2U);
  EXPECT_EQ(idx, 1U);
}

TEST(FormulonCApiStyles, AddCellXfOnFreshWorkbookKeepsZeroAsDefault) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  // A fresh workbook's tables start empty, so the first font/fill/border
  // added legitimately becomes index 0 — there is no anonymous placeholder
  // seeded ahead of it.
  EXPECT_EQ(font_idx, 0U);
  EXPECT_EQ(fill_idx, 0U);
  EXPECT_EQ(border_idx, 0U);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  // Distinguish this record from the all-zero placeholder xf that
  // `ensure_default_cell_xf` seeds at index 0 so the caller's record is
  // guaranteed to land at a fresh index instead of deduping to it. The
  // presence flag is what makes the value part of the record.
  xf.wrap_text = 1;
  xf.has_wrap_text = 1;

  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  EXPECT_NE(xf_idx, 0U);

  fm_cell_xf default_xf{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, 0, &default_xf), 0);
  EXPECT_EQ(default_xf.font_index, 0U);
  EXPECT_EQ(default_xf.fill_index, 0U);
  EXPECT_EQ(default_xf.border_index, 0U);
  EXPECT_EQ(default_xf.num_fmt_id, 0U);
  EXPECT_EQ(default_xf.wrap_text, 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsOutOfRangeIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  // Empty styles table: any non-zero index is out of range.
  fm_cell_xf xf{};
  xf.font_index = 5;  // intentionally OOR
  xf.fill_index = 0;
  xf.border_index = 0;
  xf.num_fmt_id = 0;
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddCellXfRejectsUnknownNumFmt) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 200;  // not a registered custom and not a documented built-in
  uint32_t idx = 0xFFFFFFFFU;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &idx), 0);
}

TEST(FormulonCApiStyles, AddNullArgumentRejected) {
  uint32_t idx = 0;
  EXPECT_NE(fm_styles_add_font(nullptr, MakeArial(), &idx), 0);
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  EXPECT_NE(fm_styles_add_font(wb.handle, MakeArial(), nullptr), 0);
  uint16_t id = 0;
  EXPECT_NE(fm_styles_add_num_fmt(nullptr, "General", &id), 0);
  EXPECT_NE(fm_styles_add_num_fmt(wb.handle, "General", nullptr), 0);
  EXPECT_NE(fm_styles_add_dxf(nullptr, fm_dxf_record{}, &idx), 0);
  EXPECT_NE(fm_styles_add_dxf(wb.handle, fm_dxf_record{}, nullptr), 0);
}

TEST(FormulonCApiStyles, FullLifecycleSurvivesSaveLoad) {
  // Build font -> xf -> stamp on cell -> save -> reload -> the font is
  // resolvable via the read-side accessors.
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0;
  uint32_t fill_idx = 0;
  uint32_t border_idx = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_border(wb.handle, MakeThinBoxBorder(), &border_idx), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = border_idx;
  xf.num_fmt_id = 0;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 0, 0, 1.5), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);
  uint32_t reread_xf = 0;
  EXPECT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 0, 0, &reread_xf), 0);
  fm_cell_xf reloaded{};
  EXPECT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);
}

// Regression coverage for the externally-reported symptom fixed by removing
// the spurious `ensure_default_style_roots` seeding from the font/fill/
// border adders: a non-default font + fill + custom numFmt combined into an
// `<xf>` must produce an `<xf>` index that is distinct from (and does not
// dedup to) the all-zero placeholder xf, and that index must survive
// `setCellXfIndex` -> save -> reload with the font, fill, and numFmt content
// intact on the reloaded workbook. Note that a fresh workbook's font/fill
// tables legitimately start at index 0 (that is the behavior the fix
// restored), so this test asserts on round-tripped *content*, not on the
// font/fill indices being non-zero. The WASM (`src/wasm/parts/
// workbook_styles.cpp`) and Node addon (`src/node_addon/parts/styles.cc`)
// bindings marshal the identical field set into `fm_cell_xf` and call this
// same `fm_styles_add_cell_xf` entry point, so this test also exercises
// their shared code path end to end.
TEST(FormulonCApiStyles, AddXfWithNonDefaultFontFillNumFmtRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  uint32_t font_idx = 0xFFFFFFFFU;
  uint32_t fill_idx = 0xFFFFFFFFU;
  uint16_t num_fmt_id = 0;
  ASSERT_EQ(fm_styles_add_font(wb.handle, MakeArial(), &font_idx), 0);
  ASSERT_EQ(fm_styles_add_fill(wb.handle, MakeRedFill(), &fill_idx), 0);
  ASSERT_EQ(fm_styles_add_num_fmt(wb.handle, "0.00%", &num_fmt_id), 0);

  fm_cell_xf xf{};
  xf.font_index = font_idx;
  xf.fill_index = fill_idx;
  xf.border_index = 0;
  xf.num_fmt_id = num_fmt_id;
  uint32_t xf_idx = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);
  ASSERT_NE(xf_idx, 0U) << "non-default font/fill/numFmt must not dedup to the "
                           "all-zero placeholder xf at index 0";

  // Stability: re-adding the identical record must return the same index.
  uint32_t xf_idx_again = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx_again), 0);
  EXPECT_EQ(xf_idx, xf_idx_again);

  ASSERT_EQ(fm_workbook_set_number(wb.handle, 0, 1, 1, 0.4225), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 1, 1, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  ASSERT_GT(saved.len, 0U);

  WorkbookGuard wb2;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &wb2.handle), 0);

  uint32_t reread_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(wb2.handle, 0, 1, 1, &reread_xf), 0);
  ASSERT_EQ(reread_xf, xf_idx);

  fm_cell_xf reloaded{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb2.handle, reread_xf, &reloaded), 0);
  EXPECT_EQ(reloaded.font_index, font_idx);
  EXPECT_EQ(reloaded.fill_index, fill_idx);
  EXPECT_EQ(reloaded.num_fmt_id, num_fmt_id);

  fm_font_record reloaded_font{};
  ASSERT_EQ(fm_styles_get_font(wb2.handle, reloaded.font_index, &reloaded_font), 0);
  EXPECT_STREQ(reloaded_font.name, "Arial");
  EXPECT_DOUBLE_EQ(reloaded_font.size, 12.0);
  EXPECT_EQ(reloaded_font.bold, 1);

  fm_fill_record reloaded_fill{};
  ASSERT_EQ(fm_styles_get_fill(wb2.handle, reloaded.fill_index, &reloaded_fill), 0);
  EXPECT_EQ(reloaded_fill.pattern, 1U);
  EXPECT_EQ(reloaded_fill.fg_argb, 0xFFFF0000U);

  const char* reloaded_fmt = nullptr;
  ASSERT_EQ(fm_styles_get_num_fmt_string(wb2.handle, reloaded.num_fmt_id, &reloaded_fmt), 0);
  ASSERT_NE(reloaded_fmt, nullptr);
  EXPECT_STREQ(reloaded_fmt, "0.00%");
}

TEST(FormulonCApiStyles, JustifyLastLineRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf xf{};
  xf.has_alignment = 1;
  xf.horizontal_align = 7;  // distributed
  xf.has_horizontal_align = 1;
  xf.vertical_align = 2;  // bottom
  xf.justify_last_line = 1;
  xf.has_justify_last_line = 1;
  uint32_t xf_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &xf_idx), 0);

  uint32_t duplicate_idx = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &duplicate_idx), 0);
  EXPECT_EQ(duplicate_idx, xf_idx);

  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "distributed"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, xf_idx), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);

  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0);
  uint32_t reread_idx = 0;
  ASSERT_EQ(fm_cell_get_xf_index(reloaded.handle, 0, 0, 0, &reread_idx), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(reloaded.handle, reread_idx, &reread), 0);
  EXPECT_EQ(reread.horizontal_align, 7U);
  EXPECT_EQ(reread.justify_last_line, 1);
}

TEST(FormulonCApiStyles, NamedCellStyleRoundTripsThroughSaveLoad) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf style_xf{};
  style_xf.has_alignment = 1;
  style_xf.horizontal_align = 2;
  style_xf.has_horizontal_align = 1;
  uint32_t xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);
  ASSERT_EQ(fm_styles_set_cell_style(wb.handle, "Highlight", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);

  fm_cell_xf cell_xf{};
  cell_xf.has_alignment = 1;
  cell_xf.horizontal_align = 2;
  cell_xf.has_horizontal_align = 1;
  cell_xf.xf_id = xf_id;
  uint32_t cell_xf_id = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, cell_xf, &cell_xf_id), 0);
  ASSERT_EQ(fm_workbook_set_text(wb.handle, 0, 0, 0, "styled"), 0);
  ASSERT_EQ(fm_cell_set_xf_index(wb.handle, 0, 0, 0, cell_xf_id), 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  WorkbookGuard loaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &loaded.handle), 0);
  uint32_t style_count = 0;
  ASSERT_EQ(fm_styles_get_cell_style_count(loaded.handle, &style_count), 0);
  ASSERT_EQ(style_count, 1U);
  fm_cell_style_record_t style{};
  ASSERT_EQ(fm_styles_get_cell_style(loaded.handle, 0, &style), 0);
  EXPECT_STREQ(style.name, "Highlight");
  uint32_t reread_cell_xf = 0;
  ASSERT_EQ(fm_cell_get_xf_index(loaded.handle, 0, 0, 0, &reread_cell_xf), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(loaded.handle, reread_cell_xf, &reread), 0);
  EXPECT_EQ(reread.xf_id, style.xf_id);
}

TEST(FormulonCApiStyles, NamedStyleXfRejectsDanglingReferences) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // A cell xf may only inherit from a named-style xf that already exists.
  fm_cell_xf cell_xf{};
  cell_xf.xf_id = 3;
  uint32_t cell_xf_index = 0;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, cell_xf, &cell_xf_index), 0);

  // The named-style table validates its own font / fill / border / numFmt
  // references exactly like the cell-xf table does.
  fm_cell_xf style_xf{};
  style_xf.font_index = 9;
  uint32_t xf_id = 0;
  EXPECT_NE(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);

  style_xf.font_index = 0;
  style_xf.has_alignment = 1;
  style_xf.justify_last_line = 1;
  style_xf.has_justify_last_line = 1;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style_xf, &xf_id), 0);
  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, xf_id, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);

  // 0..47 is the whole OOXML ordinal space; anything else needs the sentinel.
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, 48), 0);
  EXPECT_EQ(fm_styles_set_cell_style(wb.handle, "Custom", xf_id, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
  EXPECT_NE(fm_styles_set_cell_style(wb.handle, "Custom", xf_id + 1U, FM_CELL_STYLE_BUILTIN_ID_NONE), 0);
}

TEST(FormulonCApiStyles, AddBatchDeduplicatesAllStyleTables) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record fonts[] = {MakeArial(), MakeArial()};
  const fm_fill_record fills[] = {MakeRedFill(), MakeRedFill()};
  fm_border_record border{};
  border.left.style = 1;
  border.left.color_argb = 0xFF000000U;
  const fm_border_record borders[] = {border, border};
  uint32_t font_indices[2]{};
  uint32_t fill_indices[2]{};
  uint32_t border_indices[2]{};
  const char* const num_fmt_codes[] = {"0.000%", "0.000%"};
  uint16_t num_fmt_ids[2]{};
  fm_cell_xf xf{};
  xf.font_index = 1U;
  xf.fill_index = 1U;
  xf.border_index = 1U;
  const fm_cell_xf xfs[] = {xf, xf};
  uint32_t xf_indices[2]{};
  const fm_styles_batch batch{fonts, 2U, font_indices, fills,         2U, fill_indices, borders, 2U, border_indices,
                              xfs,   2U, xf_indices,   num_fmt_codes, 2U, num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_indices[0], font_indices[1]);
  EXPECT_EQ(fill_indices[0], fill_indices[1]);
  EXPECT_EQ(border_indices[0], border_indices[1]);
  EXPECT_EQ(xf_indices[0], xf_indices[1]);
  EXPECT_EQ(num_fmt_ids[0], num_fmt_ids[1]);
}

// The batch adder shares the single-record presence-flag contract, which is a
// behavioural change rather than a layout one: a caller that set an alignment
// value and relied on it being inferred as present still compiles against the
// widened record and still succeeds, but now emits no `<alignment>` child.
// Pin both halves so the silent half cannot regress unnoticed.
TEST(FormulonCApiStyles, AddBatchEmitsAlignmentOnlyForFlaggedAttributes) {
  const auto batch_xf_alignment = [](const fm_cell_xf& record) {
    WorkbookGuard wb;
    EXPECT_EQ(fm_workbook_create(&wb.handle), 0);
    uint32_t xf_indices[] = {0xDEADBEEFU};
    const fm_styles_batch batch{nullptr, 0U,      nullptr, nullptr,    0U,      nullptr, nullptr, 0U,
                                nullptr, &record, 1U,      xf_indices, nullptr, 0U,      nullptr};
    EXPECT_EQ(fm_styles_add_batch(wb.handle, &batch), 0) << fm_last_error_message();
    fm_cell_xf got{};
    EXPECT_EQ(fm_styles_get_cell_xf(wb.handle, xf_indices[0], &got), 0);
    return got;
  };

  // The pre-collapse caller shape: a value with no presence flag.
  fm_cell_xf inferred{};
  inferred.horizontal_align = 3;  // center-continuous
  const fm_cell_xf without_flag = batch_xf_alignment(inferred);
  EXPECT_EQ(without_flag.has_alignment, 0);
  EXPECT_EQ(without_flag.has_horizontal_align, 0);
  EXPECT_EQ(without_flag.horizontal_align, 0U);

  // The same value, now declared present, is the post-collapse spelling.
  fm_cell_xf flagged{};
  flagged.horizontal_align = 3;
  flagged.has_horizontal_align = 1;
  const fm_cell_xf with_flag = batch_xf_alignment(flagged);
  EXPECT_EQ(with_flag.has_alignment, 1);
  EXPECT_EQ(with_flag.has_horizontal_align, 1);
  EXPECT_EQ(with_flag.horizontal_align, 3U);
}

TEST(FormulonCApiStyles, AddBatchCommitsNumFmtReferencesTransactionally) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;  // default root plus the staged font
  xf.num_fmt_id = 164U;
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  ASSERT_EQ(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 1U);
  EXPECT_EQ(num_fmt_ids[0], 164U);
  EXPECT_EQ(xf_indices[0], 1U);
  ASSERT_EQ(wb.handle->workbook().styles().num_fmts.size(), 1U);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts[0].id, 164U);
  ASSERT_EQ(wb.handle->workbook().styles().cell_xfs.size(), 2U);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs[1].num_fmt_id, 164U);
}

TEST(FormulonCApiStyles, AddBatchFailureLeavesTableAndOutputsUnchanged) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  const std::string before_xml = formulon::io::write_styles(wb.handle->workbook().styles());
  const std::size_t before_fonts = wb.handle->workbook().styles().fonts.size();
  const std::size_t before_num_fmts = wb.handle->workbook().styles().num_fmts.size();
  const std::size_t before_cell_xfs = wb.handle->workbook().styles().cell_xfs.size();

  const fm_font_record font = MakeArial();
  uint32_t font_index = 0xA1A2A3A4U;
  const char* const codes[] = {"0.000%"};
  uint16_t num_fmt_ids[] = {0xBEEF};
  fm_cell_xf xf{};
  xf.font_index = 1U;    // valid after staging the font
  xf.num_fmt_id = 999U;  // neither built-in nor registered
  uint32_t xf_indices[] = {0xDEADBEEFU};
  const fm_styles_batch batch{&font,   1U,  &font_index, nullptr,    0U,    nullptr, nullptr,    0U,
                              nullptr, &xf, 1U,          xf_indices, codes, 1U,      num_fmt_ids};

  EXPECT_NE(fm_styles_add_batch(wb.handle, &batch), 0);
  EXPECT_EQ(font_index, 0xA1A2A3A4U);
  EXPECT_EQ(num_fmt_ids[0], 0xBEEFU);
  EXPECT_EQ(xf_indices[0], 0xDEADBEEFU);
  EXPECT_EQ(wb.handle->workbook().styles().fonts.size(), before_fonts);
  EXPECT_EQ(wb.handle->workbook().styles().num_fmts.size(), before_num_fmts);
  EXPECT_EQ(wb.handle->workbook().styles().cell_xfs.size(), before_cell_xfs);
  EXPECT_EQ(formulon::io::write_styles(wb.handle->workbook().styles()), before_xml);
}

TEST(FormulonCApiStyles, ZeroInitializedCellXfUsesDefaultAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);
  EXPECT_EQ(index, 0U);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 0);
  EXPECT_EQ(got.has_horizontal_align, 0);
  EXPECT_EQ(got.has_vertical_align, 0);
  EXPECT_EQ(got.has_wrap_text, 0);
  EXPECT_EQ(got.has_justify_last_line, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_EQ(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfIgnoresPoisonForOmittedAlignmentAttributes) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf poison{};
  poison.horizontal_align = 0xFFU;
  poison.vertical_align = 0xFFU;
  poison.wrap_text = 7;
  poison.justify_last_line = 9;

  uint32_t omitted_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, poison, &omitted_index), 0);
  EXPECT_EQ(omitted_index, 0U);

  poison.has_alignment = 1;
  uint32_t explicit_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, poison, &explicit_index), 0);

  fm_cell_xf empty{};
  empty.has_alignment = 1;
  uint32_t empty_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, empty, &empty_index), 0);
  EXPECT_EQ(empty_index, explicit_index);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 1);
  EXPECT_EQ(got.has_horizontal_align, 0);
  EXPECT_EQ(got.has_vertical_align, 0);
  EXPECT_EQ(got.has_wrap_text, 0);
  EXPECT_EQ(got.has_justify_last_line, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  const std::string cell_xfs = styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin);
  EXPECT_NE(cell_xfs.find("<alignment/>"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfExplicitTopAlignmentRemainsPresent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  record.vertical_align = 0;  // top
  record.has_vertical_align = 1;
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, record, &index), 0);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.vertical_align, 0U);
  EXPECT_EQ(got.has_vertical_align, 1);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_NE(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment vertical=\"top\"/>"),
            std::string::npos);
}

TEST(FormulonCApiStyles, ZeroInitializedCellStyleXfUsesDefaultAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf record{};
  uint32_t index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, record, &index), 0);

  fm_cell_xf got{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, index, &got), 0);
  EXPECT_EQ(got.horizontal_align, 0U);
  EXPECT_EQ(got.vertical_align, 2U);
  EXPECT_EQ(got.wrap_text, 0);
  EXPECT_EQ(got.justify_last_line, 0);
  EXPECT_EQ(got.has_alignment, 0);
  EXPECT_EQ(got.has_vertical_align, 0);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t style_xfs_begin = styles_xml.find("<cellStyleXfs");
  const std::size_t style_xfs_end = styles_xml.find("</cellStyleXfs>", style_xfs_begin);
  ASSERT_NE(style_xfs_begin, std::string::npos);
  ASSERT_NE(style_xfs_end, std::string::npos);
  EXPECT_EQ(styles_xml.substr(style_xfs_begin, style_xfs_end - style_xfs_begin).find("<alignment"), std::string::npos);
}

TEST(FormulonCApiStyles, CellXfVerticalAlignZeroIsExplicitTopOnlyWithItsPresenceFlag) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  // Without the presence flag, `vertical_align` is not read at all: the record
  // describes an omitted attribute and canonicalizes to the model default.
  fm_cell_xf omitted{};
  omitted.vertical_align = 0;  // top
  uint32_t omitted_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, omitted, &omitted_index), 0);
  fm_cell_xf got_omitted{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, omitted_index, &got_omitted), 0);
  EXPECT_EQ(got_omitted.vertical_align, 2U);  // bottom
  EXPECT_EQ(got_omitted.has_vertical_align, 0);

  fm_cell_xf explicit_top{};
  explicit_top.has_alignment = 1;
  explicit_top.vertical_align = 0;  // top
  explicit_top.has_vertical_align = 1;
  uint32_t explicit_index = 0xFFFFFFFFU;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, explicit_top, &explicit_index), 0);
  EXPECT_NE(explicit_index, omitted_index);

  fm_cell_xf got_explicit{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &got_explicit), 0);
  EXPECT_EQ(got_explicit.vertical_align, 0U);
  EXPECT_EQ(got_explicit.has_vertical_align, 1);

  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, &saved.data, &saved.len), 0);
  const std::string styles_xml = ExtractStylesXml(saved);
  const std::size_t cell_xfs_begin = styles_xml.find("<cellXfs");
  const std::size_t cell_xfs_end = styles_xml.find("</cellXfs>", cell_xfs_begin);
  ASSERT_NE(cell_xfs_begin, std::string::npos);
  ASSERT_NE(cell_xfs_end, std::string::npos);
  EXPECT_NE(styles_xml.substr(cell_xfs_begin, cell_xfs_end - cell_xfs_begin).find("<alignment vertical=\"top\"/>"),
            std::string::npos);
}

TEST(FormulonCApiStyles, CellXfPreservesOptionalAlignmentAndPresence) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf xf{};
  xf.vertical_align = 2;
  xf.justify_last_line = 1;
  xf.xf_id = 0;
  xf.has_alignment = 1;
  xf.has_text_rotation = 1;
  xf.text_rotation = 255;
  xf.has_indent = 1;
  xf.indent = 7;
  xf.has_relative_indent = 1;
  xf.relative_indent = -3;
  xf.has_shrink_to_fit = 1;
  xf.shrink_to_fit = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 2;
  xf.has_horizontal_align = 1;
  xf.has_vertical_align = 1;
  xf.has_wrap_text = 1;
  xf.has_justify_last_line = 1;

  uint32_t index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  uint32_t duplicate = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &duplicate), 0);
  EXPECT_EQ(duplicate, index);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, index, &reread), 0);
  EXPECT_EQ(reread.justify_last_line, 1);
  EXPECT_EQ(reread.xf_id, 0U);
  EXPECT_EQ(reread.has_alignment, 1);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 255U);
  EXPECT_EQ(reread.has_indent, 1);
  EXPECT_EQ(reread.indent, 7U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -3);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
  EXPECT_EQ(reread.has_reading_order, 1);
  EXPECT_EQ(reread.reading_order, 2U);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);

  // Presence is part of the dedup key: explicit zero / false differs from
  // an omitted attribute even when the effective value is the same.
  fm_cell_xf absent = xf;
  absent.has_text_rotation = 0;
  absent.has_indent = 0;
  absent.has_relative_indent = 0;
  absent.has_shrink_to_fit = 0;
  absent.has_reading_order = 0;
  uint32_t absent_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, absent, &absent_index), 0);
  EXPECT_NE(absent_index, index);
}

TEST(FormulonCApiStyles, CellXfRejectsInvalidExcelAlignmentRanges) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf xf{};
  uint32_t index = 0;

  xf.has_text_rotation = 1;
  xf.text_rotation = 181;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  xf.text_rotation = 255;
  EXPECT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);

  xf.has_text_rotation = 0;
  xf.has_indent = 1;
  xf.indent = 256;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
  xf.indent = 255;
  EXPECT_EQ(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);

  xf.has_indent = 0;
  xf.has_reading_order = 1;
  xf.reading_order = 3;
  EXPECT_NE(fm_styles_add_cell_xf(wb.handle, xf, &index), 0);
}

TEST(FormulonCApiStyles, CellXfPresenceFlagsDistinguishExplicitDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);

  fm_cell_xf explicit_defaults{};
  explicit_defaults.vertical_align = 2;
  explicit_defaults.has_alignment = 1;
  explicit_defaults.has_horizontal_align = 1;
  explicit_defaults.has_vertical_align = 1;
  explicit_defaults.has_wrap_text = 1;
  explicit_defaults.has_justify_last_line = 1;
  uint32_t explicit_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, explicit_defaults, &explicit_index), 0);

  fm_cell_xf omitted_defaults = explicit_defaults;
  omitted_defaults.has_horizontal_align = 0;
  omitted_defaults.has_vertical_align = 0;
  omitted_defaults.has_wrap_text = 0;
  omitted_defaults.has_justify_last_line = 0;
  uint32_t omitted_index = 0;
  ASSERT_EQ(fm_styles_add_cell_xf(wb.handle, omitted_defaults, &omitted_index), 0);
  EXPECT_NE(explicit_index, omitted_index);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_xf(wb.handle, explicit_index, &reread), 0);
  EXPECT_EQ(reread.has_horizontal_align, 1);
  EXPECT_EQ(reread.has_vertical_align, 1);
  EXPECT_EQ(reread.has_wrap_text, 1);
  EXPECT_EQ(reread.has_justify_last_line, 1);
}

TEST(FormulonCApiStyles, CellStyleXfRoundTripsOptionalAlignment) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_cell_xf style{};
  style.has_text_rotation = 1;
  style.text_rotation = 0;
  style.has_relative_indent = 1;
  style.relative_indent = -2;
  style.has_shrink_to_fit = 1;
  style.shrink_to_fit = false;
  uint32_t style_index = 0;
  ASSERT_EQ(fm_styles_add_cell_style_xf(wb.handle, style, &style_index), 0);

  fm_cell_xf reread{};
  ASSERT_EQ(fm_styles_get_cell_style_xf(wb.handle, style_index, &reread), 0);
  EXPECT_EQ(reread.has_text_rotation, 1);
  EXPECT_EQ(reread.text_rotation, 0U);
  EXPECT_EQ(reread.has_relative_indent, 1);
  EXPECT_EQ(reread.relative_indent, -2);
  EXPECT_EQ(reread.has_shrink_to_fit, 1);
  EXPECT_EQ(reread.shrink_to_fit, 0);
}
