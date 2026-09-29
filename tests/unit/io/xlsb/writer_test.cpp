//
// End-to-end round-trip tests for the XLSB package writer.
// `read_xlsb(write_xlsb(wb))` must reproduce every cell value `wb`
// carried in for the in-scope literal kinds, plus passthrough parts
// and Bundle 4.1 stub-formula cells.

#include "io/xlsb/writer.h"

#include <algorithm>
#include <cstdint>
#include <limits>
#include <map>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/styles_writer.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
#include "print/pagination.h"
#include "sheet.h"
#include "styles.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

ByteSpan SpanOf(const std::vector<std::uint8_t>& v) {
  return ByteSpan{v.data(), v.size()};
}

std::vector<std::uint8_t> WorksheetFormatPayloadFromSheet(const std::vector<std::uint8_t>& sheet_bytes) {
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtWsFmtInfo)) {
      const ByteSpan payload = record_or.value().payload;
      return std::vector<std::uint8_t>(payload.data, payload.data + payload.size);
    }
  }
  return {};
}

TEST(XlsbWriter, RejectsZeroSheetWorkbook) {
  Workbook wb = Workbook::create_empty();
  auto bytes_or = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(bytes_or));
  EXPECT_EQ(bytes_or.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(XlsbWriter, RoundTripsTwoSheetsWithLiteralCells) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Alpha");
  wb.add_sheet("Beta");
  // Add all sheets first, then index into the sheet vector. Holding
  // references across `add_sheet` calls is unsafe — the underlying
  // vector may reallocate, invalidating any prior reference.
  Sheet& s1 = wb.sheet(0);
  Sheet& s2 = wb.sheet(1);

  // Sheet 1: a mix of literal kinds spread across 2 rows.
  s1.set_cell_value(0U, 0U, Value::number(42.0));       // RK-encodable
  s1.set_cell_value(0U, 1U, Value::number(123.45));     // x100 form
  s1.set_cell_value(0U, 2U, Value::number(1.0 / 3.0));  // BrtCellReal
  s1.set_cell_value(0U, 3U, Value::boolean(true));
  s1.set_cell_value(0U, 4U, Value::text("hello"));
  s1.set_cell_value(1U, 0U, Value::error(ErrorCode::Div0));
  // Note: not setting an explicit blank at (1,1). The writer skips
  // implicitly-default-constructed columns produced by row growth, so
  // a blank slot does not survive the round-trip — there is no
  // BrtCellBlank record on the wire and no cell to query on read.
  // Tests for that path live further down.

  // Sheet 2: text dedup + another numeric cell.
  s2.set_cell_value(2U, 3U, Value::text("hello"));  // shares SST index with Sheet1
  s2.set_cell_value(2U, 4U, Value::text("world"));
  s2.set_cell_value(5U, 0U, Value::number(-7.0));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Workbook& rt = read_or.value().workbook;

  ASSERT_EQ(rt.sheet_count(), 2U);
  EXPECT_EQ(rt.sheet(0).name(), "Alpha");
  EXPECT_EQ(rt.sheet(1).name(), "Beta");

  // Sheet 1 cells.
  {
    const Cell* c = rt.sheet(0).cell_at(0U, 0U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_number());
    EXPECT_EQ(c->cached_value.as_number(), 42.0);
  }
  {
    const Cell* c = rt.sheet(0).cell_at(0U, 1U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_number());
    EXPECT_EQ(c->cached_value.as_number(), 123.45);
  }
  {
    const Cell* c = rt.sheet(0).cell_at(0U, 2U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_number());
    EXPECT_EQ(c->cached_value.as_number(), 1.0 / 3.0);
  }
  {
    const Cell* c = rt.sheet(0).cell_at(0U, 3U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_boolean());
    EXPECT_TRUE(c->cached_value.as_boolean());
  }
  {
    const Cell* c = rt.sheet(0).cell_at(0U, 4U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_text());
    EXPECT_EQ(c->cached_value.as_text(), "hello");
  }
  {
    const Cell* c = rt.sheet(0).cell_at(1U, 0U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_error());
    EXPECT_EQ(c->cached_value.as_error(), ErrorCode::Div0);
  }

  // Sheet 2 cells.
  {
    const Cell* c = rt.sheet(1).cell_at(2U, 3U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_text());
    EXPECT_EQ(c->cached_value.as_text(), "hello");
  }
  {
    const Cell* c = rt.sheet(1).cell_at(2U, 4U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_text());
    EXPECT_EQ(c->cached_value.as_text(), "world");
  }
  {
    const Cell* c = rt.sheet(1).cell_at(5U, 0U);
    ASSERT_NE(c, nullptr);
    ASSERT_TRUE(c->cached_value.is_number());
    EXPECT_EQ(c->cached_value.as_number(), -7.0);
  }
}

TEST(XlsbWriter, RoundTripsWorkbookWithoutTextCellsSkipsSstPart) {
  // No text cells -> writer must NOT emit xl/sharedStrings.bin, and
  // the round-trip should still succeed (the reader is fine without
  // an SST part because the rels file doesn't reference one).
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("Solo"));
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(0U, 1U, Value::boolean(false));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  ASSERT_EQ(read_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(read_or.value().cells_read, 2U);
}

TEST(XlsbWriter, RowHeadersDescribeTheirEmittedCellColumns) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Spans"));
  sheet.set_cell_value(0U, 5U, Value::number(1.0));
  sheet.set_cell_value(0U, 1024U, Value::number(2.0));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  ByteSpan cursor = SpanOf(sheet_or.value());
  bool dimensions_found = false;
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtWsDim)) {
      ByteSpan dimensions = record_or.value().payload;
      auto first_row = read_u32(dimensions);
      auto last_row = read_u32(dimensions);
      auto first_col = read_u32(dimensions);
      auto last_col = read_u32(dimensions);
      ASSERT_TRUE(first_row && last_row && first_col && last_col);
      EXPECT_EQ(first_row.value(), 0U);
      EXPECT_EQ(last_row.value(), 0U);
      EXPECT_EQ(first_col.value(), 5U);
      EXPECT_EQ(last_col.value(), 1024U);
      EXPECT_EQ(dimensions.size, 0U);
      dimensions_found = true;
      continue;
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    ASSERT_GE(payload.size, 17U);
    auto row_or = read_u32(payload);
    ASSERT_TRUE(static_cast<bool>(row_or));
    EXPECT_EQ(row_or.value(), 0U);
    ASSERT_TRUE(static_cast<bool>(read_u32(payload)));  // ixfe
    ASSERT_TRUE(static_cast<bool>(read_u16(payload)));  // miyRw
    ASSERT_TRUE(static_cast<bool>(read_u8(payload)));   // flags1
    auto flags2_or = read_u8(payload);
    ASSERT_TRUE(flags2_or);  // flags2
    EXPECT_EQ(flags2_or.value() & 0x40U, 0U);
    ASSERT_TRUE(static_cast<bool>(read_u8(payload)));  // fPhShow
    auto count_or = read_u32(payload);
    ASSERT_TRUE(static_cast<bool>(count_or));
    ASSERT_EQ(count_or.value(), 2U);
    auto first_start = read_u32(payload);
    auto first_end = read_u32(payload);
    auto second_start = read_u32(payload);
    auto second_end = read_u32(payload);
    ASSERT_TRUE(first_start && first_end && second_start && second_end);
    EXPECT_EQ(first_start.value(), 5U);
    EXPECT_EQ(first_end.value(), 5U);
    EXPECT_EQ(second_start.value(), 1024U);
    EXPECT_EQ(second_end.value(), 1024U);
    EXPECT_EQ(payload.size, 0U);
    EXPECT_TRUE(dimensions_found);
    return;
  }
  FAIL() << "missing BrtRowHdr";
}

TEST(XlsbWriter, EmitsRequiredWorksheetPrefixInSpecificationOrder) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Prefix"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  ByteSpan cursor = SpanOf(sheet_or.value());
  const std::vector<XlsbRecordType> expected = {
      XlsbRecordType::BrtBeginSheet,   XlsbRecordType::BrtWsProp,      XlsbRecordType::BrtWsDim,
      XlsbRecordType::BrtBeginWsViews, XlsbRecordType::BrtBeginWsView, XlsbRecordType::BrtEndWsView,
      XlsbRecordType::BrtEndWsViews,   XlsbRecordType::BrtWsFmtInfo,   XlsbRecordType::BrtBeginSheetData,
  };
  for (const XlsbRecordType type : expected) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    EXPECT_EQ(record_or.value().type, static_cast<std::uint16_t>(type));
  }
}

TEST(XlsbWriter, EmitsCanonicalAbsentWorksheetFormatDefaults) {
  Workbook wb = Workbook::create();
  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::vector<std::uint8_t> payload = WorksheetFormatPayloadFromSheet(sheet_or.value());
  ASSERT_EQ(payload.size(), 12U);
  ByteSpan cursor = SpanOf(payload);
  auto dx_or = read_u32(cursor);
  auto cch_or = read_u16(cursor);
  auto miy_or = read_u16(cursor);
  auto flags_or = read_u32(cursor);
  ASSERT_TRUE(dx_or && cch_or && miy_or && flags_or);
  EXPECT_EQ(dx_or.value(), 0xFFFFFFFFU);
  EXPECT_EQ(cch_or.value(), 8U);
  EXPECT_EQ(miy_or.value(), 300U);
  EXPECT_EQ(flags_or.value(), 0U);
  EXPECT_EQ(cursor.size, 0U);
}

TEST(XlsbWriter, EncodesWorksheetFormatDefaultsAndPresenceFlags) {
  Workbook wb = Workbook::create();
  SheetFormatDefaults& defaults = wb.sheet(0).mutable_format_defaults();
  defaults.base_col_width = 10.0;
  defaults.default_col_width = 12.5;
  defaults.has_default_col_width = true;
  defaults.default_row_height = 18.75;
  defaults.has_default_row_height = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::vector<std::uint8_t> payload = WorksheetFormatPayloadFromSheet(sheet_or.value());
  ASSERT_EQ(payload.size(), 12U);
  ByteSpan cursor = SpanOf(payload);
  auto dx_or = read_u32(cursor);
  auto cch_or = read_u16(cursor);
  auto miy_or = read_u16(cursor);
  auto flags_or = read_u32(cursor);
  ASSERT_TRUE(dx_or && cch_or && miy_or && flags_or);
  EXPECT_EQ(dx_or.value(), 3200U);
  EXPECT_EQ(cch_or.value(), 10U);
  EXPECT_EQ(miy_or.value(), 375U);
  EXPECT_EQ(flags_or.value(), 1U);

  defaults.default_col_width = 0.0;
  defaults.default_row_height = 0.0;
  auto zero_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(zero_or)) << zero_or.error().message << " | " << zero_or.error().context;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(zero_or.value()))));
  auto zero_sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(zero_sheet_or));
  const std::vector<std::uint8_t> zero_payload = WorksheetFormatPayloadFromSheet(zero_sheet_or.value());
  ASSERT_EQ(zero_payload.size(), 12U);
  ByteSpan zero_cursor = SpanOf(zero_payload);
  ASSERT_TRUE(read_u32(zero_cursor));
  ASSERT_TRUE(read_u16(zero_cursor));
  ASSERT_TRUE(read_u16(zero_cursor));
  auto zero_flags_or = read_u32(zero_cursor);
  ASSERT_TRUE(zero_flags_or);
  EXPECT_EQ(zero_flags_or.value(), 2U);

  defaults.default_col_width = 1.5 + 1.0 / 512.0;
  defaults.default_row_height = 0.001;
  auto quantized_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(quantized_or)) << quantized_or.error().message << " | " << quantized_or.error().context;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(quantized_or.value()))));
  auto quantized_sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(quantized_sheet_or));
  const std::vector<std::uint8_t> quantized_payload = WorksheetFormatPayloadFromSheet(quantized_sheet_or.value());
  ASSERT_EQ(quantized_payload.size(), 12U);
  ByteSpan quantized_cursor = SpanOf(quantized_payload);
  auto quantized_dx_or = read_u32(quantized_cursor);
  ASSERT_TRUE(quantized_dx_or);
  EXPECT_EQ(quantized_dx_or.value(), 384U);  // floor(1.5 * 256)
  ASSERT_TRUE(read_u16(quantized_cursor));
  auto quantized_miy_or = read_u16(quantized_cursor);
  ASSERT_TRUE(quantized_miy_or);
  EXPECT_EQ(quantized_miy_or.value(), 0U);  // nearest(0.02 twip)
  auto quantized_flags_or = read_u32(quantized_cursor);
  ASSERT_TRUE(quantized_flags_or);
  EXPECT_EQ(quantized_flags_or.value(), 2U);
}

TEST(XlsbWriter, InvalidWorksheetFormatDefaultsUseSafeFallbacksAndAreDeferred) {
  Workbook wb = Workbook::create();
  SheetFormatDefaults& defaults = wb.sheet(0).mutable_format_defaults();
  defaults.base_col_width = 8.5;
  defaults.default_col_width = -1.0;
  defaults.has_default_col_width = true;
  defaults.default_row_height = std::numeric_limits<double>::infinity();
  defaults.has_default_row_height = true;

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.deferred_feature_count, 3U);

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::vector<std::uint8_t> payload = WorksheetFormatPayloadFromSheet(sheet_or.value());
  ASSERT_EQ(payload.size(), 12U);
  ByteSpan cursor = SpanOf(payload);
  auto dx_or = read_u32(cursor);
  auto cch_or = read_u16(cursor);
  auto miy_or = read_u16(cursor);
  auto flags_or = read_u32(cursor);
  ASSERT_TRUE(dx_or && cch_or && miy_or && flags_or);
  EXPECT_EQ(dx_or.value(), 0xFFFFFFFFU);
  EXPECT_EQ(cch_or.value(), 8U);
  EXPECT_EQ(miy_or.value(), 300U);
  EXPECT_EQ(flags_or.value(), 0U);
}

TEST(XlsbWriter, DefinedNameCommentSurvivesWriteReadRoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("S1");
  DefinedName dn;
  dn.name = "Rate";
  dn.formula = "0.1";
  dn.local_sheet_id = -1;
  dn.comment = "The annual interest rate";
  wb.set_defined_names({dn});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const std::vector<DefinedName>& names = read_or.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_EQ(names[0].name, "Rate");
  EXPECT_EQ(names[0].formula, "0.1");
  EXPECT_EQ(names[0].local_sheet_id, -1);
  EXPECT_EQ(names[0].comment, "The annual interest rate");
}

TEST(XlsbWriter, DefinedNameAbsentCommentRoundTripsToEmptyString) {
  // A name with no Name Manager comment must decode back to an empty
  // string, not the string "null" or a dropped entry -- matching the
  // null `XLNullableWideString` sentinel a real Excel-authored name
  // without a comment carries (see `xlsb_fidelity_base.xlsb`'s own
  // "Rate" `BrtName` record).
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("S1");
  DefinedName dn;
  dn.name = "Rate";
  dn.formula = "0.1";
  dn.local_sheet_id = -1;
  wb.set_defined_names({dn});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const std::vector<DefinedName>& names = read_or.value().workbook.defined_names();
  ASSERT_EQ(names.size(), 1U);
  EXPECT_TRUE(names[0].comment.empty());
}

TEST(XlsbWriter, AbsentFormatDefaultsPreservePaginationAcrossRoundTrip) {
  // A sheet with no `<sheetFormatPr>` defaults must fall back to the same
  // engine defaults (15pt / 8.43ch) whether the source is the in-memory
  // model or a workbook that just came back through the XLSB writer/reader
  // pair -- neither path may bake in XLSB's own on-wire fallback (20pt row
  // height) as an *observable* default height. `paginate()`'s page count
  // is the sharpest end-to-end signal of that: two rows apart tall enough
  // to force a page break at the default row height.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("S1"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  sheet.set_cell_value(48U, 0U, Value::number(2.0));
  ASSERT_FALSE(sheet.format_defaults().has_default_row_height);
  ASSERT_FALSE(sheet.format_defaults().has_default_col_width);

  auto before_or = print::paginate(wb, 0U);
  ASSERT_TRUE(static_cast<bool>(before_or)) << before_or.error().message;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& round_tripped = read_or.value().workbook.sheet(0);
  EXPECT_FALSE(round_tripped.format_defaults().has_default_row_height);
  EXPECT_FALSE(round_tripped.format_defaults().has_default_col_width);

  auto after_or = print::paginate(read_or.value().workbook, 0U);
  ASSERT_TRUE(static_cast<bool>(after_or)) << after_or.error().message;
  EXPECT_EQ(after_or.value().page_count, before_or.value().page_count);
  EXPECT_EQ(after_or.value().h_breaks, before_or.value().h_breaks);
}

TEST(XlsbWriter, PassthroughPartsRoundTripVerbatim) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("S1"));
  s.set_cell_value(0U, 0U, Value::number(10.0));

  // Synthesize a passthrough part: pretend the original archive had
  // an `xl/theme/theme1.xml`. The bytes are arbitrary; the writer
  // copies them verbatim and the reader surfaces them again on
  // unknown_parts.
  std::vector<PassthroughPart> parts;
  PassthroughPart theme;
  theme.path = "xl/theme/theme1.xml";
  theme.content_type = "application/vnd.openxmlformats-officedocument.theme+xml";
  const std::string body = "<?xml version=\"1.0\"?><theme xmlns=\"x\"/>";
  theme.bytes.assign(body.begin(), body.end());
  parts.push_back(std::move(theme));
  PassthroughPart custom;
  custom.path = "docProps/custom.xml";
  custom.content_type = "application/vnd.openxmlformats-officedocument.custom-properties+xml";
  custom.bytes = {'<', 'c', 'u', 's', 't', 'o', 'm', '/', '>'};
  parts.push_back(std::move(custom));
  PassthroughPart thumbnail;
  thumbnail.path = "docProps/thumbnail.jpeg";
  thumbnail.content_type = "image/jpeg";
  thumbnail.bytes = {0xffU, 0xd8U, 0xffU, 0xd9U};
  parts.push_back(std::move(thumbnail));
  wb.set_passthrough_parts(std::move(parts));
  wb.set_unknown_package_rels(
      {UnknownRelationship{"rId9", "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail",
                           "docProps/thumbnail.jpeg", false}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto package_rels_or = zip.read_entry("_rels/.rels");
  ASSERT_TRUE(static_cast<bool>(package_rels_or));
  const std::string package_rels(package_rels_or.value().begin(), package_rels_or.value().end());
  EXPECT_NE(package_rels.find("custom-properties"), std::string::npos);
  EXPECT_NE(package_rels.find("Target=\"docProps/custom.xml\""), std::string::npos);
  EXPECT_NE(package_rels.find("relationships/metadata/thumbnail"), std::string::npos);
  EXPECT_NE(package_rels.find("Target=\"docProps/thumbnail.jpeg\""), std::string::npos);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  bool found_theme = false;
  for (const PassthroughPart& part : read_or.value().workbook.passthrough_parts()) {
    if (part.path == "xl/theme/theme1.xml") {
      found_theme = true;
      EXPECT_EQ(part.content_type, "application/vnd.openxmlformats-officedocument.theme+xml");
      const std::string round_tripped(part.bytes.begin(), part.bytes.end());
      EXPECT_EQ(round_tripped, body);
    }
  }
  EXPECT_TRUE(found_theme);
  ASSERT_EQ(read_or.value().workbook.unknown_package_rels().size(), 1U);
  EXPECT_EQ(read_or.value().workbook.unknown_package_rels()[0].target, "docProps/thumbnail.jpeg");
}

TEST(XlsbWriter, DropsXlsxMetadataAndItsWorkbookRelationship) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("S1"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  PassthroughPart metadata;
  metadata.path = "xl/metadata.xml";
  metadata.content_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml";
  metadata.bytes = {'<', 'm', '/', '>'};
  wb.set_passthrough_parts({metadata});
  wb.set_unknown_workbook_rels(
      {UnknownRelationship{"rId7", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata",
                           "xl/metadata.xml", false}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  EXPECT_FALSE(zip.has_entry("xl/metadata.xml"));
  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_EQ(rels.find("sheetMetadata"), std::string::npos);
}

TEST(XlsbWriter, EmitsDynamicArrayMetadataForSpillAnchors) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Spill"));
  sheet.set_cell_formula(0U, 0U, "=SEQUENCE(2)");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/metadata.bin"));

  auto metadata_or = zip.read_entry("xl/metadata.bin");
  ASSERT_TRUE(static_cast<bool>(metadata_or));
  ByteSpan metadata_cursor = SpanOf(metadata_or.value());
  auto metadata_record_or = read_record(metadata_cursor);
  ASSERT_TRUE(static_cast<bool>(metadata_record_or));
  EXPECT_EQ(metadata_record_or.value().type, 332U);  // BrtBeginMetadata

  auto content_types_or = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(content_types_or));
  const std::string content_types(content_types_or.value().begin(), content_types_or.value().end());
  EXPECT_NE(content_types.find("/xl/metadata.bin"), std::string::npos);
  EXPECT_NE(content_types.find("application/vnd.ms-excel.sheetMetadata"), std::string::npos);

  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("relationships/sheetMetadata"), std::string::npos);
  EXPECT_NE(rels.find("Target=\"metadata.bin\""), std::string::npos);

  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan sheet_cursor = SpanOf(sheet_or.value());
  bool found_cell_metadata = false;
  while (sheet_cursor.size > 0U) {
    auto record_or = read_record(sheet_cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto index_or = read_u32(payload);
    ASSERT_TRUE(static_cast<bool>(index_or));
    EXPECT_EQ(index_or.value(), 1U);
    EXPECT_EQ(payload.size, 0U);
    found_cell_metadata = true;
  }
  EXPECT_TRUE(found_cell_metadata);
}

// ---------------------------------------------------------------------------
// A `BrtCellMeta` index names an entry of the metadata part that actually
// ships. The generated part declares one entry, so index 1 is right by
// construction; a retained passthrough part carries its own numbering, and
// its first entry may be rich-value or cube-function metadata rather than the
// dynamic-array one. An index that misses makes Excel repair the file, so an
// unidentifiable entry means no `BrtCellMeta` record at all.
// ---------------------------------------------------------------------------

namespace {

// Builds a metadata part declaring `type_names` in order, followed by one
// cell-metadata entry per name (entry N names type N). Mirrors the record
// framing Excel writes, so the writer's finder sees a realistic part.
std::vector<std::uint8_t> BuildMetadataPart(const std::vector<std::string>& type_names) {
  constexpr std::uint16_t kBrtBeginMetadata = 332;
  constexpr std::uint16_t kBrtEndMetadata = 333;
  constexpr std::uint16_t kBrtBeginEsmdtinfo = 334;
  constexpr std::uint16_t kBrtMdtinfo = 335;
  constexpr std::uint16_t kBrtEndEsmdtinfo = 336;
  constexpr std::uint16_t kBrtBeginEsfmd = 337;
  constexpr std::uint16_t kBrtEndEsfmd = 338;
  constexpr std::uint16_t kBrtMdb = 51;

  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginMetadata, ByteSpan{});

  emit_u32(payload, static_cast<std::uint32_t>(type_names.size()));
  emit_record(out, kBrtBeginEsmdtinfo, payload);
  for (const std::string& name : type_names) {
    payload.clear();
    emit_u32(payload, 0xD86AC0B0U);
    emit_u32(payload, 0x0001D4C0U);
    emit_xlwidestring(payload, name);
    emit_record(out, kBrtMdtinfo, payload);
  }
  emit_record(out, kBrtEndEsmdtinfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, static_cast<std::uint32_t>(type_names.size()));
  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsfmd, payload);
  for (std::size_t i = 0; i < type_names.size(); ++i) {
    payload.clear();
    emit_u32(payload, 1U);                                 // One (type, id) pair.
    emit_u32(payload, static_cast<std::uint32_t>(i + 1));  // 1-based type ordinal.
    emit_u32(payload, 0U);
    emit_record(out, kBrtMdb, payload);
  }
  emit_record(out, kBrtEndEsfmd, ByteSpan{});
  emit_record(out, kBrtEndMetadata, ByteSpan{});
  return out;
}

Workbook SpillWorkbookWithRetainedMetadata(std::vector<std::uint8_t> metadata_bytes) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Spill"));
  sheet.set_cell_formula(0U, 0U, "=SEQUENCE(2)");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  EXPECT_TRUE(sheet.commit_spill(0U, 0U, 2U, 1U, {Value::number(1.0), Value::number(2.0)}));
  PassthroughPart metadata;
  metadata.path = "xl/metadata.bin";
  metadata.content_type = "application/vnd.ms-excel.sheetMetadata";
  metadata.bytes = std::move(metadata_bytes);
  wb.set_passthrough_parts({metadata});
  return wb;
}

// Collects every `BrtCellMeta` index in a worksheet body, and reports whether
// the body carries an array-formula record at all.
struct SheetMetaScan {
  std::vector<std::uint32_t> cell_meta_indices;
  bool has_array_formula = false;
};

SheetMetaScan ScanSheetForCellMeta(const std::vector<std::uint8_t>& sheet_bytes) {
  SheetMetaScan scan;
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    EXPECT_TRUE(static_cast<bool>(record_or));
    if (!record_or) {
      break;
    }
    if (record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      scan.has_array_formula = true;
      continue;
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto index_or = read_u32(payload);
    EXPECT_TRUE(static_cast<bool>(index_or));
    if (index_or) {
      scan.cell_meta_indices.push_back(index_or.value());
    }
  }
  return scan;
}

// Writes `wb`, then returns the scan of its first worksheet body.
SheetMetaScan WriteAndScanFirstSheet(const Workbook& wb) {
  auto bytes_or = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes_or));
  if (!bytes_or) {
    return SheetMetaScan{};
  }
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(sheet_or));
  if (!sheet_or) {
    return SheetMetaScan{};
  }
  return ScanSheetForCellMeta(sheet_or.value());
}

}  // namespace

TEST(XlsbMetadataBin, FindsDynamicArrayEntryBehindOtherTypes) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR", "XLMDX"});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 2U);
}

TEST(XlsbMetadataBin, ReportsZeroWhenNoDynamicArrayTypeIsDeclared) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLMDX"});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 0U);
}

TEST(XlsbMetadataBin, GeneratedPartResolvesToIndexOne) {
  const std::vector<std::uint8_t> part = build_dynamic_array_metadata_bin();
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 1U);
}

TEST(XlsbMetadataBin, ReportsZeroForEveryTruncationOfAValidPart) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR"});
  ASSERT_EQ(find_dynamic_array_cell_meta_index(SpanOf(part)), 2U);
  // Every strict prefix is either unreadable or missing the entry it would
  // have to name. None may produce an index.
  for (std::size_t length = 0; length < part.size(); ++length) {
    const ByteSpan prefix{part.data(), length};
    EXPECT_EQ(find_dynamic_array_cell_meta_index(prefix), 0U) << "prefix length " << length;
  }
}

TEST(XlsbMetadataBin, ReportsZeroForMalformedBytes) {
  EXPECT_EQ(find_dynamic_array_cell_meta_index(ByteSpan{}), 0U);
  // A record header promising more payload than the buffer holds.
  const std::vector<std::uint8_t> overrun = {0x4C, 0x02, 0x7F, 0x01, 0x02};
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(overrun)), 0U);
  // A type table whose entry payload ends before its name.
  std::vector<std::uint8_t> short_name;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, 1U);
  emit_record(short_name, 334U, payload);
  payload.clear();
  emit_u32(payload, 0U);
  emit_record(short_name, 335U, payload);
  emit_record(short_name, 336U, ByteSpan{});
  EXPECT_EQ(find_dynamic_array_cell_meta_index(SpanOf(short_name)), 0U);
}

TEST(XlsbWriter, SpillAnchorNamesTheRetainedPartsDynamicArrayEntry) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata(BuildMetadataPart({"XLRICHVALUE", "XLDAPR", "XLMDX"}));
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  ASSERT_EQ(scan.cell_meta_indices.size(), 1U);
  EXPECT_EQ(scan.cell_meta_indices[0], 2U) << "the anchor named the retained part's first entry instead of its "
                                              "dynamic-array entry";
  EXPECT_TRUE(scan.has_array_formula);
}

TEST(XlsbWriter, SpillAnchorEmitsNoCellMetaWhenRetainedPartHasNoDynamicArrayEntry) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata(BuildMetadataPart({"XLRICHVALUE", "XLMDX"}));
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  EXPECT_TRUE(scan.cell_meta_indices.empty()) << "a BrtCellMeta index that names no dynamic-array entry";
  EXPECT_TRUE(scan.has_array_formula) << "the anchor must still be written as an array formula";
}

TEST(XlsbWriter, SpillAnchorEmitsNoCellMetaWhenRetainedPartIsMalformed) {
  const Workbook wb = SpillWorkbookWithRetainedMetadata({0x4C, 0x02, 0x7F, 0x01, 0x02});
  const SheetMetaScan scan = WriteAndScanFirstSheet(wb);
  EXPECT_TRUE(scan.cell_meta_indices.empty()) << "a dangling BrtCellMeta index from unreadable metadata bytes";
  EXPECT_TRUE(scan.has_array_formula);
}

TEST(XlsbWriter, RetainedMetadataPartShipsVerbatimAndIsNotRegenerated) {
  const std::vector<std::uint8_t> part = BuildMetadataPart({"XLRICHVALUE", "XLDAPR"});
  const Workbook wb = SpillWorkbookWithRetainedMetadata(part);
  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto shipped_or = zip.read_entry("xl/metadata.bin");
  ASSERT_TRUE(static_cast<bool>(shipped_or));
  EXPECT_EQ(shipped_or.value(), part) << "the generated part displaced the retained one";
  // The index the sheet carries must resolve inside the part that shipped.
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const SheetMetaScan scan = ScanSheetForCellMeta(sheet_or.value());
  ASSERT_EQ(scan.cell_meta_indices.size(), 1U);
  EXPECT_EQ(scan.cell_meta_indices[0], find_dynamic_array_cell_meta_index(SpanOf(shipped_or.value())));
}

TEST(XlsbWriter, EmitsDynamicArrayMetadataForSingleCellArrayAnchors) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("SingleArray"));
  sheet.set_cell_formula(0U, 0U, "=IFS(TRUE,\"yes\")");
  sheet.set_cell_dynamic_array(0U, 0U, true);
  ASSERT_TRUE(sheet.commit_spill(0U, 0U, 1U, 1U, {Value::text("yes")}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/metadata.bin"));

  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan cursor = SpanOf(sheet_or.value());
  bool found_cell_metadata = false;
  bool found_array_formula = false;
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    found_cell_metadata =
        found_cell_metadata || record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtCellMeta);
    found_array_formula =
        found_array_formula || record_or.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla);
  }
  EXPECT_TRUE(found_cell_metadata);
  EXPECT_TRUE(found_array_formula);
}

// ---------------------------------------------------------------------------
// Ptg class byte at the formula's own root position. Measured against real
// Excel 365 `.xlsb` output (`backup/oracle_probe/xlsb_spill/*.xlsb`,
// `backup/oracle_probe/self_book/self_book_name.xlsb`): a bare reference
// that is the *entire* formula body -- nothing else consumes it -- is
// written in value class, not reference class. A reference-class token
// cannot stand alone as a formula's result; without this a cell whose
// whole formula is a bare range reference reads back `#VALUE!` in Excel.
// A reference used as a function argument, or as an operand the `:` fast
// path collapses into an Area/Area3d from within an otherwise-nested
// position, is unaffected -- see the "stays reference class" tests below.
// ---------------------------------------------------------------------------

namespace {

// Returns the `rgce` bytes of the first `BrtArrFmla` record whose `colFirst`
// equals `want_col`, or an empty vector if none matches.
std::vector<std::uint8_t> FindArrFmlaRgce(const std::vector<std::uint8_t>& sheet_bytes, std::uint32_t want_col) {
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    auto rw_first_or = read_u32(p);
    auto rw_last_or = read_u32(p);
    auto col_first_or = read_u32(p);
    auto col_last_or = read_u32(p);
    (void)rw_first_or;
    (void)rw_last_or;
    (void)col_last_or;
    if (!col_first_or || col_first_or.value() != want_col) {
      continue;
    }
    if (p.size < 1) {
      return {};
    }
    p.data += 1;  // flag byte
    p.size -= 1;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}

// Returns the `rgce` bytes of the `BrtFmlaNum` record at `want_col` (a
// plain, non-array formula cell), or an empty vector if none matches.
std::vector<std::uint8_t> FindFmlaNumRgce(const std::vector<std::uint8_t>& sheet_bytes, std::uint32_t want_col) {
  ByteSpan cursor = SpanOf(sheet_bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtFmlaNum)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    auto col_or = read_u32(p);
    if (!col_or || col_or.value() != want_col) {
      continue;
    }
    // Skip iStyleRef(3) + fPhShow(1) + the double value(8) + grbitFlags(2).
    if (p.size < 14) {
      return {};
    }
    p.data += 14;
    p.size -= 14;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}

}  // namespace

TEST(XlsbWriter, RootLevelSameSheetRangeUsesValueClassArea) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=A10:A11");  // C2
  ASSERT_TRUE(sheet.commit_spill(1U, 2U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x45U);  // PtgArea, value class (0x25 | value-class bit)
}

TEST(XlsbWriter, RootLevelCrossSheetRangeUsesValueClassArea3d) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.set_cell_formula(1U, 1U, "=Sheet1!A10:A11");  // B2
  ASSERT_TRUE(sheet2.commit_spill(1U, 1U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 1U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x5BU);  // PtgArea3d, value class (0x3B | value-class bit)
}

TEST(XlsbWriter, RootLevelSingleCellReferenceUsesValueClassRef) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=A10");  // C2, plain (non-array) formula cell

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x44U);  // PtgRef, value class (0x24 | value-class bit)
}

TEST(XlsbWriter, RootLevelCrossSheetSingleCellReferenceUsesValueClassRef3d) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.set_cell_formula(1U, 2U, "=Sheet1!A10");  // C2

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet2.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x5AU);  // PtgRef3d, value class (0x3A | value-class bit)
}

TEST(XlsbWriter, NestedRangeArgumentStaysReferenceClass) {
  // A reference consumed as a function argument is unaffected by the
  // root-position promotion above: it must stay reference class, exactly
  // as before. Real Excel 365 reads a value/array-class 3-D reference used
  // this way as `#REF!`.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  sheet.set_cell_formula(1U, 2U, "=SUM(A10:A11)");  // C2, plain (non-array) formula cell

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindFmlaNumRgce(sheet_or.value(), 2U);
  ASSERT_FALSE(rgce.empty());
  EXPECT_EQ(rgce[0], 0x25U);  // PtgArea, reference class -- SUM's argument
}

TEST(XlsbWriter, RootLevelChainedRangeWithDefinedNameWrapsInValueClassMemFunc) {
  // The general (non-collapsible) form of a `:` range at root position:
  // one endpoint is a defined name, so the fast-path Area/Area3d collapse
  // in `emit_range` does not apply. Measured against
  // `backup/oracle_probe/self_book/self_book_name.xlsb`'s Sheet2!D1
  // (`=Sheet1!A10:Book!Rng`): real Excel wraps `operand + operand +
  // PtgRange` in a value-class `PtgMemFunc` ("function returns a range",
  // per ptg.h) carrying the wrapped run's byte length ahead of it, because
  // `PtgRange` itself carries no class bits to promote directly.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Sheet1"));
  DefinedName dn;
  dn.name = "Rng";
  dn.formula = "$A$10:$A$11";
  dn.local_sheet_id = -1;
  wb.set_defined_names({dn});
  sheet.set_cell_formula(1U, 2U, "=A10:Rng");  // C2
  ASSERT_TRUE(sheet.commit_spill(1U, 2U, 2U, 1U, {Value::number(5.0), Value::number(7.0)}));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  const std::vector<std::uint8_t> rgce = FindArrFmlaRgce(sheet_or.value(), 2U);
  ASSERT_GE(rgce.size(), 3U);
  EXPECT_EQ(rgce[0], 0x49U);  // PtgMemFunc, value class (0x29 | value-class bit)
  const std::uint16_t wrapped_len = static_cast<std::uint16_t>(rgce[1] | (rgce[2] << 8U));
  ASSERT_EQ(static_cast<std::size_t>(wrapped_len), rgce.size() - 3U);
  EXPECT_EQ(rgce.back(), 0x11U);  // the wrapped run ends with PtgRange (`:`)
}

namespace {

// Returns the `rgce` bytes of the `BrtName` record spelling `want_name`, or
// an empty vector if none matches.
std::vector<std::uint8_t> FindNameRgce(const std::vector<std::uint8_t>& workbook_bin, std::string_view want_name) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtName)) {
      continue;
    }
    ByteSpan p = record_or.value().payload;
    if (p.size < 9U) {
      return {};
    }
    p.data += 9;  // grbit(4) + chKey(1) + itab(4)
    p.size -= 9;
    auto cch_or = read_u32(p);
    if (!cch_or || cch_or.value() * 2U > p.size) {
      return {};
    }
    std::string decoded;
    decoded.reserve(cch_or.value());
    for (std::uint32_t i = 0; i < cch_or.value(); ++i) {
      decoded.push_back(static_cast<char>(p.data[i * 2U]));
    }
    p.data += cch_or.value() * 2U;
    p.size -= cch_or.value() * 2U;
    auto cce_or = read_u32(p);
    if (!cce_or || cce_or.value() > p.size) {
      return {};
    }
    if (decoded != want_name) {
      continue;
    }
    return std::vector<std::uint8_t>(p.data, p.data + cce_or.value());
  }
  return {};
}

// Returns the first nine payload bytes (grbit, chKey, itab) of the `BrtName`
// record spelling `want_name`, or an empty vector if none matches.
std::vector<std::uint8_t> FindNameRecordPrefix(const std::vector<std::uint8_t>& workbook_bin,
                                               std::string_view want_name) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return {};
    }
    const ByteSpan p = record_or.value().payload;
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtName) || p.size < 13U) {
      continue;
    }
    const std::uint32_t cch = static_cast<std::uint32_t>(p.data[9]) | (static_cast<std::uint32_t>(p.data[10]) << 8);
    if (cch != want_name.size() || p.size < 13U + cch * 2U) {
      continue;
    }
    std::string decoded;
    for (std::uint32_t i = 0; i < cch; ++i) {
      decoded.push_back(static_cast<char>(p.data[13U + i * 2U]));
    }
    if (decoded == want_name) {
      return std::vector<std::uint8_t>(p.data, p.data + 9);
    }
  }
  return {};
}

// Returns `xl/workbook.bin`'s `BrtWbView` `itabCur` field (u32 at payload
// offset 24), or `0xFFFFFFFF` if the record is missing or truncated.
std::uint32_t FindItabCur(const std::vector<std::uint8_t>& workbook_bin) {
  ByteSpan cursor = SpanOf(workbook_bin);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    if (!record_or) {
      return 0xFFFFFFFFU;
    }
    if (record_or.value().type != 158U) {  // BrtWbView
      continue;
    }
    ByteSpan p = record_or.value().payload;
    if (p.size < 28U) {
      return 0xFFFFFFFFU;
    }
    p.data += 24;
    auto itab_or = read_u32(p);
    return itab_or ? itab_or.value() : 0xFFFFFFFFU;
  }
  return 0xFFFFFFFFU;
}

}  // namespace

TEST(XlsbWriter, DefinedNameBodyRootReferenceStaysReferenceClass) {
  // Measured against backup/oracle_probe/self_book/self_book_name.xlsb's
  // BrtName "Rng" (=Sheet1!$A$10:$A$11): unlike a cell formula, a defined
  // name's own body keeps reference class at its root.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  DefinedName dn;
  dn.name = "Rng";
  dn.formula = "Sheet1!$A$10:$A$11";
  dn.local_sheet_id = -1;
  wb.set_defined_names({dn});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  const std::vector<std::uint8_t> rgce = FindNameRgce(wb_bin_or.value(), "Rng");
  const std::vector<std::uint8_t> want = {0x3B, 0x00, 0x00, 0x09, 0x00, 0x00, 0x00, 0x0A,
                                          0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00};
  EXPECT_EQ(rgce, want);
}

TEST(XlsbWriter, ItabCurMatchesTheTabSelectedSheet) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  Sheet& sheet2 = wb.sheet(wb.add_sheet("Sheet2"));
  sheet2.mutable_view().tab_selected = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  EXPECT_EQ(FindItabCur(wb_bin_or.value()), 1U);
}

TEST(XlsbWriter, ItabCurIsZeroWhenNoSheetIsTabSelected) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));

  EXPECT_EQ(FindItabCur(wb_bin_or.value()), 0U);
}

TEST(XlsbWriter, RealFormulaRoundTripsAsFormulaCell) {
  // An engine-authored formula encodes to a Ptg stream, survives the
  // write, and decodes back to the same formula text on read.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(2U, 3U, "=A1+B2*3");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const Cell* c = read_or.value().workbook.sheet(0).cell_at(2U, 3U);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=A1+B2*3");
}

// Measured on backup/oracle_probe/qualified_xlsb/qualified.xlsb: Excel saves
// `S!Name` that neither `S` nor the workbook defines against an empty stub
// scoped to `S`, never against another sheet's definition, so `Sheet2!SVal`
// stays #NAME? while only Sheet1 defines SVal.
TEST(XlsbWriter, SheetQualifiedNamesKeepTheirQualifierAndScope) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  wb.add_sheet("Sheet2");
  DefinedName dn;
  dn.name = "SVal";
  dn.formula = "9";
  dn.local_sheet_id = 0;
  wb.set_defined_names({dn});
  struct Case {
    const char* formula;
    Value want;
  };
  const Case cases[] = {
      {"=Sheet2!SVal", Value::error(ErrorCode::Name)},      {"=Sheet1!SVal", Value::number(9.0)},
      {"=Sheet1!NOSUCH(1)", Value::error(ErrorCode::Name)}, {"=Sheet2!NOSUCH(1)", Value::error(ErrorCode::Name)},
      {"=NOSUCH(1)", Value::error(ErrorCode::Name)},        {"=Sheet1!A1(1)", Value::error(ErrorCode::Ref)},
  };
  for (std::uint32_t row = 0; row < std::size(cases); ++row) {
    ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, row, 1U, cases[row].formula)));
  }

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  Workbook& back = read_or.value().workbook;
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));

  for (std::uint32_t row = 0; row < std::size(cases); ++row) {
    const Cell* c = back.sheet(0).cell_at(row, 1U);
    ASSERT_NE(c, nullptr) << cases[row].formula;
    EXPECT_EQ(c->formula_text, cases[row].formula);
    EXPECT_EQ(c->cached_value, cases[row].want) << cases[row].formula;
  }
}

TEST(XlsbWriter, DefinedNameWithFutureFunctionRoundTrips) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Data");
  wb.set_defined_names({DefinedName{"Joined", "TEXTJOIN(\",\",TRUE,A1:A2)", -1, false, ""}});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  ASSERT_EQ(read_or.value().workbook.defined_names().size(), 1U);
  EXPECT_EQ(read_or.value().workbook.defined_names()[0].name, "Joined");
  EXPECT_EQ(read_or.value().workbook.defined_names()[0].formula, "_xlfn.TEXTJOIN(\",\",TRUE,A1:A2)");
}

// True when `haystack` contains `needle` encoded the way `BrtName`
// stores a name: UTF-16LE, no BOM. Excel names are ASCII, so each byte
// is followed by a zero byte.
bool ContainsUtf16Le(const std::vector<std::uint8_t>& haystack, std::string_view needle) {
  std::vector<std::uint8_t> wide;
  wide.reserve(needle.size() * 2U);
  for (const char c : needle) {
    wide.push_back(static_cast<std::uint8_t>(c));
    wide.push_back(0U);
  }
  return std::search(haystack.begin(), haystack.end(), wide.begin(), wide.end()) != haystack.end();
}

TEST(XlsbWriter, FutureFunctionCallRegistersHiddenName) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=XLOOKUP(\"k\",A1:A3,B1:B3)");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_TRUE(ContainsUtf16Le(workbook_or.value(), "_xlfn.XLOOKUP"));
}

// Measured on backup/oracle_probe/dyn_err: Excel stores a cached #SPILL! or
// #CALC! as #VALUE! (the real error lives in a rich value this writer does
// not produce) and #GETTING_DATA as #N/A. A newer error byte in a cell made
// Excel refuse the whole file.
TEST(XlsbWriter, NewerErrorsAreStoredAsTheirLegacyFallback) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("E"));
  const ErrorCode errors[] = {ErrorCode::Spill, ErrorCode::Calc, ErrorCode::GettingData, ErrorCode::Div0};
  const ErrorCode stored[] = {ErrorCode::Value, ErrorCode::Value, ErrorCode::NA, ErrorCode::Div0};
  for (std::uint32_t row = 0; row < std::size(errors); ++row) {
    s.set_cell_formula(row, 0U, "=1/0");
    s.set_cell_cached_value(row, 0U, Value::error(errors[row]));
    s.set_cell_value(row, 1U, Value::error(errors[row]));
  }

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& back = read_or.value().workbook.sheet(0);
  for (std::uint32_t row = 0; row < std::size(errors); ++row) {
    for (std::uint32_t col = 0; col < 2U; ++col) {
      const Cell* c = back.cell_at(row, col);
      ASSERT_NE(c, nullptr);
      const Value& v = c->cached_value;
      EXPECT_EQ(v, Value::error(stored[row])) << "row " << row << " col " << col;
    }
  }
}

// Measured on backup/oracle_probe/root_ops: Excel keeps `_xlfn.` on a name it
// does not know, in the formula bar and in the file.
TEST(XlsbWriter, UnknownPrefixedNameKeepsItsPrefix) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 0U, "=_xlfn.FOOBAR(1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=_xlfn.XLOOKUP(1,B1:B2,C1:C2)")));
  EXPECT_EQ(wb.sheet(0).cell_at(0U, 0U)->formula_text, "=_xlfn.FOOBAR(1)");
  EXPECT_EQ(wb.sheet(0).cell_at(1U, 0U)->formula_text, "=XLOOKUP(1,B1:B2,C1:C2)");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  // `_xlfn.FOOBAR` is a plain undefined-name stub (flags 0, empty body), and
  // only a function Excel knows gets the hidden future-function record.
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto wb_bin_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(wb_bin_or));
  EXPECT_EQ(FindNameRecordPrefix(wb_bin_or.value(), "_xlfn.FOOBAR"),
            (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));
  EXPECT_TRUE(FindNameRgce(wb_bin_or.value(), "_xlfn.FOOBAR").empty());
  EXPECT_EQ(FindNameRecordPrefix(wb_bin_or.value(), "_xlfn.XLOOKUP"),
            (std::vector<std::uint8_t>{0x0B, 0x00, 0x02, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& back = read_or.value().workbook.sheet(0);
  EXPECT_EQ(back.cell_at(0U, 0U)->formula_text, "=_xlfn.FOOBAR(1)");
  EXPECT_EQ(back.cell_at(1U, 0U)->formula_text, "=XLOOKUP(1,B1:B2,C1:C2)");
}

// A loaded formula keeps the meaning its file gives it. With no dynamic-array
// mark (`cm` / `BrtCellMeta`), a fixed-size CSE block and an
// implicit-intersection =SUM(A1:A2*2) are not dynamic-array formulas, and a
// save through either container keeps them so: the block its array range,
// neither of them a mark. A formula entered through `Workbook` is marked and
// keeps the mark across both containers.
TEST(XlsbWriter, LegacyFormulasKeepTheirMeaningAcrossContainers) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("Sheet1"));
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(1U, 0U, Value::number(2.0));
  s.set_cell_formula(0U, 2U, "=A1:A2*2");
  ASSERT_TRUE(s.commit_spill(0U, 2U, 2U, 1U, {Value::number(2.0), Value::number(4.0)}));
  s.set_cell_formula(0U, 4U, "=SUM(A1:A2*2)");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 5U, "=SUM(A1:A2*2)")));

  auto check = [](const Workbook& back, const char* label) {
    const Sheet& sheet = back.sheet(0);
    EXPECT_FALSE(sheet.cell_at(0U, 2U)->dynamic_array) << label;
    const SpillRegion* block = sheet.spill_region_at_anchor(0U, 2U);
    ASSERT_NE(block, nullptr) << label;
    EXPECT_EQ(block->rows, 2U) << label;
    EXPECT_FALSE(sheet.cell_at(0U, 4U)->dynamic_array) << label;
    EXPECT_TRUE(sheet.cell_at(0U, 5U)->dynamic_array) << label;
  };

  auto xlsx_or = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx_or)) << xlsx_or.error().message;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(xlsx_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.xml");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  const std::string sheet_xml(sheet_or.value().begin(), sheet_or.value().end());
  EXPECT_NE(sheet_xml.find("<c r=\"C1\"><f t=\"array\" ref=\"C1:C2\">"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<c r=\"E1\"><f>SUM(A1:A2*2)</f>"), std::string::npos) << sheet_xml;
  EXPECT_NE(sheet_xml.find("<c r=\"F1\" cm=\"1\"><f t=\"array\" ref=\"F1\">"), std::string::npos) << sheet_xml;
  auto from_xlsx = read_ooxml(SpanOf(xlsx_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << from_xlsx.error().message;
  check(from_xlsx.value().workbook, "xlsx");

  auto xlsb_or = write_xlsb(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsb_or)) << xlsb_or.error().message;
  auto from_xlsb = read_xlsb(SpanOf(xlsb_or.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << from_xlsb.error().message;
  check(from_xlsb.value().workbook, "xlsb");
}

TEST(XlsbWriter, CubeFunctionsEncodeWithTheirFunctionIds) {
  // Excel 365 saves `CUBEVALUE("c","m")` as two strings and
  // `PtgFuncVar(2, 380)`, not through a hidden `_xlfn.` name.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=CUBEVALUE(\"c\",\"m\")");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_FALSE(ContainsUtf16Le(workbook_or.value(), "CUBEVALUE"));

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=CUBEVALUE(\"c\",\"m\")");
}

TEST(XlsbWriter, UndefinedNameCallIsStoredAgainstAWorkbookStub) {
  // Measured on backup/oracle_probe/root_ops: a lone `NOSUCH(1)` is
  // `PtgName` + the argument + `PtgFuncVar(255)`, the name an empty,
  // visible, workbook-scoped BrtName.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 0U, 0U, "=NOSUCH(1)")));
  ASSERT_TRUE(static_cast<bool>(wb.set_cell_formula(0, 1U, 0U, "=NOSUCHREF")));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value().bytes))));
  auto workbook_or = zip.read_entry("xl/workbook.bin");
  ASSERT_TRUE(static_cast<bool>(workbook_or));
  EXPECT_EQ(FindNameRecordPrefix(workbook_or.value(), "NOSUCH"),
            (std::vector<std::uint8_t>{0x00, 0x00, 0x00, 0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF}));

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  Workbook& back = read_or.value().workbook;
  ASSERT_TRUE(static_cast<bool>(back.recalc(eval::default_registry())));
  for (std::uint32_t row = 0; row < 2U; ++row) {
    const Cell* cell = back.sheet(0).cell_at(row, 0U);
    ASSERT_NE(cell, nullptr);
    EXPECT_EQ(cell->formula_text, row == 0U ? "=NOSUCH(1)" : "=NOSUCHREF");
    EXPECT_EQ(cell->cached_value, Value::error(ErrorCode::Name));
  }
}

TEST(XlsbWriter, LocalisedJisSpellingSavesAsTheStoredDbcsSpelling) {
  // Excel localises the formula bar but not the file: `JIS` is the ja-JP
  // spelling of `DBCS`, and Excel stores the call as `DBCS` with
  // function id 215 in both containers. A workbook carrying either
  // spelling must save without degrading, and both come back as the
  // stored spelling.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=JIS(\"ABC\")");
  s.set_cell_formula(1U, 0U, "=DBCS(\"ABC\")");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  const Cell* from_jis = rt.cell_at(0U, 0U);
  ASSERT_NE(from_jis, nullptr);
  EXPECT_EQ(from_jis->formula_text, "=DBCS(\"ABC\")");
  const Cell* from_dbcs = rt.cell_at(1U, 0U);
  ASSERT_NE(from_dbcs, nullptr);
  EXPECT_EQ(from_dbcs->formula_text, "=DBCS(\"ABC\")");
}

TEST(XlsbWriter, HarvestedFuncIdCallsRoundTripWithIdenticalFormulaText) {
  // The ids for these callees were decoded from an Excel-produced
  // workbook (`tests/fixtures/excel/xlsb_func_ids.xlsb`). Each spells a
  // different encoding decision the table drives: a fixed-arity
  // `PtgFunc` (MROUND), a `PtgFuncVar` at its minimum and its maximum
  // arity (WEEKNUM), and an open-ended variadic (GCD). None may degrade
  // to a cached literal, and every formula must come back byte-identical.
  constexpr std::uint32_t kFormulaCount = 5U;
  const char* kFormulas[kFormulaCount] = {
      "=MROUND(17,5)", "=WEEKNUM(43922)", "=WEEKNUM(43922,1)", "=GCD(24,36,60)", "=YEARFRAC(43831,44197,0)",
  };
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    s.set_cell_formula(i, 0U, kFormulas[i]);
  }

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);

  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  for (std::uint32_t i = 0; i < kFormulaCount; ++i) {
    const Cell* cell = rt.cell_at(i, 0U);
    ASSERT_NE(cell, nullptr) << kFormulas[i];
    EXPECT_EQ(cell->formula_text, kFormulas[i]);
  }
}

TEST(XlsbWriter, UnencodableFormulaDowngradesToCachedLiteralAndReportsIt) {
  // An implicit-intersection formula (`@A1:A10`) has no Ptg lowering in
  // the common-token codec (the encoder's `NodeKind::ImplicitIntersection`
  // case still returns `unsupported_node`, unlike `NodeKind::NameRef` --
  // any defined-name reference now lowers to `PtgName`, including
  // references to names that turn out not to be genuine defined names,
  // since `collect_ptg_names` registers every `NameRef` it sees as a
  // hidden placeholder). A single unsupported formula must not make the
  // whole workbook unsaveable: it degrades to its cached literal and the
  // explicit result count records the loss.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=@A1:A10");
  s.set_cell_cached_value(0U, 0U, Value::number(42.0));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 1U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_TRUE(cell->formula_text.empty());
  ASSERT_TRUE(cell->cached_value.is_number());
  EXPECT_DOUBLE_EQ(cell->cached_value.as_number(), 42.0);
}

TEST(XlsbWriter, UnencodableSpillFormulaDowngradesAnchorWithoutPhantoms) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("F"));
  s.set_cell_formula(0U, 0U, "=@A1:A10");
  ASSERT_TRUE(s.commit_spill(0U, 0U, 1U, 2U, {Value::number(11.0), Value::number(12.0)}));

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 1U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Sheet& rt = read_or.value().workbook.sheet(0);
  const Cell* anchor = rt.cell_at(0U, 0U);
  ASSERT_NE(anchor, nullptr);
  EXPECT_TRUE(anchor->formula_text.empty());
  ASSERT_TRUE(anchor->cached_value.is_number());
  EXPECT_DOUBLE_EQ(anchor->cached_value.as_number(), 11.0);
  EXPECT_EQ(rt.cell_at(0U, 1U), nullptr);
}

TEST(XlsbWriter, GeneratedPartsBeatPassthroughOnCollision) {
  // Passthrough trying to ride on the writer's reserved
  // `xl/workbook.bin` path: must be dropped (writer logs a warning,
  // not asserted here), and the round-trip must still succeed.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.sheet(wb.add_sheet("S1"));
  s.set_cell_value(0U, 0U, Value::number(1.0));

  std::vector<PassthroughPart> parts;
  PassthroughPart bogus_workbook;
  bogus_workbook.path = "xl/workbook.bin";
  bogus_workbook.content_type = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
  bogus_workbook.bytes = {0xDE, 0xAD, 0xBE, 0xEF};  // garbage — would break the package if it survived
  parts.push_back(std::move(bogus_workbook));
  wb.set_passthrough_parts(std::move(parts));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  // Workbook is intact: we got our sheet back.
  ASSERT_EQ(read_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(read_or.value().workbook.sheet(0).name(), "S1");
}

TEST(XlsbWriter, SheetVisibilitySurvivesRoundTrip) {
  // A hidden sheet must stay hidden across write -> read. The Sheet model
  // tracks visibility via `view().tab_hidden`, which the writer maps to
  // BrtBundleSh hsState and the reader maps back.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Visible");
  wb.add_sheet("Hidden");
  wb.sheet(1).mutable_view().tab_hidden = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Workbook& rt = read_or.value().workbook;
  ASSERT_EQ(rt.sheet_count(), 2U);
  EXPECT_FALSE(rt.sheet(0).view().tab_hidden);
  EXPECT_TRUE(rt.sheet(1).view().tab_hidden);
}

TEST(XlsbWriter, Date1904SurvivesRoundTrip) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Dates");
  wb.set_date1904(true);

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  EXPECT_TRUE(read_or.value().workbook.date1904());
}

TEST(XlsbWriter, RowAndColumnLayoutSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Layout"));
  sheet.mutable_layout().columns.push_back(ColumnLayout{1U, 3U, 17.25, true, 2U});
  sheet.mutable_layout().row_overrides.push_back(RowLayout{4U, 28.5, true, 3U, true});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const SheetLayout& layout = read_or.value().workbook.sheet(0).layout();
  ASSERT_EQ(layout.columns.size(), 1U);
  EXPECT_EQ(layout.columns[0].first, 1U);
  EXPECT_EQ(layout.columns[0].last, 3U);
  EXPECT_DOUBLE_EQ(layout.columns[0].width, 17.25);
  EXPECT_TRUE(layout.columns[0].has_width);
  // BrtColInfo always carries ixfe, so XLSB canonicalizes a width-only
  // aggregate span to effective style 0 on read.
  EXPECT_TRUE(layout.columns[0].has_style);
  EXPECT_EQ(layout.columns[0].style_xf, 0U);
  EXPECT_TRUE(layout.columns[0].hidden);
  EXPECT_EQ(layout.columns[0].outline_level, 2U);
  ASSERT_EQ(layout.row_overrides.size(), 1U);
  EXPECT_EQ(layout.row_overrides[0].row, 4U);
  EXPECT_DOUBLE_EQ(layout.row_overrides[0].height, 28.5);
  EXPECT_TRUE(layout.row_overrides[0].hidden);
  EXPECT_EQ(layout.row_overrides[0].outline_level, 3U);
}

TEST(XlsbWriter, RowStyleFlagCarriesExplicitZeroAndNonZeroStyles) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("RowStyles"));
  RowLayout style_zero;
  style_zero.row = 2U;
  style_zero.has_style = true;
  style_zero.style_xf = 0U;
  RowLayout style_one;
  style_one.row = 3U;
  style_one.has_style = true;
  style_one.style_xf = 1U;
  sheet.mutable_layout().row_overrides = {style_zero, style_one};

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));

  std::map<std::uint32_t, std::pair<std::uint32_t, std::uint8_t>> raw;
  ByteSpan cursor = SpanOf(sheet_or.value());
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtRowHdr)) {
      continue;
    }
    ByteSpan payload = record_or.value().payload;
    auto row_or = read_u32(payload);
    auto style_or = read_u32(payload);
    ASSERT_TRUE(row_or && style_or);
    ASSERT_TRUE(read_u16(payload));
    ASSERT_TRUE(read_u8(payload));
    auto flags_or = read_u8(payload);
    ASSERT_TRUE(read_u8(payload));
    ASSERT_TRUE(flags_or);
    raw[row_or.value()] = {style_or.value(), flags_or.value()};
  }
  ASSERT_EQ(raw.size(), 2U);
  EXPECT_EQ(raw.at(2U).first, 0U);
  EXPECT_NE(raw.at(2U).second & 0x40U, 0U);
  EXPECT_EQ(raw.at(3U).first, 1U);
  EXPECT_NE(raw.at(3U).second & 0x40U, 0U);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const auto& rows = read_or.value().workbook.sheet(0).layout().row_overrides;
  ASSERT_EQ(rows.size(), 2U);
  EXPECT_TRUE(rows[0].has_style);
  EXPECT_TRUE(rows[1].has_style);
  const auto find_row = [&rows](std::uint32_t row) -> const RowLayout* {
    for (const RowLayout& candidate : rows) {
      if (candidate.row == row) {
        return &candidate;
      }
    }
    return nullptr;
  };
  const RowLayout* loaded_zero = find_row(2U);
  const RowLayout* loaded_one = find_row(3U);
  ASSERT_NE(loaded_zero, nullptr);
  ASSERT_NE(loaded_one, nullptr);
  EXPECT_TRUE(loaded_zero->has_style);
  EXPECT_EQ(loaded_zero->style_xf, 0U);
  EXPECT_TRUE(loaded_one->has_style);
  EXPECT_EQ(loaded_one->style_xf, 1U);
}

TEST(XlsbWriter, MergedRangesSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Merged"));
  sheet.set_cell_value(0U, 0U, Value::text("title"));
  sheet.mutable_merges().push_back(MergeRange{0U, 0U, 1U, 2U});
  sheet.mutable_merges().push_back(MergeRange{4U, 3U, 4U, 5U});

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;
  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const std::vector<MergeRange>& merges = read_or.value().workbook.sheet(0).merges();
  ASSERT_EQ(merges.size(), 2U);
  EXPECT_EQ(merges[0].first_row, 0U);
  EXPECT_EQ(merges[0].first_col, 0U);
  EXPECT_EQ(merges[0].last_row, 1U);
  EXPECT_EQ(merges[0].last_col, 2U);
  EXPECT_EQ(merges[1].first_row, 4U);
  EXPECT_EQ(merges[1].first_col, 3U);
  EXPECT_EQ(merges[1].last_row, 4U);
  EXPECT_EQ(merges[1].last_col, 5U);
}

TEST(XlsbWriter, SheetViewAndFrozenPanesSurviveRoundTrip) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("View"));
  sheet.set_cell_value(0U, 0U, Value::number(1.0));
  SheetView& view = sheet.mutable_view();
  view.zoom_scale = 135U;
  view.freeze_rows = 3U;
  view.freeze_cols = 2U;
  view.show_grid_lines = false;
  view.show_row_col_headers = false;
  view.show_zeros = false;
  view.right_to_left = true;
  view.tab_selected = true;

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto sheet_or = zip.read_entry("xl/worksheets/sheet1.bin");
  ASSERT_TRUE(static_cast<bool>(sheet_or));
  ByteSpan records = SpanOf(sheet_or.value());
  bool saw_pane = false;
  while (records.size != 0U) {
    auto rec_or = read_record(records);
    ASSERT_TRUE(static_cast<bool>(rec_or));
    if (rec_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtPane)) {
      continue;
    }
    saw_pane = true;
    ByteSpan pane = rec_or.value().payload;
    ASSERT_EQ(pane.size, 29U);
    pane.data += 16U;  // two Xnum frozen row/column counts
    pane.size -= 16U;
    auto top_row_or = read_u32(pane);
    auto left_col_or = read_u32(pane);
    auto active_pane_or = read_u32(pane);
    auto flags_or = read_u8(pane);
    ASSERT_TRUE(top_row_or && left_col_or && active_pane_or && flags_or);
    EXPECT_EQ(top_row_or.value(), 3U);
    EXPECT_EQ(left_col_or.value(), 2U);
    EXPECT_EQ(active_pane_or.value(), 0U);
    EXPECT_EQ(flags_or.value(), 0x02U);
  }
  EXPECT_TRUE(saw_pane);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const SheetView& rt = read_or.value().workbook.sheet(0).view();
  EXPECT_EQ(rt.zoom_scale, 135U);
  EXPECT_EQ(rt.freeze_rows, 3U);
  EXPECT_EQ(rt.freeze_cols, 2U);
  EXPECT_FALSE(rt.show_grid_lines);
  EXPECT_FALSE(rt.show_row_col_headers);
  EXPECT_FALSE(rt.show_zeros);
  EXPECT_TRUE(rt.right_to_left);
  EXPECT_TRUE(rt.tab_selected);
}

TEST(XlsbWriter, ReportsDeferredSheetFeatures) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Deferred"));
  Hyperlink hyperlink;
  hyperlink.target = "https://example.com";
  sheet.mutable_hyperlinks().push_back(std::move(hyperlink));
  sheet.mutable_validations().push_back(DataValidation{});
  sheet.set_auto_filter_xml("<autoFilter ref=\"A1:B2\"/>");
  sheet.mutable_view().freeze_rows = 1U;

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  // Auto-filter state remains deferred; hyperlinks and validations are
  // written from the model and therefore no longer inflate this counter.
  EXPECT_EQ(write_or.value().diagnostics.deferred_feature_count, 1U);
}

TEST(XlsbWriter, RejectsInvalidHyperlinkRectangle) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("InvalidHyperlink"));
  Hyperlink inverted;
  inverted.row = 4U;
  inverted.col = 5U;
  inverted.last_row = 3U;
  inverted.last_col = 6U;
  inverted.target = "https://invalid.example";
  sheet.mutable_hyperlinks().push_back(inverted);

  auto inverted_write = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(inverted_write));
  EXPECT_EQ(inverted_write.error().code, FormulonErrorCode::kInvalidArgument);

  sheet.mutable_hyperlinks().clear();
  Hyperlink out_of_grid;
  out_of_grid.row = 0U;
  out_of_grid.col = 0U;
  out_of_grid.last_row = Sheet::kMaxRows;
  out_of_grid.last_col = 0U;
  out_of_grid.target = "https://invalid.example";
  sheet.mutable_hyperlinks().push_back(out_of_grid);
  auto out_of_grid_write = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(out_of_grid_write));
  EXPECT_EQ(out_of_grid_write.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(XlsbWriter, RoundTripsExternalAndInternalHyperlinkRectangles) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Hyperlinks"));
  Hyperlink external;
  external.row = 1U;
  external.col = 2U;
  external.last_row = 3U;
  external.last_col = 4U;
  external.target = "https://external.example/book.xlsx";
  external.location = "#Sheet2!A1";
  external.tooltip = "external";
  external.display = "Open";
  sheet.mutable_hyperlinks().push_back(external);
  Hyperlink internal;
  internal.row = 6U;
  internal.col = 7U;
  internal.last_row = 8U;
  internal.last_col = 9U;
  internal.location = "Sheet1!A1";
  internal.tooltip = "internal";
  internal.display = "Jump";
  sheet.mutable_hyperlinks().push_back(internal);

  auto first_write = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(first_write)) << first_write.error().message << " | " << first_write.error().context;
  auto first_read = read_xlsb(SpanOf(first_write.value()));
  ASSERT_TRUE(static_cast<bool>(first_read)) << first_read.error().message << " | " << first_read.error().context;
  Workbook edited = std::move(first_read.value().workbook);
  ASSERT_EQ(edited.sheet(0).hyperlinks().size(), 2U);
  edited.sheet(0).mutable_hyperlinks()[0].row = 2U;
  edited.sheet(0).mutable_hyperlinks()[0].last_row = 4U;

  auto second_write = write_xlsb(edited);
  ASSERT_TRUE(static_cast<bool>(second_write)) << second_write.error().message << " | " << second_write.error().context;
  auto second_read = read_xlsb(SpanOf(second_write.value()));
  ASSERT_TRUE(static_cast<bool>(second_read)) << second_read.error().message << " | " << second_read.error().context;
  const auto& hyperlinks = second_read.value().workbook.sheet(0).hyperlinks();
  ASSERT_EQ(hyperlinks.size(), 2U);
  EXPECT_EQ(hyperlinks[0].row, 2U);
  EXPECT_EQ(hyperlinks[0].col, 2U);
  EXPECT_EQ(hyperlinks[0].last_row, 4U);
  EXPECT_EQ(hyperlinks[0].last_col, 4U);
  EXPECT_EQ(hyperlinks[0].target, "https://external.example/book.xlsx");
  EXPECT_FALSE(hyperlinks[0].rid.empty());
  EXPECT_EQ(hyperlinks[1].row, 6U);
  EXPECT_EQ(hyperlinks[1].col, 7U);
  EXPECT_EQ(hyperlinks[1].last_row, 8U);
  EXPECT_EQ(hyperlinks[1].last_col, 9U);
  EXPECT_TRUE(hyperlinks[1].rid.empty());
  EXPECT_EQ(hyperlinks[1].location, "Sheet1!A1");
}

TEST(XlsbWriter, ReusesSharedSourceRelationshipIdForMatchingTargets) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("SharedRids"));
  Hyperlink first;
  first.row = 1U;
  first.col = 1U;
  first.last_row = 1U;
  first.last_col = 2U;
  first.target = "https://shared.example/target";
  first.rid = "rIdShared";
  sheet.mutable_hyperlinks().push_back(first);
  Hyperlink second;
  second.row = 3U;
  second.col = 1U;
  second.last_row = 3U;
  second.last_col = 2U;
  second.target = "https://shared.example/target";
  second.rid = "rIdShared";
  sheet.mutable_hyperlinks().push_back(second);

  auto write_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(write_or.value()))));
  auto rels_or = zip.read_entry("xl/worksheets/_rels/sheet1.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  const std::string id_token = "Id=\"rIdShared\"";
  const std::size_t first_id = rels.find(id_token);
  ASSERT_NE(first_id, std::string::npos) << rels;
  EXPECT_EQ(rels.find(id_token, first_id + id_token.size()), std::string::npos) << rels;

  auto read_or = read_xlsb(SpanOf(write_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const auto& hyperlinks = read_or.value().workbook.sheet(0).hyperlinks();
  ASSERT_EQ(hyperlinks.size(), 2U);
  EXPECT_EQ(hyperlinks[0].rid, "rIdShared");
  EXPECT_EQ(hyperlinks[1].rid, "rIdShared");
  EXPECT_EQ(hyperlinks[0].target, hyperlinks[1].target);
}

TEST(XlsbWriter, GeneratesStylesPartForModelledStyles) {
  // XLSX/native workbooks do not have a raw styles.bin passthrough part.  The
  // XLSB writer must therefore materialise the modelled table and expose it
  // through both the content-type override and workbook relationship.
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Styled"));
  sheet.set_cell_value(0U, 0U, Value::number(12.5));

  StylesTable styles;
  FontRecord font;
  font.name = "Aptos";
  font.size = 12.0;
  font.bold = true;
  font.color_argb = 0xFF112233U;
  styles.fonts.push_back(font);
  FillRecord fill;
  fill.pattern = 1U;
  fill.fg_argb = 0xFFFFFF00U;
  styles.fills.push_back(fill);
  styles.borders.push_back(BorderRecord{});
  styles.num_fmt_strings.push_back("0.000");
  styles.num_fmts.push_back(NumFmtRecord{164U, 0U});
  CellXf named;
  named.vertical_align = 2U;
  styles.cell_style_xfs.push_back(named);
  CellXf xf;
  xf.font_index = 0U;
  xf.fill_index = 0U;
  xf.border_index = 0U;
  xf.num_fmt_id = 164U;
  xf.xf_id = 0U;
  xf.apply_number_format = true;
  xf.apply_font = true;
  xf.apply_fill = true;
  styles.cell_xfs.push_back(xf);
  wb.set_styles(std::move(styles));
  sheet.set_cell_xf_index(0U, 0U, 0U);

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  ASSERT_TRUE(zip.has_entry("xl/styles.bin"));
  auto rels_or = zip.read_entry("xl/_rels/workbook.bin.rels");
  ASSERT_TRUE(static_cast<bool>(rels_or));
  const std::string rels(rels_or.value().begin(), rels_or.value().end());
  EXPECT_NE(rels.find("relationships/styles"), std::string::npos);
  EXPECT_NE(rels.find("Target=\"styles.bin\""), std::string::npos);
  auto types_or = zip.read_entry("[Content_Types].xml");
  ASSERT_TRUE(static_cast<bool>(types_or));
  const std::string types(types_or.value().begin(), types_or.value().end());
  EXPECT_NE(types.find("/xl/styles.bin"), std::string::npos);
  EXPECT_NE(types.find("application/vnd.ms-excel.styles"), std::string::npos);

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const StylesTable& rt = read_or.value().workbook.styles();
  ASSERT_EQ(rt.num_fmts.size(), 1U);
  EXPECT_EQ(rt.num_fmts[0].id, 164U);
  ASSERT_EQ(rt.num_fmt_strings.size(), 1U);
  EXPECT_EQ(rt.num_fmt_strings[0], "0.000");
  ASSERT_EQ(rt.cell_xfs.size(), 1U);
  EXPECT_EQ(rt.cell_xfs[0].num_fmt_id, 164U);
  EXPECT_EQ(rt.cell_xfs[0].font_index, 0U);
  EXPECT_EQ(rt.cell_xfs[0].fill_index, 0U);
}

TEST(XlsbWriter, StylesColorsUseTheBrtColorLowValidityBit) {
  StylesTable styles;
  FontRecord font;
  font.name = "Aptos";
  font.color_argb = 0xFF112233U;
  styles.fonts.push_back(font);
  const std::vector<std::uint8_t> bytes = write_styles_bin(styles);

  ByteSpan cursor = SpanOf(bytes);
  while (cursor.size > 0U) {
    auto record_or = read_record(cursor);
    ASSERT_TRUE(static_cast<bool>(record_or));
    if (record_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtFont)) {
      continue;
    }
    const ByteSpan payload = record_or.value().payload;
    ASSERT_GE(payload.size, 20U);
    // u16 height/flags/weight/vertAlign + underline/family/charset/reserved
    // precede BrtColor. RGB is XColorType 2 with fValidRGB in bit 0.
    EXPECT_EQ(payload.data[12], 0x05U);
    EXPECT_EQ(payload.data[16], 0x11U);
    EXPECT_EQ(payload.data[17], 0x22U);
    EXPECT_EQ(payload.data[18], 0x33U);
    EXPECT_EQ(payload.data[19], 0xFFU);
    return;
  }
  FAIL() << "missing BrtFont";
}

TEST(XlsbWriter, PreservesRawStylesPartFromExistingXlsb) {
  // Generated styles are for XLSX/native input only.  When an XLSB reader has
  // retained an opaque style part, that byte stream remains authoritative so
  // unmodelled binary formatting extensions cannot be erased on save.
  Workbook wb = Workbook::create_empty();
  wb.sheet(wb.add_sheet("S")).set_cell_value(0U, 0U, Value::number(1.0));
  StylesTable raw_table;
  raw_table.num_fmt_strings.push_back("0.0000");
  raw_table.num_fmts.push_back(NumFmtRecord{164U, 0U});
  const std::vector<std::uint8_t> raw_bytes = write_styles_bin(raw_table);
  PassthroughPart raw_part;
  raw_part.path = "xl/styles.bin";
  raw_part.content_type = "application/vnd.ms-excel.styles";
  raw_part.bytes = raw_bytes;
  wb.set_passthrough_parts({raw_part});

  // Deliberately make the in-memory table differ from the opaque source.  If
  // the writer regenerated it, the byte comparison below would fail.
  StylesTable changed;
  changed.num_fmt_strings.push_back("0.0%");
  changed.num_fmts.push_back(NumFmtRecord{165U, 0U});
  wb.set_styles(std::move(changed));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));
  ZipReader zip;
  ASSERT_TRUE(static_cast<bool>(zip.open(SpanOf(bytes_or.value()))));
  auto styles_or = zip.read_entry("xl/styles.bin");
  ASSERT_TRUE(static_cast<bool>(styles_or));
  EXPECT_EQ(styles_or.value(), raw_bytes);
}

TEST(XlsbWriter, WholeColumnFormulaSavesWithoutDowngrade) {
  Workbook wb = Workbook::create_empty();
  Sheet& sheet = wb.sheet(wb.add_sheet("Ranges"));
  sheet.set_cell_formula(0U, 0U, "=SUM(A:A)");

  auto write_or = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(write_or)) << write_or.error().message << " | " << write_or.error().context;
  EXPECT_EQ(write_or.value().diagnostics.downgraded_formula_count, 0U);
  auto read_or = read_xlsb(SpanOf(write_or.value().bytes));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  const Cell* cell = read_or.value().workbook.sheet(0).cell_at(0U, 0U);
  ASSERT_NE(cell, nullptr);
  EXPECT_EQ(cell->formula_text, "=SUM(A1:A1048576)");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
