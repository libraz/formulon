// XLSB writer round-trip tests: literal cells and workbook defaults.

#include "io/xlsb/writer.h"

#include <cstdint>
#include <limits>
#include <string>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "gtest/gtest.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/zip_reader.h"
#include "print/pagination.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"
namespace formulon {
namespace io {
namespace xlsb {
namespace {

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
  EXPECT_EQ(miy_or.value(), 291U);  // engine default 102/7 pt in twips
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
  defaults.custom_height = true;

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

  defaults.custom_height = false;
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
  // The record always carries a row height: the engine default comes back
  // as an explicit default quantized to twips (102/7 pt -> 291 twips).
  EXPECT_TRUE(round_tripped.format_defaults().has_default_row_height);
  EXPECT_DOUBLE_EQ(round_tripped.format_defaults().default_row_height, 291.0 / 20.0);
  EXPECT_FALSE(round_tripped.format_defaults().custom_height);
  EXPECT_FALSE(round_tripped.format_defaults().has_default_col_width);

  auto after_or = print::paginate(read_or.value().workbook, 0U);
  ASSERT_TRUE(static_cast<bool>(after_or)) << after_or.error().message;
  EXPECT_EQ(after_or.value().page_count, before_or.value().page_count);
  EXPECT_EQ(after_or.value().h_breaks, before_or.value().h_breaks);
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
