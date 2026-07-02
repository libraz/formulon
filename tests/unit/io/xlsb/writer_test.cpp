// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// End-to-end round-trip tests for the XLSB package writer.
// `read_xlsb(write_xlsb(wb))` must reproduce every cell value `wb`
// carried in for the in-scope literal kinds, plus passthrough parts
// and Bundle 4.1 stub-formula cells.

#include "io/xlsb/writer.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/passthrough_part.h"
#include "io/xlsb/reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
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
  Sheet& s = wb.add_sheet("Solo");
  s.set_cell_value(0U, 0U, Value::number(1.0));
  s.set_cell_value(0U, 1U, Value::boolean(false));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;
  ASSERT_EQ(read_or.value().workbook.sheet_count(), 1U);
  EXPECT_EQ(read_or.value().cells_read, 2U);
}

TEST(XlsbWriter, PassthroughPartsRoundTripVerbatim) {
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("S1");
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
  wb.set_passthrough_parts(std::move(parts));

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or));

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  bool found_theme = false;
  for (const PassthroughPart& part : read_or.value().unknown_parts) {
    if (part.path == "xl/theme/theme1.xml") {
      found_theme = true;
      EXPECT_EQ(part.content_type, "application/vnd.openxmlformats-officedocument.theme+xml");
      const std::string round_tripped(part.bytes.begin(), part.bytes.end());
      EXPECT_EQ(round_tripped, body);
    }
  }
  EXPECT_TRUE(found_theme);
}

TEST(XlsbWriter, RealFormulaRoundTripsAsFormulaCell) {
  // An engine-authored formula encodes to a Ptg stream, survives the
  // write, and decodes back to the same formula text on read.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("F");
  s.set_cell_formula(2U, 3U, "=A1+B2*3");

  auto bytes_or = write_xlsb(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << bytes_or.error().message << " | " << bytes_or.error().context;

  auto read_or = read_xlsb(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << read_or.error().message << " | " << read_or.error().context;

  const Cell* c = read_or.value().workbook.sheet(0).cell_at(2U, 3U);
  ASSERT_NE(c, nullptr);
  EXPECT_EQ(c->formula_text, "=A1+B2*3");
}

TEST(XlsbWriter, UnencodableFormulaReturnsErrorNotSilentLiteral) {
  // An implicit-intersection formula (`@A1:A10`) has no Ptg lowering in
  // the common-token codec (the encoder's `NodeKind::ImplicitIntersection`
  // case still returns `unsupported_node`, unlike `NodeKind::NameRef` --
  // any defined-name reference now lowers to `PtgName`, including
  // references to names that turn out not to be genuine defined names,
  // since `collect_ptg_names` registers every `NameRef` it sees as a
  // hidden placeholder). The writer must surface the encode failure as
  // an `Expected` failure rather than silently dropping it to a literal
  // cell.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("F");
  s.set_cell_formula(0U, 0U, "=@A1:A10");

  auto bytes_or = write_xlsb(wb);
  ASSERT_FALSE(static_cast<bool>(bytes_or));
  EXPECT_EQ(bytes_or.error().code, FormulonErrorCode::kIoXlsbUnsupportedPtg);
}

TEST(XlsbWriter, GeneratedPartsBeatPassthroughOnCollision) {
  // Passthrough trying to ride on the writer's reserved
  // `xl/workbook.bin` path: must be dropped (writer logs a warning,
  // not asserted here), and the round-trip must still succeed.
  Workbook wb = Workbook::create_empty();
  Sheet& s = wb.add_sheet("S1");
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

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
