// Cross-format XLSB/OOXML symmetry tests.

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "cell.h"
#include "defined_name.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "styles.h"
#include "support/roundtrip_symmetry.h"
#include "value.h"
#include "xlsb_roundtrip_symmetry_test_helpers.h"

namespace formulon {
namespace {
using namespace xlsb_roundtrip_test_support;
TEST(XlsbRkEncoding, PredicateOnlyAcceptsValuesWhoseEncodingDecodesToTheSameBits) {
  std::uint64_t state = 0x9E3779B97F4A7C15ULL;  // any fixed seed; the sweep must be reproducible
  auto next = [&state]() {
    state = state * 6364136223846793005ULL + 1442695040888963407ULL;
    return state;
  };
  int accepted = 0;
  for (int i = 0; i < 20000; ++i) {
    // Currency-shaped magnitudes (cents over a 30-bit range), their two
    // immediate neighbours -- where multiplying by 100 rounds back onto
    // the exact cent count and the x100 form therefore looks applicable
    // while decoding to a different double -- and a spread of arbitrary
    // bit patterns.
    const double cents = static_cast<double>(static_cast<std::int64_t>(next() % 1000000000ULL) - 500000000);
    const double base = cents / 100.0;
    const double candidates[] = {base, std::nextafter(base, std::numeric_limits<double>::infinity()),
                                 std::nextafter(base, -std::numeric_limits<double>::infinity()), cents / 3.0, cents};
    for (const double v : candidates) {
      if (!io::xlsb::rk_round_trips_value(v)) {
        continue;
      }
      ++accepted;
      std::vector<std::uint8_t> bytes;
      io::xlsb::emit_rk_number(bytes, v);
      ASSERT_EQ(bytes.size(), 4U);
      const std::uint32_t rk = static_cast<std::uint32_t>(bytes[0]) | (static_cast<std::uint32_t>(bytes[1]) << 8) |
                               (static_cast<std::uint32_t>(bytes[2]) << 16) |
                               (static_cast<std::uint32_t>(bytes[3]) << 24);
      EXPECT_EQ(BitsOf(io::xlsb::decode_rk_number(rk)), BitsOf(v)) << "value=" << v;
    }
  }
  EXPECT_GT(accepted, 0) << "sweep never exercised the accepting branch";
}
TEST(XlsbWriteReadSymmetry, XlsxToXlsbToXlsxPreservesNumericBitPatterns) {
  const double kValues[] = {3611469.5700000003, -4123191.7399999998};
  for (const double v : kValues) {
    EXPECT_FALSE(io::xlsb::rk_round_trips_value(v)) << "value=" << v;
  }

  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  for (std::size_t i = 0; i < std::size(kValues); ++i) {
    wb.sheet(0).set_cell_value(static_cast<std::uint32_t>(i), 0U, Value::number(kValues[i]));
  }

  auto xlsx_in = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx_in)) << "write_ooxml failed: " << xlsx_in.error().message;
  auto from_xlsx = io::read_ooxml(test::span_of(xlsx_in.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << "read_ooxml failed: " << from_xlsx.error().message;

  auto xlsb = io::xlsb::write_xlsb(from_xlsx.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsb)) << "write_xlsb failed: " << xlsb.error().message;
  auto from_xlsb = io::xlsb::read_xlsb(test::span_of(xlsb.value()));
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << "read_xlsb failed: " << from_xlsb.error().message;

  auto xlsx_out = io::write_ooxml(from_xlsb.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx_out)) << "write_ooxml failed: " << xlsx_out.error().message;
  auto final_wb = io::read_ooxml(test::span_of(xlsx_out.value()));
  ASSERT_TRUE(static_cast<bool>(final_wb)) << "read_ooxml failed: " << final_wb.error().message;

  const Sheet& s = final_wb.value().workbook.sheet(0);
  for (std::size_t i = 0; i < std::size(kValues); ++i) {
    const Cell* c = s.cell_at(static_cast<std::uint32_t>(i), 0U);
    ASSERT_NE(c, nullptr) << "row " << i;
    ASSERT_TRUE(c->cached_value.is_number()) << "row " << i;
    EXPECT_EQ(BitsOf(c->cached_value.as_number()), BitsOf(kValues[i])) << "row " << i;
  }
}
TEST(XlsbWriteReadSymmetry, RowOverridesSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  // A cell keeps the rows from being dropped as empty, and each override
  // engages a different `BrtRowHdr` flag: a custom height (fUnsynced), the
  // hidden bit, and an outline level.
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));
  source.sheet(0).set_cell_value(4U, 0U, Value::number(2.0));
  RowLayout tall;
  tall.row = 0U;
  tall.height = 33.75;
  tall.has_height = true;
  RowLayout hidden;
  hidden.row = 2U;
  hidden.hidden = true;
  RowLayout grouped;
  grouped.row = 4U;
  grouped.outline_level = 2U;
  source.sheet(0).mutable_layout().row_overrides = {tall, hidden, grouped};

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<RowLayout>& rb = via_xlsb.sheet(0).layout().row_overrides;
  const std::vector<RowLayout>& rx = via_xlsx.sheet(0).layout().row_overrides;
  ASSERT_EQ(rb.size(), 3U) << "the xlsb path lost a row override";
  ASSERT_EQ(rx.size(), 3U) << "the xlsx path lost a row override";
  for (std::size_t i = 0; i < rb.size(); ++i) {
    SCOPED_TRACE("row override " + std::to_string(i));
    EXPECT_EQ(rb[i].row, rx[i].row);
    EXPECT_EQ(rb[i].has_height, rx[i].has_height);
    // `miyRw` is twips, so a height survives to 1/20 of a point.
    EXPECT_NEAR(rb[i].height, rx[i].height, 1.0 / 20.0);
    EXPECT_EQ(rb[i].hidden, rx[i].hidden);
    EXPECT_EQ(rb[i].outline_level, rx[i].outline_level);
  }
  // The values themselves, not just their agreement: two equally broken
  // paths would satisfy the comparison above on their own.
  EXPECT_TRUE(rb[0].has_height);
  EXPECT_NEAR(rb[0].height, 33.75, 1.0 / 20.0);
  EXPECT_TRUE(rb[1].hidden);
  EXPECT_EQ(rb[2].outline_level, 2U);
}
TEST(XlsbWriteReadSymmetry, BorderRecordContentsSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));

  StylesTable styles;
  styles.fonts.push_back(FontRecord{});
  styles.fills.push_back(FillRecord{});
  styles.borders.push_back(BorderRecord{});
  BorderRecord boxed;
  boxed.left.style = 1U;  // thin
  boxed.left.color.kind = ColorSpec::Kind::kRgb;
  boxed.left.color.rgb = 0xFF0000FFU;
  boxed.left.color_argb = 0xFF0000FFU;
  boxed.bottom.style = 2U;  // medium
  boxed.bottom.color.kind = ColorSpec::Kind::kRgb;
  boxed.bottom.color.rgb = 0xFFFF0000U;
  boxed.bottom.color_argb = 0xFFFF0000U;
  boxed.diagonal.style = 3U;  // dashed
  boxed.diagonal.color.kind = ColorSpec::Kind::kRgb;
  boxed.diagonal.color.rgb = 0xFF00FF00U;
  boxed.diagonal.color_argb = 0xFF00FF00U;
  boxed.diagonal_up = true;
  styles.borders.push_back(boxed);
  CellXf plain;
  CellXf bordered;
  bordered.border_index = 1U;
  bordered.apply_border = true;
  styles.cell_xfs = {plain, bordered};
  source.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(source.set_cell_xf_index(0U, 0U, 0U, 1U)));

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<BorderRecord>& bb = via_xlsb.styles().borders;
  const std::vector<BorderRecord>& bx = via_xlsx.styles().borders;
  ASSERT_GT(bb.size(), 1U) << "the xlsb path lost the non-default border";
  ASSERT_GT(bx.size(), 1U) << "the xlsx path lost the non-default border";
  ASSERT_EQ(bb.size(), bx.size());
  for (std::size_t i = 0; i < bb.size(); ++i) {
    SCOPED_TRACE("border index " + std::to_string(i));
    const BorderRecord& b = bb[i];
    const BorderRecord& x = bx[i];
    EXPECT_EQ(b.diagonal_up, x.diagonal_up);
    EXPECT_EQ(b.diagonal_down, x.diagonal_down);
    const BorderSide* b_sides[] = {&b.left, &b.right, &b.top, &b.bottom, &b.diagonal};
    const BorderSide* x_sides[] = {&x.left, &x.right, &x.top, &x.bottom, &x.diagonal};
    const char* names[] = {"left", "right", "top", "bottom", "diagonal"};
    for (std::size_t side = 0; side < std::size(b_sides); ++side) {
      SCOPED_TRACE(names[side]);
      EXPECT_EQ(b_sides[side]->style, x_sides[side]->style);
      ExpectColorSpecEqual(b_sides[side]->color, x_sides[side]->color);
    }
  }
  // The styles, not merely their agreement: the authored record has to come
  // back with the three sides it declared.
  EXPECT_EQ(bb[1].left.style, 1U);
  EXPECT_EQ(bb[1].bottom.style, 2U);
  EXPECT_EQ(bb[1].diagonal.style, 3U);
  EXPECT_TRUE(bb[1].diagonal_up);
}
TEST(XlsbWriteReadSymmetry, AlignmentAndApplyFlagsSurviveBothFormatsAlike) {
  Workbook source = Workbook::create_empty();
  source.add_sheet("Sheet1");
  source.sheet(0).set_cell_value(0U, 0U, Value::number(1.0));

  StylesTable styles;
  styles.fonts.push_back(FontRecord{});
  styles.fills.push_back(FillRecord{});
  styles.borders.push_back(BorderRecord{});
  CellXf plain;
  CellXf decorated;
  decorated.horizontal_align = 2U;  // center
  decorated.vertical_align = 0U;    // top -- not the schema default
  decorated.wrap_text = true;
  decorated.justify_last_line = true;
  decorated.shrink_to_fit = true;
  decorated.has_shrink_to_fit = true;
  decorated.reading_order = 2U;  // right-to-left
  decorated.has_reading_order = true;
  decorated.text_rotation = 45U;
  decorated.has_text_rotation = true;
  decorated.indent = 3U;
  decorated.has_indent = true;
  decorated.quote_prefix = true;
  decorated.has_protection = true;
  decorated.locked = false;
  decorated.hidden = true;
  decorated.apply_number_format = true;
  decorated.apply_font = true;
  decorated.apply_fill = true;
  decorated.apply_border = true;
  decorated.apply_alignment = true;
  decorated.apply_protection = true;
  styles.cell_xfs = {plain, decorated};
  source.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(source.set_cell_xf_index(0U, 0U, 0U, 1U)));

  Workbook via_xlsb = Workbook::create_empty();
  Workbook via_xlsx = Workbook::create_empty();
  ASSERT_TRUE(ThroughXlsb(source, &via_xlsb));
  ASSERT_TRUE(ThroughXlsx(source, &via_xlsx));

  const std::vector<CellXf>& xb = via_xlsb.styles().cell_xfs;
  const std::vector<CellXf>& xx = via_xlsx.styles().cell_xfs;
  ASSERT_GT(xb.size(), 1U) << "the xlsb path lost the decorated xf";
  ASSERT_GT(xx.size(), 1U) << "the xlsx path lost the decorated xf";
  ASSERT_EQ(xb.size(), xx.size());
  for (std::size_t i = 0; i < xb.size(); ++i) {
    SCOPED_TRACE("cellXf index " + std::to_string(i));
    const CellXf& b = xb[i];
    const CellXf& x = xx[i];
    EXPECT_EQ(b.horizontal_align, x.horizontal_align);
    EXPECT_EQ(b.vertical_align, x.vertical_align);
    EXPECT_EQ(b.wrap_text, x.wrap_text);
    EXPECT_EQ(b.justify_last_line, x.justify_last_line);
    EXPECT_EQ(b.shrink_to_fit, x.shrink_to_fit);
    EXPECT_EQ(b.reading_order, x.reading_order);
    EXPECT_EQ(b.text_rotation, x.text_rotation);
    EXPECT_EQ(b.indent, x.indent);
    EXPECT_EQ(b.quote_prefix, x.quote_prefix);
    EXPECT_EQ(b.has_protection, x.has_protection);
    EXPECT_EQ(b.locked, x.locked);
    EXPECT_EQ(b.hidden, x.hidden);
    EXPECT_EQ(b.apply_number_format, x.apply_number_format);
    EXPECT_EQ(b.apply_font, x.apply_font);
    EXPECT_EQ(b.apply_fill, x.apply_fill);
    EXPECT_EQ(b.apply_border, x.apply_border);
    EXPECT_EQ(b.apply_alignment, x.apply_alignment);
    EXPECT_EQ(b.apply_protection, x.apply_protection);
  }
  // The authored values themselves: two equally lossy paths would satisfy
  // the comparison above on their own.
  const CellXf& b = xb[1];
  EXPECT_EQ(b.horizontal_align, 2U);
  EXPECT_EQ(b.vertical_align, 0U);
  EXPECT_TRUE(b.wrap_text);
  EXPECT_TRUE(b.justify_last_line);
  EXPECT_TRUE(b.shrink_to_fit);
  EXPECT_EQ(b.reading_order, 2U);
  EXPECT_EQ(b.text_rotation, 45U);
  EXPECT_EQ(b.indent, 3U);
  EXPECT_TRUE(b.quote_prefix);
  EXPECT_TRUE(b.has_protection);
  EXPECT_FALSE(b.locked);
  EXPECT_TRUE(b.hidden);
  EXPECT_TRUE(b.apply_number_format);
  EXPECT_TRUE(b.apply_font);
  EXPECT_TRUE(b.apply_fill);
  EXPECT_TRUE(b.apply_border);
  EXPECT_TRUE(b.apply_alignment);
  EXPECT_TRUE(b.apply_protection);
}

}  // namespace
}  // namespace formulon
