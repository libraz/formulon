// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Integration test: a workbook carrying a populated `StylesTable` and
// per-cell `xf_index` references must survive a full
// writer -> reader cycle without loss.

#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/styles_reader.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

TEST(StylesRoundTrip, PreservesFontFillBorderAndCellXfs) {
  Workbook src = Workbook::create();

  // Construct a styles table with three records of each kind so the
  // index-based references can drift independently.
  io::StylesTable styles;
  // fonts[0] is the default; index 1 is bold red Meiryo; index 2 is
  // italic underlined.
  styles.fonts.emplace_back();  // default
  io::FontRecord red;
  red.name = "Meiryo";
  red.size = 12.0;
  red.bold = true;
  red.color_argb = 0xFFFF0000U;
  styles.fonts.push_back(red);
  io::FontRecord under;
  under.name = "Calibri";
  under.italic = true;
  under.underline = 1;  // single
  styles.fonts.push_back(under);

  styles.fills.emplace_back();  // default
  io::FillRecord solid_green;
  solid_green.pattern = 1;
  solid_green.fg_argb = 0xFF00FF00U;
  styles.fills.push_back(solid_green);

  styles.borders.emplace_back();  // default
  io::BorderRecord box;
  box.left.style = 1;
  box.right.style = 1;
  box.top.style = 1;
  box.bottom.style = 1;
  styles.borders.push_back(box);

  // Custom number-format string at id 200.
  styles.num_fmt_strings.emplace_back("0.0000");
  io::NumFmtRecord nf;
  nf.id = 200;
  nf.format_string_index = 0;
  styles.num_fmts.push_back(nf);

  // cell_xfs[0] is the default; index 1 references bold red font with
  // green fill and the custom num-fmt; index 2 is the underlined italic
  // with center alignment.
  styles.cell_xfs.emplace_back();  // default
  io::CellXf xf1;
  xf1.font_index = 1;
  xf1.fill_index = 1;
  xf1.border_index = 1;
  xf1.num_fmt_id = 200;
  xf1.horizontal_align = 3;  // right
  styles.cell_xfs.push_back(xf1);
  io::CellXf xf2;
  xf2.font_index = 2;
  xf2.horizontal_align = 2;  // center
  xf2.wrap_text = true;
  styles.cell_xfs.push_back(xf2);

  src.set_styles(std::move(styles));

  // Cells: A1 with value 42, styled xf=1; B2 with text styled xf=2.
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 0, 0, Value::number(42.0))));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_xf_index(0, 0, 0, 1)));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_xf_index(0, 1, 1, 2)));

  // Save and reload.
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  const Workbook& dst = load_or.value().workbook;
  const io::StylesTable& rt = dst.styles();

  // Default record + 2 customs in each list (the round-trip preserves
  // every entry because the writer emits them all; the reader's
  // empty-section fallback inserts a default only when a section is
  // missing).
  ASSERT_GE(rt.fonts.size(), 3U);
  EXPECT_EQ(rt.fonts[1].name, "Meiryo");
  EXPECT_TRUE(rt.fonts[1].bold);
  EXPECT_EQ(rt.fonts[1].color_argb, 0xFFFF0000U);
  EXPECT_TRUE(rt.fonts[2].italic);
  EXPECT_EQ(rt.fonts[2].underline, 1U);

  ASSERT_GE(rt.fills.size(), 2U);
  EXPECT_EQ(rt.fills[1].pattern, 1U);
  EXPECT_EQ(rt.fills[1].fg_argb, 0xFF00FF00U);

  ASSERT_GE(rt.cell_xfs.size(), 3U);
  EXPECT_EQ(rt.cell_xfs[1].num_fmt_id, 200U);
  EXPECT_EQ(rt.cell_xfs[1].font_index, 1U);
  EXPECT_EQ(rt.cell_xfs[1].fill_index, 1U);
  EXPECT_EQ(rt.cell_xfs[1].horizontal_align, 3U);  // right
  EXPECT_EQ(rt.cell_xfs[2].horizontal_align, 2U);  // center
  EXPECT_TRUE(rt.cell_xfs[2].wrap_text);

  // Custom num-fmt string preserved (built-ins never round-trip into
  // the table; only id 200 should be present here).
  ASSERT_EQ(rt.num_fmts.size(), 1U);
  EXPECT_EQ(rt.num_fmts[0].id, 200U);
  EXPECT_EQ(rt.num_fmt_strings[rt.num_fmts[0].format_string_index], "0.0000");

  // Cell-level xf indices propagated.
  const Cell* a1 = dst.sheet(0).cell_at(0, 0);
  ASSERT_NE(a1, nullptr);
  EXPECT_EQ(a1->xf_index, 1U);
  EXPECT_TRUE(a1->cached_value.is_number());
  EXPECT_DOUBLE_EQ(a1->cached_value.as_number(), 42.0);

  const Cell* b2 = dst.sheet(0).cell_at(1, 1);
  ASSERT_NE(b2, nullptr);
  EXPECT_EQ(b2->xf_index, 2U);
}

TEST(StylesRoundTrip, FontSizePreservesFullPrecision) {
  // The font-size writer must use a round-trip-safe format. A size needing
  // more than six significant digits (the old %g default) would otherwise
  // drift across a save/load cycle.
  constexpr double kPreciseSize = 12.345678;

  Workbook src = Workbook::create();
  io::StylesTable styles;
  styles.fonts.emplace_back();  // default
  io::FontRecord precise;
  precise.name = "Calibri";
  precise.size = kPreciseSize;
  styles.fonts.push_back(precise);
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.cell_xfs.emplace_back();
  src.set_styles(std::move(styles));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const io::StylesTable& rt = load_or.value().workbook.styles();
  ASSERT_GE(rt.fonts.size(), 2U);
  EXPECT_DOUBLE_EQ(rt.fonts[1].size, kPreciseSize);
}

TEST(StylesRoundTrip, PreservesNamedCellStyles) {
  Workbook src = Workbook::create();
  io::StylesTable styles;
  // The cellXfs table needs at least one default record; the writer
  // already inserts one for empty input but we exercise the named-style
  // tables alongside the per-cell table here.
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.cell_xfs.emplace_back();

  // Two named-style xf records: default + one with bold font (font_index
  // wraps to 0 because we did not push any extra fonts; we re-use the
  // default font index just to verify the record persists).
  io::CellXf style_xf0;
  io::CellXf style_xf1;
  style_xf1.horizontal_align = 2;  // center
  style_xf1.wrap_text = true;
  styles.cell_style_xfs.push_back(style_xf0);
  styles.cell_style_xfs.push_back(style_xf1);

  // Two named cell styles: built-in "Normal" pointing at xf 0, and a
  // custom user style pointing at xf 1 with hidden=true.
  io::CellStyleRecord normal;
  normal.name = "Normal";
  normal.xf_id = 0;
  normal.builtin_id = 0;  // built-in "Normal"
  styles.cell_styles.push_back(normal);
  io::CellStyleRecord custom;
  custom.name = "Project Heading";
  custom.xf_id = 1;
  custom.hidden = true;  // hidden from the style picker
  custom.custom_builtin = true;
  styles.cell_styles.push_back(custom);

  src.set_styles(std::move(styles));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;

  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;

  const io::StylesTable& rt = load_or.value().workbook.styles();

  ASSERT_EQ(rt.cell_style_xfs.size(), 2U);
  EXPECT_EQ(rt.cell_style_xfs[1].horizontal_align, 2U);
  EXPECT_TRUE(rt.cell_style_xfs[1].wrap_text);

  ASSERT_EQ(rt.cell_styles.size(), 2U);
  EXPECT_EQ(rt.cell_styles[0].name, "Normal");
  EXPECT_EQ(rt.cell_styles[0].xf_id, 0U);
  EXPECT_EQ(rt.cell_styles[0].builtin_id, 0U);
  EXPECT_FALSE(rt.cell_styles[0].hidden);

  EXPECT_EQ(rt.cell_styles[1].name, "Project Heading");
  EXPECT_EQ(rt.cell_styles[1].xf_id, 1U);
  EXPECT_EQ(rt.cell_styles[1].builtin_id, io::CellStyleRecord::kBuiltinIdNone);
  EXPECT_TRUE(rt.cell_styles[1].hidden);
  EXPECT_TRUE(rt.cell_styles[1].custom_builtin);
}

TEST(StylesRoundTrip, EmptyWorkbookHasNoNamedStyles) {
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const io::StylesTable& rt = load_or.value().workbook.styles();
  EXPECT_TRUE(rt.cell_styles.empty());
  EXPECT_TRUE(rt.cell_style_xfs.empty());
}

TEST(StylesRoundTrip, EmptyWorkbookHasDefaultStyles) {
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const io::StylesTable& rt = load_or.value().workbook.styles();
  // The minimal styles document the writer produces for a default
  // workbook carries one default font / fill / border / cellXf so
  // `xf_index = 0` always resolves.
  EXPECT_EQ(rt.fonts.size(), 1U);
  EXPECT_EQ(rt.fills.size(), 1U);
  EXPECT_EQ(rt.borders.size(), 1U);
  EXPECT_EQ(rt.cell_xfs.size(), 1U);
}

}  // namespace
}  // namespace formulon
