//
// Integration test: a workbook carrying a populated `StylesTable` and
// per-cell `xf_index` references must survive a full
// writer -> reader cycle without loss.

#include <cstdint>
#include <string>
#include <vector>

#include "cell.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/styles_reader.h"
#include "io/zip_reader.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

// Evaluates `src` against `wb` / its first sheet with the formula cell
// anchored at (row, col). Used to confirm a round-tripped protection flag
// actually drives CELL("protect").
Value EvalOn(const Workbook& wb, std::string_view src, std::uint32_t row, std::uint32_t col) {
  Arena parse_arena;
  Arena eval_arena;
  parser::Parser parser(src, parse_arena);
  parser::AstNode* root = parser.parse();
  EXPECT_NE(root, nullptr) << "parse failed for: " << src;
  if (root == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  eval::EvalState state;
  eval::EvalContext ctx(wb, wb.sheet(0), state);
  ctx = ctx.with_formula_cell(row, col);
  return eval::evaluate(*root, eval_arena, eval::default_registry(), ctx);
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
  styles.cell_xfs[0].xf_id = 1;
  styles.cell_xfs[0].apply_font = true;
  styles.cell_xfs[0].apply_alignment = true;
  styles.cell_xfs[0].quote_prefix = true;

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

  ASSERT_FALSE(rt.cell_xfs.empty());
  EXPECT_EQ(rt.cell_xfs[0].xf_id, 1U);
  EXPECT_TRUE(rt.cell_xfs[0].apply_font);
  EXPECT_TRUE(rt.cell_xfs[0].apply_alignment);
  EXPECT_TRUE(rt.cell_xfs[0].quote_prefix);

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

TEST(StylesRoundTrip, EmptyWorkbookNormalizesDefaultNamedStyle) {
  Workbook src = Workbook::create();
  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or));
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or));
  const io::StylesTable& rt = load_or.value().workbook.styles();
  ASSERT_EQ(rt.cell_style_xfs.size(), 1U);
  EXPECT_EQ(rt.cell_style_xfs[0].font_index, 0U);
  ASSERT_EQ(rt.cell_styles.size(), 1U);
  EXPECT_EQ(rt.cell_styles[0].name, "Normal");
  EXPECT_EQ(rt.cell_styles[0].xf_id, 0U);
  EXPECT_EQ(rt.cell_styles[0].builtin_id, 0U);
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

TEST(StylesRoundTrip, CellProtectReflectsRoundTrippedLockedFlag) {
  // A cell whose xf carries `<protection locked="0"/>` must survive the
  // writer -> reader cycle and drive CELL("protect") to 0, while a default
  // (locked) cell stays 1.
  Workbook src = Workbook::create();
  io::StylesTable styles;
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.cell_xfs.emplace_back();  // xf 0: default (locked).
  io::CellXf unlocked{};
  unlocked.has_protection = true;
  unlocked.locked = false;
  styles.cell_xfs.push_back(unlocked);  // xf 1: unlocked.
  src.set_styles(std::move(styles));
  src.sheet(0).set_cell_value(0, 0, Value::number(1.0));  // A1 unlocked.
  src.sheet(0).set_cell_xf_index(0, 0, 1U);
  src.sheet(0).set_cell_value(1, 0, Value::number(2.0));  // A2 default (locked).

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << "save failed: " << save_or.error().message;
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << "read failed: " << load_or.error().message;
  const Workbook& dst = load_or.value().workbook;

  const Value a1 = EvalOn(dst, "=CELL(\"protect\", A1)", 5U, 5U);
  ASSERT_TRUE(a1.is_number());
  EXPECT_DOUBLE_EQ(a1.as_number(), 0.0);

  const Value a2 = EvalOn(dst, "=CELL(\"protect\", A2)", 5U, 5U);
  ASSERT_TRUE(a2.is_number());
  EXPECT_DOUBLE_EQ(a2.as_number(), 1.0);
}

TEST(StylesRoundTrip, PreservesOptionalAlignmentAttributes) {
  Workbook src = Workbook::create();
  io::StylesTable styles;
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.cell_xfs.emplace_back();
  io::CellXf vertical;
  vertical.has_text_rotation = true;
  vertical.text_rotation = 255;
  vertical.has_indent = true;
  vertical.indent = 8;
  vertical.has_relative_indent = true;
  vertical.relative_indent = -5;
  vertical.has_shrink_to_fit = true;
  vertical.shrink_to_fit = false;
  vertical.has_reading_order = true;
  vertical.reading_order = 2;
  styles.cell_xfs.push_back(vertical);
  src.set_styles(std::move(styles));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_value(0, 0, 0, Value::text("縦書き"))));
  ASSERT_TRUE(static_cast<bool>(src.set_cell_xf_index(0, 0, 0, 1)));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message;
  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  const io::StylesTable& rt = load_or.value().workbook.styles();
  ASSERT_GE(rt.cell_xfs.size(), 2U);
  const io::CellXf& round = rt.cell_xfs[1];
  EXPECT_TRUE(round.has_text_rotation);
  EXPECT_EQ(round.text_rotation, 255U);
  EXPECT_TRUE(round.has_indent);
  EXPECT_EQ(round.indent, 8U);
  EXPECT_TRUE(round.has_relative_indent);
  EXPECT_EQ(round.relative_indent, -5);
  EXPECT_TRUE(round.has_shrink_to_fit);
  EXPECT_FALSE(round.shrink_to_fit);
  EXPECT_TRUE(round.has_reading_order);
  EXPECT_EQ(round.reading_order, 2U);
}

TEST(StylesRoundTrip, PreservesCellXfPresenceProjectionAndNormalizesParentId) {
  Workbook src = Workbook::create();
  io::StylesTable styles;
  styles.fonts.emplace_back();
  styles.fills.emplace_back();
  styles.borders.emplace_back();
  styles.cell_style_xfs.emplace_back();
  styles.cell_xfs.emplace_back();

  io::CellXf xf;
  xf.xf_id = 99U;  // normalized to the only named-style xf at serialization.
  xf.apply_number_format = true;
  xf.apply_font = true;
  xf.apply_fill = true;
  xf.apply_border = true;
  xf.apply_alignment = true;
  xf.apply_protection = true;
  xf.quote_prefix = true;
  xf.has_alignment = true;  // explicit child, even though every value is a default.
  xf.has_horizontal_align = true;
  xf.has_vertical_align = true;
  xf.has_wrap_text = true;
  xf.has_justify_last_line = true;
  xf.has_text_rotation = true;
  xf.text_rotation = 0;
  xf.has_indent = true;
  xf.indent = 0;
  xf.has_relative_indent = true;
  xf.relative_indent = -9;
  xf.has_shrink_to_fit = true;
  xf.shrink_to_fit = false;
  xf.has_reading_order = true;
  xf.reading_order = 0;
  xf.has_protection = true;
  xf.locked = true;
  xf.hidden = false;
  styles.cell_xfs.push_back(xf);
  src.set_styles(std::move(styles));

  auto save_or = src.save();
  ASSERT_TRUE(static_cast<bool>(save_or)) << save_or.error().message;
  // Normalization is serialization-only; the in-memory model remains intact.
  ASSERT_GE(src.styles().cell_xfs.size(), 2U);
  EXPECT_EQ(src.styles().cell_xfs[1].xf_id, 99U);

  auto load_or = io::read_ooxml(SpanOf(save_or.value()));
  ASSERT_TRUE(static_cast<bool>(load_or)) << load_or.error().message;
  ASSERT_GE(load_or.value().workbook.styles().cell_xfs.size(), 2U);
  const io::CellXf& round = load_or.value().workbook.styles().cell_xfs[1];
  EXPECT_EQ(round.xf_id, 0U);
  EXPECT_TRUE(round.apply_number_format);
  EXPECT_TRUE(round.apply_font);
  EXPECT_TRUE(round.apply_fill);
  EXPECT_TRUE(round.apply_border);
  EXPECT_TRUE(round.apply_alignment);
  EXPECT_TRUE(round.apply_protection);
  EXPECT_TRUE(round.quote_prefix);
  EXPECT_TRUE(round.has_alignment);
  EXPECT_TRUE(round.has_horizontal_align);
  EXPECT_TRUE(round.has_vertical_align);
  EXPECT_TRUE(round.has_wrap_text);
  EXPECT_TRUE(round.has_justify_last_line);
  EXPECT_TRUE(round.has_text_rotation);
  EXPECT_EQ(round.text_rotation, 0U);
  EXPECT_TRUE(round.has_indent);
  EXPECT_EQ(round.indent, 0U);
  EXPECT_TRUE(round.has_relative_indent);
  EXPECT_EQ(round.relative_indent, -9);
  EXPECT_TRUE(round.has_shrink_to_fit);
  EXPECT_FALSE(round.shrink_to_fit);
  EXPECT_TRUE(round.has_reading_order);
  EXPECT_EQ(round.reading_order, 0U);
  EXPECT_TRUE(round.has_protection);
  EXPECT_TRUE(round.locked);
  EXPECT_FALSE(round.hidden);
}

}  // namespace
}  // namespace formulon
