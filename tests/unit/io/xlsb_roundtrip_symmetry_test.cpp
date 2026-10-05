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
TEST(XlsbCrossFormatSymmetry, SheetStructureMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  ASSERT_EQ(xlsb.sheet_count(), xlsx.sheet_count());
  for (std::size_t i = 0; i < xlsb.sheet_count(); ++i) {
    EXPECT_EQ(xlsb.sheet(i).name(), xlsx.sheet(i).name()) << "sheet index " << i;
  }
}
TEST(XlsbCrossFormatSymmetry, DataSheetValuesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const Sheet& sb = xlsb.sheet(0);
  const Sheet& sx = xlsx.sheet(0);
  // A1:A3 are text keys; B1:B3 are the numeric column.
  for (std::uint32_t r = 0; r < 3; ++r) {
    const Cell* ab = sb.cell_at(r, 0);
    const Cell* ax = sx.cell_at(r, 0);
    ASSERT_NE(ab, nullptr);
    ASSERT_NE(ax, nullptr);
    ASSERT_TRUE(ab->cached_value.is_text());
    ASSERT_TRUE(ax->cached_value.is_text());
    EXPECT_EQ(ab->cached_value.as_text(), ax->cached_value.as_text()) << "A" << (r + 1);

    const Cell* bb = sb.cell_at(r, 1);
    const Cell* bx = sx.cell_at(r, 1);
    ASSERT_NE(bb, nullptr);
    ASSERT_NE(bx, nullptr);
    ASSERT_TRUE(bb->cached_value.is_number());
    ASSERT_TRUE(bx->cached_value.is_number());
    EXPECT_DOUBLE_EQ(bb->cached_value.as_number(), bx->cached_value.as_number()) << "B" << (r + 1);
  }
}
TEST(XlsbCrossFormatSymmetry, FormulaTextMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // F1 holds an XLOOKUP; both readers must restore identical formula text
  // (no `_xlfn.` prefix drift between the binary and OOXML paths).
  const Cell* fb = xlsb.sheet(0).cell_at(0, 5);
  const Cell* fx = xlsx.sheet(0).cell_at(0, 5);
  ASSERT_NE(fb, nullptr);
  ASSERT_NE(fx, nullptr);
  EXPECT_EQ(fb->formula_text, fx->formula_text);
  EXPECT_FALSE(fb->formula_text.empty());
}
TEST(XlsbCrossFormatSymmetry, StyleIndexMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  // D3 is a bold-red styled cell. Both readers must resolve it to the same
  // style (xf) index -- a cross-format check on the styles.bin vs styles.xml
  // parse producing equivalent style tables.
  const Cell* db = xlsb.sheet(0).cell_at(2, 3);
  const Cell* dx = xlsx.sheet(0).cell_at(2, 3);
  ASSERT_NE(db, nullptr);
  ASSERT_NE(dx, nullptr);
  EXPECT_EQ(db->xf_index, dx->xf_index);
  EXPECT_NE(db->xf_index, 0U) << "D3 should carry a non-default style";
}
TEST(XlsbCrossFormatSymmetry, FontRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<FontRecord>& fb = xlsb.styles().fonts;
  const std::vector<FontRecord>& fx = xlsx.styles().fonts;
  ASSERT_EQ(fb.size(), fx.size());

  // The fixture has to carry a font that differs from the default record in
  // more than one attribute, or every field comparison below would hold on a
  // table of blanks.
  bool saw_non_default = false;
  for (const FontRecord& font : fx) {
    if (font.bold && font.color.kind == ColorSpec::Kind::kRgb) {
      saw_non_default = true;
    }
  }
  ASSERT_TRUE(saw_non_default) << "fixture carries no bold, explicitly coloured font";

  for (std::size_t i = 0; i < fb.size(); ++i) {
    SCOPED_TRACE("font index " + std::to_string(i));
    EXPECT_EQ(fb[i].name, fx[i].name);
    EXPECT_DOUBLE_EQ(fb[i].size, fx[i].size);
    EXPECT_EQ(fb[i].bold, fx[i].bold);
    EXPECT_EQ(fb[i].has_bold, fx[i].has_bold);
    EXPECT_EQ(fb[i].italic, fx[i].italic);
    EXPECT_EQ(fb[i].has_italic, fx[i].has_italic);
    EXPECT_EQ(fb[i].strike, fx[i].strike);
    EXPECT_EQ(fb[i].has_strike, fx[i].has_strike);
    EXPECT_EQ(fb[i].underline, fx[i].underline);
    EXPECT_EQ(fb[i].vert_align, fx[i].vert_align);
    EXPECT_EQ(fb[i].has_family, fx[i].has_family);
    EXPECT_EQ(fb[i].family, fx[i].family);
    EXPECT_EQ(fb[i].has_charset, fx[i].has_charset);
    EXPECT_EQ(fb[i].charset, fx[i].charset);
    EXPECT_EQ(fb[i].scheme, fx[i].scheme);
    ExpectColorSpecEqual(fb[i].color, fx[i].color);
  }
}
TEST(XlsbCrossFormatSymmetry, FillRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<FillRecord>& fb = xlsb.styles().fills;
  const std::vector<FillRecord>& fx = xlsx.styles().fills;
  ASSERT_EQ(fb.size(), fx.size());

  // A patterned fill with an explicit foreground colour is the case that
  // distinguishes the two decode paths; the two placeholder fills Excel
  // always writes first would not.
  bool saw_coloured_pattern = false;
  for (const FillRecord& fill : fx) {
    if (fill.pattern != 0U && fill.fg.kind != ColorSpec::Kind::kNone) {
      saw_coloured_pattern = true;
    }
  }
  ASSERT_TRUE(saw_coloured_pattern) << "fixture carries no coloured pattern fill";

  for (std::size_t i = 0; i < fb.size(); ++i) {
    SCOPED_TRACE("fill index " + std::to_string(i));
    EXPECT_EQ(fb[i].pattern, fx[i].pattern);
    ExpectColorSpecEqual(fb[i].fg, fx[i].fg);
    ExpectColorSpecEqual(fb[i].bg, fx[i].bg);
  }
}
TEST(XlsbCrossFormatSymmetry, CellXfRecordContentsMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const StylesTable& tb = xlsb.styles();
  const StylesTable& tx = xlsx.styles();
  ASSERT_EQ(tb.cell_xfs.size(), tx.cell_xfs.size());

  // An xf table of nothing but copies of the default record would satisfy
  // every field comparison without exercising the decode.
  bool saw_non_default = false;
  for (const CellXf& xf : tx.cell_xfs) {
    if (xf.font_index != 0U || xf.fill_index != 0U || xf.num_fmt_id != 0U) {
      saw_non_default = true;
    }
  }
  ASSERT_TRUE(saw_non_default) << "fixture xf table selects only default records";

  // The fixture is authored in ja-JP, where `vertical="center"` is the
  // default Excel applies to every cell, so an xf table that came back
  // bottom-aligned is the shape this comparison exists to reject.
  bool saw_alignment = false;
  bool saw_apply_flag = false;
  for (const CellXf& xf : tx.cell_xfs) {
    if (HasAlignment(xf)) {
      saw_alignment = true;
    }
    if (xf.apply_number_format || xf.apply_font || xf.apply_fill) {
      saw_apply_flag = true;
    }
  }
  ASSERT_TRUE(saw_alignment) << "fixture xf table carries no alignment";
  ASSERT_TRUE(saw_apply_flag) << "fixture xf table sets no apply flag";

  for (std::size_t i = 0; i < tb.cell_xfs.size(); ++i) {
    SCOPED_TRACE("cellXf index " + std::to_string(i));
    const CellXf& b = tb.cell_xfs[i];
    const CellXf& x = tx.cell_xfs[i];
    // The selector fields, which are what makes an xf name one font, fill,
    // border and number format rather than another.
    EXPECT_EQ(b.font_index, x.font_index);
    EXPECT_EQ(b.fill_index, x.fill_index);
    EXPECT_EQ(b.border_index, x.border_index);
    EXPECT_EQ(b.num_fmt_id, x.num_fmt_id);
    EXPECT_EQ(b.xf_id, x.xf_id);
    // The `apply*` set, which decides whether the xf's own font / format
    // wins over the named style it inherits from.
    EXPECT_EQ(b.apply_number_format, x.apply_number_format);
    EXPECT_EQ(b.apply_font, x.apply_font);
    EXPECT_EQ(b.apply_fill, x.apply_fill);
    EXPECT_EQ(b.apply_border, x.apply_border);
    EXPECT_EQ(b.apply_alignment, x.apply_alignment);
    EXPECT_EQ(b.apply_protection, x.apply_protection);
    // Alignment and protection are compared on their effective values and
    // on the presence predicates the writer consults, not on the raw
    // `has_*` bits: those record how OOXML spelled a value, and `BrtXF`
    // states every field unconditionally, so an XLSB-sourced xf derives
    // presence from the value differing from its schema default.
    EXPECT_EQ(b.horizontal_align, x.horizontal_align);
    EXPECT_EQ(b.vertical_align, x.vertical_align);
    EXPECT_EQ(b.wrap_text, x.wrap_text);
    EXPECT_EQ(b.justify_last_line, x.justify_last_line);
    EXPECT_EQ(b.shrink_to_fit, x.shrink_to_fit);
    EXPECT_EQ(b.reading_order, x.reading_order);
    EXPECT_EQ(b.text_rotation, x.text_rotation);
    EXPECT_EQ(b.indent, x.indent);
    EXPECT_EQ(b.quote_prefix, x.quote_prefix);
    EXPECT_EQ(b.locked, x.locked);
    EXPECT_EQ(b.hidden, x.hidden);
    EXPECT_EQ(b.has_protection, x.has_protection);
    EXPECT_EQ(HasAlignment(b), HasAlignment(x));
    EXPECT_EQ(HasHorizontalAlign(b), HasHorizontalAlign(x));
    EXPECT_EQ(HasVerticalAlign(b), HasVerticalAlign(x));
    EXPECT_EQ(HasWrapText(b), HasWrapText(x));
    EXPECT_EQ(HasJustifyLastLine(b), HasJustifyLastLine(x));
  }
}
TEST(XlsbToOoxmlStyles, AlignmentAndApplyFlagsReachTheSavedStylesPart) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));

  auto from_xlsb = io::write_ooxml(xlsb);
  ASSERT_TRUE(static_cast<bool>(from_xlsb)) << "write_ooxml failed: " << from_xlsb.error().message;
  auto from_xlsx = io::write_ooxml(xlsx);
  ASSERT_TRUE(static_cast<bool>(from_xlsx)) << "write_ooxml failed: " << from_xlsx.error().message;

  const std::string xlsb_block = CellXfsBlockOfSavedPackage(from_xlsb.value());
  const std::string xlsx_block = CellXfsBlockOfSavedPackage(from_xlsx.value());
  ASSERT_FALSE(xlsb_block.empty()) << "saved package carries no <cellXfs>";

  // The literals first: an equality against an equally empty block would
  // hold without any of this reaching the file.
  EXPECT_EQ(CountOccurrences(xlsb_block, "<alignment vertical=\"center\"/>"), 6U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyNumberFormat=\"1\""), 3U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyFont=\"1\""), 1U) << xlsb_block;
  EXPECT_EQ(CountOccurrences(xlsb_block, "applyFill=\"1\""), 1U) << xlsb_block;
  EXPECT_EQ(xlsb_block, xlsx_block);
}
TEST(XlsbCrossFormatSymmetry, CustomNumberFormatCodesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const auto codes_by_id = [](const StylesTable& table) {
    std::vector<std::pair<std::uint16_t, std::string>> out;
    for (const NumFmtRecord& rec : table.num_fmts) {
      if (rec.format_string_index >= table.num_fmt_strings.size()) {
        ADD_FAILURE() << "numFmt id " << rec.id << " interns past the string table";
        continue;
      }
      out.emplace_back(rec.id, table.num_fmt_strings[rec.format_string_index]);
    }
    std::sort(out.begin(), out.end());
    return out;
  };
  const auto b = codes_by_id(xlsb.styles());
  const auto x = codes_by_id(xlsx.styles());
  ASSERT_FALSE(x.empty()) << "fixture declares no custom number format";
  EXPECT_EQ(b, x);
}
TEST(XlsbCrossFormatSymmetry, ColumnLayoutMatches) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<ColumnLayout>& cb = xlsb.sheet(0).layout().columns;
  const std::vector<ColumnLayout>& cx = xlsx.sheet(0).layout().columns;
  ASSERT_EQ(cb.size(), cx.size());
  ASSERT_FALSE(cx.empty()) << "fixture sheet declares no <cols> entry";
  for (std::size_t i = 0; i < cb.size(); ++i) {
    SCOPED_TRACE("column entry " + std::to_string(i));
    EXPECT_EQ(cb[i].first, cx[i].first);
    EXPECT_EQ(cb[i].last, cx[i].last);
    EXPECT_EQ(cb[i].hidden, cx[i].hidden);
    EXPECT_EQ(cb[i].outline_level, cx[i].outline_level);
    EXPECT_EQ(cb[i].has_width, cx[i].has_width);
    // The width survives the 1/256-digit quantisation to within one step of
    // it. `has_style` is deliberately not compared: `BrtColInfo` carries a
    // mandatory `ixfe` with no presence bit, so the XLSB side reports an
    // effective style 0 where OOXML reports none.
    EXPECT_NEAR(cb[i].width, cx[i].width, 1.0 / 256.0);
  }
}
TEST(XlsbCrossFormatSymmetry, DefinedNamesMatch) {
  Workbook xlsb = Workbook::create_empty();
  Workbook xlsx = Workbook::create_empty();
  ASSERT_TRUE(LoadBothFormats(&xlsb, &xlsx));
  const std::vector<DefinedName>& sb = xlsb.defined_names();
  const std::vector<DefinedName>& sx = xlsx.defined_names();
  ASSERT_EQ(sb.size(), sx.size());
  // Field-level, not just count: the XLSB reader must fill the same
  // `DefinedName` field set the OOXML reader does (name, formula,
  // scope, hidden, comment) for the same source workbook, not merely
  // produce the same number of entries.
  for (std::size_t i = 0; i < sb.size(); ++i) {
    EXPECT_EQ(sb[i].name, sx[i].name) << "index " << i;
    EXPECT_EQ(sb[i].formula, sx[i].formula) << "index " << i;
    EXPECT_EQ(sb[i].local_sheet_id, sx[i].local_sheet_id) << "index " << i;
    EXPECT_EQ(sb[i].hidden, sx[i].hidden) << "index " << i;
    EXPECT_EQ(sb[i].comment, sx[i].comment) << "index " << i;
  }
}

}  // namespace
}  // namespace formulon
