//
// C ABI - styles surface (cell xf bindings, fonts, fills, borders, num
// formats, cell styles, dedup-on-insert helpers).

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <type_traits>
#include <unordered_map>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "io/styles_reader.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

static_assert(std::is_nothrow_move_constructible_v<formulon::io::StylesTable>,
              "StylesTable must be nothrow-move-constructible for staged style commits");
static_assert(std::is_nothrow_move_assignable_v<formulon::io::StylesTable>,
              "StylesTable must be nothrow-move-assignable for staged style commits");

/// Validates the `(handle, sheet_index, row, col)` quad and resolves
/// the cell's `xf_index`. On failure populates the thread-local
/// diagnostic and returns the status. On success writes
/// `*out_xf_index` and returns `kOk`. The `xf_index` defaults to `0`
/// (the default xf) when no cell exists at the address.
fm_status_t resolve_xf(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                       std::uint32_t* out_xf_index, const char* fn) {
  if (out_xf_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (auto rc = check_sheet_u32(wb, sheet, fn); rc != 0) {
    return rc;
  }
  const formulon::Cell* cell = wb->workbook().sheet(sheet).cell_at(row, col);
  *out_xf_index = (cell != nullptr) ? cell->xf_index : 0U;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_cell_get_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                            uint32_t* out_xf_index) {
  clear_last_error();
  return resolve_xf(wb, sheet, row, col, out_xf_index, "fm_cell_get_xf_index");
}

extern "C" fm_status_t fm_cell_set_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t col,
                                            uint32_t xf_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_set_xf_index: wb is NULL");
  }
  auto r = wb->workbook().set_cell_xf_index(sheet, row, col, xf_index);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

namespace {

fm_status_t get_cell_xf_impl(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out, const char* api) {
  if (wb == nullptr || out == nullptr) {
    const std::string message = std::string(api) + ": NULL argument";
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, message.c_str());
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (xf_index >= styles.cell_xfs.size()) {
    const std::string message = std::string(api) + ": xf_index out of range";
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, message.c_str(),
        "xf_index=" + std::to_string(xf_index) + " cell_xfs_count=" + std::to_string(styles.cell_xfs.size()));
  }
  const formulon::io::CellXf& xf = styles.cell_xfs[xf_index];
  out->font_index = xf.font_index;
  out->fill_index = xf.fill_index;
  out->border_index = xf.border_index;
  out->num_fmt_id = xf.num_fmt_id;
  out->horizontal_align = xf.horizontal_align;
  out->vertical_align = xf.vertical_align;
  out->wrap_text = xf.wrap_text ? 1 : 0;
  return 0;
}

fm_status_t get_cell_style_xf_impl(fm_workbook_t* wb, uint32_t index, fm_cell_xf* out, const char* api) {
  if (wb == nullptr || out == nullptr) {
    const std::string message = std::string(api) + ": NULL argument";
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, message.c_str());
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_style_xfs.size()) {
    const std::string message = std::string(api) + ": index out of range";
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, message.c_str(),
        "index=" + std::to_string(index) + " cell_style_xfs_count=" + std::to_string(styles.cell_style_xfs.size()));
  }
  const formulon::io::CellXf& xf = styles.cell_style_xfs[index];
  out->font_index = xf.font_index;
  out->fill_index = xf.fill_index;
  out->border_index = xf.border_index;
  out->num_fmt_id = xf.num_fmt_id;
  out->horizontal_align = xf.horizontal_align;
  out->vertical_align = xf.vertical_align;
  out->wrap_text = xf.wrap_text ? 1 : 0;
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_styles_get_cell_xf(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out) {
  clear_last_error();
  return get_cell_xf_impl(wb, xf_index, out, "fm_styles_get_cell_xf");
}

extern "C" fm_status_t fm_styles_get_cell_xf_ex(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf_ex* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_xf_ex: NULL argument");
  }
  fm_status_t rc = get_cell_xf_impl(wb, xf_index, &out->base, "fm_styles_get_cell_xf_ex");
  if (rc != 0) {
    return rc;
  }
  out->justify_last_line = wb->workbook().styles().cell_xfs[xf_index].justify_last_line ? 1 : 0;
  out->xf_id = wb->workbook().styles().cell_xfs[xf_index].xf_id;
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_xf_ex2(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf_ex2* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_xf_ex2: NULL argument");
  }
  fm_status_t rc = get_cell_xf_impl(wb, xf_index, &out->base, "fm_styles_get_cell_xf_ex2");
  if (rc != 0) {
    return rc;
  }
  const formulon::io::CellXf& xf = wb->workbook().styles().cell_xfs[xf_index];
  out->justify_last_line = xf.justify_last_line ? 1 : 0;
  out->xf_id = xf.xf_id;
  out->has_alignment = xf.has_alignment ? 1 : 0;
  out->has_text_rotation = xf.has_text_rotation ? 1 : 0;
  out->text_rotation = xf.text_rotation;
  out->has_indent = xf.has_indent ? 1 : 0;
  out->indent = xf.indent;
  out->has_relative_indent = xf.has_relative_indent ? 1 : 0;
  out->relative_indent = xf.relative_indent;
  out->has_shrink_to_fit = xf.has_shrink_to_fit ? 1 : 0;
  out->shrink_to_fit = xf.shrink_to_fit ? 1 : 0;
  out->has_reading_order = xf.has_reading_order ? 1 : 0;
  out->reading_order = xf.reading_order;
  out->has_horizontal_align = xf.has_horizontal_align ? 1 : 0;
  out->has_vertical_align = xf.has_vertical_align ? 1 : 0;
  out->has_wrap_text = xf.has_wrap_text ? 1 : 0;
  out->has_justify_last_line = xf.has_justify_last_line ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_font(fm_workbook_t* wb, uint32_t font_index, fm_font_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_font: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (font_index >= styles.fonts.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_font: font_index out of range",
        "font_index=" + std::to_string(font_index) + " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  const formulon::io::FontRecord& f = styles.fonts[font_index];
  out->name = f.name.c_str();
  out->size = f.size;
  out->color_argb = f.color_argb;
  out->bold = f.bold ? 1 : 0;
  out->italic = f.italic ? 1 : 0;
  out->strike = f.strike ? 1 : 0;
  out->underline = f.underline;
  return 0;
}

extern "C" fm_status_t fm_styles_get_font_ex(fm_workbook_t* wb, uint32_t font_index, fm_font_record_ex* out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_font_ex: out is NULL");
  }
  const fm_status_t rc = fm_styles_get_font(wb, font_index, &out->base);
  if (rc != 0) {
    return rc;
  }
  out->vert_align = wb->workbook().styles().fonts[font_index].vert_align;
  return 0;
}

extern "C" fm_status_t fm_styles_get_num_fmt_string(fm_workbook_t* wb, uint16_t num_fmt_id, const char** out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_num_fmt_string: NULL argument");
  }
  // A file may define a custom `<numFmt>` whose `numFmtId` overrides a
  // built-in slot; Excel honours the file's definition over the built-in.
  // Search the custom table first so an override wins, then fall back to
  // the built-in `.rodata` table only when no override is present.
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.id == num_fmt_id && n.format_string_index < styles.num_fmt_strings.size()) {
      *out = styles.num_fmt_strings[n.format_string_index].c_str();
      return 0;
    }
  }
  // Built-in ids (0..163) resolve through the writer's `.rodata` table.
  if (num_fmt_id < 164U) {
    const char* s = formulon::io::builtin_num_fmt(num_fmt_id);
    if (s != nullptr && s[0] != '\0') {
      *out = s;
      return 0;
    }
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_num_fmt_string: id not found",
                           "num_fmt_id=" + std::to_string(num_fmt_id));
}

extern "C" fm_status_t fm_styles_get_fill(fm_workbook_t* wb, uint32_t fill_index, fm_fill_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_fill: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (fill_index >= styles.fills.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_fill: fill_index out of range",
        "fill_index=" + std::to_string(fill_index) + " fills_count=" + std::to_string(styles.fills.size()));
  }
  const formulon::io::FillRecord& f = styles.fills[fill_index];
  out->pattern = f.pattern;
  out->fg_argb = f.fg_argb;
  out->bg_argb = f.bg_argb;
  return 0;
}

extern "C" fm_status_t fm_styles_get_border(fm_workbook_t* wb, uint32_t border_index, fm_border_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_border: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (border_index >= styles.borders.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_border: border_index out of range",
        "border_index=" + std::to_string(border_index) + " borders_count=" + std::to_string(styles.borders.size()));
  }
  const formulon::io::BorderRecord& b = styles.borders[border_index];
  auto fill_side = [](const formulon::io::BorderSide& src, fm_border_side& dst) noexcept {
    dst.style = src.style;
    dst.color_argb = src.color_argb;
  };
  fill_side(b.left, out->left);
  fill_side(b.right, out->right);
  fill_side(b.top, out->top);
  fill_side(b.bottom, out->bottom);
  fill_side(b.diagonal, out->diagonal);
  out->diagonal_up = b.diagonal_up ? 1 : 0;
  out->diagonal_down = b.diagonal_down ? 1 : 0;
  return 0;
}

namespace {

void font_to_c(const formulon::io::FontRecord& f, fm_font_record* out) {
  out->name = f.name.c_str();
  out->size = f.size;
  out->color_argb = f.color_argb;
  out->bold = f.bold ? 1 : 0;
  out->italic = f.italic ? 1 : 0;
  out->strike = f.strike ? 1 : 0;
  out->underline = f.underline;
}

void fill_to_c(const formulon::io::FillRecord& f, fm_fill_record* out) {
  out->pattern = f.pattern;
  out->fg_argb = f.fg_argb;
  out->bg_argb = f.bg_argb;
}

void border_to_c(const formulon::io::BorderRecord& b, fm_border_record* out) {
  auto fill_side = [](const formulon::io::BorderSide& src, fm_border_side& dst) noexcept {
    dst.style = src.style;
    dst.color_argb = src.color_argb;
  };
  fill_side(b.left, out->left);
  fill_side(b.right, out->right);
  fill_side(b.top, out->top);
  fill_side(b.bottom, out->bottom);
  fill_side(b.diagonal, out->diagonal);
  out->diagonal_up = b.diagonal_up ? 1 : 0;
  out->diagonal_down = b.diagonal_down ? 1 : 0;
}

}  // namespace

extern "C" fm_status_t fm_styles_get_dxf_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_dxf_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().dxfs.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_dxf(fm_workbook_t* wb, uint32_t dxf_index, fm_dxf_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_dxf: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (dxf_index >= styles.dxfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_dxf: dxf_index out of range",
        "dxf_index=" + std::to_string(dxf_index) + " dxfs_count=" + std::to_string(styles.dxfs.size()));
  }
  *out = fm_dxf_record{};
  const formulon::io::DifferentialFormat& dxf = styles.dxfs[dxf_index];
  out->font_engaged = dxf.has_font ? 1 : 0;
  if (dxf.has_font) {
    font_to_c(dxf.font, &out->font);
  }
  out->fill_engaged = dxf.has_fill ? 1 : 0;
  if (dxf.has_fill) {
    fill_to_c(dxf.fill, &out->fill);
  }
  out->border_engaged = dxf.has_border ? 1 : 0;
  if (dxf.has_border) {
    border_to_c(dxf.border, &out->border);
  }
  out->num_fmt_engaged = dxf.has_num_fmt ? 1 : 0;
  out->num_fmt_id = dxf.num_fmt_id;
  out->num_fmt_code = dxf.num_fmt_code.c_str();
  return 0;
}

// `fm_styles_get_{font,fill,border,cell_xf}_count` are now emitted by
// the binding codegen (see `src/c_api/generated/styles_counts.cpp`).
// `fm_styles_get_cell_style_count` /
// `fm_styles_get_cell_style_xf_count` stay hand-written because the JS
// surface only exposes them on the embind binding; they are not part of
// the cross-binding manifest.

extern "C" fm_status_t fm_styles_get_cell_style_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().cell_styles.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style(fm_workbook_t* wb, uint32_t index, fm_cell_style_record_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_styles.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style: index out of range",
        "index=" + std::to_string(index) + " cell_styles_count=" + std::to_string(styles.cell_styles.size()));
  }
  const formulon::io::CellStyleRecord& cs = styles.cell_styles[index];
  out->name = cs.name.c_str();
  out->xf_id = cs.xf_id;
  out->builtin_id = cs.builtin_id;
  out->i_level = cs.i_level;
  out->hidden = cs.hidden ? 1 : 0;
  out->custom_builtin = cs.custom_builtin ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_xf_count(fm_workbook_t* wb, uint32_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf_count: NULL argument");
  }
  *out_count = static_cast<uint32_t>(wb->workbook().styles().cell_style_xfs.size());
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_xf(fm_workbook_t* wb, uint32_t index, fm_cell_xf* out) {
  clear_last_error();
  return get_cell_style_xf_impl(wb, index, out, "fm_styles_get_cell_style_xf");
}

extern "C" fm_status_t fm_styles_get_cell_style_xf_ex(fm_workbook_t* wb, uint32_t index, fm_cell_xf_ex* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf_ex: NULL argument");
  }
  fm_status_t rc = get_cell_style_xf_impl(wb, index, &out->base, "fm_styles_get_cell_style_xf_ex");
  if (rc != 0) {
    return rc;
  }
  out->justify_last_line = wb->workbook().styles().cell_style_xfs[index].justify_last_line ? 1 : 0;
  out->xf_id = 0;
  return 0;
}

extern "C" fm_status_t fm_styles_get_cell_style_xf_ex2(fm_workbook_t* wb, uint32_t index, fm_cell_xf_ex2* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf_ex2: NULL argument");
  }
  fm_status_t rc = get_cell_style_xf_impl(wb, index, &out->base, "fm_styles_get_cell_style_xf_ex2");
  if (rc != 0) {
    return rc;
  }
  const formulon::io::CellXf& xf = wb->workbook().styles().cell_style_xfs[index];
  out->justify_last_line = xf.justify_last_line ? 1 : 0;
  out->xf_id = 0;
  out->has_alignment = xf.has_alignment ? 1 : 0;
  out->has_text_rotation = xf.has_text_rotation ? 1 : 0;
  out->text_rotation = xf.text_rotation;
  out->has_indent = xf.has_indent ? 1 : 0;
  out->indent = xf.indent;
  out->has_relative_indent = xf.has_relative_indent ? 1 : 0;
  out->relative_indent = xf.relative_indent;
  out->has_shrink_to_fit = xf.has_shrink_to_fit ? 1 : 0;
  out->shrink_to_fit = xf.shrink_to_fit ? 1 : 0;
  out->has_reading_order = xf.has_reading_order ? 1 : 0;
  out->reading_order = xf.reading_order;
  out->has_horizontal_align = xf.has_horizontal_align ? 1 : 0;
  out->has_vertical_align = xf.has_vertical_align ? 1 : 0;
  out->has_wrap_text = xf.has_wrap_text ? 1 : 0;
  out->has_justify_last_line = xf.has_justify_last_line ? 1 : 0;
  return 0;
}

namespace {

/// Field-for-field equality for the C++ side `FontRecord`. Excluded
/// from `<algorithm>` because the struct has no `operator==`; defining
/// one here keeps the dedup-comparator confined to the C ABI's needs.
bool font_records_equal(const formulon::io::FontRecord& a, const formulon::io::FontRecord& b) {
  return a.name == b.name && a.size == b.size && a.bold == b.bold && a.italic == b.italic && a.strike == b.strike &&
         a.underline == b.underline && a.vert_align == b.vert_align && a.color_argb == b.color_argb;
}

bool fill_records_equal(const formulon::io::FillRecord& a, const formulon::io::FillRecord& b) noexcept {
  return a.pattern == b.pattern && a.fg_argb == b.fg_argb && a.bg_argb == b.bg_argb;
}

bool border_sides_equal(const formulon::io::BorderSide& a, const formulon::io::BorderSide& b) noexcept {
  return a.style == b.style && a.color_argb == b.color_argb;
}

bool border_records_equal(const formulon::io::BorderRecord& a, const formulon::io::BorderRecord& b) noexcept {
  return border_sides_equal(a.left, b.left) && border_sides_equal(a.right, b.right) &&
         border_sides_equal(a.top, b.top) && border_sides_equal(a.bottom, b.bottom) &&
         border_sides_equal(a.diagonal, b.diagonal) && a.diagonal_up == b.diagonal_up &&
         a.diagonal_down == b.diagonal_down;
}

template <typename T>
bool optional_xf_field_equal(bool a_has, const T& a_value, bool b_has, const T& b_value) noexcept {
  return a_has == b_has && (!a_has || a_value == b_value);
}

bool cell_xfs_equal(const formulon::io::CellXf& a, const formulon::io::CellXf& b) noexcept {
  return a.font_index == b.font_index && a.fill_index == b.fill_index && a.border_index == b.border_index &&
         a.num_fmt_id == b.num_fmt_id && a.xf_id == b.xf_id && a.apply_number_format == b.apply_number_format &&
         a.apply_font == b.apply_font && a.apply_fill == b.apply_fill && a.apply_border == b.apply_border &&
         a.apply_alignment == b.apply_alignment && a.apply_protection == b.apply_protection &&
         a.quote_prefix == b.quote_prefix && formulon::io::HasAlignment(a) == formulon::io::HasAlignment(b) &&
         optional_xf_field_equal(formulon::io::HasHorizontalAlign(a), a.horizontal_align,
                                 formulon::io::HasHorizontalAlign(b), b.horizontal_align) &&
         optional_xf_field_equal(formulon::io::HasVerticalAlign(a), a.vertical_align, formulon::io::HasVerticalAlign(b),
                                 b.vertical_align) &&
         optional_xf_field_equal(formulon::io::HasWrapText(a), a.wrap_text, formulon::io::HasWrapText(b),
                                 b.wrap_text) &&
         optional_xf_field_equal(formulon::io::HasJustifyLastLine(a), a.justify_last_line,
                                 formulon::io::HasJustifyLastLine(b), b.justify_last_line) &&
         optional_xf_field_equal(a.has_text_rotation, a.text_rotation, b.has_text_rotation, b.text_rotation) &&
         optional_xf_field_equal(a.has_indent, a.indent, b.has_indent, b.indent) &&
         optional_xf_field_equal(a.has_relative_indent, a.relative_indent, b.has_relative_indent, b.relative_indent) &&
         optional_xf_field_equal(a.has_shrink_to_fit, a.shrink_to_fit, b.has_shrink_to_fit, b.shrink_to_fit) &&
         optional_xf_field_equal(a.has_reading_order, a.reading_order, b.has_reading_order, b.reading_order) &&
         a.has_protection == b.has_protection && (!a.has_protection || (a.locked == b.locked && a.hidden == b.hidden));
}

template <typename T>
void append_key_scalar(std::string& out, const T& value) {
  out.append(reinterpret_cast<const char*>(&value), sizeof(value));
}

std::string font_key(const formulon::io::FontRecord& value) {
  std::string out = value.name;
  out.push_back('\0');
  append_key_scalar(out, value.size);
  append_key_scalar(out, value.bold);
  append_key_scalar(out, value.italic);
  append_key_scalar(out, value.strike);
  append_key_scalar(out, value.underline);
  append_key_scalar(out, value.vert_align);
  append_key_scalar(out, value.color_argb);
  return out;
}

std::string fill_key(const formulon::io::FillRecord& value) {
  std::string out;
  append_key_scalar(out, value.pattern);
  append_key_scalar(out, value.fg_argb);
  append_key_scalar(out, value.bg_argb);
  return out;
}

std::string border_key(const formulon::io::BorderRecord& value) {
  std::string out;
  for (const formulon::io::BorderSide* side : {&value.left, &value.right, &value.top, &value.bottom, &value.diagonal}) {
    append_key_scalar(out, side->style);
    append_key_scalar(out, side->color_argb);
  }
  append_key_scalar(out, value.diagonal_up);
  append_key_scalar(out, value.diagonal_down);
  return out;
}

std::string cell_xf_key(const formulon::io::CellXf& value) {
  std::string out;
  append_key_scalar(out, value.font_index);
  append_key_scalar(out, value.fill_index);
  append_key_scalar(out, value.border_index);
  append_key_scalar(out, value.num_fmt_id);
  append_key_scalar(out, value.xf_id);
  append_key_scalar(out, value.apply_number_format);
  append_key_scalar(out, value.apply_font);
  append_key_scalar(out, value.apply_fill);
  append_key_scalar(out, value.apply_border);
  append_key_scalar(out, value.apply_alignment);
  append_key_scalar(out, value.apply_protection);
  append_key_scalar(out, value.quote_prefix);
  const bool has_alignment = formulon::io::HasAlignment(value);
  append_key_scalar(out, has_alignment);
  const bool has_horizontal_align = formulon::io::HasHorizontalAlign(value);
  append_key_scalar(out, has_horizontal_align);
  if (has_horizontal_align) {
    append_key_scalar(out, value.horizontal_align);
  }
  const bool has_vertical_align = formulon::io::HasVerticalAlign(value);
  append_key_scalar(out, has_vertical_align);
  if (has_vertical_align) {
    append_key_scalar(out, value.vertical_align);
  }
  const bool has_wrap_text = formulon::io::HasWrapText(value);
  append_key_scalar(out, has_wrap_text);
  if (has_wrap_text) {
    append_key_scalar(out, value.wrap_text);
  }
  const bool has_justify_last_line = formulon::io::HasJustifyLastLine(value);
  append_key_scalar(out, has_justify_last_line);
  if (has_justify_last_line) {
    append_key_scalar(out, value.justify_last_line);
  }
  append_key_scalar(out, value.has_text_rotation);
  if (value.has_text_rotation) {
    append_key_scalar(out, value.text_rotation);
  }
  append_key_scalar(out, value.has_indent);
  if (value.has_indent) {
    append_key_scalar(out, value.indent);
  }
  append_key_scalar(out, value.has_relative_indent);
  if (value.has_relative_indent) {
    append_key_scalar(out, value.relative_indent);
  }
  append_key_scalar(out, value.has_shrink_to_fit);
  if (value.has_shrink_to_fit) {
    append_key_scalar(out, value.shrink_to_fit);
  }
  append_key_scalar(out, value.has_reading_order);
  if (value.has_reading_order) {
    append_key_scalar(out, value.reading_order);
  }
  append_key_scalar(out, value.has_protection);
  if (value.has_protection) {
    append_key_scalar(out, value.locked);
    append_key_scalar(out, value.hidden);
  }
  return out;
}

bool dxf_records_equal(const formulon::io::DifferentialFormat& a, const formulon::io::DifferentialFormat& b) noexcept {
  return a.has_font == b.has_font && (!a.has_font || font_records_equal(a.font, b.font)) && a.has_fill == b.has_fill &&
         (!a.has_fill || fill_records_equal(a.fill, b.fill)) && a.has_border == b.has_border &&
         (!a.has_border || border_records_equal(a.border, b.border)) && a.has_num_fmt == b.has_num_fmt &&
         (!a.has_num_fmt || (a.num_fmt_id == b.num_fmt_id && a.num_fmt_code == b.num_fmt_code));
}

void ensure_default_style_roots(formulon::io::StylesTable& styles) {
  if (styles.fonts.empty()) {
    styles.fonts.emplace_back();
  }
  if (styles.fills.empty()) {
    styles.fills.emplace_back();
  }
  if (styles.borders.empty()) {
    styles.borders.emplace_back();
  }
}

void ensure_default_cell_xf(formulon::io::StylesTable& styles) {
  ensure_default_style_roots(styles);
  if (styles.cell_xfs.empty()) {
    styles.cell_xfs.emplace_back();
  }
}

/// Rejects an `<xf>` whose font / fill / border / number-format reference
/// points past the corresponding table. Auto-growing the parallel tables
/// would hide the caller's ordering mistake and emit a dangling index that
/// Excel reports as a repairable file.
fm_status_t validate_xf_references(const formulon::io::StylesTable& styles, const fm_cell_xf& record, const char* api) {
  auto reject = [api](const char* what, std::string context) {
    const std::string message = std::string(api) + ": " + what;
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, message.c_str(), std::move(context));
  };
  if (record.font_index >= styles.fonts.size()) {
    return reject("font_index out of range", "font_index=" + std::to_string(record.font_index) +
                                                 " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  if (record.fill_index >= styles.fills.size()) {
    return reject("fill_index out of range", "fill_index=" + std::to_string(record.fill_index) +
                                                 " fills_count=" + std::to_string(styles.fills.size()));
  }
  if (record.border_index >= styles.borders.size()) {
    return reject("border_index out of range", "border_index=" + std::to_string(record.border_index) +
                                                   " borders_count=" + std::to_string(styles.borders.size()));
  }
  // num_fmt_id must be a documented built-in or a registered custom id.
  bool num_fmt_ok = false;
  if (record.num_fmt_id < 164U) {
    const char* s = formulon::io::builtin_num_fmt(record.num_fmt_id);
    num_fmt_ok = (s != nullptr && s[0] != '\0');
  }
  if (!num_fmt_ok) {
    for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
      if (n.id == record.num_fmt_id) {
        num_fmt_ok = true;
        break;
      }
    }
  }
  if (!num_fmt_ok) {
    return reject("num_fmt_id not registered", "num_fmt_id=" + std::to_string(record.num_fmt_id));
  }
  return 0;
}

fm_status_t validate_xf_alignment_extras(const fm_cell_xf_ex2& record, const char* api) {
  auto reject = [api](const char* what, std::string context) {
    const std::string message = std::string(api) + ": " + what;
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, message.c_str(), std::move(context));
  };
  if (record.has_text_rotation != 0 && record.text_rotation > 180U && record.text_rotation != 255U) {
    return reject("text_rotation out of range", "text_rotation=" + std::to_string(record.text_rotation));
  }
  if (record.has_indent != 0 && record.indent > 255U) {
    return reject("indent out of range", "indent=" + std::to_string(record.indent));
  }
  if (record.has_reading_order != 0 && record.reading_order > 2U) {
    return reject("reading_order out of range", "reading_order=" + std::to_string(record.reading_order));
  }
  return 0;
}

fm_status_t validate_xf_alignment_enums(const fm_cell_xf& record, const char* api) {
  auto reject = [api](const char* what, std::string context) {
    const std::string message = std::string(api) + ": " + what;
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, message.c_str(), std::move(context));
  };
  // These ordinals are the canonical writer mapping for the corresponding
  // XML string enums (general..distributed and top..distributed).
  if (record.horizontal_align > 7U) {
    return reject("horizontal_align out of range", "horizontal_align=" + std::to_string(record.horizontal_align));
  }
  if (record.vertical_align > 4U) {
    return reject("vertical_align out of range", "vertical_align=" + std::to_string(record.vertical_align));
  }
  return 0;
}

formulon::io::FontRecord font_from_c(const fm_font_record& record) {
  formulon::io::FontRecord out;
  out.name = (record.name != nullptr) ? std::string(record.name) : std::string();
  out.size = record.size;
  out.bold = record.bold != 0;
  out.italic = record.italic != 0;
  out.strike = record.strike != 0;
  out.underline = record.underline;
  out.color_argb = record.color_argb;
  return out;
}

formulon::io::FontRecord font_from_c_ex(const fm_font_record_ex& record) {
  formulon::io::FontRecord out = font_from_c(record.base);
  out.vert_align = record.vert_align;
  return out;
}

formulon::io::FillRecord fill_from_c(const fm_fill_record& record) noexcept {
  formulon::io::FillRecord out;
  out.pattern = record.pattern;
  out.fg_argb = record.fg_argb;
  out.bg_argb = record.bg_argb;
  return out;
}

formulon::io::BorderSide border_side_from_c(const fm_border_side& src) noexcept {
  formulon::io::BorderSide dst;
  dst.style = src.style;
  dst.color_argb = src.color_argb;
  return dst;
}

formulon::io::BorderRecord border_from_c(const fm_border_record& record) noexcept {
  formulon::io::BorderRecord out;
  out.left = border_side_from_c(record.left);
  out.right = border_side_from_c(record.right);
  out.top = border_side_from_c(record.top);
  out.bottom = border_side_from_c(record.bottom);
  out.diagonal = border_side_from_c(record.diagonal);
  out.diagonal_up = record.diagonal_up != 0;
  out.diagonal_down = record.diagonal_down != 0;
  return out;
}

}  // namespace

extern "C" fm_status_t fm_styles_add_font(fm_workbook_t* wb, fm_font_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_font: NULL argument");
  }
  formulon::io::FontRecord candidate = font_from_c(record);

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.fonts.size(); ++i) {
    if (font_records_equal(styles.fonts[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.fonts.push_back(std::move(candidate));
  *out_index = static_cast<uint32_t>(styles.fonts.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_font_ex(fm_workbook_t* wb, fm_font_record_ex record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_font_ex: NULL argument");
  }
  formulon::io::FontRecord candidate = font_from_c_ex(record);
  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.fonts.size(); ++i) {
    if (font_records_equal(styles.fonts[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.fonts.push_back(std::move(candidate));
  *out_index = static_cast<uint32_t>(styles.fonts.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_fill(fm_workbook_t* wb, fm_fill_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_fill: NULL argument");
  }
  formulon::io::FillRecord candidate = fill_from_c(record);

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.fills.size(); ++i) {
    if (fill_records_equal(styles.fills[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.fills.push_back(candidate);
  *out_index = static_cast<uint32_t>(styles.fills.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_border(fm_workbook_t* wb, fm_border_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_border: NULL argument");
  }
  formulon::io::BorderRecord candidate = border_from_c(record);

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.borders.size(); ++i) {
    if (border_records_equal(styles.borders[i], candidate)) {
      *out_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.borders.push_back(candidate);
  *out_index = static_cast<uint32_t>(styles.borders.size() - 1);
  return 0;
}

extern "C" fm_status_t fm_styles_add_num_fmt(fm_workbook_t* wb, const char* format_code, uint16_t* out_num_fmt_id) {
  clear_last_error();
  if (wb == nullptr || out_num_fmt_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_num_fmt: NULL argument");
  }
  const std::string code = (format_code != nullptr) ? std::string(format_code) : std::string();

  // Step 1: built-in match.
  for (uint16_t id = 0; id < 164U; ++id) {
    const char* s = formulon::io::builtin_num_fmt(id);
    if (s != nullptr && s[0] != '\0' && code == s) {
      *out_num_fmt_id = id;
      return 0;
    }
  }

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();

  // Step 2: existing custom entry.
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.format_string_index < styles.num_fmt_strings.size() &&
        styles.num_fmt_strings[n.format_string_index] == code) {
      *out_num_fmt_id = n.id;
      return 0;
    }
  }

  // Step 3: append a new custom entry. New id is one past the largest
  // existing custom id, with `163` as the lower bound (the last
  // built-in slot).
  std::uint32_t next_id = 163U;
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.id > next_id) {
      next_id = n.id;
    }
  }
  if (next_id >= std::numeric_limits<std::uint16_t>::max()) {
    return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed,
                             "fm_styles_add_num_fmt: custom numFmtId space exhausted",
                             "max_num_fmt_id=" + std::to_string(std::numeric_limits<std::uint16_t>::max()));
  }
  ++next_id;

  styles.num_fmt_strings.push_back(code);
  formulon::io::NumFmtRecord rec;
  rec.id = static_cast<std::uint16_t>(next_id);
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1);
  styles.num_fmts.push_back(rec);
  *out_num_fmt_id = static_cast<std::uint16_t>(next_id);
  return 0;
}

namespace {

fm_status_t add_cell_xf_impl(fm_workbook_t* wb, fm_cell_xf record, bool justify_last_line, uint32_t xf_id,
                             const fm_cell_xf_ex2* extras, uint32_t* out_xf_index, const char* api) {
  if (wb == nullptr || out_xf_index == nullptr) {
    const std::string message = std::string(api) + ": NULL argument";
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, message.c_str());
  }
  if (fm_status_t rc = validate_xf_alignment_enums(record, api); rc != 0) {
    return rc;
  }
  if (extras != nullptr) {
    if (fm_status_t rc = validate_xf_alignment_extras(*extras, api); rc != 0) {
      return rc;
    }
  }
  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_cell_xf(styles);

  // Reject out-of-range references rather than auto-growing the parallel
  // tables; callers must register fonts / fills / borders beforehand.
  if (fm_status_t rc = validate_xf_references(styles, record, api); rc != 0) {
    return rc;
  }
  // A non-zero `xfId` must name an existing `<cellStyleXfs>` entry, or the
  // emitted attribute dangles. Zero stays legal with an empty table so the
  // plain `fm_styles_add_cell_xf` path keeps its historic behaviour.
  if (xf_id != 0 && xf_id >= styles.cell_style_xfs.size()) {
    const std::string message = std::string(api) + ": xf_id out of range";
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, message.c_str(),
        "xf_id=" + std::to_string(xf_id) + " cell_style_xfs_count=" + std::to_string(styles.cell_style_xfs.size()));
  }

  // Confirm the read-side `fm_cell_xf` struct mirrors the engine
  // `formulon::io::CellXf` field-for-field; the write side reuses the
  // same shape so a layout drift would silently corrupt records.
  static_assert(sizeof(record.font_index) == sizeof(formulon::io::CellXf::font_index),
                "fm_cell_xf::font_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.fill_index) == sizeof(formulon::io::CellXf::fill_index),
                "fm_cell_xf::fill_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.border_index) == sizeof(formulon::io::CellXf::border_index),
                "fm_cell_xf::border_index width must match formulon::io::CellXf");
  static_assert(sizeof(record.num_fmt_id) == sizeof(formulon::io::CellXf::num_fmt_id),
                "fm_cell_xf::num_fmt_id width must match formulon::io::CellXf");

  formulon::io::CellXf candidate;
  candidate.font_index = record.font_index;
  candidate.fill_index = record.fill_index;
  candidate.border_index = record.border_index;
  candidate.num_fmt_id = record.num_fmt_id;
  candidate.horizontal_align = record.horizontal_align;
  candidate.vertical_align = record.vertical_align;
  candidate.wrap_text = record.wrap_text != 0;
  candidate.justify_last_line = justify_last_line;
  candidate.xf_id = xf_id;
  candidate.has_horizontal_align = candidate.horizontal_align != 0U;
  candidate.has_vertical_align = candidate.vertical_align != 2U;
  candidate.has_wrap_text = candidate.wrap_text;
  candidate.has_justify_last_line = candidate.justify_last_line;
  if (extras != nullptr) {
    candidate.has_alignment = extras->has_alignment != 0;
    candidate.has_horizontal_align = extras->has_horizontal_align != 0 || candidate.has_horizontal_align;
    candidate.has_vertical_align = extras->has_vertical_align != 0 || candidate.has_vertical_align;
    candidate.has_wrap_text = extras->has_wrap_text != 0 || candidate.has_wrap_text;
    candidate.has_justify_last_line = extras->has_justify_last_line != 0 || candidate.has_justify_last_line;
    if (!candidate.has_alignment && !candidate.has_horizontal_align && !candidate.has_vertical_align &&
        !candidate.has_wrap_text && !candidate.has_justify_last_line && record.horizontal_align == 0U &&
        record.vertical_align == 0U && record.wrap_text == 0 && !candidate.justify_last_line &&
        extras->has_text_rotation == 0 && extras->has_indent == 0 && extras->has_relative_indent == 0 &&
        extras->has_shrink_to_fit == 0 && extras->has_reading_order == 0) {
      // A zero-initialized public record denotes the omitted child. The
      // engine's normalized effective default for vertical alignment is
      // bottom, not the ordinal used for an explicit top alignment.
      candidate.vertical_align = 2U;
    }
    candidate.has_text_rotation = extras->has_text_rotation != 0;
    candidate.text_rotation = extras->text_rotation;
    candidate.has_indent = extras->has_indent != 0;
    candidate.indent = extras->indent;
    candidate.has_relative_indent = extras->has_relative_indent != 0;
    candidate.relative_indent = extras->relative_indent;
    candidate.has_shrink_to_fit = extras->has_shrink_to_fit != 0;
    candidate.shrink_to_fit = extras->shrink_to_fit != 0;
    candidate.has_reading_order = extras->has_reading_order != 0;
    candidate.reading_order = extras->reading_order;
    // Keep the C ABI's inferred presence semantics in lockstep with the
    // binding adapters and the writer.  A caller may supply an alignment
    // value without setting the top-level presence bit; non-default values
    // still require an alignment child.  Conversely, an explicit empty
    // child remains present through `has_alignment=true`.
    candidate.has_alignment = formulon::io::HasAlignment(candidate);
  }
  candidate.has_alignment = formulon::io::HasAlignment(candidate);

  for (std::size_t i = 0; i < styles.cell_xfs.size(); ++i) {
    if (cell_xfs_equal(styles.cell_xfs[i], candidate)) {
      *out_xf_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.cell_xfs.push_back(candidate);
  *out_xf_index = static_cast<uint32_t>(styles.cell_xfs.size() - 1);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_styles_add_cell_xf(fm_workbook_t* wb, fm_cell_xf record, uint32_t* out_xf_index) {
  clear_last_error();
  return add_cell_xf_impl(wb, record, false, 0, nullptr, out_xf_index, "fm_styles_add_cell_xf");
}

extern "C" fm_status_t fm_styles_add_cell_xf_ex(fm_workbook_t* wb, fm_cell_xf_ex record, uint32_t* out_xf_index) {
  clear_last_error();
  return add_cell_xf_impl(wb, record.base, record.justify_last_line != 0, record.xf_id, nullptr, out_xf_index,
                          "fm_styles_add_cell_xf_ex");
}

extern "C" fm_status_t fm_styles_add_cell_xf_ex2(fm_workbook_t* wb, fm_cell_xf_ex2 record, uint32_t* out_xf_index) {
  clear_last_error();
  return add_cell_xf_impl(wb, record.base, record.justify_last_line != 0, record.xf_id, &record, out_xf_index,
                          "fm_styles_add_cell_xf_ex2");
}

extern "C" fm_status_t fm_styles_add_cell_style_xf_ex(fm_workbook_t* wb, fm_cell_xf_ex record, uint32_t* out_xf_id) {
  clear_last_error();
  if (wb == nullptr || out_xf_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_add_cell_style_xf_ex: NULL argument");
  }
  if (fm_status_t rc = validate_xf_alignment_enums(record.base, "fm_styles_add_cell_style_xf_ex"); rc != 0) {
    return rc;
  }
  auto& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  if (fm_status_t rc = validate_xf_references(styles, record.base, "fm_styles_add_cell_style_xf_ex"); rc != 0) {
    return rc;
  }
  formulon::io::CellXf candidate;
  candidate.font_index = record.base.font_index;
  candidate.fill_index = record.base.fill_index;
  candidate.border_index = record.base.border_index;
  candidate.num_fmt_id = record.base.num_fmt_id;
  candidate.horizontal_align = record.base.horizontal_align;
  candidate.vertical_align = record.base.vertical_align;
  candidate.wrap_text = record.base.wrap_text != 0;
  candidate.justify_last_line = record.justify_last_line != 0;
  candidate.has_horizontal_align = candidate.horizontal_align != 0U;
  candidate.has_vertical_align = candidate.vertical_align != 2U;
  candidate.has_wrap_text = candidate.wrap_text;
  candidate.has_justify_last_line = candidate.justify_last_line;
  candidate.has_alignment = formulon::io::HasAlignment(candidate);
  for (std::size_t i = 0; i < styles.cell_style_xfs.size(); ++i) {
    if (cell_xfs_equal(styles.cell_style_xfs[i], candidate)) {
      *out_xf_id = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.cell_style_xfs.push_back(std::move(candidate));
  *out_xf_id = static_cast<uint32_t>(styles.cell_style_xfs.size() - 1U);
  return 0;
}

extern "C" fm_status_t fm_styles_add_cell_style_xf_ex2(fm_workbook_t* wb, fm_cell_xf_ex2 record, uint32_t* out_xf_id) {
  clear_last_error();
  if (wb == nullptr || out_xf_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_add_cell_style_xf_ex2: NULL argument");
  }
  if (fm_status_t rc = validate_xf_alignment_enums(record.base, "fm_styles_add_cell_style_xf_ex2"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = validate_xf_alignment_extras(record, "fm_styles_add_cell_style_xf_ex2"); rc != 0) {
    return rc;
  }
  auto& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  if (fm_status_t rc = validate_xf_references(styles, record.base, "fm_styles_add_cell_style_xf_ex2"); rc != 0) {
    return rc;
  }
  formulon::io::CellXf candidate;
  candidate.font_index = record.base.font_index;
  candidate.fill_index = record.base.fill_index;
  candidate.border_index = record.base.border_index;
  candidate.num_fmt_id = record.base.num_fmt_id;
  candidate.horizontal_align = record.base.horizontal_align;
  candidate.vertical_align = record.base.vertical_align;
  candidate.wrap_text = record.base.wrap_text != 0;
  candidate.justify_last_line = record.justify_last_line != 0;
  candidate.has_alignment = record.has_alignment != 0;
  candidate.has_horizontal_align = record.has_horizontal_align != 0 || candidate.horizontal_align != 0U;
  candidate.has_vertical_align = record.has_vertical_align != 0 || candidate.vertical_align != 2U;
  candidate.has_wrap_text = record.has_wrap_text != 0 || candidate.wrap_text;
  candidate.has_justify_last_line = record.has_justify_last_line != 0 || candidate.justify_last_line;
  candidate.has_text_rotation = record.has_text_rotation != 0;
  candidate.text_rotation = record.text_rotation;
  candidate.has_indent = record.has_indent != 0;
  candidate.indent = record.indent;
  candidate.has_relative_indent = record.has_relative_indent != 0;
  candidate.relative_indent = record.relative_indent;
  candidate.has_shrink_to_fit = record.has_shrink_to_fit != 0;
  candidate.shrink_to_fit = record.shrink_to_fit != 0;
  candidate.has_reading_order = record.has_reading_order != 0;
  candidate.reading_order = record.reading_order;
  if (!candidate.has_alignment && !candidate.has_horizontal_align && !candidate.has_vertical_align &&
      !candidate.has_wrap_text && !candidate.has_justify_last_line && record.base.horizontal_align == 0U &&
      record.base.vertical_align == 0U && record.base.wrap_text == 0 && !candidate.justify_last_line &&
      record.has_text_rotation == 0 && record.has_indent == 0 && record.has_relative_indent == 0 &&
      record.has_shrink_to_fit == 0 && record.has_reading_order == 0) {
    candidate.vertical_align = 2U;
  }
  candidate.has_alignment = formulon::io::HasAlignment(candidate);
  for (std::size_t i = 0; i < styles.cell_style_xfs.size(); ++i) {
    if (cell_xfs_equal(styles.cell_style_xfs[i], candidate)) {
      *out_xf_id = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.cell_style_xfs.push_back(std::move(candidate));
  *out_xf_id = static_cast<uint32_t>(styles.cell_style_xfs.size() - 1U);
  return 0;
}

extern "C" fm_status_t fm_styles_set_cell_style(fm_workbook_t* wb, const char* name, uint32_t xf_id,
                                                uint32_t builtin_id) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_set_cell_style: NULL argument");
  }
  auto& styles = wb->workbook().mutable_styles();
  if (name[0] == '\0' || xf_id >= styles.cell_style_xfs.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_set_cell_style: invalid name or xf_id");
  }
  // The OOXML ordinal space is 0..47; anything else must use the sentinel,
  // which suppresses the attribute instead of writing an unknown ordinal.
  if (builtin_id > 47U && builtin_id != FM_CELL_STYLE_BUILTIN_ID_NONE) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_set_cell_style: builtin_id out of range",
                             "builtin_id=" + std::to_string(builtin_id));
  }
  for (auto& style : styles.cell_styles) {
    if (style.name == name) {
      style.xf_id = xf_id;
      style.builtin_id = builtin_id;
      return 0;
    }
  }
  formulon::io::CellStyleRecord style;
  style.name = name;
  style.xf_id = xf_id;
  style.builtin_id = builtin_id;
  styles.cell_styles.push_back(std::move(style));
  return 0;
}

extern "C" fm_status_t fm_styles_add_batch(fm_workbook_t* wb, const fm_styles_batch* batch) {
  clear_last_error();
  if (wb == nullptr || batch == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_batch: NULL argument");
  }
  if ((batch->font_count != 0U && (batch->fonts == nullptr || batch->font_indices == nullptr)) ||
      (batch->fill_count != 0U && (batch->fills == nullptr || batch->fill_indices == nullptr)) ||
      (batch->border_count != 0U && (batch->borders == nullptr || batch->border_indices == nullptr)) ||
      (batch->cell_xf_count != 0U && (batch->cell_xfs == nullptr || batch->cell_xf_indices == nullptr)) ||
      (batch->num_fmt_count != 0U && (batch->num_fmt_codes == nullptr || batch->num_fmt_ids == nullptr))) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_add_batch: non-empty array is NULL");
  }

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  // Stage every mutation, including the default roots that are required for
  // a valid xf. StylesTable owns strings and all nested records by value, so
  // this is a deep copy. The live table and caller-owned output arrays remain
  // untouched until every input record has passed validation.
  formulon::io::StylesTable staged = styles;
  ensure_default_style_roots(staged);

  std::vector<uint32_t> staged_font_indices(batch->font_count);
  std::vector<uint32_t> staged_fill_indices(batch->fill_count);
  std::vector<uint32_t> staged_border_indices(batch->border_count);
  std::vector<uint32_t> staged_cell_xf_indices(batch->cell_xf_count);
  std::vector<uint16_t> staged_num_fmt_ids(batch->num_fmt_count);

  std::unordered_map<std::string, uint32_t> fonts;
  std::unordered_map<std::string, uint32_t> fills;
  std::unordered_map<std::string, uint32_t> borders;
  fonts.reserve(staged.fonts.size() + batch->font_count);
  fills.reserve(staged.fills.size() + batch->fill_count);
  borders.reserve(staged.borders.size() + batch->border_count);
  for (uint32_t i = 0; i < staged.fonts.size(); ++i)
    fonts.emplace(font_key(staged.fonts[i]), i);
  for (uint32_t i = 0; i < staged.fills.size(); ++i)
    fills.emplace(fill_key(staged.fills[i]), i);
  for (uint32_t i = 0; i < staged.borders.size(); ++i)
    borders.emplace(border_key(staged.borders[i]), i);
  for (size_t i = 0; i < batch->font_count; ++i) {
    formulon::io::FontRecord value = font_from_c(batch->fonts[i]);
    const std::string key = font_key(value);
    const auto [it, inserted] = fonts.emplace(key, static_cast<uint32_t>(staged.fonts.size()));
    if (inserted)
      staged.fonts.push_back(std::move(value));
    staged_font_indices[i] = it->second;
  }
  for (size_t i = 0; i < batch->fill_count; ++i) {
    formulon::io::FillRecord value = fill_from_c(batch->fills[i]);
    const std::string key = fill_key(value);
    const auto [it, inserted] = fills.emplace(key, static_cast<uint32_t>(staged.fills.size()));
    if (inserted)
      staged.fills.push_back(value);
    staged_fill_indices[i] = it->second;
  }
  for (size_t i = 0; i < batch->border_count; ++i) {
    formulon::io::BorderRecord value = border_from_c(batch->borders[i]);
    const std::string key = border_key(value);
    const auto [it, inserted] = borders.emplace(key, static_cast<uint32_t>(staged.borders.size()));
    if (inserted)
      staged.borders.push_back(value);
    staged_border_indices[i] = it->second;
  }
  std::unordered_map<std::string, uint16_t> num_fmts;
  num_fmts.reserve(staged.num_fmts.size() + batch->num_fmt_count);
  std::uint32_t next_num_fmt_id = 163U;
  for (const formulon::io::NumFmtRecord& record : staged.num_fmts) {
    if (record.format_string_index < staged.num_fmt_strings.size()) {
      num_fmts.emplace(staged.num_fmt_strings[record.format_string_index], record.id);
    }
    if (record.id > next_num_fmt_id) {
      next_num_fmt_id = record.id;
    }
  }
  for (size_t i = 0; i < batch->num_fmt_count; ++i) {
    const std::string code = batch->num_fmt_codes[i] != nullptr ? batch->num_fmt_codes[i] : "";
    uint16_t builtin_id = 0U;
    bool is_builtin = false;
    for (uint16_t id = 0U; id < 164U; ++id) {
      const char* builtin = formulon::io::builtin_num_fmt(id);
      if (builtin != nullptr && builtin[0] != '\0' && code == builtin) {
        builtin_id = id;
        is_builtin = true;
        break;
      }
    }
    if (is_builtin) {
      staged_num_fmt_ids[i] = builtin_id;
      continue;
    }
    const auto existing = num_fmts.find(code);
    if (existing != num_fmts.end()) {
      staged_num_fmt_ids[i] = existing->second;
      continue;
    }
    if (next_num_fmt_id >= std::numeric_limits<std::uint16_t>::max()) {
      return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed,
                               "fm_styles_add_batch: custom numFmtId space exhausted",
                               "max_num_fmt_id=" + std::to_string(std::numeric_limits<std::uint16_t>::max()));
    }
    const std::uint32_t candidate_num_fmt_id = next_num_fmt_id + 1U;
    num_fmts.emplace(code, static_cast<std::uint16_t>(candidate_num_fmt_id));
    ++next_num_fmt_id;
    staged.num_fmt_strings.push_back(code);
    formulon::io::NumFmtRecord record;
    record.id = static_cast<std::uint16_t>(next_num_fmt_id);
    record.format_string_index = static_cast<uint32_t>(staged.num_fmt_strings.size() - 1U);
    staged.num_fmts.push_back(record);
    staged_num_fmt_ids[i] = record.id;
  }
  ensure_default_cell_xf(staged);
  std::unordered_map<std::string, uint32_t> xfs;
  xfs.reserve(staged.cell_xfs.size() + batch->cell_xf_count);
  for (uint32_t i = 0; i < staged.cell_xfs.size(); ++i)
    xfs.emplace(cell_xf_key(staged.cell_xfs[i]), i);
  for (size_t i = 0; i < batch->cell_xf_count; ++i) {
    const fm_cell_xf record = batch->cell_xfs[i];
    if (fm_status_t rc = validate_xf_alignment_enums(record, "fm_styles_add_batch"); rc != 0) {
      return rc;
    }
    if (fm_status_t rc = validate_xf_references(staged, record, "fm_styles_add_batch"); rc != 0) {
      return rc;
    }
    formulon::io::CellXf value;
    value.font_index = record.font_index;
    value.fill_index = record.fill_index;
    value.border_index = record.border_index;
    value.num_fmt_id = record.num_fmt_id;
    value.horizontal_align = record.horizontal_align;
    value.vertical_align = record.vertical_align;
    value.wrap_text = record.wrap_text != 0;
    value.has_horizontal_align = value.horizontal_align != 0U;
    value.has_vertical_align = value.vertical_align != 2U;
    value.has_wrap_text = value.wrap_text;
    value.has_alignment = formulon::io::HasAlignment(value);
    const std::string key = cell_xf_key(value);
    const auto [it, inserted] = xfs.emplace(key, static_cast<uint32_t>(staged.cell_xfs.size()));
    if (inserted)
      staged.cell_xfs.push_back(value);
    staged_cell_xf_indices[i] = it->second;
  }

  // Commit the complete staged table and publish output indices only after
  // all records have succeeded. A failed batch therefore leaves both the
  // workbook and every caller-provided sentinel byte-identical.
  styles = std::move(staged);
  for (size_t i = 0; i < batch->font_count; ++i)
    batch->font_indices[i] = staged_font_indices[i];
  for (size_t i = 0; i < batch->fill_count; ++i)
    batch->fill_indices[i] = staged_fill_indices[i];
  for (size_t i = 0; i < batch->border_count; ++i)
    batch->border_indices[i] = staged_border_indices[i];
  for (size_t i = 0; i < batch->cell_xf_count; ++i)
    batch->cell_xf_indices[i] = staged_cell_xf_indices[i];
  for (size_t i = 0; i < batch->num_fmt_count; ++i)
    batch->num_fmt_ids[i] = staged_num_fmt_ids[i];
  return 0;
}

extern "C" fm_status_t fm_styles_add_dxf(fm_workbook_t* wb, fm_dxf_record record, uint32_t* out_dxf_index) {
  clear_last_error();
  if (wb == nullptr || out_dxf_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_dxf: NULL argument");
  }

  formulon::io::DifferentialFormat candidate;
  candidate.has_font = record.font_engaged != 0;
  if (candidate.has_font) {
    candidate.font = font_from_c(record.font);
  }
  candidate.has_fill = record.fill_engaged != 0;
  if (candidate.has_fill) {
    candidate.fill = fill_from_c(record.fill);
  }
  candidate.has_border = record.border_engaged != 0;
  if (candidate.has_border) {
    candidate.border = border_from_c(record.border);
  }
  candidate.has_num_fmt = record.num_fmt_engaged != 0;
  if (candidate.has_num_fmt) {
    candidate.num_fmt_id = record.num_fmt_id;
    candidate.num_fmt_code = (record.num_fmt_code != nullptr) ? std::string(record.num_fmt_code) : std::string();
  }

  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.dxfs.size(); ++i) {
    if (dxf_records_equal(styles.dxfs[i], candidate)) {
      *out_dxf_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.dxfs.push_back(std::move(candidate));
  *out_dxf_index = static_cast<uint32_t>(styles.dxfs.size() - 1);
  return 0;
}
