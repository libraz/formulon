//
// C ABI - styles surface (cell xf bindings, fonts, fills, borders, num
// formats, cell styles). The mutation side lives in styles_mutate.cpp.

#include "styles.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cell.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

using formulon::effective_num_fmt;
using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

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

extern "C" fm_status_t fm_sheet_set_range_xf_index(fm_workbook_t* wb, uint32_t sheet, uint32_t first_row,
                                                   uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                   uint32_t xf_index) {
  static constexpr const char* kFn = "fm_sheet_set_range_xf_index";
  clear_last_error();
  if (auto rc = check_sheet_u32(wb, sheet, kFn); rc != 0) {
    return rc;
  }
  const std::uint32_t row_lo = std::min(first_row, last_row);
  const std::uint32_t row_hi = std::max(first_row, last_row);
  const std::uint32_t col_lo = std::min(first_col, last_col);
  const std::uint32_t col_hi = std::max(first_col, last_col);
  if (row_hi >= formulon::Sheet::kMaxRows || col_hi >= formulon::Sheet::kMaxCols) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, kFn,
                             "last_row=" + std::to_string(row_hi) + " last_col=" + std::to_string(col_hi));
  }
  const std::uint64_t cells =
      (static_cast<std::uint64_t>(row_hi - row_lo) + 1U) * (static_cast<std::uint64_t>(col_hi - col_lo) + 1U);
  if (cells > formulon::kMaxStyledRangeCells) {
    return set_binding_error(
        formulon::FormulonErrorCode::kPreconditionFailed, kFn,
        "cells=" + std::to_string(cells) + " limit=" + std::to_string(formulon::kMaxStyledRangeCells));
  }
  for (std::uint32_t row = row_lo; row <= row_hi; ++row) {
    for (std::uint32_t col = col_lo; col <= col_hi; ++col) {
      auto r = wb->workbook().set_cell_xf_index(sheet, row, col, xf_index);
      if (!r) {
        return set_last_error(r.error());
      }
    }
  }
  return 0;
}

namespace {

/// Projects one engine record onto the flat C record. The whole model is
/// carried, so both XF tables share this and neither getter is lossy.
void cell_xf_to_c(const formulon::CellXf& xf, fm_cell_xf* out) {
  out->font_index = xf.font_index;
  out->fill_index = xf.fill_index;
  out->border_index = xf.border_index;
  out->num_fmt_id = xf.num_fmt_id;
  out->horizontal_align = xf.horizontal_align;
  out->vertical_align = xf.vertical_align;
  out->wrap_text = xf.wrap_text ? 1 : 0;
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
}

}  // namespace

extern "C" fm_status_t fm_styles_get_cell_xf(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_cell_xf: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (xf_index >= styles.cell_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_xf: xf_index out of range",
        "xf_index=" + std::to_string(xf_index) + " cell_xfs_count=" + std::to_string(styles.cell_xfs.size()));
  }
  cell_xf_to_c(styles.cell_xfs[xf_index], out);
  return 0;
}

namespace {

void color_to_c(const formulon::ColorSpec& src, fm_color_spec* out) noexcept {
  out->tint = src.tint;
  out->rgb = src.rgb;
  out->theme = src.theme;
  out->indexed = src.indexed;
  out->kind = static_cast<uint8_t>(src.kind);
}

/// Projects a model font onto the public record. Every field the styles
/// writer can observe is carried, including the presence flags and the
/// round-trip colour specification, so `fm_styles_add_font(out)` names the
/// record it was read from instead of appending a lossy copy of it.
void font_to_c(const formulon::FontRecord& f, fm_font_record* out) noexcept {
  out->name = f.name.c_str();
  out->size = f.size;
  out->color_argb = f.color_argb;
  out->bold = f.bold ? 1 : 0;
  out->italic = f.italic ? 1 : 0;
  out->strike = f.strike ? 1 : 0;
  out->has_bold = f.has_bold ? 1 : 0;
  out->has_italic = f.has_italic ? 1 : 0;
  out->has_strike = f.has_strike ? 1 : 0;
  out->has_family = f.has_family ? 1 : 0;
  out->has_charset = f.has_charset ? 1 : 0;
  out->underline = f.underline;
  out->vert_align = f.vert_align;
  out->family = f.family;
  out->charset = f.charset;
  out->scheme = f.scheme;
  color_to_c(f.color, &out->color);
}

void fill_to_c(const formulon::FillRecord& f, fm_fill_record* out) noexcept {
  out->pattern = f.pattern;
  out->fg_argb = f.fg_argb;
  out->bg_argb = f.bg_argb;
  color_to_c(f.fg, &out->fg);
  color_to_c(f.bg, &out->bg);
}

void border_side_to_c(const formulon::BorderSide& src, fm_border_side* out) noexcept {
  out->style = src.style;
  out->color_argb = src.color_argb;
  color_to_c(src.color, &out->color);
}

void border_to_c(const formulon::BorderRecord& b, fm_border_record* out) noexcept {
  border_side_to_c(b.left, &out->left);
  border_side_to_c(b.right, &out->right);
  border_side_to_c(b.top, &out->top);
  border_side_to_c(b.bottom, &out->bottom);
  border_side_to_c(b.diagonal, &out->diagonal);
  out->diagonal_up = b.diagonal_up ? 1 : 0;
  out->diagonal_down = b.diagonal_down ? 1 : 0;
}

}  // namespace

extern "C" fm_status_t fm_styles_get_font(fm_workbook_t* wb, uint32_t font_index, fm_font_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_font: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (font_index >= styles.fonts.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_font: font_index out of range",
        "font_index=" + std::to_string(font_index) + " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  font_to_c(styles.fonts[font_index], out);
  return 0;
}

extern "C" fm_status_t fm_styles_get_num_fmt_string(fm_workbook_t* wb, uint16_t num_fmt_id, const char** out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_num_fmt_string: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  const std::optional<std::string_view> resolved = effective_num_fmt(styles, num_fmt_id);
  if (resolved.has_value()) {
    *out = resolved->data();
    return 0;
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_num_fmt_string: id not found",
                           "num_fmt_id=" + std::to_string(num_fmt_id));
}

extern "C" fm_status_t fm_styles_get_fill(fm_workbook_t* wb, uint32_t fill_index, fm_fill_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_fill: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (fill_index >= styles.fills.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_fill: fill_index out of range",
        "fill_index=" + std::to_string(fill_index) + " fills_count=" + std::to_string(styles.fills.size()));
  }
  fill_to_c(styles.fills[fill_index], out);
  return 0;
}

extern "C" fm_status_t fm_styles_get_border(fm_workbook_t* wb, uint32_t border_index, fm_border_record* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_border: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (border_index >= styles.borders.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_border: border_index out of range",
        "border_index=" + std::to_string(border_index) + " borders_count=" + std::to_string(styles.borders.size()));
  }
  border_to_c(styles.borders[border_index], out);
  return 0;
}

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
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (dxf_index >= styles.dxfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_dxf: dxf_index out of range",
        "dxf_index=" + std::to_string(dxf_index) + " dxfs_count=" + std::to_string(styles.dxfs.size()));
  }
  *out = fm_dxf_record{};
  const formulon::DifferentialFormat& dxf = styles.dxfs[dxf_index];
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
  out->alignment_xml = dxf.alignment_xml.c_str();
  out->protection_xml = dxf.protection_xml.c_str();
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
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_styles.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style: index out of range",
        "index=" + std::to_string(index) + " cell_styles_count=" + std::to_string(styles.cell_styles.size()));
  }
  const formulon::CellStyleRecord& cs = styles.cell_styles[index];
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
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf: NULL argument");
  }
  const formulon::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_style_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style_xf: index out of range",
        "index=" + std::to_string(index) + " cell_style_xfs_count=" + std::to_string(styles.cell_style_xfs.size()));
  }
  cell_xf_to_c(styles.cell_style_xfs[index], out);
  // A named-style xf is itself the parent, so it never carries an `xfId`.
  out->xf_id = 0;
  return 0;
}
