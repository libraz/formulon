// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - styles surface (cell xf bindings, fonts, fills, borders, num
// formats, cell styles, dedup-on-insert helpers).

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>

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

extern "C" fm_status_t fm_styles_get_cell_xf(fm_workbook_t* wb, uint32_t xf_index, fm_cell_xf* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_get_cell_xf: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (xf_index >= styles.cell_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_xf: xf_index out of range",
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
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_styles_get_cell_style_xf: NULL argument");
  }
  const formulon::io::StylesTable& styles = wb->workbook().styles();
  if (index >= styles.cell_style_xfs.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_get_cell_style_xf: index out of range",
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

namespace {

/// Field-for-field equality for the C++ side `FontRecord`. Excluded
/// from `<algorithm>` because the struct has no `operator==`; defining
/// one here keeps the dedup-comparator confined to the C ABI's needs.
bool font_records_equal(const formulon::io::FontRecord& a, const formulon::io::FontRecord& b) {
  return a.name == b.name && a.size == b.size && a.bold == b.bold && a.italic == b.italic && a.strike == b.strike &&
         a.underline == b.underline && a.color_argb == b.color_argb;
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

bool cell_xfs_equal(const formulon::io::CellXf& a, const formulon::io::CellXf& b) noexcept {
  return a.font_index == b.font_index && a.fill_index == b.fill_index && a.border_index == b.border_index &&
         a.num_fmt_id == b.num_fmt_id && a.horizontal_align == b.horizontal_align &&
         a.vertical_align == b.vertical_align && a.wrap_text == b.wrap_text;
}

}  // namespace

extern "C" fm_status_t fm_styles_add_font(fm_workbook_t* wb, fm_font_record record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_font: NULL argument");
  }
  formulon::io::FontRecord candidate;
  candidate.name = (record.name != nullptr) ? std::string(record.name) : std::string();
  candidate.size = record.size;
  candidate.bold = record.bold != 0;
  candidate.italic = record.italic != 0;
  candidate.strike = record.strike != 0;
  candidate.underline = record.underline;
  candidate.color_argb = record.color_argb;

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
  formulon::io::FillRecord candidate;
  candidate.pattern = record.pattern;
  candidate.fg_argb = record.fg_argb;
  candidate.bg_argb = record.bg_argb;

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
  auto pull_side = [](const fm_border_side& src) noexcept {
    formulon::io::BorderSide dst;
    dst.style = src.style;
    dst.color_argb = src.color_argb;
    return dst;
  };
  formulon::io::BorderRecord candidate;
  candidate.left = pull_side(record.left);
  candidate.right = pull_side(record.right);
  candidate.top = pull_side(record.top);
  candidate.bottom = pull_side(record.bottom);
  candidate.diagonal = pull_side(record.diagonal);
  candidate.diagonal_up = record.diagonal_up != 0;
  candidate.diagonal_down = record.diagonal_down != 0;

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
  uint16_t next_id = 163U;
  for (const formulon::io::NumFmtRecord& n : styles.num_fmts) {
    if (n.id > next_id) {
      next_id = n.id;
    }
  }
  ++next_id;

  styles.num_fmt_strings.push_back(code);
  formulon::io::NumFmtRecord rec;
  rec.id = next_id;
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1);
  styles.num_fmts.push_back(rec);
  *out_num_fmt_id = next_id;
  return 0;
}

extern "C" fm_status_t fm_styles_add_cell_xf(fm_workbook_t* wb, fm_cell_xf record, uint32_t* out_xf_index) {
  clear_last_error();
  if (wb == nullptr || out_xf_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_styles_add_cell_xf: NULL argument");
  }
  formulon::io::StylesTable& styles = wb->workbook().mutable_styles();

  // Validate referenced indices. Reject out-of-range references rather
  // than auto-growing the parallel tables; callers must register fonts /
  // fills / borders before referencing them.
  if (record.font_index >= styles.fonts.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_add_cell_xf: font_index out of range",
        "font_index=" + std::to_string(record.font_index) + " fonts_count=" + std::to_string(styles.fonts.size()));
  }
  if (record.fill_index >= styles.fills.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_styles_add_cell_xf: fill_index out of range",
        "fill_index=" + std::to_string(record.fill_index) + " fills_count=" + std::to_string(styles.fills.size()));
  }
  if (record.border_index >= styles.borders.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_add_cell_xf: border_index out of range",
                             "border_index=" + std::to_string(record.border_index) +
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
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_styles_add_cell_xf: num_fmt_id not registered",
                             "num_fmt_id=" + std::to_string(record.num_fmt_id));
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
