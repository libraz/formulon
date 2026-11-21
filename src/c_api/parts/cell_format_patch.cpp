//
// C ABI - cell-format patch: `fm_sheet_update_cell_format` merges a font / fill / border / alignment / protection /
// number-format patch into every XF a selection resolves to and remaps the selection onto the results.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/styles_records.h"
#include "font_completion.h"
#include "io/theme_part.h"
#include "style_region.h"
#include "styles.h"
#include "workbook.h"

using formulon::c_api::parts::add_num_fmt_code;
using formulon::c_api::parts::border_key;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_font_size_pt;
using formulon::c_api::parts::check_patch_op;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_sheet_use;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_from_c;
using formulon::c_api::parts::ensure_default_cell_xf;
using formulon::c_api::parts::fill_from_c;
using formulon::c_api::parts::fill_key;
using formulon::c_api::parts::find_or_append_record;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::font_key;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::patch_op_is_set;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::SheetUse;
using formulon::c_api::parts::validate;

namespace {

// A font charset is one byte in both file formats.
// @limit: format
constexpr std::int64_t kMaxCharset = 255;
// The xlsb cell format stores the indent level in one byte.
// @limit: format
constexpr std::uint32_t kMaxIndent = 255U;
// Rotation degrees run 0..180 in both file formats; 255 is the vertical-text marker.
// @limit: format
constexpr std::uint32_t kMaxTextRotation = 180U;
constexpr std::uint32_t kVerticalTextRotation = 255U;
constexpr std::int32_t kLastReadingOrder = FM_READING_ORDER_RIGHT_TO_LEFT;

bool states(fm_patch_op_t op) noexcept {
  return op != FM_PATCH_OP_KEEP;
}

bool states_any(const fm_font_patch_t& p) noexcept {
  return states(p.name_op) || states(p.size_pt_op) || states(p.bold_op) || states(p.italic_op) || states(p.strike_op) ||
         states(p.outline_op) || states(p.shadow_op) || states(p.condense_op) || states(p.extend_op) ||
         states(p.underline_op) || states(p.vert_align_op) || states(p.family_op) || states(p.charset_op) ||
         states(p.scheme_op) || states(p.color_op);
}

bool states_any(const fm_fill_patch_t& p) noexcept {
  return states(p.pattern_op) || states(p.foreground_op) || states(p.background_op) || states(p.gradient_op);
}

bool states_any(const fm_border_side_patch_t& p) noexcept {
  return states(p.style_op) || states(p.color_op);
}

bool states_any(const fm_border_patch_t& p) noexcept {
  return states_any(p.left) || states_any(p.right) || states_any(p.top) || states_any(p.bottom) ||
         states_any(p.diagonal) || states(p.diagonal_up_op) || states(p.diagonal_down_op);
}

bool states_any(const fm_alignment_patch_t& p) noexcept {
  return states(p.horizontal_op) || states(p.vertical_op) || states(p.wrap_text_op) || states(p.shrink_to_fit_op) ||
         states(p.indent_op) || states(p.text_rotation_op) || states(p.reading_order_op) ||
         states(p.justify_last_line_op) || states(p.relative_indent_op);
}

bool states_any(const fm_cell_protection_patch_t& p) noexcept {
  return states(p.locked_op) || states(p.hidden_op);
}

// ---- validation -------------------------------------------------------------------------------------------------

// `value` is an ordinal of `0..max` when `op` sets it.
fm_status_t check_set_enum(fm_patch_op_t op, std::int64_t value, std::int64_t max, const char* field) {
  return patch_op_is_set(op) ? check_enum_domain(value, 0, max, field) : 0;
}

fm_status_t check_set_color(fm_patch_op_t op, const fm_color_t& color, const char* field) {
  return check_set_enum(op, color.kind, FM_COLOR_KIND_AUTO, field);
}

// Every `(op, field)` of `ops` is KEEP or SET, or CLEAR where `clearable`.
struct OpField {
  fm_patch_op_t op;
  const char* field;
};

fm_status_t check_ops(std::initializer_list<OpField> ops, bool clearable) {
  for (const OpField& entry : ops) {
    if (auto rc = check_patch_op(entry.op, entry.field, clearable); rc != 0) {
      return rc;
    }
  }
  return 0;
}

fm_status_t validate_font(const fm_font_patch_t& p) {
  if (auto rc = check_ops({{p.name_op, "font.name"},
                           {p.size_pt_op, "font.size_pt"},
                           {p.bold_op, "font.bold"},
                           {p.italic_op, "font.italic"},
                           {p.strike_op, "font.strike"},
                           {p.outline_op, "font.outline"},
                           {p.shadow_op, "font.shadow"},
                           {p.condense_op, "font.condense"},
                           {p.extend_op, "font.extend"},
                           {p.family_op, "font.family"},
                           {p.charset_op, "font.charset"},
                           {p.color_op, "font.color"}},
                          true);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_ops(
          {{p.underline_op, "font.underline"}, {p.vert_align_op, "font.vert_align"}, {p.scheme_op, "font.scheme"}},
          false);
      rc != 0) {
    return rc;
  }
  if (patch_op_is_set(p.name_op)) {
    if (p.name == nullptr) {
      return null_field("font.name");
    }
    if (auto rc = check_utf8(p.name, "field", "font.name"); rc != 0) {
      return rc;
    }
  }
  if (patch_op_is_set(p.size_pt_op)) {
    if (!std::isfinite(p.size_pt)) {
      return invalid_argument("font.size_pt must be finite");
    }
    if (auto rc = check_font_size_pt(p.size_pt); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_set_enum(p.underline_op, p.underline, FM_UNDERLINE_STYLE_DOUBLE_ACCOUNTING, "font.underline");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.vert_align_op, p.vert_align, FM_FONT_VERTICAL_ALIGN_SUBSCRIPT, "font.vert_align");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.family_op, p.family, FM_FONT_FAMILY_DECORATIVE, "font.family"); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.charset_op, p.charset, kMaxCharset, "font.charset"); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.scheme_op, p.scheme, FM_FONT_SCHEME_MINOR, "font.scheme"); rc != 0) {
    return rc;
  }
  return check_set_color(p.color_op, p.color, "font.color.kind");
}

// The `fm_fill_t` a gradient `SET` states: the caller's gradient fields, read as a gradient fill.
fm_fill_t gradient_fill_of(const fm_fill_patch_t& p) {
  fm_fill_t fill = *p.gradient;
  fill.has_gradient = 1;
  fill.pattern = FM_FILL_PATTERN_NONE;
  return fill;
}

fm_status_t validate_fill(const fm_fill_patch_t& p) {
  if (auto rc = check_ops(
          {{p.pattern_op, "fill.pattern"}, {p.foreground_op, "fill.foreground"}, {p.background_op, "fill.background"}},
          false);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_patch_op(p.gradient_op, "fill.gradient", true); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.pattern_op, p.pattern, FM_FILL_PATTERN_GRAY0625, "fill.pattern"); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_color(p.foreground_op, p.foreground, "fill.foreground.kind"); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_color(p.background_op, p.background, "fill.background.kind"); rc != 0) {
    return rc;
  }
  if (!patch_op_is_set(p.gradient_op)) {
    return 0;
  }
  if (p.gradient == nullptr) {
    return null_field("fill.gradient");
  }
  if (p.gradient->gradient_stops == nullptr && p.gradient->gradient_stop_count > 0U) {
    return null_field("fill.gradient.gradient_stops");
  }
  return validate(gradient_fill_of(p));
}

fm_status_t validate_side(const fm_border_side_patch_t& p, const char* style_field, const char* color_field) {
  if (auto rc = check_ops({{p.style_op, style_field}, {p.color_op, color_field}}, false); rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.style_op, p.style, FM_BORDER_STYLE_SLANT_DASH_DOT, style_field); rc != 0) {
    return rc;
  }
  return check_set_color(p.color_op, p.color, color_field);
}

fm_status_t validate_border(const fm_border_patch_t& p) {
  if (auto rc =
          check_ops({{p.diagonal_up_op, "border.diagonal_up"}, {p.diagonal_down_op, "border.diagonal_down"}}, false);
      rc != 0) {
    return rc;
  }
  if (auto rc = validate_side(p.left, "border.left.style", "border.left.color.kind"); rc != 0) {
    return rc;
  }
  if (auto rc = validate_side(p.right, "border.right.style", "border.right.color.kind"); rc != 0) {
    return rc;
  }
  if (auto rc = validate_side(p.top, "border.top.style", "border.top.color.kind"); rc != 0) {
    return rc;
  }
  if (auto rc = validate_side(p.bottom, "border.bottom.style", "border.bottom.color.kind"); rc != 0) {
    return rc;
  }
  return validate_side(p.diagonal, "border.diagonal.style", "border.diagonal.color.kind");
}

fm_status_t validate_alignment(const fm_alignment_patch_t& p) {
  if (auto rc = check_ops({{p.horizontal_op, "alignment.horizontal"},
                           {p.vertical_op, "alignment.vertical"},
                           {p.wrap_text_op, "alignment.wrap_text"},
                           {p.shrink_to_fit_op, "alignment.shrink_to_fit"},
                           {p.indent_op, "alignment.indent"},
                           {p.text_rotation_op, "alignment.text_rotation"},
                           {p.reading_order_op, "alignment.reading_order"},
                           {p.justify_last_line_op, "alignment.justify_last_line"},
                           {p.relative_indent_op, "alignment.relative_indent"}},
                          false);
      rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.horizontal_op, p.horizontal, FM_HORIZONTAL_ALIGN_DISTRIBUTED, "alignment.horizontal");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.vertical_op, p.vertical, FM_VERTICAL_ALIGN_DISTRIBUTED, "alignment.vertical");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_set_enum(p.reading_order_op, p.reading_order, kLastReadingOrder, "alignment.reading_order");
      rc != 0) {
    return rc;
  }
  if (patch_op_is_set(p.indent_op) && p.indent > kMaxIndent) {
    return invalid_argument("alignment.indent must be in 0..255", "indent=" + std::to_string(p.indent));
  }
  if (patch_op_is_set(p.text_rotation_op) && p.text_rotation > kMaxTextRotation &&
      p.text_rotation != kVerticalTextRotation) {
    return invalid_argument("alignment.text_rotation must be in 0..180 or 255",
                            "text_rotation=" + std::to_string(p.text_rotation));
  }
  return 0;
}

fm_status_t validate_protection(const fm_cell_protection_patch_t& p) {
  return check_ops({{p.locked_op, "protection.locked"}, {p.hidden_op, "protection.hidden"}}, false);
}

// ---- merging ----------------------------------------------------------------------------------------------------

// An optional on/off style flag: SET writes it (explicit off included), CLEAR returns it to absent.
void apply_toggle(fm_patch_op_t op, int32_t value, bool* has, bool* flag) noexcept {
  if (patch_op_is_set(op)) {
    *has = true;
    *flag = value != 0;
  } else if (op == FM_PATCH_OP_CLEAR) {
    *has = false;
    *flag = false;
  }
}

// The family and charset a font renamed to a new face takes when the patch states neither: the Normal font's own when
// the face is the Normal face, else the measured row of the face under `locale`, else neither.
void complete_renamed_font(formulon::FontRecord* font, const formulon::FontRecord& normal,
                           formulon::ExcelLocale locale) {
  if (font->name == normal.name) {
    font->has_family = normal.has_family;
    font->family = normal.family;
    font->has_charset = normal.has_charset;
    font->charset = normal.charset;
    return;
  }
  const formulon::FontCompletion* row = formulon::find_font_completion(font->name);
  font->family = row != nullptr ? formulon::completion_family(*row, locale) : 0U;
  font->has_family = font->family != 0U;
  font->charset = row != nullptr ? row->charset : 0U;
  font->has_charset = font->charset != 0U;
}

formulon::FontRecord merged_font(const formulon::FontRecord& current, const formulon::FontRecord& normal,
                                 formulon::ExcelLocale locale, const fm_font_patch_t& p) {
  formulon::FontRecord out = current;
  if (patch_op_is_set(p.name_op)) {
    out.name = p.name;
    if (!patch_op_is_set(p.scheme_op)) {
      out.scheme = 0U;
    }
    if (out.name != current.name && !states(p.family_op) && !states(p.charset_op)) {
      complete_renamed_font(&out, normal, locale);
    }
  } else if (p.name_op == FM_PATCH_OP_CLEAR) {
    out.name = normal.name;
  }
  if (patch_op_is_set(p.size_pt_op)) {
    out.size = p.size_pt;
  } else if (p.size_pt_op == FM_PATCH_OP_CLEAR) {
    out.size = normal.size;
  }
  out.has_size = true;
  apply_toggle(p.bold_op, p.bold, &out.has_bold, &out.bold);
  apply_toggle(p.italic_op, p.italic, &out.has_italic, &out.italic);
  apply_toggle(p.strike_op, p.strike, &out.has_strike, &out.strike);
  apply_toggle(p.outline_op, p.outline, &out.has_outline, &out.outline);
  apply_toggle(p.shadow_op, p.shadow, &out.has_shadow, &out.shadow);
  apply_toggle(p.condense_op, p.condense, &out.has_condense, &out.condense);
  apply_toggle(p.extend_op, p.extend, &out.has_extend, &out.extend);
  if (patch_op_is_set(p.underline_op)) {
    out.underline = static_cast<std::uint8_t>(p.underline);
  }
  if (patch_op_is_set(p.vert_align_op)) {
    out.vert_align = static_cast<std::uint8_t>(p.vert_align);
  }
  if (patch_op_is_set(p.family_op)) {
    out.has_family = true;
    out.family = static_cast<std::uint8_t>(p.family);
  } else if (p.family_op == FM_PATCH_OP_CLEAR) {
    out.has_family = false;
    out.family = 0U;
  }
  if (patch_op_is_set(p.charset_op)) {
    out.has_charset = true;
    out.charset = static_cast<std::uint8_t>(p.charset);
  } else if (p.charset_op == FM_PATCH_OP_CLEAR) {
    out.has_charset = false;
    out.charset = 0U;
  }
  if (patch_op_is_set(p.scheme_op)) {
    out.scheme = static_cast<std::uint8_t>(p.scheme);
  }
  if (patch_op_is_set(p.color_op)) {
    out.color = color_from_c(p.color, &out.color_argb);
  } else if (p.color_op == FM_PATCH_OP_CLEAR) {
    out.color = color_from_c(fm_color_t{}, &out.color_argb);
  }
  return out;
}

formulon::FillRecord merged_fill(const formulon::FillRecord& current, const fm_fill_patch_t& p) {
  formulon::FillRecord out = current;
  out.pattern_stated = true;
  // Any change to a fill drops a gradient unless the patch sets one.
  out.has_gradient = false;
  out.gradient = formulon::GradientFill{};
  if (patch_op_is_set(p.pattern_op)) {
    out.pattern = static_cast<std::uint8_t>(p.pattern);
  }
  if (patch_op_is_set(p.foreground_op)) {
    out.fg = color_from_c(p.foreground, &out.fg_argb);
  }
  if (patch_op_is_set(p.background_op)) {
    out.bg = color_from_c(p.background, &out.bg_argb);
  }
  const bool colored = patch_op_is_set(p.foreground_op) || patch_op_is_set(p.background_op);
  if (colored && !patch_op_is_set(p.pattern_op) && out.pattern == 0U) {
    out.pattern = 1U;
  }
  if (patch_op_is_set(p.gradient_op)) {
    // The gradient's own pattern and colours do not exist; `fill_from_c` yields the canonical gradient record.
    out = fill_from_c(gradient_fill_of(p));
  }
  return out;
}

void merge_side(formulon::BorderSide* side, const fm_border_side_patch_t& p) {
  if (patch_op_is_set(p.style_op)) {
    side->style = static_cast<std::uint8_t>(p.style);
  }
  if (patch_op_is_set(p.color_op)) {
    side->color = color_from_c(p.color, &side->color_argb);
  }
}

formulon::BorderRecord merged_border(const formulon::BorderRecord& current, const fm_border_patch_t& p) {
  formulon::BorderRecord out = current;
  merge_side(&out.left, p.left);
  merge_side(&out.right, p.right);
  merge_side(&out.top, p.top);
  merge_side(&out.bottom, p.bottom);
  merge_side(&out.diagonal, p.diagonal);
  if (patch_op_is_set(p.diagonal_up_op)) {
    out.diagonal_up = p.diagonal_up != 0;
  }
  if (patch_op_is_set(p.diagonal_down_op)) {
    out.diagonal_down = p.diagonal_down != 0;
  }
  return out;
}

void merge_alignment(formulon::CellXf* xf, const fm_alignment_patch_t& p) {
  if (patch_op_is_set(p.horizontal_op)) {
    xf->has_horizontal_align = true;
    xf->horizontal_align = static_cast<std::uint8_t>(p.horizontal);
  }
  if (patch_op_is_set(p.vertical_op)) {
    xf->has_vertical_align = true;
    xf->vertical_align = static_cast<std::uint8_t>(p.vertical);
  }
  if (patch_op_is_set(p.wrap_text_op)) {
    xf->has_wrap_text = true;
    xf->wrap_text = p.wrap_text != 0;
  }
  if (patch_op_is_set(p.shrink_to_fit_op)) {
    xf->has_shrink_to_fit = true;
    xf->shrink_to_fit = p.shrink_to_fit != 0;
  }
  if (patch_op_is_set(p.indent_op)) {
    xf->has_indent = true;
    xf->indent = p.indent;
  }
  if (patch_op_is_set(p.text_rotation_op)) {
    xf->has_text_rotation = true;
    xf->text_rotation = p.text_rotation;
  }
  if (patch_op_is_set(p.reading_order_op)) {
    xf->has_reading_order = true;
    xf->reading_order = static_cast<std::uint32_t>(p.reading_order);
  }
  if (patch_op_is_set(p.justify_last_line_op)) {
    xf->has_justify_last_line = true;
    xf->justify_last_line = p.justify_last_line != 0;
  }
  if (patch_op_is_set(p.relative_indent_op)) {
    xf->has_relative_indent = true;
    xf->relative_indent = p.relative_indent;
  }
  xf->has_alignment = formulon::HasAlignment(*xf);
}

void merge_protection(formulon::CellXf* xf, const fm_cell_protection_patch_t& p) {
  xf->has_protection = true;
  if (patch_op_is_set(p.locked_op)) {
    xf->locked = p.locked != 0;
  }
  if (patch_op_is_set(p.hidden_op)) {
    xf->hidden = p.hidden != 0;
  }
}

// `table[index]`, or a default record for an index past the table (a package that left the table short).
template <typename Record>
Record record_or_default(const std::vector<Record>& table, std::uint32_t index) {
  return index < table.size() ? table[index] : Record{};
}

}  // namespace

extern "C" fm_status_t fm_sheet_update_cell_format(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* range,
                                                   const fm_cell_format_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || range == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_use(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*range); rc != 0) {
    return rc;
  }
  if (auto rc = check_patch_op(patch->number_format_op, "number_format", false); rc != 0) {
    return rc;
  }
  if (patch_op_is_set(patch->number_format_op)) {
    if (patch->number_format == nullptr) {
      return null_field("number_format");
    }
    if (patch->number_format[0] == '\0') {
      return invalid_argument("number_format must not be empty", "number_format=");
    }
    if (auto rc = check_utf8(patch->number_format, "field", "number_format"); rc != 0) {
      return rc;
    }
  }
  if (auto rc = validate_font(patch->font); rc != 0) {
    return rc;
  }
  if (auto rc = validate_fill(patch->fill); rc != 0) {
    return rc;
  }
  if (auto rc = validate_border(patch->border); rc != 0) {
    return rc;
  }
  if (auto rc = validate_alignment(patch->alignment); rc != 0) {
    return rc;
  }
  if (auto rc = validate_protection(patch->protection); rc != 0) {
    return rc;
  }

  const bool number_format = states(patch->number_format_op);
  const bool font = states_any(patch->font);
  const bool fill = states_any(patch->fill);
  const bool border = states_any(patch->border);
  const bool alignment = states_any(patch->alignment);
  const bool protection = states_any(patch->protection);
  if (!number_format && !font && !fill && !border && !alignment && !protection) {
    return 0;
  }

  const formulon::StyleRect rect{range->first_row, range->first_col, range->last_row, range->last_col};
  if (auto area = formulon::check_styled_range_area(rect); !area) {
    return set_last_error(area.error());
  }
  auto bases = wb->workbook().range_style_bases(sheet_index, rect);
  if (!bases) {
    return set_last_error(bases.error());
  }

  // The number format is the one step that can still fail; the tables below only grow.
  std::uint32_t num_fmt_id = 0;
  if (number_format) {
    const std::string code = patch->number_format;
    if (code != "General") {
      if (auto rc = add_num_fmt_code(wb->workbook(), code, &num_fmt_id); rc != 0) {
        return rc;
      }
    }
  }

  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_cell_xf(styles);
  const formulon::FontRecord normal = styles.fonts[0];
  const formulon::ExcelLocale locale = wb->workbook().excel_profile().locale;
  std::vector<formulon::XfRemap> map;
  map.reserve(bases.value().size());
  for (const std::uint32_t base : bases.value()) {
    formulon::CellXf xf = record_or_default(styles.cell_xfs, base);
    if (number_format) {
      xf.apply_number_format = true;
      xf.num_fmt_id = static_cast<std::uint16_t>(num_fmt_id);
    }
    if (font) {
      xf.apply_font = true;
      formulon::FontRecord merged =
          merged_font(record_or_default(styles.fonts, xf.font_index), normal, locale, patch->font);
      if (patch_op_is_set(patch->font.scheme_op)) {
        formulon::io::resolve_scheme_font_name(merged, wb->workbook());
      }
      xf.font_index = find_or_append_record(styles.fonts, std::move(merged), &font_key);
    }
    if (fill) {
      xf.apply_fill = true;
      xf.fill_index = find_or_append_record(
          styles.fills, merged_fill(record_or_default(styles.fills, xf.fill_index), patch->fill), &fill_key);
    }
    if (border) {
      xf.apply_border = true;
      xf.border_index = find_or_append_record(
          styles.borders, merged_border(record_or_default(styles.borders, xf.border_index), patch->border),
          &border_key);
    }
    if (alignment) {
      xf.apply_alignment = true;
      merge_alignment(&xf, patch->alignment);
    }
    if (protection) {
      xf.apply_protection = true;
      merge_protection(&xf, patch->protection);
    }
    map.push_back({base, formulon::intern_cell_xf(styles, std::move(xf))});
  }
  return finish_status(wb->workbook().remap_range_style(sheet_index, rect, std::move(map)));
}
