//
// C ABI - styles mutation surface (fonts, fills, borders, num formats, cell
// xfs, cell styles, dxfs and the batch commit, all with dedup-on-insert).

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <type_traits>
#include <unordered_map>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "io/styles_writer.h"
#include "sheet_name.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/number_text.h"
#include "workbook.h"

using formulon::effective_num_fmt;
using formulon::c_api::parts::add_num_fmt_code;
using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_font_enums;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::num_fmt_id_known;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::validate;

namespace {

static_assert(std::is_nothrow_move_constructible_v<formulon::StylesTable>,
              "StylesTable must be nothrow-move-constructible for staged style commits");
static_assert(std::is_nothrow_move_assignable_v<formulon::StylesTable>,
              "StylesTable must be nothrow-move-assignable for staged style commits");

/// Dedup identity for the font / fill / border tables.
///
/// `fm_styles_add_*` may only hand back an existing index when the styles
/// writer cannot tell the stored record from the candidate, so identity is
/// the XML fragment the writer emits for the record. Deriving the key from
/// the writer rather than restating its rules here folds in every
/// writer-visible field by construction — the presence flags, `<family>` /
/// `<charset>`, and the round-trip colour specifications that a
/// hand-written comparator kept missing — and keeps the linear-scan
/// comparison below and the batch hash key from drifting apart, because
/// they are the same function.
std::string font_key(const formulon::FontRecord& value) {
  return formulon::io::font_fragment(value);
}

std::string fill_key(const formulon::FillRecord& value) {
  return formulon::io::fill_fragment(value);
}

std::string border_key(const formulon::BorderRecord& value) {
  return formulon::io::border_fragment(value);
}

bool font_records_equal(const formulon::FontRecord& a, const formulon::FontRecord& b) {
  return font_key(a) == font_key(b);
}

bool fill_records_equal(const formulon::FillRecord& a, const formulon::FillRecord& b) {
  return fill_key(a) == fill_key(b);
}

bool border_records_equal(const formulon::BorderRecord& a, const formulon::BorderRecord& b) {
  return border_key(a) == border_key(b);
}

/// Index of the record in `table` whose key equals `candidate`'s, appending
/// `candidate` when none does. The candidate's key is built once, so the scan
/// stays linear in the table size.
template <typename Record>
uint32_t find_or_append_record(std::vector<Record>& table, Record candidate, std::string (*key)(const Record&)) {
  const std::string candidate_key = key(candidate);
  for (std::size_t i = 0; i < table.size(); ++i) {
    if (key(table[i]) == candidate_key) {
      return static_cast<uint32_t>(i);
    }
  }
  table.push_back(std::move(candidate));
  return static_cast<uint32_t>(table.size() - 1U);
}

/// A dxf's engaged children are compared through the same writer-derived
/// keys as the global tables, so a differential font that only differs in
/// `<vertAlign>` or in an explicit `<b val="0"/>` is a distinct record.
///
bool dxf_records_equal(const formulon::DifferentialFormat& a, const formulon::DifferentialFormat& b) {
  return a.has_font == b.has_font && (!a.has_font || font_records_equal(a.font, b.font)) && a.has_fill == b.has_fill &&
         (!a.has_fill || fill_records_equal(a.fill, b.fill)) && a.has_border == b.has_border &&
         (!a.has_border || border_records_equal(a.border, b.border)) && a.has_num_fmt == b.has_num_fmt &&
         (!a.has_num_fmt || (a.num_fmt_id == b.num_fmt_id && a.num_fmt_code == b.num_fmt_code)) &&
         a.alignment_xml == b.alignment_xml && a.protection_xml == b.protection_xml;
}

void ensure_default_style_roots(formulon::StylesTable& styles) {
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

void ensure_default_cell_xf(formulon::StylesTable& styles) {
  ensure_default_style_roots(styles);
  if (styles.cell_xfs.empty()) {
    styles.cell_xfs.emplace_back();
  }
}

/// Rejects an `<xf>` whose font / fill / border / number-format reference
/// points past the corresponding table. Auto-growing the parallel tables
/// would hide the caller's ordering mistake and emit a dangling index that
/// Excel reports as a repairable file.
fm_status_t validate_xf_references(const formulon::StylesTable& styles, const fm_cell_xf_t& record) {
  if (auto rc = check_index(record.font_index, styles.fonts.size(), "font_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(record.fill_index, styles.fills.size(), "fill_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(record.border_index, styles.borders.size(), "border_index"); rc != 0) {
    return rc;
  }
  // Stored ids are 16-bit (XLSB writes a 2-byte ifmt), so a wider id is refused rather than truncated.
  if (record.num_fmt_id > std::numeric_limits<std::uint16_t>::max() || !num_fmt_id_known(styles, record.num_fmt_id)) {
    return invalid_argument(
        ("num_fmt_id " + std::to_string(record.num_fmt_id) + " is not a registered number format").c_str(),
        "num_fmt_id=" + std::to_string(record.num_fmt_id));
  }
  return 0;
}

fm_status_t validate_xf_alignment(const fm_cell_xf_t& record) {
  // The presence flags are authoritative. Values for an omitted attribute
  // are ignored, including malformed alignment enum ordinals. The accepted
  // ordinals are the canonical writer mapping for the corresponding XML
  // string enums (general..distributed and top..distributed).
  if (record.has_horizontal != 0 && (record.horizontal < 0 || record.horizontal > 7)) {
    return invalid_argument("horizontal must be in 0..7", "horizontal=" + std::to_string(record.horizontal));
  }
  if (record.has_vertical != 0 && (record.vertical < 0 || record.vertical > 4)) {
    return invalid_argument("vertical must be in 0..4", "vertical=" + std::to_string(record.vertical));
  }
  if (record.has_text_rotation != 0 && record.text_rotation > 180U && record.text_rotation != 255U) {
    return invalid_argument("text_rotation must be in 0..180 or 255",
                            "text_rotation=" + std::to_string(record.text_rotation));
  }
  if (record.has_indent != 0 && record.indent > 255U) {
    return invalid_argument("indent must be in 0..255", "indent=" + std::to_string(record.indent));
  }
  if (record.has_reading_order != 0 && (record.reading_order < 0 || record.reading_order > 2)) {
    return invalid_argument("reading_order must be in 0..2", "reading_order=" + std::to_string(record.reading_order));
  }
  return 0;
}

formulon::CellXf cell_xf_from_c(const fm_cell_xf_t& record, std::uint32_t xf_id) {
  formulon::CellXf candidate;
  candidate.font_index = record.font_index;
  candidate.fill_index = record.fill_index;
  candidate.border_index = record.border_index;
  candidate.num_fmt_id = static_cast<std::uint16_t>(record.num_fmt_id);
  candidate.xf_id = xf_id;

  // The record carries explicit presence bits. Keep values for present
  // attributes and canonicalize omitted ones so Has* does not infer a child
  // from ignored poison values.
  candidate.horizontal_align = record.has_horizontal != 0 ? static_cast<std::uint8_t>(record.horizontal) : 0U;
  candidate.vertical_align = record.has_vertical != 0 ? static_cast<std::uint8_t>(record.vertical) : 2U;
  candidate.wrap_text = record.has_wrap_text != 0 && record.wrap_text != 0;
  candidate.justify_last_line = record.has_justify_last_line != 0 && record.justify_last_line != 0;
  candidate.has_alignment = record.has_alignment != 0;
  candidate.has_horizontal_align = record.has_horizontal != 0;
  candidate.has_vertical_align = record.has_vertical != 0;
  candidate.has_wrap_text = record.has_wrap_text != 0;
  candidate.has_justify_last_line = record.has_justify_last_line != 0;

  // These optional attributes retain their existing presence-plus-value
  // contract; values are copied even when the corresponding bit is false.
  candidate.has_text_rotation = record.has_text_rotation != 0;
  candidate.text_rotation = record.text_rotation;
  candidate.has_indent = record.has_indent != 0;
  candidate.indent = record.indent;
  candidate.has_relative_indent = record.has_relative_indent != 0;
  candidate.relative_indent = record.relative_indent;
  candidate.has_shrink_to_fit = record.has_shrink_to_fit != 0;
  candidate.shrink_to_fit = record.shrink_to_fit != 0;
  candidate.has_reading_order = record.has_reading_order != 0;
  candidate.reading_order = static_cast<std::uint32_t>(record.reading_order);

  candidate.apply_number_format = record.apply_number_format != 0;
  candidate.apply_font = record.apply_font != 0;
  candidate.apply_fill = record.apply_fill != 0;
  candidate.apply_border = record.apply_border != 0;
  candidate.apply_alignment = record.apply_alignment != 0;
  candidate.apply_protection = record.apply_protection != 0;
  candidate.quote_prefix = record.quote_prefix != 0;
  // An omitted `<protection>` child canonicalizes to the schema defaults so
  // ignored values cannot split otherwise identical records.
  candidate.has_protection = record.has_protection != 0;
  candidate.locked = !candidate.has_protection || record.locked != 0;
  candidate.hidden = candidate.has_protection && record.hidden != 0;

  candidate.has_alignment = formulon::HasAlignment(candidate);
  return candidate;
}

/// Accepts a colour specification from the caller. An out-of-range `kind`
/// degrades to `kNone` rather than reaching the writer as an unknown
/// attribute; `out_argb` receives the record's legacy ARGB field.
formulon::ColorSpec color_from_c(fm_color_t src, std::uint32_t* out_argb) {
  if (src.kind < 0 || src.kind > static_cast<int32_t>(formulon::ColorSpec::Kind::kAuto)) {
    src.kind = FM_COLOR_KIND_NONE;
  }
  return formulon::c_api::parts::color_spec_from_c(src, out_argb);
}

formulon::FontRecord font_from_c(const fm_font_t& record) {
  formulon::FontRecord out;
  out.name = text_or_empty(record.name);
  out.size = record.size_pt;
  out.bold = record.bold != 0;
  out.italic = record.italic != 0;
  out.strike = record.strike != 0;
  out.has_bold = record.has_bold != 0;
  out.has_italic = record.has_italic != 0;
  out.has_strike = record.has_strike != 0;
  out.has_family = record.has_family != 0;
  out.has_charset = record.has_charset != 0;
  out.has_size = record.has_size_pt != 0;
  out.outline = record.outline != 0;
  out.shadow = record.shadow != 0;
  out.condense = record.condense != 0;
  out.extend = record.extend != 0;
  out.has_outline = record.has_outline != 0;
  out.has_shadow = record.has_shadow != 0;
  out.has_condense = record.has_condense != 0;
  out.has_extend = record.has_extend != 0;
  out.underline = static_cast<std::uint8_t>(record.underline);
  out.vert_align = static_cast<std::uint8_t>(record.vert_align);
  out.family = static_cast<std::uint8_t>(record.family);
  out.charset = record.charset;
  out.scheme = static_cast<std::uint8_t>(record.scheme);
  out.color = color_from_c(record.color, &out.color_argb);
  return out;
}

/// Cell fonts always carry a size: only a differential font may leave it
/// unchanged. An unstated size or a `NULL` name takes the Normal font's.
formulon::FontRecord cell_font_from_c(const fm_font_t& record, const formulon::FontRecord& normal) {
  formulon::FontRecord out = font_from_c(record);
  if (record.has_size_pt == 0) {
    out.size = normal.size;
  }
  if (record.name == nullptr) {
    out.name = normal.name;
  }
  out.has_size = true;
  return out;
}

/// A stated font size is `1..409` points, Excel's font-size box range.
fm_status_t check_font_size_pt(double size_pt) {
  if (size_pt < 1.0 || size_pt > 409.0) {
    return invalid_argument("size_pt must be in 1..409", "size_pt=" + formulon::fixed_string(size_pt, 6));
  }
  return 0;
}

/// `validate` of a cell font plus its size range.
fm_status_t validate_cell_font(const fm_font_t& record) {
  if (auto rc = validate(record); rc != 0) {
    return rc;
  }
  return record.has_size_pt != 0 ? check_font_size_pt(record.size_pt) : 0;
}

double finite_or_zero(double value) noexcept {
  return std::isfinite(value) ? value : 0.0;
}

formulon::FillRecord fill_from_c(const fm_fill_t& record) {
  formulon::FillRecord out;
  out.pattern = static_cast<std::uint8_t>(record.pattern);
  out.fg = color_from_c(record.foreground, &out.fg_argb);
  out.bg = color_from_c(record.background, &out.bg_argb);
  if (record.has_gradient != 0) {
    out.has_gradient = true;
    // The pattern fields do not exist on a gradient fill; clearing them keeps
    // records that differ only there from comparing unequal.
    out.pattern = 0;
    out.fg_argb = 0;
    out.bg_argb = 0;
    out.fg = formulon::ColorSpec{};
    out.bg = formulon::ColorSpec{};
    formulon::GradientFill& gradient = out.gradient;
    gradient.path = record.gradient_type == FM_GRADIENT_TYPE_PATH;
    gradient.degree = finite_or_zero(record.gradient_angle_deg);
    gradient.left = finite_or_zero(record.gradient_left);
    gradient.right = finite_or_zero(record.gradient_right);
    gradient.top = finite_or_zero(record.gradient_top);
    gradient.bottom = finite_or_zero(record.gradient_bottom);
    // `validate` has already refused a count past `kMaxGradientStops`.
    const std::size_t count = record.gradient_stops == nullptr ? 0U : record.gradient_stop_count;
    gradient.stops.reserve(count);
    for (std::size_t i = 0; i < count; ++i) {
      const fm_gradient_stop_t& in = record.gradient_stops[i];
      formulon::GradientStop stop;
      stop.position = finite_or_zero(in.position);
      stop.color = color_from_c(in.color, &stop.argb);
      gradient.stops.push_back(stop);
    }
  }
  return out;
}

formulon::BorderSide border_side_from_c(const fm_border_side_t& src) noexcept {
  formulon::BorderSide dst;
  dst.style = static_cast<std::uint8_t>(src.style);
  dst.color = color_from_c(src.color, &dst.color_argb);
  return dst;
}

formulon::BorderRecord border_from_c(const fm_border_t& record) noexcept {
  formulon::BorderRecord out;
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

extern "C" fm_status_t fm_workbook_add_font(fm_workbook_t* wb, const fm_font_t* record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = validate_cell_font(*record); rc != 0) {
    return rc;
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  formulon::FontRecord candidate = cell_font_from_c(*record, styles.fonts[0]);
  *out_index = find_or_append_record(styles.fonts, std::move(candidate), &font_key);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_font(fm_workbook_t* wb, uint32_t font_index, const fm_font_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  // No `ensure_default_style_roots` here: growing the table to satisfy an
  // out-of-range index would install a default-constructed font at every
  // slot below it, which is a different workbook than the caller asked for.
  if (auto rc = check_index(font_index, styles.fonts.size(), "font_index"); rc != 0) {
    return rc;
  }
  if (auto rc = validate_cell_font(*record); rc != 0) {
    return rc;
  }
  styles.fonts[font_index] = cell_font_from_c(*record, styles.fonts[0]);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_default_font(fm_workbook_t* wb, const fm_font_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  if (auto rc = validate_cell_font(*record); rc != 0) {
    return rc;
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  // Unlike `fm_workbook_set_font`, seeding is right here: index 0 is reserved
  // by construction, so a table without it is an empty table rather than one
  // whose slot the caller mis-numbered.
  ensure_default_style_roots(styles);
  styles.fonts[0] = cell_font_from_c(*record, styles.fonts[0]);
  return 0;
}

extern "C" fm_status_t fm_workbook_add_fill(fm_workbook_t* wb, const fm_fill_t* record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = validate(*record); rc != 0) {
    return rc;
  }
  formulon::FillRecord candidate = fill_from_c(*record);

  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  *out_index = find_or_append_record(styles.fills, std::move(candidate), &fill_key);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_fill(fm_workbook_t* wb, uint32_t fill_index, const fm_fill_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  if (auto rc = check_index(fill_index, styles.fills.size(), "fill_index"); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*record); rc != 0) {
    return rc;
  }
  styles.fills[fill_index] = fill_from_c(*record);
  return 0;
}

extern "C" fm_status_t fm_workbook_add_border(fm_workbook_t* wb, const fm_border_t* record, uint32_t* out_index) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = validate(*record); rc != 0) {
    return rc;
  }
  formulon::BorderRecord candidate = border_from_c(*record);

  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  *out_index = formulon::io::intern_border(styles, std::move(candidate));
  return 0;
}

extern "C" fm_status_t fm_workbook_set_border(fm_workbook_t* wb, uint32_t border_index, const fm_border_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  if (auto rc = check_index(border_index, styles.borders.size(), "border_index"); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*record); rc != 0) {
    return rc;
  }
  styles.borders[border_index] = border_from_c(*record);
  return 0;
}

extern "C" fm_status_t fm_workbook_add_number_format(fm_workbook_t* wb, const char* number_format,
                                                     uint32_t* out_num_fmt_id) {
  clear_last_error();
  if (wb == nullptr || out_num_fmt_id == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({number_format}); rc != 0) {
    return rc;
  }
  return add_num_fmt_code(wb->workbook(), text_or_empty(number_format), out_num_fmt_id);
}

namespace {

// Checks the table references of a cell-XF record against `styles`, and
// returns the engine record to store.
fm_status_t build_cell_xf(const formulon::StylesTable& styles, const fm_cell_xf_t& record, formulon::CellXf* out) {
  // Reject out-of-range references rather than auto-growing the parallel
  // tables; callers must register fonts / fills / borders beforehand.
  if (fm_status_t rc = validate_xf_references(styles, record); rc != 0) {
    return rc;
  }
  // A non-zero `xfId` must name an existing `<cellStyleXfs>` entry, or the
  // emitted attribute dangles. Zero stays legal with an empty table, which is
  // what a record that names no parent style looks like.
  if (record.style_xf_index != 0) {
    if (fm_status_t rc = check_index(record.style_xf_index, styles.cell_style_xfs.size(), "style_xf_index"); rc != 0) {
      return rc;
    }
  }

  // Confirm the read-side `fm_cell_xf_t` struct mirrors the engine
  // `formulon::CellXf` field-for-field; the write side reuses the
  // same shape so a layout drift would silently corrupt records.
  static_assert(sizeof(record.font_index) == sizeof(formulon::CellXf::font_index),
                "cell xf field font_index width must match the engine type");
  static_assert(sizeof(record.fill_index) == sizeof(formulon::CellXf::fill_index),
                "cell xf field fill_index width must match the engine type");
  static_assert(sizeof(record.border_index) == sizeof(formulon::CellXf::border_index),
                "cell xf field border_index width must match the engine type");

  *out = cell_xf_from_c(record, record.style_xf_index);
  return 0;
}

fm_status_t add_cell_xf_impl(fm_workbook_t* wb, const fm_cell_xf_t* record, uint32_t* out_xf_index) {
  if (wb == nullptr || record == nullptr || out_xf_index == nullptr) {
    return null_argument();
  }
  if (fm_status_t rc = validate_xf_alignment(*record); rc != 0) {
    return rc;
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  ensure_default_cell_xf(styles);
  formulon::CellXf candidate;
  if (fm_status_t rc = build_cell_xf(styles, *record, &candidate); rc != 0) {
    return rc;
  }
  *out_xf_index = formulon::intern_cell_xf(styles, std::move(candidate));
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_add_cell_xf(fm_workbook_t* wb, const fm_cell_xf_t* record, uint32_t* out_xf_index) {
  clear_last_error();
  return add_cell_xf_impl(wb, record, out_xf_index);
}

extern "C" fm_status_t fm_workbook_set_cell_xf(fm_workbook_t* wb, uint32_t xf_index, const fm_cell_xf_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  if (auto rc = check_index(xf_index, styles.cell_xfs.size(), "xf_index"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = validate_xf_alignment(*record); rc != 0) {
    return rc;
  }
  formulon::CellXf candidate;
  if (fm_status_t rc = build_cell_xf(styles, *record, &candidate); rc != 0) {
    return rc;
  }
  styles.cell_xfs[xf_index] = std::move(candidate);
  return 0;
}

extern "C" fm_status_t fm_workbook_add_cell_style_xf(fm_workbook_t* wb, const fm_cell_xf_t* record,
                                                     uint32_t* out_xf_index) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || out_xf_index == nullptr) {
    return null_argument();
  }
  if (fm_status_t rc = validate_xf_alignment(*record); rc != 0) {
    return rc;
  }
  auto& styles = wb->workbook().mutable_styles();
  ensure_default_style_roots(styles);
  if (fm_status_t rc = validate_xf_references(styles, *record); rc != 0) {
    return rc;
  }
  formulon::CellXf candidate = cell_xf_from_c(*record, 0U);
  *out_xf_index = find_or_append_record(styles.cell_style_xfs, std::move(candidate), &formulon::cell_xf_key);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_cell_style_xf(fm_workbook_t* wb, uint32_t cell_style_xf_index,
                                                     const fm_cell_xf_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  if (auto rc = check_index(cell_style_xf_index, styles.cell_style_xfs.size(), "cell_style_xf_index"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = validate_xf_alignment(*record); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = validate_xf_references(styles, *record); rc != 0) {
    return rc;
  }
  return finish_status(formulon::set_cell_style_xf(styles, cell_style_xf_index, cell_xf_from_c(*record, 0U)));
}

extern "C" fm_status_t fm_workbook_set_cell_style(fm_workbook_t* wb, const fm_cell_style_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || record->name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({record->name}); rc != 0) {
    return rc;
  }
  auto& styles = wb->workbook().mutable_styles();
  if (record->name[0] == '\0') {
    return invalid_argument("name must not be empty");
  }
  if (auto rc = check_index(record->style_xf_index, styles.cell_style_xfs.size(), "style_xf_index"); rc != 0) {
    return rc;
  }
  // A custom style carries no `builtinId` attribute, which the engine spells as the sentinel.
  const bool builtin = record->has_builtin_id != 0;
  if (builtin && record->builtin_id > formulon::CellStyleRecord::kBuiltinIdMax) {
    return invalid_argument(
        ("builtin_id must be in 0.." + std::to_string(formulon::CellStyleRecord::kBuiltinIdMax)).c_str(),
        "builtin_id=" + std::to_string(record->builtin_id));
  }
  const uint32_t builtin_id = builtin ? record->builtin_id : formulon::CellStyleRecord::kBuiltinIdNone;
  // Only RowLevel_n / ColLevel_n carry an outline level.
  const uint32_t max_level = (builtin_id == 1U || builtin_id == 2U) ? 6U : 0U;
  if (record->outline_level > max_level) {
    return invalid_argument(
        "outline_level out of range",
        "builtin_id=" + std::to_string(builtin_id) + " outline_level=" + std::to_string(record->outline_level));
  }
  formulon::CellStyleRecord style;
  style.name = record->name;
  style.xf_id = record->style_xf_index;
  style.builtin_id = builtin_id;
  style.i_level = record->outline_level;
  style.hidden = record->hidden != 0;
  style.custom_builtin = record->custom_builtin != 0;
  // An upsert: a custom style the name resolves to takes the record whole; a
  // built-in one keeps its identity and takes the rest.
  if (const std::size_t found = formulon::resolve_cell_style(styles, style.name, wb->workbook().excel_profile());
      found != formulon::sheet_names::kNotFound) {
    formulon::CellStyleRecord& existing = styles.cell_styles[found];
    if (existing.builtin_id == formulon::CellStyleRecord::kBuiltinIdNone) {
      existing = std::move(style);
      return 0;
    }
    if (style.builtin_id != formulon::CellStyleRecord::kBuiltinIdNone && style.builtin_id != existing.builtin_id) {
      return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "name already in use",
                               "name=" + style.name);
    }
    existing.name = std::move(style.name);
    existing.xf_id = style.xf_id;
    existing.hidden = style.hidden;
    existing.custom_builtin = style.custom_builtin;
    return 0;
  }
  if (std::any_of(styles.cell_styles.begin(), styles.cell_styles.end(), [&](const formulon::CellStyleRecord& s) {
        return formulon::sheet_names::write_collides(s.name, style.name);
      })) {
    return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "name already in use", "name=" + style.name);
  }
  styles.cell_styles.push_back(std::move(style));
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_cell_style(fm_workbook_t* wb, const char* name) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  return finish_status(
      formulon::remove_cell_style(wb->workbook().mutable_styles(), name, wb->workbook().excel_profile()));
}

extern "C" fm_status_t fm_workbook_add_batch(fm_workbook_t* wb, const fm_styles_batch_spec_t* batch) {
  clear_last_error();
  if (wb == nullptr || batch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*batch); rc != 0) {
    return rc;
  }
  const struct {
    std::size_t count;
    const void* data;
    const char* name;
    const void* indices;
    const char* indices_name;
  } arrays[] = {
      {batch->font_count, batch->fonts, "fonts", batch->font_indices, "font_indices"},
      {batch->fill_count, batch->fills, "fills", batch->fill_indices, "fill_indices"},
      {batch->border_count, batch->borders, "borders", batch->border_indices, "border_indices"},
      {batch->cell_xf_count, batch->cell_xfs, "cell_xfs", batch->cell_xf_indices, "cell_xf_indices"},
      {batch->num_fmt_count, batch->num_fmt_codes, "num_fmt_codes", batch->num_fmt_ids, "num_fmt_ids"},
  };
  for (const auto& array : arrays) {
    if (array.count != 0U && array.data == nullptr) {
      return null_field(array.name);
    }
    if (array.count != 0U && array.indices == nullptr) {
      return null_field(array.indices_name);
    }
  }
  for (size_t i = 0; i < batch->font_count; ++i) {
    if (auto rc = validate_cell_font(batch->fonts[i]); rc != 0) {
      return rc;
    }
  }
  for (size_t i = 0; i < batch->fill_count; ++i) {
    if (auto rc = validate(batch->fills[i]); rc != 0) {
      return rc;
    }
  }
  for (size_t i = 0; i < batch->border_count; ++i) {
    if (auto rc = validate(batch->borders[i]); rc != 0) {
      return rc;
    }
  }
  for (size_t i = 0; i < batch->num_fmt_count; ++i) {
    if (auto rc = check_utf8({batch->num_fmt_codes[i]}); rc != 0) {
      return rc;
    }
  }

  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  // Stage every mutation, including the default roots that are required for
  // a valid xf. StylesTable owns strings and all nested records by value, so
  // this is a deep copy. The live table and caller-owned output arrays remain
  // untouched until every input record has passed validation.
  formulon::StylesTable staged = styles;
  ensure_default_style_roots(staged);

  std::vector<uint32_t> staged_font_indices(batch->font_count);
  std::vector<uint32_t> staged_fill_indices(batch->fill_count);
  std::vector<uint32_t> staged_border_indices(batch->border_count);
  std::vector<uint32_t> staged_cell_xf_indices(batch->cell_xf_count);
  std::vector<uint32_t> staged_num_fmt_ids(batch->num_fmt_count);

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
    formulon::FontRecord value = cell_font_from_c(batch->fonts[i], staged.fonts[0]);
    const std::string key = font_key(value);
    const auto [it, inserted] = fonts.emplace(key, static_cast<uint32_t>(staged.fonts.size()));
    if (inserted)
      staged.fonts.push_back(std::move(value));
    staged_font_indices[i] = it->second;
  }
  for (size_t i = 0; i < batch->fill_count; ++i) {
    formulon::FillRecord value = fill_from_c(batch->fills[i]);
    const std::string key = fill_key(value);
    const auto [it, inserted] = fills.emplace(key, static_cast<uint32_t>(staged.fills.size()));
    if (inserted)
      staged.fills.push_back(value);
    staged_fill_indices[i] = it->second;
  }
  for (size_t i = 0; i < batch->border_count; ++i) {
    formulon::BorderRecord value = border_from_c(batch->borders[i]);
    const std::string key = border_key(value);
    const auto [it, inserted] = borders.emplace(key, static_cast<uint32_t>(staged.borders.size()));
    if (inserted)
      staged.borders.push_back(value);
    staged_border_indices[i] = it->second;
  }
  std::unordered_map<std::string, uint16_t> num_fmts;
  num_fmts.reserve(staged.num_fmts.size() + batch->num_fmt_count);
  std::uint32_t next_num_fmt_id = formulon::c_api::parts::kLastBuiltinNumFmtId;
  for (const formulon::NumFmtRecord& record : staged.num_fmts) {
    const std::optional<std::string_view> resolved = effective_num_fmt(staged, record.id);
    if (record.format_string_index < staged.num_fmt_strings.size() && resolved.has_value() &&
        *resolved == staged.num_fmt_strings[record.format_string_index]) {
      num_fmts.emplace(staged.num_fmt_strings[record.format_string_index], record.id);
    }
    if (record.id > next_num_fmt_id) {
      next_num_fmt_id = record.id;
    }
  }
  for (size_t i = 0; i < batch->num_fmt_count; ++i) {
    const std::string code = text_or_empty(batch->num_fmt_codes[i]);
    if (const std::optional<std::uint16_t> builtin =
            formulon::c_api::parts::builtin_num_fmt_id_for_code(staged, wb->workbook().excel_profile(), code)) {
      staged_num_fmt_ids[i] = *builtin;
      continue;
    }
    const auto existing = num_fmts.find(code);
    if (existing != num_fmts.end()) {
      staged_num_fmt_ids[i] = existing->second;
      continue;
    }
    if (next_num_fmt_id >= std::numeric_limits<std::uint16_t>::max()) {
      return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed, "custom numFmtId space exhausted",
                               "max_num_fmt_id=" + std::to_string(std::numeric_limits<std::uint16_t>::max()));
    }
    const std::uint32_t candidate_num_fmt_id = next_num_fmt_id + 1U;
    num_fmts.emplace(code, static_cast<std::uint16_t>(candidate_num_fmt_id));
    ++next_num_fmt_id;
    staged.num_fmt_strings.push_back(code);
    formulon::NumFmtRecord record;
    record.id = static_cast<std::uint16_t>(next_num_fmt_id);
    record.format_string_index = static_cast<uint32_t>(staged.num_fmt_strings.size() - 1U);
    staged.num_fmts.push_back(record);
    staged_num_fmt_ids[i] = record.id;
  }
  ensure_default_cell_xf(staged);
  std::unordered_map<std::string, uint32_t> xfs;
  xfs.reserve(staged.cell_xfs.size() + batch->cell_xf_count);
  for (uint32_t i = 0; i < staged.cell_xfs.size(); ++i)
    xfs.emplace(formulon::cell_xf_key(staged.cell_xfs[i]), i);
  for (size_t i = 0; i < batch->cell_xf_count; ++i) {
    const fm_cell_xf_t& record = batch->cell_xfs[i];
    if (fm_status_t rc = validate_xf_alignment(record); rc != 0) {
      return rc;
    }
    if (fm_status_t rc = validate_xf_references(staged, record); rc != 0) {
      return rc;
    }
    // The batch shares the single-record presence-flag contract; `xfId` is
    // staged as given so a batch may point at a named style added earlier,
    // and is rejected when it would dangle.
    if (record.style_xf_index != 0) {
      if (fm_status_t rc = check_index(record.style_xf_index, staged.cell_style_xfs.size(), "style_xf_index");
          rc != 0) {
        return rc;
      }
    }
    const formulon::CellXf value = cell_xf_from_c(record, record.style_xf_index);
    const std::string key = formulon::cell_xf_key(value);
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

namespace {

// Validates `record` and builds the engine record, registering its number
// format with the workbook.
fm_status_t build_dxf(formulon::Workbook& workbook, const fm_dxf_t* record, formulon::DifferentialFormat* out) {
  if (record->has_font != 0 && record->font.has_size_pt != 0) {
    if (fm_status_t rc = check_finite_non_negative(record->font.size_pt, "font.size_pt"); rc != 0) {
      return rc;
    }
    if (fm_status_t rc = check_font_size_pt(record->font.size_pt); rc != 0) {
      return rc;
    }
  }
  if (record->has_font != 0) {
    if (fm_status_t rc = check_font_enums(record->font); rc != 0) {
      return rc;
    }
  }
  if (record->has_fill != 0) {
    if (fm_status_t rc = validate(record->fill); rc != 0) {
      return rc;
    }
  }
  if (record->has_border != 0) {
    if (fm_status_t rc = validate(record->border); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_utf8({record->has_font != 0 ? record->font.name : nullptr, record->number_format,
                            record->alignment_xml, record->protection_xml});
      rc != 0) {
    return rc;
  }

  formulon::DifferentialFormat& candidate = *out;
  candidate.has_font = record->has_font != 0;
  if (candidate.has_font) {
    candidate.font = font_from_c(record->font);
  }
  candidate.has_fill = record->has_fill != 0;
  if (candidate.has_fill) {
    candidate.fill = fill_from_c(record->fill);
  }
  candidate.has_border = record->has_border != 0;
  if (candidate.has_border) {
    candidate.border = border_from_c(record->border);
  }
  candidate.has_num_fmt = record->number_format != nullptr;
  if (candidate.has_num_fmt) {
    candidate.num_fmt_code = record->number_format;
    std::uint32_t num_fmt_id = 0;
    if (fm_status_t rc = add_num_fmt_code(workbook, candidate.num_fmt_code, &num_fmt_id); rc != 0) {
      return rc;
    }
    candidate.num_fmt_id = static_cast<std::uint16_t>(num_fmt_id);
  }
  candidate.alignment_xml = text_or_empty(record->alignment_xml);
  candidate.protection_xml = text_or_empty(record->protection_xml);

  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_add_dxf(fm_workbook_t* wb, const fm_dxf_t* record, uint32_t* out_dxf_index) {
  clear_last_error();
  if (wb == nullptr || record == nullptr || out_dxf_index == nullptr) {
    return null_argument();
  }
  formulon::DifferentialFormat candidate;
  if (fm_status_t rc = build_dxf(wb->workbook(), record, &candidate); rc != 0) {
    return rc;
  }
  formulon::StylesTable& styles = wb->workbook().mutable_styles();
  for (std::size_t i = 0; i < styles.dxfs.size(); ++i) {
    if (dxf_records_equal(styles.dxfs[i], candidate)) {
      *out_dxf_index = static_cast<uint32_t>(i);
      return 0;
    }
  }
  styles.dxfs.push_back(std::move(candidate));
  *out_dxf_index = static_cast<uint32_t>(styles.dxfs.size() - 1U);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_dxf(fm_workbook_t* wb, uint32_t dxf_index, const fm_dxf_t* record) {
  clear_last_error();
  if (wb == nullptr || record == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(dxf_index, wb->workbook().styles().dxfs.size(), "dxf_index"); rc != 0) {
    return rc;
  }
  formulon::DifferentialFormat candidate;
  if (fm_status_t rc = build_dxf(wb->workbook(), record, &candidate); rc != 0) {
    return rc;
  }
  wb->workbook().mutable_styles().dxfs[dxf_index] = std::move(candidate);
  return 0;
}
