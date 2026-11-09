
#include "c_api/parts/common.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <initializer_list>
#include <limits>
#include <new>
#include <optional>
#include <string>
#include <string_view>
#include <type_traits>
#include <utility>

#include "c_api/binding_status.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/enum_map.h"
#include "cf/cf_types.h"
#include "drawing/drawing_xml.h"
#include "eval/builtin_names.h"
#include "eval/coerce.h"
#include "eval/formula_entry_r1c1.h"
#include "eval/function_registry.h"
#include "num_fmt_locale.h"
#include "parser/formula_prefix.h"
#include "parser/literal_spelling.h"
#include "parser/parse_status.h"
#include "parser/reference.h"
#include "parser/structured_ref_spelling.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "utils/number_text.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook.h"

// The binding layers mirror these codes without including utils/error.h.
static_assert(formulon::c_api::kBindingInvalidHandle ==
              static_cast<std::int32_t>(formulon::FormulonErrorCode::kBindingInvalidHandle));
static_assert(formulon::c_api::kBindingCallbackException ==
              static_cast<std::int32_t>(formulon::FormulonErrorCode::kBindingCallbackException));
static_assert(formulon::c_api::kInvalidArgument ==
              static_cast<std::int32_t>(formulon::FormulonErrorCode::kInvalidArgument));

namespace formulon {
namespace c_api {
namespace parts {

// ABI identity: the `FM_VAL_*` discriminators are numerically equal to the
// `formulon::ValueKind` ordinals. `value_to_fm` below maps each ValueKind
// to its `FM_VAL_*` by name; these static_asserts pin the parity so a
// reorder of either enum is a build break rather than a silent wire-format
// drift across the C ABI.
static_assert(FM_VALUE_KIND_BLANK == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Blank),
              "FM_VALUE_KIND_BLANK must match ValueKind::Blank");
static_assert(FM_VALUE_KIND_NUMBER == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Number),
              "FM_VALUE_KIND_NUMBER must match ValueKind::Number");
static_assert(FM_VALUE_KIND_BOOL == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Bool),
              "FM_VALUE_KIND_BOOL must match ValueKind::Bool");
static_assert(FM_VALUE_KIND_TEXT == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Text),
              "FM_VALUE_KIND_TEXT must match ValueKind::Text");
static_assert(FM_VALUE_KIND_ERROR == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Error),
              "FM_VALUE_KIND_ERROR must match ValueKind::Error");
static_assert(FM_VALUE_KIND_ARRAY == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Array),
              "FM_VALUE_KIND_ARRAY must match ValueKind::Array");
static_assert(FM_VALUE_KIND_REF == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Ref),
              "FM_VALUE_KIND_REF must match ValueKind::Ref");
static_assert(FM_VALUE_KIND_LAMBDA == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Lambda),
              "FM_VALUE_KIND_LAMBDA must match ValueKind::Lambda");

namespace {

// Thread-local diagnostics for the most recent C API call on this thread.
//
// Both buffers are always non-empty in the sense that they own valid string
// storage, but their `c_str()` may point to an empty string. Returning
// `c_str()` therefore satisfies the public "never NULL" contract.
thread_local std::string g_last_error_message;
thread_local std::string g_last_error_context;

// Position of the current call's formula syntax failure, if it had one.
thread_local ParseErrorRecord g_last_parse_error;
thread_local bool g_has_last_parse_error = false;

}  // namespace

const char* store_cstr(TextStore& store, std::string_view text) {
  store.emplace_back(text.data(), text.size());
  return store.back().c_str();
}

void copy_owned_bytes(const std::vector<std::uint8_t>& bytes, std::uint8_t** out_bytes, std::size_t* out_len) {
  auto* buffer = new std::uint8_t[bytes.size()];
  if (!bytes.empty()) {
    std::memcpy(buffer, bytes.data(), bytes.size());
  }
  *out_bytes = buffer;
  *out_len = bytes.size();
}

void clear_last_error() {
  g_last_error_message.clear();
  g_last_error_context.clear();
  g_has_last_parse_error = false;
}

fm_status_t set_last_error(const formulon::Error& err) {
  g_last_error_message = err.message;
  g_last_error_context = err.context;
  return static_cast<fm_status_t>(err.code);
}

fm_status_t finish_status(const formulon::Expected<void, formulon::Error>& r) {
  return r ? 0 : set_last_error(r.error());
}

fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context) {
  formulon::Error err;
  err.code = code;
  err.message = message != nullptr ? message : "";
  err.context = std::move(context);
  return set_last_error(err);
}

fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message) {
  return set_binding_error(code, message, std::string());
}

const char* last_error_message() {
  return g_last_error_message.c_str();
}

const char* last_error_context() {
  return g_last_error_context.c_str();
}

const ParseErrorRecord* last_parse_error() {
  return g_has_last_parse_error ? &g_last_parse_error : nullptr;
}

fm_status_t formula_parse_failed(const formulon::parser::ParseFailure& failure, std::string context) {
  g_last_parse_error.code = static_cast<std::int32_t>(failure.code);
  g_last_parse_error.start = failure.range.start;
  g_last_parse_error.end = failure.range.end;
  g_last_parse_error.line = failure.range.line;
  g_last_parse_error.column = failure.range.column;
  g_has_last_parse_error = true;
  formulon::Error err;
  err.code = formulon::parser::parse_error_status(failure.code);
  err.message = std::string(failure.message);
  err.context = std::move(context);
  return set_last_error(err);
}

fm_status_t check_struct_size(std::uint32_t got, std::size_t v1_size, const char* type_name) {
  if (got >= v1_size) {
    return 0;
  }
  return invalid_argument("struct_size is smaller than the record",
                          "struct_size=" + std::to_string(got) + " type=" + std::string(type_name));
}

fm_status_t ensure_calculated(const fm_workbook_t* wb) {
  return finish_status(
      const_cast<fm_workbook_t*>(wb)->workbook().ensure_calculated(formulon::eval::default_registry()));
}

// These C enums carry the engine ordinals unchanged; pin the last member of each.
// The comparison operator and the date period cross through `enum_map.h`.
static_assert(FM_COLOR_KIND_AUTO == static_cast<int>(formulon::ColorSpec::Kind::kAuto), "fm_color_kind_t drift");
static_assert(FM_CF_RULE_TYPE_UNIQUE_VALUES == static_cast<int>(formulon::cf::RuleType::UniqueValues),
              "fm_cf_rule_type_t drift");
static_assert(FM_CF_VALUE_OBJECT_TYPE_AUTO_MAX == static_cast<int>(formulon::cf::CfvoType::AutoMax),
              "fm_cf_value_object_type_t drift");
static_assert(FM_ICON_SET_NAME_5_QUARTERS == static_cast<int>(formulon::cf::IconSetName::Five_Quarters),
              "fm_icon_set_name_t drift");

fm_color_t color_spec_to_c(const formulon::ColorSpec& src, std::uint32_t fallback_argb) {
  fm_color_t out{};
  if (src.kind == formulon::ColorSpec::Kind::kNone) {
    if (fallback_argb != 0U) {
      out.kind = FM_COLOR_KIND_RGB;
      out.argb = fallback_argb;
    }
    return out;
  }
  out.tint = src.tint;
  out.argb = src.rgb;
  out.theme = src.theme;
  out.indexed = src.indexed;
  out.kind = static_cast<std::uint8_t>(src.kind);
  return out;
}

formulon::ColorSpec color_spec_from_c(const fm_color_t& src, std::uint32_t* out_argb) {
  if (out_argb != nullptr) {
    *out_argb = src.kind == FM_COLOR_KIND_RGB ? src.argb : 0U;
  }
  formulon::ColorSpec out;
  if (src.kind == FM_COLOR_KIND_NONE) {
    return out;
  }
  out.kind = static_cast<formulon::ColorSpec::Kind>(src.kind);
  out.rgb = src.argb;
  out.theme = src.theme;
  out.tint = src.tint;
  out.indexed = src.indexed;
  return out;
}

namespace {

fm_status_t check_range_count(std::size_t n) {
  if (n > kMaxRangesPerCApiCall) {
    return invalid_argument("range_count exceeds per-call cap",
                            "range_count=" + std::to_string(n) + " cap=" + std::to_string(kMaxRangesPerCApiCall));
  }
  return 0;
}

}  // namespace

fm_status_t check_profile_id(const char* profile_id, formulon::ExcelProfile* out) {
  // The id is echoed into the diagnostic context, which must stay UTF-8.
  if (auto rc = check_utf8(profile_id, "field=profile_id"); rc != 0) {
    return rc;
  }
  if (formulon::parse_excel_profile_id(profile_id, out)) {
    return 0;
  }
  return invalid_argument("unknown profile", std::string("profile_id=") + profile_id);
}

fm_status_t check_sheet_rect(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                             std::uint32_t last_col) {
  if (formulon::Sheet::rect_in_grid(first_row, first_col, last_row, last_col)) {
    return 0;
  }
  return invalid_argument("cell rectangle out of range",
                          "first_row=" + std::to_string(first_row) + " first_col=" + std::to_string(first_col) +
                              " last_row=" + std::to_string(last_row) + " last_col=" + std::to_string(last_col));
}

fm_status_t check_sheet_rect(const fm_cell_rect_t& rect) {
  return check_sheet_rect(rect.first_row, rect.first_col, rect.last_row, rect.last_col);
}

fm_status_t check_rect_list(const fm_cell_rect_t* ranges, std::size_t range_count) {
  if (range_count == 0U) {
    return invalid_argument("range_count must be >= 1", "range_count=0");
  }
  if (ranges == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "ranges is NULL while range_count > 0");
  }
  if (auto rc = check_range_count(range_count); rc != 0) {
    return rc;
  }
  for (std::size_t i = 0; i < range_count; ++i) {
    if (auto rc = check_sheet_rect(ranges[i]); rc != 0) {
      return rc;
    }
  }
  return 0;
}

const fm_cell_rect_t* publish_rects(const fm_workbook_t* wb, const std::vector<formulon::MergeRange>& ranges) {
  auto& scratch = const_cast<formulon::c_api::BorrowedArrayArena<fm_cell_rect_t>&>(wb->rect_scratch);
  scratch.clear();
  std::vector<fm_cell_rect_t> block;
  block.reserve(ranges.size());
  for (const formulon::MergeRange& r : ranges) {
    block.push_back(to_cell_rect(r));
  }
  return scratch.adopt(std::move(block));
}

namespace {

// One endpoint of a whole-column / whole-row pair: the text before the `:`
// `format_a1` writes between the two copies of the axis.
std::string whole_axis_endpoint(const formulon::parser::Reference& ref) {
  std::string text = formulon::parser::format_a1(ref);
  text.resize(text.rfind(':'));
  return text;
}

}  // namespace

std::string format_rect_a1(std::string_view sheet_name, const fm_cell_rect_t& rect, bool absolute) {
  formulon::parser::Reference first;
  first.sheet = sheet_name;
  first.row = rect.first_row;
  first.col = rect.first_col;
  first.row_abs = absolute;
  first.col_abs = absolute;
  formulon::parser::Reference second = first;
  second.sheet = std::string_view();
  second.row = rect.last_row;
  second.col = rect.last_col;
  if (rect.first_row == 0U && rect.last_row == formulon::Sheet::kMaxRows - 1U) {
    first.is_full_col = true;
    second.is_full_col = true;
    return whole_axis_endpoint(first) + ":" + whole_axis_endpoint(second);
  }
  if (rect.first_col == 0U && rect.last_col == formulon::Sheet::kMaxCols - 1U) {
    first.is_full_row = true;
    second.is_full_row = true;
    return whole_axis_endpoint(first) + ":" + whole_axis_endpoint(second);
  }
  if (rect.first_row == rect.last_row && rect.first_col == rect.last_col) {
    return formulon::parser::format_a1(first);
  }
  return formulon::parser::format_a1(first) + ":" + formulon::parser::format_a1(second);
}

fm_status_t check_column_span(std::uint32_t first, std::uint32_t last) {
  if (first <= last && last < formulon::Sheet::kMaxCols) {
    return 0;
  }
  return invalid_argument("column span out of range",
                          "first=" + std::to_string(first) + " last=" + std::to_string(last) +
                              " last_col=" + std::to_string(formulon::Sheet::kMaxCols - 1U));
}

fm_status_t check_row_span(std::uint32_t first, std::uint32_t last) {
  if (first <= last && last < formulon::Sheet::kMaxRows) {
    return 0;
  }
  return invalid_argument("row span out of range", "first=" + std::to_string(first) + " last=" + std::to_string(last) +
                                                       " last_row=" + std::to_string(formulon::Sheet::kMaxRows - 1U));
}

fm_status_t check_row_index(std::uint32_t row) {
  if (row < formulon::Sheet::kMaxRows) {
    return 0;
  }
  return invalid_argument("row index out of range",
                          "row=" + std::to_string(row) + " last_row=" + std::to_string(formulon::Sheet::kMaxRows - 1U));
}

fm_status_t check_column_index(std::uint32_t col) {
  if (col < formulon::Sheet::kMaxCols) {
    return 0;
  }
  return invalid_argument("column index out of range",
                          "col=" + std::to_string(col) + " last_col=" + std::to_string(formulon::Sheet::kMaxCols - 1U));
}

fm_status_t check_finite_non_negative(double value, const char* field) {
  if (std::isfinite(value) && value >= 0.0) {
    return 0;
  }
  return invalid_argument("value must be finite and non-negative",
                          std::string(field) + "=" + formulon::fixed_string(value, 6));
}

fm_status_t check_finite(double value, const char* field) {
  if (std::isfinite(value)) {
    return 0;
  }
  return invalid_argument("value must be finite", std::string(field) + "=" + formulon::fixed_string(value, 6));
}

fm_status_t value_from_fm(const fm_value_t& in, formulon::Value* out) {
  switch (in.kind) {
    case FM_VALUE_KIND_BLANK:
      *out = formulon::Value::blank();
      return 0;
    case FM_VALUE_KIND_NUMBER:
      if (auto rc = check_finite(in.u.number, "value"); rc != 0) {
        return rc;
      }
      *out = formulon::Value::number(in.u.number);
      return 0;
    case FM_VALUE_KIND_BOOL:
      *out = formulon::Value::boolean(in.u.boolean != 0);
      return 0;
    case FM_VALUE_KIND_TEXT:
      if (in.u.text == nullptr) {
        return null_argument("value text");
      }
      if (auto rc = check_utf8(in.u.text, "field=value"); rc != 0) {
        return rc;
      }
      *out = formulon::Value::text(in.u.text);
      return 0;
    case FM_VALUE_KIND_ERROR:
      if (!is_valid_error_code(in.u.error_code)) {
        return invalid_error_code(in.u.error_code);
      }
      *out = formulon::Value::error(static_cast<formulon::ErrorCode>(in.u.error_code));
      return 0;
    default:
      return invalid_argument("value kind not accepted", "kind=" + std::to_string(static_cast<int>(in.kind)));
  }
}

fm_status_t check_enum_domain(std::int64_t value, std::int64_t min, std::int64_t max, const char* field) {
  if (value >= min && value <= max) {
    return 0;
  }
  return invalid_argument("enum ordinal out of range", std::string(field) + "=" + std::to_string(value) + " min=" +
                                                           std::to_string(min) + " max=" + std::to_string(max));
}

namespace {

// Rejects a string that is not GUID-shaped (`{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`,
// hexadecimal digits). NULL or empty passes: the entry point synthesizes an id.
// `ST_Guid` attributes holding anything else make Excel drop the whole block.
fm_status_t check_guid(const char* text, const char* field) {
  if (text == nullptr || text[0] == '\0') {
    return 0;
  }
  // `{8-4-4-4-12}`: 32 hexadecimal digits, four hyphens, two braces.
  static constexpr std::string_view kShape = "{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}";
  const std::string_view value(text);
  bool shaped = value.size() == kShape.size();
  for (std::size_t i = 0; shaped && i < kShape.size(); ++i) {
    switch (kShape[i]) {
      case 'X':
        shaped = formulon::strings::hex_digit_value(value[i]) >= 0;
        break;
      default:
        shaped = value[i] == kShape[i];
        break;
    }
  }
  if (shaped) {
    return 0;
  }
  return invalid_argument("value is not a GUID", std::string(field) + "=" + std::string(value));
}

}  // namespace

namespace {

// Runs `check_enum_domain` over every `fm_cfvo_t` in a threshold array.
// The array bound is the caller's, so an entry point that has not yet
// established `count` against its pointer must not reach here.
fm_status_t check_cfvo_types(const fm_cfvo_t* thresholds, std::uint32_t count, const char* field) {
  for (std::uint32_t i = 0; i < count; ++i) {
    if (auto rc = check_enum_domain(thresholds[i].type, 0, static_cast<std::int64_t>(cf::CfvoType::AutoMax), field);
        rc != 0) {
      return rc;
    }
  }
  return 0;
}

}  // namespace

fm_status_t validate(const fm_cf_rule_t& rule) {
  // A range the A1 encoder cannot spell reaches the writer as an empty
  // reference, which Excel reads as a broken conditional-formatting
  // block. Rejecting it here names the offending rectangle at set time
  // instead of losing the format at save time.
  if (auto rc = check_rect_list(rule.ranges, rule.range_count); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(rule.type, 0, static_cast<std::int64_t>(cf::RuleType::UniqueValues), "type");
      rc != 0) {
    return rc;
  }
  if (rule.has_op != 0) {
    if (auto rc = check_enum_domain(rule.op, 0, FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, "op"); rc != 0) {
      return rc;
    }
  }
  if (rule.has_time_period != 0) {
    if (auto rc = check_enum_domain(rule.time_period, 0, FM_DATE_PERIOD_NEXT_WEEK, "time_period"); rc != 0) {
      return rc;
    }
  }
  if (rule.has_data_bar != 0) {
    if (auto rc = check_cfvo_types(&rule.data_bar_min, 1U, "data_bar_min.type"); rc != 0) {
      return rc;
    }
    if (auto rc = check_cfvo_types(&rule.data_bar_max, 1U, "data_bar_max.type"); rc != 0) {
      return rc;
    }
    if (rule.has_data_bar_axis_position != 0) {
      if (auto rc = check_enum_domain(rule.data_bar_axis_position, 0, FM_CF_DATA_BAR_AXIS_POSITION_NONE,
                                      "data_bar_axis_position");
          rc != 0) {
        return rc;
      }
    }
    if (auto rc =
            check_enum_domain(rule.data_bar_direction, 0, FM_CF_DATA_BAR_DIRECTION_RIGHT_TO_LEFT, "data_bar_direction");
        rc != 0) {
      return rc;
    }
  }
  // The colour-scale scan is bounded by the array shape the entry point
  // accepts, so a record that declares more thresholds than the schema
  // allows is rejected there rather than scanned here.
  if (rule.color_scale_thresholds != nullptr && rule.color_scale_count >= 2U && rule.color_scale_count <= 3U) {
    if (auto rc =
            check_cfvo_types(rule.color_scale_thresholds, rule.color_scale_count, "color_scale_thresholds[].type");
        rc != 0) {
      return rc;
    }
  }
  if (rule.has_icon_set != 0) {
    if (auto rc = check_enum_domain(rule.icon_set_name, 0, static_cast<std::int64_t>(cf::IconSetName::Five_Quarters),
                                    "icon_set_name");
        rc != 0) {
      return rc;
    }
    if (rule.icon_set_thresholds != nullptr) {
      if (auto rc =
              check_cfvo_types(rule.icon_set_thresholds, rule.icon_set_threshold_count, "icon_set_thresholds[].type");
          rc != 0) {
        return rc;
      }
    }
    if (rule.has_icon_set_floor != 0) {
      if (auto rc = check_cfvo_types(&rule.icon_set_floor, 1U, "icon_set_floor.type"); rc != 0) {
        return rc;
      }
    }
    if (rule.icon_set_icon_count != 0U) {
      // One icon per level, each naming an icon its source set has.
      const std::size_t levels = cf::icon_count(static_cast<cf::IconSetName>(rule.icon_set_name));
      if (rule.icon_set_icons == nullptr || rule.icon_set_icon_count != levels) {
        return invalid_argument(
            "icon_set_icons must hold one icon per level",
            "icon_set_icons count=" + std::to_string(rule.icon_set_icon_count) + " levels=" + std::to_string(levels));
      }
      for (std::uint32_t i = 0; i < rule.icon_set_icon_count; ++i) {
        const fm_cf_icon_t& icon = rule.icon_set_icons[i];
        const std::size_t available = icon.icon_set < 0 || icon.icon_set > FM_ICON_SET_NAME_NO_ICONS
                                          ? 0U
                                          : cf::cf_icon_set_icon_count(token_of(kIconSetNameTokens, icon.icon_set));
        if (available == 0U || icon.icon_index < 0 || static_cast<std::size_t>(icon.icon_index) >= available) {
          return invalid_argument("icon_set_icons entry names no icon", "icon_set_icons[" + std::to_string(i) + "]");
        }
      }
    }
  }
  if (auto rc = check_utf8({rule.text}); rc != 0) {
    return rc;
  }
  // The id reaches the file as `<x14:cfRule id="...">` and `<x14:id>`,
  // both of schema type `ST_Guid`. Excel repairs a workbook whose
  // extension block holds anything else, and repairing means dropping
  // every conditional format in it.
  return check_guid(rule.id, "id");
}

fm_status_t validate(const fm_data_validation_t& v) {
  if (auto rc = check_rect_list(v.ranges, v.range_count); rc != 0) {
    return rc;
  }
  // The three domains the writer's attribute tables spell out
  // (`sheet_xml_builder.cpp`): a value past them is emitted as an absent
  // attribute, so a mistyped `type` saves as "no validation at all".
  if (auto rc = check_enum_domain(v.type, 0, 7, "type"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.op, 0, FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, "op"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.error_style, 0, 2, "error_style"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.allow_blank, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "allow_blank"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.show_input_message, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_input_message");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.show_error_message, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_error_message");
      rc != 0) {
    return rc;
  }
  if (auto rc =
          check_enum_domain(v.show_in_cell_dropdown, FM_TRISTATE_DEFAULT, FM_TRISTATE_TRUE, "show_in_cell_dropdown");
      rc != 0) {
    return rc;
  }
  return check_utf8({v.error_title, v.error_message, v.prompt_title, v.prompt_message});
}

fm_status_t validate(const fm_font_t& font) {
  if (fm_status_t rc = check_finite_non_negative(font.size_pt, "font.size_pt"); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_font_enums(font); rc != 0) {
    return rc;
  }
  return check_utf8({font.name});
}

fm_status_t check_font_enums(const fm_font_t& font) {
  if (fm_status_t rc = check_enum_domain(font.underline, 0, FM_UNDERLINE_STYLE_DOUBLE_ACCOUNTING, "font.underline");
      rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(font.vert_align, 0, FM_FONT_VERTICAL_ALIGN_SUBSCRIPT, "font.vert_align");
      rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_enum_domain(font.family, 0, FM_FONT_FAMILY_DECORATIVE, "font.family"); rc != 0) {
    return rc;
  }
  return check_enum_domain(font.scheme, 0, FM_FONT_SCHEME_MINOR, "font.scheme");
}

fm_status_t validate(const fm_fill_t& fill) {
  if (fm_status_t rc = check_enum_domain(fill.pattern, 0, FM_FILL_PATTERN_GRAY0625, "fill.pattern"); rc != 0) {
    return rc;
  }
  return check_enum_domain(fill.gradient_type, 0, FM_GRADIENT_TYPE_PATH, "fill.gradient_type");
}

fm_status_t validate(const fm_border_t& border) {
  const std::pair<const fm_border_side_t*, const char*> sides[] = {{&border.left, "border.left.style"},
                                                                   {&border.right, "border.right.style"},
                                                                   {&border.top, "border.top.style"},
                                                                   {&border.bottom, "border.bottom.style"},
                                                                   {&border.diagonal, "border.diagonal.style"}};
  for (const auto& [side, field] : sides) {
    if (fm_status_t rc = check_enum_domain(side->style, 0, FM_BORDER_STYLE_SLANT_DASH_DOT, field); rc != 0) {
      return rc;
    }
  }
  return 0;
}

bool num_fmt_id_known(const formulon::StylesTable& styles, std::uint16_t id) {
  if (formulon::is_builtin_num_fmt_id(id)) {
    return true;
  }
  for (const formulon::NumFmtRecord& record : styles.num_fmts) {
    if (record.id == id) {
      return true;
    }
  }
  return false;
}

std::optional<std::uint16_t> builtin_num_fmt_id_for_code(const formulon::StylesTable& styles,
                                                         formulon::ExcelProfile profile, std::string_view code) {
  for (std::uint16_t id = 0; id < kBuiltinNumFmtIdLimit; ++id) {
    const std::optional<std::string_view> rendered = effective_num_fmt(styles, id, profile);
    if (rendered.has_value() && *rendered == code) {
      return id;
    }
  }
  return std::nullopt;
}

fm_status_t add_num_fmt_code(formulon::Workbook& wb, const std::string& code, std::uint32_t* out_id) {
  // Step 1: a built-in slot whose profile-rendered code is `code`. A custom
  // record may override the built-in.
  if (const std::optional<std::uint16_t> builtin = builtin_num_fmt_id_for_code(wb.styles(), wb.excel_profile(), code)) {
    *out_id = *builtin;
    return 0;
  }

  formulon::StylesTable& styles = wb.mutable_styles();

  // Step 2: existing custom entry whose id still resolves to its code. This
  // skips shadowed duplicate ids and invalid string indexes.
  for (const formulon::NumFmtRecord& n : styles.num_fmts) {
    const std::optional<std::string_view> resolved = effective_num_fmt(styles, n.id);
    if (resolved.has_value() && code == *resolved) {
      *out_id = n.id;
      return 0;
    }
  }

  // Step 3: append a new custom entry. New id is one past the largest
  // existing custom id, with `kLastBuiltinNumFmtId` as the lower bound (the last
  // built-in slot).
  std::uint32_t next_id = kLastBuiltinNumFmtId;
  for (const formulon::NumFmtRecord& n : styles.num_fmts) {
    if (n.id > next_id) {
      next_id = n.id;
    }
  }
  if (next_id >= std::numeric_limits<std::uint16_t>::max()) {
    return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed, "custom numFmtId space exhausted",
                             "max_num_fmt_id=" + std::to_string(std::numeric_limits<std::uint16_t>::max()));
  }
  ++next_id;

  styles.num_fmt_strings.push_back(code);
  formulon::NumFmtRecord rec;
  rec.id = static_cast<std::uint16_t>(next_id);
  rec.format_string_index = static_cast<std::uint32_t>(styles.num_fmt_strings.size() - 1);
  styles.num_fmts.push_back(rec);
  *out_id = static_cast<std::uint16_t>(next_id);
  return 0;
}

fm_status_t pivot_num_fmt_id_text(formulon::Workbook& wb, const char* code, std::string* out) {
  out->clear();
  if (code == nullptr || code[0] == '\0') {
    return 0;
  }
  if (auto rc = check_utf8(code, "field=number_format"); rc != 0) {
    return rc;
  }
  const std::string_view text(code);
  std::uint32_t id = 0;
  if (auto rc = add_num_fmt_code(wb, std::string(text), &id); rc != 0) {
    return rc;
  }
  *out = std::to_string(id);
  return 0;
}

std::optional<std::string_view> resolve_num_fmt_code(const formulon::Workbook& wb, std::uint16_t id) {
  return effective_num_fmt(wb.styles(), id, wb.excel_profile());
}

const char* pivot_num_fmt_code(const formulon::Workbook& wb, std::string_view id_text) {
  if (id_text.empty()) {
    return nullptr;
  }
  std::uint32_t id = 0;
  for (const char c : id_text) {
    if (c < '0' || c > '9') {
      return nullptr;
    }
    id = id * 10U + static_cast<std::uint32_t>(c - '0');
    if (id > std::numeric_limits<std::uint16_t>::max()) {
      return nullptr;
    }
  }
  const std::optional<std::string_view> resolved = resolve_num_fmt_code(wb, static_cast<std::uint16_t>(id));
  return resolved.has_value() ? resolved->data() : nullptr;
}

fm_status_t validate(const fm_pivot_field_spec_t& spec) {
  if (spec.source_name == nullptr) {
    return null_argument("spec->source_name");
  }
  if (auto rc = check_utf8({spec.source_name, spec.custom_name}); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.axis, FM_PIVOT_AXIS_ABSENT, FM_PIVOT_AXIS_PAGE, "axis");
}

fm_status_t validate(const fm_pivot_data_field_t& spec) {
  if (auto rc = check_utf8({spec.name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec.aggregation, 0, FM_PIVOT_AGGREGATION_VARP, "aggregation"); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.show_as, 0, FM_PIVOT_SHOW_VALUES_AS_RANK_DESCENDING, "show_as");
}

fm_status_t validate(const fm_pivot_filter_t& spec) {
  if (auto rc = check_utf8({spec.value_text, spec.value_high_text}); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.type, 0, FM_PIVOT_FILTER_KIND_M12, "type");
}

void value_to_fm(const formulon::Value& v, TextStore& store, fm_value_t* out) {
  switch (v.kind()) {
    case formulon::ValueKind::Blank:
      out->kind = FM_VALUE_KIND_BLANK;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Number:
      out->kind = FM_VALUE_KIND_NUMBER;
      out->u.number = v.as_number();
      return;
    case formulon::ValueKind::Bool:
      out->kind = FM_VALUE_KIND_BOOL;
      out->u.boolean = v.as_boolean() ? 1 : 0;
      return;
    case formulon::ValueKind::Text: {
      const std::string_view text = v.as_text();
      out->kind = FM_VALUE_KIND_TEXT;
      out->u.text = store_cstr(store, text);
      return;
    }
    case formulon::ValueKind::Error:
      out->kind = FM_VALUE_KIND_ERROR;
      // The C-ABI wire format for errors is the raw `ErrorCode` ordinal
      // (Null=0, Div0=1, Value=2, ...), NOT the OOXML wire code. This is
      // the single source of truth for the boundary; see the
      // `fm_value_t::u.error_code` contract in `formulon_c.h`.
      out->u.error_code = static_cast<int32_t>(v.as_error());
      return;
    case formulon::ValueKind::Array:
      // Array passthrough is reserved; the kind is reported but no
      // payload is exposed across the boundary in this bundle.
      out->kind = FM_VALUE_KIND_ARRAY;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Ref:
      out->kind = FM_VALUE_KIND_REF;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Lambda:
      out->kind = FM_VALUE_KIND_LAMBDA;
      out->u.number = 0.0;
      return;
  }
  // Defensive default: surface as Blank rather than leaving uninitialised
  // bytes on the boundary. The switch above is exhaustive over every
  // `ValueKind` enumerator, so this branch is unreachable in practice.
  out->kind = FM_VALUE_KIND_BLANK;
  out->u.number = 0.0;
}

fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index) {
  if (wb == nullptr) {
    return null_argument("wb");
  }
  const std::size_t count = wb->workbook().sheet_count();
  if (sheet_index >= count) {
    return set_binding_error(formulon::FormulonErrorCode::kSheetIndexOutOfRange,
                             ("sheet index " + std::to_string(sheet_index) + " is out of range (" +
                              std::to_string(count) + (count == 1U ? " sheet)" : " sheets)"))
                                 .c_str(),
                             "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(count));
  }
  return 0;
}

TextStore& fresh_read_scratch(const fm_workbook_t* wb) {
  TextStore& store = const_cast<TextStore&>(wb->read_scratch);
  store.clear();
  return store;
}

const char* publish_scratch(const fm_workbook_t* wb, std::string text) {
  TextStore& store = fresh_read_scratch(wb);
  store.emplace_back(std::move(text));
  return store.back().c_str();
}

fm_status_t check_cell_coord(std::uint32_t row, std::uint32_t col) {
  if (formulon::Sheet::coord_in_grid(row, col)) {
    return 0;
  }
  return invalid_argument("cell coordinate out of range", "row=" + std::to_string(row) + " col=" + std::to_string(col));
}

fm_status_t check_index(std::size_t index, std::size_t size, const char* field) {
  if (index < size) {
    return 0;
  }
  return set_binding_error(
      formulon::FormulonErrorCode::kIndexOutOfRange,
      (std::string(field) + " " + std::to_string(index) + " is out of range (" + std::to_string(size) + " items)")
          .c_str(),
      std::string(field) + "=" + std::to_string(index) + " count=" + std::to_string(size));
}

fm_status_t check_limit(std::size_t value, std::size_t bound, const char* field) {
  if (value < bound) {
    return 0;
  }
  return invalid_argument((std::string(field) + " " + std::to_string(value) + " is out of range (maximum " +
                           std::to_string(bound - 1U) + ")")
                              .c_str(),
                          std::string(field) + "=" + std::to_string(value) + " max=" + std::to_string(bound - 1U));
}

fm_status_t null_argument() {
  return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "a required argument is NULL");
}

fm_status_t null_argument(const char* what) {
  return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, (std::string(what) + " is NULL").c_str());
}

fm_status_t invalid_argument(const char* what, std::string context) {
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, what, std::move(context));
}

bool is_valid_error_code(fm_error_code_t error) {
  return error >= 0 && error <= static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown);
}

fm_status_t invalid_error_code(fm_error_code_t error) {
  return invalid_argument("error code out of range", "error=" + std::to_string(error));
}

fm_status_t not_found(const char* what, std::string context) {
  return set_binding_error(formulon::FormulonErrorCode::kNotFound, what, std::move(context));
}

fm_status_t precondition_failed(const char* what, std::string context) {
  return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed, what, std::move(context));
}

fm_status_t check_sheet_row(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_row_index(row);
}

fm_status_t check_anchor_enums(std::int64_t anchor_kind, std::int64_t edit_as) {
  if (auto rc = check_enum_domain(anchor_kind, 0, FM_ANCHOR_KIND_ABSOLUTE, "anchor_kind"); rc != 0) {
    return rc;
  }
  return check_enum_domain(edit_as, 0, FM_ANCHOR_EDIT_AS_ABSOLUTE, "edit_as");
}

namespace {

fm_status_t utf8_error(std::string context) {
  return set_binding_error(formulon::FormulonErrorCode::kBindingUtf8EncodingError, "text is not valid UTF-8",
                           std::move(context));
}

}  // namespace

fm_status_t check_utf8(std::string_view text, std::string_view context) {
  return formulon::sheet_names::valid_utf8(text) ? 0 : utf8_error(std::string(context));
}

fm_status_t check_utf8(std::string_view text, std::string_view key, std::string_view value) {
  if (formulon::sheet_names::valid_utf8(text)) {
    return 0;
  }
  std::string context(key);
  context.append("=").append(value);
  return utf8_error(std::move(context));
}

fm_status_t check_utf8(std::string_view text, std::string_view key, std::size_t index) {
  return formulon::sheet_names::valid_utf8(text) ? 0 : utf8_error(std::string(key) + "=" + std::to_string(index));
}

fm_status_t check_cell_text_length(std::string_view text, const char* field) {
  constexpr std::uint64_t kCap = formulon::eval::kExcelTextCapUnits;
  // A UTF-8 sequence of at most 4 bytes yields at least one unit, so a longer
  // byte string is over the cap without walking it (or wrapping the count).
  if (text.size() <= kCap) {
    return 0;
  }
  const std::uint64_t units = text.size() > 4U * kCap ? (text.size() + 3U) / 4U : formulon::utf16_units_in(text);
  if (units <= kCap) {
    return 0;
  }
  return invalid_argument("text exceeds 32767 UTF-16 units", std::string(field) + "_units=" + std::to_string(units));
}

fm_status_t check_utf8(std::initializer_list<const char*> texts) {
  std::size_t index = 0;
  for (const char* text : texts) {
    if (text != nullptr) {
      if (auto rc = check_utf8(text, "text_index", index); rc != 0) {
        return rc;
      }
    }
    ++index;
  }
  return 0;
}

fm_status_t check_cell_xf_index(const fm_workbook_t* wb, std::size_t xf_index) {
  return check_index(xf_index, wb->workbook().styles().cell_xfs.size(), "xf_index");
}

fm_status_t check_keeps_visible_sheet(const formulon::Workbook& workbook, std::uint32_t index) {
  if (workbook.sheet(index).view().visibility() != formulon::SheetVisibility::kVisible ||
      workbook.nearest_visible_sheet(index) != index) {
    return 0;
  }
  return precondition_failed("workbook needs a visible sheet", "sheet_index=" + std::to_string(index));
}

namespace {

using formulon::parser::AstNode;
using formulon::parser::ParseErrorCode;
using formulon::parser::ParseFailure;

bool is_builtin_function_name(std::string_view name) {
  return formulon::eval::resolve_builtin_function_name(name) != nullptr;
}

// `code` over the whole of `text`.
ParseFailure whole_text_failure(ParseErrorCode code, std::string_view text) {
  ParseFailure failure;
  failure.code = code;
  failure.range.end = formulon::utf16_units_in(text);
  failure.message = formulon::parser::default_message(code);
  return failure;
}

// `failure` moved right by `units` UTF-16 code units of first-line text.
ParseFailure shifted(ParseFailure failure, std::uint32_t units) {
  failure.range.start += units;
  failure.range.end += units;
  if (failure.range.line == 1U) {
    failure.range.column += units;
  }
  return failure;
}

// UTF-16 code units of the text ahead of `body`, a suffix of `text`.
std::uint32_t prefix_units(std::string_view text, std::string_view body) {
  return formulon::utf16_units_in(text.substr(0, text.size() - body.size()));
}

// Bytes spanned by the first `units` UTF-16 code units of UTF-8 `text`.
std::size_t utf8_prefix_bytes(std::string_view text, std::uint32_t units) {
  std::size_t i = 0;
  std::uint32_t seen = 0;
  while (i < text.size() && seen < units) {
    const auto lead = static_cast<unsigned char>(text[i]);
    const std::size_t len = lead < 0x80U ? 1U : lead >= 0xF0U ? 4U : lead >= 0xE0U ? 3U : lead >= 0xC0U ? 2U : 1U;
    seen += len == 4U ? 2U : 1U;
    i += len;
  }
  return std::min(i, text.size());
}

// The call in `node` that takes `arg` as one of its arguments.
const AstNode* call_taking(const AstNode& node, const AstNode* arg) {
  if (node.kind() == formulon::parser::NodeKind::Call) {
    for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
      if (&node.as_call_arg(i) == arg) {
        return &node;
      }
    }
  }
  const AstNode* hit = nullptr;
  formulon::parser::any_child_node(node, [&](const AstNode& child) {
    hit = call_taking(child, arg);
    return hit != nullptr;
  });
  return hit;
}

// AST of formula body `src` (no leading '=') as entry accepts it, or nullptr
// with `*failure` set in `src`'s coordinates: the parse diagnostic, a
// qualified built-in, and with `refusals` a non-reference argument to a
// reference-only parameter (`SUBTOTAL(9,)`) or an argument count the callee
// does not take (`ROWS()`), the last two over the call.
const AstNode* diagnose_entry(std::string_view src, formulon::Arena& arena, bool refusals, ParseFailure* failure) {
  const AstNode* root = formulon::eval::parse_formula_entry_with_diagnostics(src, arena, failure);
  if (root == nullptr || !refusals) {
    return root;
  }
  ParseErrorCode code = ParseErrorCode::EntryReferenceRequired;
  const AstNode* call = nullptr;
  if (const AstNode* arg = formulon::eval::find_non_reference_argument(*root); arg != nullptr) {
    call = call_taking(*root, arg);
  } else if ((call = formulon::eval::find_wrong_arity_call(*root)) != nullptr) {
    code = ParseErrorCode::EntryWrongArity;
  } else {
    return root;
  }
  failure->code = code;
  failure->range = call != nullptr ? call->range() : whole_text_failure(code, src).range;
  failure->message = formulon::parser::default_message(code);
  return nullptr;
}

// `failure`, from `diagnose_entry` on the body of `parsed` (the caller's
// `caller` after normalization), placed in `caller`: as is when `parsed` and
// `caller` agree up to its end, else as a re-diagnosis of `caller` finds it
// when that yields the same code, else over all of `caller`.
ParseFailure place_failure(std::string_view caller, std::string_view parsed, bool refusals,
                           const ParseFailure& failure) {
  const ParseFailure at = shifted(failure, prefix_units(parsed, strip_formula_prefix(parsed)));
  const std::size_t bytes = utf8_prefix_bytes(parsed, at.range.end);
  if (caller.substr(0, bytes) == parsed.substr(0, bytes)) {
    return at;
  }
  const std::string_view body = strip_formula_prefix(caller);
  formulon::Arena arena;
  ParseFailure again;
  if (diagnose_entry(body, arena, refusals, &again) == nullptr && again.code == failure.code) {
    return shifted(again, prefix_units(caller, body));
  }
  return whole_text_failure(failure.code, caller);
}

// The refusal of a leading byte-order mark or a second leading '=', which
// precede any parse.
std::optional<ParseFailure> refused_entry_prefix(std::string_view formula) {
  ParseFailure failure;
  if (formula.substr(0, 3) == "\xEF\xBB\xBF") {
    failure.code = ParseErrorCode::EntryBomNotSupported;
    failure.range.end = 1;
  } else if (formula.substr(0, 2) == "==") {
    failure.code = ParseErrorCode::EntryDoubleEquals;
    failure.range = formulon::parser::TextRange{1U, 2U, 1U, 2U};
  } else {
    return std::nullopt;
  }
  failure.message = formulon::parser::default_message(failure.code);
  return failure;
}

// AST of entered `formula`'s stored spelling, written to `*stored`, or
// nullptr with `*failure` placed in `caller` (the text the caller passed,
// which `formula` binds): whitespace Excel refuses, or a stored spelling
// `diagnose_entry` refuses.
const AstNode* parse_entered_formula(std::string_view caller, std::string_view formula, formulon::Arena& arena,
                                     std::string* stored, ParseFailure* failure) {
  *stored =
      formulon::Workbook::normalize_formula_text(std::string(formula), formulon::Workbook::FormulaSource::kEntered);
  if (formulon::parser::has_refused_whitespace(formula, &is_builtin_function_name)) {
    *failure = whole_text_failure(ParseErrorCode::EntryRefusedWhitespace, caller);
    return nullptr;
  }
  ParseFailure raw;
  const AstNode* root = diagnose_entry(strip_formula_prefix(*stored), arena, /*refusals=*/true, &raw);
  if (root == nullptr) {
    *failure = place_failure(caller, *stored, /*refusals=*/true, raw);
  }
  return root;
}

// Parse check of entered text: typed `_xlpm.` before a cell-shaped name fails
// it, and so does `parse_entered_formula`. `*stored` receives the stored
// spelling.
fm_status_t check_entered_formula_parses(std::string_view formula, std::string_view context, std::string* stored) {
  if (auto rc = check_utf8(formula, context); rc != 0) {
    return rc;
  }
  if (formulon::eval::has_entered_cell_shaped_param(formula)) {
    return formula_parse_failed(whole_text_failure(ParseErrorCode::EntryParamShapedName, formula),
                                std::string(context));
  }
  formulon::Arena arena;
  ParseFailure failure;
  if (parse_entered_formula(formula, formula, arena, stored, &failure) == nullptr) {
    return formula_parse_failed(failure, std::string(context));
  }
  return 0;
}

}  // namespace

fm_status_t check_stored_formula(std::string_view formula, std::string_view context, std::string* out) {
  if (formula.empty()) {
    out->clear();
    return 0;
  }
  // The parser skips one '=' of its own, so a second one is refused here.
  if (const std::optional<ParseFailure> refused = refused_entry_prefix(formula)) {
    return formula_parse_failed(*refused, std::string(context));
  }
  std::string stored;
  if (auto rc = check_entered_formula_parses(formula, context, &stored); rc != 0) {
    return rc;
  }
  *out = std::string(strip_formula_prefix(stored));
  return 0;
}

std::string ingest_entered_formula(fm_workbook_t* wb, std::string_view formula) {
  if (formula.empty()) {
    return std::string();
  }
  std::string text(formula);
  wb->workbook().ingest_entered_formula(text);
  return std::string(strip_formula_prefix(text));
}

std::string format_runtime_id(std::uint32_t id) {
  return std::to_string(id);
}

bool parse_runtime_id(std::string_view text, std::uint32_t* out) {
  if (text.empty() || text.size() > 10U || text.front() == '0') {
    return false;
  }
  std::uint64_t value = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return false;
    }
    value = value * 10U + static_cast<std::uint64_t>(c - '0');
  }
  if (value > std::numeric_limits<std::uint32_t>::max()) {
    return false;
  }
  *out = static_cast<std::uint32_t>(value);
  return true;
}

fm_status_t check_entered_cell_formula(const fm_workbook_t* wb, std::size_t sheet, std::uint32_t row, std::uint32_t col,
                                       std::string_view formula, std::string_view context, std::string* bound) {
  if (auto rc = check_utf8(formula, context); rc != 0) {
    return rc;
  }
  if (const std::optional<ParseFailure> refused = refused_entry_prefix(formula)) {
    return formula_parse_failed(*refused, std::string(context));
  }
  *bound = wb->workbook().bind_entered_structured_refs(formula, sheet, row, col);
  if (formulon::eval::has_entered_cell_shaped_param(*bound)) {
    return formula_parse_failed(whole_text_failure(ParseErrorCode::EntryParamShapedName, formula),
                                std::string(context));
  }
  std::string stored;
  ParseFailure a1_failure;
  bool a1_failed = false;
  {
    formulon::Arena arena;
    const AstNode* root = parse_entered_formula(formula, *bound, arena, &stored, &a1_failure);
    if (root != nullptr && !formulon::eval::needs_r1c1_entry(*root, *bound, wb->workbook())) {
      *bound = with_formula_prefix(*bound);
      return 0;
    }
    a1_failed = root == nullptr;
  }
  // Excel reads text that fails as A1 again as R1C1 relative to the cell.
  std::optional<std::string> translated = formulon::eval::translate_r1c1_entry(*bound, row, col);
  if (translated) {
    formulon::Arena arena;
    ParseFailure r1c1_failure;
    if (parse_entered_formula(*translated, *translated, arena, &stored, &r1c1_failure) != nullptr) {
      *bound = with_formula_prefix(*translated);
      return 0;
    }
  }
  // A failed A1 reading is the one the caller's text locates; text that read
  // as A1 failed only as the R1C1 reference it names.
  return formula_parse_failed(a1_failed ? a1_failure : whole_text_failure(ParseErrorCode::InvalidReference, formula),
                              std::string(context));
}

const formulon::parser::AstNode* parse_evaluated_formula(std::string_view formula, std::string_view bound,
                                                         formulon::Arena& arena, fm_status_t* rc) {
  ParseFailure raw;
  const AstNode* root = diagnose_entry(strip_formula_prefix(bound), arena, /*refusals=*/false, &raw);
  *rc =
      root == nullptr ? formula_parse_failed(place_failure(formula, bound, /*refusals=*/false, raw), std::string()) : 0;
  return root;
}

std::string_view table_display_name_at(const formulon::Workbook& book, std::size_t sheet, std::uint32_t row,
                                       std::uint32_t col) {
  const std::optional<std::size_t> table = book.table_containing(sheet, row, col);
  return table ? std::string_view(book.tables()[*table].display_name) : std::string_view();
}

std::string display_formula_text(const formulon::Workbook& book, std::string_view model, std::string_view own_table) {
  if (!may_respell_formula(model)) {
    return std::string(model);
  }
  const std::optional<formulon::sheet_names::NameList> table_list =
      own_table.empty()
          ? std::nullopt
          : std::optional<formulon::sheet_names::NameList>(formulon::sheet_names::table_name_list(book.tables()));
  std::string shown = formulon::parser::display_structured_refs(model, own_table, table_list ? &*table_list : nullptr);
  if (book.spells_self_book_names(shown)) {
    shown = book.spell_self_book_names(shown);
  }
  return formulon::parser::display_cell_shaped_names(shown);
}

const char* display_formula(const formulon::Workbook& book, const std::string& model, TextStore& store,
                            std::string_view own_table) {
  if (!may_respell_formula(model)) {
    return model.c_str();
  }
  std::string shown = display_formula_text(book, model, own_table);
  return shown == model ? model.c_str() : store.emplace_back(std::move(shown)).c_str();
}

fm_status_t resolve_sheet_table(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t table_index,
                                std::size_t* out_global) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& tables = wb->workbook().tables();
  std::size_t local = 0;
  for (std::size_t i = 0; i < tables.size(); ++i) {
    if (tables[i].sheet_index != sheet_index) {
      continue;
    }
    if (local == table_index) {
      *out_global = i;
      return 0;
    }
    ++local;
  }
  return check_index(table_index, local, "table_index");
}

std::size_t local_table_index(const formulon::Workbook& book, std::size_t global) {
  const auto& tables = book.tables();
  std::size_t local = 0;
  for (std::size_t i = 0; i < global; ++i) {
    if (tables[i].sheet_index == tables[global].sheet_index) {
      ++local;
    }
  }
  return local;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
