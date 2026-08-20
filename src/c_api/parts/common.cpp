
#include "c_api/parts/common.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <type_traits>
#include <utility>

#include "c_api/formulon_c.h"
#include "cf/cf_types.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace c_api {
namespace parts {

// ABI identity: the `FM_VAL_*` discriminators are numerically equal to the
// `formulon::ValueKind` ordinals. `value_to_fm` below maps each ValueKind
// to its `FM_VAL_*` by name; these static_asserts pin the parity so a
// reorder of either enum is a build break rather than a silent wire-format
// drift across the C ABI.
static_assert(FM_VAL_BLANK == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Blank),
              "FM_VAL_BLANK must match ValueKind::Blank");
static_assert(FM_VAL_NUMBER == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Number),
              "FM_VAL_NUMBER must match ValueKind::Number");
static_assert(FM_VAL_BOOL == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Bool),
              "FM_VAL_BOOL must match ValueKind::Bool");
static_assert(FM_VAL_TEXT == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Text),
              "FM_VAL_TEXT must match ValueKind::Text");
static_assert(FM_VAL_ERROR == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Error),
              "FM_VAL_ERROR must match ValueKind::Error");
static_assert(FM_VAL_ARRAY == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Array),
              "FM_VAL_ARRAY must match ValueKind::Array");
static_assert(FM_VAL_REF == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Ref),
              "FM_VAL_REF must match ValueKind::Ref");
static_assert(FM_VAL_LAMBDA == static_cast<std::underlying_type_t<ValueKind>>(ValueKind::Lambda),
              "FM_VAL_LAMBDA must match ValueKind::Lambda");

namespace {

// Thread-local diagnostics for the most recent C API call on this thread.
//
// Both buffers are always non-empty in the sense that they own valid string
// storage, but their `c_str()` may point to an empty string. Returning
// `c_str()` therefore satisfies the public "never NULL" contract.
thread_local std::string g_last_error_message;
thread_local std::string g_last_error_context;

}  // namespace

void clear_last_error() {
  g_last_error_message.clear();
  g_last_error_context.clear();
}

fm_status_t set_last_error(const formulon::Error& err) {
  g_last_error_message = err.message;
  g_last_error_context = err.context;
  return static_cast<fm_status_t>(err.code);
}

fm_status_t set_binding_error(formulon::FormulonErrorCode code, const char* message, std::string context) {
  formulon::Error err;
  err.code = code;
  err.message = message != nullptr ? message : "";
  err.context = std::move(context);
  return set_last_error(err);
}

const char* last_error_message() {
  return g_last_error_message.c_str();
}

const char* last_error_context() {
  return g_last_error_context.c_str();
}

bool check_range_count(std::uint32_t n, const char* api) {
  if (n > kMaxRangesPerCApiCall) {
    set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "range_count exceeds per-call cap",
        std::string(api) + ": range_count=" + std::to_string(n) + " cap=" + std::to_string(kMaxRangesPerCApiCall));
    return false;
  }
  return true;
}

fm_status_t check_sheet_rect(std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                             std::uint32_t last_col, const char* api) {
  if (formulon::Sheet::rect_in_grid(first_row, first_col, last_row, last_col)) {
    return 0;
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "cell rectangle out of range",
                           std::string(api) + ": first_row=" + std::to_string(first_row) +
                               " first_col=" + std::to_string(first_col) + " last_row=" + std::to_string(last_row) +
                               " last_col=" + std::to_string(last_col));
}

fm_status_t check_enum_domain(std::int64_t value, std::int64_t max, const char* api, const char* field) {
  if (value >= 0 && value <= max) {
    return 0;
  }
  return set_binding_error(
      formulon::FormulonErrorCode::kInvalidArgument, "enum ordinal out of range",
      std::string(api) + ": " + field + "=" + std::to_string(value) + " max=" + std::to_string(max));
}

fm_status_t check_guid(const char* text, const char* api, const char* field) {
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
        shaped = (value[i] >= '0' && value[i] <= '9') || (value[i] >= 'a' && value[i] <= 'f') ||
                 (value[i] >= 'A' && value[i] <= 'F');
        break;
      default:
        shaped = value[i] == kShape[i];
        break;
    }
  }
  if (shaped) {
    return 0;
  }
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "value is not a GUID",
                           std::string(api) + ": " + field + "=" + std::string(value));
}

namespace {

// Runs `check_enum_domain` over every `fm_cfvo_t` in a threshold array.
// The array bound is the caller's, so an entry point that has not yet
// established `count` against its pointer must not reach here.
fm_status_t check_cfvo_types(const fm_cfvo_t* thresholds, std::uint32_t count, const char* api, const char* field) {
  for (std::uint32_t i = 0; i < count; ++i) {
    if (auto rc = check_enum_domain(thresholds[i].type, static_cast<std::int64_t>(cf::CfvoType::AutoMax), api, field);
        rc != 0) {
      return rc;
    }
  }
  return 0;
}

}  // namespace

fm_status_t validate(const fm_cf_rule_t& rule, const char* api) {
  if (rule.sqref_count == 0) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "sqref_count must be >= 1",
                             std::string(api) + ": sqref_count=0");
  }
  if (rule.sqref == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "sqref is NULL while sqref_count > 0",
                             std::string(api));
  }
  if (!check_range_count(rule.sqref_count, api)) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  for (std::uint32_t i = 0; i < rule.sqref_count; ++i) {
    // A range the A1 encoder cannot spell reaches the writer as an empty
    // reference, which Excel reads as a broken conditional-formatting
    // block. Rejecting it here names the offending rectangle at set time
    // instead of losing the format at save time.
    if (auto rc = check_sheet_rect(rule.sqref[i].first_row, rule.sqref[i].first_col, rule.sqref[i].last_row,
                                   rule.sqref[i].last_col, api);
        rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_enum_domain(rule.type, static_cast<std::int64_t>(cf::RuleType::UniqueValues), api, "type");
      rc != 0) {
    return rc;
  }
  if (rule.op_engaged != 0) {
    if (auto rc = check_enum_domain(rule.op, static_cast<std::int64_t>(cf::CellIsOperator::NotBetween), api, "op");
        rc != 0) {
      return rc;
    }
  }
  if (rule.time_period_engaged != 0) {
    if (auto rc = check_enum_domain(rule.time_period, static_cast<std::int64_t>(cf::TimePeriod::NextMonth), api,
                                    "time_period");
        rc != 0) {
      return rc;
    }
  }
  if (rule.data_bar_engaged != 0) {
    if (auto rc = check_cfvo_types(&rule.data_bar_min, 1U, api, "data_bar_min.type"); rc != 0) {
      return rc;
    }
    if (auto rc = check_cfvo_types(&rule.data_bar_max, 1U, api, "data_bar_max.type"); rc != 0) {
      return rc;
    }
    if (rule.data_bar_axis_position_engaged != 0) {
      if (auto rc =
              check_enum_domain(rule.data_bar_axis_position, static_cast<std::int64_t>(cf::DataBarAxisPosition::None),
                                api, "data_bar_axis_position");
          rc != 0) {
        return rc;
      }
    }
  }
  // The colour-scale scan is bounded by the array shape the entry point
  // accepts, so a record that declares more thresholds than the schema
  // allows is rejected there rather than scanned here.
  if (rule.color_scale_thresholds != nullptr && rule.color_scale_count >= 2U && rule.color_scale_count <= 3U) {
    if (auto rc =
            check_cfvo_types(rule.color_scale_thresholds, rule.color_scale_count, api, "color_scale_thresholds[].type");
        rc != 0) {
      return rc;
    }
  }
  if (rule.icon_set_engaged != 0) {
    if (auto rc = check_enum_domain(rule.icon_set_name, static_cast<std::int64_t>(cf::IconSetName::Five_Quarters), api,
                                    "icon_set_name");
        rc != 0) {
      return rc;
    }
    if (rule.icon_set_thresholds != nullptr) {
      if (auto rc = check_cfvo_types(rule.icon_set_thresholds, rule.icon_set_threshold_count, api,
                                     "icon_set_thresholds[].type");
          rc != 0) {
        return rc;
      }
    }
  }
  // The id reaches the file as `<x14:cfRule id="...">` and `<x14:id>`,
  // both of schema type `ST_Guid`. Excel repairs a workbook whose
  // extension block holds anything else, and repairing means dropping
  // every conditional format in it.
  return check_guid(rule.id, api, "id");
}

fm_status_t validate(const fm_data_validation& v, const char* api) {
  if (v.range_count > 0 && v.ranges == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "ranges is NULL while range_count > 0",
                             std::string(api) + ": range_count=" + std::to_string(v.range_count));
  }
  if (!check_range_count(v.range_count, api)) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  for (std::uint32_t i = 0; i < v.range_count; ++i) {
    if (auto rc = check_sheet_rect(v.ranges[i].first_row, v.ranges[i].first_col, v.ranges[i].last_row,
                                   v.ranges[i].last_col, api);
        rc != 0) {
      return rc;
    }
  }
  // The three domains the writer's attribute tables spell out
  // (`sheet_xml_builder.cpp`): a value past them is emitted as an absent
  // attribute, so a mistyped `type` saves as "no validation at all".
  if (auto rc = check_enum_domain(v.type, 7, api, "type"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(v.op, 7, api, "op"); rc != 0) {
    return rc;
  }
  return check_enum_domain(v.error_style, 2, api, "error_style");
}

fm_status_t validate(const fm_pivot_field_spec_t& spec, const char* api) {
  if (spec.source_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "spec->source_name is NULL",
                             std::string(api));
  }
  return check_enum_domain(spec.axis, FM_PIVOT_AXIS_PAGE, api, "axis");
}

fm_status_t validate(const fm_pivot_data_field_spec_t& spec, const char* api) {
  if (spec.name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "spec->name is NULL", std::string(api));
  }
  if (auto rc = check_enum_domain(spec.aggregation, FM_PIVOT_AGG_VARP, api, "aggregation"); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.show_as, FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT, api, "show_as");
}

fm_status_t validate(const fm_pivot_filter_spec_t& spec, const char* api) {
  if (auto rc = check_enum_domain(spec.axis, FM_PIVOT_AXIS_PAGE, api, "axis"); rc != 0) {
    return rc;
  }
  return check_enum_domain(spec.type, FM_PIVOT_FILTER_LABEL_DATE, api, "type");
}

fm_status_t validate(const fm_viewport& viewport, const char* api) {
  // `SheetCellRange::sheet_id` is `std::uint16_t`; reject the narrowing
  // path so a caller-supplied sheet > 0xFFFF does not silently address a
  // different sheet (or wrap to 0). Excel's own cap is far below 0xFFFF.
  if (viewport.sheet > 0xFFFFU) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "viewport->sheet exceeds 16-bit sheet id range",
                             std::string(api) + ": sheet=" + std::to_string(viewport.sheet));
  }
  return 0;
}

void value_to_fm(const formulon::Value& v, TextStore& store, fm_value_t* out) {
  switch (v.kind()) {
    case formulon::ValueKind::Blank:
      out->kind = FM_VAL_BLANK;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Number:
      out->kind = FM_VAL_NUMBER;
      out->u.number = v.as_number();
      return;
    case formulon::ValueKind::Bool:
      out->kind = FM_VAL_BOOL;
      out->u.boolean = v.as_boolean() ? 1 : 0;
      return;
    case formulon::ValueKind::Text: {
      const std::string_view text = v.as_text();
      store.emplace_back(text.data(), text.size());
      out->kind = FM_VAL_TEXT;
      out->u.text = store.back().c_str();
      return;
    }
    case formulon::ValueKind::Error:
      out->kind = FM_VAL_ERROR;
      // The C-ABI wire format for errors is the raw `ErrorCode` ordinal
      // (Null=0, Div0=1, Value=2, ...), NOT the OOXML wire code. This is
      // the single source of truth for the boundary; see the
      // `fm_value_t::u.error_code` contract in `formulon_c.h`.
      out->u.error_code = static_cast<int32_t>(v.as_error());
      return;
    case formulon::ValueKind::Array:
      // Array passthrough is reserved; the kind is reported but no
      // payload is exposed across the boundary in this bundle.
      out->kind = FM_VAL_ARRAY;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Ref:
      out->kind = FM_VAL_REF;
      out->u.number = 0.0;
      return;
    case formulon::ValueKind::Lambda:
      out->kind = FM_VAL_LAMBDA;
      out->u.number = 0.0;
      return;
  }
  // Defensive default: surface as Blank rather than leaving uninitialised
  // bytes on the boundary. The switch above is exhaustive over every
  // `ValueKind` enumerator, so this branch is unreachable in practice.
  out->kind = FM_VAL_BLANK;
  out->u.number = 0.0;
}

fm_status_t check_sheet_index(const fm_workbook_t* wb, std::size_t sheet_index, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

fm_status_t check_sheet_u32(const fm_workbook_t* wb, std::uint32_t sheet, const char* fn) {
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, fn);
  }
  if (static_cast<std::size_t>(sheet) >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, fn,
        "sheet=" + std::to_string(sheet) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  return 0;
}

}  // namespace parts
}  // namespace c_api
}  // namespace formulon
