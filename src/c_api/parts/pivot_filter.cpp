//
// C ABI - PivotTable filter surface. Shares helpers with `pivot_table.cpp`
// via `pivot_internal.h`.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <variant>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::pivot_axis_from_fm;
using formulon::c_api::parts::pivot_axis_to_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_filter_type_from_fm;
using formulon::c_api::parts::PivotEdit;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::validate;

// ---------------------------------------------------------------------------
// PivotTable filters
// ---------------------------------------------------------------------------

namespace {

// Stores one numeric payload slot. NaN and infinities have no spelling in a
// saved filter, so they are refused.
fm_status_t numeric_payload(fm_pivot_filter_value_kind_t kind, std::int32_t as_int, double as_double,
                            std::variant<std::monostate, int, double, std::string>& out, const char* api,
                            const char* field) {
  if (kind == FM_PIVOT_FILTER_VALUE_INT) {
    out = static_cast<int>(as_int);
    return 0;
  }
  if (!std::isfinite(as_double)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": " + field + " is not finite").c_str());
  }
  out = as_double;
  return 0;
}

fm_status_t add_pivot_filter_impl(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                  const fm_pivot_filter_spec_t* spec, const char* api) {
  if (wb == nullptr || spec == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": NULL argument").c_str());
  }
  if (spec->field_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api) + ": spec->field_name is NULL").c_str());
  }
  if (auto rc = validate(*spec, api); rc != 0) {
    return rc;
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, api);
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  const formulon::pivot::FilterType type = pivot_filter_type_from_fm(spec->type);
  const bool caption = formulon::pivot::is_caption_filter_type(type);
  const bool ranged = formulon::pivot::is_range_filter_type(type);
  const auto invalid = [&](const char* what, const std::string& detail) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, (std::string(api) + ": " + what).c_str(),
                             detail);
  };
  const std::string kind_detail = "type=" + std::to_string(static_cast<int>(spec->type)) +
                                  " value_kind=" + std::to_string(static_cast<int>(spec->value_kind));
  if (formulon::pivot::is_operandless_filter_type(type)) {
    if (spec->value_kind != FM_PIVOT_FILTER_VALUE_NONE) {
      return invalid("this filter type takes no value", kind_detail);
    }
  } else if (spec->value_kind == FM_PIVOT_FILTER_VALUE_NONE) {
    return invalid("spec->value_kind is unset", kind_detail);
  } else if (caption != (spec->value_kind == FM_PIVOT_FILTER_VALUE_TEXT)) {
    return invalid(caption ? "a label filter takes a text value" : "this filter type takes a number", kind_detail);
  }
  const bool text_high = caption && ranged;
  if (text_high != (spec->value_high_text != nullptr) ||
      text_high != (spec->value_high_kind == FM_PIVOT_FILTER_VALUE_TEXT)) {
    return invalid(text_high ? "a label range takes its upper bound in value_high_text"
                             : "value_high_text is only taken by a label range",
                   "type=" + std::to_string(static_cast<int>(spec->type)));
  }
  if (formulon::pivot::is_value_filter_type(type)) {
    if (auto rc = check_index(spec->data_field_index, table->data_fields().size(), api, "data_field_index"); rc != 0) {
      return rc;
    }
  }
  // The engine skips a filter whose field does not resolve, so an unresolvable
  // name would leave any filter kind a silent no-op.
  if (!formulon::pivot::resolve_field_by_any_name(*table, spec->field_name).has_value()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             (std::string(api) + ": field_name does not name a pivot field").c_str(),
                             std::string("field_name=") + spec->field_name);
  }
  formulon::pivot::PivotFilter filter;
  filter.axis = pivot_axis_from_fm(spec->axis);
  filter.field_name = spec->field_name;
  filter.type = type;
  filter.data_field_index = spec->data_field_index;
  switch (spec->value_kind) {
    case FM_PIVOT_FILTER_VALUE_INT:
    case FM_PIVOT_FILTER_VALUE_DOUBLE:
      if (auto rc = numeric_payload(spec->value_kind, spec->value_int, spec->value_double, filter.value, api, "value");
          rc != 0) {
        return rc;
      }
      break;
    case FM_PIVOT_FILTER_VALUE_TEXT:
      if (spec->value_text == nullptr) {
        return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                                 (std::string(api) + ": spec->value_text is NULL").c_str());
      }
      filter.value = std::string(spec->value_text);
      break;
    case FM_PIVOT_FILTER_VALUE_NONE:
      break;
    default:
      return invalid("spec->value_kind is out of range", kind_detail);
  }
  switch (spec->value_high_kind) {
    case FM_PIVOT_FILTER_VALUE_INT:
    case FM_PIVOT_FILTER_VALUE_DOUBLE:
      if (auto rc = numeric_payload(spec->value_high_kind, spec->value_high_int, spec->value_high_double,
                                    filter.value_high, api, "value_high");
          rc != 0) {
        return rc;
      }
      break;
    case FM_PIVOT_FILTER_VALUE_TEXT:
      filter.value_high = std::string(spec->value_high_text);
      break;
    case FM_PIVOT_FILTER_VALUE_NONE:
      // Leave default monostate.
      break;
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(api) + ": spec->value_high_kind not supported").c_str(),
                               "value_high_kind=" + std::to_string(static_cast<int>(spec->value_high_kind)));
  }
  // The top / bottom amounts: a percent within 0..100, a count or sum above 0.
  if (formulon::pivot::is_value_filter_type(type)) {
    const double amount = formulon::pivot::filter_number_value(filter);
    const bool percent =
        type == formulon::pivot::FilterType::ValueTopPercent || type == formulon::pivot::FilterType::ValueBottomPercent;
    const bool count_or_sum =
        type == formulon::pivot::FilterType::ValueTop10 || type == formulon::pivot::FilterType::ValueBottomCount ||
        type == formulon::pivot::FilterType::ValueTopSum || type == formulon::pivot::FilterType::ValueBottomSum;
    if ((percent && (amount < 0.0 || amount > 100.0)) || (count_or_sum && amount <= 0.0)) {
      return invalid("value is out of range for this filter type", kind_detail);
    }
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->mutable_active_filters().push_back(std::move(filter));
  return edit.commit();
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_filter_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count, "fm_workbook_pivot_filter_count",
                     [](const formulon::pivot::PivotTable& table) { return table.active_filters().size(); });
}

extern "C" fm_status_t fm_workbook_pivot_filter_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                   std::size_t pivot_index, std::size_t filter_idx,
                                                   fm_pivot_filter_spec_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, "fm_workbook_pivot_filter_at", &table);
      rc != 0) {
    return rc;
  }
  const auto& filters = table->active_filters();
  if (auto rc = check_index(filter_idx, filters.size(), "fm_workbook_pivot_filter_at", "filter_idx"); rc != 0) {
    return rc;
  }

  const auto& filter = filters[filter_idx];
  const auto axis = pivot_axis_to_fm(filter.axis);
  if (!axis.has_value()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_filter_at: model axis is not representable");
  }
  if (filter.type > formulon::pivot::FilterType::DateM12) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_filter_at: model filter type is not representable");
  }

  fm_pivot_filter_spec_t result{};
  result.axis = *axis;
  result.field_name = filter.field_name.c_str();
  result.type = static_cast<fm_pivot_filter_type_t>(filter.type);
  result.value_kind = FM_PIVOT_FILTER_VALUE_NONE;
  result.value_high_kind = FM_PIVOT_FILTER_VALUE_NONE;
  result.data_field_index = filter.data_field_index;

  if (const auto* int_value = std::get_if<int>(&filter.value); int_value != nullptr) {
    result.value_kind = FM_PIVOT_FILTER_VALUE_INT;
    result.value_int = static_cast<std::int32_t>(*int_value);
  } else if (const auto* double_value = std::get_if<double>(&filter.value); double_value != nullptr) {
    result.value_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    result.value_double = *double_value;
  } else if (const auto* text_value = std::get_if<std::string>(&filter.value); text_value != nullptr) {
    result.value_kind = FM_PIVOT_FILTER_VALUE_TEXT;
    result.value_text = text_value->c_str();
  }

  if (const auto* int_value_high = std::get_if<int>(&filter.value_high); int_value_high != nullptr) {
    result.value_high_kind = FM_PIVOT_FILTER_VALUE_INT;
    result.value_high_int = static_cast<std::int32_t>(*int_value_high);
  } else if (const auto* double_value_high = std::get_if<double>(&filter.value_high); double_value_high != nullptr) {
    result.value_high_kind = FM_PIVOT_FILTER_VALUE_DOUBLE;
    result.value_high_double = *double_value_high;
  } else if (const auto* text_value_high = std::get_if<std::string>(&filter.value_high); text_value_high != nullptr) {
    result.value_high_kind = FM_PIVOT_FILTER_VALUE_TEXT;
    result.value_high_text = text_value_high->c_str();
  }
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_filter_add(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    const fm_pivot_filter_spec_t* spec) {
  clear_last_error();
  return add_pivot_filter_impl(wb, sheet_index, pivot_index, spec, "fm_workbook_pivot_filter_add");
}

extern "C" fm_status_t fm_workbook_pivot_filter_clear(fm_workbook_t* wb, std::size_t sheet_index,
                                                      std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, "fm_workbook_pivot_filter_clear", &table);
      rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->erase_active_filters_if([](const formulon::pivot::PivotFilter&) { return true; });
  return edit.commit();
}

extern "C" fm_status_t fm_workbook_pivot_filter_remove_at(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::size_t filter_idx) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_filter_remove_at: wb is NULL");
  }
  auto* table = resolve_pivot_mut(wb->workbook(), sheet_index, pivot_index, "fm_workbook_pivot_filter_remove_at");
  if (table == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (auto rc =
          check_index(filter_idx, table->active_filters().size(), "fm_workbook_pivot_filter_remove_at", "filter_idx");
      rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  std::size_t position = 0;
  table->erase_active_filters_if([&](const formulon::pivot::PivotFilter&) { return position++ == filter_idx; });
  return edit.commit();
}
