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
#include "c_api/parts/struct_v1_sizes.h"
#include "pivot/field_lookup.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"

using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_filter_type_from_fm;
using formulon::c_api::parts::PivotEdit;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::validate;

// ---------------------------------------------------------------------------
// PivotTable filters
// ---------------------------------------------------------------------------

namespace {

// Stores one numeric payload slot. NaN and infinities have no spelling in a
// saved filter, so they are refused.
using FilterPayload = std::variant<std::monostate, int, double, std::string>;

fm_status_t numeric_payload(std::int32_t kind, std::int32_t as_int, double as_double, FilterPayload& out,
                            const char* field) {
  if (kind == FM_PIVOT_FILTER_OPERAND_KIND_INT) {
    out = static_cast<int>(as_int);
    return 0;
  }
  if (auto rc = check_finite(as_double, field); rc != 0) {
    return rc;
  }
  out = as_double;
  return 0;
}

// Spells one payload slot as its kind / int / double / text fields; the text
// view borrows from `payload`. An unset slot leaves the fields as they are.
void payload_to_fm(const FilterPayload& payload, std::int32_t* kind, std::int32_t* as_int, double* as_double,
                   const char** as_text) {
  if (const auto* int_value = std::get_if<int>(&payload); int_value != nullptr) {
    *kind = FM_PIVOT_FILTER_OPERAND_KIND_INT;
    *as_int = static_cast<std::int32_t>(*int_value);
  } else if (const auto* double_value = std::get_if<double>(&payload); double_value != nullptr) {
    *kind = FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE;
    *as_double = *double_value;
  } else if (const auto* text_value = std::get_if<std::string>(&payload); text_value != nullptr) {
    *kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
    *as_text = text_value->c_str();
  }
}

fm_status_t add_pivot_filter_impl(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                  const fm_pivot_filter_t* spec, std::size_t* out_index) {
  if (wb == nullptr || spec == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*spec); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const formulon::pivot::FilterType type = pivot_filter_type_from_fm(static_cast<fm_pivot_filter_kind_t>(spec->kind));
  const bool caption = formulon::pivot::is_caption_filter_type(type);
  const bool ranged = formulon::pivot::is_range_filter_type(type);
  const std::string kind_detail = "kind=" + std::to_string(static_cast<int>(spec->kind)) +
                                  " value_kind=" + std::to_string(static_cast<int>(spec->value_kind));
  if (formulon::pivot::is_operandless_filter_type(type)) {
    if (spec->value_kind != FM_PIVOT_FILTER_OPERAND_KIND_ABSENT) {
      return invalid_argument("value_kind must be absent for this filter type", kind_detail);
    }
  } else if (spec->value_kind == FM_PIVOT_FILTER_OPERAND_KIND_ABSENT) {
    return invalid_argument("value_kind is required for this filter type", kind_detail);
  } else if (caption != (spec->value_kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT)) {
    return invalid_argument(
        caption ? "value_kind must be text for a label filter" : "value_kind must be a number for this filter type",
        kind_detail);
  }
  const bool text_high = caption && ranged;
  if (text_high != (spec->value_high_text != nullptr) ||
      text_high != (spec->value_high_kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT)) {
    return invalid_argument(text_high ? "value_high_text is required for a label range"
                                      : "value_high_text must be absent unless the filter is a label range",
                            "kind=" + std::to_string(static_cast<int>(spec->kind)));
  }
  if (formulon::pivot::is_value_filter_type(type)) {
    if (auto rc = check_index(spec->data_field_index, table->data_fields().size(), "data_field_index"); rc != 0) {
      return rc;
    }
  }
  if (auto rc = check_index(spec->field_index, table->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  // The filter takes its field's axis; a field on no axis would leave it a silent no-op.
  const formulon::pivot::PivotField& target = table->fields()[spec->field_index];
  if (target.axis == formulon::pivot::PivotAxis::None) {
    return invalid_argument(("field_index " + std::to_string(spec->field_index) + " names a field on no axis").c_str(),
                            "field_index=" + std::to_string(spec->field_index));
  }
  formulon::pivot::PivotFilter filter;
  filter.axis = target.axis;
  filter.field_name = target.source_name;
  filter.type = type;
  filter.data_field_index = spec->data_field_index;
  switch (spec->value_kind) {
    case FM_PIVOT_FILTER_OPERAND_KIND_INT:
    case FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE:
      if (auto rc = numeric_payload(spec->value_kind, spec->value_int, spec->value_double, filter.value, "value");
          rc != 0) {
        return rc;
      }
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_TEXT:
      if (spec->value_text == nullptr) {
        return null_field("value_text");
      }
      filter.value = std::string(spec->value_text);
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_ABSENT:
      break;
    default:
      return invalid_argument(
          ("value_kind " + std::to_string(static_cast<int>(spec->value_kind)) + " is not a FilterOperandKind").c_str(),
          kind_detail);
  }
  switch (spec->value_high_kind) {
    case FM_PIVOT_FILTER_OPERAND_KIND_INT:
    case FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE:
      if (auto rc = numeric_payload(spec->value_high_kind, spec->value_high_int, spec->value_high_double,
                                    filter.value_high, "value_high");
          rc != 0) {
        return rc;
      }
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_TEXT:
      filter.value_high = std::string(spec->value_high_text);
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_ABSENT:
      // Leave default monostate.
      break;
    default:
      return invalid_argument(
          ("value_high_kind " + std::to_string(static_cast<int>(spec->value_high_kind)) + " is not a FilterOperandKind")
              .c_str(),
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
    if (percent && (amount < 0.0 || amount > 100.0)) {
      return invalid_argument("value must be in 0..100 for a percent filter", kind_detail);
    }
    if (count_or_sum && amount <= 0.0) {
      return invalid_argument("value must be above 0 for this filter type", kind_detail);
    }
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->mutable_active_filters().push_back(std::move(filter));
  const std::size_t added = table->active_filters().size() - 1U;
  const fm_status_t status = edit.commit();
  if (status == 0) {
    *out_index = added;
  }
  return status;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_pivot_filter_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                       std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count,
                     [](const formulon::pivot::PivotTable& table) { return table.active_filters().size(); });
}

extern "C" fm_status_t fm_sheet_get_pivot_filter_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                    std::size_t pivot_index, std::size_t filter_index,
                                                    fm_pivot_filter_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const auto& filters = table->active_filters();
  if (auto rc = check_index(filter_index, filters.size(), "filter_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }

  const auto& filter = filters[filter_index];
  const auto field = formulon::pivot::resolve_field_by_any_name(*table, filter.field_name);
  if (!field.has_value()) {
    return invalid_argument("field_name names no pivot field", "field_name=" + filter.field_name);
  }
  if (filter.type > formulon::pivot::FilterType::DateM12) {
    return invalid_argument("kind is not a representable filter kind");
  }

  fm_pivot_filter_t result{};
  result.struct_size = out->struct_size;
  result.field_index = static_cast<std::uint32_t>(*field);
  result.kind = static_cast<fm_pivot_filter_kind_t>(filter.type);
  result.value_kind = FM_PIVOT_FILTER_OPERAND_KIND_ABSENT;
  result.value_high_kind = FM_PIVOT_FILTER_OPERAND_KIND_ABSENT;
  result.data_field_index = filter.data_field_index;

  payload_to_fm(filter.value, &result.value_kind, &result.value_int, &result.value_double, &result.value_text);
  payload_to_fm(filter.value_high, &result.value_high_kind, &result.value_high_int, &result.value_high_double,
                &result.value_high_text);
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_sheet_add_pivot_filter(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                 const fm_pivot_filter_t* spec, std::size_t* out_index) {
  clear_last_error();
  return add_pivot_filter_impl(wb, sheet_index, pivot_index, spec, out_index);
}

extern "C" fm_status_t fm_sheet_clear_pivot_filters(fm_workbook_t* wb, std::size_t sheet_index,
                                                    std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->erase_active_filters_if([](const formulon::pivot::PivotFilter&) { return true; });
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_remove_pivot_filter(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                    std::size_t filter_index) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(filter_index, table->active_filters().size(), "filter_index"); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  std::size_t position = 0;
  table->erase_active_filters_if([&](const formulon::pivot::PivotFilter&) { return position++ == filter_index; });
  return edit.commit();
}
