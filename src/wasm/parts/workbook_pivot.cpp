//
// JsWorkbook PivotTable mutators. Thin wrappers over the `fm_sheet_*_pivot_*`
// C ABI: numeric / boolean / string args map straight through; the
// spec-style mutators accept an `emscripten::val` and unpack the fields
// into the matching `fm_pivot_*_spec_t`. The std::string locals keep
// ownership for the borrowed `const char*` pointers handed to the C ABI;
// the spec is never retained beyond the call. The PivotCache side lives in
// `parts/workbook_pivot_cache.cpp`; the companion `pivotCount` /
// `pivotLayout` readers live in `parts/workbook_cells.cpp` so the
// iteration surface stays together.

#include <emscripten/val.h>

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

bool build_data_field_spec_checked(JsNarrowNumericReader& reader, const emscripten::val& spec,
                                   fm_pivot_data_field_spec_t& out, std::string& name_buf, std::string& nfmt_buf) {
  // Read the optional strings before the numeric fields. Once the reader
  // rejects a numeric value, no further access to the caller's object is
  // allowed before returning the binding error.
  out.name = reader.optional_string(spec, "name", name_buf, "pivotDataField.name");
  // `null` is an absent format, as the `listDataFields()` read gives it.
  out.number_format = reader.optional_string(spec, "numberFormat", nfmt_buf, "pivotDataField.numberFormat");
  if (!reader.ok()) {
    return false;
  }
  out.field_index = reader.required_u32(spec, "fieldIndex", "pivotDataField.fieldIndex");
  out.aggregation = static_cast<fm_pivot_aggregation_t>(
      reader.u32(spec, "aggregation", FM_PIVOT_AGGREGATION_SUM, "pivotDataField.aggregation"));
  out.show_as = static_cast<fm_pivot_show_values_as_t>(reader.u32(spec, "showAs", 0U, "pivotDataField.showAs"));
  out.show_as_base_field = reader.i32(spec, "showAsBaseField", -1, "pivotDataField.showAsBaseField");
  out.show_as_base_item = reader.i32(spec, "showAsBaseItem", -1, "pivotDataField.showAsBaseItem");
  return reader.ok();
}

/// One pivot filter operand pulled from a JS `number | string | null` slot.
struct FilterOperand {
  int32_t kind = FM_PIVOT_FILTER_OPERAND_KIND_NONE;
  int32_t as_int = 0;
  double as_double = 0.0;
  std::string text;
};

/// Reads `owner[key]`: absent or null is no operand, a Number is an INT when
/// `integral` and whole, else a DOUBLE, a String is TEXT.
FilterOperand pull_filter_operand(JsNarrowNumericReader& reader, const emscripten::val& owner, const char* key,
                                  const char* field, bool integral) {
  FilterOperand out;
  const emscripten::val value = reader.value(owner, key, field);
  if (!reader.ok() || !js_value_present(value)) {
    return out;
  }
  if (value.isNumber()) {
    const double number = value.as<double>();
    if (integral && number >= -2147483648.0 && number <= 2147483647.0 && std::floor(number) == number) {
      out.kind = FM_PIVOT_FILTER_OPERAND_KIND_INT;
      out.as_int = static_cast<int32_t>(number);
    } else {
      out.kind = FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE;
      out.as_double = number;
    }
  } else if (value.isString()) {
    out.kind = FM_PIVOT_FILTER_OPERAND_KIND_TEXT;
    out.text = reader.string_value(value, field);
  } else {
    reader.fail_type(field, "a Number or a String");
  }
  return out;
}

/// Sets `o[key]` from one operand slot of a C filter spec.
void set_filter_operand(emscripten::val& o, const char* key, int32_t kind, int32_t as_int, double as_double,
                        const char* text) {
  switch (kind) {
    case FM_PIVOT_FILTER_OPERAND_KIND_INT:
      o.set(key, as_int);
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE:
      o.set(key, as_double);
      break;
    case FM_PIVOT_FILTER_OPERAND_KIND_TEXT:
      o.set(key, std::string(text != nullptr ? text : ""));
      break;
    default:
      o.set(key, emscripten::val::null());
      break;
  }
}

}  // namespace

// ---- PivotTable --------------------------------------------------------

uint32_t JsWorkbook::pivotCreate(uint32_t sheet, const std::string& name, uint32_t cacheId, uint32_t anchorRow,
                                 uint32_t anchorCol) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  fm_pivot_spec_t spec = FM_STRUCT_INIT(fm_pivot_spec_t);
  spec.name = name.c_str();
  spec.cache_id = cacheId;
  spec.anchor_row = anchorRow;
  spec.anchor_col = anchorCol;
  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot(handle_, sheet, &spec, &out);
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::pivotRemove(uint32_t sheet, uint32_t pivotIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_pivot(handle_, sheet, pivotIndex) : kBindingInvalidHandle);
}

void JsWorkbook::pivotSetName(uint32_t sheet, uint32_t pivotIndex, const std::string& name) {
  succeeded(handle_ != nullptr ? fm_sheet_rename_pivot(handle_, sheet, pivotIndex, name.c_str())
                               : kBindingInvalidHandle);
}

void JsWorkbook::pivotSetAnchor(uint32_t sheet, uint32_t pivotIndex, emscripten::val cell) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setPivotAnchor");
  const uint32_t row = reader.required_u32(cell, "row", "anchor.row");
  const uint32_t col = reader.required_u32(cell, "col", "anchor.col");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_pivot_anchor(handle_, sheet, pivotIndex, row, col));
}

void JsWorkbook::pivotSetGrandTotals(uint32_t sheet, uint32_t pivotIndex, bool rowsEnabled, bool colsEnabled) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  fm_status_t rc =
      fm_sheet_update_pivot_grand_totals(handle_, sheet, pivotIndex, rowsEnabled ? 1 : 0, colsEnabled ? 1 : 0);
  succeeded(rc);
}

void JsWorkbook::pivotUpdateStyleOptions(uint32_t sheet, uint32_t pivotIndex, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePivotStyleOptions");
  fm_pivot_style_options_patch_t p = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  std::string style_name;
  // `null` writes the element without a style name; a string replaces it; an omitted key keeps it.
  const emscripten::val name_value = reader.value(patch, "styleName", "patch.styleName");
  switch (js_patch_key(reader, name_value, "patch.styleName", true)) {
    case JsPatchKey::kClear:
      p.style_name_op = FM_CHART_PATCH_OP_CLEAR;
      break;
    case JsPatchKey::kSet:
      p.style_name_op = FM_CHART_PATCH_OP_SET;
      p.style_name = reader.optional_string_value(name_value, style_name, "patch.styleName");
      break;
    case JsPatchKey::kKeep:
      break;
  }
  js_patch_flag(reader, patch, "showRowHeaders", "patch.showRowHeaders", &p.show_row_headers_op, &p.show_row_headers);
  js_patch_flag(reader, patch, "showColumnHeaders", "patch.showColumnHeaders", &p.show_column_headers_op,
                &p.show_column_headers);
  js_patch_flag(reader, patch, "showRowStripes", "patch.showRowStripes", &p.show_row_stripes_op, &p.show_row_stripes);
  js_patch_flag(reader, patch, "showColumnStripes", "patch.showColumnStripes", &p.show_column_stripes_op,
                &p.show_column_stripes);
  js_patch_flag(reader, patch, "showLastColumn", "patch.showLastColumn", &p.show_last_column_op, &p.show_last_column);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_pivot_style_options(handle_, sheet, pivotIndex, &p));
}

double JsWorkbook::pivotGetLayout(uint32_t sheet, uint32_t pivotIndex) const {
  std::int32_t layout = FM_PIVOT_REPORT_LAYOUT_COMPACT;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_report_layout(handle_, sheet, pivotIndex, &layout)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(layout);
}

void JsWorkbook::pivotSetLayout(uint32_t sheet, uint32_t pivotIndex, uint32_t layout) {
  succeeded(handle_ != nullptr
                ? fm_sheet_set_pivot_report_layout(handle_, sheet, pivotIndex, static_cast<std::int32_t>(layout))
                : kBindingInvalidHandle);
}

uint32_t JsWorkbook::pivotCacheId(uint32_t sheet, uint32_t pivotIndex) const {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  uint32_t out = 0;
  const fm_status_t rc = fm_sheet_get_pivot_cache_id(handle_, sheet, pivotIndex, &out);
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

std::string JsWorkbook::pivotName(uint32_t sheet, uint32_t pivotIndex) const {
  const char* name = nullptr;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_name(handle_, sheet, pivotIndex, &name) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return std::string();
  }
  return string_from_cstr(name);
}

uint32_t JsWorkbook::pivotRuntimeId(uint32_t sheet, uint32_t pivotIndex) const {
  uint32_t out = 0;
  succeeded(handle_ != nullptr ? fm_sheet_get_pivot_runtime_id(handle_, sheet, pivotIndex, &out)
                               : kBindingInvalidHandle);
  return out;
}

uint32_t JsWorkbook::resolvePivotRuntimeId(uint32_t sheet, uint32_t id) const {
  std::size_t out = 0;
  succeeded(handle_ != nullptr ? fm_sheet_resolve_pivot_runtime_id(handle_, sheet, id, &out) : kBindingInvalidHandle);
  return static_cast<uint32_t>(out);
}

emscripten::val JsWorkbook::pivotGetGrandTotals(uint32_t sheet, uint32_t pivotIndex) const {
  int32_t rows = 0;
  int32_t cols = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_grand_totals(handle_, sheet, pivotIndex, &rows, &cols)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("rows", rows != 0);
  o.set("columns", cols != 0);
  return o;
}

emscripten::val JsWorkbook::pivotGetStyleOptions(uint32_t sheet, uint32_t pivotIndex) const {
  fm_pivot_style_options_t info = FM_STRUCT_INIT(fm_pivot_style_options_t);
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_style_options(handle_, sheet, pivotIndex, &info) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  if (info.present == 0) {
    return emscripten::val::null();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null(o, "styleName", info.style_name);
  o.set("showRowHeaders", info.show_row_headers != 0);
  o.set("showColumnHeaders", info.show_column_headers != 0);
  o.set("showRowStripes", info.show_row_stripes != 0);
  o.set("showColumnStripes", info.show_column_stripes != 0);
  o.set("showLastColumn", info.show_last_column != 0);
  return o;
}

double JsWorkbook::pivotFieldCount(uint32_t sheet, uint32_t pivotIndex) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_field_count(handle_, sheet, pivotIndex, &count) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

uint32_t JsWorkbook::pivotFieldAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  std::string source_name;
  std::string custom_name;
  std::string number_format;
  JsNarrowNumericReader reader("addPivotField");
  // Every nested read uses this per-call reader so getter/proxy failures and
  // lossy integer conversions are reported before the C ABI is called.
  const char* source_name_ptr = reader.optional_string(spec, "sourceName", source_name, "pivotField.sourceName");
  const char* custom_name_ptr = reader.optional_string(spec, "customName", custom_name, "pivotField.customName");
  const emscripten::val subtotal_top = reader.value(spec, "subtotalTop", "pivotField.subtotalTop");
  const char* number_format_ptr =
      reader.optional_string(spec, "numberFormat", number_format, "pivotField.numberFormat");

  fm_pivot_field_spec_t c_spec = FM_STRUCT_INIT(fm_pivot_field_spec_t);
  // `sourceName` is required; an omitted key passes NULL through rather
  // than `js_pull_string`'s empty-string default, which the C ABI's own
  // required-field check does not treat the same as a missing argument.
  c_spec.source_name = source_name_ptr;
  c_spec.custom_name = custom_name_ptr;
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.required_u32(spec, "axis", "pivotField.axis"));
  // An omitted `subtotalTop` takes the engine default (the group head).
  c_spec.subtotal_top = !js_value_present(subtotal_top)
                            ? FM_TRISTATE_DEFAULT
                            : (reader.boolean_value(subtotal_top, true) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE);
  c_spec.number_format = number_format_ptr;

  if (!reader.ok()) {
    return 0U;
  }

  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_field(handle_, sheet, pivotIndex, &c_spec, &out);
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

emscripten::val JsWorkbook::pivotFieldAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const {
  fm_pivot_field_info_t info = FM_STRUCT_INIT(fm_pivot_field_info_t);
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_field(handle_, sheet, pivotIndex, fieldIndex, &info)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "sourceName", info.source_name);
  js_set_cstr_or_null(o, "customName", info.custom_name);
  js_set_index_or_null(o, "axis", info.axis);
  o.set("subtotalTop", info.subtotal_top != 0);
  js_set_cstr_or_null(o, "numberFormat", info.number_format);
  js_set_index_or_null(o, "cacheFieldIndex", info.cache_field_index);
  o.set("sortOrder", info.sort_order);
  js_set_cstr_or_null(o, "sortByField",
                      info.sort_by_field != nullptr && info.sort_by_field[0] != '\0' ? info.sort_by_field : nullptr);
  o.set("datePeriods", info.date_periods);
  o.set("dateCalendar", info.date_calendar);
  js_set_index_or_null(o, "groupBaseIndex", info.group_base_index);
  o.set("defaultSubtotal", info.default_subtotal != 0);
  o.set("insertBlankRow", info.insert_blank_row != 0);
  o.set("fillDownLabels", info.fill_down_labels != 0);
  o.set("showAll", info.show_all != 0);
  o.set("showDetail", info.show_detail != 0);
  emscripten::val fns = emscripten::val::array();
  for (std::uint32_t agg = 0; agg < 32U; ++agg) {
    if ((info.subtotal_functions & (1U << agg)) != 0U) {
      fns.call<void>("push", agg);
    }
  }
  o.set("subtotalFunctions", fns);
  return o;
}

double JsWorkbook::pivotFieldItemCount(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_field_item_count(handle_, sheet, pivotIndex, fieldIndex, &count)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

emscripten::val JsWorkbook::pivotFieldItemAt(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex,
                                             uint32_t itemIndex) const {
  fm_pivot_item_info_t info = FM_STRUCT_INIT(fm_pivot_item_info_t);
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_field_item(handle_, sheet, pivotIndex, fieldIndex, itemIndex, &info)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr_or_null(o, "name", info.name);
  js_set_index_or_null(o, "cacheIndex", info.cache_index);
  o.set("visible", info.visible != 0);
  if (info.caption != nullptr) {
    js_set_cstr(o, "caption", info.caption);
  } else {
    o.set("caption", emscripten::val::null());
  }
  o.set("showDetail", info.show_detail != 0);
  return o;
}

void JsWorkbook::pivotFieldUpdate(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePivotField");
  fm_pivot_field_patch_t p = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  std::string custom_name;
  std::string number_format;
  std::string sort_by_field;
  js_patch_string(reader, patch, "customName", "patch.customName", true, custom_name, &p.custom_name_op,
                  &p.custom_name);
  js_patch_string(reader, patch, "numberFormat", "patch.numberFormat", true, number_format, &p.number_format_op,
                  &p.number_format);
  js_patch_flag(reader, patch, "defaultSubtotal", "patch.defaultSubtotal", &p.default_subtotal_op, &p.default_subtotal);
  js_patch_flag(reader, patch, "insertBlankRow", "patch.insertBlankRow", &p.insert_blank_row_op, &p.insert_blank_row);
  js_patch_flag(reader, patch, "fillDownLabels", "patch.fillDownLabels", &p.fill_down_labels_op, &p.fill_down_labels);
  js_patch_flag(reader, patch, "showAll", "patch.showAll", &p.show_all_op, &p.show_all);
  js_patch_flag(reader, patch, "subtotalTop", "patch.subtotalTop", &p.subtotal_top_op, &p.subtotal_top);
  js_patch_int(reader, patch, "sortOrder", "patch.sortOrder", false, &p.sort_order_op, &p.sort_order);
  js_patch_string(reader, patch, "sortByField", "patch.sortByField", true, sort_by_field, &p.sort_by_field_op,
                  &p.sort_by_field);
  js_patch_flag(reader, patch, "showDetail", "patch.showDetail", &p.show_detail_op, &p.show_detail);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_pivot_field(handle_, sheet, pivotIndex, fieldIndex, &p));
}

void JsWorkbook::pivotFieldItemUpdate(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t itemIndex,
                                      emscripten::val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updatePivotFieldItem");
  fm_pivot_item_patch_t p = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  std::string caption;
  js_patch_flag(reader, patch, "visible", "patch.visible", &p.visible_op, &p.visible);
  js_patch_flag(reader, patch, "showDetail", "patch.showDetail", &p.show_detail_op, &p.show_detail);
  js_patch_string(reader, patch, "caption", "patch.caption", true, caption, &p.caption_op, &p.caption);
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_update_pivot_field_item(handle_, sheet, pivotIndex, fieldIndex, itemIndex, &p));
}

void JsWorkbook::pivotFieldClear(uint32_t sheet, uint32_t pivotIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_fields(handle_, sheet, pivotIndex) : kBindingInvalidHandle);
}

void JsWorkbook::pivotFieldSetAxis(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, int32_t axis) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  fm_status_t rc =
      fm_sheet_set_pivot_field_axis(handle_, sheet, pivotIndex, fieldIndex, static_cast<std::int32_t>(axis));
  succeeded(rc);
}

emscripten::val JsWorkbook::pivotGetOptions(uint32_t sheet, uint32_t pivotIndex) const {
  fm_pivot_options_t opt = FM_STRUCT_INIT(fm_pivot_options_t);
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_options(handle_, sheet, pivotIndex, &opt) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("showError", opt.show_error != 0);
  js_set_cstr_or_null(o, "errorCaption", opt.error_caption);
  o.set("showMissing", opt.show_missing != 0);
  js_set_cstr_or_null(o, "missingCaption", opt.missing_caption);
  o.set("showHeaders", opt.show_headers != 0);
  o.set("showDrill", opt.show_drill != 0);
  o.set("showDataTips", opt.show_data_tips != 0);
  o.set("gridDropZones", opt.grid_drop_zones != 0);
  o.set("showEmptyRow", opt.show_empty_row != 0);
  o.set("showEmptyColumn", opt.show_empty_column != 0);
  o.set("showDataDropDown", opt.show_data_drop_down != 0);
  o.set("printDrill", opt.print_drill != 0);
  o.set("fieldPrintTitles", opt.field_print_titles != 0);
  o.set("itemPrintTitles", opt.item_print_titles != 0);
  o.set("mergeItem", opt.merge_item != 0);
  o.set("indent", opt.indent);
  o.set("pageOverThenDown", opt.page_over_then_down != 0);
  o.set("pageWrap", opt.page_wrap);
  o.set("multipleFieldFilters", opt.multiple_field_filters != 0);
  o.set("preserveFormatting", opt.preserve_formatting != 0);
  o.set("customListSort", opt.custom_list_sort != 0);
  js_set_cstr_or_null(o, "altText", opt.alt_text);
  js_set_cstr_or_null(o, "altTextSummary", opt.alt_text_summary);
  js_set_cstr_or_null(o, "grandTotalCaption", opt.grand_total_caption);
  o.set("autoFormatting", opt.auto_formatting != 0);
  o.set("enableDrill", opt.enable_drill != 0);
  return o;
}

void JsWorkbook::pivotSetOptions(uint32_t sheet, uint32_t pivotIndex, emscripten::val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  if (!js_value_present(options)) {
    record_type_error("setPivotOptions: `options` must be an object");
    return;
  }
  // The C call writes every member, so members the caller omits keep the
  // pivot's current value.
  fm_pivot_options_t opt = FM_STRUCT_INIT(fm_pivot_options_t);
  if (!succeeded(fm_sheet_get_pivot_options(handle_, sheet, pivotIndex, &opt))) {
    return;
  }
  std::string error_caption = string_from_cstr(opt.error_caption);
  std::string missing_caption = string_from_cstr(opt.missing_caption);
  std::string alt_text = string_from_cstr(opt.alt_text);
  std::string alt_summary = string_from_cstr(opt.alt_text_summary);
  std::string grand_total = string_from_cstr(opt.grand_total_caption);

  JsNarrowNumericReader reader("setPivotOptions");
  const auto flag = [&](const char* key, int32_t current) {
    const std::string field = std::string("pivotOptions.") + key;
    return reader.boolean(options, key, current != 0, field.c_str()) ? 1 : 0;
  };
  // An omitted caption keeps the stored one; `null` removes it.
  const auto text = [&](const char* key, std::string& storage) {
    const std::string field = std::string("pivotOptions.") + key;
    const emscripten::val value = reader.value(options, key, field.c_str());
    if (!reader.ok() || value.isUndefined()) {
      return;
    }
    storage = value.isNull() ? std::string() : reader.string_value(value, field.c_str());
  };
  opt.show_error = flag("showError", opt.show_error);
  text("errorCaption", error_caption);
  opt.show_missing = flag("showMissing", opt.show_missing);
  text("missingCaption", missing_caption);
  opt.show_headers = flag("showHeaders", opt.show_headers);
  opt.show_drill = flag("showDrill", opt.show_drill);
  opt.show_data_tips = flag("showDataTips", opt.show_data_tips);
  opt.grid_drop_zones = flag("gridDropZones", opt.grid_drop_zones);
  opt.show_empty_row = flag("showEmptyRow", opt.show_empty_row);
  opt.show_empty_column = flag("showEmptyColumn", opt.show_empty_column);
  opt.show_data_drop_down = flag("showDataDropDown", opt.show_data_drop_down);
  opt.print_drill = flag("printDrill", opt.print_drill);
  opt.field_print_titles = flag("fieldPrintTitles", opt.field_print_titles);
  opt.item_print_titles = flag("itemPrintTitles", opt.item_print_titles);
  opt.merge_item = flag("mergeItem", opt.merge_item);
  opt.indent = reader.u32(options, "indent", opt.indent, "pivotOptions.indent");
  opt.page_over_then_down = flag("pageOverThenDown", opt.page_over_then_down);
  opt.page_wrap = reader.u32(options, "pageWrap", opt.page_wrap, "pivotOptions.pageWrap");
  opt.multiple_field_filters = flag("multipleFieldFilters", opt.multiple_field_filters);
  opt.preserve_formatting = flag("preserveFormatting", opt.preserve_formatting);
  opt.custom_list_sort = flag("customListSort", opt.custom_list_sort);
  text("altText", alt_text);
  text("altTextSummary", alt_summary);
  text("grandTotalCaption", grand_total);
  opt.auto_formatting = flag("autoFormatting", opt.auto_formatting);
  opt.enable_drill = flag("enableDrill", opt.enable_drill);
  if (!reader.ok()) {
    return;
  }
  const auto view = [](const std::string& s) { return s.empty() ? nullptr : s.c_str(); };
  opt.error_caption = view(error_caption);
  opt.missing_caption = view(missing_caption);
  opt.alt_text = view(alt_text);
  opt.alt_text_summary = view(alt_summary);
  opt.grand_total_caption = view(grand_total);
  succeeded(fm_sheet_update_pivot_options(handle_, sheet, pivotIndex, &opt));
}

std::string JsWorkbook::pivotDataFieldDefaultName(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex,
                                                  int32_t aggregation, std::size_t excludeDataField) const {
  const char* name = nullptr;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_data_field_default_name(handle_, sheet, pivotIndex, fieldIndex,
                                                                          aggregation, excludeDataField, &name)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return std::string();
  }
  return string_from_cstr(name);
}

void JsWorkbook::pivotProject(uint32_t sheet, uint32_t pivotIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_project_pivot(handle_, sheet, pivotIndex) : kBindingInvalidHandle);
}

bool JsWorkbook::pivotProjectionPending(uint32_t sheet, uint32_t pivotIndex) const {
  int32_t pending = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_is_pivot_projection_pending(handle_, sheet, pivotIndex, &pending)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return false;
  }
  return pending != 0;
}

uint32_t JsWorkbook::pivotShowDetails(uint32_t sheet, uint32_t pivotIndex, uint32_t row, uint32_t col) {
  std::size_t newSheet = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_show_pivot_details(handle_, sheet, pivotIndex, row, col, &newSheet)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(newSheet);
}

uint32_t JsWorkbook::pivotFieldAddItem(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, emscripten::val item,
                                       bool visible) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  // `item` is a new item's name, or `{ cacheIndex }` for the cache's shared item at that position.
  std::size_t out = 0;
  fm_status_t rc = 0;
  if (item.typeOf().strictlyEquals(emscripten::val("object"))) {
    JsNarrowNumericReader reader("addPivotFieldItem");
    const uint32_t cacheIndex = reader.required_u32(item, "cacheIndex", "item.cacheIndex");
    if (!reader.ok()) {
      return 0U;
    }
    rc = fm_sheet_add_pivot_field_item_by_cache_index(handle_, sheet, pivotIndex, fieldIndex, cacheIndex,
                                                      visible ? 1 : 0, &out);
  } else {
    JsNarrowNumericReader reader("addPivotFieldItem");
    const std::string name = reader.string_value(item, "item");
    if (!reader.ok()) {
      return 0U;
    }
    rc = fm_sheet_add_pivot_field_item(handle_, sheet, pivotIndex, fieldIndex, name.c_str(), visible ? 1 : 0, &out);
  }
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

void JsWorkbook::pivotFieldClearItems(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_field_items(handle_, sheet, pivotIndex, fieldIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::pivotFieldAddSubtotalFn(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t agg) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  fm_status_t rc = fm_sheet_add_pivot_field_subtotal_function(handle_, sheet, pivotIndex, fieldIndex,
                                                              static_cast<std::int32_t>(agg));
  succeeded(rc);
}

void JsWorkbook::pivotFieldClearSubtotalFns(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_field_subtotal_functions(handle_, sheet, pivotIndex, fieldIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::pivotFieldSetDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex, uint32_t periods,
                                        uint32_t calendar, uint32_t intervalDays, double start, double end) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // A NaN bound is an absent one: the group takes the field's first / last value.
  fm_pivot_date_group_spec_t group = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  group.periods = periods;
  group.calendar = static_cast<std::int32_t>(calendar);
  group.interval_days = intervalDays;
  group.has_start = std::isnan(start) ? 0 : 1;
  group.start = std::isnan(start) ? 0.0 : start;
  group.has_end = std::isnan(end) ? 0 : 1;
  group.end = std::isnan(end) ? 0.0 : end;
  succeeded(fm_sheet_set_pivot_field_date_group(handle_, sheet, pivotIndex, fieldIndex, &group));
}

void JsWorkbook::pivotFieldClearDateGroup(uint32_t sheet, uint32_t pivotIndex, uint32_t fieldIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_field_date_group(handle_, sheet, pivotIndex, fieldIndex)
                               : kBindingInvalidHandle);
}

namespace {

// The JS spelling of `FM_PIVOT_VALUES_FIELD` (`PIVOT_VALUES_FIELD`).
constexpr int32_t kPivotValuesFieldJs = -2;

// Shared body of `pivotSetRowFieldOrder` / `pivotSetColFieldOrder`.
using PivotFieldOrderFn = fm_status_t (*)(fm_workbook_t*, size_t, size_t, const uint32_t*, size_t);

void set_pivot_field_order(fm_workbook_t* wb, uint32_t sheet, uint32_t pivotIndex, const emscripten::val& indices,
                           PivotFieldOrderFn fn) {
  if (wb == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("pivotFieldOrder");
  if (!reader.is_array(indices, "pivotFieldOrder")) {
    record_type_error("pivot field order must be an array");
    return;
  }
  const uint32_t length = reader.length(indices, "pivotFieldOrder");
  if (!reader.ok()) {
    return;
  }
  std::vector<uint32_t> v;
  v.reserve(length);
  for (uint32_t i = 0; i < length && reader.ok(); ++i) {
    // Read the raw element once through the guarded property reader, then
    // let the checked conversion reject a symbol, fraction, non-finite value,
    // or uint32 overflow. No subsequent JS reads occur after failure.
    const std::string key = std::to_string(i);
    const emscripten::val value = reader.value(indices, key.c_str(), "pivotFieldOrder.indices");
    // -2 is the Values level; any other negative reaches the C ABI as an
    // out-of-range index and is rejected there.
    const int32_t index = reader.i32_value(value, 0, "pivotFieldOrder.indices");
    v.push_back(index == kPivotValuesFieldJs ? FM_PIVOT_VALUES_FIELD : static_cast<uint32_t>(index));
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fn(wb, sheet, pivotIndex, v.empty() ? nullptr : v.data(), v.size()));
}

}  // namespace

void JsWorkbook::pivotSetRowFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIndex, indices, &fm_sheet_set_pivot_row_field_order);
}

void JsWorkbook::pivotSetColFieldOrder(uint32_t sheet, uint32_t pivotIndex, emscripten::val indices) {
  return set_pivot_field_order(handle_, sheet, pivotIndex, indices, &fm_sheet_set_pivot_column_field_order);
}

namespace {

// Shared body of `pivotRowFieldOrder` / `pivotColFieldOrder`.
using PivotFieldOrderGetter = fm_status_t (*)(const fm_workbook_t*, size_t, size_t, const uint32_t**, size_t*);

emscripten::val get_pivot_field_order(const fm_workbook_t* wb, uint32_t sheet, uint32_t pivotIndex,
                                      PivotFieldOrderGetter fn) {
  emscripten::val indices = emscripten::val::array();
  const uint32_t* data = nullptr;
  std::size_t count = 0;
  if (!succeeded(wb != nullptr ? fn(wb, sheet, pivotIndex, &data, &count) : kBindingInvalidHandle)) {
    return emscripten::val::undefined();
  }
  for (std::size_t i = 0; i < count; ++i) {
    indices.call<void>("push", data[i] == FM_PIVOT_VALUES_FIELD ? kPivotValuesFieldJs : static_cast<int32_t>(data[i]));
  }
  return indices;
}

}  // namespace

emscripten::val JsWorkbook::pivotRowFieldOrder(uint32_t sheet, uint32_t pivotIndex) const {
  return get_pivot_field_order(handle_, sheet, pivotIndex, &fm_sheet_get_pivot_row_field_order);
}

emscripten::val JsWorkbook::pivotColFieldOrder(uint32_t sheet, uint32_t pivotIndex) const {
  return get_pivot_field_order(handle_, sheet, pivotIndex, &fm_sheet_get_pivot_column_field_order);
}

double JsWorkbook::pivotDataFieldCount(uint32_t sheet, uint32_t pivotIndex) const {
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_data_field_count(handle_, sheet, pivotIndex, &count)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

uint32_t JsWorkbook::pivotDataFieldAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  fm_pivot_data_field_spec_t c_spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  std::string name_buf;
  std::string nfmt_buf;
  JsNarrowNumericReader reader("addPivotDataField");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf)) {
    return 0U;
  }
  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_data_field(handle_, sheet, pivotIndex, &c_spec, &out);
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

emscripten::val JsWorkbook::pivotDataFieldAt(uint32_t sheet, uint32_t pivotIndex, uint32_t dataFieldIndex) const {
  fm_pivot_data_field_spec_t spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_pivot_data_field(handle_, sheet, pivotIndex, dataFieldIndex, &spec)
                             : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  js_set_cstr(o, "name", spec.name);
  o.set("fieldIndex", spec.field_index);
  o.set("aggregation", static_cast<int32_t>(spec.aggregation));
  js_set_cstr_or_null(o, "numberFormat", spec.number_format);
  o.set("showAs", static_cast<int32_t>(spec.show_as));
  js_set_index_or_null(o, "showAsBaseField", spec.show_as_base_field);
  js_set_index_or_null(o, "showAsBaseItem", spec.show_as_base_item);
  return o;
}

void JsWorkbook::pivotDataFieldClear(uint32_t sheet, uint32_t pivotIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_data_fields(handle_, sheet, pivotIndex) : kBindingInvalidHandle);
}

void JsWorkbook::pivotDataFieldSet(uint32_t sheet, uint32_t pivotIndex, uint32_t dataFieldIndex, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  fm_pivot_data_field_spec_t c_spec = FM_STRUCT_INIT(fm_pivot_data_field_spec_t);
  std::string name_buf;
  std::string nfmt_buf;
  JsNarrowNumericReader reader("setPivotDataField");
  if (!build_data_field_spec_checked(reader, spec, c_spec, name_buf, nfmt_buf)) {
    return;
  }
  fm_status_t rc = fm_sheet_set_pivot_data_field(handle_, sheet, pivotIndex, dataFieldIndex, &c_spec);
  succeeded(rc);
}

double JsWorkbook::pivotFilterCount(uint32_t sheet, uint32_t pivotIndex) const {
  std::size_t count = 0;
  const fm_status_t rc =
      handle_ != nullptr ? fm_sheet_get_pivot_filter_count(handle_, sheet, pivotIndex, &count) : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return 0.0;
  }
  return static_cast<double>(count);
}

uint32_t JsWorkbook::pivotFilterAdd(uint32_t sheet, uint32_t pivotIndex, emscripten::val spec) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  // `fieldName` is required; an omitted key passes NULL through rather
  // than `js_pull_string`'s empty-string default, which the C ABI's own
  // required-field check does not treat the same as a missing argument.
  std::string field_name;
  JsNarrowNumericReader reader("addPivotFilter");
  const char* field_name_ptr = reader.optional_string(spec, "fieldName", field_name, "pivotFilter.fieldName");

  fm_pivot_filter_spec_t c_spec = FM_STRUCT_INIT(fm_pivot_filter_spec_t);
  c_spec.axis = static_cast<fm_pivot_axis_t>(reader.required_u32(spec, "axis", "pivotFilter.axis"));
  c_spec.field_name = field_name_ptr;
  c_spec.type = static_cast<fm_pivot_filter_kind_t>(reader.required_u32(spec, "type", "pivotFilter.type"));
  // A count filter takes a whole number; every other numeric operand is a double.
  const bool count_type =
      c_spec.type == FM_PIVOT_FILTER_KIND_VALUE_TOP_10 || c_spec.type == FM_PIVOT_FILTER_KIND_VALUE_BOTTOM_COUNT;
  const FilterOperand value = pull_filter_operand(reader, spec, "value", "pivotFilter.value", count_type);
  const FilterOperand value_high = pull_filter_operand(reader, spec, "valueHigh", "pivotFilter.valueHigh", false);
  c_spec.value_kind = value.kind;
  c_spec.value_int = value.as_int;
  c_spec.value_double = value.as_double;
  c_spec.value_text = value.kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT ? value.text.c_str() : nullptr;
  c_spec.value_high_kind = value_high.kind;
  c_spec.value_high_int = value_high.as_int;
  c_spec.value_high_double = value_high.as_double;
  c_spec.value_high_text = value_high.kind == FM_PIVOT_FILTER_OPERAND_KIND_TEXT ? value_high.text.c_str() : nullptr;
  c_spec.data_field_index = reader.u32(spec, "dataFieldIndex", 0U, "pivotFilter.dataFieldIndex");

  if (!reader.ok()) {
    return 0U;
  }

  std::size_t out = 0;
  fm_status_t rc = fm_sheet_add_pivot_filter(handle_, sheet, pivotIndex, &c_spec, &out);
  if (!succeeded(rc)) {
    return 0U;
  }
  return static_cast<uint32_t>(out);
}

emscripten::val JsWorkbook::pivotFilterAt(uint32_t sheet, uint32_t pivotIndex, uint32_t filterIndex) const {
  fm_pivot_filter_spec_t spec = FM_STRUCT_INIT(fm_pivot_filter_spec_t);
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_pivot_filter(handle_, sheet, pivotIndex, filterIndex, &spec)
                                            : kBindingInvalidHandle;
  if (!succeeded(rc)) {
    return emscripten::val::undefined();
  }
  emscripten::val o = emscripten::val::object();
  o.set("axis", static_cast<int32_t>(spec.axis));
  js_set_cstr(o, "fieldName", spec.field_name);
  o.set("type", static_cast<int32_t>(spec.type));
  o.set("dataFieldIndex", spec.data_field_index);
  set_filter_operand(o, "value", spec.value_kind, spec.value_int, spec.value_double, spec.value_text);
  set_filter_operand(o, "valueHigh", spec.value_high_kind, spec.value_high_int, spec.value_high_double,
                     spec.value_high_text);
  return o;
}

void JsWorkbook::pivotFilterClear(uint32_t sheet, uint32_t pivotIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_clear_pivot_filters(handle_, sheet, pivotIndex) : kBindingInvalidHandle);
}

void JsWorkbook::pivotFilterRemoveAt(uint32_t sheet, uint32_t pivotIndex, uint32_t filterIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_pivot_filter(handle_, sheet, pivotIndex, filterIndex)
                               : kBindingInvalidHandle);
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
