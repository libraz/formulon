#pragma once

#include <cmath>
#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"

// Positional wrappers over the C setters that take a record, for tests that
// set every field.

// Sets the defined name `name` to `formula` in scope `local_sheet_id`
// (`FM_DEFINED_NAME_SCOPE_WORKBOOK` for workbook scope), with no comment and not hidden.
inline fm_status_t SetDefinedName(fm_workbook_t* wb, const char* name, const char* formula,
                                  std::int32_t local_sheet_id = FM_DEFINED_NAME_SCOPE_WORKBOOK) {
  const fm_defined_name_t input{sizeof(fm_defined_name_t), name, formula, local_sheet_id, nullptr, 0};
  return fm_workbook_set_defined_name(wb, &input);
}

// Sets the note at (`row`, `col`) to `author` and `text`.
inline fm_status_t SetComment(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                              const char* author, const char* text) {
  fm_comment_t comment = FM_STRUCT_INIT(fm_comment_t);
  comment.author = author;
  comment.text = text;
  return fm_sheet_set_comment(wb, sheet_index, row, col, &comment);
}

// Sets the three iterative-calculation settings; an argument of -1 keeps that
// setting.
inline fm_status_t SetIterative(fm_workbook_t* wb, std::int32_t enabled, std::int32_t max_iterations,
                                double max_change) {
  fm_iterative_settings_patch_t patch = FM_STRUCT_INIT(fm_iterative_settings_patch_t);
  if (enabled != -1) {
    patch.enabled_op = FM_PATCH_OP_SET;
    patch.enabled = enabled;
  }
  if (max_iterations != -1) {
    patch.max_iterations_op = FM_PATCH_OP_SET;
    patch.max_iterations = static_cast<std::uint32_t>(max_iterations);
  }
  if (max_change != -1.0) {
    patch.max_change_op = FM_PATCH_OP_SET;
    patch.max_change = max_change;
  }
  return fm_workbook_update_iterative(wb, &patch);
}

// Creates a table from positional arguments. `header_row` and `totals_row` are
// 0 or 1 and map onto the explicit tristate values; the four style toggles
// keep their creation defaults. An empty `style_name` creates an unstyled
// table (`styled` false). A null `rect` passes a null spec.
inline fm_status_t AddTable(fm_workbook_t* wb, std::size_t sheet_index, const fm_cell_rect_t* rect, const char* name,
                            const char* const* column_names, std::size_t column_count, const char* style_name,
                            std::int32_t header_row, std::int32_t totals_row, std::size_t* out_table_index) {
  if (rect == nullptr) {
    return fm_sheet_add_table(wb, sheet_index, nullptr, out_table_index);
  }
  fm_table_spec_t spec = FM_STRUCT_INIT(fm_table_spec_t);
  spec.range = *rect;
  spec.name = name;
  spec.column_names = column_names;
  spec.column_count = static_cast<uint32_t>(column_count);
  const bool unstyled = style_name != nullptr && *style_name == '\0';
  spec.style_name = unstyled ? nullptr : style_name;
  spec.styled = unstyled ? FM_TRISTATE_FALSE : FM_TRISTATE_DEFAULT;
  spec.header_row = header_row != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  spec.totals_row = totals_row != 0 ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  return fm_sheet_add_table(wb, sheet_index, &spec, out_table_index);
}

// Patches table `table_index`'s range, leaving its header and totals rows.
inline fm_status_t UpdateTableRange(fm_workbook_t* wb, std::size_t sheet_index, std::size_t table_index,
                                    const fm_cell_rect_t& range) {
  fm_table_patch_t patch = FM_STRUCT_INIT(fm_table_patch_t);
  patch.range_op = FM_PATCH_OP_SET;
  patch.range = range;
  return fm_sheet_update_table(wb, sheet_index, table_index, &patch);
}

// Creates a pivot table named `name` over cache `cache_id`, anchored at
// (`anchor_row`, `anchor_col`).
inline fm_status_t AddPivot(fm_workbook_t* wb, std::size_t sheet_index, const char* name, std::uint32_t cache_id,
                            std::uint32_t anchor_row, std::uint32_t anchor_col, std::size_t* out_pivot_index) {
  fm_pivot_spec_t spec = FM_STRUCT_INIT(fm_pivot_spec_t);
  spec.name = name;
  spec.cache_id = cache_id;
  spec.anchor_row = anchor_row;
  spec.anchor_col = anchor_col;
  return fm_sheet_add_pivot(wb, sheet_index, &spec, out_pivot_index);
}

// Pivot field and item setters over the patch records. Each writes one member
// and leaves the rest at KEEP, with the positional shape the property setters
// had; an empty or null string where a setter accepts one means CLEAR.
namespace pivot_patch_detail {

inline fm_patch_op_t StringOp(const char* text) {
  return text == nullptr || *text == '\0' ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
}

}  // namespace pivot_patch_detail

#define FORMULON_PIVOT_FIELD_FLAG_SETTER(Name, member)                                         \
  inline fm_status_t Name(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index, \
                          std::size_t field_index, std::int32_t on) {                          \
    fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);                     \
    patch.member##_op = FM_PATCH_OP_SET;                                                       \
    patch.member = on;                                                                         \
    return fm_sheet_update_pivot_field(wb, sheet_index, pivot_index, field_index, &patch);     \
  }

FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldDefaultSubtotal, default_subtotal)
FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldInsertBlankRow, insert_blank_row)
FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldFillDownLabels, fill_down_labels)
FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldShowAll, show_all)
FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldSubtotalTop, subtotal_top)
FORMULON_PIVOT_FIELD_FLAG_SETTER(SetPivotFieldShowDetail, show_detail)

#undef FORMULON_PIVOT_FIELD_FLAG_SETTER

// Sets the sort order and the by-field key together; a null or empty
// `by_field` sorts by label.
inline fm_status_t SetPivotFieldSort(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                     std::size_t field_index, std::int32_t order, const char* by_field) {
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.sort_order_op = FM_PATCH_OP_SET;
  patch.sort_order = order;
  patch.sort_by_field_op = pivot_patch_detail::StringOp(by_field);
  patch.sort_by_field = by_field;
  return fm_sheet_update_pivot_field(wb, sheet_index, pivot_index, field_index, &patch);
}

// A null or empty `name` restores the source name.
inline fm_status_t SetPivotFieldCustomName(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                           std::size_t field_index, const char* name) {
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.custom_name_op = pivot_patch_detail::StringOp(name);
  patch.custom_name = name;
  return fm_sheet_update_pivot_field(wb, sheet_index, pivot_index, field_index, &patch);
}

// An empty `format_code` removes the number format.
inline fm_status_t SetPivotFieldNumberFormat(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                             std::size_t field_index, const char* format_code) {
  fm_pivot_field_patch_t patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  patch.number_format_op = pivot_patch_detail::StringOp(format_code);
  patch.number_format = format_code;
  return fm_sheet_update_pivot_field(wb, sheet_index, pivot_index, field_index, &patch);
}

inline fm_status_t SetPivotFieldItemVisible(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                            std::size_t field_index, std::size_t item_index, std::int32_t visible) {
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.visible_op = FM_PATCH_OP_SET;
  patch.visible = visible;
  return fm_sheet_update_pivot_field_item(wb, sheet_index, pivot_index, field_index, item_index, &patch);
}

inline fm_status_t SetPivotFieldItemShowDetail(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                               std::size_t field_index, std::size_t item_index, std::int32_t on) {
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.show_detail_op = FM_PATCH_OP_SET;
  patch.show_detail = on;
  return fm_sheet_update_pivot_field_item(wb, sheet_index, pivot_index, field_index, item_index, &patch);
}

// A null or empty `caption` removes the caption.
inline fm_status_t SetPivotFieldItemCaption(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                            std::size_t field_index, std::size_t item_index, const char* caption) {
  fm_pivot_item_patch_t patch = FM_STRUCT_INIT(fm_pivot_item_patch_t);
  patch.caption_op = pivot_patch_detail::StringOp(caption);
  patch.caption = caption;
  return fm_sheet_update_pivot_field_item(wb, sheet_index, pivot_index, field_index, item_index, &patch);
}

// Sets every member of the pivot style; a null or empty `style_name` writes
// the style element without a style name.
inline fm_status_t UpdatePivotStyleOptions(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                           const char* style_name, std::int32_t show_row_headers,
                                           std::int32_t show_column_headers, std::int32_t show_row_stripes,
                                           std::int32_t show_column_stripes, std::int32_t show_last_column) {
  fm_pivot_style_options_patch_t patch = FM_STRUCT_INIT(fm_pivot_style_options_patch_t);
  patch.style_name_op = pivot_patch_detail::StringOp(style_name);
  patch.style_name = style_name;
  patch.show_row_headers_op = FM_PATCH_OP_SET;
  patch.show_row_headers = show_row_headers;
  patch.show_column_headers_op = FM_PATCH_OP_SET;
  patch.show_column_headers = show_column_headers;
  patch.show_row_stripes_op = FM_PATCH_OP_SET;
  patch.show_row_stripes = show_row_stripes;
  patch.show_column_stripes_op = FM_PATCH_OP_SET;
  patch.show_column_stripes = show_column_stripes;
  patch.show_last_column_op = FM_PATCH_OP_SET;
  patch.show_last_column = show_last_column;
  return fm_sheet_update_pivot_style_options(wb, sheet_index, pivot_index, &patch);
}

// Groups a date field. A NaN bound is not set, so the group takes the field's
// first or last value.
inline fm_status_t SetPivotFieldDateGroup(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                          std::size_t field_index, std::uint32_t periods, std::int32_t calendar,
                                          std::uint32_t interval_days, double start_or_nan, double end_or_nan) {
  fm_pivot_date_group_spec_t group = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  group.periods = periods;
  group.calendar = calendar;
  group.interval_days = interval_days;
  group.has_start = std::isnan(start_or_nan) ? 0 : 1;
  group.start = start_or_nan;
  group.has_end = std::isnan(end_or_nan) ? 0 : 1;
  group.end = end_or_nan;
  return fm_sheet_set_pivot_field_date_group(wb, sheet_index, pivot_index, field_index, &group);
}

// Writes the cache's worksheet source; `present == 0` clears it.
inline fm_status_t SetPivotCacheWorksheetSource(fm_workbook_t* wb, std::uint32_t cache_id, std::int32_t present,
                                                const char* ref, const char* sheet, const char* name) {
  if (present == 0) {
    return fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, nullptr);
  }
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  source.ref = ref;
  source.sheet = sheet;
  source.name = name;
  return fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, &source);
}

// Reads the cache's worksheet source into positional outputs.
inline fm_status_t GetPivotCacheWorksheetSource(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                std::int32_t* out_present, const char** out_ref, const char** out_sheet,
                                                const char** out_name) {
  fm_pivot_worksheet_source_t source = FM_STRUCT_INIT(fm_pivot_worksheet_source_t);
  const fm_status_t status = fm_workbook_get_pivot_cache_worksheet_source(wb, cache_id, out_present, &source);
  if (status == FM_STATUS_OK) {
    *out_ref = source.ref;
    *out_sheet = source.sheet;
    *out_name = source.name;
  }
  return status;
}
