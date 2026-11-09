//
// C ABI - PivotTable mutation surface (tables, field order, data fields).
// Fields and filters live in `pivot_field.cpp` / `pivot_filter.cpp`; shared
// helpers come from `pivot_internal.h`.

#include "pivot/pivot_table.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <tuple>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "c_api/parts/table_style_defaults.h"
#include "pivot/calculated.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_locale.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/utf8_length.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_patch_op;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kMaxPivotNameUnits;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::null_field;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::pivot_agg_from_fm;
using formulon::c_api::parts::pivot_count;
using formulon::c_api::parts::pivot_num_fmt_code;
using formulon::c_api::parts::pivot_num_fmt_id_text;
using formulon::c_api::parts::pivot_show_as_from_fm;
using formulon::c_api::parts::PivotEdit;
using formulon::c_api::parts::project_pivot;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::resolve_cache;
using formulon::c_api::parts::resolve_pivot_checked;
using formulon::c_api::parts::resolve_pivot_mut;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::set_values_position;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::text_or_null;
using formulon::c_api::parts::validate;

namespace {

// Largest `indent` and `pageWrap` the options dialog accepts.
constexpr std::uint32_t kMaxIndent = 126U;
constexpr std::uint32_t kMaxPageWrap = 255U;

// A NULL or empty caption is no caption.
std::optional<std::string> optional_text(const char* text) {
  return text != nullptr && text[0] != '\0' ? std::optional<std::string>(text) : std::nullopt;
}

fm_pivot_report_layout_t pivot_layout_to_fm(formulon::pivot::PivotLayout layout) {
  switch (layout) {
    case formulon::pivot::PivotLayout::Tabular:
      return FM_PIVOT_REPORT_LAYOUT_TABULAR;
    case formulon::pivot::PivotLayout::Outline:
      return FM_PIVOT_REPORT_LAYOUT_OUTLINE;
    case formulon::pivot::PivotLayout::Compact:
    default:
      return FM_PIVOT_REPORT_LAYOUT_COMPACT;
  }
}

formulon::pivot::PivotLayout pivot_layout_from_fm(std::int32_t layout) {
  switch (layout) {
    case FM_PIVOT_REPORT_LAYOUT_TABULAR:
      return formulon::pivot::PivotLayout::Tabular;
    case FM_PIVOT_REPORT_LAYOUT_OUTLINE:
      return formulon::pivot::PivotLayout::Outline;
    case FM_PIVOT_REPORT_LAYOUT_COMPACT:
    default:
      return formulon::pivot::PivotLayout::Compact;
  }
}

// The default style the desktop spreadsheet gives a new PivotTable.
formulon::pivot::PivotStyleInfo new_pivot_style(const formulon::Workbook& book) {
  formulon::pivot::PivotStyleInfo style;
  style.name = std::string(formulon::c_api::parts::default_pivot_style_name(book));
  style.show_row_headers = true;
  style.show_col_headers = true;
  style.show_last_column = true;
  return style;
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotTable lifecycle / metadata
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_pivot(fm_workbook_t* wb, std::size_t sheet_index, const fm_pivot_spec_t* spec,
                                          std::size_t* out_pivot_index) {
  clear_last_error();
  if (spec == nullptr || out_pivot_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*spec); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (spec->name == nullptr) {
    return null_argument("name");
  }
  if (auto rc = check_utf8({spec->name}); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (auto rc = resolve_cache(wb, spec->cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(spec->anchor_row, spec->anchor_col); rc != 0) {
    return rc;
  }
  auto table = std::make_unique<formulon::pivot::PivotTable>();
  table->set_name(spec->name);
  table->set_pivot_cache_id(spec->cache_id);
  // Default span of 1x1 anchored at the requested cell; callers can
  // adjust via `fm_sheet_set_pivot_anchor`.
  table->set_anchor(spec->anchor_row, spec->anchor_col, 1U, 1U);
  // Excel stores the Values caption in the creating UI's wording and keeps it
  // when the file is opened under another locale.
  table->set_data_caption(formulon::pivot::pivot_layout_options_for(wb->workbook().excel_profile()).values_label);
  table->set_style_info(new_pivot_style(wb->workbook()));
  table->apply_new_pivot_defaults();
  // The engine refuses a taken or empty name and gives the pivot one field per cache field.
  auto added = wb->workbook().add_pivot(sheet_index, std::move(table));
  if (!added) {
    return set_last_error(added.error());
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  const std::size_t pivot_index = added.value();
  // A report that cannot be placed is not created.
  if (auto projected = project_pivot(wb->workbook(), sheet_index, pivot_index); !projected) {
    auto& pivots = sheet.mutable_pivot_tables();
    pivots.erase(pivots.begin() + static_cast<std::ptrdiff_t>(pivot_index));
    return set_last_error(projected.error());
  }
  *out_pivot_index = pivot_index;
  wb->workbook().mark_all_formulas_dirty();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_pivot(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  formulon::Sheet& sheet = wb->workbook().sheet(sheet_index);
  auto& pivots = sheet.mutable_pivot_tables();
  if (auto rc = check_index(pivot_index, pivots.size(), "pivot_index"); rc != 0) {
    return rc;
  }
  // Deleting a report clears the cells and merges it wrote.
  if (auto cleared = wb->workbook().clear_pivot_projection(sheet_index, pivot_index); !cleared) {
    return set_last_error(cleared.error());
  }
  pivots.erase(pivots.begin() + static_cast<std::ptrdiff_t>(pivot_index));
  wb->workbook().mark_all_formulas_dirty();
  return 0;
}

extern "C" fm_status_t fm_sheet_rename_pivot(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                             const char* utf8_name) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({utf8_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  if (auto renamed = wb->workbook().rename_pivot(sheet_index, pivot_index, utf8_name); !renamed) {
    return set_last_error(renamed.error());
  }
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_set_pivot_anchor(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                 std::uint32_t anchor_row, std::uint32_t anchor_col) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  // Reject before `set_anchor` so downstream layout never starts outside the grid.
  if (auto rc = check_cell_coord(anchor_row, anchor_col); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->set_anchor(anchor_row, anchor_col, 1U, 1U);
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_update_pivot_grand_totals(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index,
                                                          const fm_pivot_grand_totals_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  // Neither band can be cleared; the whole patch is checked before anything changes.
  if (auto rc = check_patch_op(patch->rows_op, "rows", false); rc != 0) {
    return rc;
  }
  if (auto rc = check_patch_op(patch->columns_op, "columns", false); rc != 0) {
    return rc;
  }
  if (patch->rows_op == FM_PATCH_OP_KEEP && patch->columns_op == FM_PATCH_OP_KEEP) {
    return 0;
  }
  const bool rows = patch->rows_op == FM_PATCH_OP_SET ? patch->rows != 0 : table->grand_totals_rows();
  const bool columns = patch->columns_op == FM_PATCH_OP_SET ? patch->columns != 0 : table->grand_totals_cols();
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->set_grand_totals(rows, columns);
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_show_pivot_details(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                                                   std::uint32_t row, std::uint32_t col, std::size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_sheet_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(pivot_index, wb->workbook().sheet(sheet_index).pivot_tables().size(), "pivot_index");
      rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  formulon::Expected<std::size_t, formulon::Error> created =
      wb->workbook().show_pivot_details(sheet_index, pivot_index, row, col);
  if (!created) {
    return set_last_error(created.error());
  }
  *out_sheet_index = created.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_update_pivot_style_options(fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index,
                                                           const fm_pivot_style_options_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  // The whole patch is checked before anything changes.
  if (auto rc = check_patch_op(patch->style_name_op, "style_name", true); rc != 0) {
    return rc;
  }
  if (patch->style_name_op == FM_PATCH_OP_SET) {
    if (patch->style_name == nullptr) {
      return null_field("style_name");
    }
    if (patch->style_name[0] == '\0') {
      return invalid_argument("style_name must not be empty", "style_name=");
    }
    if (auto rc = check_utf8({patch->style_name}); rc != 0) {
      return rc;
    }
  }
  const std::tuple<fm_patch_op_t, std::int32_t, bool formulon::pivot::PivotStyleInfo::*, const char*> toggles[] = {
      {patch->show_row_headers_op, patch->show_row_headers, &formulon::pivot::PivotStyleInfo::show_row_headers,
       "show_row_headers"},
      {patch->show_column_headers_op, patch->show_column_headers, &formulon::pivot::PivotStyleInfo::show_col_headers,
       "show_column_headers"},
      {patch->show_row_stripes_op, patch->show_row_stripes, &formulon::pivot::PivotStyleInfo::show_row_stripes,
       "show_row_stripes"},
      {patch->show_column_stripes_op, patch->show_column_stripes, &formulon::pivot::PivotStyleInfo::show_col_stripes,
       "show_column_stripes"},
      {patch->show_last_column_op, patch->show_last_column, &formulon::pivot::PivotStyleInfo::show_last_column,
       "show_last_column"}};
  bool changes = patch->style_name_op != FM_PATCH_OP_KEEP;
  for (const auto& [op, value, member, field] : toggles) {
    if (auto rc = check_patch_op(op, field, false); rc != 0) {
      return rc;
    }
    if (op == FM_PATCH_OP_SET) {
      if (auto rc = check_enum_domain(value, 0, 1, field); rc != 0) {
        return rc;
      }
      changes = true;
    }
  }
  if (!changes) {
    return 0;
  }
  formulon::pivot::PivotStyleInfo info =
      table->style_info().has_value() ? *table->style_info() : new_pivot_style(wb->workbook());
  if (patch->style_name_op == FM_PATCH_OP_SET) {
    info.name = patch->style_name;
  } else if (patch->style_name_op == FM_PATCH_OP_CLEAR) {
    info.name.clear();
  }
  for (const auto& [op, value, member, field] : toggles) {
    if (op == FM_PATCH_OP_SET) {
      info.*member = value != 0;
    }
  }
  table->set_style_info(std::move(info));
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_report_layout(const fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::int32_t* out_layout) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_layout}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  *out_layout = pivot_layout_to_fm(table->layout());
  return 0;
}

extern "C" fm_status_t fm_sheet_set_pivot_report_layout(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::int32_t layout) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (layout != FM_PIVOT_REPORT_LAYOUT_COMPACT && layout != FM_PIVOT_REPORT_LAYOUT_TABULAR &&
      layout != FM_PIVOT_REPORT_LAYOUT_OUTLINE) {
    return invalid_argument(("layout " + std::to_string(static_cast<int>(layout)) + " is not a ReportLayout").c_str(),
                            "layout=" + std::to_string(static_cast<int>(layout)));
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->set_layout(pivot_layout_from_fm(layout));
  return edit.commit();
}

// PivotTable field order (row / col)
// ---------------------------------------------------------------------------

namespace {

// The axis the Values level sits on, as the layout resolves it: the rows when
// they record a position, else the columns. Meaningful only with two or more
// data fields; fewer leave no Values level at all.
bool values_on_rows(const formulon::pivot::PivotTable& table) {
  return table.row_values_position().has_value();
}

bool has_values_level(const formulon::pivot::PivotTable& table) {
  return table.data_fields().size() >= 2U;
}

// Common backing for `set_row_field_order` / `set_col_field_order`.
fm_status_t set_field_order(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                            const std::uint32_t* indices, std::size_t count, bool row) {
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (count > 0 && indices == nullptr) {
    return null_argument("indices");
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const std::size_t field_count = table->fields().size();
  constexpr std::uint8_t kNotRequested = 0U;
  constexpr std::uint8_t kRequested = 1U;
  constexpr std::uint8_t kAlreadyRemoved = 2U;
  std::vector<std::uint8_t> requested_field(field_count, kNotRequested);
  // The getter exposes the axis vector's storage, so copy the request before
  // updating either order in case `indices` aliases one of them. The Values
  // level is split off: `values_slot` is its index among the fields.
  std::vector<std::uint32_t> requested_order;
  std::optional<std::size_t> values_slot;
  for (std::size_t i = 0; i < count; ++i) {
    if (indices[i] == FM_PIVOT_VALUES_FIELD) {
      if (values_slot.has_value()) {
        return invalid_argument("indices must not list the Values field twice", "i=" + std::to_string(i));
      }
      values_slot = requested_order.size();
      continue;
    }
    if (auto rc = check_index(indices[i], field_count, "indices[i]"); rc != 0) {
      return rc;
    }
    if (requested_field[indices[i]] != kNotRequested) {
      return invalid_argument("indices must not list a field twice",
                              "i=" + std::to_string(i) + " value=" + std::to_string(indices[i]));
    }
    requested_field[indices[i]] = kRequested;
    requested_order.push_back(indices[i]);
  }
  const bool holds_values = has_values_level(*table) && values_on_rows(*table) == row;
  if (values_slot.has_value() && !has_values_level(*table)) {
    return invalid_argument("indices lists the Values field, which needs at least two data fields",
                            "data_fields=" + std::to_string(table->data_fields().size()));
  }
  if (holds_values && !values_slot.has_value()) {
    return invalid_argument("indices must list the Values field that is on this axis");
  }
  // The list is the axis's whole membership: a field left out comes off the
  // axis, and a listed field moves onto it from wherever it was.
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  const formulon::pivot::PivotAxis axis = row ? formulon::pivot::PivotAxis::Row : formulon::pivot::PivotAxis::Col;
  std::vector<std::uint32_t>& target = row ? table->mutable_row_field_order() : table->mutable_col_field_order();
  for (const std::uint32_t field : target) {
    if (field < field_count && requested_field[field] == kNotRequested) {
      table->set_field_axis(field, formulon::pivot::PivotAxis::None);
    }
  }
  for (const std::uint32_t field : requested_order) {
    table->set_field_axis(field, axis);
  }

  // Remove each requested field's first occurrence from the opposite order,
  // preserving its stable order even if a loaded document contains duplicates.
  std::vector<std::uint32_t>& opposite = row ? table->mutable_col_field_order() : table->mutable_row_field_order();
  const std::optional<std::size_t> opposite_values_position =
      row ? table->col_values_position() : table->row_values_position();
  std::size_t removed_before_values = 0U;
  std::size_t kept = 0U;
  for (std::size_t read = 0U; read < opposite.size(); ++read) {
    const std::uint32_t field = opposite[read];
    if (field < field_count && requested_field[field] == kRequested) {
      requested_field[field] = kAlreadyRemoved;
      if (opposite_values_position.has_value() && read < *opposite_values_position) {
        ++removed_before_values;
      }
      continue;
    }
    if (kept != read) {
      opposite[kept] = field;
    }
    ++kept;
  }
  opposite.resize(kept);
  if (opposite_values_position.has_value() && removed_before_values > 0U) {
    set_values_position(*table, !row, *opposite_values_position - removed_before_values);
  }
  target.assign(requested_order.begin(), requested_order.end());
  if (values_slot.has_value()) {
    // The Values level moves here; the other axis gives it up.
    set_values_position(*table, row, *values_slot);
    set_values_position(*table, !row, std::nullopt);
  } else {
    // Keep a Values marker within the new order's slots.
    const std::optional<std::size_t> values_position =
        row ? table->row_values_position() : table->col_values_position();
    if (values_position.has_value() && *values_position > requested_order.size()) {
      set_values_position(*table, row, requested_order.size());
    }
  }
  return edit.commit();
}

}  // namespace

extern "C" fm_status_t fm_sheet_set_pivot_row_field_order(fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, const std::uint32_t* indices,
                                                          std::size_t count) {
  clear_last_error();
  return set_field_order(wb, sheet_index, pivot_index, indices, count, true);
}

extern "C" fm_status_t fm_sheet_set_pivot_column_field_order(fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, const std::uint32_t* indices,
                                                             std::size_t count) {
  clear_last_error();
  return set_field_order(wb, sheet_index, pivot_index, indices, count, false);
}

// ---------------------------------------------------------------------------
// PivotTable data fields
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_get_pivot_data_field_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                           std::size_t pivot_index, std::size_t* out_count) {
  clear_last_error();
  return pivot_count(wb, sheet_index, pivot_index, out_count,
                     [](const formulon::pivot::PivotTable& table) { return table.data_fields().size(); });
}

namespace {

// Name Excel gives a new data field of `field_index` under `agg`; `cache` is the
// table's cache, when it has one.
std::string default_data_field_name(const formulon::pivot::PivotTable& table, const formulon::pivot::PivotCache* cache,
                                    formulon::pivot::Aggregation agg, std::size_t field_index,
                                    formulon::ExcelProfile profile, std::optional<std::size_t> exclude) {
  return cache != nullptr ? formulon::pivot::data_field_default_name(table, *cache, agg, field_index, profile, exclude)
                          : formulon::pivot::data_field_display_name(
                                agg, formulon::pivot::pivot_field_display_name(table.fields()[field_index]), profile);
}

// Materialises a `fm_pivot_data_field_t` into a `PivotDataField`.
// The spec is already validated by `validate(*spec, api)` at the entry
// point; `number_format` is its format code already converted to the
// numFmtId text the model stores. `field_index` and a set
// `show_as_base_field` must index `table.fields()`, else the binding error
// is set and the status returned. A NULL or empty name is the one Excel
// gives a new data field, `exclude` being the entry the spec replaces.
// `kPreconditionFailed` for a function the cache's calculated items exclude.
fm_status_t fill_data_field(const formulon::Workbook& workbook, const formulon::pivot::PivotTable& table,
                            const fm_pivot_data_field_t& spec, std::string number_format,
                            std::optional<std::size_t> exclude, formulon::pivot::PivotDataField* out) {
  if (auto rc = check_index(spec.field_index, table.fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec.has_show_as_base_field, 0, 1, "has_show_as_base_field"); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(spec.has_show_as_base_item, 0, 1, "has_show_as_base_item"); rc != 0) {
    return rc;
  }
  if (spec.has_show_as_base_field != 0) {
    if (auto rc = check_index(spec.show_as_base_field, table.fields().size(), "show_as_base_field"); rc != 0) {
      return rc;
    }
  }
  out->aggregation = pivot_agg_from_fm(static_cast<fm_pivot_aggregation_t>(spec.aggregation));
  const formulon::pivot::PivotCache* cache = find_cache(workbook, table.pivot_cache_id());
  if (cache != nullptr) {
    if (auto checked = formulon::pivot::check_aggregation_with_calculated_items(*cache, out->aggregation); !checked) {
      return set_last_error(checked.error());
    }
  }
  if (spec.name != nullptr && spec.name[0] != '\0') {
    out->name = spec.name;
  } else {
    out->name =
        default_data_field_name(table, cache, out->aggregation, spec.field_index, workbook.excel_profile(), exclude);
  }
  out->field_index = spec.field_index;
  out->number_format = std::move(number_format);
  out->show_as = pivot_show_as_from_fm(static_cast<fm_pivot_show_values_as_t>(spec.show_as));
  if (spec.has_show_as_base_field != 0) {
    out->show_as_base_field = spec.show_as_base_field;
  } else {
    out->show_as_base_field.reset();
  }
  if (spec.has_show_as_base_item != 0) {
    out->show_as_base_item = spec.show_as_base_item;
  } else {
    out->show_as_base_item.reset();
  }
  return 0;
}

// Validates `spec`, resolves the table and builds the data field. With
// `replace_index` the index must name an existing data field.
fm_status_t prepare_data_field(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                               const fm_pivot_data_field_t& spec, std::optional<std::size_t> replace_index,
                               formulon::pivot::PivotTable** table, formulon::pivot::PivotDataField* df) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(spec); rc != 0) {
    return rc;
  }
  if (auto rc = validate(spec); rc != 0) {
    return rc;
  }
  if (auto rc = resolve_pivot_mut(wb, sheet_index, pivot_index, table); rc != 0) {
    return rc;
  }
  if (replace_index.has_value()) {
    if (auto rc = check_index(*replace_index, (*table)->data_fields().size(), "data_field_index"); rc != 0) {
      return rc;
    }
  }
  std::string number_format;
  if (auto rc = pivot_num_fmt_id_text(wb->workbook(), spec.number_format, &number_format); rc != 0) {
    return rc;
  }
  return fill_data_field(wb->workbook(), **table, spec, std::move(number_format), replace_index, df);
}

}  // namespace

extern "C" fm_status_t fm_sheet_add_pivot_data_field(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, const fm_pivot_data_field_t* spec,
                                                     std::size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr || out_index == nullptr) {
    return null_argument();
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotDataField df;
  if (auto rc = prepare_data_field(wb, sheet_index, pivot_index, *spec, std::nullopt, &table, &df); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->mutable_data_fields().push_back(std::move(df));
  const std::size_t added = table->data_fields().size() - 1U;
  const fm_status_t status = edit.commit();
  if (status == 0) {
    *out_index = added;
  }
  return status;
}

extern "C" fm_status_t fm_sheet_clear_pivot_data_fields(fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->mutable_data_fields().clear();
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_set_pivot_data_field(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::size_t data_field_index,
                                                     const fm_pivot_data_field_t* spec) {
  clear_last_error();
  if (wb == nullptr || spec == nullptr) {
    return null_argument();
  }
  formulon::pivot::PivotTable* table = nullptr;
  formulon::pivot::PivotDataField df;
  if (auto rc = prepare_data_field(wb, sheet_index, pivot_index, *spec, data_field_index, &table, &df); rc != 0) {
    return rc;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  table->mutable_data_fields()[data_field_index] = std::move(df);
  return edit.commit();
}

// ---------------------------------------------------------------------------
// PivotTable definition read-back
// ---------------------------------------------------------------------------

namespace {

fm_status_t get_field_order(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index,
                            std::uint32_t* out, std::size_t capacity, std::size_t* out_count, bool row) {
  if (out == nullptr && capacity > 0U) {
    return null_argument("out");
  }
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_count}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const std::vector<std::uint32_t>& order = row ? table->row_field_order() : table->col_field_order();
  // The Values level is listed on the axis that holds it; the columns hold it,
  // at the end, when the table records no position.
  std::optional<std::size_t> values_slot;
  if (has_values_level(*table) && values_on_rows(*table) == row) {
    values_slot = row ? table->row_values_position() : table->col_values_position().value_or(order.size());
  }
  std::vector<std::uint32_t> listing(order.begin(), order.end());
  if (values_slot.has_value()) {
    const std::size_t slot = std::min(*values_slot, order.size());
    listing.insert(listing.begin() + static_cast<std::ptrdiff_t>(slot), FM_PIVOT_VALUES_FIELD);
  }
  std::copy_n(listing.begin(), std::min(capacity, listing.size()), out);
  *out_count = listing.size();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_pivot_cache_id(const fm_workbook_t* wb, std::size_t sheet_index,
                                                   std::size_t pivot_index, std::uint32_t* out_cache_id) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_cache_id}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  *out_cache_id = table->pivot_cache_id();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_name(const fm_workbook_t* wb, std::size_t sheet_index,
                                               std::size_t pivot_index, const char** out_name) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_name}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  *out_name = table->name().c_str();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_runtime_id(const fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, std::uint32_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  formulon::RuntimeId& id = table->runtime_id();
  if (id.value() == 0U) {
    id.assign(wb->workbook().allocate_runtime_id());
  }
  *out = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_resolve_pivot_runtime_id(const fm_workbook_t* wb, std::size_t sheet_index,
                                                         std::uint32_t id, std::size_t* out_pivot_index) {
  clear_last_error();
  if (wb == nullptr || out_pivot_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const auto& pivots = wb->workbook().sheet(sheet_index).pivot_tables();
  // Zero is the id of a pivot nobody has asked about, so it names none.
  for (std::size_t i = 0; id != 0U && i < pivots.size(); ++i) {
    if (pivots[i] != nullptr && pivots[i]->runtime_id().value() == id) {
      *out_pivot_index = i;
      return 0;
    }
  }
  return not_found(("id " + std::to_string(id) + " names no pivot on the sheet").c_str(),
                   "sheet_index=" + std::to_string(sheet_index) + " id=" + std::to_string(id));
}

extern "C" fm_status_t fm_sheet_get_pivot_grand_totals(const fm_workbook_t* wb, std::size_t sheet_index,
                                                       std::size_t pivot_index, fm_pivot_grand_totals_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_pivot_grand_totals_t{};
  out->struct_size = struct_size;
  out->rows = table->grand_totals_rows() ? 1 : 0;
  out->columns = table->grand_totals_cols() ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_style_options(const fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, int32_t* out_present,
                                                        fm_pivot_style_options_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out_present, out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  fm_pivot_style_options_t result{};
  result.struct_size = out->struct_size;
  *out_present = 0;
  if (const auto& info = table->style_info(); info.has_value()) {
    result.style_name = text_or_null(info->name.c_str());
    *out_present = 1;
    result.show_row_headers = info->show_row_headers ? 1 : 0;
    result.show_column_headers = info->show_col_headers ? 1 : 0;
    result.show_row_stripes = info->show_row_stripes ? 1 : 0;
    result.show_column_stripes = info->show_col_stripes ? 1 : 0;
    result.show_last_column = info->show_last_column ? 1 : 0;
  }
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_data_field_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                        std::size_t pivot_index, std::size_t data_field_index,
                                                        fm_pivot_data_field_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(data_field_index, table->data_fields().size(), "data_field_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotDataField& df = table->data_fields()[data_field_index];
  fm_pivot_data_field_t result{};
  result.struct_size = out->struct_size;
  result.name = df.name.c_str();
  result.field_index = df.field_index;
  result.aggregation = static_cast<fm_pivot_aggregation_t>(df.aggregation);
  result.number_format = text_or_null(pivot_num_fmt_code(wb->workbook(), df.number_format));
  result.show_as = static_cast<fm_pivot_show_values_as_t>(df.show_as);
  result.has_show_as_base_field = df.show_as_base_field.has_value() ? 1 : 0;
  result.show_as_base_field = df.show_as_base_field.value_or(0U);
  result.has_show_as_base_item = df.show_as_base_item.has_value() ? 1 : 0;
  result.show_as_base_item = df.show_as_base_item.value_or(0U);
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_pivot_row_field_order(const fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t pivot_index, std::uint32_t* out,
                                                          std::size_t capacity, std::size_t* out_count) {
  clear_last_error();
  return get_field_order(wb, sheet_index, pivot_index, out, capacity, out_count, true);
}

extern "C" fm_status_t fm_sheet_get_pivot_column_field_order(const fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t pivot_index, std::uint32_t* out,
                                                             std::size_t capacity, std::size_t* out_count) {
  clear_last_error();
  return get_field_order(wb, sheet_index, pivot_index, out, capacity, out_count, false);
}

// ---------------------------------------------------------------------------
// PivotTable options, default data-field names, projection
// ---------------------------------------------------------------------------

namespace {

// The stored value of a scratch-backed option string: NULL when unset.
const char* option_text(formulon::c_api::parts::TextStore& store, const std::optional<std::string>& text) {
  return text.has_value() ? store_cstr(store, *text) : nullptr;
}

const char* option_text(formulon::c_api::parts::TextStore& store, const std::string& text) {
  return text.empty() ? nullptr : store_cstr(store, text);
}

// `value` within 0..`max`, or the value the table already holds: a loaded
// definition keeps an out-of-range value across a get / set round trip.
fm_status_t check_option_range(std::uint32_t value, std::uint32_t current, std::uint32_t max, const char* field) {
  if (value <= max || value == current) {
    return 0;
  }
  return invalid_argument((std::string(field) + " must be in 0.." + std::to_string(max)).c_str(),
                          std::string(field) + "=" + std::to_string(value));
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_pivot_options(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t pivot_index, fm_pivot_options_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  formulon::c_api::parts::TextStore& store = fresh_read_scratch(wb);
  fm_pivot_options_t result{};
  result.struct_size = out->struct_size;
  result.show_error = table->show_error() ? 1 : 0;
  result.error_caption = option_text(store, table->error_caption());
  result.show_missing = table->show_missing() ? 1 : 0;
  result.missing_caption = option_text(store, table->missing_caption());
  result.show_headers = table->show_headers() ? 1 : 0;
  result.show_drill = table->show_drill() ? 1 : 0;
  result.show_data_tips = table->show_data_tips() ? 1 : 0;
  result.grid_drop_zones = table->grid_drop_zones() ? 1 : 0;
  result.show_empty_row = table->show_empty_row() ? 1 : 0;
  result.show_empty_column = table->show_empty_col() ? 1 : 0;
  result.show_data_drop_down = table->show_data_drop_down() ? 1 : 0;
  result.print_drill = table->print_drill() ? 1 : 0;
  result.field_print_titles = table->field_print_titles() ? 1 : 0;
  result.item_print_titles = table->item_print_titles() ? 1 : 0;
  result.merge_item = table->merge_item() ? 1 : 0;
  result.indent = table->indent();
  result.page_over_then_down = table->page_over_then_down() ? 1 : 0;
  result.page_wrap = table->page_wrap();
  result.multiple_field_filters = table->multiple_field_filters() ? 1 : 0;
  result.preserve_formatting = table->preserve_formatting() ? 1 : 0;
  result.custom_list_sort = table->custom_list_sort() ? 1 : 0;
  result.alt_text = option_text(store, table->alt_text());
  result.alt_text_summary = option_text(store, table->alt_text_summary());
  result.grand_total_caption = option_text(store, table->grand_total_caption());
  result.auto_formatting = table->auto_formatting() ? 1 : 0;
  result.enable_drill = table->enable_drill() ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_sheet_update_pivot_options(fm_workbook_t* wb, std::size_t sheet_index,
                                                     std::size_t pivot_index, const fm_pivot_options_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  const fm_pivot_options_patch_t& o = *patch;
  using Table = formulon::pivot::PivotTable;
  struct FlagField {
    fm_patch_op_t op;
    std::int32_t value;
    void (Table::*set)(bool);
    const char* field;
  };
  const FlagField flags[] = {
      {o.show_error_op, o.show_error, &Table::set_show_error, "show_error"},
      {o.show_missing_op, o.show_missing, &Table::set_show_missing, "show_missing"},
      {o.show_headers_op, o.show_headers, &Table::set_show_headers, "show_headers"},
      {o.show_drill_op, o.show_drill, &Table::set_show_drill, "show_drill"},
      {o.show_data_tips_op, o.show_data_tips, &Table::set_show_data_tips, "show_data_tips"},
      {o.grid_drop_zones_op, o.grid_drop_zones, &Table::set_grid_drop_zones, "grid_drop_zones"},
      {o.show_empty_row_op, o.show_empty_row, &Table::set_show_empty_row, "show_empty_row"},
      {o.show_empty_column_op, o.show_empty_column, &Table::set_show_empty_col, "show_empty_column"},
      {o.show_data_drop_down_op, o.show_data_drop_down, &Table::set_show_data_drop_down, "show_data_drop_down"},
      {o.print_drill_op, o.print_drill, &Table::set_print_drill, "print_drill"},
      {o.field_print_titles_op, o.field_print_titles, &Table::set_field_print_titles, "field_print_titles"},
      {o.item_print_titles_op, o.item_print_titles, &Table::set_item_print_titles, "item_print_titles"},
      {o.merge_item_op, o.merge_item, &Table::set_merge_item, "merge_item"},
      {o.page_over_then_down_op, o.page_over_then_down, &Table::set_page_over_then_down, "page_over_then_down"},
      {o.multiple_field_filters_op, o.multiple_field_filters, &Table::set_multiple_field_filters,
       "multiple_field_filters"},
      {o.preserve_formatting_op, o.preserve_formatting, &Table::set_preserve_formatting, "preserve_formatting"},
      {o.custom_list_sort_op, o.custom_list_sort, &Table::set_custom_list_sort, "custom_list_sort"},
      {o.auto_formatting_op, o.auto_formatting, &Table::set_auto_formatting, "auto_formatting"},
      {o.enable_drill_op, o.enable_drill, &Table::set_enable_drill, "enable_drill"},
  };
  struct TextField {
    fm_patch_op_t op;
    const char* value;
    bool capped;
    const char* field;
  };
  const TextField texts[] = {
      {o.error_caption_op, o.error_caption, true, "error_caption"},
      {o.missing_caption_op, o.missing_caption, true, "missing_caption"},
      {o.alt_text_op, o.alt_text, true, "alt_text"},
      {o.alt_text_summary_op, o.alt_text_summary, false, "alt_text_summary"},
      {o.grand_total_caption_op, o.grand_total_caption, true, "grand_total_caption"},
  };
  // The whole patch is checked before anything changes; only the five strings are clearable.
  bool changes = false;
  for (const FlagField& flag : flags) {
    if (auto rc = check_patch_op(flag.op, flag.field, false); rc != 0) {
      return rc;
    }
    changes = changes || flag.op == FM_PATCH_OP_SET;
  }
  for (const TextField& text : texts) {
    if (auto rc = check_patch_op(text.op, text.field, true); rc != 0) {
      return rc;
    }
    if (text.op == FM_PATCH_OP_SET) {
      if (text.value == nullptr) {
        return null_field(text.field);
      }
      if (auto rc = check_utf8(text.value, "field", text.field); rc != 0) {
        return rc;
      }
      if (text.capped && formulon::utf16_units_in(text.value) > kMaxPivotNameUnits) {
        return invalid_argument((std::string(text.field) + " must be at most 255 UTF-16 units").c_str());
      }
    }
    changes = changes || text.op != FM_PATCH_OP_KEEP;
  }
  for (const auto& [op, value, current, max, field] :
       {std::tuple<fm_patch_op_t, std::uint32_t, std::uint32_t, std::uint32_t, const char*>{
            o.indent_op, o.indent, table->indent(), kMaxIndent, "indent"},
        {o.page_wrap_op, o.page_wrap, table->page_wrap(), kMaxPageWrap, "page_wrap"}}) {
    if (auto rc = check_patch_op(op, field, false); rc != 0) {
      return rc;
    }
    if (op == FM_PATCH_OP_SET) {
      if (auto rc = check_option_range(value, current, max, field); rc != 0) {
        return rc;
      }
      changes = true;
    }
  }
  if (!changes) {
    return 0;
  }
  PivotEdit edit(wb->workbook(), sheet_index, pivot_index, *table);
  for (const FlagField& flag : flags) {
    if (flag.op == FM_PATCH_OP_SET) {
      (table->*flag.set)(flag.value != 0);
    }
  }
  // CLEAR and SET of "" both unset a string option.
  const auto next_text = [](const TextField& text) {
    return text.op == FM_PATCH_OP_SET ? optional_text(text.value) : std::nullopt;
  };
  if (texts[0].op != FM_PATCH_OP_KEEP) {
    table->set_error_caption(next_text(texts[0]));
  }
  if (texts[1].op != FM_PATCH_OP_KEEP) {
    table->set_missing_caption(next_text(texts[1]));
  }
  if (texts[2].op != FM_PATCH_OP_KEEP) {
    table->set_alt_text(next_text(texts[2]).value_or(std::string()));
  }
  if (texts[3].op != FM_PATCH_OP_KEEP) {
    table->set_alt_text_summary(next_text(texts[3]).value_or(std::string()));
  }
  if (texts[4].op != FM_PATCH_OP_KEEP) {
    table->set_grand_total_caption(next_text(texts[4]));
  }
  if (o.indent_op == FM_PATCH_OP_SET) {
    table->set_indent(o.indent);
  }
  if (o.page_wrap_op == FM_PATCH_OP_SET) {
    table->set_page_wrap(o.page_wrap);
  }
  return edit.commit();
}

extern "C" fm_status_t fm_sheet_get_pivot_data_field_default_name(const fm_workbook_t* wb, std::size_t sheet_index,
                                                                  std::size_t pivot_index, std::size_t field_index,
                                                                  std::int32_t aggregation,
                                                                  std::size_t exclude_data_field, const char** out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(aggregation, 0, FM_PIVOT_AGGREGATION_VARP, "aggregation"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(field_index, table->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  std::optional<std::size_t> exclude;
  if (exclude_data_field != FM_PIVOT_DATA_FIELD_NONE) {
    if (auto rc = check_index(exclude_data_field, table->data_fields().size(), "exclude_data_field"); rc != 0) {
      return rc;
    }
    exclude = exclude_data_field;
  }
  const formulon::pivot::Aggregation agg = pivot_agg_from_fm(static_cast<fm_pivot_aggregation_t>(aggregation));
  const formulon::ExcelProfile profile = wb->workbook().excel_profile();
  const formulon::pivot::PivotCache* cache = find_cache(wb->workbook(), table->pivot_cache_id());
  *out = publish_scratch(wb, default_data_field_name(*table, cache, agg, field_index, profile, exclude));
  return 0;
}

extern "C" fm_status_t fm_sheet_project_pivot(fm_workbook_t* wb, std::size_t sheet_index, std::size_t pivot_index) {
  clear_last_error();
  formulon::pivot::PivotTable* table = nullptr;
  if (fm_status_t rc = resolve_pivot_checked(wb, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  if (auto projected = project_pivot(wb->workbook(), sheet_index, pivot_index); !projected) {
    return set_last_error(projected.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_is_pivot_projection_pending(const fm_workbook_t* wb, std::size_t sheet_index,
                                                            std::size_t pivot_index, std::int32_t* out) {
  clear_last_error();
  const formulon::pivot::PivotTable* table = nullptr;
  if (auto rc = resolve_pivot_checked(wb, {out}, sheet_index, pivot_index, &table); rc != 0) {
    return rc;
  }
  *out = table->projection_pending() ? 1 : 0;
  return 0;
}
