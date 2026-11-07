//
// C ABI - chart creation and the typed chart edits.
//
// Patch structs are translated field by field into the C++ patch types
// (`PatchReader`); the edits themselves live in `chart/chart_edit`.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/chart_patch_reader.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "chart/chart_edit.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::anchor_kind_from_c;
using formulon::c_api::parts::chart_axis_crosses_from_c;
using formulon::c_api::parts::chart_blanks_as_from_c;
using formulon::c_api::parts::chart_legend_position_from_c;
using formulon::c_api::parts::chart_title_kind_from_c;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::edit_as_from_c;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::ingest_entered_formula;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::PatchReader;
using formulon::c_api::parts::saturate_index;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::text_or_empty;

namespace {

// ---------------------------------------------------------------------------
// C patches -> model
// ---------------------------------------------------------------------------

/// A C title as the model's; `NULL` strings read as empty.
formulon::ChartTitle title_of(const fm_chart_title_t& t) {
  return formulon::ChartTitle{chart_title_kind_from_c(t.kind), text_or_empty(t.text), text_or_empty(t.formula)};
}

formulon::ChartFillPatch read_fill(PatchReader& r, const fm_chart_fill_patch_t& p) {
  formulon::ChartFillPatch out;
  r.rgb_field(p.color_op, p.color, out.color, "fill.color_op");
  r.field(p.transparency_op, p.transparency_pct, out.transparency_pct, "fill.transparency_op");
  return out;
}

formulon::ChartLinePatch read_line(PatchReader& r, const fm_chart_line_patch_t& p) {
  formulon::ChartLinePatch out;
  r.rgb_field(p.color_op, p.color, out.color, "line.color_op");
  r.field(p.transparency_op, p.transparency_pct, out.transparency_pct, "line.transparency_op");
  r.field(p.width_op, p.width_pt, out.width_pt, "line.width_op");
  r.enum_field(p.dash_op, p.dash, formulon::ChartLineDash::kOther, out.dash, "line.dash");
  r.flag(p.none_op, p.none, out.none, "line.none_op");
  return out;
}

formulon::ChartMarkerPatch read_marker(PatchReader& r, const fm_chart_marker_patch_t& p) {
  formulon::ChartMarkerPatch out;
  r.enum_optional(p.symbol_op, p.symbol, formulon::ChartMarkerSymbol::kDash, out.symbol, "marker.symbol");
  r.field(p.size_op, p.size, out.size, "marker.size_op");
  out.fill = read_fill(r, p.fill);
  out.line = read_line(r, p.line);
  return out;
}

formulon::ChartLabelsPatch read_labels(PatchReader& r, const fm_chart_labels_patch_t& p) {
  formulon::ChartLabelsPatch out;
  r.flag(p.show_value_op, p.show_value, out.show_val, "labels.show_value_op");
  r.flag(p.show_category_op, p.show_category, out.show_cat, "labels.show_category_op");
  r.flag(p.show_series_op, p.show_series, out.show_ser, "labels.show_series_op");
  r.flag(p.show_percent_op, p.show_percent, out.show_pct, "labels.show_percent_op");
  r.flag(p.show_legend_key_op, p.show_legend_key, out.show_legend_key, "labels.show_legend_key_op");
  r.text_field(p.separator_op, p.separator, out.separator, "labels.separator_op");
  r.enum_field(p.position_op, p.position, formulon::ChartLabelPosition::kRight, out.position, "labels.position");
  r.text_field(p.number_format_op, p.number_format, out.number_format, "labels.number_format_op");
  r.rgb_field(p.font_color_op, p.font_color, out.font_color, "labels.font_color_op");
  r.field(p.font_size_op, p.font_size_pt, out.font_size_pt, "labels.font_size_op");
  return out;
}

void read_ref(PatchReader& r, const char* text, std::optional<std::string>& out, const char* name) {
  if (text != nullptr && r.text_ok(text, name)) {
    out = std::string(text);
  }
}

formulon::ChartSeriesPatch read_series(PatchReader& r, const fm_chart_series_patch_t& p) {
  formulon::ChartSeriesPatch out;
  read_ref(r, p.name, out.name, "name");
  read_ref(r, p.categories, out.cat, "categories");
  read_ref(r, p.values, out.val, "values");
  read_ref(r, p.x, out.x, "x");
  read_ref(r, p.y, out.y, "y");
  read_ref(r, p.size, out.size, "size");
  out.fill = read_fill(r, p.fill);
  out.line = read_line(r, p.line);
  r.field(p.marker_op, p.marker_op == FM_CHART_PATCH_SET ? read_marker(r, p.marker) : formulon::ChartMarkerPatch{},
          out.marker, "marker_op");
  r.flag(p.smooth_op, p.smooth, out.smooth, "smooth_op");
  r.field(p.labels_op, p.labels_op == FM_CHART_PATCH_SET ? read_labels(r, p.labels) : formulon::ChartLabelsPatch{},
          out.labels, "labels_op");
  return out;
}

/// Stores the formulas a patch enters, once its checks passed (see `ingest_entered_formula`).
void ingest_series_refs(fm_workbook_t* wb, formulon::ChartSeriesPatch& patch) {
  for (std::optional<std::string>* ref : {&patch.name, &patch.cat, &patch.val, &patch.x, &patch.y, &patch.size}) {
    if (ref->has_value() && !(*ref)->empty()) {
      **ref = ingest_entered_formula(wb, **ref);
    }
  }
}

void ingest_title_ref(fm_workbook_t* wb, formulon::ChartTitle& title) {
  if (!title.ref_formula.empty()) {
    title.ref_formula = ingest_entered_formula(wb, title.ref_formula);
  }
}

/// Shared entry checks of the edit functions.
fm_status_t edit_entry(const fm_workbook_t* wb, std::size_t sheet_index, const void* required) {
  clear_last_error();
  if (required == nullptr) {
    return null_argument();
  }
  return check_sheet_index(wb, sheet_index);
}

}  // namespace

// ---------------------------------------------------------------------------
// Creation and edits
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_chart(fm_workbook_t* wb, size_t sheet_index, const fm_chart_insert_t* opts,
                                          uint32_t* out_object_id) {
  if (out_object_id == nullptr) {
    return null_argument();
  }
  if (auto rc = edit_entry(wb, sheet_index, opts); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*opts); rc != 0) {
    return rc;
  }
  if (opts->source == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({opts->source, opts->title, opts->placement.name, opts->placement.description}); rc != 0) {
    return rc;
  }
  PatchReader r;
  r.enum_ok(opts->kind, formulon::ChartInsertKind::kDoughnut, "kind");
  r.enum_ok(opts->grouping, formulon::ChartGrouping::kStandard, "grouping");
  r.enum_ok(opts->direction, formulon::ChartSeriesDirection::kRows, "direction");
  r.enum_ok(opts->header_row, formulon::ChartHeaderMode::kNo, "header_row");
  r.enum_ok(opts->header_column, formulon::ChartHeaderMode::kNo, "header_column");
  r.enum_ok(opts->legend, FM_CHART_LEGEND_NONE, FM_CHART_LEGEND_TOP_RIGHT, "legend");
  r.enum_ok(opts->placement.anchor.anchor_kind, FM_ANCHOR_KIND_ABSOLUTE, "placement.anchor.anchor_kind");
  r.enum_ok(opts->placement.anchor.edit_as, FM_ANCHOR_EDIT_AS_ABSOLUTE, "placement.anchor.edit_as");
  if (r.status() != 0) {
    return r.status();
  }
  formulon::ChartInsertOptions model;
  model.kind = static_cast<formulon::ChartInsertKind>(opts->kind);
  model.grouping = static_cast<formulon::ChartGrouping>(opts->grouping);
  model.source = opts->source;
  model.direction = static_cast<formulon::ChartSeriesDirection>(opts->direction);
  model.header_row = static_cast<formulon::ChartHeaderMode>(opts->header_row);
  model.header_column = static_cast<formulon::ChartHeaderMode>(opts->header_column);
  if (opts->title != nullptr) {
    model.title = std::string(opts->title);
  }
  model.legend = chart_legend_position_from_c(opts->legend);
  const fm_drawing_anchor_t& p = opts->placement.anchor;
  model.placement.anchor_kind = anchor_kind_from_c(p.anchor_kind);
  model.placement.edit_as = edit_as_from_c(p.edit_as);
  model.placement.row = p.row;
  model.placement.col = p.col;
  model.placement.row_off = p.row_offset_emu;
  model.placement.col_off = p.col_offset_emu;
  model.placement.width_emu = p.width_emu;
  model.placement.height_emu = p.height_emu;
  model.name = text_or_empty(opts->placement.name);
  model.descr = text_or_empty(opts->placement.description);
  auto id = formulon::insert_chart(wb->workbook(), sheet_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_set_chart_type(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                               int32_t kind, int32_t grouping) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  PatchReader r;
  r.enum_ok(kind, formulon::ChartInsertKind::kDoughnut, "kind");
  r.enum_ok(grouping, formulon::ChartGrouping::kStandard, "grouping");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_type(wb->workbook(), sheet_index, object_id, plot,
                                                static_cast<formulon::ChartInsertKind>(kind),
                                                static_cast<formulon::ChartGrouping>(grouping)));
}

extern "C" fm_status_t fm_sheet_set_chart_title(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                const fm_chart_title_t* title) {
  if (auto rc = edit_entry(wb, sheet_index, title); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(title->kind, 0, FM_CHART_TITLE_AUTO, "kind"); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({title->text, title->formula}); rc != 0) {
    return rc;
  }
  formulon::ChartTitle model = title_of(*title);
  ingest_title_ref(wb, model);
  return finish_status(formulon::set_chart_title(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const fm_chart_legend_patch_t* legend) {
  if (auto rc = edit_entry(wb, sheet_index, legend); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*legend); rc != 0) {
    return rc;
  }
  if (legend->hidden_series == nullptr && legend->hidden_series_count != 0) {
    return null_argument();
  }
  PatchReader r;
  r.enum_ok(legend->position, FM_CHART_LEGEND_NONE, FM_CHART_LEGEND_TOP_RIGHT, "position");
  formulon::ChartLegendPatch model;
  model.position = chart_legend_position_from_c(legend->position);
  model.overlay = legend->overlay != 0;
  std::vector<std::uint32_t> hidden;
  if (legend->hidden_series_count != 0) {
    hidden.assign(legend->hidden_series, legend->hidden_series + legend->hidden_series_count);
  }
  r.optional(legend->hidden_series_op, std::move(hidden), model.hidden_series_idx, "hidden_series_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_legend(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_axis(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               uint32_t axis_id, const fm_chart_axis_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartAxisPatch model;
  r.flag(patch->visible_op, patch->visible, model.visible, "visible_op");
  if (patch->title_op == FM_CHART_PATCH_SET) {
    r.enum_ok(patch->title.kind, FM_CHART_TITLE_AUTO, "title.kind");
    r.text_ok(patch->title.text, "title.text");
    r.text_ok(patch->title.formula, "title.formula");
  }
  r.field(patch->title_op, title_of(patch->title), model.title, "title_op");
  r.field(patch->min_op, patch->min, model.min, "min_op");
  r.field(patch->max_op, patch->max, model.max, "max_op");
  r.field(patch->major_unit_op, patch->major_unit, model.major_unit, "major_unit_op");
  r.text_field(patch->number_format_op, patch->number_format, model.number_format, "number_format_op");
  r.flag(patch->major_gridlines_op, patch->major_gridlines, model.major_gridlines, "major_gridlines_op");
  r.flag(patch->reversed_op, patch->reversed, model.reversed, "reversed_op");
  r.field(patch->log_base_op, patch->log_base, model.log_base, "log_base_op");
  r.enum_optional(patch->crosses_op, patch->crosses, FM_CHART_CROSSES_ABSENT, FM_CHART_CROSSES_AT,
                  chart_axis_crosses_from_c, model.crosses, "crosses");
  model.crosses_at = patch->crosses_at;
  r.text_field(patch->tick_label_pos_op, patch->tick_label_pos, model.tick_label_pos, "tick_label_pos_op");
  if (r.status() != 0) {
    return r.status();
  }
  if (model.title.is_set()) {
    ingest_title_ref(wb, model.title.value);
  }
  return finish_status(formulon::set_chart_axis(wb->workbook(), sheet_index, object_id, axis_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_display(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                  const fm_chart_display_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartDisplayPatch model;
  if (patch->blanks_as_op == FM_CHART_PATCH_SET && patch->blanks_as == FM_CHART_BLANKS_ABSENT) {
    r.fail("blanks_as", "cannot be set to ABSENT");
  }
  r.enum_field(patch->blanks_as_op, patch->blanks_as, FM_CHART_BLANKS_ABSENT, FM_CHART_BLANKS_SPAN,
               chart_blanks_as_from_c, model.blanks_as, "blanks_as");
  r.bool_field(patch->plot_vis_only_op, patch->plot_vis_only, model.plot_vis_only, "plot_vis_only");
  r.bool_field(patch->na_as_blank_op, patch->na_as_blank, model.na_as_blank, "na_as_blank");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_display(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_set_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                 size_t series_index, const fm_chart_series_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartSeriesPatch model = read_series(r, *patch);
  if (r.status() != 0) {
    return r.status();
  }
  ingest_series_refs(wb, model);
  return finish_status(
      formulon::set_chart_series(wb->workbook(), sheet_index, object_id, plot, saturate_index(series_index), model));
}

extern "C" fm_status_t fm_sheet_set_chart_point(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                size_t series_index, size_t point_index,
                                                const fm_chart_point_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartPointPatch model;
  model.fill = read_fill(r, patch->fill);
  model.line = read_line(r, patch->line);
  r.field(patch->marker_op,
          patch->marker_op == FM_CHART_PATCH_SET ? read_marker(r, patch->marker) : formulon::ChartMarkerPatch{},
          model.marker, "marker_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_point(wb->workbook(), sheet_index, object_id, plot,
                                                 saturate_index(series_index), saturate_index(point_index), model));
}

extern "C" fm_status_t fm_sheet_set_chart_plot(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                               const fm_chart_plot_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartPlotPatch model;
  r.optional(patch->gap_width_op, patch->gap_width, model.gap_width, "gap_width_op");
  r.optional(patch->overlap_op, patch->overlap, model.overlap, "overlap_op");
  r.optional(patch->first_slice_angle_op, patch->first_slice_angle, model.first_slice_angle, "first_slice_angle_op");
  r.flag(patch->vary_colors_op, patch->vary_colors, model.vary_colors, "vary_colors_op");
  r.optional(patch->hole_size_op, patch->hole_size, model.hole_size, "hole_size_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_plot(wb->workbook(), sheet_index, object_id, plot, model));
}

extern "C" fm_status_t fm_sheet_set_chart_filter(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const uint32_t* hidden_series, size_t hidden_series_count,
                                                 const uint32_t* hidden_categories, size_t hidden_category_count) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  if ((hidden_series == nullptr && hidden_series_count != 0) ||
      (hidden_categories == nullptr && hidden_category_count != 0)) {
    return null_argument();
  }
  std::vector<std::uint32_t> series;
  std::vector<std::uint32_t> categories;
  if (hidden_series_count != 0) {
    series.assign(hidden_series, hidden_series + hidden_series_count);
  }
  if (hidden_category_count != 0) {
    categories.assign(hidden_categories, hidden_categories + hidden_category_count);
  }
  return finish_status(formulon::set_chart_filter(wb->workbook(), sheet_index, object_id, series, categories));
}

extern "C" fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id, size_t plot,
                                                 const fm_chart_series_patch_t* patch, size_t* out_series_index) {
  if (out_series_index == nullptr) {
    return null_argument();
  }
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartSeriesPatch model = read_series(r, *patch);
  if (r.status() != 0) {
    return r.status();
  }
  ingest_series_refs(wb, model);
  auto idx = formulon::add_chart_series(wb->workbook(), sheet_index, object_id, plot, model);
  if (!idx) {
    return set_last_error(idx.error());
  }
  *out_series_index = idx.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    size_t plot, size_t series_index) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  return finish_status(
      formulon::remove_chart_series(wb->workbook(), sheet_index, object_id, plot, saturate_index(series_index)));
}

extern "C" fm_status_t fm_sheet_set_chart_source(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 const char* source, int32_t direction, int32_t header_row,
                                                 int32_t header_column) {
  if (auto rc = edit_entry(wb, sheet_index, source); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({source}); rc != 0) {
    return rc;
  }
  PatchReader r;
  r.enum_ok(direction, formulon::ChartSeriesDirection::kRows, "direction");
  r.enum_ok(header_row, formulon::ChartHeaderMode::kNo, "header_row");
  r.enum_ok(header_column, formulon::ChartHeaderMode::kNo, "header_column");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_source(
      wb->workbook(), sheet_index, object_id, source, static_cast<formulon::ChartSeriesDirection>(direction),
      static_cast<formulon::ChartHeaderMode>(header_row), static_cast<formulon::ChartHeaderMode>(header_column)));
}
