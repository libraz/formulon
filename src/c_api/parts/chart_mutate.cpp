//
// C ABI - chart creation and the typed chart edits.
//
// Patch structs are translated field by field into the C++ patch types
// (`PatchReader`); the edits themselves live in `chart/chart_edit`.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/chart_patch_reader.h"
#include "c_api/parts/common.h"
#include "c_api/parts/drawing_anchor.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "chart/chart_edit.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

using formulon::c_api::parts::AnchorPlan;
using formulon::c_api::parts::apply_placement_extras;
using formulon::c_api::parts::chart_axis_crosses_from_c;
using formulon::c_api::parts::chart_blanks_as_from_c;
using formulon::c_api::parts::chart_legend_position_from_c;
using formulon::c_api::parts::chart_title_kind_from_c;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::ingest_entered_formula;
using formulon::c_api::parts::kTickLabelPositionTokens;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::PatchReader;
using formulon::c_api::parts::plan_add_anchor;
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
  r.color_field(p.color_op, p.color, out.color, "fill.color_op");
  r.field(p.transparency_pct_op, p.transparency_pct, out.transparency_pct, "fill.transparency_pct_op");
  return out;
}

formulon::ChartLinePatch read_line(PatchReader& r, const fm_chart_line_patch_t& p) {
  formulon::ChartLinePatch out;
  r.color_field(p.color_op, p.color, out.color, "line.color_op");
  r.field(p.transparency_pct_op, p.transparency_pct, out.transparency_pct, "line.transparency_pct_op");
  r.field(p.width_pt_op, p.width_pt, out.width_pt, "line.width_pt_op");
  r.enum_field(p.dash_op, p.dash, formulon::ChartLineDash::kOther, out.dash, "line.dash");
  // CLEAR drops an explicit no-line as SET 0 does, so a patch clearing every member stays valid.
  if (p.none_op == FM_PATCH_OP_CLEAR) {
    out.none = false;
  } else {
    r.flag(p.none_op, p.none, out.none, "line.none_op");
  }
  return out;
}

formulon::ChartMarkerPatch read_marker(PatchReader& r, const fm_chart_marker_patch_t& p) {
  formulon::ChartMarkerPatch out;
  r.enum_optional(p.symbol_op, p.symbol, formulon::ChartMarkerSymbol::kDash, out.symbol, "marker.symbol");
  r.field(p.size_pt_op, p.size_pt, out.size, "marker.size_pt_op");
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
  r.color_field(p.font_color_op, p.font_color, out.font_color, "labels.font_color_op");
  r.field(p.font_size_pt_op, p.font_size_pt, out.font_size_pt, "labels.font_size_pt_op");
  r.flag(p.source_linked_op, p.source_linked, out.source_linked, "labels.source_linked_op");
  r.flag(p.deleted_op, p.deleted, out.deleted, "labels.deleted_op");
  return out;
}

/// Reads one series input: `KEEP` leaves `out` unset, `CLEAR` and a `NONE` source remove the input.
void read_source(PatchReader& r, std::int32_t op, const fm_chart_data_source_spec_t& spec,
                 std::optional<formulon::ChartSeriesSource>& out, const char* op_name, const char* kind_name) {
  if (!r.enum_ok(op, FM_PATCH_OP_SET, op_name) || op == FM_PATCH_OP_KEEP) {
    return;
  }
  if (op == FM_PATCH_OP_CLEAR) {
    out = formulon::ChartSeriesSource::None();
    return;
  }
  if (!r.enum_ok(spec.kind, FM_CHART_SOURCE_KIND_LITERAL, kind_name)) {
    return;
  }
  if (spec.kind == FM_CHART_SOURCE_KIND_NONE) {
    out = formulon::ChartSeriesSource::None();
    return;
  }
  if (spec.kind == FM_CHART_SOURCE_KIND_REF) {
    if (r.text_ok(spec.formula, kind_name)) {
      out = formulon::ChartSeriesSource::Ref(text_or_empty(spec.formula));
    }
    return;
  }
  if (spec.literal == nullptr && spec.literal_count != 0 && !r.check(null_argument())) {
    return;
  }
  std::vector<formulon::ChartValue> points(spec.literal_count);
  for (std::size_t i = 0; i < spec.literal_count; ++i) {
    const fm_chart_literal_info_t& e = spec.literal[i];
    if (!r.check(check_record_size(e)) || !r.enum_ok(e.is_text, 1, kind_name) || !r.text_ok(e.text, kind_name)) {
      return;
    }
    formulon::ChartValue& v = points[i];
    v.idx = e.idx;
    if (e.is_text != 0) {
      v.kind = formulon::ChartValue::Kind::kText;
      v.text = text_or_empty(e.text);
    } else {
      v.kind = formulon::ChartValue::Kind::kNumber;
      v.number = e.number;
    }
  }
  out = formulon::ChartSeriesSource::Literal(std::move(points));
}

formulon::ChartSeriesPatch read_series(PatchReader& r, const fm_chart_series_patch_t& p) {
  formulon::ChartSeriesPatch out;
  read_source(r, p.name_op, p.name, out.name, "name_op", "name.kind");
  read_source(r, p.categories_op, p.categories, out.cat, "categories_op", "categories.kind");
  read_source(r, p.values_op, p.values, out.val, "values_op", "values.kind");
  read_source(r, p.x_op, p.x, out.x, "x_op", "x.kind");
  read_source(r, p.y_op, p.y, out.y, "y_op", "y.kind");
  read_source(r, p.size_op, p.size, out.size, "size_op", "size.kind");
  out.fill = read_fill(r, p.fill);
  out.line = read_line(r, p.line);
  r.field(p.marker_op, p.marker_op == FM_PATCH_OP_SET ? read_marker(r, p.marker) : formulon::ChartMarkerPatch{},
          out.marker, "marker_op");
  r.flag(p.smooth_op, p.smooth, out.smooth, "smooth_op");
  r.field(p.labels_op, p.labels_op == FM_PATCH_OP_SET ? read_labels(r, p.labels) : formulon::ChartLabelsPatch{},
          out.labels, "labels_op");
  return out;
}

/// Stores the formulas a patch enters, once its checks passed (see `ingest_entered_formula`).
void ingest_series_refs(fm_workbook_t* wb, formulon::ChartSeriesPatch& patch) {
  for (std::optional<formulon::ChartSeriesSource>* source :
       {&patch.name, &patch.cat, &patch.val, &patch.x, &patch.y, &patch.size}) {
    if (source->has_value() && (*source)->kind == formulon::ChartDataSource::Kind::kRef &&
        !(*source)->formula.empty()) {
      (*source)->formula = ingest_entered_formula(wb, (*source)->formula);
    }
  }
}

void ingest_title_ref(fm_workbook_t* wb, formulon::ChartTitle& title) {
  if (!title.ref_formula.empty()) {
    title.ref_formula = ingest_entered_formula(wb, title.ref_formula);
  }
}

/// Maps a validated `fm_tristate_t` to the engine's header detection mode.
formulon::ChartHeaderMode header_mode_from_c(std::int32_t value) {
  switch (value) {
    case FM_TRISTATE_TRUE:
      return formulon::ChartHeaderMode::kYes;
    case FM_TRISTATE_FALSE:
      return formulon::ChartHeaderMode::kNo;
    default:
      return formulon::ChartHeaderMode::kAuto;
  }
}

/// The insert template of a domain-checked C plot type and bar direction. Only `BAR` is an explicit direction:
/// `COLUMN` (zero) and `ABSENT` leave it to the type, so a zeroed `bar_direction` stays valid on a non-bar type.
formulon::Expected<formulon::ChartInsertKind, formulon::Error> insert_kind_of(std::int32_t type,
                                                                              std::int32_t bar_direction) {
  return formulon::chart_insert_kind(static_cast<formulon::ChartPlotType>(type),
                                     bar_direction == FM_CHART_BAR_DIRECTION_BAR
                                         ? formulon::ChartBarDirection::kBar
                                         : formulon::ChartBarDirection::kAbsent);
}

/// Shared entry checks of the edit functions.
fm_status_t edit_entry(const fm_workbook_t* wb, std::size_t sheet_index, const void* required) {
  clear_last_error();
  if (required == nullptr) {
    return null_argument();
  }
  return check_sheet_index(wb, sheet_index);
}

/// The axis' own `c:crossesAt` (NaN without one, which the edit refuses), for a crossing set to `AT` without a
/// value. An unknown chart or axis leaves NaN for the edit to report.
fm_status_t stored_crosses_at(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t object_id,
                              std::uint32_t axis_id, double* out) {
  *out = std::numeric_limits<double>::quiet_NaN();
  auto chart = formulon::get_chart(wb->workbook(), sheet_index, object_id);
  if (!chart) {
    return chart.error().code == formulon::FormulonErrorCode::kNotFound ? 0 : set_last_error(chart.error());
  }
  for (const formulon::ChartAxis& axis : chart.value().axes) {
    if (axis.id == axis_id) {
      *out = axis.crosses_at;
    }
  }
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// Creation and edits
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_sheet_add_chart(fm_workbook_t* wb, size_t sheet_index, const fm_chart_insert_spec_t* opts,
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
  if (auto rc = check_record_size(opts->placement); rc != 0) {
    return rc;
  }
  if (opts->source == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({opts->source, opts->title, opts->placement.name, opts->placement.description}); rc != 0) {
    return rc;
  }
  PatchReader r;
  r.enum_ok(opts->type, FM_CHART_PLOT_TYPE_EXTENDED, "type");
  r.enum_ok(opts->bar_direction, FM_CHART_BAR_DIRECTION_ABSENT, FM_CHART_BAR_DIRECTION_BAR, "bar_direction");
  r.enum_ok(opts->grouping, formulon::ChartGrouping::kStandard, "grouping");
  r.enum_ok(opts->direction, formulon::ChartSeriesDirection::kRows, "direction");
  r.enum_ok(opts->header_row, FM_TRISTATE_TRUE, "header_row");
  r.enum_ok(opts->header_column, FM_TRISTATE_TRUE, "header_column");
  r.enum_ok(opts->legend, FM_CHART_LEGEND_POSITION_ABSENT, FM_CHART_LEGEND_POSITION_TOP, "legend");
  r.enum_ok(opts->placement.hidden, FM_TRISTATE_TRUE, "placement.hidden");
  if (r.status() != 0) {
    return r.status();
  }
  auto kind = insert_kind_of(opts->type, opts->bar_direction);
  if (!kind) {
    return set_last_error(kind.error());
  }
  formulon::ChartInsertOptions model;
  model.kind = kind.value();
  model.grouping = static_cast<formulon::ChartGrouping>(opts->grouping);
  model.source = opts->source;
  model.direction = static_cast<formulon::ChartSeriesDirection>(opts->direction);
  model.header_row = header_mode_from_c(opts->header_row);
  model.header_column = header_mode_from_c(opts->header_column);
  if (opts->title != nullptr) {
    model.title = std::string(opts->title);
  }
  model.legend = chart_legend_position_from_c(opts->legend);
  AnchorPlan plan;
  if (auto rc = plan_add_anchor(wb->workbook(), sheet_index, opts->placement.anchor, formulon::AnchorKind::kTwoCell,
                                false, &plan);
      rc != 0) {
    return rc;
  }
  model.placement.anchor_kind = plan.kind;
  model.placement.edit_as = plan.edit_as;
  model.placement.row = plan.from.row;
  model.placement.col = plan.from.col;
  model.placement.row_off = plan.from.row_off;
  model.placement.col_off = plan.from.col_off;
  model.placement.width_emu = plan.cx;
  model.placement.height_emu = plan.cy;
  model.name = text_or_empty(opts->placement.name);
  model.descr = text_or_empty(opts->placement.description);
  auto id = formulon::insert_chart(wb->workbook(), sheet_index, model);
  if (!id) {
    return set_last_error(id.error());
  }
  *out_object_id = id.value();
  return apply_placement_extras(wb->workbook(), sheet_index, id.value(), opts->placement, false);
}

extern "C" fm_status_t fm_sheet_set_chart_type(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                               size_t plot_index, int32_t type, int32_t bar_direction,
                                               int32_t grouping) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  PatchReader r;
  r.enum_ok(type, FM_CHART_PLOT_TYPE_EXTENDED, "type");
  r.enum_ok(bar_direction, FM_CHART_BAR_DIRECTION_ABSENT, FM_CHART_BAR_DIRECTION_BAR, "bar_direction");
  r.enum_ok(grouping, formulon::ChartGrouping::kStandard, "grouping");
  if (r.status() != 0) {
    return r.status();
  }
  auto kind = insert_kind_of(type, bar_direction);
  if (!kind) {
    return set_last_error(kind.error());
  }
  return finish_status(formulon::set_chart_type(wb->workbook(), sheet_index, object_id, plot_index, kind.value(),
                                                static_cast<formulon::ChartGrouping>(grouping)));
}

extern "C" fm_status_t fm_sheet_set_chart_title(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                const fm_chart_title_t* title) {
  if (auto rc = edit_entry(wb, sheet_index, title); rc != 0) {
    return rc;
  }
  if (auto rc = check_enum_domain(title->kind, 0, FM_CHART_TITLE_KIND_AUTO, "kind"); rc != 0) {
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
                                                 const fm_chart_legend_t* legend) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  formulon::ChartLegendPatch model;
  if (legend == nullptr) {
    model.position = formulon::ChartLegendPosition::kNone;
    return finish_status(formulon::set_chart_legend(wb->workbook(), sheet_index, object_id, model));
  }
  if (auto rc = check_record_size(*legend); rc != 0) {
    return rc;
  }
  if (legend->hidden_entries == nullptr && legend->hidden_entries_count != 0) {
    return null_argument();
  }
  PatchReader r;
  r.enum_ok(legend->position, FM_CHART_LEGEND_POSITION_RIGHT, FM_CHART_LEGEND_POSITION_TOP, "position");
  r.enum_ok(legend->overlay, 1, "overlay");
  if (r.status() != 0) {
    return r.status();
  }
  model.position = chart_legend_position_from_c(legend->position);
  model.overlay = legend->overlay != 0;
  model.hidden_series_idx = std::vector<std::uint32_t>();
  if (legend->hidden_entries_count != 0) {
    model.hidden_series_idx->assign(legend->hidden_entries, legend->hidden_entries + legend->hidden_entries_count);
  }
  return finish_status(formulon::set_chart_legend(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_update_chart_legend(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    const fm_chart_legend_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  if (patch->hidden_entries == nullptr && patch->hidden_entries_count != 0) {
    return null_argument();
  }
  PatchReader r;
  // A legend always has a position and an overlay; removal is fm_sheet_set_chart_legend(..., NULL).
  std::optional<formulon::ChartLegendPosition> position;
  r.enum_optional(patch->position_op, patch->position, FM_CHART_LEGEND_POSITION_RIGHT, FM_CHART_LEGEND_POSITION_TOP,
                  chart_legend_position_from_c, position, "position");
  std::optional<bool> overlay;
  if (patch->overlay_op == FM_PATCH_OP_SET) {
    r.enum_ok(patch->overlay, 1, "overlay");
  }
  r.flag(patch->overlay_op, patch->overlay, overlay, "overlay");
  formulon::ChartLegendPatch model;
  std::vector<std::uint32_t> hidden;
  if (patch->hidden_entries_count != 0) {
    hidden.assign(patch->hidden_entries, patch->hidden_entries + patch->hidden_entries_count);
  }
  r.optional(patch->hidden_entries_op, std::move(hidden), model.hidden_series_idx, "hidden_entries");
  if (r.status() != 0) {
    return r.status();
  }
  // A KEEP field keeps the legend's value, or takes Excel's default on a chart without a legend.
  if (!position || !overlay) {
    auto chart = formulon::get_chart(wb->workbook(), sheet_index, object_id);
    if (!chart) {
      return set_last_error(chart.error());
    }
    const formulon::ChartLegend& current = chart.value().legend;
    model.position = current.present ? current.position : formulon::ChartLegendPosition::kRight;
    model.overlay = current.present && current.overlay;
  }
  if (position) {
    model.position = *position;
  }
  if (overlay) {
    model.overlay = *overlay;
  }
  return finish_status(formulon::set_chart_legend(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_update_chart_axis(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
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
  if (patch->title_op == FM_PATCH_OP_SET) {
    r.enum_ok(patch->title.kind, FM_CHART_TITLE_KIND_AUTO, "title.kind");
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
  r.enum_optional(patch->crosses_op, patch->crosses, FM_CHART_AXIS_CROSSES_ABSENT, FM_CHART_AXIS_CROSSES_AT,
                  chart_axis_crosses_from_c, model.crosses, "crosses");
  std::optional<double> crosses_at;
  r.optional(patch->crosses_at_op, patch->crosses_at, crosses_at, "crosses_at");
  r.token_field(patch->tick_label_position_op, patch->tick_label_position, kTickLabelPositionTokens,
                model.tick_label_pos, "tick_label_position");
  r.field(patch->minor_unit_op, patch->minor_unit, model.minor_unit, "minor_unit_op");
  r.flag(patch->minor_gridlines_op, patch->minor_gridlines, model.minor_gridlines, "minor_gridlines_op");
  r.enum_optional(patch->position_op, patch->position, formulon::ChartAxisPosition::kRight, model.position, "position");
  r.flag(patch->source_linked_op, patch->source_linked, model.source_linked, "source_linked_op");
  if (r.status() != 0) {
    return r.status();
  }
  if (crosses_at && !model.crosses) {
    // A crossing value alone moves the crossing there.
    model.crosses = formulon::ChartAxisCrosses::kAt;
  }
  if (model.crosses == formulon::ChartAxisCrosses::kAt) {
    if (crosses_at) {
      model.crosses_at = *crosses_at;
    } else if (auto rc = stored_crosses_at(wb, sheet_index, object_id, axis_id, &model.crosses_at); rc != 0) {
      return rc;
    }
  }
  if (model.title.is_set()) {
    ingest_title_ref(wb, model.title.value);
  }
  return finish_status(formulon::set_chart_axis(wb->workbook(), sheet_index, object_id, axis_id, model));
}

extern "C" fm_status_t fm_sheet_update_chart_display(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                     const fm_chart_display_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartDisplayPatch model;
  if (patch->blanks_as_op == FM_PATCH_OP_SET && patch->blanks_as == FM_BLANKS_AS_ABSENT) {
    r.fail("blanks_as", "cannot be set to ABSENT");
  }
  r.enum_field(patch->blanks_as_op, patch->blanks_as, FM_BLANKS_AS_ABSENT, FM_BLANKS_AS_SPAN, chart_blanks_as_from_c,
               model.blanks_as, "blanks_as");
  r.bool_field(patch->plot_visible_only_op, patch->plot_visible_only, model.plot_vis_only, "plot_visible_only");
  r.bool_field(patch->na_as_blank_op, patch->na_as_blank, model.na_as_blank, "na_as_blank");
  r.field(patch->style_op, patch->style, model.style, "style_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_display(wb->workbook(), sheet_index, object_id, model));
}

extern "C" fm_status_t fm_sheet_update_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    size_t plot_index, uint32_t series_id,
                                                    const fm_chart_series_patch_t* patch) {
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
      formulon::set_chart_series(wb->workbook(), sheet_index, object_id, plot_index, series_id, model));
}

extern "C" fm_status_t fm_sheet_update_chart_point(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                   size_t plot_index, uint32_t series_id, size_t point_index,
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
          patch->marker_op == FM_PATCH_OP_SET ? read_marker(r, patch->marker) : formulon::ChartMarkerPatch{},
          model.marker, "marker_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_point(wb->workbook(), sheet_index, object_id, plot_index, series_id,
                                                 saturate_index(point_index), model));
}

extern "C" fm_status_t fm_sheet_update_chart_plot(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                  size_t plot_index, const fm_chart_plot_patch_t* patch) {
  if (auto rc = edit_entry(wb, sheet_index, patch); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  PatchReader r;
  formulon::ChartPlotPatch model;
  r.optional(patch->gap_width_pct_op, patch->gap_width_pct, model.gap_width, "gap_width_pct_op");
  r.optional(patch->overlap_op, patch->overlap, model.overlap, "overlap_op");
  r.optional(patch->first_slice_angle_deg_op, patch->first_slice_angle_deg, model.first_slice_angle,
             "first_slice_angle_deg_op");
  r.flag(patch->vary_colors_op, patch->vary_colors, model.vary_colors, "vary_colors_op");
  r.optional(patch->hole_size_op, patch->hole_size, model.hole_size, "hole_size_op");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_plot(wb->workbook(), sheet_index, object_id, plot_index, model));
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

extern "C" fm_status_t fm_sheet_add_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                 size_t plot_index, const fm_chart_series_patch_t* patch,
                                                 uint32_t* out_series_id) {
  if (out_series_id == nullptr) {
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
  auto idx = formulon::add_chart_series(wb->workbook(), sheet_index, object_id, plot_index, model);
  if (!idx) {
    return set_last_error(idx.error());
  }
  *out_series_id = idx.value();
  return 0;
}

extern "C" fm_status_t fm_sheet_remove_chart_series(fm_workbook_t* wb, size_t sheet_index, uint32_t object_id,
                                                    size_t plot_index, uint32_t series_id) {
  if (auto rc = edit_entry(wb, sheet_index, wb); rc != 0) {
    return rc;
  }
  return finish_status(formulon::remove_chart_series(wb->workbook(), sheet_index, object_id, plot_index, series_id));
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
  r.enum_ok(header_row, FM_TRISTATE_TRUE, "header_row");
  r.enum_ok(header_column, FM_TRISTATE_TRUE, "header_column");
  if (r.status() != 0) {
    return r.status();
  }
  return finish_status(formulon::set_chart_source(wb->workbook(), sheet_index, object_id, source,
                                                  static_cast<formulon::ChartSeriesDirection>(direction),
                                                  header_mode_from_c(header_row), header_mode_from_c(header_column)));
}
