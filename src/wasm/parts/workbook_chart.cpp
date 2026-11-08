//
// JsWorkbook chart creation and the typed edits. Patches are driven by field
// tables over the C structs; a patch field maps `undefined` to keep, `null` to
// clear and a value to set.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "wasm/parts/embind_common.h"
#include "wasm/parts/workbook.h"

namespace formulon {
namespace wasm {
namespace parts {

namespace {

using emscripten::val;

template <class T>
void store(void* base, size_t off, T v) {
  std::memcpy(static_cast<char*>(base) + off, &v, sizeof(T));
}

// ---- Patch input ----------------------------------------------------

enum class Wr : uint8_t { kU32, kI32, kF64, kBool, kStr, kTitle, kColor, kRef, kPlainF64, kObj, kObjOp };

/// One input property: the `_op` and value slots in the C patch struct. A `kObj` nests without an op,
/// a `kObjOp` carries one.
struct WrField {
  const char* key;
  uint16_t op;
  uint16_t value;
  Wr kind;
  const WrField* sub;
};

// clang-format off
#define WR(key, T, kind, op, value) {key, offsetof(T, op), offsetof(T, value), Wr::kind, nullptr}
#define WR_REF(key, T, member) {key, 0, offsetof(T, member), Wr::kRef, nullptr}
#define WR_PLAIN(key, T, member) {key, 0, offsetof(T, member), Wr::kPlainF64, nullptr}
#define WR_OBJ(key, T, member, table) {key, 0, offsetof(T, member), Wr::kObj, table}
#define WR_OBJ_OP(key, T, op, member, table) {key, offsetof(T, op), offsetof(T, member), Wr::kObjOp, table}
#define WR_END {nullptr, 0, 0, Wr::kU32, nullptr}
// clang-format on

constexpr WrField kFillWr[] = {
    WR("color", fm_chart_fill_patch_t, kColor, color_op, color),
    WR("transparencyPct", fm_chart_fill_patch_t, kF64, transparency_pct_op, transparency_pct),
    WR_END,
};

constexpr WrField kLineWr[] = {
    WR("color", fm_chart_line_patch_t, kColor, color_op, color),
    WR("transparencyPct", fm_chart_line_patch_t, kF64, transparency_pct_op, transparency_pct),
    WR("widthPt", fm_chart_line_patch_t, kF64, width_pt_op, width_pt),
    WR("dash", fm_chart_line_patch_t, kI32, dash_op, dash),
    WR("none", fm_chart_line_patch_t, kBool, none_op, none),
    WR_END,
};

constexpr WrField kMarkerWr[] = {
    WR("symbol", fm_chart_marker_patch_t, kI32, symbol_op, symbol),
    WR("sizePt", fm_chart_marker_patch_t, kF64, size_pt_op, size_pt),
    WR_OBJ("fill", fm_chart_marker_patch_t, fill, kFillWr),
    WR_OBJ("line", fm_chart_marker_patch_t, line, kLineWr),
    WR_END,
};

constexpr WrField kLabelsWr[] = {
    WR("showValue", fm_chart_labels_patch_t, kBool, show_value_op, show_value),
    WR("showCategory", fm_chart_labels_patch_t, kBool, show_category_op, show_category),
    WR("showSeries", fm_chart_labels_patch_t, kBool, show_series_op, show_series),
    WR("showPercent", fm_chart_labels_patch_t, kBool, show_percent_op, show_percent),
    WR("showLegendKey", fm_chart_labels_patch_t, kBool, show_legend_key_op, show_legend_key),
    WR("separator", fm_chart_labels_patch_t, kStr, separator_op, separator),
    WR("position", fm_chart_labels_patch_t, kI32, position_op, position),
    WR("numberFormat", fm_chart_labels_patch_t, kStr, number_format_op, number_format),
    WR("fontColor", fm_chart_labels_patch_t, kColor, font_color_op, font_color),
    WR("fontSizePt", fm_chart_labels_patch_t, kF64, font_size_op, font_size_pt),
    WR_END,
};

constexpr WrField kSeriesWr[] = {
    WR_REF("name", fm_chart_series_patch_t, name),
    WR_REF("categories", fm_chart_series_patch_t, categories),
    WR_REF("values", fm_chart_series_patch_t, values),
    WR_REF("x", fm_chart_series_patch_t, x),
    WR_REF("y", fm_chart_series_patch_t, y),
    WR_REF("size", fm_chart_series_patch_t, size),
    WR_OBJ("fill", fm_chart_series_patch_t, fill, kFillWr),
    WR_OBJ("line", fm_chart_series_patch_t, line, kLineWr),
    WR_OBJ_OP("marker", fm_chart_series_patch_t, marker_op, marker, kMarkerWr),
    WR("smooth", fm_chart_series_patch_t, kBool, smooth_op, smooth),
    WR_OBJ_OP("labels", fm_chart_series_patch_t, labels_op, labels, kLabelsWr),
    WR_END,
};

constexpr WrField kPointWr[] = {
    WR_OBJ("fill", fm_chart_point_patch_t, fill, kFillWr),
    WR_OBJ("line", fm_chart_point_patch_t, line, kLineWr),
    WR_OBJ_OP("marker", fm_chart_point_patch_t, marker_op, marker, kMarkerWr),
    WR_END,
};

constexpr WrField kPlotWr[] = {
    WR("gapWidthPct", fm_chart_plot_patch_t, kU32, gap_width_pct_op, gap_width_pct),
    WR("overlap", fm_chart_plot_patch_t, kI32, overlap_op, overlap),
    WR("firstSliceAngleDeg", fm_chart_plot_patch_t, kU32, first_slice_angle_deg_op, first_slice_angle_deg),
    WR("varyColors", fm_chart_plot_patch_t, kBool, vary_colors_op, vary_colors),
    WR("holeSize", fm_chart_plot_patch_t, kU32, hole_size_op, hole_size),
    WR_END,
};

constexpr WrField kAxisWr[] = {
    WR("visible", fm_chart_axis_patch_t, kBool, visible_op, visible),
    WR("title", fm_chart_axis_patch_t, kTitle, title_op, title),
    WR("min", fm_chart_axis_patch_t, kF64, min_op, min),
    WR("max", fm_chart_axis_patch_t, kF64, max_op, max),
    WR("majorUnit", fm_chart_axis_patch_t, kF64, major_unit_op, major_unit),
    WR("numberFormat", fm_chart_axis_patch_t, kStr, number_format_op, number_format),
    WR("majorGridlines", fm_chart_axis_patch_t, kBool, major_gridlines_op, major_gridlines),
    WR("reversed", fm_chart_axis_patch_t, kBool, reversed_op, reversed),
    WR("logBase", fm_chart_axis_patch_t, kF64, log_base_op, log_base),
    WR("crosses", fm_chart_axis_patch_t, kI32, crosses_op, crosses),
    WR_PLAIN("crossesAt", fm_chart_axis_patch_t, crosses_at),
    WR("tickLabelPosition", fm_chart_axis_patch_t, kI32, tick_label_position_op, tick_label_position),
    WR_END,
};

constexpr WrField kDisplayWr[] = {
    WR("blanksAs", fm_chart_display_patch_t, kI32, blanks_as_op, blanks_as),
    WR("plotVisibleOnly", fm_chart_display_patch_t, kBool, plot_visible_only_op, plot_visible_only),
    WR("naAsBlank", fm_chart_display_patch_t, kBool, na_as_blank_op, na_as_blank),
    WR_END,
};

/// Keeps the bytes behind the `const char*` fields of a patch alive.
using Arena = std::deque<std::string>;

/// A title given as its text, or in the read-back shape `{ kind, text, formula }`. An out-of-domain (absent)
/// kind is left for the C ABI to reject.
fm_chart_title_t read_title(JsNarrowNumericReader& r, Arena& arena, const val& v, const char* field) {
  fm_chart_title_t t{};
  if (v.typeOf().strictlyEquals(val("object"))) {
    t.kind = r.i32(v, "kind", -1, "kind");
    arena.push_back(r.string(v, "text", "text"));
    t.text = arena.back().c_str();
    arena.push_back(r.string(v, "formula", "formula"));
    t.formula = arena.back().c_str();
  } else {
    t.kind = FM_CHART_TITLE_KIND_TEXT;
    arena.push_back(r.string_value(v, field));
    t.text = arena.back().c_str();
  }
  return t;
}

/// Fills the zeroed patch struct at `base` from the properties of `owner`.
void read_patch(JsNarrowNumericReader& r, Arena& arena, const val& owner, void* base, const WrField* f) {
  for (; f->key != nullptr && r.ok(); ++f) {
    const val v = r.value(owner, f->key, f->key);
    if (!r.ok() || v.isUndefined()) {
      continue;
    }
    const bool clear = v.isNull();
    const int32_t op = clear ? FM_CHART_PATCH_OP_CLEAR : FM_CHART_PATCH_OP_SET;
    switch (f->kind) {
      case Wr::kRef:
        // A reference is kept as NULL, removed as "" and replaced by a string.
        arena.push_back(clear ? std::string() : r.string_value(v, f->key));
        store(base, f->value, arena.back().c_str());
        break;
      case Wr::kPlainF64:
        if (!clear) {
          store(base, f->value, r.number_value(v, 0.0, f->key));
        }
        break;
      case Wr::kObj:
        if (!clear) {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub);
        }
        break;
      case Wr::kObjOp:
        store(base, f->op, op);
        if (!clear) {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub);
        }
        break;
      default:
        store(base, f->op, op);
        if (clear) {
          break;
        }
        switch (f->kind) {
          case Wr::kU32:
            store(base, f->value, r.u32_value(v, 0U, f->key));
            break;
          case Wr::kI32:
            store(base, f->value, r.i32_value(v, 0, f->key));
            break;
          case Wr::kF64:
            store(base, f->value, r.number_value(v, 0.0, f->key));
            break;
          case Wr::kBool:
            store(base, f->value, static_cast<int32_t>(r.boolean_value(v, false)));
            break;
          case Wr::kTitle:
            store(base, f->value, read_title(r, arena, v, f->key));
            break;
          case Wr::kColor:
            store(base, f->value, js_pull_color_spec_value(v, f->key, &r));
            break;
          default:
            arena.push_back(r.string_value(v, f->key));
            store(base, f->value, arena.back().c_str());
            break;
        }
        break;
    }
  }
}

/// Reads a `number[]`; absent reads as empty, a non-array is rejected.
std::vector<uint32_t> read_indices(JsNarrowNumericReader& r, const val& list, const char* field) {
  std::vector<uint32_t> out;
  if (!js_value_present(list)) {
    return out;
  }
  // `array_element` reports a non-array as a type error.
  const uint32_t n = r.is_array(list, field) ? r.length(list, field) : 1U;
  for (uint32_t i = 0; i < n && r.ok(); ++i) {
    out.push_back(r.u32_value(r.array_element(list, i, field), 0U, field));
  }
  return out;
}

/// A `boolean | null | undefined` header flag as an `fm_tristate_t`.
int32_t header_mode(JsNarrowNumericReader& r, const val& v, const char* field) {
  if (!js_value_present(v)) {
    return FM_TRISTATE_DEFAULT;
  }
  return r.boolean_value(v, false) ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
}

const uint32_t* data_or_null(const std::vector<uint32_t>& v) {
  return v.empty() ? nullptr : v.data();
}

/// Reads `patch` (absent reads as `{}`) into `base`; false after an argument rejection.
bool read_patch_arg(JsNarrowNumericReader& r, const val& patch, void* base, const WrField* table, Arena& arena) {
  read_patch(r, arena, js_object_or_empty(patch), base, table);
  return r.ok();
}

}  // namespace

// ---- Creation and edits ---------------------------------------------

uint32_t JsWorkbook::insertChart(uint32_t sheet, val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  if (!options.typeOf().strictlyEquals(val("object"))) {
    record_type_error("addChart expects (sheet:number, options:object)");
    return 0U;
  }
  JsNarrowNumericReader reader("addChart");
  std::string name_store;
  std::string descr_store;
  std::string title_store;
  fm_chart_insert_t opts = FM_STRUCT_INIT(fm_chart_insert_t);
  opts.kind = reader.i32(options, "kind", FM_CHART_INSERT_KIND_COLUMN, "kind");
  opts.grouping = reader.i32(options, "grouping", FM_CHART_GROUPING_CLUSTERED, "grouping");
  const std::string source = reader.string(options, "source", "source");
  opts.source = source.c_str();
  opts.direction = reader.i32(options, "direction", FM_CHART_SERIES_DIRECTION_AUTO, "direction");
  opts.header_row = header_mode(reader, reader.value(options, "headerRow", "headerRow"), "headerRow");
  opts.header_column = header_mode(reader, reader.value(options, "headerColumn", "headerColumn"), "headerColumn");
  opts.title = reader.optional_string(options, "title", title_store, "title");
  // An omitted legend is Excel's default bottom legend; `null` is none.
  const val legend = reader.value(options, "legend", "legend");
  opts.legend = legend.isNull() ? FM_CHART_LEGEND_POSITION_NONE
                                : reader.i32_value(legend, FM_CHART_LEGEND_POSITION_BOTTOM, "legend");
  const val placement = reader.value(options, "placement", "placement");
  opts.placement = pull_image_insert(reader, placement, FM_ANCHOR_KIND_TWO_CELL, name_store, descr_store);
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t object_id = 0;
  if (!succeeded(fm_sheet_add_chart(handle_, sheet, &opts, &object_id))) {
    return 0U;
  }
  return object_id;
}

void JsWorkbook::setChartType(uint32_t sheet, uint32_t objectId, uint32_t plot, int32_t kind, val grouping) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartType");
  const int32_t group = reader.i32_value(grouping, FM_CHART_GROUPING_CLUSTERED, "grouping");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_chart_type(handle_, sheet, objectId, plot, kind, group));
}

void JsWorkbook::setChartTitle(uint32_t sheet, uint32_t objectId, val title) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartTitle");
  Arena arena;
  fm_chart_title_t t{};
  t.kind = FM_CHART_TITLE_KIND_NONE;
  if (js_value_present(title)) {
    t = read_title(reader, arena, title, "title");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_chart_title(handle_, sheet, objectId, &t));
}

void JsWorkbook::setChartLegend(uint32_t sheet, uint32_t objectId, val legend) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  // `null` removes the legend.
  if (legend.isNull()) {
    succeeded(fm_sheet_set_chart_legend(handle_, sheet, objectId, nullptr));
    return;
  }
  JsNarrowNumericReader reader("setChartLegend");
  const int32_t position = reader.i32(legend, "position", FM_CHART_LEGEND_POSITION_RIGHT, "position");
  const bool overlay = reader.boolean(legend, "overlay", false, "overlay");
  const std::vector<uint32_t> hidden =
      read_indices(reader, reader.value(legend, "hiddenSeries", "hiddenSeries"), "hiddenSeries");
  if (!reader.ok()) {
    return;
  }
  fm_chart_legend_t record = FM_STRUCT_INIT(fm_chart_legend_t);
  record.position = position;
  record.overlay = overlay ? 1 : 0;
  record.hidden_series = data_or_null(hidden);
  record.hidden_series_count = static_cast<uint32_t>(hidden.size());
  succeeded(fm_sheet_set_chart_legend(handle_, sheet, objectId, &record));
}

void JsWorkbook::updateChartLegend(uint32_t sheet, uint32_t objectId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updateChartLegend");
  const val spec = js_object_or_empty(patch);
  const val position_value = reader.value(spec, "position", "position");
  const val overlay_value = reader.value(spec, "overlay", "overlay");
  const val hidden_value = reader.value(spec, "hiddenSeries", "hiddenSeries");
  const std::vector<uint32_t> hidden = read_indices(reader, hidden_value, "hiddenSeries");
  if (!reader.ok()) {
    return;
  }
  // The C patch always writes position and overlay, so an omitted one is read back from the chart
  // (Excel's default legend when it has none).
  int32_t position = FM_CHART_LEGEND_POSITION_RIGHT;
  int32_t overlay = 0;
  if (!js_value_present(position_value) || !js_value_present(overlay_value)) {
    int32_t present = 0;
    fm_chart_t* chart = nullptr;
    if (!succeeded(fm_sheet_get_chart(handle_, sheet, objectId, &present, &chart))) {
      return;
    }
    fm_chart_info_t info = FM_STRUCT_INIT(fm_chart_info_t);
    const fm_status_t rc = present != 0 ? fm_chart_get_info(chart, &info) : 0;
    fm_chart_destroy(chart);
    if (!succeeded(rc)) {
      return;
    }
    if (info.legend_position != FM_CHART_LEGEND_POSITION_NONE) {
      position = info.legend_position;
      overlay = info.legend_overlay;
    }
  }
  if (js_value_present(position_value)) {
    position = reader.i32_value(position_value, position, "position");
  }
  if (js_value_present(overlay_value)) {
    overlay = reader.boolean_value(overlay_value, false) ? 1 : 0;
  }
  if (!reader.ok()) {
    return;
  }
  fm_chart_legend_patch_t p = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  p.position = position;
  p.overlay = overlay;
  if (js_value_present(hidden_value)) {
    p.hidden_series_op = FM_CHART_PATCH_OP_SET;
    p.hidden_series = data_or_null(hidden);
    p.hidden_series_count = static_cast<uint32_t>(hidden.size());
  }
  succeeded(fm_sheet_update_chart_legend(handle_, sheet, objectId, &p));
}

void JsWorkbook::setChartAxis(uint32_t sheet, uint32_t objectId, uint32_t axisId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartAxis");
  Arena arena;
  fm_chart_axis_patch_t p = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  if (!read_patch_arg(reader, patch, &p, kAxisWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_axis(handle_, sheet, objectId, axisId, &p));
}

void JsWorkbook::setChartDisplay(uint32_t sheet, uint32_t objectId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartDisplay");
  Arena arena;
  fm_chart_display_patch_t p = FM_STRUCT_INIT(fm_chart_display_patch_t);
  if (!read_patch_arg(reader, patch, &p, kDisplayWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_display(handle_, sheet, objectId, &p));
}

void JsWorkbook::setChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartSeries");
  Arena arena;
  fm_chart_series_patch_t p = FM_STRUCT_INIT(fm_chart_series_patch_t);
  if (!read_patch_arg(reader, patch, &p, kSeriesWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_series(handle_, sheet, objectId, plot, seriesIndex, &p));
}

void JsWorkbook::updateChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex,
                                  uint32_t pointIndex, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("updateChartPoint");
  Arena arena;
  fm_chart_point_patch_t p = FM_STRUCT_INIT(fm_chart_point_patch_t);
  if (!read_patch_arg(reader, patch, &p, kPointWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_point(handle_, sheet, objectId, plot, seriesIndex, pointIndex, &p));
}

void JsWorkbook::setChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex,
                               uint32_t pointIndex, val point) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartPoint");
  Arena arena;
  fm_chart_point_patch_t p = FM_STRUCT_INIT(fm_chart_point_patch_t);
  // `null` removes the point's formatting; a record replaces it.
  if (!point.isNull() && !read_patch_arg(reader, point, &p, kPointWr, arena)) {
    return;
  }
  // A patch that changes nothing removes the point's formatting, so the record is merged into a bare point.
  fm_chart_point_patch_t bare = FM_STRUCT_INIT(fm_chart_point_patch_t);
  if (!succeeded(fm_sheet_update_chart_point(handle_, sheet, objectId, plot, seriesIndex, pointIndex, &bare))) {
    return;
  }
  if (!point.isNull()) {
    succeeded(fm_sheet_update_chart_point(handle_, sheet, objectId, plot, seriesIndex, pointIndex, &p));
  }
}

void JsWorkbook::setChartPlot(uint32_t sheet, uint32_t objectId, uint32_t plot, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartPlot");
  Arena arena;
  fm_chart_plot_patch_t p = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  if (!read_patch_arg(reader, patch, &p, kPlotWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_plot(handle_, sheet, objectId, plot, &p));
}

void JsWorkbook::setChartFilter(uint32_t sheet, uint32_t objectId, val hiddenSeries, val hiddenCategories) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartFilter");
  const std::vector<uint32_t> series = read_indices(reader, hiddenSeries, "hiddenSeries");
  const std::vector<uint32_t> categories = read_indices(reader, hiddenCategories, "hiddenCategories");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_chart_filter(handle_, sheet, objectId, data_or_null(series), series.size(),
                                      data_or_null(categories), categories.size()));
}

uint32_t JsWorkbook::addChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  JsNarrowNumericReader reader("addChartSeries");
  Arena arena;
  fm_chart_series_patch_t p = FM_STRUCT_INIT(fm_chart_series_patch_t);
  if (!read_patch_arg(reader, patch, &p, kSeriesWr, arena)) {
    return 0U;
  }
  std::size_t series_index = 0;
  if (!succeeded(fm_sheet_add_chart_series(handle_, sheet, objectId, plot, &p, &series_index))) {
    return 0U;
  }
  return static_cast<uint32_t>(series_index);
}

void JsWorkbook::removeChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesIndex) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_chart_series(handle_, sheet, objectId, plot, seriesIndex)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setChartSource(uint32_t sheet, uint32_t objectId, const std::string& source, int32_t direction,
                                val headerRow, val headerColumn) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("setChartSource");
  const int32_t row = header_mode(reader, headerRow, "headerRow");
  const int32_t column = header_mode(reader, headerColumn, "headerColumn");
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_chart_source(handle_, sheet, objectId, source.c_str(), direction, row, column));
}

}  // namespace parts
}  // namespace wasm
}  // namespace formulon
