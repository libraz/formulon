//
// JsWorkbook chart creation and the typed edits. Patches are driven by field
// tables over the C structs; a patch field maps `undefined` to keep, `null` to
// clear and a value to set.

#include <emscripten/val.h>

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <iterator>
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

enum class Wr : uint8_t { kU32, kI32, kF64, kBool, kStr, kTitle, kColor, kSource, kObj, kObjOp };

/// One input property: the `_op` and value slots in the C patch struct. A `kObj` nests without an op,
/// a `kObjOp` carries one, a `kSource` value is an `fm_chart_data_source_spec_t`.
struct WrField {
  const char* key;
  uint16_t op;
  uint16_t value;
  Wr kind;
  const WrField* sub;
  bool nullable;  ///< `null` clears the member; otherwise it is a type error.
};

// clang-format off
#define WR(key, T, kind, op, value) {key, offsetof(T, op), offsetof(T, value), Wr::kind, nullptr, false}
#define WR_SRC(key, T, op, member) {key, offsetof(T, op), offsetof(T, member), Wr::kSource, nullptr, false}
#define WR_OBJ(key, T, member, table) {key, 0, offsetof(T, member), Wr::kObj, table, false}
#define WR_OBJ_OP(key, T, op, member, table) {key, offsetof(T, op), offsetof(T, member), Wr::kObjOp, table, false}
// The same members as above, for a key the patch type declares clearable.
#define WRN(key, T, kind, op, value) {key, offsetof(T, op), offsetof(T, value), Wr::kind, nullptr, true}
#define WRN_SRC(key, T, op, member) {key, offsetof(T, op), offsetof(T, member), Wr::kSource, nullptr, true}
#define WRN_OBJ(key, T, member, table) {key, 0, offsetof(T, member), Wr::kObj, table, true}
#define WRN_OBJ_OP(key, T, op, member, table) {key, offsetof(T, op), offsetof(T, member), Wr::kObjOp, table, true}
#define WR_END {nullptr, 0, 0, Wr::kU32, nullptr, false}
// clang-format on

constexpr WrField kFillWr[] = {
    WRN("color", fm_chart_fill_patch_t, kColor, color_op, color),
    WRN("transparencyPct", fm_chart_fill_patch_t, kF64, transparency_pct_op, transparency_pct),
    WR("none", fm_chart_fill_patch_t, kBool, none_op, none),
    WR_END,
};

constexpr WrField kLineWr[] = {
    WRN("color", fm_chart_line_patch_t, kColor, color_op, color),
    WRN("transparencyPct", fm_chart_line_patch_t, kF64, transparency_pct_op, transparency_pct),
    WRN("widthPt", fm_chart_line_patch_t, kF64, width_pt_op, width_pt),
    WRN("dash", fm_chart_line_patch_t, kI32, dash_op, dash),
    WR("none", fm_chart_line_patch_t, kBool, none_op, none),
    WR_END,
};

constexpr WrField kMarkerWr[] = {
    WR("symbol", fm_chart_marker_patch_t, kI32, symbol_op, symbol),
    WRN("sizePt", fm_chart_marker_patch_t, kF64, size_pt_op, size_pt),
    WRN_OBJ("fill", fm_chart_marker_patch_t, fill, kFillWr),
    WRN_OBJ("line", fm_chart_marker_patch_t, line, kLineWr),
    WR_END,
};

constexpr WrField kLabelsWr[] = {
    WR("showValue", fm_chart_labels_patch_t, kBool, show_value_op, show_value),
    WR("showCategory", fm_chart_labels_patch_t, kBool, show_category_op, show_category),
    WR("showSeries", fm_chart_labels_patch_t, kBool, show_series_op, show_series),
    WR("showPercent", fm_chart_labels_patch_t, kBool, show_percent_op, show_percent),
    WR("showLegendKey", fm_chart_labels_patch_t, kBool, show_legend_key_op, show_legend_key),
    WRN("separator", fm_chart_labels_patch_t, kStr, separator_op, separator),
    WRN("position", fm_chart_labels_patch_t, kI32, position_op, position),
    WRN("numberFormat", fm_chart_labels_patch_t, kStr, number_format_op, number_format),
    WRN("fontColor", fm_chart_labels_patch_t, kColor, font_color_op, font_color),
    WRN("fontSizePt", fm_chart_labels_patch_t, kF64, font_size_pt_op, font_size_pt),
    WR("sourceLinked", fm_chart_labels_patch_t, kBool, source_linked_op, source_linked),
    WR("deleted", fm_chart_labels_patch_t, kBool, deleted_op, deleted),
    WR_END,
};

constexpr WrField kSeriesWr[] = {
    WRN_SRC("name", fm_chart_series_patch_t, name_op, name),
    WRN_SRC("categories", fm_chart_series_patch_t, categories_op, categories),
    WR_SRC("values", fm_chart_series_patch_t, values_op, values),
    WRN_SRC("x", fm_chart_series_patch_t, x_op, x),
    WR_SRC("y", fm_chart_series_patch_t, y_op, y),
    WR_SRC("size", fm_chart_series_patch_t, size_op, size),
    WRN_OBJ("fill", fm_chart_series_patch_t, fill, kFillWr),
    WRN_OBJ("line", fm_chart_series_patch_t, line, kLineWr),
    WRN_OBJ_OP("marker", fm_chart_series_patch_t, marker_op, marker, kMarkerWr),
    WR("smooth", fm_chart_series_patch_t, kBool, smooth_op, smooth),
    WRN_OBJ_OP("labels", fm_chart_series_patch_t, labels_op, labels, kLabelsWr),
    WR_END,
};

constexpr WrField kPointWr[] = {
    WRN_OBJ("fill", fm_chart_point_patch_t, fill, kFillWr),
    WRN_OBJ("line", fm_chart_point_patch_t, line, kLineWr),
    WRN_OBJ_OP("marker", fm_chart_point_patch_t, marker_op, marker, kMarkerWr),
    WR_END,
};

constexpr WrField kPlotWr[] = {
    WR("gapWidthPct", fm_chart_plot_patch_t, kU32, gap_width_pct_op, gap_width_pct),
    WR("overlapPct", fm_chart_plot_patch_t, kI32, overlap_pct_op, overlap_pct),
    WR("firstSliceAngleDeg", fm_chart_plot_patch_t, kU32, first_slice_angle_deg_op, first_slice_angle_deg),
    WR("varyColors", fm_chart_plot_patch_t, kBool, vary_colors_op, vary_colors),
    WR("holeSizePct", fm_chart_plot_patch_t, kU32, hole_size_pct_op, hole_size_pct),
    WR_END,
};

constexpr WrField kAxisWr[] = {
    WR("visible", fm_chart_axis_patch_t, kBool, visible_op, visible),
    WRN("title", fm_chart_axis_patch_t, kTitle, title_op, title),
    WRN("min", fm_chart_axis_patch_t, kF64, min_op, min),
    WRN("max", fm_chart_axis_patch_t, kF64, max_op, max),
    WRN("majorUnit", fm_chart_axis_patch_t, kF64, major_unit_op, major_unit),
    WRN("numberFormat", fm_chart_axis_patch_t, kStr, number_format_op, number_format),
    WR("majorGridlines", fm_chart_axis_patch_t, kBool, major_gridlines_op, major_gridlines),
    WR("reversed", fm_chart_axis_patch_t, kBool, reversed_op, reversed),
    WRN("logBase", fm_chart_axis_patch_t, kF64, log_base_op, log_base),
    WR("crosses", fm_chart_axis_patch_t, kI32, crosses_op, crosses),
    WR("crossesAt", fm_chart_axis_patch_t, kF64, crosses_at_op, crosses_at),
    WRN("tickLabelPosition", fm_chart_axis_patch_t, kI32, tick_label_position_op, tick_label_position),
    WRN("minorUnit", fm_chart_axis_patch_t, kF64, minor_unit_op, minor_unit),
    WR("minorGridlines", fm_chart_axis_patch_t, kBool, minor_gridlines_op, minor_gridlines),
    WR("position", fm_chart_axis_patch_t, kI32, position_op, position),
    WR("sourceLinked", fm_chart_axis_patch_t, kBool, source_linked_op, source_linked),
    WR_END,
};

constexpr WrField kDisplayWr[] = {
    WRN("blanksAs", fm_chart_display_patch_t, kI32, blanks_as_op, blanks_as),
    WRN("plotVisibleOnly", fm_chart_display_patch_t, kBool, plot_visible_only_op, plot_visible_only),
    WRN("naAsBlank", fm_chart_display_patch_t, kBool, na_as_blank_op, na_as_blank),
    WRN("style", fm_chart_display_patch_t, kI32, style_op, style),
    WR_END,
};

/// Keeps the bytes behind the `const char*` fields of a patch, and the literal points of its sources, alive.
struct Arena : std::deque<std::string> {
  std::deque<std::vector<fm_chart_literal_info_t>> literals;
};

/// A series input: reference text (empty is no input), or the read-back shape `{ kind, formula, literal }`.
fm_chart_data_source_spec_t read_source(JsNarrowNumericReader& r, Arena& arena, const val& v, const char* field) {
  fm_chart_data_source_spec_t spec{};
  if (!v.typeOf().strictlyEquals(val("object"))) {
    arena.push_back(r.string_value(v, field));
    spec.kind = arena.back().empty() ? FM_CHART_SOURCE_KIND_NONE : FM_CHART_SOURCE_KIND_REF;
    spec.formula = arena.back().c_str();
    return spec;
  }
  spec.kind = r.i32(v, "kind", FM_CHART_SOURCE_KIND_NONE, "kind");
  if (spec.kind == FM_CHART_SOURCE_KIND_REF) {
    arena.push_back(r.string(v, "formula", "formula"));
    spec.formula = arena.back().c_str();
  } else if (spec.kind == FM_CHART_SOURCE_KIND_LITERAL) {
    const val literal = r.value(v, "literal", "literal");
    std::vector<fm_chart_literal_info_t> points;
    if (r.ok() && js_value_present(literal) && r.require_array(literal, "literal")) {
      const uint32_t n = r.length(literal, "literal.length");
      for (uint32_t i = 0; i < n && r.ok(); ++i) {
        const val point = r.array_element(literal, i, "literal[]");
        fm_chart_literal_info_t lit = FM_STRUCT_INIT(fm_chart_literal_info_t);
        lit.idx = i;
        if (point.isString()) {
          arena.push_back(r.string_value(point, "literal[]"));
          lit.is_text = 1;
          lit.text = arena.back().c_str();
        } else {
          lit.number = r.number_value(point, 0.0, "literal[]");
        }
        points.push_back(lit);
      }
    }
    arena.literals.push_back(std::move(points));
    spec.literal = arena.literals.back().data();
    spec.literal_count = static_cast<uint32_t>(arena.literals.back().size());
  }
  return spec;
}

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

void clear_unset(void* base, const WrField* f);

/// Fills the zeroed patch struct at `base` from the properties of `owner`; `path` names `owner` in a rejection.
void read_patch(JsNarrowNumericReader& r, Arena& arena, const val& owner, void* base, const WrField* f,
                const std::string& path) {
  for (; f->key != nullptr && r.ok(); ++f) {
    const std::string label = path + "." + f->key;
    const char* const field = label.c_str();
    const val v = r.value(owner, f->key, field);
    if (!r.ok() || v.isUndefined()) {
      continue;
    }
    const bool clear = v.isNull();
    if (clear && !f->nullable) {
      r.fail_null(field);
      break;
    }
    const int32_t op = clear ? FM_PATCH_OP_CLEAR : FM_PATCH_OP_SET;
    switch (f->kind) {
      case Wr::kSource:
        store(base, f->op, op);
        if (!clear) {
          store(base, f->value, read_source(r, arena, v, field));
        }
        break;
      case Wr::kObj:
        // `null` clears every member, which removes the element (a fill or line then takes the style's paint).
        if (clear) {
          clear_unset(static_cast<char*>(base) + f->value, f->sub);
        } else {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub, label);
        }
        break;
      case Wr::kObjOp:
        store(base, f->op, op);
        if (!clear) {
          read_patch(r, arena, v, static_cast<char*>(base) + f->value, f->sub, label);
        }
        break;
      default:
        store(base, f->op, op);
        if (clear) {
          break;
        }
        switch (f->kind) {
          case Wr::kU32:
            store(base, f->value, r.u32_value(v, 0U, field));
            break;
          case Wr::kI32:
            store(base, f->value, r.i32_value(v, 0, field));
            break;
          case Wr::kF64:
            store(base, f->value, r.number_value(v, 0.0, field));
            break;
          case Wr::kBool:
            store(base, f->value, static_cast<int32_t>(r.boolean_value(v, false)));
            break;
          case Wr::kTitle:
            store(base, f->value, read_title(r, arena, v, field));
            break;
          case Wr::kColor:
            store(base, f->value, js_pull_color_spec_value(v, field, &r));
            break;
          default:
            arena.push_back(r.string_value(v, field));
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

/// Turns every `KEEP` member of a patch read from a whole record into `CLEAR`, so the record replaces the
/// stored value instead of merging into it.
void clear_unset(void* base, const WrField* f) {
  for (; f->key != nullptr; ++f) {
    switch (f->kind) {
      case Wr::kObj:
        clear_unset(static_cast<char*>(base) + f->value, f->sub);
        break;
      case Wr::kObjOp: {
        int32_t op = 0;
        std::memcpy(&op, static_cast<char*>(base) + f->op, sizeof(op));
        if (op == FM_PATCH_OP_KEEP) {
          store(base, f->op, static_cast<int32_t>(FM_PATCH_OP_CLEAR));
        } else if (op == FM_PATCH_OP_SET) {
          clear_unset(static_cast<char*>(base) + f->value, f->sub);
        }
        break;
      }
      default: {
        int32_t op = 0;
        std::memcpy(&op, static_cast<char*>(base) + f->op, sizeof(op));
        if (op == FM_PATCH_OP_KEEP) {
          store(base, f->op, static_cast<int32_t>(FM_PATCH_OP_CLEAR));
        }
        break;
      }
    }
  }
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

/// Reads `patch` (absent reads as `{}`), the argument named `param`, into `base`; false after an argument
/// rejection.
bool read_patch_arg(JsNarrowNumericReader& r, const val& patch, const char* param, void* base, const WrField* table,
                    Arena& arena) {
  read_patch(r, arena, js_object_or_empty(patch), base, table, param);
  return r.ok();
}

/// A table reading `owner.fill` / `owner.line` into a bare paint patch, which then sits at offset 0.
constexpr WrField kFillKeyWr[] = {
    {"fill", 0, 0, Wr::kObj, kFillWr, true},
    WR_END,
};
constexpr WrField kLineKeyWr[] = {
    {"line", 0, 0, Wr::kObj, kLineWr, true},
    WR_END,
};

/// Where a paint patch member lands in its record: the `has_` flag (`kNoFlag` when the member has none),
/// the value and its width. Entry i describes entry i of the matching `WrField` table.
struct RecSlot {
  uint16_t has;
  uint16_t value;
  uint16_t size;
};
constexpr uint16_t kNoFlag = 0xFFFFU;

// clang-format off
#define REC(T, member) {kNoFlag, offsetof(T, member), sizeof(T::member)}
#define REC_OPT(T, has, member) {offsetof(T, has), offsetof(T, member), sizeof(T::member)}
// clang-format on

constexpr RecSlot kFillRec[] = {
    REC(fm_chart_fill_t, color),
    REC_OPT(fm_chart_fill_t, has_transparency_pct, transparency_pct),
    REC(fm_chart_fill_t, none),
};
static_assert(std::size(kFillRec) + 1U == std::size(kFillWr), "one record slot per fill patch member");

constexpr RecSlot kLineRec[] = {
    REC(fm_chart_line_t, color),
    REC_OPT(fm_chart_line_t, has_transparency_pct, transparency_pct),
    REC_OPT(fm_chart_line_t, has_width_pt, width_pt),
    REC_OPT(fm_chart_line_t, has_dash, dash),
    REC(fm_chart_line_t, none),
};
static_assert(std::size(kLineRec) + 1U == std::size(kLineWr), "one record slot per line patch member");

/// Copies every `SET` member of the paint patch at `patch` into the zeroed record at `out`, raising its
/// `has_` flag; members left `KEEP` or `CLEAR` stay zero.
void patch_to_record(const void* patch, const WrField* f, const RecSlot* slot, void* out) {
  for (; f->key != nullptr; ++f, ++slot) {
    int32_t op = 0;
    std::memcpy(&op, static_cast<const char*>(patch) + f->op, sizeof(op));
    if (op != FM_PATCH_OP_SET) {
      continue;
    }
    if (slot->has != kNoFlag) {
      store(out, slot->has, int32_t{1});
    }
    std::memcpy(static_cast<char*>(out) + slot->value, static_cast<const char*>(patch) + f->value, slot->size);
  }
}

}  // namespace

// ---- Paint records ----------------------------------------------------

void pull_chart_fill_patch(JsNarrowNumericReader& reader, const val& owner, const char* param,
                           fm_chart_fill_patch_t* out) {
  Arena arena;
  read_patch(reader, arena, owner, out, kFillKeyWr, param);
}

void pull_chart_line_patch(JsNarrowNumericReader& reader, const val& owner, const char* param,
                           fm_chart_line_patch_t* out) {
  Arena arena;
  read_patch(reader, arena, owner, out, kLineKeyWr, param);
}

fm_chart_fill_t pull_chart_fill(JsNarrowNumericReader& reader, const val& owner, const char* param) {
  fm_chart_fill_patch_t p{};
  pull_chart_fill_patch(reader, owner, param, &p);
  fm_chart_fill_t out{};
  patch_to_record(&p, kFillWr, kFillRec, &out);
  return out;
}

fm_chart_line_t pull_chart_line(JsNarrowNumericReader& reader, const val& owner, const char* param) {
  fm_chart_line_patch_t p{};
  pull_chart_line_patch(reader, owner, param, &p);
  fm_chart_line_t out{};
  patch_to_record(&p, kLineWr, kLineRec, &out);
  return out;
}

// ---- Creation and edits ---------------------------------------------

uint32_t JsWorkbook::insertChart(uint32_t sheet, val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return 0U;
  }
  if (!options.typeOf().strictlyEquals(val("object"))) {
    record_type_error("Sheet.addChart input must be an object");
    return 0U;
  }
  JsNarrowNumericReader reader("Sheet.addChart");
  std::string name_store;
  std::string descr_store;
  Arena title_arena;
  fm_chart_spec_t opts = FM_STRUCT_INIT(fm_chart_spec_t);
  opts.type = reader.i32(options, "type", FM_CHART_PLOT_TYPE_BAR, "input.type");
  opts.bar_direction = reader.i32(options, "barDirection", FM_CHART_BAR_DIRECTION_COLUMN, "input.barDirection");
  opts.grouping = reader.i32(options, "grouping", FM_CHART_GROUPING_CLUSTERED, "input.grouping");
  const std::string source = reader.string(options, "source", "input.range");
  opts.source = source.c_str();
  opts.direction = reader.i32(options, "direction", FM_CHART_SERIES_DIRECTION_AUTO, "input.direction");
  opts.header_row = header_mode(reader, reader.value(options, "headerRow", "input.headerRow"), "input.headerRow");
  opts.header_column =
      header_mode(reader, reader.value(options, "headerColumn", "input.headerColumn"), "input.headerColumn");
  // An omitted title is Excel's automatic one; `null` is no title.
  const val title = reader.value(options, "title", "input.title");
  if (reader.ok() && !title.isUndefined()) {
    opts.has_title = 1;
    if (title.isNull()) {
      opts.title.kind = FM_CHART_TITLE_KIND_NONE;
    } else {
      opts.title = read_title(reader, title_arena, title, "input.title");
    }
  }
  // An omitted legend is Excel's default bottom legend; `null` is none.
  const val legend = reader.value(options, "legend", "input.legend");
  opts.legend = legend.isNull() ? FM_CHART_LEGEND_POSITION_ABSENT
                                : reader.i32_value(legend, FM_CHART_LEGEND_POSITION_BOTTOM, "input.legend");
  // The placement keys sit at the top level, beside the chart keys.
  opts.placement = pull_drawing_placement(reader, options, "input", name_store, descr_store);
  if (!reader.ok()) {
    return 0U;
  }
  uint32_t object_id = 0;
  if (!succeeded(fm_sheet_add_chart(handle_, sheet, &opts, &object_id))) {
    return 0U;
  }
  return object_id;
}

void JsWorkbook::setChartType(uint32_t sheet, uint32_t objectId, uint32_t plot, int32_t type, val options) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.setType");
  int32_t bar_direction = FM_CHART_BAR_DIRECTION_COLUMN;
  int32_t grouping = FM_CHART_GROUPING_CLUSTERED;
  if (js_value_present(options)) {
    bar_direction = reader.i32(options, "barDirection", FM_CHART_BAR_DIRECTION_COLUMN, "barDirection");
    grouping = reader.i32(options, "grouping", FM_CHART_GROUPING_CLUSTERED, "grouping");
  }
  if (!reader.ok()) {
    return;
  }
  succeeded(fm_sheet_set_chart_type(handle_, sheet, objectId, plot, type, bar_direction, grouping));
}

void JsWorkbook::setChartTitle(uint32_t sheet, uint32_t objectId, val title) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.setTitle");
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
  JsNarrowNumericReader reader("Chart.setLegend");
  const int32_t position = reader.i32(legend, "position", FM_CHART_LEGEND_POSITION_RIGHT, "position");
  const bool overlay = reader.boolean(legend, "overlay", false, "overlay");
  const std::vector<uint32_t> hidden =
      read_indices(reader, reader.value(legend, "hiddenEntries", "hiddenEntries"), "hiddenEntries");
  if (!reader.ok()) {
    return;
  }
  fm_chart_legend_t record = FM_STRUCT_INIT(fm_chart_legend_t);
  record.position = position;
  record.overlay = overlay ? 1 : 0;
  record.hidden_entries = data_or_null(hidden);
  record.hidden_entries_count = static_cast<uint32_t>(hidden.size());
  succeeded(fm_sheet_set_chart_legend(handle_, sheet, objectId, &record));
}

void JsWorkbook::updateChartLegend(uint32_t sheet, uint32_t objectId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updateLegend");
  const val spec = js_object_or_empty(patch);
  const val position_value = reader.value(spec, "position", "position");
  const val overlay_value = reader.value(spec, "overlay", "overlay");
  const val hidden_value = reader.value(spec, "hiddenEntries", "hiddenEntries");
  const std::vector<uint32_t> hidden = read_indices(reader, hidden_value, "hiddenEntries");
  if (!reader.ok()) {
    return;
  }
  fm_chart_legend_patch_t p = FM_STRUCT_INIT(fm_chart_legend_patch_t);
  if (js_value_present(position_value)) {
    p.position_op = FM_PATCH_OP_SET;
    p.position = reader.i32_value(position_value, 0, "position");
  }
  if (js_value_present(overlay_value)) {
    p.overlay_op = FM_PATCH_OP_SET;
    p.overlay = reader.boolean_value(overlay_value, false) ? 1 : 0;
  }
  if (!reader.ok()) {
    return;
  }
  if (js_value_present(hidden_value)) {
    p.hidden_entries_op = FM_PATCH_OP_SET;
    p.hidden_entries = data_or_null(hidden);
    p.hidden_entries_count = static_cast<uint32_t>(hidden.size());
  }
  succeeded(fm_sheet_update_chart_legend(handle_, sheet, objectId, &p));
}

void JsWorkbook::setChartAxis(uint32_t sheet, uint32_t objectId, uint32_t axisId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updateAxis");
  Arena arena;
  fm_chart_axis_patch_t p = FM_STRUCT_INIT(fm_chart_axis_patch_t);
  if (!read_patch_arg(reader, patch, "patch", &p, kAxisWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_axis(handle_, sheet, objectId, axisId, &p));
}

void JsWorkbook::setChartDisplay(uint32_t sheet, uint32_t objectId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updateDisplay");
  Arena arena;
  fm_chart_display_patch_t p = FM_STRUCT_INIT(fm_chart_display_patch_t);
  if (!read_patch_arg(reader, patch, "patch", &p, kDisplayWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_display(handle_, sheet, objectId, &p));
}

void JsWorkbook::setChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updateSeries");
  Arena arena;
  fm_chart_series_patch_t p = FM_STRUCT_INIT(fm_chart_series_patch_t);
  if (!read_patch_arg(reader, patch, "patch", &p, kSeriesWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_series(handle_, sheet, objectId, plot, seriesId, &p));
}

void JsWorkbook::updateChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId,
                                  uint32_t pointIndex, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updatePoint");
  Arena arena;
  fm_chart_point_patch_t p = FM_STRUCT_INIT(fm_chart_point_patch_t);
  if (!read_patch_arg(reader, patch, "patch", &p, kPointWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_point(handle_, sheet, objectId, plot, seriesId, pointIndex, &p));
}

void JsWorkbook::setChartPoint(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId, uint32_t pointIndex,
                               val point) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.setPoint");
  Arena arena;
  fm_chart_point_patch_t p = FM_STRUCT_INIT(fm_chart_point_patch_t);
  // A record replaces the point's formatting: members it carries are set, the rest cleared. `null` clears
  // every member, which removes the override.
  if (!point.isNull() && !read_patch_arg(reader, point, "point", &p, kPointWr, arena)) {
    return;
  }
  clear_unset(&p, kPointWr);
  succeeded(fm_sheet_update_chart_point(handle_, sheet, objectId, plot, seriesId, pointIndex, &p));
}

void JsWorkbook::setChartPlot(uint32_t sheet, uint32_t objectId, uint32_t plot, val patch) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.updatePlot");
  Arena arena;
  fm_chart_plot_patch_t p = FM_STRUCT_INIT(fm_chart_plot_patch_t);
  if (!read_patch_arg(reader, patch, "patch", &p, kPlotWr, arena)) {
    return;
  }
  succeeded(fm_sheet_update_chart_plot(handle_, sheet, objectId, plot, &p));
}

void JsWorkbook::setChartFilter(uint32_t sheet, uint32_t objectId, val hiddenSeries, val hiddenCategories) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.setFilter");
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
  JsNarrowNumericReader reader("Chart.addSeries");
  Arena arena;
  fm_chart_series_patch_t p = FM_STRUCT_INIT(fm_chart_series_patch_t);
  if (!read_patch_arg(reader, patch, "input", &p, kSeriesWr, arena)) {
    return 0U;
  }
  uint32_t series_id = 0;
  if (!succeeded(fm_sheet_add_chart_series(handle_, sheet, objectId, plot, &p, &series_id))) {
    return 0U;
  }
  return series_id;
}

void JsWorkbook::removeChartSeries(uint32_t sheet, uint32_t objectId, uint32_t plot, uint32_t seriesId) {
  succeeded(handle_ != nullptr ? fm_sheet_remove_chart_series(handle_, sheet, objectId, plot, seriesId)
                               : kBindingInvalidHandle);
}

void JsWorkbook::setChartSource(uint32_t sheet, uint32_t objectId, const std::string& source, int32_t direction,
                                val headerRow, val headerColumn) {
  if (handle_ == nullptr) {
    record_invalid_handle();
    return;
  }
  JsNarrowNumericReader reader("Chart.setSource");
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
