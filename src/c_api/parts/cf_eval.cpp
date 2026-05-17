// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - conditional-format evaluation (range walker).
//
// Bridges the engine-side `cf::evaluate_cf_for_range` walker into the
// stable C ABI. The opaque `fm_cf_results_t` owns the per-cell match
// list returned by the evaluator; index accessors copy the relevant
// fields out into the `fm_cf_match_t` POD so JS / Python / CLI consumers
// never see the internal `cf::CFMatch` layout.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <new>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cf/cf_evaluator.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

// Opaque results handle. Owns the engine-produced match list verbatim so
// the address / index accessors can return data without re-walking the
// evaluator. The engine returns one `CFRangeCellMatches` entry per cell
// that produced at least one match.
struct fm_cf_results {
  std::vector<formulon::cf::CFRangeCellMatches> matches;
};

namespace {

// Translates an engine-side `cf::Color` into the C ABI's RGBA POD.
fm_cf_color_t to_c_color(formulon::cf::Color c) {
  fm_cf_color_t out{};
  out.r = c.r;
  out.g = c.g;
  out.b = c.b;
  out.a = c.a;
  return out;
}

// Projects a single `cf::CFMatch` into the wire-format `fm_cf_match_t`.
// The output is zero-initialised first so non-active fields are
// deterministic (the public ABI promises default-zero for all
// kind-irrelevant payload).
void fill_match(const formulon::cf::CFMatch& match, fm_cf_match_t* out) {
  *out = fm_cf_match_t{};
  out->priority = match.priority;
  switch (match.kind) {
    case formulon::cf::CFMatchKind::DifferentialFormat:
      out->kind = FM_CF_DIFFERENTIAL_FORMAT;
      if (match.dxf_id.has_value()) {
        out->dxf_id_engaged = 1;
        out->dxf_id = *match.dxf_id;
      }
      return;
    case formulon::cf::CFMatchKind::ColorScale:
      out->kind = FM_CF_COLOR_SCALE;
      if (match.resolved_fill_color.has_value()) {
        out->color = to_c_color(*match.resolved_fill_color);
      }
      return;
    case formulon::cf::CFMatchKind::DataBar:
      out->kind = FM_CF_DATA_BAR;
      if (match.data_bar_render.has_value()) {
        const auto& bar = *match.data_bar_render;
        out->bar_length_pct = bar.length_pct;
        out->bar_axis_position_pct = bar.axis_position_pct;
        out->bar_is_negative = bar.is_negative ? 1 : 0;
        out->bar_fill = to_c_color(bar.fill);
        if (bar.border.has_value()) {
          out->bar_border_engaged = 1;
          out->bar_border = to_c_color(*bar.border);
        }
        out->bar_gradient = bar.gradient ? 1 : 0;
      }
      return;
    case formulon::cf::CFMatchKind::IconSet:
      out->kind = FM_CF_ICON_SET;
      if (match.icon_render.has_value()) {
        out->icon_set_name = static_cast<std::int32_t>(match.icon_render->set_name);
        out->icon_index = match.icon_render->icon_index;
      }
      return;
  }
  // Defensive default - every enumerator above returns. If a future
  // kind enumerator is added without a corresponding case the output
  // surfaces as DifferentialFormat with no engaged fields.
  out->kind = FM_CF_DIFFERENTIAL_FORMAT;
}

}  // namespace

extern "C" fm_status_t fm_workbook_cf_evaluate_range(const fm_workbook_t* wb, size_t sheet_index, uint32_t first_row,
                                                     uint32_t first_col, uint32_t last_row, uint32_t last_col,
                                                     double today_serial, fm_cf_results_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_cf_evaluate_range: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_workbook_cf_evaluate_range"); rc != 0) {
    return rc;
  }
  // Stack-allocated arena and eval context: the CF walker is purely
  // synchronous and the engine consumes both before returning. Binding
  // them here keeps the C ABI free of long-lived per-handle CF state.
  formulon::Arena arena;
  formulon::eval::EvalContext eval_ctx(wb->workbook().sheet(sheet_index));
  formulon::cf::CFHost host;
  host.arena = &arena;
  host.registry = &formulon::eval::default_registry();
  host.eval_ctx = &eval_ctx;
  host.today_serial = std::isnan(today_serial) ? std::optional<double>{} : std::optional<double>{today_serial};

  formulon::cf::CFCellRange range{};
  range.first = formulon::CellAddress{first_row, first_col};
  range.last = formulon::CellAddress{last_row, last_col};

  auto matches = formulon::cf::evaluate_cf_for_range(wb->workbook().sheet(sheet_index), range, host);

  auto handle = std::unique_ptr<fm_cf_results_t>(new fm_cf_results_t{});
  handle->matches = std::move(matches);
  *out = handle.release();
  return 0;
}

extern "C" void fm_cf_results_destroy(fm_cf_results_t* results) {
  // Mirrors `free(NULL)` semantics.
  delete results;
}

extern "C" size_t fm_cf_results_cell_count(const fm_cf_results_t* results) {
  if (results == nullptr) {
    return 0;
  }
  return results->matches.size();
}

extern "C" fm_status_t fm_cf_results_cell_at(const fm_cf_results_t* results, size_t cell_idx, uint32_t* out_row,
                                             uint32_t* out_col, size_t* out_match_count) {
  clear_last_error();
  if (results == nullptr || out_row == nullptr || out_col == nullptr || out_match_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cf_results_cell_at: NULL argument");
  }
  if (cell_idx >= results->matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_cell_at: cell_idx out of range",
        "cell_idx=" + std::to_string(cell_idx) + " cell_count=" + std::to_string(results->matches.size()));
  }
  const auto& entry = results->matches[cell_idx];
  *out_row = entry.cell.row;
  *out_col = entry.cell.col;
  *out_match_count = entry.matches.size();
  return 0;
}

extern "C" fm_status_t fm_cf_results_match_at(const fm_cf_results_t* results, size_t cell_idx, size_t match_idx,
                                              fm_cf_match_t* out) {
  clear_last_error();
  if (results == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cf_results_match_at: NULL argument");
  }
  if (cell_idx >= results->matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_match_at: cell_idx out of range",
        "cell_idx=" + std::to_string(cell_idx) + " cell_count=" + std::to_string(results->matches.size()));
  }
  const auto& entry = results->matches[cell_idx];
  if (match_idx >= entry.matches.size()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_cf_results_match_at: match_idx out of range",
        "match_idx=" + std::to_string(match_idx) + " match_count=" + std::to_string(entry.matches.size()));
  }
  fill_match(entry.matches[match_idx], out);
  return 0;
}
