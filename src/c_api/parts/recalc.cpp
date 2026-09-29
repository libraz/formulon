//
// C ABI - recalc / iterative / calc-mode / profile / partial-recalc.

#include <cstdint>
#include <string>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::validate;

namespace {

// Engine-side shim for the C ABI iterative-solver progress callback. The
// engine passes the registering handle as `user_data`; the caller's own
// opaque pointer is stored beside the callback on that handle. Per the
// threading model (formulon_c.h), a caller must not race a clear against a
// solve on the same handle; the null check here is not a guard against
// such a race, only against `fm_workbook_set_iterative_progress(wb,
// nullptr, ...)` having cleared the callback before this solve started.
bool iterative_progress_adapter(std::uint32_t iteration, double max_residual, std::uint32_t max_iterations,
                                void* user_data) {
  auto* handle = static_cast<fm_workbook_t*>(user_data);
  if (handle == nullptr || handle->iterative_progress_cb == nullptr) {
    return true;
  }
  return handle->iterative_progress_cb(iteration, max_residual, max_iterations, handle->iterative_progress_user_data) !=
         0;
}

}  // namespace

// ---------------------------------------------------------------------------
// Recalc / iterative / calc-mode / profile / partial-recalc
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_recalc(fm_workbook_t* wb) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_recalc: wb is NULL");
  }
  auto r = wb->workbook().recalc(formulon::eval::default_registry());
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, uint32_t thread_count,
                                                   fm_parallel_recalc_stats* out_stats) {
  clear_last_error();
  if (out_stats != nullptr) {
    *out_stats = fm_parallel_recalc_stats{};
  }
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_recalc_parallel: wb is NULL");
  }
  if (thread_count > 8U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_recalc_parallel: thread_count must be 0..8",
                             "thread_count=" + std::to_string(thread_count) + " max=8");
  }

  formulon::eval::SchedulerConfig config;
  config.num_threads = thread_count;
  formulon::eval::SchedulerStats stats;
  auto result = wb->workbook().recalc_parallel(formulon::eval::default_registry(), config, &stats);
  if (!result) {
    // `stats` is intentionally not copied on failure. The public contract
    // promises an all-zero output for every failed entry, including an
    // evaluator / spill-scheduler failure after partial internal work.
    return set_last_error(result.error());
  }
  if (out_stats != nullptr) {
    out_stats->cells_evaluated = stats.cells_evaluated;
    out_stats->sccs_processed = stats.sccs_processed;
    out_stats->parallel_steps = stats.parallel_steps;
    out_stats->serial_fallback_steps = stats.serial_fallback_steps;
    out_stats->cycle_recoveries = stats.cycle_recoveries;
    out_stats->worker_threads_started = stats.worker_threads_started;
    out_stats->worker_threads_used = stats.worker_threads_used;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_iterative(fm_workbook_t* wb, int32_t enabled, int32_t max_iterations,
                                                 double max_change) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_iterative: wb is NULL");
  }
  formulon::IterativeOptions opts;
  opts.enabled = (enabled != 0);
  // Clamped into Excel's own dialog range rather than rejected: the
  // pre-existing contract for this argument is "out-of-range values are
  // adjusted", and a binding author who passes a large cap wants a long
  // solve, not an error. The upper bound is what keeps that from becoming
  // a hang the host cannot cancel without a progress callback.
  opts.max_iterations = max_iterations < 1 ? 1U
                        : static_cast<std::uint32_t>(max_iterations) > formulon::kMaxIterationsCap
                            ? formulon::kMaxIterationsCap
                            : static_cast<std::uint32_t>(max_iterations);
  opts.max_change = max_change;
  wb->workbook().set_iterative_options(opts);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_iterative_enabled(fm_workbook_t* wb, int32_t enabled) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_iterative_enabled: wb is NULL");
  }
  formulon::IterativeOptions opts = wb->workbook().iterative_options();
  opts.enabled = enabled != 0;
  wb->workbook().set_iterative_options(opts);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_iterative(const fm_workbook_t* wb, int32_t* out_enabled,
                                                 uint32_t* out_max_iterations, double* out_max_change) {
  clear_last_error();
  if (wb == nullptr || out_enabled == nullptr || out_max_iterations == nullptr || out_max_change == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_get_iterative: NULL argument");
  }
  const formulon::IterativeOptions& opts = wb->workbook().iterative_options();
  *out_enabled = opts.enabled ? 1 : 0;
  *out_max_iterations = opts.max_iterations;
  *out_max_change = opts.max_change;
  return 0;
}

extern "C" fm_status_t fm_workbook_calc_mode(const fm_workbook_t* wb, fm_calc_mode_t* out_mode) {
  clear_last_error();
  if (wb == nullptr || out_mode == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_calc_mode: NULL argument");
  }
  switch (wb->workbook().calc_mode()) {
    case formulon::Workbook::CalcMode::kAuto:
      *out_mode = FM_CALC_MODE_AUTO;
      break;
    case formulon::Workbook::CalcMode::kManual:
      *out_mode = FM_CALC_MODE_MANUAL;
      break;
    case formulon::Workbook::CalcMode::kAutoNoTable:
      *out_mode = FM_CALC_MODE_AUTO_NO_TABLE;
      break;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, std::int32_t mode) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_calc_mode: wb is NULL");
  }
  formulon::Workbook::CalcMode resolved = formulon::Workbook::CalcMode::kAuto;
  switch (mode) {
    case FM_CALC_MODE_AUTO:
      resolved = formulon::Workbook::CalcMode::kAuto;
      break;
    case FM_CALC_MODE_MANUAL:
      resolved = formulon::Workbook::CalcMode::kManual;
      break;
    case FM_CALC_MODE_AUTO_NO_TABLE:
      resolved = formulon::Workbook::CalcMode::kAutoNoTable;
      break;
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_set_calc_mode: unknown mode");
  }
  wb->workbook().set_calc_mode(resolved);
  return 0;
}

extern "C" fm_status_t fm_workbook_pinned_now(const fm_workbook_t* wb, fm_civil_time_t* out_now,
                                              std::int32_t* out_pinned) {
  clear_last_error();
  if (wb == nullptr || out_now == nullptr || out_pinned == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_pinned_now: NULL argument");
  }
  *out_now = fm_civil_time_t{};
  const auto& pinned = wb->workbook().pinned_now();
  *out_pinned = pinned.has_value() ? 1 : 0;
  if (pinned.has_value()) {
    out_now->year = pinned->date.y;
    out_now->month = static_cast<std::int32_t>(pinned->date.m);
    out_now->day = static_cast<std::int32_t>(pinned->date.d);
    out_now->hour = static_cast<std::int32_t>(pinned->time.h);
    out_now->minute = static_cast<std::int32_t>(pinned->time.m);
    out_now->second = static_cast<std::int32_t>(pinned->time.s);
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pinned_now(fm_workbook_t* wb, const fm_civil_time_t* now) {
  clear_last_error();
  if (wb == nullptr || now == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_pinned_now: NULL argument");
  }
  // Validated rather than normalised: `days_from_civil` would happily roll
  // month 13 into the next January, and a pin that silently moved would make
  // every result computed under it unexplainable to the host that set it.
  const bool in_range =
      now->year >= 1900 && now->year <= 9999 && now->month >= 1 && now->month <= 12 && now->day >= 1 &&
      now->day <=
          static_cast<std::int32_t>(formulon::date_time::days_in_month(now->year, static_cast<unsigned>(now->month))) &&
      now->hour >= 0 && now->hour <= 23 && now->minute >= 0 && now->minute <= 59 && now->second >= 0 &&
      now->second <= 59;
  if (!in_range) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_pinned_now: field out of range");
  }
  wb->workbook().set_pinned_now(formulon::date_time::CivilTime{
      {now->year, static_cast<unsigned>(now->month), static_cast<unsigned>(now->day)},
      {static_cast<unsigned>(now->hour), static_cast<unsigned>(now->minute), static_cast<unsigned>(now->second)}});
  return 0;
}

extern "C" fm_status_t fm_workbook_clear_pinned_now(fm_workbook_t* wb) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_clear_pinned_now: wb is NULL");
  }
  wb->workbook().clear_pinned_now();
  return 0;
}

extern "C" fm_status_t fm_workbook_excel_profile_id(const fm_workbook_t* wb, const char** out_profile_id) {
  clear_last_error();
  if (wb == nullptr || out_profile_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_excel_profile_id: NULL argument");
  }
  *out_profile_id = formulon::excel_profile_id(wb->workbook().excel_profile());
  return 0;
}

extern "C" fm_status_t fm_workbook_set_excel_profile_id(fm_workbook_t* wb, const char* profile_id) {
  clear_last_error();
  if (wb == nullptr || profile_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_excel_profile_id: NULL argument");
  }
  formulon::ExcelProfile profile;
  if (!formulon::parse_excel_profile_id(profile_id, &profile)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_excel_profile_id: unknown profile");
  }
  formulon::Workbook& book = wb->workbook();
  const bool changed = !formulon::same_profile(book.excel_profile(), profile);
  book.set_excel_profile(profile);
  if (changed) {
    // Profile-sensitive evaluation (SUMIF/COUNTIF criteria matching,
    // INFO/CELL, other host-dependent branches) is not confined to a
    // enumerable set of function names the way SUBTOTAL/AGGREGATE are, so
    // every formula cell needs the recompute rather than a scoped subset.
    book.mark_all_formulas_dirty();
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_partial_recalc(fm_workbook_t* wb, const fm_viewport* viewport,
                                                  uint32_t* out_recomputed_count) {
  clear_last_error();
  if (wb == nullptr || viewport == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_partial_recalc: NULL argument");
  }
  // A sheet index past the end reaches the engine as a closure that
  // matches nothing, so the call would report a successful no-op recalc
  // and the caller would keep rendering stale values. Every other entry
  // point that takes a sheet rejects the same input.
  if (auto rc = check_sheet_index(wb, viewport->sheet, "fm_workbook_partial_recalc"); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*viewport, "fm_workbook_partial_recalc"); rc != 0) {
    return rc;
  }
  formulon::eval::SheetCellRange range;
  range.sheet_id = static_cast<std::uint16_t>(viewport->sheet);
  range.first_row = viewport->first_row;
  range.last_row = viewport->last_row;
  range.first_col = viewport->first_col;
  range.last_col = viewport->last_col;
  auto r = wb->workbook().partial_recalc(formulon::eval::default_registry(), range);
  if (!r) {
    return set_last_error(r.error());
  }
  if (out_recomputed_count != nullptr) {
    *out_recomputed_count = r.value().cells_evaluated;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_iterative_progress(fm_workbook_t* wb, fm_iterative_progress_cb cb,
                                                          void* user_data) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_iterative_progress: wb is NULL");
  }
  wb->iterative_progress_cb = cb;
  wb->iterative_progress_user_data = cb == nullptr ? nullptr : user_data;
  // The engine's callback returns `bool`, which no C ABI declaration uses.
  // Hand it the adapter above instead of the caller's pointer so both sides
  // keep their own return type and neither assignment needs a cast.
  const formulon::eval::IterativeProgressCb engine_cb = cb == nullptr ? nullptr : &iterative_progress_adapter;
  wb->workbook().set_iterative_progress(engine_cb, cb == nullptr ? nullptr : wb);
  return 0;
}
