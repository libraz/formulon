//
// C ABI - recalc / iterative / calc-mode / profile / partial-recalc.

#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/recalc_reentry.h"
#include "eval/scheduler.h"
#include "utils/date_time.h"
#include "utils/error.h"
#include "utils/resource_budget.h"
#include "workbook.h"

using formulon::c_api::parts::check_finite_non_negative;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_sheet_rect;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::TextStore;

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
    return null_argument("wb");
  }
  auto r = wb->workbook().recalc(formulon::eval::default_registry());
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_has_dirty_cells(const fm_workbook_t* wb, int32_t* out_has_dirty) {
  clear_last_error();
  if (wb == nullptr || out_has_dirty == nullptr) {
    return null_argument();
  }
  // A recalc on this thread holds the engine and is mid-way through the dirty set.
  if (formulon::eval::detail::g_in_recalc) {
    return set_binding_error(formulon::FormulonErrorCode::kGraphRecalcReentrant,
                             "dirty cells cannot be read inside a recalc");
  }
  *out_has_dirty = wb->workbook().recalc_engine().dirty().size() != 0U ? 1 : 0;
  return 0;
}

namespace {

// Checks a circular-reference read of `sheet_index`: a sheet that exists, outside a recalc on this thread.
fm_status_t check_circular_reference_read(const fm_workbook_t* wb, std::size_t sheet_index) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (formulon::eval::detail::g_in_recalc) {
    return set_binding_error(formulon::FormulonErrorCode::kGraphRecalcReentrant,
                             "circular references cannot be read inside a recalc");
  }
  return 0;
}

std::vector<formulon::eval::CellNodeId> circular_reference_cells(const fm_workbook_t* wb, std::size_t sheet_index) {
  return wb->workbook().recalc_engine().unresolved_cycle_cells(static_cast<std::uint32_t>(sheet_index));
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_circular_reference_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                             std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_circular_reference_read(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = formulon::c_api::parts::ensure_calculated(wb); rc != 0) {
    return rc;
  }
  *out_count = circular_reference_cells(wb, sheet_index).size();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_circular_reference_at(const fm_workbook_t* wb, std::size_t sheet_index,
                                                          std::size_t index, fm_cell_address_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_circular_reference_read(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = formulon::c_api::parts::ensure_calculated(wb); rc != 0) {
    return rc;
  }
  const std::vector<formulon::eval::CellNodeId> cells = circular_reference_cells(wb, sheet_index);
  if (auto rc = formulon::c_api::parts::check_index(index, cells.size(), "index"); rc != 0) {
    return rc;
  }
  out->row = cells[index].row;
  out->col = cells[index].col;
  return 0;
}

extern "C" fm_status_t fm_sheet_is_cell_stale(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                              std::uint32_t col, int32_t* out_stale) {
  clear_last_error();
  if (wb == nullptr || out_stale == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = formulon::c_api::parts::check_cell_coord(row, col); rc != 0) {
    return rc;
  }
  if (formulon::eval::detail::g_in_recalc) {
    return set_binding_error(formulon::FormulonErrorCode::kGraphRecalcReentrant,
                             "stale cells cannot be read inside a recalc");
  }
  const formulon::Workbook& book = wb->workbook();
  const formulon::Cell* cell = book.sheet(sheet_index).cell_at(row, col);
  const bool formula = cell != nullptr && !cell->formula_text.empty();
  *out_stale = formula && book.calc_mode() != formulon::CalcMode::kAuto &&
                       book.recalc_engine().awaits_recalc(
                           formulon::eval::CellNodeId{static_cast<std::uint16_t>(sheet_index), row, col})
                   ? 1
                   : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_recalc_parallel(fm_workbook_t* wb, uint32_t thread_count,
                                                   fm_parallel_recalc_stats_info_t* out_stats) {
  clear_last_error();
  if (out_stats != nullptr) {
    if (auto rc = check_record_size(*out_stats); rc != 0) {
      return rc;
    }
    const std::uint32_t struct_size = out_stats->struct_size;
    *out_stats = FM_STRUCT_INIT(fm_parallel_recalc_stats_info_t);
    out_stats->struct_size = struct_size;
  }
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (thread_count > formulon::eval::kMaxSchedulerThreads) {
    return set_binding_error(formulon::FormulonErrorCode::kSecurityResourceLimit,
                             "thread_count is above the scheduler limit",
                             formulon::limit_context("threads", thread_count, formulon::eval::kMaxSchedulerThreads));
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

extern "C" fm_status_t fm_workbook_update_iterative(fm_workbook_t* wb, const fm_iterative_settings_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (patch == nullptr) {
    return null_argument("patch");
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  // No field can be cleared; every check runs before the write.
  if (!patch_op_ok_without_clear(patch->enabled_op)) {
    return invalid_argument("enabled_op is not a patch op", "enabled_op=" + std::to_string(patch->enabled_op));
  }
  if (!patch_op_ok_without_clear(patch->max_iterations_op)) {
    return invalid_argument("max_iterations_op is not a patch op",
                            "max_iterations_op=" + std::to_string(patch->max_iterations_op));
  }
  if (!patch_op_ok_without_clear(patch->max_change_op)) {
    return invalid_argument("max_change_op is not a patch op", "max_change_op=" + std::to_string(patch->max_change_op));
  }
  if (patch->max_iterations_op == FM_PATCH_OP_SET &&
      (patch->max_iterations == 0U || patch->max_iterations > formulon::kMaxIterationsCap)) {
    return invalid_argument(("max_iterations must be in 1.." + std::to_string(formulon::kMaxIterationsCap)).c_str(),
                            "max_iterations=" + std::to_string(patch->max_iterations));
  }
  if (patch->max_change_op == FM_PATCH_OP_SET) {
    if (auto rc = check_finite_non_negative(patch->max_change, "max_change"); rc != 0) {
      return rc;
    }
  }
  formulon::IterativeOptions opts = wb->workbook().iterative_options();
  if (patch->enabled_op == FM_PATCH_OP_SET) {
    opts.enabled = patch->enabled != 0;
  }
  if (patch->max_iterations_op == FM_PATCH_OP_SET) {
    opts.max_iterations = patch->max_iterations;
  }
  if (patch->max_change_op == FM_PATCH_OP_SET) {
    opts.max_change = patch->max_change;
  }
  wb->workbook().set_iterative_options(opts);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_iterative(const fm_workbook_t* wb, fm_iterative_settings_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::IterativeOptions& opts = wb->workbook().iterative_options();
  const std::uint32_t struct_size = out->struct_size;
  *out = FM_STRUCT_INIT(fm_iterative_settings_info_t);
  out->struct_size = struct_size;
  out->enabled = opts.enabled ? 1 : 0;
  out->max_iterations = opts.max_iterations;
  out->max_change = opts.max_change;
  return 0;
}

extern "C" fm_status_t fm_workbook_get_calc_mode(const fm_workbook_t* wb, std::int32_t* out_mode) {
  clear_last_error();
  if (wb == nullptr || out_mode == nullptr) {
    return null_argument();
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

extern "C" fm_status_t fm_workbook_get_compatibility_version(const fm_workbook_t* wb, std::int32_t* out_version) {
  clear_last_error();
  if (wb == nullptr || out_version == nullptr) {
    return null_argument();
  }
  const std::uint32_t version = wb->workbook().compatibility_version();
  // A value outside the enum (a later Excel's) evaluates as 2 and is kept in the file as read.
  *out_version = version >= 1U && version <= 3U ? static_cast<std::int32_t>(version) : FM_COMPATIBILITY_VERSION_2;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_compatibility_version(fm_workbook_t* wb, std::int32_t version) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (version < FM_COMPATIBILITY_VERSION_1 || version > FM_COMPATIBILITY_VERSION_3) {
    return invalid_argument(("version " + std::to_string(version) + " is not a CompatibilityVersion").c_str());
  }
  return formulon::c_api::parts::finish_status(
      wb->workbook().set_compatibility_version(static_cast<std::uint32_t>(version)));
}

extern "C" fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, std::int32_t mode) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
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
      return invalid_argument(("mode " + std::to_string(mode) + " is not a CalcMode").c_str());
  }
  wb->workbook().set_calc_mode(resolved);
  return 0;
}

namespace {

/// Reads a boolean workbook flag into an `int32_t` out-parameter.
template <typename Read>
fm_status_t read_workbook_flag(const fm_workbook_t* wb, std::int32_t* out, Read read) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  *out = read(wb->workbook()) ? 1 : 0;
  return 0;
}

template <typename Write>
fm_status_t write_workbook_flag(fm_workbook_t* wb, Write write) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  write(wb->workbook());
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_full_precision(const fm_workbook_t* wb, std::int32_t* out) {
  return read_workbook_flag(wb, out, [](const formulon::Workbook& w) { return w.calc_options().full_precision; });
}

extern "C" fm_status_t fm_workbook_set_full_precision(fm_workbook_t* wb, std::int32_t full_precision) {
  return write_workbook_flag(
      wb, [full_precision](formulon::Workbook& w) { w.mutable_calc_options().full_precision = full_precision != 0; });
}

extern "C" fm_status_t fm_workbook_get_calc_on_save(const fm_workbook_t* wb, std::int32_t* out) {
  return read_workbook_flag(wb, out, [](const formulon::Workbook& w) { return w.calc_options().calc_on_save; });
}

extern "C" fm_status_t fm_workbook_set_calc_on_save(fm_workbook_t* wb, std::int32_t calc_on_save) {
  return write_workbook_flag(
      wb, [calc_on_save](formulon::Workbook& w) { w.mutable_calc_options().calc_on_save = calc_on_save != 0; });
}

extern "C" fm_status_t fm_workbook_get_ref_mode(const fm_workbook_t* wb, std::int32_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  *out = wb->workbook().calc_options().ref_mode == formulon::RefMode::kR1C1 ? FM_REF_MODE_R1C1 : FM_REF_MODE_A1;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_ref_mode(fm_workbook_t* wb, std::int32_t mode) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (mode != FM_REF_MODE_A1 && mode != FM_REF_MODE_R1C1) {
    return invalid_argument(("mode " + std::to_string(mode) + " is not a RefMode").c_str());
  }
  wb->workbook().mutable_calc_options().ref_mode =
      mode == FM_REF_MODE_R1C1 ? formulon::RefMode::kR1C1 : formulon::RefMode::kA1;
  return 0;
}

extern "C" fm_status_t fm_workbook_get_date1904(const fm_workbook_t* wb, std::int32_t* out) {
  return read_workbook_flag(wb, out, [](const formulon::Workbook& w) { return w.date1904(); });
}

extern "C" fm_status_t fm_workbook_set_date1904(fm_workbook_t* wb, std::int32_t date1904) {
  return write_workbook_flag(wb, [date1904](formulon::Workbook& w) { w.set_date1904(date1904 != 0); });
}

extern "C" fm_status_t fm_workbook_get_pinned_now(const fm_workbook_t* wb, std::int32_t* out_present,
                                                  fm_civil_time_t* out_now) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out_now == nullptr) {
    return null_argument();
  }
  *out_now = fm_civil_time_t{};
  const auto& pinned = wb->workbook().pinned_now();
  *out_present = pinned.has_value() ? 1 : 0;
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
  if (wb == nullptr) {
    return null_argument();
  }
  if (now == nullptr) {
    wb->workbook().clear_pinned_now();
    return 0;
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
    return invalid_argument("now must be a date and time between 1900-01-01 and 9999-12-31");
  }
  wb->workbook().set_pinned_now(formulon::date_time::CivilTime{
      {now->year, static_cast<unsigned>(now->month), static_cast<unsigned>(now->day)},
      {static_cast<unsigned>(now->hour), static_cast<unsigned>(now->minute), static_cast<unsigned>(now->second)}});
  return 0;
}

extern "C" fm_status_t fm_workbook_get_profile_id(const fm_workbook_t* wb, const char** out_profile_id) {
  clear_last_error();
  if (wb == nullptr || out_profile_id == nullptr) {
    return null_argument();
  }
  *out_profile_id = formulon::excel_profile_id(wb->workbook().excel_profile());
  return 0;
}

extern "C" fm_status_t fm_workbook_set_profile_id(fm_workbook_t* wb, const char* profile_id) {
  clear_last_error();
  if (wb == nullptr || profile_id == nullptr) {
    return null_argument();
  }
  formulon::ExcelProfile profile;
  if (auto rc = check_profile_id(profile_id, &profile); rc != 0) {
    return rc;
  }
  formulon::Workbook& book = wb->workbook_before_refresh_on_load();
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

extern "C" fm_status_t fm_workbook_get_location(const fm_workbook_t* wb, std::int32_t* out_present,
                                                fm_workbook_location_t* out) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_workbook_location_t{};
  out->struct_size = struct_size;
  const std::optional<formulon::WorkbookLocation>& location = wb->workbook().location();
  *out_present = location ? 1 : 0;
  if (!location) {
    return 0;
  }
  TextStore& store = fresh_read_scratch(wb);
  out->directory = location->directory ? store_cstr(store, *location->directory) : nullptr;
  out->file_name = store_cstr(store, location->file_name);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_location(fm_workbook_t* wb, const fm_workbook_location_t* location) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  std::optional<formulon::WorkbookLocation> next;
  if (location != nullptr) {
    // The size is checked first so a short record's fields are never read.
    if (auto rc = check_record_size(*location); rc != 0) {
      return rc;
    }
    if (location->file_name == nullptr) {
      return null_argument();
    }
    next = formulon::WorkbookLocation{
        location->directory != nullptr ? std::optional<std::string>(location->directory) : std::nullopt,
        location->file_name};
  }
  // Nothing is marked dirty: the values that read the location (CELL,
  // FORMULATEXT) are volatile and refresh at the next recalculation, and a save
  // writes the values as computed, as Excel's Save As does.
  if (auto result = wb->workbook().set_location(std::move(next)); !result) {
    return set_last_error(result.error());
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_recalc_viewport(fm_workbook_t* wb, size_t sheet_index, const fm_cell_rect_t* viewport,
                                                uint32_t* out_recomputed_count) {
  clear_last_error();
  if (wb == nullptr || viewport == nullptr) {
    return null_argument();
  }
  // A sheet index past the end reaches the engine as a closure that
  // matches nothing, so the call would report a successful no-op recalc
  // and the caller would keep rendering stale values. Every other entry
  // point that takes a sheet rejects the same input.
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_rect(*viewport); rc != 0) {
    return rc;
  }
  // `SheetCellRange::sheet_id` is 16-bit; refuse an index it cannot hold
  // rather than recalculating a different sheet.
  if (sheet_index > 0xFFFFU) {
    return invalid_argument(("sheet_index " + std::to_string(sheet_index) + " is out of range (maximum 65535)").c_str(),
                            "sheet_index=" + std::to_string(sheet_index));
  }
  formulon::eval::SheetCellRange range;
  range.sheet_id = static_cast<std::uint16_t>(sheet_index);
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
    return null_argument("wb");
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
