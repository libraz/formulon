//
// C ABI - workbook lifecycle, save/load, sheet management, recalc /
// iterative / partial-recalc / calc-mode / profile, defined names,
// tables, passthrough parts, structural row/column insertion + deletion.

#include "workbook.h"

#include <cstdint>
#include <cstring>
#include <memory>
#include <new>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

// ---------------------------------------------------------------------------
// Construction / lifecycle
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_create(fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_create: out is NULL");
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(formulon::Workbook::create());
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_create_empty(fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_create_empty: out is NULL");
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(formulon::Workbook::create_empty());
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_load: NULL or empty input");
  }
  *out = nullptr;
  if (bytes == nullptr || len == 0) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_load: NULL or empty input");
  }
  formulon::io::ByteSpan span;
  span.data = bytes;
  span.size = len;
  // Detect the container format from the package bytes (the C ABI takes
  // bytes, not a path, so extension-based routing is impossible). An
  // `.xlsb` package declares the binary `xl/workbook.bin` workbook part;
  // `.xlsx` declares `xl/workbook.xml`. `Unknown` falls through to the
  // OOXML reader, which owns the authoritative "not a workbook" /
  // encryption / corruption diagnostics.
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  if (formulon::io::detect_workbook_format(span) == formulon::io::WorkbookFormat::Xlsb) {
    auto result = formulon::io::xlsb::read_xlsb(span);
    if (!result) {
      return set_last_error(result.error());
    }
    formulon::io::xlsb::XlsbReadResult read_result = std::move(result.value());
    handle->xlsb_undecoded_formula_count = read_result.undecoded_formula_count;
    handle->xlsb_undecoded_defined_name_count = read_result.undecoded_defined_name_count;
    handle->wb.emplace(std::move(read_result.workbook));
  } else {
    auto result = formulon::io::read_ooxml(span);
    if (!result) {
      return set_last_error(result.error());
    }
    // The workbook now owns the text-storage deque that backs every
    // Text-cell `string_view`, so we move only the workbook out of the
    // result; the rest of `OoxmlReadResult` (passthrough parts mirror,
    // audit counter) is discarded.
    handle->wb.emplace(std::move(result.value().workbook));
  }
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_xlsb_read_diagnostics(const fm_workbook_t* wb, size_t* out_undecoded_formula_count,
                                                         size_t* out_undecoded_defined_name_count) {
  clear_last_error();
  if (wb == nullptr || out_undecoded_formula_count == nullptr || out_undecoded_defined_name_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_xlsb_read_diagnostics: NULL argument");
  }
  *out_undecoded_formula_count = wb->xlsb_undecoded_formula_count;
  *out_undecoded_defined_name_count = wb->xlsb_undecoded_defined_name_count;
  return 0;
}

extern "C" void fm_workbook_destroy(fm_workbook_t* wb) {
  // Mirrors `free(NULL)` semantics: silently accept NULL handles.
  delete wb;
}

// ---------------------------------------------------------------------------
// Save
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_save(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || out_len == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_save: NULL argument");
  }
  auto bytes = wb->workbook().save();
  if (!bytes) {
    return set_last_error(bytes.error());
  }
  const std::vector<std::uint8_t>& src = bytes.value();
  // Allocate with `new[]` so `fm_buffer_free`'s matching `delete[]` is
  // well-defined. The header documents the pairing.
  auto* buffer = new uint8_t[src.size()];
  if (!src.empty()) {
    std::memcpy(buffer, src.data(), src.size());
  }
  *out_bytes = buffer;
  *out_len = src.size();
  return 0;
}

extern "C" fm_status_t fm_workbook_save_ex(const fm_workbook_t* wb, fm_workbook_format_t format, uint8_t** out_bytes,
                                           size_t* out_len) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || out_len == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_save_ex: NULL argument");
  }
  formulon::io::WorkbookFormat engine_format;
  switch (format) {
    case FM_WORKBOOK_FORMAT_XLSX:
      engine_format = formulon::io::WorkbookFormat::Ooxml;
      break;
    case FM_WORKBOOK_FORMAT_XLSB:
      engine_format = formulon::io::WorkbookFormat::Xlsb;
      break;
    case FM_WORKBOOK_FORMAT_UNKNOWN:
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_save_ex: unsupported format",
                               "format=" + std::to_string(static_cast<int>(format)));
  }
  auto bytes = wb->workbook().save_ex(engine_format);
  if (!bytes) {
    return set_last_error(bytes.error());
  }
  const std::vector<std::uint8_t>& src = bytes.value();
  auto* buffer = new uint8_t[src.size()];
  if (!src.empty()) {
    std::memcpy(buffer, src.data(), src.size());
  }
  *out_bytes = buffer;
  *out_len = src.size();
  return 0;
}

extern "C" fm_status_t fm_workbook_save_xlsb_with_result(const fm_workbook_t* wb, uint8_t** out_bytes, size_t* out_len,
                                                         size_t* out_downgraded_formula_count) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || out_len == nullptr || out_downgraded_formula_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_save_xlsb_with_result: NULL argument");
  }
  auto result_or = formulon::io::xlsb::write_xlsb_with_result(wb->workbook());
  if (!result_or) {
    return set_last_error(result_or.error());
  }
  formulon::io::xlsb::XlsbWriteResult result = std::move(result_or.value());
  auto* buffer = new uint8_t[result.bytes.size()];
  if (!result.bytes.empty()) {
    std::memcpy(buffer, result.bytes.data(), result.bytes.size());
  }
  *out_bytes = buffer;
  *out_len = result.bytes.size();
  *out_downgraded_formula_count = result.downgraded_formula_count;
  return 0;
}

extern "C" void fm_buffer_free(uint8_t* bytes) {
  delete[] bytes;
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------
//
// `fm_workbook_sheet_count` is now emitted by the binding codegen (see
// `src/c_api/generated/workbook_counts.cpp`).

extern "C" fm_status_t fm_workbook_sheet_name(const fm_workbook_t* wb, size_t index, const char** out_utf8) {
  clear_last_error();
  if (wb == nullptr || out_utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_sheet_name: NULL argument");
  }
  if (index >= wb->workbook().sheet_count()) {
    return set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_sheet_name: sheet_index out of range",
        "sheet_index=" + std::to_string(index) + " sheet_count=" + std::to_string(wb->workbook().sheet_count()));
  }
  // `Sheet::name()` returns `const std::string&`, so `c_str()` is
  // NUL-terminated and stable until the sheet is mutated or destroyed.
  *out_utf8 = wb->workbook().sheet(index).name().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* utf8_name) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_add_sheet: NULL argument");
  }
  auto r = wb->workbook().add_sheet_validated(std::string(utf8_name));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, uint32_t from_index, uint32_t to_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_move_sheet: wb is NULL");
  }
  auto r = wb->workbook().move_sheet(from_index, to_index);
  if (!r) {
    return set_last_error(r.error());
  }
  // The enumeration cache is keyed by sheet index. Moving sheets can put a
  // different Sheet at a cached index with the same cell revision.
  wb->cell_enumeration_cache.addresses.clear();
  wb->cell_enumeration_cache.sheet_index = std::numeric_limits<std::size_t>::max();
  wb->cell_enumeration_cache.revision = std::numeric_limits<std::uint64_t>::max();
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, uint32_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_remove_sheet: wb is NULL");
  }
  auto r = wb->workbook().remove_sheet(index);
  if (!r) {
    return set_last_error(r.error());
  }
  // Removing a sheet shifts later sheet indices, so invalidate the
  // index-keyed coordinate cache even though the remaining Sheet objects
  // themselves were not mutated.
  wb->cell_enumeration_cache.addresses.clear();
  wb->cell_enumeration_cache.sheet_index = std::numeric_limits<std::size_t>::max();
  wb->cell_enumeration_cache.revision = std::numeric_limits<std::uint64_t>::max();
  return 0;
}

extern "C" fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, uint32_t index, const char* new_name) {
  clear_last_error();
  if (wb == nullptr || new_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_rename_sheet: NULL argument");
  }
  auto r = wb->workbook().rename_sheet(index, std::string(new_name));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_defined_name: NULL argument");
  }
  auto r = wb->workbook().set_defined_name(std::string(name), std::string(formula));
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_defined_name_scoped(fm_workbook_t* wb, const char* name, const char* formula,
                                                           int32_t local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_defined_name_scoped: NULL argument");
  }
  auto r = wb->workbook().set_defined_name_scoped(std::string(name), std::string(formula), local_sheet_id);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_insert_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_insert_rows: NULL argument");
  }
  auto r = wb->workbook().insert_rows(sheet, row, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_delete_rows(fm_workbook_t* wb, uint32_t sheet, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_delete_rows: NULL argument");
  }
  auto r = wb->workbook().delete_rows(sheet, row, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_insert_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_insert_cols: NULL argument");
  }
  auto r = wb->workbook().insert_cols(sheet, col, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_delete_cols(fm_workbook_t* wb, uint32_t sheet, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_delete_cols: NULL argument");
  }
  auto r = wb->workbook().delete_cols(sheet, col, count);
  if (!r) {
    return set_last_error(r.error());
  }
  return 0;
}

// ---------------------------------------------------------------------------
// Defined names / tables / passthrough parts (read-side iteration)
// ---------------------------------------------------------------------------
//
// `fm_workbook_defined_name_count`, `fm_workbook_table_count`, and
// `fm_workbook_passthrough_count` are now emitted by the binding
// codegen (see `src/c_api/generated/workbook_counts.cpp`).

extern "C" fm_status_t fm_workbook_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                   const char** out_formula) {
  return fm_workbook_defined_name_at_ex(wb, idx, out_name, out_formula, nullptr);
}

extern "C" fm_status_t fm_workbook_defined_name_at_ex(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                      const char** out_formula, int32_t* out_local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_defined_name_at_ex: NULL argument");
  }
  const auto& names = wb->workbook().defined_names();
  if (idx >= names.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_defined_name_at_ex: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(names.size()));
  }
  *out_name = names[idx].name.c_str();
  *out_formula = names[idx].formula.c_str();
  if (out_local_sheet_id != nullptr) {
    *out_local_sheet_id = names[idx].local_sheet_id;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_table_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                            const char** out_display_name, const char** out_ref,
                                            size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_display_name == nullptr || out_ref == nullptr ||
      out_sheet_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_table_at: NULL argument");
  }
  const auto& tables = wb->workbook().tables();
  if (idx >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_workbook_table_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(tables.size()));
  }
  *out_name = tables[idx].name.c_str();
  *out_display_name = tables[idx].display_name.c_str();
  *out_ref = tables[idx].ref.c_str();
  *out_sheet_index = tables[idx].sheet_index;
  return 0;
}

extern "C" fm_status_t fm_workbook_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path) {
  clear_last_error();
  if (wb == nullptr || out_path == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_passthrough_at: NULL argument");
  }
  const auto& parts = wb->workbook().passthrough_parts();
  if (idx >= parts.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_passthrough_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(parts.size()));
  }
  *out_path = parts[idx].path.c_str();
  return 0;
}

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

extern "C" fm_status_t fm_workbook_set_iterative(fm_workbook_t* wb, int32_t enabled, int32_t max_iterations,
                                                 double max_change) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_set_iterative: wb is NULL");
  }
  formulon::eval::IterativeOptions opts;
  opts.enabled = (enabled != 0);
  opts.max_iterations = max_iterations < 1 ? 1U : static_cast<std::uint32_t>(max_iterations);
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
  formulon::eval::IterativeOptions opts = wb->workbook().iterative_options();
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
  const formulon::eval::IterativeOptions& opts = wb->workbook().iterative_options();
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

extern "C" fm_status_t fm_workbook_set_calc_mode(fm_workbook_t* wb, fm_calc_mode_t mode) {
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

extern "C" fm_status_t fm_workbook_excel_profile_id(const fm_workbook_t* wb, const char** out_profile_id) {
  clear_last_error();
  if (wb == nullptr || out_profile_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_excel_profile_id: NULL argument");
  }
  *out_profile_id = formulon::eval::excel_profile_id(wb->workbook().excel_profile());
  return 0;
}

extern "C" fm_status_t fm_workbook_set_excel_profile_id(fm_workbook_t* wb, const char* profile_id) {
  clear_last_error();
  if (wb == nullptr || profile_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_excel_profile_id: NULL argument");
  }
  formulon::eval::ExcelProfile profile;
  if (!formulon::eval::parse_excel_profile_id(profile_id, &profile)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_excel_profile_id: unknown profile");
  }
  wb->workbook().set_excel_profile(profile);
  return 0;
}

extern "C" fm_status_t fm_workbook_partial_recalc(fm_workbook_t* wb, const fm_viewport* viewport,
                                                  uint32_t* out_recomputed_count) {
  clear_last_error();
  if (wb == nullptr || viewport == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_partial_recalc: NULL argument");
  }
  // SheetCellRange::sheet_id is std::uint16_t; reject the narrowing path so
  // a caller-supplied sheet > 0xFFFF does not silently address a different
  // sheet (or wrap to 0). Excel's hard cap is far below 0xFFFF anyway.
  if (viewport->sheet > 0xFFFFU) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_partial_recalc: viewport->sheet exceeds 16-bit sheet id range",
                             "sheet=" + std::to_string(viewport->sheet));
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
  // The C ABI callback signature
  //   `bool(*)(uint32_t, double, uint32_t, void*)`
  // is bit-identical to the engine's `IterativeProgressCb` typedef, so
  // a direct assignment is well-defined under both C and C++ rules.
  formulon::eval::IterativeProgressCb engine_cb = cb;
  wb->workbook().set_iterative_progress(engine_cb, user_data);
  return 0;
}
