//
// C ABI - workbook lifecycle, save/load, sheet management, defined names,
// passthrough parts, structural row/column insertion + deletion. Tables live
// in `parts/tables.cpp`; recalc and calculation settings in
// `parts/recalc.cpp`.

#include "workbook.h"

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
#include "c_api/parts/pivot_internal.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/function_registry.h"
#include "eval/recalc_engine.h"
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_keeps_visible_sheet;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::display_formula_text;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::precondition_failed;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::saturate_index;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::with_formula_prefix;

namespace {

// An inclusive span `[first, last]` inside a sheet with `limit` lines.
fm_status_t check_span(std::uint32_t first, std::uint32_t last, std::uint32_t limit) {
  if (last < first || last >= limit) {
    return invalid_argument("span is empty or past the sheet bound",
                            "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  return 0;
}

fm_status_t sheet_index_out_of_range(std::size_t position, const char* what) {
  return formulon::c_api::parts::set_binding_error(formulon::FormulonErrorCode::kSheetIndexOutOfRange,
                                                   "sheet position is out of range",
                                                   std::string(what) + "=" + std::to_string(position));
}

// Projects the io layer's `WriteDiagnostics` onto the C ABI struct. Both
// carry the same five `uint32_t` counters under the same names; the copy is
// explicit so a field added on one side fails to compile rather than
// silently reading zero on the other.
fm_save_diagnostics_info_t to_c_save_diagnostics(const formulon::io::WriteDiagnostics& src) {
  fm_save_diagnostics_info_t out = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
  out.downgraded_formula_count = src.downgraded_formula_count;
  out.deferred_feature_count = src.deferred_feature_count;
  out.dropped_part_count = src.dropped_part_count;
  out.dropped_relationship_count = src.dropped_relationship_count;
  out.renumbered_part_count = src.renumbered_part_count;
  return out;
}

// Zeroes the save out-params (a diagnostics record keeps its `struct_size`)
// and validates them and `format`; they stay zeroed on every failure.
fm_status_t begin_save(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes, size_t* out_len,
                       fm_save_diagnostics_info_t* out_diagnostics) {
  clear_last_error();
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0;
  }
  if (out_diagnostics != nullptr) {
    if (auto rc = check_record_size(*out_diagnostics); rc != 0) {
      return rc;
    }
    const std::uint32_t struct_size = out_diagnostics->struct_size;
    *out_diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
    out_diagnostics->struct_size = struct_size;
  }
  if (wb == nullptr || !wb->wb.has_value() || out_bytes == nullptr || out_len == nullptr ||
      out_diagnostics == nullptr) {
    return null_argument();
  }
  if (format != FM_WORKBOOK_FORMAT_XLSX && format != FM_WORKBOOK_FORMAT_XLSB) {
    return invalid_argument("unsupported format", "format=" + std::to_string(static_cast<int>(format)));
  }
  return 0;
}

// The calculation a save adds to `ensure_calculated`: a manual workbook
// with `calcOnSave` computes its pending cells before it is written.
fm_status_t calculate_on_save(const fm_workbook_t* wb) {
  formulon::Workbook& book = const_cast<fm_workbook_t*>(wb)->workbook();
  if (book.calc_mode() != formulon::CalcMode::kManual || !book.calc_options().calc_on_save ||
      book.recalc_engine().dirty().size() == 0U) {
    return 0;
  }
  auto r = book.recalc(formulon::eval::default_registry());
  return r ? 0 : set_last_error(r.error());
}

// Writes `wb` as `format`, which `begin_save` accepted, into the out-params.
fm_status_t write_workbook(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes, size_t* out_len,
                           fm_save_diagnostics_info_t* out_diagnostics) {
  // Pivots left stale by a cache edit are projected into their cells first. The
  // handle is never a const object, so the cast is well-defined.
  formulon::c_api::parts::project_pending_pivots(const_cast<fm_workbook_t*>(wb)->workbook());
  std::vector<std::uint8_t> bytes;
  fm_save_diagnostics_info_t diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
  if (format == FM_WORKBOOK_FORMAT_XLSX) {
    auto result = formulon::io::write_ooxml_with_result(wb->workbook());
    if (!result) {
      return set_last_error(result.error());
    }
    formulon::io::OoxmlWriteResult write_result = std::move(result.value());
    bytes = std::move(write_result.bytes);
    diagnostics = to_c_save_diagnostics(write_result.diagnostics);
  } else {
    auto result = formulon::io::xlsb::write_xlsb_with_result(wb->workbook());
    if (!result) {
      return set_last_error(result.error());
    }
    formulon::io::xlsb::XlsbWriteResult write_result = std::move(result.value());
    bytes = std::move(write_result.bytes);
    diagnostics = to_c_save_diagnostics(write_result.diagnostics);
  }

  copy_owned_bytes(bytes, out_bytes, out_len);
  diagnostics.struct_size = out_diagnostics->struct_size;
  *out_diagnostics = diagnostics;
  return 0;
}

/// Wraps `book` in a new handle and hands it to `*out`.
void publish_new_handle(fm_workbook_t** out, formulon::Workbook book) {
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(std::move(book));
  *out = handle.release();
}

/// Shared body of `fm_workbook_create_default` / `fm_workbook_create_empty`.
fm_status_t create_handle(fm_workbook_t** out, formulon::Workbook (*factory)()) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  *out = nullptr;
  publish_new_handle(out, factory());
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// Construction / lifecycle
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_create_default(fm_workbook_t** out) {
  return create_handle(out, &formulon::Workbook::create);
}

extern "C" fm_status_t fm_workbook_create_for_profile(const char* profile_id, fm_workbook_t** out) {
  clear_last_error();
  if (profile_id == nullptr || out == nullptr) {
    return null_argument();
  }
  *out = nullptr;
  formulon::ExcelProfile profile;
  if (const fm_status_t st = check_profile_id(profile_id, &profile)) {
    return st;
  }
  publish_new_handle(out, formulon::Workbook::create_for_profile(profile));
  return 0;
}

extern "C" fm_status_t fm_workbook_create_empty(fm_workbook_t** out) {
  return create_handle(out, &formulon::Workbook::create_empty);
}

extern "C" fm_status_t fm_workbook_load(const uint8_t* bytes, size_t len, fm_workbook_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  *out = nullptr;
  if (bytes == nullptr) {
    return null_argument("bytes");
  }
  if (len == 0) {
    return invalid_argument("input is empty", "len=0");
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
  if (formulon::io::detect_workbook_format(span) == formulon::WorkbookFormat::Xlsb) {
    auto result = formulon::io::xlsb::read_xlsb(span);
    if (!result) {
      return set_last_error(result.error());
    }
    formulon::io::xlsb::XlsbReadResult read_result = std::move(result.value());
    handle->read_diagnostics.undecoded_formula_count = read_result.undecoded_formula_count;
    handle->read_diagnostics.undecoded_defined_name_count = read_result.undecoded_defined_name_count;
    handle->read_diagnostics.undecoded_part_count = read_result.dropped_part_count;
    handle->wb.emplace(std::move(read_result.workbook));
  } else {
    auto result = formulon::io::read_ooxml(span);
    if (!result) {
      return set_last_error(result.error());
    }
    // The workbook now owns the text-storage deque that backs every
    // Text-cell `string_view` as well as the passthrough payload, so
    // moving it out takes everything the handle needs; the read result's
    // remaining audit counter is discarded.
    handle->read_diagnostics.skipped_feature_count = result.value().diagnostics.skipped_feature_count;
    handle->read_diagnostics.unknown_content_type_count = result.value().diagnostics.unknown_content_type_count;
    handle->wb.emplace(std::move(result.value().workbook));
  }
  *out = handle.release();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_diagnostics(const fm_workbook_t* wb,
                                                   fm_read_diagnostics_info_t* out_diagnostics) {
  clear_last_error();
  if (out_diagnostics == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out_diagnostics); rc != 0) {
    return rc;
  }
  const std::uint32_t struct_size = out_diagnostics->struct_size;
  *out_diagnostics = FM_STRUCT_INIT(fm_read_diagnostics_info_t);
  out_diagnostics->struct_size = struct_size;
  if (wb == nullptr) {
    return null_argument();
  }
  *out_diagnostics = wb->read_diagnostics;
  out_diagnostics->struct_size = struct_size;
  return 0;
}

extern "C" fm_status_t fm_workbook_get_memory_usage(const fm_workbook_t* wb, size_t* out_bytes) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || !wb->wb.has_value()) {
    return null_argument();
  }
  *out_bytes = wb->wb->approximate_memory_bytes();
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
  // The whole save family shares one failure-path contract: every out-param
  // is zeroed before validation and stays zeroed on every non-`kOk` return.
  // `Workbook::save()` is itself `save_as(Ooxml)`, so the produced bytes are
  // unchanged.
  fm_save_diagnostics_info_t diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
  if (auto rc = begin_save(wb, FM_WORKBOOK_FORMAT_XLSX, out_bytes, out_len, &diagnostics); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  if (auto rc = calculate_on_save(wb); rc != 0) {
    return rc;
  }
  return write_workbook(wb, FM_WORKBOOK_FORMAT_XLSX, out_bytes, out_len, &diagnostics);
}

extern "C" fm_status_t fm_workbook_save_with_diagnostics(const fm_workbook_t* wb, std::int32_t format,
                                                         uint8_t** out_bytes, size_t* out_len,
                                                         fm_save_diagnostics_info_t* out_diagnostics) {
  if (auto rc = begin_save(wb, format, out_bytes, out_len, out_diagnostics); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  if (auto rc = calculate_on_save(wb); rc != 0) {
    return rc;
  }
  return write_workbook(wb, format, out_bytes, out_len, out_diagnostics);
}

extern "C" fm_status_t fm_workbook_save_as(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes,
                                           size_t* out_len) {
  fm_save_diagnostics_info_t diagnostics = FM_STRUCT_INIT(fm_save_diagnostics_info_t);
  if (auto rc = begin_save(wb, format, out_bytes, out_len, &diagnostics); rc != 0) {
    return rc;
  }
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  if (auto rc = calculate_on_save(wb); rc != 0) {
    return rc;
  }
  return write_workbook(wb, format, out_bytes, out_len, &diagnostics);
}

extern "C" void fm_buffer_destroy(uint8_t* bytes) {
  delete[] bytes;
}

// ---------------------------------------------------------------------------
// Sheets
// ---------------------------------------------------------------------------
//
// `fm_workbook_get_sheet_count` is now emitted by the binding codegen (see
// `src/c_api/generated/workbook_counts.cpp`).

extern "C" fm_status_t fm_workbook_get_sheet_name(const fm_workbook_t* wb, size_t index, const char** out_utf8) {
  clear_last_error();
  if (wb == nullptr || out_utf8 == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, index); rc != 0) {
    return rc;
  }
  // `Sheet::name()` returns `const std::string&`, so `c_str()` is
  // NUL-terminated and stable until the sheet is mutated or destroyed.
  *out_utf8 = wb->workbook().sheet(index).name().c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_sheet_index(const fm_workbook_t* wb, const char* name, int32_t* out_present,
                                                   size_t* out_sheet_index) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out_sheet_index == nullptr) {
    return null_argument();
  }
  *out_present = 0;
  *out_sheet_index = 0;
  if (name == nullptr) {
    return null_argument("name");
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  const std::size_t found = wb->workbook().sheet_index_by_name(name);
  if (found != static_cast<std::size_t>(-1)) {
    *out_present = 1;
    *out_sheet_index = found;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* utf8_name, size_t insert_before,
                                             size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (utf8_name != nullptr) {
    if (auto rc = check_utf8({utf8_name}); rc != 0) {
      return rc;
    }
  }
  formulon::Workbook& book = wb->workbook();
  const std::size_t count = book.sheet_count();
  if (insert_before > count) {
    return sheet_index_out_of_range(insert_before, "insert_before");
  }
  auto r = book.add_sheet_entered(utf8_name != nullptr ? std::string(utf8_name) : book.next_default_sheet_name());
  if (!r) {
    return set_last_error(r.error());
  }
  // The new sheet is last; the engine's move interprets its target in the list without it.
  if (insert_before < count) {
    book.move_sheet(static_cast<std::uint32_t>(count), static_cast<std::uint32_t>(insert_before));
    wb->clear_cell_enumeration_cache();
  }
  *out_index = std::min(insert_before, count);
  return 0;
}

extern "C" fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, size_t from_index, size_t insert_before) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  const std::size_t count = wb->workbook().sheet_count();
  if (from_index >= count) {
    return sheet_index_out_of_range(from_index, "from_index");
  }
  if (insert_before > count) {
    return sheet_index_out_of_range(insert_before, "insert_before");
  }
  // The engine's target index is a position in the list after the sheet is lifted out.
  const std::size_t to_index = insert_before > from_index ? insert_before - 1U : insert_before;
  auto r = wb->workbook().move_sheet(static_cast<std::uint32_t>(from_index), static_cast<std::uint32_t>(to_index));
  if (!r) {
    return set_last_error(r.error());
  }
  // The enumeration cache is keyed by sheet index. Moving sheets can put a
  // different Sheet at a cached index with the same cell revision.
  wb->clear_cell_enumeration_cache();
  return 0;
}

extern "C" fm_status_t fm_workbook_copy_sheet(fm_workbook_t* wb, size_t source_index, size_t insert_before,
                                              const char* new_name, size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return null_argument();
  }
  std::optional<std::string> name;
  if (new_name != nullptr) {
    if (auto rc = check_utf8({new_name}); rc != 0) {
      return rc;
    }
    name = std::string(new_name);
  }
  auto r = wb->workbook().copy_sheet(source_index, insert_before, std::move(name));
  if (!r) {
    return set_last_error(r.error());
  }
  // The enumeration cache is keyed by sheet index; inserting a sheet shifts
  // every later index.
  wb->clear_cell_enumeration_cache();
  *out_index = r.value();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_active_sheet(const fm_workbook_t* wb, size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return null_argument();
  }
  *out_index = wb->workbook().active_sheet();
  return 0;
}

extern "C" fm_status_t fm_workbook_set_active_sheet(fm_workbook_t* wb, size_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_sheet_index(wb, index); rc != 0) {
    return rc;
  }
  if (wb->workbook().sheet(index).view().visibility() != formulon::SheetVisibility::kVisible) {
    return precondition_failed("sheet is hidden", "sheet_index=" + std::to_string(index));
  }
  wb->workbook().set_active_sheet(static_cast<std::uint32_t>(index));
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_sheet(fm_workbook_t* wb, size_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  // Range and last-sheet refusals stay with `Workbook::remove_sheet`.
  if (index < wb->workbook().sheet_count() && wb->workbook().sheet_count() > 1U) {
    if (auto rc = check_keeps_visible_sheet(wb->workbook(), static_cast<std::uint32_t>(index)); rc != 0) {
      return rc;
    }
  }
  auto r = wb->workbook().remove_sheet(saturate_index(index));
  if (!r) {
    return set_last_error(r.error());
  }
  // Removing a sheet shifts later sheet indices, so invalidate the
  // index-keyed coordinate cache even though the remaining Sheet objects
  // themselves were not mutated.
  wb->clear_cell_enumeration_cache();
  return 0;
}

extern "C" fm_status_t fm_workbook_rename_sheet(fm_workbook_t* wb, size_t index, const char* new_name) {
  clear_last_error();
  if (wb == nullptr || new_name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({new_name}); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().rename_sheet(saturate_index(index), std::string(new_name)));
}

namespace {

// Context text for a defined-name scope: `workbook` or the sheet index.
std::string scope_context(std::int32_t local_sheet_id) {
  return local_sheet_id == FM_DEFINED_NAME_SCOPE_WORKBOOK ? std::string("scope=workbook")
                                                          : "scope=" + std::to_string(local_sheet_id);
}

// Rejects a defined-name scope that is neither workbook scope nor a sheet of the workbook.
fm_status_t check_local_sheet_id(const fm_workbook_t* wb, std::int32_t local_sheet_id) {
  const std::size_t count = wb->workbook().sheet_count();
  if (local_sheet_id == FM_DEFINED_NAME_SCOPE_WORKBOOK ||
      (local_sheet_id >= 0 && static_cast<std::size_t>(local_sheet_id) < count)) {
    return 0;
  }
  return formulon::c_api::parts::set_binding_error(
      formulon::FormulonErrorCode::kSheetIndexOutOfRange, "defined-name scope is neither the workbook nor a sheet",
      scope_context(local_sheet_id) + " sheet_count=" + std::to_string(count));
}

}  // namespace

extern "C" fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const fm_defined_name_t* input) {
  clear_last_error();
  if (wb == nullptr || input == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*input); rc != 0) {
    return rc;
  }
  if (input->name == nullptr || input->formula == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({input->name, input->comment}); rc != 0) {
    return rc;
  }
  if (auto rc = check_local_sheet_id(wb, input->local_sheet_id); rc != 0) {
    return rc;
  }
  std::string stored;
  if (auto rc = check_stored_formula(input->formula, "field=formula", &stored); rc != 0) {
    return rc;
  }
  // The model reads an empty formula as a removal; a name always has one.
  if (stored.empty()) {
    return invalid_argument("formula is empty", "name=" + std::string(input->name));
  }
  auto& book = wb->workbook();
  const std::int32_t scope = input->local_sheet_id;
  if (auto result = book.set_defined_name_scoped(std::string(input->name), std::move(stored), scope); !result) {
    return finish_status(std::move(result));
  }
  // The record replaces the whole entry, so absent attributes reset to their defaults.
  if (auto result = book.set_defined_name_comment(input->name, scope, text_or_empty(input->comment)); !result) {
    return finish_status(std::move(result));
  }
  return finish_status(book.set_defined_name_hidden(input->name, scope, input->hidden != 0));
}

extern "C" fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_span(first, last, formulon::Sheet::kMaxRows); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_rows(sheet_index, first, last - first + 1U));
}

extern "C" fm_status_t fm_sheet_delete_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_span(first, last, formulon::Sheet::kMaxRows); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().delete_rows(sheet_index, first, last - first + 1U));
}

extern "C" fm_status_t fm_sheet_insert_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_span(first, last, formulon::Sheet::kMaxCols); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_cols(sheet_index, first, last - first + 1U));
}

extern "C" fm_status_t fm_sheet_delete_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t first, uint32_t last) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_span(first, last, formulon::Sheet::kMaxCols); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().delete_cols(sheet_index, first, last - first + 1U));
}

// ---------------------------------------------------------------------------
// Defined names / passthrough parts (read-side iteration)
// ---------------------------------------------------------------------------
//
// `fm_workbook_get_table_count` is emitted by the binding codegen (see
// `src/c_api/generated/workbook_counts.cpp`).

extern "C" fm_status_t fm_workbook_get_defined_name_count(const fm_workbook_t* wb, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().defined_names().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_defined_name_at(const fm_workbook_t* wb, size_t idx, fm_defined_name_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const auto& names = wb->workbook().defined_names();
  if (auto rc = check_index(idx, names.size(), "index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto& entry = names[idx];
  out->name = entry.name.c_str();
  out->formula = publish_scratch(wb, with_formula_prefix(display_formula_text(wb->workbook(), entry.formula, {})));
  out->local_sheet_id = entry.local_sheet_id;
  out->comment = entry.comment.empty() ? nullptr : entry.comment.c_str();
  out->hidden = entry.hidden ? 1 : 0;
  return 0;
}

namespace {

// Position of the defined name `name` resolves to (`sheet_names::resolve_by`) in scope `local_sheet_id`, or the
// list size.
std::size_t find_defined_name_index(const fm_workbook_t* wb, const char* name, int32_t local_sheet_id) {
  const auto& names = wb->workbook().defined_names();
  const std::size_t found =
      formulon::sheet_names::resolve_by(names.size(), [&](std::size_t i, formulon::sheet_names::Fold fold) {
        return names[i].local_sheet_id == local_sheet_id && formulon::sheet_names::same(names[i].name, name, fold);
      });
  return found == formulon::sheet_names::kNotFound ? names.size() : found;
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_defined_name(const fm_workbook_t* wb, const char* name, int32_t local_sheet_id,
                                                    int32_t* out_present, fm_defined_name_t* out) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  *out_present = 0;
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_defined_name_t{};
  out->struct_size = struct_size;
  if (auto rc = check_local_sheet_id(wb, local_sheet_id); rc != 0) {
    return rc;
  }
  const std::size_t idx = find_defined_name_index(wb, name, local_sheet_id);
  if (idx == wb->workbook().defined_names().size()) {
    return 0;
  }
  if (auto rc = fm_workbook_get_defined_name_at(wb, idx, out); rc != 0) {
    return rc;
  }
  *out_present = 1;
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_defined_name(fm_workbook_t* wb, const char* name, int32_t local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  if (auto rc = check_local_sheet_id(wb, local_sheet_id); rc != 0) {
    return rc;
  }
  if (find_defined_name_index(wb, name, local_sheet_id) == wb->workbook().defined_names().size()) {
    return not_found("no such name in that scope", "name=" + std::string(name) + " " + scope_context(local_sheet_id));
  }
  return finish_status(wb->workbook().set_defined_name_scoped(std::string(name), std::string(), local_sheet_id));
}

extern "C" fm_status_t fm_workbook_get_passthrough_count(const fm_workbook_t* wb, size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().passthrough_parts().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path) {
  clear_last_error();
  if (wb == nullptr || out_path == nullptr) {
    return null_argument();
  }
  const auto& parts = wb->workbook().passthrough_parts();
  if (auto rc = check_index(idx, parts.size(), "index"); rc != 0) {
    return rc;
  }
  *out_path = parts[idx].path.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_has_passthrough(const fm_workbook_t* wb, const char* path, int32_t* out_present) {
  clear_last_error();
  if (wb == nullptr || path == nullptr || out_present == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({path}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  for (const auto& part : wb->workbook().passthrough_parts()) {
    if (part.path == path) {
      *out_present = 1;
      break;
    }
  }
  return 0;
}
