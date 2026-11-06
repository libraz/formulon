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
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_keeps_visible_sheet;
using formulon::c_api::parts::check_profile_id;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::copy_owned_bytes;
using formulon::c_api::parts::display_formula;
using formulon::c_api::parts::finish_status;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::may_respell_formula;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::saturate_index;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

// Projects the io layer's `WriteDiagnostics` onto the C ABI struct. Both
// carry the same five `uint32_t` counters under the same names; the copy is
// explicit so a field added on one side fails to compile rather than
// silently reading zero on the other.
fm_save_diagnostics_t to_c_save_diagnostics(const formulon::io::WriteDiagnostics& src) {
  fm_save_diagnostics_t out{};
  out.downgraded_formula_count = src.downgraded_formula_count;
  out.deferred_feature_count = src.deferred_feature_count;
  out.dropped_part_count = src.dropped_part_count;
  out.dropped_relationship_count = src.dropped_relationship_count;
  out.renumbered_part_count = src.renumbered_part_count;
  return out;
}

fm_status_t save_with_diagnostics_impl(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes,
                                       size_t* out_len, fm_save_diagnostics_t* out_diagnostics) {
  clear_last_error();
  if (out_bytes != nullptr) {
    *out_bytes = nullptr;
  }
  if (out_len != nullptr) {
    *out_len = 0;
  }
  if (out_diagnostics != nullptr) {
    *out_diagnostics = fm_save_diagnostics_t{};
  }
  if (wb == nullptr || !wb->wb.has_value() || out_bytes == nullptr || out_len == nullptr ||
      out_diagnostics == nullptr) {
    return null_argument();
  }

  // Pivots left stale by a cache edit are projected into their cells first. The
  // handle is never a const object, so the cast is well-defined.
  if (format == FM_WORKBOOK_FORMAT_XLSX || format == FM_WORKBOOK_FORMAT_XLSB) {
    formulon::c_api::parts::project_pending_pivots(const_cast<fm_workbook_t*>(wb)->workbook());
  }
  std::vector<std::uint8_t> bytes;
  fm_save_diagnostics_t diagnostics{};
  switch (format) {
    case FM_WORKBOOK_FORMAT_XLSX: {
      auto result = formulon::io::write_ooxml_with_result(wb->workbook());
      if (!result) {
        return set_last_error(result.error());
      }
      formulon::io::OoxmlWriteResult write_result = std::move(result.value());
      bytes = std::move(write_result.bytes);
      diagnostics = to_c_save_diagnostics(write_result.diagnostics);
      break;
    }
    case FM_WORKBOOK_FORMAT_XLSB: {
      auto result = formulon::io::xlsb::write_xlsb_with_result(wb->workbook());
      if (!result) {
        return set_last_error(result.error());
      }
      formulon::io::xlsb::XlsbWriteResult write_result = std::move(result.value());
      bytes = std::move(write_result.bytes);
      diagnostics = to_c_save_diagnostics(write_result.diagnostics);
      break;
    }
    case FM_WORKBOOK_FORMAT_UNKNOWN:
    default:
      return invalid_argument("unsupported format", "format=" + std::to_string(static_cast<int>(format)));
  }

  copy_owned_bytes(bytes, out_bytes, out_len);
  *out_diagnostics = diagnostics;
  return 0;
}

/// Shared body of `fm_workbook_create_default` / `fm_workbook_create_empty`.
fm_status_t create_handle(fm_workbook_t** out, formulon::Workbook (*factory)()) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  *out = nullptr;
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(factory());
  *out = handle.release();
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
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(formulon::Workbook::create_for_profile(profile));
  *out = handle.release();
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

extern "C" fm_status_t fm_workbook_get_diagnostics(const fm_workbook_t* wb, fm_read_diagnostics_t* out_diagnostics) {
  clear_last_error();
  if (out_diagnostics != nullptr) {
    *out_diagnostics = fm_read_diagnostics_t{};
  }
  if (wb == nullptr || out_diagnostics == nullptr) {
    return null_argument();
  }
  *out_diagnostics = wb->read_diagnostics;
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
  // Delegates so the whole save family shares one failure-path contract:
  // every out-param is zeroed before validation and stays zeroed on every
  // non-`kOk` return. `Workbook::save()` is itself `save_as(Ooxml)`, so the
  // produced bytes are unchanged.
  fm_save_diagnostics_t diagnostics{};
  return save_with_diagnostics_impl(wb, FM_WORKBOOK_FORMAT_XLSX, out_bytes, out_len, &diagnostics);
}

extern "C" fm_status_t fm_workbook_save_with_diagnostics(const fm_workbook_t* wb, std::int32_t format,
                                                         uint8_t** out_bytes, size_t* out_len,
                                                         fm_save_diagnostics_t* out_diagnostics) {
  return save_with_diagnostics_impl(wb, format, out_bytes, out_len, out_diagnostics);
}

extern "C" fm_status_t fm_workbook_save_as(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes,
                                           size_t* out_len) {
  fm_save_diagnostics_t diagnostics{};
  return save_with_diagnostics_impl(wb, format, out_bytes, out_len, &diagnostics);
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

extern "C" fm_status_t fm_workbook_get_sheet_index(const fm_workbook_t* wb, const char* name, size_t* out_index,
                                                   int32_t* out_present) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr || out_present == nullptr) {
    return null_argument();
  }
  *out_index = 0;
  *out_present = 0;
  if (name == nullptr) {
    return null_argument("name");
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  const std::size_t found = wb->workbook().sheet_index_by_name(name);
  if (found != static_cast<std::size_t>(-1)) {
    *out_index = found;
    *out_present = 1;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_add_sheet(fm_workbook_t* wb, const char* utf8_name, size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({utf8_name}); rc != 0) {
    return rc;
  }
  auto r = wb->workbook().add_sheet_entered(std::string(utf8_name));
  if (!r) {
    return set_last_error(r.error());
  }
  *out_index = wb->workbook().sheet_count() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_move_sheet(fm_workbook_t* wb, size_t from_index, size_t to_index) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  auto r = wb->workbook().move_sheet(saturate_index(from_index), saturate_index(to_index));
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
    return invalid_argument("sheet is hidden", "index=" + std::to_string(index));
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

// Shared entry check of the defined-name setters: the name must be UTF-8, a
// non-empty formula must parse, and is stored without its leading '='. An
// empty one removes the name.
fm_status_t check_defined_name_formula(const char* name, const char* formula, std::string* out) {
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  return check_stored_formula(formula, "name=" + std::string(name), out);
}

}  // namespace

extern "C" fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return null_argument();
  }
  std::string stored;
  if (auto rc = check_defined_name_formula(name, formula, &stored); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().set_defined_name(std::string(name), std::move(stored)));
}

extern "C" fm_status_t fm_workbook_set_defined_name_scoped(fm_workbook_t* wb, const char* name, const char* formula,
                                                           int32_t local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return null_argument();
  }
  std::string stored;
  if (auto rc = check_defined_name_formula(name, formula, &stored); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().set_defined_name_scoped(std::string(name), std::move(stored), local_sheet_id));
}

extern "C" fm_status_t fm_workbook_set_defined_name_attrs(fm_workbook_t* wb, const char* name, int32_t local_sheet_id,
                                                          const char* comment, int32_t hidden) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name, comment}); rc != 0) {
    return rc;
  }
  // Both setters fail only on a missing name, so the first one reports it
  // before anything is changed.
  if (comment != nullptr) {
    if (auto result = wb->workbook().set_defined_name_comment(name, local_sheet_id, comment); !result) {
      return finish_status(std::move(result));
    }
  }
  if (hidden >= 0) {
    return finish_status(wb->workbook().set_defined_name_hidden(name, local_sheet_id, hidden != 0));
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_insert_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_rows(sheet_index, row, count));
}

extern "C" fm_status_t fm_sheet_delete_rows(fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().delete_rows(sheet_index, row, count));
}

extern "C" fm_status_t fm_sheet_insert_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().insert_cols(sheet_index, col, count));
}

extern "C" fm_status_t fm_sheet_delete_columns(fm_workbook_t* wb, size_t sheet_index, uint32_t col, uint32_t count) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return finish_status(wb->workbook().delete_cols(sheet_index, col, count));
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

extern "C" fm_status_t fm_workbook_get_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                       const char** out_formula, int32_t* out_local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_formula == nullptr) {
    return null_argument();
  }
  const auto& names = wb->workbook().defined_names();
  if (auto rc = check_index(idx, names.size(), "index"); rc != 0) {
    return rc;
  }
  *out_name = names[idx].name.c_str();
  // Read scratch holds a respelled formula; any other one is the model's.
  *out_formula = may_respell_formula(names[idx].formula)
                     ? display_formula(wb->workbook(), names[idx].formula, fresh_read_scratch(wb), {})
                     : names[idx].formula.c_str();
  if (out_local_sheet_id != nullptr) {
    *out_local_sheet_id = names[idx].local_sheet_id;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_get_defined_name_attrs_at(const fm_workbook_t* wb, size_t idx,
                                                             const char** out_comment, int32_t* out_hidden) {
  clear_last_error();
  if (wb == nullptr || out_comment == nullptr || out_hidden == nullptr) {
    return null_argument();
  }
  const auto& names = wb->workbook().defined_names();
  if (auto rc = check_index(idx, names.size(), "index"); rc != 0) {
    return rc;
  }
  *out_comment = names[idx].comment.c_str();
  *out_hidden = names[idx].hidden ? 1 : 0;
  return 0;
}

namespace {

// Position of the defined name `name` (ASCII case-insensitive) in scope `local_sheet_id`, or the list size.
std::size_t find_defined_name_index(const fm_workbook_t* wb, const char* name, int32_t local_sheet_id) {
  const auto& names = wb->workbook().defined_names();
  for (std::size_t i = 0; i < names.size(); ++i) {
    if (names[i].local_sheet_id == local_sheet_id && formulon::strings::case_insensitive_eq(names[i].name, name)) {
      return i;
    }
  }
  return names.size();
}

}  // namespace

extern "C" fm_status_t fm_workbook_get_defined_name(const fm_workbook_t* wb, const char* name, int32_t local_sheet_id,
                                                    int32_t* out_present, const char** out_name,
                                                    const char** out_formula, const char** out_comment,
                                                    int32_t* out_hidden) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || out_present == nullptr || out_name == nullptr || out_formula == nullptr ||
      out_comment == nullptr || out_hidden == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  *out_present = 0;
  *out_name = nullptr;
  *out_formula = nullptr;
  *out_comment = nullptr;
  *out_hidden = 0;
  const std::size_t idx = find_defined_name_index(wb, name, local_sheet_id);
  if (idx == wb->workbook().defined_names().size()) {
    return 0;
  }
  if (auto rc = fm_workbook_get_defined_name_at(wb, idx, out_name, out_formula, nullptr); rc != 0) {
    return rc;
  }
  *out_present = 1;
  return fm_workbook_get_defined_name_attrs_at(wb, idx, out_comment, out_hidden);
}

extern "C" fm_status_t fm_workbook_remove_defined_name(fm_workbook_t* wb, const char* name, int32_t local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || name == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({name}); rc != 0) {
    return rc;
  }
  if (find_defined_name_index(wb, name, local_sheet_id) == wb->workbook().defined_names().size()) {
    return not_found("no such name in that scope",
                     "name=" + std::string(name) + " local_sheet_id=" + std::to_string(local_sheet_id));
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

extern "C" fm_status_t fm_workbook_get_passthrough(const fm_workbook_t* wb, const char* path, int32_t* out_present) {
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
