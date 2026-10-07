//
// C ABI - workbook lifecycle, save/load, sheet management, defined names,
// passthrough parts, structural row/column insertion + deletion. Tables live
// in `parts/tables.cpp`; recalc and calculation settings in
// `parts/recalc.cpp`.

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
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::check_formula_parses;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

fm_status_t set_api_error(const formulon::Error& error, const char* api_name) {
  formulon::Error named = error;
  named.message = std::string(api_name) + ": " + error.message;
  return set_last_error(named);
}

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
                                       size_t* out_len, fm_save_diagnostics_t* out_diagnostics, const char* api_name) {
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
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             (std::string(api_name) + ": NULL argument").c_str());
  }

  std::vector<std::uint8_t> bytes;
  fm_save_diagnostics_t diagnostics{};
  switch (format) {
    case FM_WORKBOOK_FORMAT_XLSX: {
      auto result = formulon::io::write_ooxml_with_result(wb->workbook());
      if (!result) {
        return set_api_error(result.error(), api_name);
      }
      formulon::io::OoxmlWriteResult write_result = std::move(result.value());
      bytes = std::move(write_result.bytes);
      diagnostics = to_c_save_diagnostics(write_result.diagnostics);
      break;
    }
    case FM_WORKBOOK_FORMAT_XLSB: {
      auto result = formulon::io::xlsb::write_xlsb_with_result(wb->workbook());
      if (!result) {
        return set_api_error(result.error(), api_name);
      }
      formulon::io::xlsb::XlsbWriteResult write_result = std::move(result.value());
      bytes = std::move(write_result.bytes);
      diagnostics = to_c_save_diagnostics(write_result.diagnostics);
      break;
    }
    case FM_WORKBOOK_FORMAT_UNKNOWN:
    default:
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               (std::string(api_name) + ": unsupported format").c_str(),
                               "format=" + std::to_string(static_cast<int>(format)));
  }

  auto* buffer = new uint8_t[bytes.size()];
  if (!bytes.empty()) {
    std::memcpy(buffer, bytes.data(), bytes.size());
  }
  *out_bytes = buffer;
  *out_len = bytes.size();
  *out_diagnostics = diagnostics;
  return 0;
}

/// Shared body of `fm_workbook_create` / `fm_workbook_create_empty`.
fm_status_t create_handle(fm_workbook_t** out, formulon::Workbook (*factory)(), const char* null_out_message) {
  clear_last_error();
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, null_out_message);
  }
  auto handle = std::unique_ptr<fm_workbook_t>(new fm_workbook_t{});
  handle->wb.emplace(factory());
  *out = handle.release();
  return 0;
}

}  // namespace

// ---------------------------------------------------------------------------
// Construction / lifecycle
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_create(fm_workbook_t** out) {
  return create_handle(out, &formulon::Workbook::create, "fm_workbook_create: out is NULL");
}

extern "C" fm_status_t fm_workbook_create_empty(fm_workbook_t** out) {
  return create_handle(out, &formulon::Workbook::create_empty, "fm_workbook_create_empty: out is NULL");
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

extern "C" fm_status_t fm_workbook_read_diagnostics(const fm_workbook_t* wb, fm_read_diagnostics_t* out_diagnostics) {
  clear_last_error();
  if (out_diagnostics != nullptr) {
    *out_diagnostics = fm_read_diagnostics_t{};
  }
  if (wb == nullptr || out_diagnostics == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_read_diagnostics: NULL argument");
  }
  *out_diagnostics = wb->read_diagnostics;
  return 0;
}

extern "C" fm_status_t fm_workbook_memory_usage(const fm_workbook_t* wb, size_t* out_bytes) {
  clear_last_error();
  if (wb == nullptr || out_bytes == nullptr || !wb->wb.has_value()) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_memory_usage: NULL argument");
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
  return save_with_diagnostics_impl(wb, FM_WORKBOOK_FORMAT_XLSX, out_bytes, out_len, &diagnostics, "fm_workbook_save");
}

extern "C" fm_status_t fm_workbook_save_with_diagnostics(const fm_workbook_t* wb, std::int32_t format,
                                                         uint8_t** out_bytes, size_t* out_len,
                                                         fm_save_diagnostics_t* out_diagnostics) {
  return save_with_diagnostics_impl(wb, format, out_bytes, out_len, out_diagnostics,
                                    "fm_workbook_save_with_diagnostics");
}

extern "C" fm_status_t fm_workbook_save_as(const fm_workbook_t* wb, std::int32_t format, uint8_t** out_bytes,
                                           size_t* out_len) {
  fm_save_diagnostics_t diagnostics{};
  return save_with_diagnostics_impl(wb, format, out_bytes, out_len, &diagnostics, "fm_workbook_save_as");
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
  wb->cell_enumeration_cache = {};
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
  wb->cell_enumeration_cache = {};
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

namespace {

// Shared entry check of the defined-name setters: a non-empty formula must parse.
fm_status_t check_defined_name_formula(const char* fn, const char* name, const char* formula) {
  if (formula[0] == '\0') {
    return 0;
  }
  return check_formula_parses(fn, formula, "name=" + std::string(name));
}

}  // namespace

extern "C" fm_status_t fm_workbook_set_defined_name(fm_workbook_t* wb, const char* name, const char* formula) {
  clear_last_error();
  if (wb == nullptr || name == nullptr || formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_set_defined_name: NULL argument");
  }
  if (auto rc = check_defined_name_formula("fm_workbook_set_defined_name", name, formula); rc != 0) {
    return rc;
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
  if (auto rc = check_defined_name_formula("fm_workbook_set_defined_name_scoped", name, formula); rc != 0) {
    return rc;
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
// Defined names / passthrough parts (read-side iteration)
// ---------------------------------------------------------------------------
//
// `fm_workbook_defined_name_count`, `fm_workbook_table_count`, and
// `fm_workbook_passthrough_count` are now emitted by the binding
// codegen (see `src/c_api/generated/workbook_counts.cpp`).

extern "C" fm_status_t fm_workbook_defined_name_at(const fm_workbook_t* wb, size_t idx, const char** out_name,
                                                   const char** out_formula, int32_t* out_local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_defined_name_at: NULL argument");
  }
  const auto& names = wb->workbook().defined_names();
  if (auto rc = check_index(idx, names.size(), "fm_workbook_defined_name_at", "idx"); rc != 0) {
    return rc;
  }
  *out_name = names[idx].name.c_str();
  *out_formula = names[idx].formula.c_str();
  if (out_local_sheet_id != nullptr) {
    *out_local_sheet_id = names[idx].local_sheet_id;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_passthrough_at(const fm_workbook_t* wb, size_t idx, const char** out_path) {
  clear_last_error();
  if (wb == nullptr || out_path == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_passthrough_at: NULL argument");
  }
  const auto& parts = wb->workbook().passthrough_parts();
  if (auto rc = check_index(idx, parts.size(), "fm_workbook_passthrough_at", "idx"); rc != 0) {
    return rc;
  }
  *out_path = parts[idx].path.c_str();
  return 0;
}
