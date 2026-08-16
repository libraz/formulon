//
// C ABI - workbook lifecycle, save/load, sheet management, recalc /
// iterative / partial-recalc / calc-mode / profile, defined names,
// tables, passthrough parts, structural row/column insertion + deletion.

#include "workbook.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cstring>
#include <memory>
#include <new>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/date_time.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "io/a1_ref.h"
#include "io/format_detect.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/package_diagnostics.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/writer.h"
#include "io/xml_escape.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;

namespace {

// Engine-side shim for the C ABI iterative-solver progress callback. The
// engine passes the registering handle as `user_data`; the caller's own
// opaque pointer is stored beside the callback on that handle. A cleared
// callback is never installed, so the null check here only guards a racing
// clear from another thread and defaults to continuing the solve.
bool iterative_progress_adapter(std::uint32_t iteration, double max_residual, std::uint32_t max_iterations,
                                void* user_data) {
  auto* handle = static_cast<fm_workbook_t*>(user_data);
  if (handle == nullptr || handle->iterative_progress_cb == nullptr) {
    return true;
  }
  return handle->iterative_progress_cb(iteration, max_residual, max_iterations, handle->iterative_progress_user_data) !=
         0;
}

std::string ascii_lower(std::string value) {
  for (char& ch : value) {
    ch = static_cast<char>(std::tolower(static_cast<unsigned char>(ch)));
  }
  return value;
}

std::string xml_attr_escape(std::string_view value) {
  std::string out;
  out.reserve(value.size());
  formulon::io::AppendXmlAttrEscaped(out, value);
  return out;
}

/// Column span of an A1 range such as `"A1:C10"`, or 0 when the text is not
/// the plain single-area form that `xl/tables/tableN.xml` accepts for `ref`.
/// A table whose column count disagrees with its range is a file Excel
/// refuses to open without repair, so the span is checked up front.
std::size_t table_ref_column_count(std::string_view ref) {
  std::uint32_t first_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
  const std::size_t colon = ref.find(':');
  if (colon == std::string_view::npos) {
    return formulon::io::parse_a1_ref(ref, &first_row, &first_col) ? 1U : 0U;
  }
  if (!formulon::io::parse_a1_ref(ref.substr(0, colon), &first_row, &first_col) ||
      !formulon::io::parse_a1_ref(ref.substr(colon + 1), &last_row, &last_col)) {
    return 0U;
  }
  if (last_col < first_col || last_row < first_row) {
    return 0U;
  }
  return static_cast<std::size_t>(last_col - first_col) + 1U;
}

std::string table_style_xml(std::string_view style_name) {
  if (style_name.empty()) {
    return {};
  }
  return "<tableStyleInfo name=\"" + xml_attr_escape(style_name) +
         "\" showFirstColumn=\"0\" showLastColumn=\"0\" showRowStripes=\"1\" showColumnStripes=\"0\"/>";
}

// Keep the raw table-level autoFilter payload, changing only its opening
// element's ref attribute. The reader stores this fragment verbatim because
// filterColumn criteria and extension payloads are not modelled by the
// evaluator.
std::string table_auto_filter_xml(std::string_view raw_xml, std::string_view ref) {
  const std::string escaped_ref = xml_attr_escape(ref);
  if (raw_xml.empty()) {
    return "<autoFilter ref=\"" + escaped_ref + "\"/>";
  }

  const std::size_t name_start = raw_xml.find("<autoFilter");
  if (name_start == std::string_view::npos) {
    return std::string(raw_xml);
  }
  const std::size_t name_end = name_start + std::string_view("<autoFilter").size();
  if (name_end < raw_xml.size() && raw_xml[name_end] != ' ' && raw_xml[name_end] != '\t' && raw_xml[name_end] != '\r' &&
      raw_xml[name_end] != '\n' && raw_xml[name_end] != '/' && raw_xml[name_end] != '>') {
    return std::string(raw_xml);
  }

  bool in_quote = false;
  char quote = '\0';
  std::size_t opening_end = std::string_view::npos;
  for (std::size_t i = name_end; i < raw_xml.size(); ++i) {
    const char ch = raw_xml[i];
    if (in_quote) {
      if (ch == quote) {
        in_quote = false;
      }
    } else if (ch == '\'' || ch == '"') {
      in_quote = true;
      quote = ch;
    } else if (ch == '>') {
      opening_end = i;
      break;
    }
  }
  if (opening_end == std::string_view::npos || in_quote) {
    return std::string(raw_xml);
  }

  // Locate a ref attribute in the opening element. This intentionally edits
  // only the attribute value, retaining its original quoting and whitespace.
  for (std::size_t i = name_end; i + 3U <= opening_end; ++i) {
    if (raw_xml.substr(i, 3U) != "ref") {
      continue;
    }
    const char before = i == name_end ? ' ' : raw_xml[i - 1U];
    const char after = i + 3U < opening_end ? raw_xml[i + 3U] : ' ';
    const bool before_is_space = before == ' ' || before == '\t' || before == '\r' || before == '\n';
    const bool after_is_space = after == ' ' || after == '\t' || after == '\r' || after == '\n' || after == '=';
    if (!before_is_space || !after_is_space) {
      continue;
    }
    std::size_t equal = i + 3U;
    while (equal < opening_end &&
           (raw_xml[equal] == ' ' || raw_xml[equal] == '\t' || raw_xml[equal] == '\r' || raw_xml[equal] == '\n')) {
      ++equal;
    }
    if (equal >= opening_end || raw_xml[equal] != '=') {
      continue;
    }
    ++equal;
    while (equal < opening_end &&
           (raw_xml[equal] == ' ' || raw_xml[equal] == '\t' || raw_xml[equal] == '\r' || raw_xml[equal] == '\n')) {
      ++equal;
    }
    if (equal >= opening_end) {
      continue;
    }
    const char value_quote = raw_xml[equal];
    const bool quoted = value_quote == '\'' || value_quote == '"';
    const std::size_t value_start = quoted ? equal + 1U : equal;
    std::size_t value_end = value_start;
    if (quoted) {
      value_end = raw_xml.find(value_quote, value_start);
      if (value_end == std::string_view::npos || value_end > opening_end) {
        continue;
      }
    } else {
      while (value_end < opening_end && raw_xml[value_end] != ' ' && raw_xml[value_end] != '\t' &&
             raw_xml[value_end] != '\r' && raw_xml[value_end] != '\n' && raw_xml[value_end] != '/') {
        ++value_end;
      }
    }
    std::string out(raw_xml);
    out.replace(value_start, value_end - value_start, escaped_ref);
    return out;
  }

  // No ref attribute: insert it immediately before the closing `>` (or the
  // self-closing slash) without touching any existing attributes.
  std::size_t insert_at = opening_end;
  while (insert_at > name_end && (raw_xml[insert_at - 1U] == ' ' || raw_xml[insert_at - 1U] == '\t' ||
                                  raw_xml[insert_at - 1U] == '\r' || raw_xml[insert_at - 1U] == '\n')) {
    --insert_at;
  }
  if (insert_at > name_end && raw_xml[insert_at - 1U] == '/') {
    --insert_at;
  }
  std::string out(raw_xml);
  out.insert(insert_at, " ref=\"" + escaped_ref + "\"");
  return out;
}

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

}  // namespace

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
                                                   const char** out_formula, int32_t* out_local_sheet_id) {
  clear_last_error();
  if (wb == nullptr || out_name == nullptr || out_formula == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_defined_name_at: NULL argument");
  }
  const auto& names = wb->workbook().defined_names();
  if (idx >= names.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_defined_name_at: idx out of range",
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

extern "C" fm_status_t fm_workbook_table_create(fm_workbook_t* wb, size_t sheet_index, const char* ref,
                                                const char* name, const char* display_name,
                                                const char* const* column_names, size_t column_count,
                                                const char* style_name, int32_t header_row, int32_t totals_row,
                                                size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || ref == nullptr || name == nullptr || display_name == nullptr || column_names == nullptr ||
      out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_create: NULL argument");
  }
  formulon::Workbook& book = wb->workbook();
  if (sheet_index >= book.sheet_count() || ref[0] == '\0' || name[0] == '\0' || display_name[0] == '\0' ||
      column_count == 0U) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_create: invalid sheet, table identity, range, or columns");
  }
  if (table_ref_column_count(ref) != column_count) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_create: column_count does not match the range width",
                             "ref=" + std::string(ref) + " column_count=" + std::to_string(column_count));
  }
  const std::string name_key = ascii_lower(name);
  for (const formulon::io::TableMetadata& existing : book.tables()) {
    if (ascii_lower(existing.name) == name_key || ascii_lower(existing.display_name) == name_key ||
        ascii_lower(existing.name) == ascii_lower(display_name) ||
        ascii_lower(existing.display_name) == ascii_lower(display_name)) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: table name already exists");
    }
  }
  formulon::io::TableMetadata table;
  table.sheet_index = sheet_index;
  table.name = name;
  table.display_name = display_name;
  table.ref = ref;
  table.header_row = header_row != 0;
  table.totals_row = totals_row != 0;
  uint32_t next_id = 1;
  for (const formulon::io::TableMetadata& existing : book.tables()) {
    next_id = std::max(next_id, existing.id + 1U);
  }
  table.id = next_id;
  table.columns.reserve(column_count);
  for (size_t i = 0; i < column_count; ++i) {
    if (column_names[i] == nullptr || column_names[i][0] == '\0') {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_table_create: column name is empty");
    }
    const std::string column_key = ascii_lower(column_names[i]);
    for (const formulon::io::TableColumn& existing : table.columns) {
      if (ascii_lower(existing.name) == column_key) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 "fm_workbook_table_create: column name is duplicated");
      }
    }
    formulon::io::TableColumn column;
    column.id = static_cast<uint32_t>(i + 1U);
    column.name = column_names[i];
    table.columns.push_back(std::move(column));
  }
  table.auto_filter_xml = "<autoFilter ref=\"" + xml_attr_escape(table.ref) + "\"/>";
  table.table_style_info_xml = table_style_xml(style_name != nullptr ? style_name : "");
  auto& tables = book.mutable_tables();
  tables.push_back(std::move(table));
  *out_index = tables.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_table_update(fm_workbook_t* wb, size_t index, const char* ref,
                                                const char* style_name, int32_t header_row, int32_t totals_row) {
  clear_last_error();
  if (wb == nullptr || ref == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_table_update: NULL argument");
  }
  auto& tables = wb->workbook().mutable_tables();
  if (index >= tables.size() || ref[0] == '\0') {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_update: table index or ref is invalid");
  }
  formulon::io::TableMetadata& table = tables[index];
  if (!table.columns.empty() && table_ref_column_count(ref) != table.columns.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_update: ref width does not match the table's column count",
                             "ref=" + std::string(ref) + " column_count=" + std::to_string(table.columns.size()));
  }

  // Compute every replacement before committing any metadata. In
  // particular, a raw style payload survives a NULL style_name, while an
  // empty string explicitly removes it.
  const std::string next_ref = ref;
  const std::string next_auto_filter_xml = table_auto_filter_xml(table.auto_filter_xml, next_ref);
  const std::string next_style_xml = style_name == nullptr ? table.table_style_info_xml : table_style_xml(style_name);
  const bool next_header_row = header_row < 0 ? table.header_row : header_row > 0;
  const bool next_totals_row = totals_row < 0 ? table.totals_row : totals_row > 0;

  table.ref = next_ref;
  table.header_row = next_header_row;
  table.totals_row = next_totals_row;
  table.auto_filter_xml = next_auto_filter_xml;
  table.table_style_info_xml = next_style_xml;
  return 0;
}

extern "C" fm_status_t fm_workbook_table_remove(fm_workbook_t* wb, size_t index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_workbook_table_remove: wb is NULL");
  }
  auto& tables = wb->workbook().mutable_tables();
  if (index >= tables.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_table_remove: table index out of range");
  }
  tables.erase(tables.begin() + static_cast<std::ptrdiff_t>(index));
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
  const bool in_range = now->year >= 1900 && now->year <= 9999 && now->month >= 1 && now->month <= 12 &&
                        now->day >= 1 &&
                        now->day <= static_cast<std::int32_t>(formulon::eval::date_time::days_in_month(
                                        now->year, static_cast<unsigned>(now->month))) &&
                        now->hour >= 0 && now->hour <= 23 && now->minute >= 0 && now->minute <= 59 &&
                        now->second >= 0 && now->second <= 59;
  if (!in_range) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_set_pinned_now: field out of range");
  }
  wb->workbook().set_pinned_now(formulon::eval::date_time::CivilTime{
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
  wb->iterative_progress_cb = cb;
  wb->iterative_progress_user_data = cb == nullptr ? nullptr : user_data;
  // The engine's callback returns `bool`, which no C ABI declaration uses.
  // Hand it the adapter above instead of the caller's pointer so both sides
  // keep their own return type and neither assignment needs a cast.
  const formulon::eval::IterativeProgressCb engine_cb = cb == nullptr ? nullptr : &iterative_progress_adapter;
  wb->workbook().set_iterative_progress(engine_cb, cb == nullptr ? nullptr : wb);
  return 0;
}
