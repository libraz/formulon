//
// C ABI - PivotCache mutation surface (caches, fields, shared items,
// records). The PivotTable mutation surface lives in `pivot_table.cpp`;
// the two share helpers from `pivot_internal.h`.

#include "pivot/pivot_cache.h"

#include <cstddef>
#include <cstdint>
#include <limits>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/builtin_names.h"
#include "eval/eval_profile_scope.h"
#include "eval/pivot_formula.h"
#include "parser/parse_status.h"
#include "pivot/calculated.h"
#include "pivot/date_serial.h"
#include "pivot/pivot_formula_host.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::cache_tables;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_patch_op;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula_text;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::formula_parse_failed;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::grow_record_cells;
using formulon::c_api::parts::intern_cache_text;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::invalid_error_code;
using formulon::c_api::parts::invalidate_pivot_results_for_cache;
using formulon::c_api::parts::is_valid_error_code;
using formulon::c_api::parts::kMaxPivotNameUnits;
using formulon::c_api::parts::mark_cache_pivots_pending;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::patch_op_ok;
using formulon::c_api::parts::patch_op_ok_without_clear;
using formulon::c_api::parts::publish_scratch;
using formulon::c_api::parts::resolve_cache;
using formulon::c_api::parts::resolve_cache_mut;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::value_to_fm;
using formulon::c_api::parts::with_formula_prefix;

namespace {

// Read-only lookup and field-bounds check shared by the cache-field getters.
// Returns NULL after writing the binding error and reporting its status
// through `out_status`.
const formulon::pivot::PivotCacheField* lookup_cache_field(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                           std::size_t field_index, fm_status_t* out_status) {
  const formulon::pivot::PivotCache* cache = nullptr;
  if (*out_status = resolve_cache(wb, cache_id, &cache); *out_status != 0) {
    return nullptr;
  }
  if (*out_status = check_index(field_index, cache->fields().size(), "field_index"); *out_status != 0) {
    return nullptr;
  }
  return &cache->fields()[field_index];
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotCache lifecycle / enumeration
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_pivot_cache_count(const fm_workbook_t* wb, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = wb->workbook().pivot_caches().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_id_at(const fm_workbook_t* wb, std::size_t idx,
                                                         std::uint32_t* out_cache_id) {
  clear_last_error();
  if (wb == nullptr || out_cache_id == nullptr) {
    return null_argument();
  }
  const auto& caches = wb->workbook().pivot_caches();
  if (auto rc = check_index(idx, caches.size(), "index"); rc != 0) {
    return rc;
  }
  *out_cache_id = caches[idx]->cache_id();
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache(fm_workbook_t* wb, std::uint32_t requested_id,
                                                   std::uint32_t* out_cache_id) {
  clear_last_error();
  if (wb == nullptr || out_cache_id == nullptr) {
    return null_argument();
  }
  std::uint32_t assigned = requested_id;
  if (requested_id == 0U) {
    std::uint32_t max_id = 0U;
    bool any = false;
    for (const auto& c : wb->workbook().pivot_caches()) {
      if (c == nullptr) {
        continue;
      }
      any = true;
      if (c->cache_id() > max_id) {
        max_id = c->cache_id();
      }
    }
    if (any && max_id == std::numeric_limits<std::uint32_t>::max()) {
      return invalid_argument("requested_id must be given because no cache id is left to assign",
                              "max_existing_id=" + std::to_string(max_id));
    }
    assigned = any ? (max_id + 1U) : 1U;
  } else {
    if (find_cache(wb->workbook(), requested_id) != nullptr) {
      return set_binding_error(formulon::FormulonErrorCode::kAlreadyExists, "requested_id already in use",
                               "requested_id=" + std::to_string(requested_id));
    }
  }
  auto cache = std::make_unique<formulon::pivot::PivotCache>();
  cache->set_cache_id(assigned);
  wb->workbook().add_pivot_cache(std::move(cache));
  *out_cache_id = assigned;
  return 0;
}

extern "C" fm_status_t fm_workbook_remove_pivot_cache(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  formulon::Workbook& book = wb->workbook();
  // Reject removal when any pivot still references the cache.
  for (std::size_t s = 0; s < book.sheet_count(); ++s) {
    const formulon::Sheet& sheet = book.sheet(s);
    for (const auto& pt : sheet.pivot_tables()) {
      if (pt != nullptr && pt->pivot_cache_id() == cache_id) {
        return set_binding_error(formulon::FormulonErrorCode::kPreconditionFailed,
                                 "cache still referenced by a pivot table",
                                 "cache_id=" + std::to_string(cache_id) + " sheet_index=" + std::to_string(s));
      }
    }
  }
  auto& caches = book.mutable_pivot_caches();
  for (auto it = caches.begin(); it != caches.end(); ++it) {
    if (*it != nullptr && (*it)->cache_id() == cache_id) {
      caches.erase(it);
      return 0;
    }
  }
  return not_found(("cache_id " + std::to_string(cache_id) + " names no pivot cache").c_str(),
                   "cache_id=" + std::to_string(cache_id));
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_worksheet_source(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::int32_t* out_present,
                                                                    fm_pivot_worksheet_source_t* out) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  const formulon::pivot::WorksheetSource& src = cache->worksheet_source();
  const auto view = [](const std::string& text) { return text.empty() ? nullptr : text.c_str(); };
  fm_pivot_worksheet_source_t result{};
  result.struct_size = out->struct_size;
  if (src.present) {
    result.ref = view(src.ref);
    result.sheet = view(src.sheet);
    result.name = view(src.name);
  }
  *out_present = src.present ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_worksheet_source(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    const fm_pivot_worksheet_source_t* source) {
  clear_last_error();
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  formulon::pivot::WorksheetSource& src = cache->mutable_worksheet_source();
  if (source == nullptr) {
    src = formulon::pivot::WorksheetSource{};
    invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
    return 0;
  }
  if (auto rc = check_record_size(*source); rc != 0) {
    return rc;
  }
  if (auto rc = check_utf8({source->ref, source->sheet, source->name}); rc != 0) {
    return rc;
  }
  // A worksheet source replaces whatever source kind the cache was read with,
  // so the captured `<cacheSource>` type and its other children go with it.
  cache->mutable_cache_source_attrs().clear();
  cache->mutable_cache_source_children_xml().clear();
  src.present = true;
  src.ref = text_or_empty(source->ref);
  src.sheet = text_or_empty(source->sheet);
  src.name = text_or_empty(source->name);
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

// ---------------------------------------------------------------------------
// PivotCache fields + shared items
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_count(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  *out_count = cache->fields().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_name(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                              std::size_t field_index, const char** out_utf8) {
  clear_last_error();
  if (wb == nullptr || out_utf8 == nullptr) {
    return null_argument();
  }
  fm_status_t status = 0;
  const auto* field = lookup_cache_field(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  *out_utf8 = field->name.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field(fm_workbook_t* wb, std::uint32_t cache_id,
                                                         const char* utf8_name, std::size_t* out_field_index) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr || out_field_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_utf8({utf8_name}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotCacheField field;
  field.name = utf8_name;
  // The engine appends one unplaced field to every pivot drawing from the cache.
  auto added = wb->workbook().add_pivot_cache_field(cache_id, std::move(field));
  if (!added) {
    return set_last_error(added.error());
  }
  *out_field_index = added.value();
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_clear_pivot_cache_fields(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  cache->mutable_fields().clear();
  cache->mutable_records().clear();
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_shared_item_count(const fm_workbook_t* wb,
                                                                           std::uint32_t cache_id,
                                                                           std::size_t field_index,
                                                                           std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  fm_status_t status = 0;
  const auto* field = lookup_cache_field(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  *out_count = field->shared_items.size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_shared_item_value(const fm_workbook_t* wb,
                                                                           std::uint32_t cache_id,
                                                                           std::size_t field_index,
                                                                           std::size_t item_index, fm_value_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  fm_status_t status = 0;
  const auto* field = lookup_cache_field(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  if (fm_status_t rc = check_index(item_index, field->shared_items.size(), "item_index"); rc != 0) {
    return rc;
  }
  value_to_fm(field->shared_items[item_index], fresh_read_scratch(wb), out);
  return 0;
}

extern "C" fm_status_t fm_workbook_is_pivot_cache_field_shared_item_date(const fm_workbook_t* wb,
                                                                         std::uint32_t cache_id,
                                                                         std::size_t field_index,
                                                                         std::size_t item_index,
                                                                         std::int32_t* out_is_date) {
  clear_last_error();
  if (wb == nullptr || out_is_date == nullptr) {
    return null_argument();
  }
  fm_status_t status = 0;
  const auto* field = lookup_cache_field(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  if (fm_status_t rc = check_index(item_index, field->shared_items.size(), "item_index"); rc != 0) {
    return rc;
  }
  *out_is_date = item_index < field->shared_item_is_date.size() && field->shared_item_is_date[item_index] ? 1 : 0;
  return 0;
}

namespace {

// NULL guard for `wb` and the valid-date check for `serial`.
fm_status_t check_date_serial(const fm_workbook_t* wb, double serial) {
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (!formulon::pivot::is_valid_date_serial(serial, wb->workbook().date1904())) {
    return invalid_argument("serial is not a valid date serial");
  }
  return 0;
}

// Common front-half for the four `add_shared_item_*` entry points: NULL
// guard, lookup, and field-bounds check. Returns the cache field on
// success, or NULL after writing the binding error and reporting the
// matching status through `out_status`, so the caller returns the same
// code `fm_last_error_code` records.
formulon::pivot::PivotCacheField* lookup_cache_field_mut(fm_workbook_t* wb, std::uint32_t cache_id,
                                                         std::size_t field_index, fm_status_t* out_status,
                                                         formulon::pivot::PivotCache** out_cache = nullptr) {
  formulon::pivot::PivotCache* cache = nullptr;
  if (*out_status = resolve_cache_mut(wb, cache_id, &cache); *out_status != 0) {
    return nullptr;
  }
  if (*out_status = check_index(field_index, cache->fields().size(), "field_index"); *out_status != 0) {
    return nullptr;
  }
  // The shared-item mutators all resolve their field through here, so
  // invalidate dependent layout caches once at the shared choke point.
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  if (out_cache != nullptr) {
    *out_cache = cache;
  }
  return &cache->mutable_fields()[field_index];
}

}  // namespace

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_number(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                            std::size_t field_index, double value,
                                                                            std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  if (const fm_status_t st = check_finite(value, "value")) {
    return st;
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::number(value));
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_date(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, double serial,
                                                                          std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  if (const fm_status_t st = check_finite(serial, "serial")) {
    return st;
  }
  if (const fm_status_t st = check_date_serial(wb, serial)) {
    return st;
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_item_is_date.resize(field->shared_items.size(), false);
  field->shared_items.push_back(formulon::Value::number(serial));
  field->shared_item_is_date.push_back(true);
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_text(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, const char* utf8,
                                                                          std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  if (utf8 == nullptr) {
    return null_argument("utf8");
  }
  if (auto rc = check_utf8({utf8}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotCache* cache = nullptr;
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status, &cache);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(intern_cache_text(*cache, std::string_view(utf8)));
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_bool(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, std::int32_t value,
                                                                          std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::boolean(value != 0));
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_blank(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_index,
                                                                           std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::blank());
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_error(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_index,
                                                                           fm_error_code_t error,
                                                                           std::size_t* out_shared_item_index) {
  clear_last_error();
  if (out_shared_item_index == nullptr) {
    return null_argument("out_shared_item_index");
  }
  if (!is_valid_error_code(error)) {
    return invalid_error_code(error);
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
  *out_shared_item_index = field->shared_items.size() - 1U;
  return 0;
}

extern "C" fm_status_t fm_workbook_clear_pivot_cache_field_shared_items(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                        std::size_t field_index) {
  clear_last_error();
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.clear();
  field->shared_item_is_date.clear();
  return 0;
}

// ---------------------------------------------------------------------------
// PivotCache records
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_get_pivot_cache_record_count(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  *out_count = cache->records().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_record_value(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t record_index, std::size_t field_index,
                                                                fm_value_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(record_index, cache->records().size(), "record_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(field_index, cache->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotCacheRecord& record = cache->records()[record_index];
  formulon::Value value = formulon::Value::blank();
  if (field_index < record.cells.size()) {
    value = record.cells[field_index];
    const bool indexed = field_index < record.cell_is_index.size() && record.cell_is_index[field_index];
    if (indexed && value.is_number()) {
      const std::vector<formulon::Value>& shared = cache->fields()[field_index].shared_items;
      const double slot = value.as_number();
      value = slot >= 0.0 && slot < static_cast<double>(shared.size()) ? shared[static_cast<std::size_t>(slot)]
                                                                       : formulon::Value::blank();
    }
  }
  value_to_fm(value, fresh_read_scratch(wb), out);
  return 0;
}

extern "C" fm_status_t fm_workbook_is_pivot_cache_record_value_date(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::size_t record_index, std::size_t field_index,
                                                                    std::int32_t* out_is_date) {
  clear_last_error();
  if (wb == nullptr || out_is_date == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(record_index, cache->records().size(), "record_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(field_index, cache->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotCacheRecord& record = cache->records()[record_index];
  bool is_date = false;
  if (field_index < record.cells.size()) {
    const formulon::Value& value = record.cells[field_index];
    const bool indexed = field_index < record.cell_is_index.size() && record.cell_is_index[field_index];
    if (indexed && value.is_number()) {
      const formulon::pivot::PivotCacheField& field = cache->fields()[field_index];
      const double slot = value.as_number();
      is_date = slot >= 0.0 && slot < static_cast<double>(field.shared_items.size()) &&
                static_cast<std::size_t>(slot) < field.shared_item_is_date.size() &&
                field.shared_item_is_date[static_cast<std::size_t>(slot)];
    } else if (!indexed) {
      is_date = field_index < record.cell_is_date.size() && record.cell_is_date[field_index];
    }
  }
  *out_is_date = is_date ? 1 : 0;
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_record(fm_workbook_t* wb, std::uint32_t cache_id,
                                                          std::size_t* out_record_index) {
  clear_last_error();
  if (wb == nullptr || out_record_index == nullptr) {
    return null_argument();
  }
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  cache->mutable_records().emplace_back();
  *out_record_index = cache->records().size() - 1U;
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_clear_pivot_cache_records(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  cache->mutable_records().clear();
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

namespace {

// Resolves a record cell pointer for the `record_set_*` family. Returns
// the cell to overwrite on success, or NULL after writing the binding
// error and reporting the matching status through `out_status`, so the
// caller returns the same code `fm_last_error_code` records.
formulon::pivot::PivotCacheRecord* lookup_record_mut(fm_workbook_t* wb, std::uint32_t cache_id,
                                                     std::size_t record_index, std::size_t field_index,
                                                     formulon::pivot::PivotCache** out_cache, fm_status_t* out_status) {
  formulon::pivot::PivotCache* cache = nullptr;
  if (*out_status = resolve_cache_mut(wb, cache_id, &cache); *out_status != 0) {
    return nullptr;
  }
  if (*out_status = check_index(record_index, cache->records().size(), "record_index"); *out_status != 0) {
    return nullptr;
  }
  // A record's cell arity is bounded by the cache's declared field count (the
  // reader builds records with `cells.assign(field_count, ...)`). Reject any
  // `field_index` at or beyond that count before it reaches `grow_record_cells`,
  // where `field_index + 1` would otherwise wrap on `SIZE_MAX` and drive an
  // out-of-bounds write / runaway resize.
  if (*out_status = check_index(field_index, cache->fields().size(), "field_index"); *out_status != 0) {
    return nullptr;
  }
  formulon::pivot::PivotCacheRecord* rec = &cache->mutable_records()[record_index];
  grow_record_cells(*rec, field_index);
  if (out_cache != nullptr) {
    *out_cache = cache;
  }
  // Every record setter routes through here, so invalidating the layout
  // cache of tables drawing from this cache in one place keeps a stale
  // projection from surviving a record edit.
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return rec;
}

// A cell written through the C API is an inline value: it is never the index
// of a shared item (unless the field has shared items and the value is a
// number, the hand-built convention) and never a `<d>` date. Flags the record
// carries for the cell would otherwise survive the overwrite.
void reset_cell_encoding(formulon::pivot::PivotCacheRecord& rec, const formulon::pivot::PivotCacheField& field,
                         std::size_t field_index, const formulon::Value& value) {
  if (field_index < rec.cell_is_index.size()) {
    rec.cell_is_index[field_index] = !field.shared_items.empty() && value.is_number();
  }
  if (field_index < rec.cell_is_date.size()) {
    rec.cell_is_date[field_index] = false;
  }
}

// Shared body of the scalar `record_set_*` setters.
fm_status_t set_record_cell(fm_workbook_t* wb, std::uint32_t cache_id, std::size_t record_index,
                            std::size_t field_index, formulon::Value value) {
  fm_status_t status = 0;
  formulon::pivot::PivotCache* cache = nullptr;
  auto* rec = lookup_record_mut(wb, cache_id, record_index, field_index, &cache, &status);
  if (rec == nullptr) {
    return status;
  }
  reset_cell_encoding(*rec, cache->fields()[field_index], field_index, value);
  rec->cells[field_index] = std::move(value);
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_number(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                 std::size_t record_index, std::size_t field_index,
                                                                 double value) {
  clear_last_error();
  if (const fm_status_t st = check_finite(value, "value")) {
    return st;
  }
  return set_record_cell(wb, cache_id, record_index, field_index, formulon::Value::number(value));
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_date(fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t record_index, std::size_t field_index,
                                                               double serial) {
  clear_last_error();
  if (const fm_status_t st = check_finite(serial, "serial")) {
    return st;
  }
  if (const fm_status_t st = check_date_serial(wb, serial)) {
    return st;
  }
  fm_status_t status = 0;
  auto* rec = lookup_record_mut(wb, cache_id, record_index, field_index, nullptr, &status);
  if (rec == nullptr) {
    return status;
  }
  rec->cells[field_index] = formulon::Value::number(serial);
  rec->cell_is_date.resize(rec->cells.size(), false);
  rec->cell_is_date[field_index] = true;
  if (field_index < rec->cell_is_index.size()) {
    rec->cell_is_index[field_index] = false;
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_text(fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t record_index, std::size_t field_index,
                                                               const char* utf8) {
  clear_last_error();
  if (utf8 == nullptr) {
    return null_argument("utf8");
  }
  if (auto rc = check_utf8({utf8}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotCache* cache = nullptr;
  fm_status_t status = 0;
  auto* rec = lookup_record_mut(wb, cache_id, record_index, field_index, &cache, &status);
  if (rec == nullptr) {
    return status;
  }
  // Overwrite-aware: the cache reuses this coordinate's existing storage
  // slot, so repeated writes to one cell do not grow the text store.
  // `lookup_record_mut` has bounded both indices and grown the record's cells.
  FM_CHECK(cache->set_record_text(record_index, field_index, std::string_view(utf8)),
           "record text coordinate out of range");
  reset_cell_encoding(*rec, cache->fields()[field_index], field_index, rec->cells[field_index]);
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_bool(fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t record_index, std::size_t field_index,
                                                               std::int32_t value) {
  clear_last_error();
  return set_record_cell(wb, cache_id, record_index, field_index, formulon::Value::boolean(value != 0));
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_blank(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t record_index, std::size_t field_index) {
  clear_last_error();
  return set_record_cell(wb, cache_id, record_index, field_index, formulon::Value::blank());
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_record_error(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t record_index, std::size_t field_index,
                                                                fm_error_code_t error) {
  clear_last_error();
  if (!is_valid_error_code(error)) {
    return invalid_error_code(error);
  }
  return set_record_cell(wb, cache_id, record_index, field_index,
                         formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
}

// ---------------------------------------------------------------------------
// PivotCache options
// ---------------------------------------------------------------------------

namespace {

// The largest `missingItemsLimit` the API sets: one item per worksheet row.
constexpr std::uint32_t kMaxMissingItemsLimit = 1048576;

}  // namespace

extern "C" fm_status_t fm_workbook_get_pivot_cache_options(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                           fm_pivot_cache_options_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  fm_pivot_cache_options_t result{};
  result.struct_size = out->struct_size;
  result.save_data = cache->save_data() ? 1 : 0;
  result.has_missing_items_limit = cache->missing_items_limit().has_value() ? 1 : 0;
  result.missing_items_limit = cache->missing_items_limit().value_or(0U);
  result.refresh_on_load = cache->refresh_on_load() ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_update_pivot_cache_options(fm_workbook_t* wb, std::uint32_t cache_id,
                                                              const fm_pivot_cache_options_patch_t* patch) {
  clear_last_error();
  if (wb == nullptr || patch == nullptr) {
    return null_argument();
  }
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*patch); rc != 0) {
    return rc;
  }
  // The whole patch is checked before anything changes; only the limit is clearable.
  if (auto rc = check_patch_op(patch->save_data_op, "save_data", false); rc != 0) {
    return rc;
  }
  if (auto rc = check_patch_op(patch->missing_items_limit_op, "missing_items_limit", true); rc != 0) {
    return rc;
  }
  if (auto rc = check_patch_op(patch->refresh_on_load_op, "refresh_on_load", false); rc != 0) {
    return rc;
  }
  const std::optional<std::uint32_t> held = cache->missing_items_limit();
  // A loaded value outside the settable range survives a get / set round trip unchanged.
  const std::uint32_t limit = patch->missing_items_limit;
  if (patch->missing_items_limit_op == FM_PATCH_OP_SET && limit > kMaxMissingItemsLimit && held != limit) {
    return invalid_argument(("missing_items_limit must be in 0.." + std::to_string(kMaxMissingItemsLimit)).c_str(),
                            "missing_items_limit=" + std::to_string(limit));
  }
  if (patch->save_data_op == FM_PATCH_OP_KEEP && patch->missing_items_limit_op == FM_PATCH_OP_KEEP &&
      patch->refresh_on_load_op == FM_PATCH_OP_KEEP) {
    return 0;
  }
  if (patch->save_data_op == FM_PATCH_OP_SET) {
    cache->set_save_data(patch->save_data != 0);
  }
  if (patch->missing_items_limit_op == FM_PATCH_OP_CLEAR) {
    cache->set_missing_items_limit(std::nullopt);
  } else if (patch->missing_items_limit_op == FM_PATCH_OP_SET && held != limit) {
    cache->set_missing_items_limit(limit);
  }
  if (patch->refresh_on_load_op == FM_PATCH_OP_SET) {
    cache->set_refresh_on_load(patch->refresh_on_load != 0);
  }
  mark_cache_pivots_pending(wb->workbook(), cache_id);
  return 0;
}

// ---------------------------------------------------------------------------
// Calculated fields and calculated items
// ---------------------------------------------------------------------------

namespace {

// Longest calculated formula, in UTF-16 code units.
constexpr std::uint32_t kMaxFormulaUnits = 8192U;

// Validates a required name or formula argument: NULL, malformed UTF-8, empty
// and over-long text are refused.
fm_status_t check_calculated_text(const char* text, const char* what, std::uint32_t max_units) {
  if (text == nullptr) {
    return null_argument(what);
  }
  if (auto rc = check_utf8({text}); rc != 0) {
    return rc;
  }
  if (text[0] == '\0') {
    return invalid_argument((std::string(what) + " must not be empty").c_str());
  }
  if (formulon::utf16_units_in(text) > max_units) {
    return invalid_argument(
        (std::string(what) + " must be at most " + std::to_string(max_units) + " UTF-16 units").c_str());
  }
  return 0;
}

// The evaluator, noting whether its first compile was refused as a syntax
// error. Every edit taking a formula compiles the entered one first.
class EntryHost final : public formulon::pivot::PivotFormulaHost {
 public:
  formulon::Expected<std::uint32_t, formulon::Error> compile(std::string_view formula,
                                                             const formulon::pivot::FormulaScope& scope,
                                                             std::vector<formulon::pivot::FormulaRef>* refs,
                                                             std::string* canonical) override {
    formulon::Expected<std::uint32_t, formulon::Error> id = evaluator_.compile(formula, scope, refs, canonical);
    if (compiles_++ == 0U && !id) {
      const auto code = static_cast<std::int32_t>(id.error().code);
      entry_refused_ = code >= static_cast<std::int32_t>(formulon::FormulonErrorCode::kParserUnexpectedToken) &&
                       code <= static_cast<std::int32_t>(formulon::FormulonErrorCode::kParserTooManyErrors);
    }
    return id;
  }

  formulon::Value evaluate(std::uint32_t formula_id, const formulon::Value* ref_values,
                           std::size_t ref_count) override {
    return evaluator_.evaluate(formula_id, ref_values, ref_count);
  }

  bool entry_refused() const noexcept { return entry_refused_; }

 private:
  formulon::eval::PivotFormulaEvaluator evaluator_;
  std::size_t compiles_ = 0U;
  bool entry_refused_ = false;
};

// Reports `formula` through `formula_parse_failed` (context `field=formula`,
// range into the caller's text) when it does not parse as a plain formula;
// returns 0 when it does.
fm_status_t entered_formula_syntax(std::string_view formula) {
  const std::uint32_t prefix = !formula.empty() && formula.front() == '=' ? 1U : 0U;
  formulon::Arena arena;
  formulon::parser::ParseFailure failure;
  if (formulon::eval::parse_formula_entry_with_diagnostics(formula.substr(prefix), arena, &failure) != nullptr) {
    return 0;
  }
  failure.range.start += prefix;
  failure.range.end += prefix;
  if (failure.range.line == 1U) {
    failure.range.column += prefix;
  }
  return formula_parse_failed(failure, "field=formula");
}

// Runs one calculated-field / -item model edit on cache `cache_id` and every
// table drawing from it, formulas compiled and names collated under the
// workbook's profile. The edit changes the cache, so the tables are left
// pending rather than projected. When the edit refuses `entered` (non-NULL
// for the edits taking a formula) and the text is malformed as a plain
// formula too, the failure is the parser's. Pivot names the plain parser
// cannot read are left to the host, which accepts them.
template <typename Edit>
fm_status_t edit_calculated(fm_workbook_t* wb, std::uint32_t cache_id, const char* entered, Edit&& edit) {
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  const formulon::eval::ScopedEvalProfile profile_scope(wb->workbook().excel_profile());
  EntryHost host;
  const formulon::Expected<void, formulon::Error> done =
      std::forward<Edit>(edit)(*cache, cache_tables(wb->workbook(), cache_id), host);
  if (!done) {
    if (entered != nullptr && host.entry_refused()) {
      if (fm_status_t rc = entered_formula_syntax(entered); rc != 0) {
        return rc;
      }
    }
    return set_last_error(done.error());
  }
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

using CalculatedTables = std::vector<formulon::pivot::PivotTable*>;

// `edit_calculated` of an add entering `formula`: `add(cache, tables, host)`
// returns the new index, which lands in `*out_index` on success.
template <typename Add>
fm_status_t add_calculated(fm_workbook_t* wb, std::uint32_t cache_id, const char* formula, std::size_t* out_index,
                           Add&& add) {
  std::size_t added = 0;
  const fm_status_t status =
      edit_calculated(wb, cache_id, formula,
                      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables,
                          formulon::pivot::PivotFormulaHost& host) -> formulon::Expected<void, formulon::Error> {
                        formulon::Expected<std::size_t, formulon::Error> index = add(cache, tables, host);
                        if (!index) {
                          return index.error();
                        }
                        added = index.value();
                        return formulon::Expected<void, formulon::Error>::Ok();
                      });
  if (status == 0) {
    *out_index = added;
  }
  return status;
}

}  // namespace

extern "C" fm_status_t fm_workbook_add_pivot_cache_calculated_field(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    const char* name, const char* formula,
                                                                    std::size_t* out_field_index) {
  clear_last_error();
  if (wb == nullptr || out_field_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_calculated_text(name, "name", kMaxPivotNameUnits); rc != 0) {
    return rc;
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  return add_calculated(
      wb, cache_id, formula, out_field_index,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::add_calculated_field(cache, tables, name, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_calculated_field_formula(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                            std::size_t field_index,
                                                                            const char* formula) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  return edit_calculated(
      wb, cache_id, formula,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::set_calculated_field_formula(cache, field_index, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_remove_pivot_cache_calculated_field(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                       std::size_t field_index) {
  clear_last_error();
  return edit_calculated(
      wb, cache_id, nullptr,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::remove_calculated_field(cache, tables, field_index, host);
      });
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_formula(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                 std::size_t field_index, std::int32_t* out_present,
                                                                 const char** out_formula) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out_formula == nullptr) {
    return null_argument();
  }
  fm_status_t status = 0;
  const auto* field = lookup_cache_field(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  if (!field->formula.has_value() || field->formula->empty()) {
    *out_present = 0;
    *out_formula = nullptr;
  } else {
    *out_present = 1;
    *out_formula = publish_scratch(wb, with_formula_prefix(display_formula_text(wb->workbook(), *field->formula, {})));
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_calculated_item(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                   std::size_t field_index, const char* name,
                                                                   const char* formula, std::size_t* out_index) {
  clear_last_error();
  if (wb == nullptr || out_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_calculated_text(name, "name", kMaxPivotNameUnits); rc != 0) {
    return rc;
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  return add_calculated(
      wb, cache_id, formula, out_index,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::add_calculated_item(cache, tables, field_index, name, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_calculated_item_count(const fm_workbook_t* wb,
                                                                         std::uint32_t cache_id,
                                                                         std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  *out_count = cache->calculated_items().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_calculated_item_at(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                      std::size_t idx,
                                                                      fm_pivot_calculated_item_info_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_index(idx, cache->calculated_items().size(), "item_index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const formulon::pivot::PivotCalculatedItem& entry = cache->calculated_items()[idx];
  // The item is the one the entry's area selects on its own field; any
  // reference to another field scopes the formula to an intersection.
  std::string_view name;
  bool scoped = false;
  for (const formulon::pivot::PivotAreaReference& reference : entry.area.references) {
    if (!reference.field.has_value()) {
      continue;
    }
    if (*reference.field != entry.cache_field_index) {
      scoped = true;
    } else if (name.empty() && !reference.items.empty() && entry.cache_field_index < cache->fields().size()) {
      const std::vector<formulon::Value>& shared = cache->fields()[entry.cache_field_index].shared_items;
      const std::uint32_t item = reference.items.front();
      if (item < shared.size() && shared[item].is_text()) {
        name = shared[item].as_text();
      }
    }
  }
  formulon::c_api::parts::TextStore& store = fresh_read_scratch(wb);
  fm_pivot_calculated_item_info_t result{};
  result.struct_size = out->struct_size;
  result.field_index = entry.cache_field_index;
  result.name = name.empty() ? nullptr : store_cstr(store, name);
  result.formula = store_cstr(store, with_formula_prefix(entry.formula));
  result.scoped = entry.has_area && scoped ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_calculated_item_formula(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t idx, const char* formula) {
  clear_last_error();
  if (wb == nullptr) {
    return null_argument("wb");
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  return edit_calculated(
      wb, cache_id, formula,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::set_calculated_item_formula(cache, idx, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_move_pivot_cache_calculated_item(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::size_t idx, std::size_t new_index) {
  clear_last_error();
  return edit_calculated(
      wb, cache_id, nullptr,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::pivot::PivotFormulaHost&) {
        return formulon::pivot::move_calculated_item(cache, idx, new_index);
      });
}

extern "C" fm_status_t fm_workbook_remove_pivot_cache_calculated_item(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                      std::size_t idx) {
  clear_last_error();
  return edit_calculated(
      wb, cache_id, nullptr,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables, formulon::pivot::PivotFormulaHost& host) {
        return formulon::pivot::remove_calculated_item(cache, tables, idx, host);
      });
}
