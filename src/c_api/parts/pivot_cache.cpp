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
#include "eval/eval_profile_scope.h"
#include "eval/pivot_formula.h"
#include "pivot/calculated.h"
#include "pivot/date_serial.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::cache_tables;
using formulon::c_api::parts::check_finite;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_utf8;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::display_formula;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::find_cache_mut;
using formulon::c_api::parts::fresh_read_scratch;
using formulon::c_api::parts::grow_record_cells;
using formulon::c_api::parts::intern_cache_text;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::invalidate_pivot_results_for_cache;
using formulon::c_api::parts::mark_cache_pivots_pending;
using formulon::c_api::parts::may_respell_formula;
using formulon::c_api::parts::mutable_pivot_caches;
using formulon::c_api::parts::not_found;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::store_cstr;
using formulon::c_api::parts::value_to_fm;

namespace {

bool is_valid_error_code(fm_error_code_t error) {
  return error >= 0 && error <= static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown);
}

fm_status_t set_invalid_error_code(fm_error_code_t error) {
  return invalid_argument("error code out of range", "error=" + std::to_string(error));
}

fm_status_t cache_not_found(std::uint32_t cache_id) {
  return invalid_argument("cache_id not found", "cache_id=" + std::to_string(cache_id));
}

// Resolves `cache_id` to its cache, writing the binding error and returning
// its status on a miss.
fm_status_t resolve_cache(const fm_workbook_t* wb, std::uint32_t cache_id, const formulon::pivot::PivotCache** out) {
  *out = find_cache(wb->workbook(), cache_id);
  if (*out == nullptr) {
    return cache_not_found(cache_id);
  }
  return 0;
}

// Mutable variant of `resolve_cache` that also rejects a NULL workbook.
fm_status_t resolve_cache_mut(fm_workbook_t* wb, std::uint32_t cache_id, formulon::pivot::PivotCache** out) {
  if (wb == nullptr) {
    return null_argument("wb");
  }
  *out = find_cache_mut(wb->workbook(), cache_id);
  if (*out == nullptr) {
    return cache_not_found(cache_id);
  }
  return 0;
}

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
      return invalid_argument("auto-assignment would overflow", "max_existing_id=" + std::to_string(max_id));
    }
    assigned = any ? (max_id + 1U) : 1U;
  } else {
    if (find_cache(wb->workbook(), requested_id) != nullptr) {
      return invalid_argument("requested_id already in use", "requested_id=" + std::to_string(requested_id));
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
        return invalid_argument("cache still referenced by a pivot table",
                                "cache_id=" + std::to_string(cache_id) + " sheet_index=" + std::to_string(s));
      }
    }
  }
  auto& caches = mutable_pivot_caches(book);
  for (auto it = caches.begin(); it != caches.end(); ++it) {
    if (*it != nullptr && (*it)->cache_id() == cache_id) {
      caches.erase(it);
      return 0;
    }
  }
  return not_found("no pivot cache has that cache_id", "cache_id=" + std::to_string(cache_id));
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_worksheet_source(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::int32_t* out_present, const char** out_ref,
                                                                    const char** out_sheet, const char** out_name) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out_ref == nullptr || out_sheet == nullptr || out_name == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  const formulon::pivot::WorksheetSource& src = cache->worksheet_source();
  *out_present = src.present ? 1 : 0;
  *out_ref = src.ref.c_str();
  *out_sheet = src.sheet.c_str();
  *out_name = src.name.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_worksheet_source(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::int32_t present, const char* ref,
                                                                    const char* sheet, const char* name) {
  clear_last_error();
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  formulon::pivot::WorksheetSource& src = cache->mutable_worksheet_source();
  if (present == 0) {
    src = formulon::pivot::WorksheetSource{};
    invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
    return 0;
  }
  if (auto rc = check_utf8({ref, sheet, name}); rc != 0) {
    return rc;
  }
  // A worksheet source replaces whatever source kind the cache was read with,
  // so the captured `<cacheSource>` type and its other children go with it.
  cache->mutable_cache_source_attrs().clear();
  cache->mutable_cache_source_children_xml().clear();
  src.present = true;
  src.ref = ref != nullptr ? ref : "";
  src.sheet = sheet != nullptr ? sheet : "";
  src.name = name != nullptr ? name : "";
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
  cache->mutable_fields().push_back(std::move(field));
  *out_field_index = cache->fields().size() - 1U;
  // Every table keeps one pivot field per cache field.
  formulon::pivot::sync_table_fields(*cache, cache_tables(wb->workbook(), cache_id));
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
    return invalid_argument("serial is not a valid date");
  }
  return 0;
}

// Common front-half for the four `add_shared_item_*` entry points: NULL
// guard, lookup, and field-bounds check. Returns the cache field on
// success, or NULL after writing the binding error and reporting the
// matching status through `out_status`, so the caller returns the same
// code `fm_last_error_code` records.
formulon::pivot::PivotCacheField* lookup_cache_field_mut(fm_workbook_t* wb, std::uint32_t cache_id,
                                                         std::size_t field_index, fm_status_t* out_status) {
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
  return &cache->mutable_fields()[field_index];
}

}  // namespace

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_number(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                            std::size_t field_index, double value) {
  clear_last_error();
  if (const fm_status_t st = check_finite(value, "value")) {
    return st;
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::number(value));
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_date(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, double serial) {
  clear_last_error();
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
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_text(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, const char* utf8) {
  clear_last_error();
  if (utf8 == nullptr) {
    return null_argument("utf8");
  }
  if (auto rc = check_utf8({utf8}); rc != 0) {
    return rc;
  }
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (auto rc = check_index(field_index, cache->fields().size(), "field_index"); rc != 0) {
    return rc;
  }
  cache->mutable_fields()[field_index].shared_items.push_back(intern_cache_text(*cache, std::string_view(utf8)));
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_bool(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_index, std::int32_t value) {
  clear_last_error();
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::boolean(value != 0));
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_blank(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_index) {
  clear_last_error();
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::blank());
  return 0;
}

extern "C" fm_status_t fm_workbook_add_pivot_cache_field_shared_item_error(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_index,
                                                                           fm_error_code_t error) {
  clear_last_error();
  if (!is_valid_error_code(error)) {
    return set_invalid_error_code(error);
  }
  fm_status_t status = 0;
  auto* field = lookup_cache_field_mut(wb, cache_id, field_index, &status);
  if (field == nullptr) {
    return status;
  }
  field->shared_items.push_back(formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
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
  if (!cache->set_record_text(record_index, field_index, std::string_view(utf8))) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
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
    return set_invalid_error_code(error);
  }
  return set_record_cell(wb, cache_id, record_index, field_index,
                         formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
}

// ---------------------------------------------------------------------------
// PivotCache options
// ---------------------------------------------------------------------------

namespace {

// The largest `missingItemsLimit` the API sets: one item per worksheet row.
constexpr std::int64_t kMaxMissingItemsLimit = 1048576;
constexpr std::int32_t kAutomaticMissingItemsLimit = -1;

std::int32_t missing_items_limit_to_fm(std::optional<std::uint32_t> limit) {
  if (!limit.has_value()) {
    return kAutomaticMissingItemsLimit;
  }
  return *limit > static_cast<std::uint32_t>(std::numeric_limits<std::int32_t>::max())
             ? std::numeric_limits<std::int32_t>::max()
             : static_cast<std::int32_t>(*limit);
}

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
  fm_pivot_cache_options_t result{};
  result.save_data = cache->save_data() ? 1 : 0;
  result.missing_items_limit = missing_items_limit_to_fm(cache->missing_items_limit());
  result.refresh_on_load = cache->refresh_on_load() ? 1 : 0;
  *out = result;
  return 0;
}

extern "C" fm_status_t fm_workbook_set_pivot_cache_options(fm_workbook_t* wb, std::uint32_t cache_id,
                                                           const fm_pivot_cache_options_t* options) {
  clear_last_error();
  if (wb == nullptr || options == nullptr) {
    return null_argument();
  }
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  // -1 is automatic; a loaded value outside the settable range survives a
  // get / set round trip unchanged.
  const std::int32_t limit = options->missing_items_limit;
  if (limit != kAutomaticMissingItemsLimit && (limit < 0 || limit > kMaxMissingItemsLimit) &&
      limit != missing_items_limit_to_fm(cache->missing_items_limit())) {
    return invalid_argument("missing_items_limit out of range", "missing_items_limit=" + std::to_string(limit));
  }
  cache->set_save_data(options->save_data != 0);
  if (limit != missing_items_limit_to_fm(cache->missing_items_limit())) {
    cache->set_missing_items_limit(limit == kAutomaticMissingItemsLimit
                                       ? std::nullopt
                                       : std::optional<std::uint32_t>(static_cast<std::uint32_t>(limit)));
  }
  cache->set_refresh_on_load(options->refresh_on_load != 0);
  mark_cache_pivots_pending(wb->workbook(), cache_id);
  return 0;
}

// ---------------------------------------------------------------------------
// Calculated fields and calculated items
// ---------------------------------------------------------------------------

namespace {

// Longest calculated field or item name, in UTF-16 code units.
constexpr std::uint32_t kMaxCalculatedNameUnits = 255U;
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
    return invalid_argument((std::string(what) + " is empty").c_str());
  }
  if (formulon::utf16_units_in(text) > max_units) {
    return invalid_argument((std::string(what) + " is too long").c_str());
  }
  return 0;
}

// Runs one calculated-field / -item model edit on cache `cache_id` and every
// table drawing from it, formulas compiled and names collated under the
// workbook's profile. The edit changes the cache, so the tables are left
// pending rather than projected.
template <typename Edit>
fm_status_t edit_calculated(fm_workbook_t* wb, std::uint32_t cache_id, Edit&& edit) {
  formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache_mut(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  const formulon::eval::ScopedEvalProfile profile_scope(wb->workbook().excel_profile());
  formulon::eval::PivotFormulaEvaluator host;
  const formulon::Expected<void, formulon::Error> done =
      std::forward<Edit>(edit)(*cache, cache_tables(wb->workbook(), cache_id), host);
  if (!done) {
    return set_last_error(done.error());
  }
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

using CalculatedTables = std::vector<formulon::pivot::PivotTable*>;

}  // namespace

extern "C" fm_status_t fm_workbook_add_pivot_cache_calculated_field(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    const char* name, const char* formula,
                                                                    std::size_t* out_field_index) {
  clear_last_error();
  if (wb == nullptr || out_field_index == nullptr) {
    return null_argument();
  }
  if (auto rc = check_calculated_text(name, "name", kMaxCalculatedNameUnits); rc != 0) {
    return rc;
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  std::size_t added = 0;
  const fm_status_t status =
      edit_calculated(wb, cache_id,
                      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables,
                          formulon::eval::PivotFormulaEvaluator& host) -> formulon::Expected<void, formulon::Error> {
                        formulon::Expected<std::size_t, formulon::Error> index =
                            formulon::pivot::add_calculated_field(cache, tables, name, formula, host);
                        if (!index) {
                          return index.error();
                        }
                        added = index.value();
                        return formulon::Expected<void, formulon::Error>::Ok();
                      });
  if (status == 0) {
    *out_field_index = added;
  }
  return status;
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
      wb, cache_id,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::eval::PivotFormulaEvaluator& host) {
        return formulon::pivot::set_calculated_field_formula(cache, field_index, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_remove_pivot_cache_calculated_field(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                       std::size_t field_index) {
  clear_last_error();
  return edit_calculated(wb, cache_id,
                         [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables,
                             formulon::eval::PivotFormulaEvaluator& host) {
                           return formulon::pivot::remove_calculated_field(cache, tables, field_index, host);
                         });
}

extern "C" fm_status_t fm_workbook_get_pivot_cache_field_formula(const fm_workbook_t* wb, std::uint32_t cache_id,
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
  if (!field->formula.has_value()) {
    *out_utf8 = "";
  } else if (may_respell_formula(*field->formula)) {
    *out_utf8 = display_formula(wb->workbook(), *field->formula, fresh_read_scratch(wb), {});
  } else {
    *out_utf8 = field->formula->c_str();
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
  if (auto rc = check_calculated_text(name, "name", kMaxCalculatedNameUnits); rc != 0) {
    return rc;
  }
  if (auto rc = check_calculated_text(formula, "formula", kMaxFormulaUnits); rc != 0) {
    return rc;
  }
  std::size_t added = 0;
  const fm_status_t status =
      edit_calculated(wb, cache_id,
                      [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables,
                          formulon::eval::PivotFormulaEvaluator& host) -> formulon::Expected<void, formulon::Error> {
                        formulon::Expected<std::size_t, formulon::Error> index =
                            formulon::pivot::add_calculated_item(cache, tables, field_index, name, formula, host);
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

extern "C" fm_status_t fm_workbook_get_pivot_cache_calculated_item(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                   std::size_t idx, fm_pivot_calculated_item_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return null_argument();
  }
  const formulon::pivot::PivotCache* cache = nullptr;
  if (fm_status_t rc = resolve_cache(wb, cache_id, &cache); rc != 0) {
    return rc;
  }
  if (fm_status_t rc = check_index(idx, cache->calculated_items().size(), "index"); rc != 0) {
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
  fm_pivot_calculated_item_t result{};
  result.field_index = entry.cache_field_index;
  result.name = store_cstr(store, name);
  result.formula = store_cstr(store, entry.formula);
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
      wb, cache_id,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::eval::PivotFormulaEvaluator& host) {
        return formulon::pivot::set_calculated_item_formula(cache, idx, formula, host);
      });
}

extern "C" fm_status_t fm_workbook_move_pivot_cache_calculated_item(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::size_t idx, std::size_t new_index) {
  clear_last_error();
  return edit_calculated(
      wb, cache_id,
      [&](formulon::pivot::PivotCache& cache, const CalculatedTables&, formulon::eval::PivotFormulaEvaluator&) {
        return formulon::pivot::move_calculated_item(cache, idx, new_index);
      });
}

extern "C" fm_status_t fm_workbook_remove_pivot_cache_calculated_item(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                      std::size_t idx) {
  clear_last_error();
  return edit_calculated(wb, cache_id,
                         [&](formulon::pivot::PivotCache& cache, const CalculatedTables& tables,
                             formulon::eval::PivotFormulaEvaluator& host) {
                           return formulon::pivot::remove_calculated_item(cache, tables, idx, host);
                         });
}
