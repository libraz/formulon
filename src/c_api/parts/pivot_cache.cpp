//
// C ABI - PivotCache mutation surface (caches, fields, shared items,
// records). The PivotTable mutation surface lives in `pivot_table.cpp`;
// the two share helpers from `pivot_internal.h`.

#include "pivot/pivot_cache.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/pivot_internal.h"
#include "sheet.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::find_cache;
using formulon::c_api::parts::find_cache_mut;
using formulon::c_api::parts::grow_record_cells;
using formulon::c_api::parts::intern_cache_text;
using formulon::c_api::parts::invalidate_pivot_results_for_cache;
using formulon::c_api::parts::mutable_pivot_caches;
using formulon::c_api::parts::set_binding_error;

namespace {

bool is_valid_error_code(fm_error_code_t error) {
  return error >= 0 && error <= static_cast<fm_error_code_t>(formulon::ErrorCode::Unknown);
}

fm_status_t set_invalid_error_code(const char* fn, fm_error_code_t error) {
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                           (std::string(fn) + ": error code out of range").c_str(), "error=" + std::to_string(error));
}

}  // namespace

// ---------------------------------------------------------------------------
// PivotCache lifecycle / enumeration
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_cache_count(const fm_workbook_t* wb, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_count: NULL argument");
  }
  *out_count = wb->workbook().pivot_caches().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_id_at(const fm_workbook_t* wb, std::size_t idx,
                                                     std::uint32_t* out_cache_id) {
  clear_last_error();
  if (wb == nullptr || out_cache_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_id_at: NULL argument");
  }
  const auto& caches = wb->workbook().pivot_caches();
  if (idx >= caches.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_id_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(caches.size()));
  }
  *out_cache_id = caches[idx]->cache_id();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_create(fm_workbook_t* wb, std::uint32_t requested_id,
                                                      std::uint32_t* out_cache_id) {
  clear_last_error();
  if (wb == nullptr || out_cache_id == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_create: NULL argument");
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
    assigned = any ? (max_id + 1U) : 1U;
  } else {
    if (find_cache(wb->workbook(), requested_id) != nullptr) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_workbook_pivot_cache_create: requested_id already in use",
                               "requested_id=" + std::to_string(requested_id));
    }
  }
  auto cache = std::make_unique<formulon::pivot::PivotCache>();
  cache->set_cache_id(assigned);
  wb->workbook().add_pivot_cache(std::move(cache));
  *out_cache_id = assigned;
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_remove(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_remove: wb is NULL");
  }
  formulon::Workbook& book = wb->workbook();
  // Reject removal when any pivot still references the cache.
  for (std::size_t s = 0; s < book.sheet_count(); ++s) {
    const formulon::Sheet& sheet = book.sheet(s);
    for (const auto& pt : sheet.pivot_tables()) {
      if (pt != nullptr && pt->pivot_cache_id() == cache_id) {
        return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                                 "fm_workbook_pivot_cache_remove: cache still referenced by a pivot table",
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
  return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                           "fm_workbook_pivot_cache_remove: cache_id not found",
                           "cache_id=" + std::to_string(cache_id));
}

extern "C" fm_status_t fm_workbook_pivot_cache_get_worksheet_source(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::int32_t* out_present, const char** out_ref,
                                                                    const char** out_sheet, const char** out_name) {
  clear_last_error();
  if (wb == nullptr || out_present == nullptr || out_ref == nullptr || out_sheet == nullptr || out_name == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_get_worksheet_source: NULL argument");
  }
  const auto* cache = find_cache(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_get_worksheet_source: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  const formulon::pivot::WorksheetSource& src = cache->worksheet_source();
  *out_present = src.present ? 1 : 0;
  *out_ref = src.ref.c_str();
  *out_sheet = src.sheet.c_str();
  *out_name = src.name.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_set_worksheet_source(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                    std::int32_t present, const char* ref,
                                                                    const char* sheet, const char* name) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_set_worksheet_source: wb is NULL");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_set_worksheet_source: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  formulon::pivot::WorksheetSource& src = cache->mutable_worksheet_source();
  if (present == 0) {
    src = formulon::pivot::WorksheetSource{};
    invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
    return 0;
  }
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

extern "C" fm_status_t fm_workbook_pivot_cache_field_count(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                           std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_count: NULL argument");
  }
  const auto* cache = find_cache(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_count: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  *out_count = cache->fields().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_name(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                          std::size_t field_idx, const char** out_utf8) {
  clear_last_error();
  if (wb == nullptr || out_utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_name: NULL argument");
  }
  const auto* cache = find_cache(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_name: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  if (field_idx >= cache->fields().size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_name: field_idx out of range",
                             "field_idx=" + std::to_string(field_idx));
  }
  *out_utf8 = cache->fields()[field_idx].name.c_str();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_add(fm_workbook_t* wb, std::uint32_t cache_id,
                                                         const char* utf8_name, std::size_t* out_field_idx) {
  clear_last_error();
  if (wb == nullptr || utf8_name == nullptr || out_field_idx == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_add: NULL argument");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_add: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  formulon::pivot::PivotCacheField field;
  field.name = utf8_name;
  cache->mutable_fields().push_back(std::move(field));
  *out_field_idx = cache->fields().size() - 1U;
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_clear(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_clear: wb is NULL");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_clear: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  cache->mutable_fields().clear();
  cache->mutable_records().clear();
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_shared_item_count(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                                       std::size_t field_idx, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_shared_item_count: NULL argument");
  }
  const auto* cache = find_cache(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_shared_item_count: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  if (field_idx >= cache->fields().size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_shared_item_count: field_idx out of range",
                             "field_idx=" + std::to_string(field_idx));
  }
  *out_count = cache->fields()[field_idx].shared_items.size();
  return 0;
}

namespace {

// Common front-half for the four `add_shared_item_*` entry points: NULL
// guard, lookup, and field-bounds check. Returns the cache field on
// success, or NULL after writing the binding error.
formulon::pivot::PivotCacheField* lookup_cache_field_mut(fm_workbook_t* wb, std::uint32_t cache_id,
                                                         std::size_t field_idx, const char* fn) {
  if (wb == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, (std::string(fn) + ": wb is NULL").c_str());
    return nullptr;
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, (std::string(fn) + ": cache_id not found").c_str(),
                      "cache_id=" + std::to_string(cache_id));
    return nullptr;
  }
  if (field_idx >= cache->fields().size()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                      (std::string(fn) + ": field_idx out of range").c_str(), "field_idx=" + std::to_string(field_idx));
    return nullptr;
  }
  // The shared-item mutators all resolve their field through here, so
  // invalidate dependent layout caches once at the shared choke point.
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return &cache->mutable_fields()[field_idx];
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_cache_field_add_shared_item_number(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                            std::size_t field_idx, double value) {
  clear_last_error();
  auto* field = lookup_cache_field_mut(wb, cache_id, field_idx, "fm_workbook_pivot_cache_field_add_shared_item_number");
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->shared_items.push_back(formulon::Value::number(value));
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_add_shared_item_text(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_idx, const char* utf8) {
  clear_last_error();
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_add_shared_item_text: utf8 is NULL");
  }
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_field_add_shared_item_text: wb is NULL");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_add_shared_item_text: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  if (field_idx >= cache->fields().size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_field_add_shared_item_text: field_idx out of range",
                             "field_idx=" + std::to_string(field_idx));
  }
  cache->mutable_fields()[field_idx].shared_items.push_back(intern_cache_text(*cache, std::string_view(utf8)));
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_add_shared_item_bool(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                          std::size_t field_idx, std::int32_t value) {
  clear_last_error();
  auto* field = lookup_cache_field_mut(wb, cache_id, field_idx, "fm_workbook_pivot_cache_field_add_shared_item_bool");
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->shared_items.push_back(formulon::Value::boolean(value != 0));
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_add_shared_item_blank(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_idx) {
  clear_last_error();
  auto* field = lookup_cache_field_mut(wb, cache_id, field_idx, "fm_workbook_pivot_cache_field_add_shared_item_blank");
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->shared_items.push_back(formulon::Value::blank());
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_add_shared_item_error(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                           std::size_t field_idx,
                                                                           fm_error_code_t error) {
  clear_last_error();
  if (!is_valid_error_code(error)) {
    return set_invalid_error_code("fm_workbook_pivot_cache_field_add_shared_item_error", error);
  }
  auto* field = lookup_cache_field_mut(wb, cache_id, field_idx, "fm_workbook_pivot_cache_field_add_shared_item_error");
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->shared_items.push_back(formulon::Value::error(static_cast<formulon::ErrorCode>(error)));
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_field_clear_shared_items(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                        std::size_t field_idx) {
  clear_last_error();
  auto* field = lookup_cache_field_mut(wb, cache_id, field_idx, "fm_workbook_pivot_cache_field_clear_shared_items");
  if (field == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  field->shared_items.clear();
  return 0;
}

// ---------------------------------------------------------------------------
// PivotCache records
// ---------------------------------------------------------------------------

extern "C" fm_status_t fm_workbook_pivot_cache_record_count(const fm_workbook_t* wb, std::uint32_t cache_id,
                                                            std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_record_count: NULL argument");
  }
  const auto* cache = find_cache(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_record_count: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  *out_count = cache->records().size();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_add(fm_workbook_t* wb, std::uint32_t cache_id,
                                                          std::size_t* out_record_idx) {
  clear_last_error();
  if (wb == nullptr || out_record_idx == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_record_add: NULL argument");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_record_add: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  cache->mutable_records().emplace_back();
  *out_record_idx = cache->records().size() - 1U;
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_clear(fm_workbook_t* wb, std::uint32_t cache_id) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_record_clear: wb is NULL");
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_workbook_pivot_cache_record_clear: cache_id not found",
                             "cache_id=" + std::to_string(cache_id));
  }
  cache->mutable_records().clear();
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return 0;
}

namespace {

// Resolves a record cell pointer for the `record_set_*` family. Returns
// the cell to overwrite on success, or NULL after writing the binding
// error.
formulon::pivot::PivotCacheRecord* lookup_record_mut(fm_workbook_t* wb, std::uint32_t cache_id, std::size_t record_idx,
                                                     std::size_t field_idx, const char* fn,
                                                     formulon::pivot::PivotCache** out_cache) {
  if (wb == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, (std::string(fn) + ": wb is NULL").c_str());
    return nullptr;
  }
  auto* cache = find_cache_mut(wb->workbook(), cache_id);
  if (cache == nullptr) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, (std::string(fn) + ": cache_id not found").c_str(),
                      "cache_id=" + std::to_string(cache_id));
    return nullptr;
  }
  if (record_idx >= cache->records().size()) {
    set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                      (std::string(fn) + ": record_idx out of range").c_str(),
                      "record_idx=" + std::to_string(record_idx));
    return nullptr;
  }
  // A record's cell arity is bounded by the cache's declared field count (the
  // reader builds records with `cells.assign(field_count, ...)`). Reject any
  // `field_idx` at or beyond that count before it reaches `grow_record_cells`,
  // where `field_idx + 1` would otherwise wrap on `SIZE_MAX` and drive an
  // out-of-bounds write / runaway resize.
  if (field_idx >= cache->fields().size()) {
    set_binding_error(
        formulon::FormulonErrorCode::kInvalidArgument, (std::string(fn) + ": field_idx out of range").c_str(),
        "field_idx=" + std::to_string(field_idx) + " field_count=" + std::to_string(cache->fields().size()));
    return nullptr;
  }
  formulon::pivot::PivotCacheRecord* rec = &cache->mutable_records()[record_idx];
  grow_record_cells(*rec, field_idx);
  if (out_cache != nullptr) {
    *out_cache = cache;
  }
  // Every record setter routes through here, so invalidating the layout
  // cache of tables drawing from this cache in one place keeps a stale
  // projection from surviving a record edit.
  invalidate_pivot_results_for_cache(wb->workbook(), cache_id);
  return rec;
}

}  // namespace

extern "C" fm_status_t fm_workbook_pivot_cache_record_set_number(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                 std::size_t record_idx, std::size_t field_idx,
                                                                 double value) {
  clear_last_error();
  auto* rec =
      lookup_record_mut(wb, cache_id, record_idx, field_idx, "fm_workbook_pivot_cache_record_set_number", nullptr);
  if (rec == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  rec->cells[field_idx] = formulon::Value::number(value);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_set_text(fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t record_idx, std::size_t field_idx,
                                                               const char* utf8) {
  clear_last_error();
  if (utf8 == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_workbook_pivot_cache_record_set_text: utf8 is NULL");
  }
  formulon::pivot::PivotCache* cache = nullptr;
  auto* rec = lookup_record_mut(wb, cache_id, record_idx, field_idx, "fm_workbook_pivot_cache_record_set_text", &cache);
  if (rec == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  // Overwrite-aware: the cache reuses this coordinate's existing storage
  // slot, so repeated writes to one cell do not grow the text store.
  if (!cache->set_record_text(record_idx, field_idx, std::string_view(utf8))) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_set_bool(fm_workbook_t* wb, std::uint32_t cache_id,
                                                               std::size_t record_idx, std::size_t field_idx,
                                                               std::int32_t value) {
  clear_last_error();
  auto* rec =
      lookup_record_mut(wb, cache_id, record_idx, field_idx, "fm_workbook_pivot_cache_record_set_bool", nullptr);
  if (rec == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  rec->cells[field_idx] = formulon::Value::boolean(value != 0);
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_set_blank(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t record_idx, std::size_t field_idx) {
  clear_last_error();
  auto* rec =
      lookup_record_mut(wb, cache_id, record_idx, field_idx, "fm_workbook_pivot_cache_record_set_blank", nullptr);
  if (rec == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  rec->cells[field_idx] = formulon::Value::blank();
  return 0;
}

extern "C" fm_status_t fm_workbook_pivot_cache_record_set_error(fm_workbook_t* wb, std::uint32_t cache_id,
                                                                std::size_t record_idx, std::size_t field_idx,
                                                                fm_error_code_t error) {
  clear_last_error();
  if (!is_valid_error_code(error)) {
    return set_invalid_error_code("fm_workbook_pivot_cache_record_set_error", error);
  }
  auto* rec =
      lookup_record_mut(wb, cache_id, record_idx, field_idx, "fm_workbook_pivot_cache_record_set_error", nullptr);
  if (rec == nullptr) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  rec->cells[field_idx] = formulon::Value::error(static_cast<formulon::ErrorCode>(error));
  return 0;
}
