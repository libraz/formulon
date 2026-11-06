// PivotCache mutation bindings.
//
// Each method is a thin wrapper over `fm_workbook_pivot_cache_*`. The
// shape mirrors the embind binding: number / boolean / string args
// map straight through; a method returns nothing, or the freshly assigned
// id / index as a number, and throws on failure. See `c_api/formulon_c.h`
// "Pivot cache mutation" for the underlying contract.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct CacheRecordArgs {
  uint32_t cache_id;
  std::size_t record_index;
  std::size_t field_index;
};

CacheRecordArgs ReadCacheRecord(const Napi::CallbackInfo& info) {
  return {Workbook::ArgU32(info, 0), static_cast<std::size_t>(Workbook::ArgU32(info, 1)),
          static_cast<std::size_t>(Workbook::ArgU32(info, 2))};
}

// Shared body of the per-cache count getters (fields, records).
using CacheCountFn = fm_status_t (*)(const fm_workbook_t*, uint32_t, size_t*);

Napi::Value CacheCountResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, CacheCountFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = Workbook::ArgU32(info, 0);
  std::size_t count = 0;
  const fm_status_t rc = fn(handle, cache_id, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

}  // namespace

Napi::Value Workbook::ListPivotCaches(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  Napi::Array arr = Napi::Array::New(env);
  const fm_status_t rc = CollectList<uint32_t>(
      arr, [&](size_t* n) { return fm_workbook_get_pivot_cache_count(handle_, n); },
      [&](size_t i, uint32_t* id) { return fm_workbook_get_pivot_cache_id_at(handle_, i, id); },
      [&](uint32_t id) { return Napi::Number::New(env, id); });
  return ValueOrThrow(env, rc, arr);
}

Napi::Value Workbook::PivotCacheCreate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t requested = ArgU32(info, 0);
  uint32_t out = 0;
  const fm_status_t rc = fm_workbook_add_pivot_cache(handle_, requested, &out);
  return ValueOrThrow(env, rc, JsNumber(env, out));
}

Napi::Value Workbook::PivotCacheRemove(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_remove_pivot_cache);
}

Napi::Value Workbook::PivotCacheGetWorksheetSource(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t present = 0;
  const char* ref = nullptr;
  const char* sheet = nullptr;
  const char* name = nullptr;
  const fm_status_t rc =
      fm_workbook_get_pivot_cache_worksheet_source(handle_, ArgU32(info, 0), &present, &ref, &sheet, &name);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  if (present == 0) {
    return env.Null();
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("present", Napi::Boolean::New(env, true));
  out.Set("ref", JsStringOrNull(env, ref));
  out.Set("sheet", JsStringOrNull(env, sheet));
  out.Set("name", JsStringOrNull(env, name));
  return out;
}

Napi::Value Workbook::PivotCacheSetWorksheetSource(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  Napi::Object source = ArgObject(info, 1);
  CheckedSpecReader reader(env);
  const bool present = reader.Bool(source, "present", true);
  std::string ref;
  std::string sheet;
  std::string name;
  const bool has_ref = reader.String(source, "ref", &ref);
  const bool has_sheet = reader.String(source, "sheet", &sheet);
  const bool has_name = reader.String(source, "name", &name);
  if (!reader.ok()) {
    return env.Undefined();
  }
  const fm_status_t rc = fm_workbook_set_pivot_cache_worksheet_source(
      handle_, cache_id, present ? 1 : 0, has_ref ? ref.c_str() : nullptr, has_sheet ? sheet.c_str() : nullptr,
      has_name ? name.c_str() : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::PivotCacheFieldCount(const Napi::CallbackInfo& info) {
  return CacheCountResult(info, handle_, &fm_workbook_get_pivot_cache_field_count);
}

Napi::Value Workbook::PivotCacheFieldName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
  const char* name = nullptr;
  const fm_status_t rc = fm_workbook_get_pivot_cache_field_name(handle_, cache_id, field_index, &name);
  return ValueOrThrow(env, rc, JsString(env, name));
}

Napi::Value Workbook::PivotCacheFieldAdd(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const uint32_t cache_id = ArgU32(info, 0);
    const std::string name = ArgString(info, 1);
    return fm_workbook_add_pivot_cache_field(handle, cache_id, name.c_str(), out);
  });
}

Napi::Value Workbook::PivotCacheFieldClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_clear_pivot_cache_fields);
}

Napi::Value Workbook::PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_get_pivot_cache_field_shared_item_count(handle_, cache_id, field_index, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::PivotCacheFieldSharedItemValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t item_index = static_cast<std::size_t>(ArgU32(info, 2));
  fm_value_t v{};
  const fm_status_t rc =
      fm_workbook_get_pivot_cache_field_shared_item_value(handle_, cache_id, field_index, item_index, &v);
  return ValueOrThrow(env, rc, rc == 0 ? ValueToJs(env, v) : env.Undefined());
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_add_pivot_cache_field_shared_item_number);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemDate(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_add_pivot_cache_field_shared_item_date);
}

Napi::Value Workbook::PivotCacheFieldSharedItemIsDate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  int32_t is_date = 0;
  const fm_status_t rc = fm_workbook_is_pivot_cache_field_shared_item_date(
      handle_, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), static_cast<std::size_t>(ArgU32(info, 2)),
      &is_date);
  return ValueOrThrow(env, rc, Napi::Boolean::New(env, is_date != 0));
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const uint32_t cache_id = ArgU32(info, 0);
    const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
    const std::string utf8 = ArgString(info, 2);
    return fm_workbook_add_pivot_cache_field_shared_item_text(handle, cache_id, field_index, utf8.c_str());
  });
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const uint32_t cache_id = ArgU32(info, 0);
    const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
    return fm_workbook_add_pivot_cache_field_shared_item_bool(handle, cache_id, field_index, ArgFlag(info, 2));
  });
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_add_pivot_cache_field_shared_item_blank);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemError(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const uint32_t cache_id = ArgU32(info, 0);
    const std::size_t field_index = static_cast<std::size_t>(ArgU32(info, 1));
    const fm_error_code_t error = static_cast<fm_error_code_t>(ArgI32(info, 2));
    return fm_workbook_add_pivot_cache_field_shared_item_error(handle, cache_id, field_index, error);
  });
}

Napi::Value Workbook::PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_clear_pivot_cache_field_shared_items);
}

Napi::Value Workbook::PivotCacheRecordCount(const Napi::CallbackInfo& info) {
  return CacheCountResult(info, handle_, &fm_workbook_get_pivot_cache_record_count);
}

Napi::Value Workbook::PivotCacheRecordAdd(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    return fm_workbook_add_pivot_cache_record(handle, ArgU32(info, 0), out);
  });
}

Napi::Value Workbook::PivotCacheRecordClear(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_clear_pivot_cache_records);
}

Napi::Value Workbook::PivotCacheRecordSetNumber(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_set_pivot_cache_record_number);
}

Napi::Value Workbook::PivotCacheRecordSetDate(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_set_pivot_cache_record_date);
}

Napi::Value Workbook::PivotCacheRecordSetText(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CacheRecordArgs record = ReadCacheRecord(info);
    const std::string utf8 = ArgString(info, 3);
    return fm_workbook_set_pivot_cache_record_text(handle, record.cache_id, record.record_index, record.field_index,
                                                   utf8.c_str());
  });
}

Napi::Value Workbook::PivotCacheRecordSetBool(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CacheRecordArgs record = ReadCacheRecord(info);
    return fm_workbook_set_pivot_cache_record_bool(handle, record.cache_id, record.record_index, record.field_index,
                                                   ArgFlag(info, 3));
  });
}

Napi::Value Workbook::PivotCacheRecordSetBlank(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_workbook_set_pivot_cache_record_blank);
}

Napi::Value Workbook::PivotCacheRecordSetError(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CacheRecordArgs record = ReadCacheRecord(info);
    const int32_t error_code = ArgI32(info, 3);
    return fm_workbook_set_pivot_cache_record_error(handle, record.cache_id, record.record_index, record.field_index,
                                                    static_cast<fm_error_code_t>(error_code));
  });
}

Napi::Value Workbook::PivotCacheGetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_pivot_cache_options_t o{};
  const fm_status_t rc = fm_workbook_get_pivot_cache_options(handle_, ArgU32(info, 0), &o);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("saveData", Napi::Boolean::New(env, o.save_data != 0));
  out.Set("missingItemsLimit", JsIndexOrNull(env, o.missing_items_limit));
  out.Set("refreshOnLoad", Napi::Boolean::New(env, o.refresh_on_load != 0));
  return out;
}

// Every member is optional; an omitted key keeps the cache's current value.
Napi::Value Workbook::PivotCacheSetOptions(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const Napi::Object spec = ArgObject(info, 1);
  fm_pivot_cache_options_t o{};
  const fm_status_t read_rc = fm_workbook_get_pivot_cache_options(handle_, cache_id, &o);
  if (read_rc != 0) {
    return ThrowStatus(env, read_rc);
  }
  CheckedSpecReader reader(env);
  o.save_data = reader.Bool(spec, "saveData", o.save_data != 0) ? 1 : 0;
  // `null` is the automatic limit (no attribute).
  const Napi::Value limit = reader.Get(spec, "missingItemsLimit");
  switch (ClassifyPatchKey(reader, limit, "missingItemsLimit", true)) {
    case PatchKey::kClear:
      o.missing_items_limit = -1;
      break;
    case PatchKey::kSet:
      o.missing_items_limit = reader.I32(limit, "missingItemsLimit", -1);
      if (reader.ok() && o.missing_items_limit < 0) {
        reader.FailRange("missingItemsLimit", "[0, 2147483647]");
      }
      break;
    case PatchKey::kKeep:
      break;
  }
  o.refresh_on_load = reader.Bool(spec, "refreshOnLoad", o.refresh_on_load != 0) ? 1 : 0;
  if (!reader.ok()) {
    return env.Undefined();
  }
  return ValueOrThrow(env, fm_workbook_set_pivot_cache_options(handle_, cache_id, &o), env.Undefined());
}

Napi::Value Workbook::PivotCacheCalculatedFieldAdd(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const std::string name = ArgString(info, 1);
    const std::string formula = ArgString(info, 2);
    return fm_workbook_add_pivot_cache_calculated_field(handle, ArgU32(info, 0), name.c_str(), formula.c_str(), out);
  });
}

Napi::Value Workbook::PivotCacheCalculatedFieldSetFormula(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string formula = ArgString(info, 2);
    return fm_workbook_set_pivot_cache_calculated_field_formula(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), formula.c_str());
  });
}

Napi::Value Workbook::PivotCacheCalculatedFieldRemove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_remove_pivot_cache_calculated_field(handle, ArgU32(info, 0),
                                                           static_cast<std::size_t>(ArgU32(info, 1)));
  });
}

Napi::Value Workbook::PivotCacheFieldFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const char* formula = nullptr;
  const fm_status_t rc = fm_workbook_get_pivot_cache_field_formula(handle_, ArgU32(info, 0),
                                                                   static_cast<std::size_t>(ArgU32(info, 1)), &formula);
  return ValueOrThrow(env, rc, JsStringOrNull(env, formula));
}

Napi::Value Workbook::PivotCacheCalculatedItemAdd(const Napi::CallbackInfo& info) {
  return IndexCall(info, [&](fm_workbook_t* handle, std::size_t* out) {
    const std::string name = ArgString(info, 2);
    const std::string formula = ArgString(info, 3);
    return fm_workbook_add_pivot_cache_calculated_item(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), name.c_str(), formula.c_str(), out);
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::size_t count = 0;
  const fm_status_t rc = fm_workbook_get_pivot_cache_calculated_item_count(handle_, ArgU32(info, 0), &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::PivotCacheCalculatedItemAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  fm_pivot_calculated_item_t item{};
  const fm_status_t rc = fm_workbook_get_pivot_cache_calculated_item(handle_, ArgU32(info, 0),
                                                                     static_cast<std::size_t>(ArgU32(info, 1)), &item);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("fieldIndex", Napi::Number::New(env, item.field_index));
  out.Set("name", JsString(env, item.name));
  out.Set("formula", JsString(env, item.formula));
  out.Set("scoped", Napi::Boolean::New(env, item.scoped != 0));
  return out;
}

Napi::Value Workbook::PivotCacheCalculatedItemSetFormula(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const std::string formula = ArgString(info, 2);
    return fm_workbook_set_pivot_cache_calculated_item_formula(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), formula.c_str());
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemMove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_move_pivot_cache_calculated_item(
        handle, ArgU32(info, 0), static_cast<std::size_t>(ArgU32(info, 1)), static_cast<std::size_t>(ArgU32(info, 2)));
  });
}

Napi::Value Workbook::PivotCacheCalculatedItemRemove(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    return fm_workbook_remove_pivot_cache_calculated_item(handle, ArgU32(info, 0),
                                                          static_cast<std::size_t>(ArgU32(info, 1)));
  });
}

}  // namespace formulon_node
