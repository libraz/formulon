// Copyright 2026 libraz. Licensed under the MIT License.
//
// PivotCache mutation bindings.
//
// Each method is a thin wrapper over `fm_workbook_pivot_cache_*`. The
// shape mirrors the embind binding: number / boolean / string args
// map straight through; the return shape is the JS Status envelope or
// an `{status, index}` pair (for entries that surface a freshly
// assigned id / index). See `c_api/formulon_c.h` "Pivot cache mutation"
// for the underlying contract.

#include <cstddef>
#include <cstdint>
#include <string>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

Napi::Value Workbook::PivotCacheCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_count(handle_, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheIdAt(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const std::size_t idx = static_cast<std::size_t>(ArgU32(info, 0));
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_id_at(handle_, idx, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), out);
}

Napi::Value Workbook::PivotCacheCreate(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t requested = ArgU32(info, 0);
  uint32_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_create(handle_, requested, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), out);
}

Napi::Value Workbook::PivotCacheRemove(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_remove(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_count(handle_, cache_id, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheFieldName(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const char* name = nullptr;
  fm_status_t rc = fm_workbook_pivot_cache_field_name(handle_, cache_id, field_idx, &name);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    out.Set("value", Napi::String::New(env, ""));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("value", Napi::String::New(env, name != nullptr ? name : ""));
  return out;
}

Napi::Value Workbook::PivotCacheFieldAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::string name = ArgString(info, 1);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_field_add(handle_, cache_id, name.c_str(), &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotCacheFieldClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_field_clear(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldSharedItemCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_field_shared_item_count(handle_, cache_id, field_idx, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const double value = ArgDouble(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_number(handle_, cache_id, field_idx, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::string utf8 = ArgString(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_text(handle_, cache_id, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const bool value = ArgBool(info, 2);
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_bool(handle_, cache_id, field_idx, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldAddSharedItemBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_cache_field_add_shared_item_blank(handle_, cache_id, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheFieldClearSharedItems(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 1));
  fm_status_t rc = fm_workbook_pivot_cache_field_clear_shared_items(handle_, cache_id, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return Napi::Number::New(env, 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t count = 0;
  if (fm_workbook_pivot_cache_record_count(handle_, cache_id, &count) != 0) {
    return Napi::Number::New(env, 0);
  }
  return Napi::Number::New(env, static_cast<double>(count));
}

Napi::Value Workbook::PivotCacheRecordAdd(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeIndexResult(env, NullHandleError(env), 0);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  std::size_t out = 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_add(handle_, cache_id, &out);
  if (rc != 0) {
    return MakeIndexResult(env, MakeErrorStatus(env, rc), 0);
  }
  return MakeIndexResult(env, MakeOkStatus(env), static_cast<uint32_t>(out));
}

Napi::Value Workbook::PivotCacheRecordClear(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  fm_status_t rc = fm_workbook_pivot_cache_record_clear(handle_, cache_id);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetNumber(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const double value = ArgDouble(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_number(handle_, cache_id, record_idx, field_idx, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const std::string utf8 = ArgString(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_text(handle_, cache_id, record_idx, field_idx, utf8.c_str());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const bool value = ArgBool(info, 3);
  fm_status_t rc = fm_workbook_pivot_cache_record_set_bool(handle_, cache_id, record_idx, field_idx, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  fm_status_t rc = fm_workbook_pivot_cache_record_set_blank(handle_, cache_id, record_idx, field_idx);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::PivotCacheRecordSetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t cache_id = ArgU32(info, 0);
  const std::size_t record_idx = static_cast<std::size_t>(ArgU32(info, 1));
  const std::size_t field_idx = static_cast<std::size_t>(ArgU32(info, 2));
  const int32_t error_code = info.Length() > 3 ? info[3].As<Napi::Number>().Int32Value() : 0;
  fm_status_t rc = fm_workbook_pivot_cache_record_set_error(handle_, cache_id, record_idx, field_idx,
                                                            static_cast<fm_error_code_t>(error_code));
  return MakeStatus(env, rc);
}

}  // namespace formulon_node
