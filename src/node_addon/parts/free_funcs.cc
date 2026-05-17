// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the module-level free functions exported alongside
// the `Workbook` class. See `free_funcs.h` for the contract.

#include "node_addon/parts/free_funcs.h"

#include <cstdint>
#include <string>

namespace formulon_node {

Napi::Value EvalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  std::string formula;
  if (info.Length() > 0) {
    formula = info[0].ToString().Utf8Value();
  }

  fm_workbook_t* wb = nullptr;
  fm_status_t rc = fm_workbook_create(&wb);
  if (rc != 0) {
    return MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
  }
  rc = fm_workbook_set_formula(wb, 0, 0, 0, formula.c_str());
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  rc = fm_workbook_recalc(wb);
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  fm_value_t v{};
  rc = fm_workbook_get_value(wb, 0, 0, 0, &v);
  if (rc != 0) {
    Napi::Object out = MakeEmptyValueResult(env, MakeErrorStatus(env, rc));
    fm_workbook_destroy(wb);
    return out;
  }
  Napi::Object out = MakeValueResult(env, MakeOkStatus(env), v);
  fm_workbook_destroy(wb);
  return out;
}

Napi::Value Version(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_version_string();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorMessage(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_message();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value LastErrorContext(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const char* s = fm_last_error_context();
  return Napi::String::New(env, s != nullptr ? s : "");
}

Napi::Value StatusString(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  int32_t code = 0;
  if (info.Length() > 0) {
    code = info[0].ToNumber().Int32Value();
  }
  const char* s = fm_status_string(static_cast<fm_status_t>(code));
  return Napi::String::New(env, s != nullptr ? s : "");
}

}  // namespace formulon_node
