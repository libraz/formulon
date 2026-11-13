// `Workbook` ObjectWrap lifecycle and class registration. Holds the
// ctor / dtor, the shared argument extraction helpers, and the
// `DefineClass` table that wires every per-area method into the JS
// class, and the module functions that share its argument guard. The
// per-area method bodies live in `parts/*.cc`.

#include "node_addon/parts/workbook_class.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstring>
#include <iterator>
#include <limits>
#include <string>

#include "node_addon/parts/free_funcs.h"

namespace formulon_node {
namespace {

// Positional arguments are checked by the registration wrapper before a
// part-TU method reads them. The masks are indexed by the JavaScript argument
// position. A typed, `object` or `required` position must be present; an
// unset bit means the position is optional and handled by the method itself.
// `tools/codegen/gen_wasm_positional_guards.py` emits the same table into the
// WASM pre-js, so both surfaces reject the same arguments with the same words.
struct PositionalArgSpec {
  std::uint64_t u32_mask;
  std::uint64_t i32_mask;
  std::uint64_t u64_mask;
  std::uint64_t double_mask;
  std::uint64_t string_mask;
  std::uint64_t optional_u32_mask;
  std::uint64_t optional_u64_mask;
  std::uint64_t optional_null_mask;
  /// A non-null object (arrays, typed arrays and functions included).
  std::uint64_t object_mask;
  /// Present, of any type: `undefined` and a missing argument are rejected.
  std::uint64_t required_mask;
};

// The generated facade scope tables, from the ledger:
//   * `kSheetMethods` -- the raw methods the `Sheet` facade calls with the sheet index first, sorted;
//   * `kHandleMethods` -- the raw methods a handle class calls, each a `HandleMethod` with the
//     `Class.method` label its errors carry and the count of leading arguments the handle supplies.
#include "node_addon/generated/facade_scope.inc"

/// One registered method or module function: its masks and the name its errors carry.
struct GuardedMethodData {
  PositionalArgSpec spec;
  const char* name;
  /// `"Workbook"` for a method; null for a module function, whose errors carry the bare name.
  const char* owner;
  /// The `Sheet` facade supplies the first argument (the sheet index).
  bool sheet_scoped;
  /// The handle class that calls the method, or null; it takes precedence over `sheet_scoped`.
  const HandleMethod* handle;
};

// A uint64 argument arrives as a Number, which is exact only up to 2^53 - 1;
// a larger value could not name the cursor the caller holds.
// @limit: format
constexpr double kMaxSafeInteger = 9007199254740991.0;

bool IsSheetMethod(const char* name) {
  return std::binary_search(std::begin(kSheetMethods), std::end(kSheetMethods), name,
                            [](const char* a, const char* b) { return std::strcmp(a, b) < 0; });
}

const HandleMethod* FindHandleMethod(const char* name) {
  const auto found = std::find_if(std::begin(kHandleMethods), std::end(kHandleMethods),
                                  [name](const HandleMethod& entry) { return std::strcmp(entry.raw, name) == 0; });
  return found != std::end(kHandleMethods) ? &*found : nullptr;
}

// The public `Class.method` name of a registered method, for reader failures.
std::string PublicMethodLabel(const void* data) {
  const auto& method = *static_cast<const GuardedMethodData*>(data);
  if (method.handle != nullptr) {
    return method.handle->label;
  }
  if (method.sheet_scoped) {
    return std::string("Sheet.") + method.name;
  }
  return std::string(method.owner == nullptr ? "Formulon" : method.owner) + "." + method.name;
}

// One past the highest set bit of `mask`.
std::size_t MaskWidth(std::uint64_t mask) {
  std::size_t width = 0;
  while (mask != 0) {
    mask >>= 1U;
    ++width;
  }
  return width;
}

// The d.ts name of raw argument `raw_index` of a registered method, for its positional refusals. The arguments the
// facade supplies have none; the generated table leaves out a call whose raw layout is not its d.ts signature.
std::string PublicArgumentName(const void* data, std::size_t raw_index) {
  const auto& method = *static_cast<const GuardedMethodData*>(data);
  const std::size_t supplied =
      method.handle != nullptr ? static_cast<std::size_t>(method.handle->supplied) : (method.sheet_scoped ? 1U : 0U);
  if (raw_index < supplied) {
    return {};
  }
  const std::string label = PublicMethodLabel(data);
  const std::size_t index = raw_index - supplied;
  std::string_view name = formulon::binding::param_name(label, index);
  // The `ooxml` namespaces share the raw methods of their owner, so their rows are keyed by the namespace class
  // (the facade's KEY_FALLBACKS).
  if (name.empty() && label.rfind("Sheet.", 0) == 0) {
    name = formulon::binding::param_name("SheetOoxml." + label.substr(6), index);
  } else if (name.empty() && label.rfind("Workbook.", 0) == 0) {
    name = formulon::binding::param_name("WorkbookOoxml." + label.substr(9), index);
  }
  return std::string(name);
}

// The positional refusals below word their message for the call `ScopedReaderLabel` names.
bool RejectArgumentType(const Napi::CallbackInfo& info, std::size_t idx, const char* expected) {
  ThrowPositionalTypeError(info.Env(), idx, std::string("must be ") + expected);
  return false;
}

bool RejectArgumentRequired(const Napi::CallbackInfo& info, std::size_t idx) {
  ThrowPositionalTypeError(info.Env(), idx, "is required");
  return false;
}

bool RejectArgumentInteger(const Napi::CallbackInfo& info, std::size_t idx, double number, double min_value) {
  ThrowPositionalIntegerRefusal(info.Env(), idx, number, min_value);
  return false;
}

bool ValidatePositionalArgs(const Napi::CallbackInfo& info, const GuardedMethodData& method) {
  const PositionalArgSpec& spec = method.spec;
  // A number that is an integer in [min_value, max_value].
  const auto validate_integer = [&](std::size_t idx, double min_value, double max_value) {
    const Napi::Value value = info[idx];
    if (!value.IsNumber()) {
      return RejectArgumentType(info, idx, "a number");
    }
    const double number = NumberValue(value, 0.0);
    if (!std::isfinite(number) || std::trunc(number) != number || number < min_value || number > max_value) {
      return RejectArgumentInteger(info, idx, number, min_value);
    }
    return true;
  };
  // An optional position may be absent or `undefined`; `null` only when its bit is in `optional_null_mask`.
  const auto validate_optional_integer = [&](std::size_t idx, std::uint64_t bit, double max_value) {
    if (idx >= info.Length()) {
      return true;
    }
    const Napi::Value value = info[idx];
    if (value.IsNull() && (spec.optional_null_mask & bit) == 0) {
      return RejectArgumentType(info, idx, "a number");
    }
    return value.IsUndefined() || value.IsNull() || validate_integer(idx, 0.0, max_value);
  };

  // Required positions are checked even past the supplied arguments, where
  // `info[idx]` reads `undefined`.
  const std::size_t required_width = MaskWidth(spec.u32_mask | spec.i32_mask | spec.u64_mask | spec.double_mask |
                                               spec.string_mask | spec.object_mask | spec.required_mask);
  const std::size_t supplied = info.Length() < 64 ? info.Length() : 64;
  const std::size_t positional_count = supplied > required_width ? supplied : required_width;
  for (std::size_t idx = 0; idx < positional_count; ++idx) {
    const std::uint64_t bit = std::uint64_t{1} << idx;
    if ((spec.object_mask & bit) != 0 && !info[idx].IsObject()) {
      return RejectArgumentType(info, idx, "an object");
    }
    if ((spec.required_mask & bit) != 0 && info[idx].IsUndefined()) {
      return RejectArgumentRequired(info, idx);
    }
    if ((spec.u32_mask & bit) != 0 && !validate_integer(idx, 0.0, kUint32MaxNumber)) {
      return false;
    }
    if ((spec.i32_mask & bit) != 0 &&
        !validate_integer(idx, static_cast<double>(std::numeric_limits<std::int32_t>::min()),
                          static_cast<double>(std::numeric_limits<std::int32_t>::max()))) {
      return false;
    }
    if ((spec.u64_mask & bit) != 0 && !validate_integer(idx, 0.0, kMaxSafeInteger)) {
      return false;
    }
    if ((spec.double_mask & bit) != 0 && !info[idx].IsNumber()) {
      return RejectArgumentType(info, idx, "a number");
    }
    if ((spec.string_mask & bit) != 0 && !info[idx].IsString()) {
      return RejectArgumentType(info, idx, "a string");
    }
    if ((spec.optional_u32_mask & bit) != 0 && !validate_optional_integer(idx, bit, kUint32MaxNumber)) {
      return false;
    }
    if ((spec.optional_u64_mask & bit) != 0 && !validate_optional_integer(idx, bit, kMaxSafeInteger)) {
      return false;
    }
  }
  return true;
}

using InstanceMethodCallback = Napi::Value (Workbook::*)(const Napi::CallbackInfo&);

template <InstanceMethodCallback method>
napi_value GuardedInstanceMethodCallback(napi_env raw_env, napi_callback_info raw_info) {
  return Napi::details::WrapCallback(raw_env, [&] {
    Napi::CallbackInfo info(raw_env, raw_info);
    const auto* const data = static_cast<const GuardedMethodData*>(info.Data());
    if (data == nullptr) {
      return info.Env().Undefined();
    }
    const ScopedReaderLabel label(&PublicMethodLabel, &PublicArgumentName, data);
    if (!ValidatePositionalArgs(info, *data)) {
      return info.Env().Undefined();
    }
    Workbook* const instance = Workbook::Unwrap(info.This().As<Napi::Object>());
    return instance == nullptr ? Napi::Value() : instance->RunActive(method, info);
  });
}

// Each method is registered once, so the per-instantiation static holds its
// one name for every environment.
template <InstanceMethodCallback method, std::uint64_t u32_mask = 0, std::uint64_t i32_mask = 0,
          std::uint64_t u64_mask = 0, std::uint64_t double_mask = 0, std::uint64_t string_mask = 0,
          std::uint64_t optional_u32_mask = 0, std::uint64_t optional_u64_mask = 0,
          std::uint64_t optional_null_mask = 0, std::uint64_t object_mask = 0, std::uint64_t required_mask = 0>
Napi::ClassPropertyDescriptor<Workbook> GuardedInstanceMethod(const char* name) {
  static const GuardedMethodData data{{u32_mask, i32_mask, u64_mask, double_mask, string_mask, optional_u32_mask,
                                       optional_u64_mask, optional_null_mask, object_mask, required_mask},
                                      name,
                                      "Workbook",
                                      IsSheetMethod(name),
                                      FindHandleMethod(name)};
  napi_property_descriptor descriptor{};
  descriptor.utf8name = name;
  descriptor.method = &GuardedInstanceMethodCallback<method>;
  descriptor.data = const_cast<void*>(static_cast<const void*>(&data));
  descriptor.attributes = napi_default;
  return descriptor;
}

using ModuleFunctionCallback = Napi::Value (*)(const Napi::CallbackInfo&);

// The module-function counterpart of `GuardedInstanceMethod`, with the same
// mask order. A module function holds no workbook, so no active-call count
// is kept.
template <ModuleFunctionCallback function, std::uint64_t u32_mask = 0, std::uint64_t i32_mask = 0,
          std::uint64_t u64_mask = 0, std::uint64_t double_mask = 0, std::uint64_t string_mask = 0,
          std::uint64_t optional_u32_mask = 0, std::uint64_t optional_u64_mask = 0,
          std::uint64_t optional_null_mask = 0, std::uint64_t object_mask = 0, std::uint64_t required_mask = 0>
void SetGuardedModuleFunction(Napi::Env env, Napi::Object exports, const char* name) {
  static const GuardedMethodData data{{u32_mask, i32_mask, u64_mask, double_mask, string_mask, optional_u32_mask,
                                       optional_u64_mask, optional_null_mask, object_mask, required_mask},
                                      name,
                                      nullptr,
                                      false,
                                      nullptr};
  exports.Set(name, Napi::Function::New(
                        env,
                        [](const Napi::CallbackInfo& info) -> Napi::Value {
                          const ScopedReaderLabel label(&PublicMethodLabel, &PublicArgumentName, &data);
                          if (!ValidatePositionalArgs(info, data)) {
                            return info.Env().Undefined();
                          }
                          return function(info);
                        },
                        name));
}

}  // namespace

Workbook::Workbook(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Workbook>(info) {
  // Default constructor used by the static factories. They populate
  // `handle_` after construction via `wb->handle_ = ...`.
  //
  // External JS callers should NOT invoke `new Workbook()` directly;
  // the JS-side index.mjs only re-exports the static factories.
  (void)info;
}

Workbook::~Workbook() {
  DestroyHandle(Env());
}

Napi::Symbol Workbook::IterativeProgressKey(Napi::Env env) {
  return Napi::Symbol::For(env, "formulon.iterativeProgress");
}

void Workbook::DestroyHandle(Napi::Env env) {
  if (handle_ != nullptr) {
    // Clear this workbook's C callback before destroying its user-data
    // wrapper.
    (void)fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
  // With the handle gone the measurement is zero, so this hands the
  // whole reported amount back to V8. Doing it after the destroy (rather
  // than before) keeps the accounting from ever going negative if the
  // release path is entered twice.
  SyncExternalMemory(env);
}

void Workbook::SyncExternalMemory(Napi::Env env) {
  size_t bytes = 0;
  if (handle_ != nullptr) {
    (void)fm_workbook_get_memory_usage(handle_, &bytes);
  }
  ApplyExternalMemoryEstimate(env, bytes);
}

void Workbook::ApplyExternalMemoryEstimate(Napi::Env env, size_t bytes) {
  const int64_t current = static_cast<int64_t>(bytes);
  const int64_t delta = current - reported_external_bytes_;
  if (delta == 0) {
    return;
  }
  reported_external_bytes_ = current;
  Napi::MemoryManagement::AdjustExternalMemory(env, delta);
}

uint32_t Workbook::ArgU32(const Napi::CallbackInfo& info, size_t idx) {
  const double number = NumberArg(info, idx, 0.0);
  return number >= 0.0 && number <= kUint32MaxNumber ? static_cast<uint32_t>(number) : 0U;
}

int32_t Workbook::ArgI32(const Napi::CallbackInfo& info, size_t idx) {
  const double number = NumberArg(info, idx, 0.0);
  return number >= static_cast<double>(std::numeric_limits<int32_t>::min()) &&
                 number <= static_cast<double>(std::numeric_limits<int32_t>::max())
             ? static_cast<int32_t>(number)
             : 0;
}

double Workbook::ArgDouble(const Napi::CallbackInfo& info, size_t idx) {
  return NumberArg(info, idx, 0.0);
}

std::string Workbook::ArgString(const Napi::CallbackInfo& info, size_t idx) {
  return StringArg(info, idx);
}

bool Workbook::ArgBool(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length() || info[idx].IsUndefined()) {
    return false;
  }
  if (!info[idx].IsBoolean()) {
    ThrowPositionalTypeError(info.Env(), idx, "must be a boolean");
    return false;
  }
  return info[idx].As<Napi::Boolean>().Value();
}

Napi::Object Workbook::ArgObject(const Napi::CallbackInfo& info, size_t idx) {
  Napi::Object object;
  return ObjectArg(info, idx, &object) ? object : Napi::Object::New(info.Env());
}

Napi::Value Workbook::InvokeRowColEdit(const Napi::CallbackInfo& info, RowColEditFn fn) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const std::size_t sheet = ArgU32(info, 0);
  const uint32_t first = ArgU32(info, 1);
  const uint32_t last = ArgU32(info, 2);
  return ValueOrThrow(env, fn(handle_, sheet, first, last), env.Undefined());
}

// ---- Static factories -----------------------------------------------

Napi::Value Workbook::FactoryResult(Napi::Env env, Napi::Object jsobj, Workbook* wb, fm_status_t rc) {
  if (rc != 0) {
    wb->handle_ = nullptr;
    return ThrowStatus(env, rc);
  }
  wb->SyncExternalMemory(env);
  return jsobj;
}

Napi::Value Workbook::CreateWith(const Napi::CallbackInfo& info, CreateFn create) {
  Napi::Env env = info.Env();
  Napi::Object jsobj = GetClass(env).New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  return FactoryResult(env, jsobj, wb, create(&wb->handle_));
}

Napi::Value Workbook::CreateDefault(const Napi::CallbackInfo& info) {
  return CreateWith(info, &fm_workbook_create_default);
}

Napi::Value Workbook::CreateEmpty(const Napi::CallbackInfo& info) {
  return CreateWith(info, &fm_workbook_create_empty);
}

Napi::Value Workbook::CreateForProfile(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (!info[0].IsString()) {
    Napi::TypeError::New(env, formulon::binding::argument_message(
                                  "Workbook.createForProfile",
                                  formulon::binding::param_name("Workbook.createForProfile", 0), "must be a string"))
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object jsobj = GetClass(env).New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  const std::string profile_id = ArgString(info, 0);
  return FactoryResult(env, jsobj, wb, fm_workbook_create_for_profile(profile_id.c_str(), &wb->handle_));
}

Napi::Value Workbook::LoadBytes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (!info[0].IsObject()) {
    Napi::TypeError::New(env, "Workbook.load bytes must be an object").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Uint8Array u8;
  if (!BytesArg(info, 0, &u8)) {
    Napi::TypeError::New(env, "Workbook.load bytes must be a Uint8Array").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  Napi::Object jsobj = GetClass(env).New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  // An empty array may have no backing store; a non-NULL pointer lets the
  // C ABI report the empty input rather than a NULL argument.
  static const uint8_t kEmptyInput = 0;
  const std::size_t len = u8.ElementLength();
  const uint8_t* data = len == 0 ? &kEmptyInput : u8.Data();
  return FactoryResult(env, jsobj, wb, fm_workbook_load(data, len, &wb->handle_));
}

// ---- Class registration ---------------------------------------------

Napi::Function Workbook::GetClass(Napi::Env env) {
  AddonData& data = GetAddonData(env);
  if (!data.workbook_class.IsEmpty()) {
    return data.workbook_class.Value();
  }

  // clang-format off
  Napi::Function constructor = DefineClass(
      env, "Workbook",
      {
          StaticMethod<&Workbook::CreateDefault>("createDefault"),
          StaticMethod<&Workbook::CreateEmpty>("createEmpty"),
          StaticMethod<&Workbook::CreateForProfile>("createForProfile"),
          StaticMethod<&Workbook::LoadBytes>("load"),
          GuardedInstanceMethod<&Workbook::AddBorder, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addBorder"),
          GuardedInstanceMethod<&Workbook::AddConditionalFormat, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addCfRule"),
          GuardedInstanceMethod<&Workbook::AddDxf, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addDxf"),
          GuardedInstanceMethod<&Workbook::AddFill, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addFill"),
          GuardedInstanceMethod<&Workbook::AddFont, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addFont"),
          GuardedInstanceMethod<&Workbook::SetHyperlink, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setHyperlink"),
          GuardedInstanceMethod<&Workbook::AddMerge, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addMerge"),
          GuardedInstanceMethod<&Workbook::AddNumFmt, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addNumberFormat"),
          GuardedInstanceMethod<&Workbook::AddSheet, 0x2ULL>("addSheet"),
          GuardedInstanceMethod<&Workbook::AddValidation, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addValidation"),
          GuardedInstanceMethod<&Workbook::AddPerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addPerson"),
          GuardedInstanceMethod<&Workbook::AddThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addThreadedComment"),
          GuardedInstanceMethod<&Workbook::ApplyAutoFilter, 0x1ULL>("applyAutoFilter"),
          GuardedInstanceMethod<&Workbook::ApplyTableAutoFilter, 0x3ULL>("applyTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::ClearAutoFilter, 0x1ULL>("clearAutoFilterCriteria"),
          GuardedInstanceMethod<&Workbook::ClearTableAutoFilter, 0x3ULL>("clearTableAutoFilterCriteria"),
          GuardedInstanceMethod<&Workbook::EditThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateThreadedComment"),
          GuardedInstanceMethod<&Workbook::EvaluateAutoFilter, 0x1ULL>("evaluateAutoFilter"),
          GuardedInstanceMethod<&Workbook::EvaluateTableAutoFilter, 0x3ULL>("evaluateTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetAutoFilter, 0x1ULL>("getAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetAutoFilterXml, 0x1ULL>("getAutoFilterXml"),
          GuardedInstanceMethod<&Workbook::GetPersons>("listPeople"),
          GuardedInstanceMethod<&Workbook::GetTableAutoFilter, 0x3ULL>("getTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetThreadedComments, 0x1ULL>("listThreadedComments"),
          GuardedInstanceMethod<&Workbook::ListInvalidCells, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x2ULL, 0x6ULL>("getInvalidCells"),
          GuardedInstanceMethod<&Workbook::SetAutoFilterXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setAutoFilterXml"),
          GuardedInstanceMethod<&Workbook::RemovePerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removePerson"),
          GuardedInstanceMethod<&Workbook::GetPerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getPerson"),
          GuardedInstanceMethod<&Workbook::RemoveThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeThreadedComment"),
          GuardedInstanceMethod<&Workbook::GetThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getThreadedComment"),
          GuardedInstanceMethod<&Workbook::SetAutoFilter, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetTableAutoFilter, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::ValidateValue, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("validateValue"),
          GuardedInstanceMethod<&Workbook::ListDrawingObjects, 0x1ULL>("listDrawingObjects"),
          GuardedInstanceMethod<&Workbook::GetDrawingObject, 0x3ULL>("getDrawingObject"),
          GuardedInstanceMethod<&Workbook::GetImage, 0x3ULL>("getPicture"),
          GuardedInstanceMethod<&Workbook::InsertImage, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addImage"),
          GuardedInstanceMethod<&Workbook::RemoveDrawing, 0x3ULL>("removeDrawingObject"),
          GuardedInstanceMethod<&Workbook::UpdateDrawing, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateDrawing"),
          GuardedInstanceMethod<&Workbook::SetDrawingZOrder, 0x7ULL>("setDrawingZOrder"),
          GuardedInstanceMethod<&Workbook::SnapshotDrawing, 0x3ULL>("snapshotDrawing"),
          GuardedInstanceMethod<&Workbook::RestoreDrawing, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("restoreDrawing"),
          GuardedInstanceMethod<&Workbook::InsertShape, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addShape"),
          GuardedInstanceMethod<&Workbook::UpdateShape, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateShape"),
          GuardedInstanceMethod<&Workbook::GetShape, 0x3ULL>("getShape"),
          GuardedInstanceMethod<&Workbook::GetSlicer, 0x3ULL>("getSlicer"),
          GuardedInstanceMethod<&Workbook::AddSlicer, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("addSlicer"),
          GuardedInstanceMethod<&Workbook::UpdateSlicer, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateSlicer"),
          GuardedInstanceMethod<&Workbook::AddDataTable, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addDataTable"),
          GuardedInstanceMethod<&Workbook::GetDataTable, 0x7ULL>("getDataTable"),
          GuardedInstanceMethod<&Workbook::RemoveDataTable, 0x7ULL>("removeDataTable"),
          GuardedInstanceMethod<&Workbook::RecalcDataTables>("recalcDataTables"),
          GuardedInstanceMethod<&Workbook::GetSparklineGroupCount, 0x1ULL>("getSparklineGroupCount"),
          GuardedInstanceMethod<&Workbook::GetSparklineGroup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getSparklineGroup"),
          GuardedInstanceMethod<&Workbook::AddSparklineGroup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addSparklineGroup"),
          GuardedInstanceMethod<&Workbook::UpdateSparklineGroup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateSparklineGroup"),
          GuardedInstanceMethod<&Workbook::RemoveSparklineGroup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeSparklineGroup"),
          GuardedInstanceMethod<&Workbook::GetChart, 0x3ULL>("getChart"),
          GuardedInstanceMethod<&Workbook::GetChartData, 0x3ULL>("getChartData"),
          GuardedInstanceMethod<&Workbook::PreviewChartData, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("getChartPreview"),
          GuardedInstanceMethod<&Workbook::InsertChart, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addChart"),
          GuardedInstanceMethod<&Workbook::SetChartType, 0x7ULL, 0x8ULL>("setChartType"),
          GuardedInstanceMethod<&Workbook::SetChartTitle, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setChartTitle"),
          GuardedInstanceMethod<&Workbook::SetChartLegend, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setChartLegend"),
          GuardedInstanceMethod<&Workbook::UpdateChartLegend, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateChartLegend"),
          GuardedInstanceMethod<&Workbook::SetChartAxis, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("updateChartAxis"),
          GuardedInstanceMethod<&Workbook::SetChartDisplay, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateChartDisplay"),
          GuardedInstanceMethod<&Workbook::SetChartSeries, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("updateChartSeries"),
          GuardedInstanceMethod<&Workbook::SetChartPoint, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("setChartPoint"),
          GuardedInstanceMethod<&Workbook::UpdateChartPoint, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("updateChartPoint"),
          GuardedInstanceMethod<&Workbook::SetChartPlot, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("updateChartPlot"),
          GuardedInstanceMethod<&Workbook::SetChartFilter, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xCULL>("setChartFilter"),
          GuardedInstanceMethod<&Workbook::AddChartSeries, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("addChartSeries"),
          GuardedInstanceMethod<&Workbook::RemoveChartSeries, 0xFULL>("removeChartSeries"),
          GuardedInstanceMethod<&Workbook::SetChartSource, 0x3ULL, 0x8ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setChartSource"),
          GuardedInstanceMethod<&Workbook::AddXf, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addCellXf"),
          GuardedInstanceMethod<&Workbook::AddCellStyleXf, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addCellStyleXf"),
          GuardedInstanceMethod<&Workbook::BorderCount>("getBorderCount"),
          GuardedInstanceMethod<&Workbook::CalcMode>("getCalcMode"),
          GuardedInstanceMethod<&Workbook::FullPrecision>("getFullPrecision"),
          GuardedInstanceMethod<&Workbook::CalcOnSave>("getCalcOnSave"),
          GuardedInstanceMethod<&Workbook::RefMode>("getRefMode"),
          GuardedInstanceMethod<&Workbook::Date1904>("getDate1904"),
          GuardedInstanceMethod<&Workbook::ClearRowHeight, 0x7ULL>("resetRowHeight"),
          GuardedInstanceMethod<&Workbook::ClearColumnWidth, 0x7ULL>("resetColumnWidth"),
          GuardedInstanceMethod<&Workbook::ClearColumnStyle, 0x7ULL>("resetColumnXfIndex"),
          GuardedInstanceMethod<&Workbook::ClearRowStyle, 0x7ULL>("resetRowXfIndex"),
          GuardedInstanceMethod<&Workbook::ColumnCharsToPt, 0x1ULL, 0x2ULL, 0x0ULL, 0x4ULL>("convertColumnCharsToPt"),
          GuardedInstanceMethod<&Workbook::ColumnPtToChars, 0x1ULL, 0x2ULL, 0x0ULL, 0x4ULL>("convertColumnPtToChars"),
          GuardedInstanceMethod<&Workbook::FormatValue, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("formatValue"),
          GuardedInstanceMethod<&Workbook::GetCellRectPt, 0x1ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getCellRectPt"),
          GuardedInstanceMethod<&Workbook::GetCellsInRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL, 0x4ULL, 0xCULL, 0x2ULL>("getCellsInRange"),
          GuardedInstanceMethod<&Workbook::GetColumnWidthPt, 0x3ULL, 0x4ULL>("getColumnWidthPt"),
          GuardedInstanceMethod<&Workbook::GetDisplayText, 0x7ULL>("getDisplayText"),
          GuardedInstanceMethod<&Workbook::GetFormula, 0x7ULL>("getFormula"),
          GuardedInstanceMethod<&Workbook::GetFormulaR1C1, 0x7ULL>("getFormulaR1C1"),
          GuardedInstanceMethod<&Workbook::GetMergesInRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getMergesInRange"),
          GuardedInstanceMethod<&Workbook::GetRowHeightPt, 0x3ULL>("getRowHeightPt"),
          GuardedInstanceMethod<&Workbook::GetSheetFormatDefaults, 0x1ULL>("getFormatDefaults"),
          GuardedInstanceMethod<&Workbook::GetWidthModel, 0x1ULL, 0x2ULL>("getWidthModel"),
          GuardedInstanceMethod<&Workbook::PinnedNow>("getPinnedNow"),
          GuardedInstanceMethod<&Workbook::ListCellStyles>("listCellStyles"),
          GuardedInstanceMethod<&Workbook::CellStyleXfCount>("getCellStyleXfCount"),
          GuardedInstanceMethod<&Workbook::ClearConditionalFormats, 0x1ULL>("clearCfRules"),
          GuardedInstanceMethod<&Workbook::ClearHyperlinks, 0x1ULL>("clearHyperlinks"),
          GuardedInstanceMethod<&Workbook::ClearMerges, 0x1ULL>("clearMerges"),
          GuardedInstanceMethod<&Workbook::ClearValidations, 0x1ULL>("clearValidations"),
          GuardedInstanceMethod<&Workbook::ListDefinedNames>("listDefinedNames"),
          GuardedInstanceMethod<&Workbook::GetDefinedName, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getDefinedName"),
          GuardedInstanceMethod<&Workbook::RemoveDefinedName, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeDefinedName"),
          GuardedInstanceMethod<&Workbook::DeleteCols, 0x7ULL>("deleteColumns"),
          GuardedInstanceMethod<&Workbook::DeleteRows, 0x7ULL>("deleteRows"),
          GuardedInstanceMethod<&Workbook::DeleteCells, 0x1ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("deleteCells"),
          GuardedInstanceMethod<&Workbook::Dependents, 0xFULL>("getDependents"),
          GuardedInstanceMethod<&Workbook::Dispose>("dispose"),
          InstanceAccessor<&Workbook::Disposed>("disposed"),
          GuardedInstanceMethod<&Workbook::DxfCount>("getDxfCount"),
          GuardedInstanceMethod<&Workbook::EvaluateCfRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("evaluateCfRange"),
          GuardedInstanceMethod<&Workbook::EvaluateConditionalFormula, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("evaluateCfFormula"),
          GuardedInstanceMethod<&Workbook::EvaluateFormulaArray, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("evaluateFormulaArray"),
          GuardedInstanceMethod<&Workbook::ProfileId>("getProfileId"),
          GuardedInstanceMethod<&Workbook::GetLocation>("getLocation"),
          GuardedInstanceMethod<&Workbook::FillCount>("getFillCount"),
          GuardedInstanceMethod<&Workbook::FontCount>("getFontCount"),
          GuardedInstanceMethod<&Workbook::GetBorder, 0x1ULL>("getBorder"),
          GuardedInstanceMethod<&Workbook::GetCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getCellStyle"),
          GuardedInstanceMethod<&Workbook::GetCellStyleXf, 0x1ULL>("getCellStyleXf"),
          GuardedInstanceMethod<&Workbook::GetEffectiveStyle, 0x7ULL>("getEffectiveStyle"),
          GuardedInstanceMethod<&Workbook::GetThemeColors>("getThemeColors"),
          GuardedInstanceMethod<&Workbook::GetThemeFonts>("getThemeFonts"),
          GuardedInstanceMethod<&Workbook::RemoveCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeCellStyle"),
          GuardedInstanceMethod<&Workbook::ResolveColor, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("resolveColor"),
          GuardedInstanceMethod<&Workbook::SetCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setCellStyle"),
          GuardedInstanceMethod<&Workbook::ListTableStyles>("listTableStyles"),
          GuardedInstanceMethod<&Workbook::GetTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getTableStyle"),
          GuardedInstanceMethod<&Workbook::SetTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setTableStyle"),
          GuardedInstanceMethod<&Workbook::RemoveTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeTableStyle"),
          GuardedInstanceMethod<&Workbook::GetTableStyleDefaults>("getTableStyleDefaults"),
          GuardedInstanceMethod<&Workbook::SetTableStyleDefaults, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("updateTableStyleDefaults"),
          GuardedInstanceMethod<&Workbook::SetThemeColors, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setThemeColors"),
          GuardedInstanceMethod<&Workbook::SetThemeFonts, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setThemeFonts"),
          GuardedInstanceMethod<&Workbook::SetThemeScriptFont, 0x0ULL, 0x1ULL, 0x0ULL, 0x0ULL, 0x6ULL>("setThemeScriptFont"),
          GuardedInstanceMethod<&Workbook::GetThemeScriptFont, 0x0ULL, 0x1ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getThemeScriptFont"),
          GuardedInstanceMethod<&Workbook::ResetTheme>("resetTheme"),
          GuardedInstanceMethod<&Workbook::GetCellXf, 0x1ULL>("getCellXf"),
          GuardedInstanceMethod<&Workbook::GetCellXfIndex, 0x7ULL>("getCellXfIndex"),
          GuardedInstanceMethod<&Workbook::GetCellPhonetic, 0x7ULL>("getCellPhonetic"),
          GuardedInstanceMethod<&Workbook::GetCellPhoneticRunCount, 0x7ULL>("getCellPhoneticRunCount"),
          GuardedInstanceMethod<&Workbook::GetCellPhoneticRun, 0xFULL>("getCellPhoneticRun"),
          GuardedInstanceMethod<&Workbook::GetCellPhoneticProperties, 0x7ULL>("getCellPhoneticProperties"),
          GuardedInstanceMethod<&Workbook::GetCellShowPhonetic, 0x7ULL>("getCellShowPhonetic"),
          GuardedInstanceMethod<&Workbook::GetComment, 0x7ULL>("getComment"),
          GuardedInstanceMethod<&Workbook::GetComments, 0x1ULL>("listComments"),
          GuardedInstanceMethod<&Workbook::GetCommentVisible, 0x7ULL>("getCommentVisible"),
          GuardedInstanceMethod<&Workbook::GetWorkbookProtection>("getWorkbookProtection"),
          GuardedInstanceMethod<&Workbook::GetProtectedRangeCount, 0x1ULL>("getProtectedRangeCount"),
          GuardedInstanceMethod<&Workbook::GetProtectedRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getProtectedRange"),
          GuardedInstanceMethod<&Workbook::GetCfRuleCount, 0x1ULL>("getCfRuleCount"),
          GuardedInstanceMethod<&Workbook::GetCfRule, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getCfRule"),
          GuardedInstanceMethod<&Workbook::GetDxf, 0x1ULL>("getDxf"),
          GuardedInstanceMethod<&Workbook::GetExternalLinkCount>("getExternalLinkCount"),
          GuardedInstanceMethod<&Workbook::GetExternalLink, 0x1ULL>("getExternalLink"),
          GuardedInstanceMethod<&Workbook::GetFill, 0x1ULL>("getFill"),
          GuardedInstanceMethod<&Workbook::GetFont, 0x1ULL>("getFont"),
          GuardedInstanceMethod<&Workbook::GetHyperlinks, 0x1ULL>("listHyperlinks"),
          GuardedInstanceMethod<&Workbook::GetHyperlink, 0x7ULL>("getHyperlink"),
          GuardedInstanceMethod<&Workbook::GetIterative>("getIterative"),
          GuardedInstanceMethod<&Workbook::GetLambdaText, 0x7ULL>("getLambdaText"),
          GuardedInstanceMethod<&Workbook::GetMerges, 0x1ULL>("listMerges"),
          GuardedInstanceMethod<&Workbook::GetNumFmt, 0x1ULL>("getNumberFormat"),
          GuardedInstanceMethod<&Workbook::ListNumberFormats>("listNumberFormats"),
          GuardedInstanceMethod<&Workbook::GetSheetColumns, 0x1ULL>("listColumnOverrides"),
          GuardedInstanceMethod<&Workbook::GetSheetProtection, 0x1ULL>("getSheetProtection"),
          GuardedInstanceMethod<&Workbook::GetSheetRowOverrides, 0x1ULL>("listRowOverrides"),
          GuardedInstanceMethod<&Workbook::GetSheetSelection, 0x1ULL>("getSelection"),
          GuardedInstanceMethod<&Workbook::GetSheetView, 0x1ULL>("getView"),
          GuardedInstanceMethod<&Workbook::GetVisibility, 0x1ULL>("getVisibility"),
          GuardedInstanceMethod<&Workbook::SetVisibility, 0x1ULL, 0x2ULL>("setVisibility"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordGetValue, 0x7ULL>("getPivotCacheRecordValue"),
          GuardedInstanceMethod<&Workbook::ListThemeScriptFonts, 0x0ULL, 0x1ULL>("listThemeScriptFonts"),
          GuardedInstanceMethod<&Workbook::GetValidationCount, 0x1ULL>("getValidationCount"),
          GuardedInstanceMethod<&Workbook::GetValidation, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getValidation"),
          GuardedInstanceMethod<&Workbook::GetValue, 0x7ULL>("getValue"),
          GuardedInstanceMethod<&Workbook::InsertCols, 0x7ULL>("insertColumns"),
          GuardedInstanceMethod<&Workbook::InsertRows, 0x7ULL>("insertRows"),
          GuardedInstanceMethod<&Workbook::InsertCells, 0x1ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("insertCells"),
          GuardedInstanceMethod<&Workbook::MemoryUsage>("getMemoryUsage"),
          GuardedInstanceMethod<&Workbook::MoveSheet, 0x3ULL>("moveSheet"),
          GuardedInstanceMethod<&Workbook::CopySheet, 0x3ULL>("copySheet"),
          GuardedInstanceMethod<&Workbook::PartialRecalc, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("recalcViewport"),
          GuardedInstanceMethod<&Workbook::Paginate, 0x1ULL>("getPagination"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetupXml, 0x1ULL>("getPageSetupXml"),
          GuardedInstanceMethod<&Workbook::UpdateSheetFormatDefaults, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updateFormatDefaults"),
          GuardedInstanceMethod<&Workbook::SetSheetPageSetupXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageSetupXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMarginsXml, 0x1ULL>("getPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPageMarginsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintOptionsXml, 0x1ULL>("getPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintOptionsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetHeaderFooterXml, 0x1ULL>("getHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::SetSheetHeaderFooterXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::GetSheetSheetPrXml, 0x1ULL>("getSheetPrXml"),
          GuardedInstanceMethod<&Workbook::SetSheetSheetPrXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPrXml"),
          GuardedInstanceMethod<&Workbook::GetSheetTabColor, 0x1ULL>("getTabColor"),
          GuardedInstanceMethod<&Workbook::SetSheetTabColor, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setTabColor"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintArea, 0x1ULL>("getPrintArea"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintArea, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintArea"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintTitles, 0x1ULL>("getPrintTitles"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintTitles, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintTitles"),
          GuardedInstanceMethod<&Workbook::SetRowBreak, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setRowBreak"),
          GuardedInstanceMethod<&Workbook::SetColumnBreak, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setColumnBreak"),
          GuardedInstanceMethod<&Workbook::ClearSheetBreaks, 0x1ULL>("clearBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetRowBreaks, 0x1ULL>("listRowBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetColBreaks, 0x1ULL>("listColumnBreaks"),
          GuardedInstanceMethod<&Workbook::UpdateSheetPageSetup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updatePageSetup"),
          GuardedInstanceMethod<&Workbook::UpdateSheetPageMargins, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updatePageMargins"),
          GuardedInstanceMethod<&Workbook::UpdateSheetPrintOptions, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updatePrintOptions"),
          GuardedInstanceMethod<&Workbook::UpdateSheetHeaderFooter, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updateHeaderFooter"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetup, 0x1ULL>("getPageSetup"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMargins, 0x1ULL>("getPageMargins"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintOptions, 0x1ULL>("getPrintOptions"),
          GuardedInstanceMethod<&Workbook::GetSheetHeaderFooter, 0x1ULL>("getHeaderFooter"),
          GuardedInstanceMethod<&Workbook::ListPassthroughs>("listPassthroughs"),
          GuardedInstanceMethod<&Workbook::HasPassthrough, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("hasPassthrough"),
          GuardedInstanceMethod<&Workbook::ListPivotCaches>("listPivotCaches"),
          GuardedInstanceMethod<&Workbook::PivotCacheCreate, 0x1ULL>("addPivotCache"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAdd, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addPivotCacheField"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBlank, 0x3ULL>("addPivotCacheFieldSharedItemBlank"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBool, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotCacheFieldSharedItemBool"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemError, 0x3ULL, 0x4ULL>("addPivotCacheFieldSharedItemError"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemNumber, 0x3ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotCacheFieldSharedItemNumber"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemDate, 0x3ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotCacheFieldSharedItemDate"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemIsDate, 0x7ULL>("isPivotCacheFieldSharedItemDate"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemText, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotCacheFieldSharedItemText"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldClear, 0x1ULL>("clearPivotCacheFields"),
          GuardedInstanceMethod<&Workbook::PivotCacheRefresh, 0x1ULL>("refreshPivotCache"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldTypes, 0x3ULL>("getPivotCacheFieldTypes"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldClearSharedItems, 0x3ULL>("clearPivotCacheFieldSharedItems"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldCount, 0x1ULL>("getPivotCacheFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldName, 0x3ULL>("getPivotCacheFieldName"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemCount, 0x3ULL>("getPivotCacheFieldSharedItemCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemValue, 0x7ULL>("getPivotCacheFieldSharedItemValue"),
          GuardedInstanceMethod<&Workbook::PivotCacheGetWorksheetSource, 0x1ULL>("getPivotCacheWorksheetSource"),
          GuardedInstanceMethod<&Workbook::PivotCacheWorksheetSourceRange, 0x1ULL>("getPivotCacheWorksheetSourceRange"),
          GuardedInstanceMethod<&Workbook::PivotCacheId, 0x3ULL>("getPivotCacheId"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordAdd, 0x1ULL>("addPivotCacheRecord"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordClear, 0x1ULL>("clearPivotCacheRecords"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordCount, 0x1ULL>("getPivotCacheRecordCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetBlank, 0x7ULL>("setPivotCacheRecordBlank"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetBool, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotCacheRecordBool"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetError, 0x7ULL, 0x8ULL>("setPivotCacheRecordError"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetNumber, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotCacheRecordNumber"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetDate, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotCacheRecordDate"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotCacheRecordText"),
          GuardedInstanceMethod<&Workbook::PivotCacheRemove, 0x1ULL>("removePivotCache"),
          GuardedInstanceMethod<&Workbook::PivotCacheSetWorksheetSource, 0x1ULL>("setPivotCacheWorksheetSource"),
          GuardedInstanceMethod<&Workbook::PivotColFieldOrder, 0x3ULL>("getPivotColumnFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotCount, 0x1ULL>("getPivotCount"),
          GuardedInstanceMethod<&Workbook::PivotCreate, 0x1DULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addPivot"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldAt, 0x7ULL>("getPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldClear, 0x3ULL>("clearPivotDataFields"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldCount, 0x3ULL>("getPivotDataFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldSet, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddItem, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x18ULL>("addPivotFieldItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetSubtotalFunctions, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldSubtotalFunctions"),
          GuardedInstanceMethod<&Workbook::PivotFieldAt, 0x7ULL>("getPivotField"),
          GuardedInstanceMethod<&Workbook::PivotFieldClear, 0x3ULL>("clearPivotFields"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearItems, 0x7ULL>("clearPivotFieldItems"),
          GuardedInstanceMethod<&Workbook::PivotFieldCount, 0x3ULL>("getPivotFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotFieldItemAt, 0xFULL>("getPivotFieldItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldItemCount, 0x7ULL>("getPivotFieldItemCount"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetAxis, 0x7ULL, 0x8ULL>("setPivotFieldAxis"),
          GuardedInstanceMethod<&Workbook::PivotUpdateField, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("updatePivotField"),
          GuardedInstanceMethod<&Workbook::PivotUpdateFieldItem, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("updatePivotFieldItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetNumericGroup, 0x7ULL>("setPivotFieldNumericGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldGetDateGroup, 0x7ULL>("getPivotFieldDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldGetNumericGroup, 0x7ULL>("getPivotFieldNumericGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetDateGroup, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL, 0x0ULL, 0x8ULL>("setPivotFieldDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFilterAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotFilterAt, 0x7ULL>("getPivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotFilterClear, 0x3ULL>("clearPivotFilters"),
          GuardedInstanceMethod<&Workbook::PivotFilterCount, 0x3ULL>("getPivotFilterCount"),
          GuardedInstanceMethod<&Workbook::PivotFilterRemoveAt, 0x7ULL>("removePivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotLayout, 0x3ULL>("getPivotLayout"),
          GuardedInstanceMethod<&Workbook::PivotRemove, 0x3ULL>("removePivot"),
          GuardedInstanceMethod<&Workbook::PivotSetAnchor, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotAnchor"),
          GuardedInstanceMethod<&Workbook::PivotSetPageFieldOrder, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotPageFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotSetColFieldOrder, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotColumnFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotUpdateGrandTotals, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xCULL>("updatePivotGrandTotals"),
          GuardedInstanceMethod<&Workbook::PivotUpdateStyleOptions, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updatePivotStyleOptions"),
          GuardedInstanceMethod<&Workbook::PivotGetGrandTotals, 0x3ULL>("getPivotGrandTotals"),
          GuardedInstanceMethod<&Workbook::PivotGetStyleOptions, 0x3ULL>("getPivotStyleOptions"),
          GuardedInstanceMethod<&Workbook::PivotGetLayout, 0x3ULL>("getPivotReportLayout"),
          GuardedInstanceMethod<&Workbook::PivotGetOptions, 0x3ULL>("getPivotOptions"),
          GuardedInstanceMethod<&Workbook::PivotUpdateOptions, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updatePivotOptions"),
          GuardedInstanceMethod<&Workbook::PivotCacheGetOptions, 0x1ULL>("getPivotCacheOptions"),
          GuardedInstanceMethod<&Workbook::PivotCacheUpdateOptions, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updatePivotCacheOptions"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedFieldAdd, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL>("addPivotCacheCalculatedField"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedFieldSetFormula, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotCacheCalculatedFieldFormula"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedFieldRemove, 0x3ULL>("removePivotCacheCalculatedField"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldFormula, 0x3ULL>("getPivotCacheFieldFormula"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xCULL>("addPivotCacheCalculatedItem"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemCount, 0x1ULL>("getPivotCacheCalculatedItemCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemAt, 0x3ULL>("getPivotCacheCalculatedItem"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemSetFormula, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotCacheCalculatedItemFormula"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemMove, 0x7ULL>("movePivotCacheCalculatedItem"),
          GuardedInstanceMethod<&Workbook::PivotCacheCalculatedItemRemove, 0x3ULL>("removePivotCacheCalculatedItem"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldDefaultName, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("getPivotDataFieldDefaultName"),
          GuardedInstanceMethod<&Workbook::PivotProject, 0x3ULL>("projectPivot"),
          GuardedInstanceMethod<&Workbook::PivotProjectionPending, 0x3ULL>("isPivotProjectionPending"),
          GuardedInstanceMethod<&Workbook::GetPivotSpill, 0x3ULL>("getPivotSpill"),
          GuardedInstanceMethod<&Workbook::PivotShowDetails, 0xFULL>("showPivotDetails"),
          GuardedInstanceMethod<&Workbook::PivotName, 0x3ULL>("getPivotName"),
          GuardedInstanceMethod<&Workbook::PivotPageFieldOrder, 0x3ULL>("getPivotPageFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotRuntimeId, 0x3ULL>("getPivotRuntimeId"),
          GuardedInstanceMethod<&Workbook::ResolvePivotRuntimeId, 0x3ULL>("resolvePivotRuntimeId"),
          GuardedInstanceMethod<&Workbook::PivotRowFieldOrder, 0x3ULL>("getPivotRowFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotSetLayout, 0x3ULL, 0x4ULL>("setPivotReportLayout"),
          GuardedInstanceMethod<&Workbook::PivotRename, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("renamePivot"),
          GuardedInstanceMethod<&Workbook::PivotSetRowFieldOrder, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotRowFieldOrder"),
          GuardedInstanceMethod<&Workbook::Precedents, 0xFULL>("getPrecedents"),
          GuardedInstanceMethod<&Workbook::Recalc>("recalc"),
          GuardedInstanceMethod<&Workbook::RecalcParallel>("recalcParallel"),
          GuardedInstanceMethod<&Workbook::HasDirtyCells>("hasDirtyCells"),
          GuardedInstanceMethod<&Workbook::RemoveConditionalFormatAt, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeCfRule"),
          GuardedInstanceMethod<&Workbook::InsertHyperlinkAt, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("insertHyperlinkAt"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlink, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeHyperlink"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlinkAt, 0x3ULL>("removeHyperlinkAt"),
          GuardedInstanceMethod<&Workbook::RemoveMerge, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeMerge"),
          GuardedInstanceMethod<&Workbook::RemoveMergeAt, 0x3ULL>("removeMergeAt"),
          GuardedInstanceMethod<&Workbook::RemoveSheet, 0x1ULL>("removeSheet"),
          GuardedInstanceMethod<&Workbook::RemoveValidationAt, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeValidation"),
          GuardedInstanceMethod<&Workbook::UpdateValidation, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateValidation"),
          GuardedInstanceMethod<&Workbook::UpdateCfRule, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateCfRule"),
          GuardedInstanceMethod<&Workbook::RenameSheet, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("renameSheet"),
          GuardedInstanceMethod<&Workbook::Save, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("save"),
          GuardedInstanceMethod<&Workbook::SaveWithDiagnostics, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("saveWithDiagnostics"),
          GuardedInstanceMethod<&Workbook::ReadDiagnostics>("getDiagnostics"),
          GuardedInstanceMethod<&Workbook::SetBlank, 0x7ULL>("setBlank"),
          GuardedInstanceMethod<&Workbook::SetBool, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setBool"),
          GuardedInstanceMethod<&Workbook::SetCalcMode, 0x0ULL, 0x1ULL>("setCalcMode"),
          GuardedInstanceMethod<&Workbook::SetFullPrecision, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setFullPrecision"),
          GuardedInstanceMethod<&Workbook::SetCalcOnSave, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setCalcOnSave"),
          GuardedInstanceMethod<&Workbook::SetRefMode, 0x0ULL, 0x1ULL>("setRefMode"),
          GuardedInstanceMethod<&Workbook::SetDate1904, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setDate1904"),
          GuardedInstanceMethod<&Workbook::SetPinnedNow, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setPinnedNow"),
          GuardedInstanceMethod<&Workbook::SetCellXfIndex, 0xFULL>("setCellXfIndex"),
          GuardedInstanceMethod<&Workbook::ClearCellXfIndex, 0x7ULL>("resetCellXfIndex"),
          GuardedInstanceMethod<&Workbook::SetCellPhonetic, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhonetic"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticRuns, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhoneticRuns"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticProperties, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhoneticProperties"),
          GuardedInstanceMethod<&Workbook::SetCellShowPhonetic, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellShowPhonetic"),
          GuardedInstanceMethod<&Workbook::SetRangeXfIndex, 0x5ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setRangeXfIndex"),
          GuardedInstanceMethod<&Workbook::SetColumnHidden, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnHidden"),
          GuardedInstanceMethod<&Workbook::SetColumnOutline, 0xFULL>("setColumnOutline"),
          GuardedInstanceMethod<&Workbook::SetColumnCollapsed, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnCollapsed"),
          GuardedInstanceMethod<&Workbook::SetColumnStyle, 0xFULL>("setColumnXfIndex"),
          GuardedInstanceMethod<&Workbook::SetColumnWidth, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnWidth"),
          GuardedInstanceMethod<&Workbook::SetComment, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setComment"),
          GuardedInstanceMethod<&Workbook::SetCommentVisible, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCommentVisible"),
          GuardedInstanceMethod<&Workbook::SetWorkbookProtection, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setWorkbookProtection"),
          GuardedInstanceMethod<&Workbook::AddProtectedRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addProtectedRange"),
          GuardedInstanceMethod<&Workbook::UpdateProtectedRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateProtectedRange"),
          GuardedInstanceMethod<&Workbook::RemoveProtectedRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeProtectedRange"),
          GuardedInstanceMethod<&Workbook::SetDefinedName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setDefinedName"),
          GuardedInstanceMethod<&Workbook::SetDefaultFont, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setDefaultFont"),
          GuardedInstanceMethod<&Workbook::SetFill, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setFill"),
          GuardedInstanceMethod<&Workbook::SetBorder, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setBorder"),
          GuardedInstanceMethod<&Workbook::SetCellXf, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setCellXf"),
          GuardedInstanceMethod<&Workbook::SetCellStyleXf, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setCellStyleXf"),
          GuardedInstanceMethod<&Workbook::SetDxf, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setDxf"),
          GuardedInstanceMethod<&Workbook::SetFont, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setFont"),
          GuardedInstanceMethod<&Workbook::SetError, 0x7ULL, 0x8ULL>("setError"),
          GuardedInstanceMethod<&Workbook::SetProfileId, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setProfileId"),
          GuardedInstanceMethod<&Workbook::SetLocation, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setLocation"),
          GuardedInstanceMethod<&Workbook::SetFormula, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setFormula"),
          GuardedInstanceMethod<&Workbook::SetIterative, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("updateIterative"),
          GuardedInstanceMethod<&Workbook::SetIterativeProgress, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setIterativeProgress"),
          GuardedInstanceMethod<&Workbook::SetNumber, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setNumber"),
          GuardedInstanceMethod<&Workbook::SetRowAutoHeight, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowAutoHeight"),
          GuardedInstanceMethod<&Workbook::SetRowHeight, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowHeight"),
          GuardedInstanceMethod<&Workbook::SetRowHidden, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowHidden"),
          GuardedInstanceMethod<&Workbook::SetRowOutline, 0xFULL>("setRowOutline"),
          GuardedInstanceMethod<&Workbook::SetRowCollapsed, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowCollapsed"),
          GuardedInstanceMethod<&Workbook::SetRowStyle, 0xFULL>("setRowXfIndex"),
          GuardedInstanceMethod<&Workbook::RangeStyleBases, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getRangeXfBases"),
          GuardedInstanceMethod<&Workbook::ListCellXfIndices, 0x1ULL>("listCellXfIndices"),
          GuardedInstanceMethod<&Workbook::ListShowPhoneticCells, 0x1ULL>("listShowPhoneticCells"),
          GuardedInstanceMethod<&Workbook::RangeStyleBasesExcluding, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xEULL>("getRangeXfBasesExcluding"),
          GuardedInstanceMethod<&Workbook::RemapRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xEULL>("remapRangeXfs"),
          GuardedInstanceMethod<&Workbook::SnapshotRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("snapshotRangeXfs"),
          GuardedInstanceMethod<&Workbook::RestoreRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("restoreRangeXfs"),
          GuardedInstanceMethod<&Workbook::SetSheetProtection, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetProtection"),
          GuardedInstanceMethod<&Workbook::UpdateSheetView, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updateView"),
          GuardedInstanceMethod<&Workbook::UpdateSheetProtection, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updateSheetProtection"),
          GuardedInstanceMethod<&Workbook::SetSheetSelection, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSelection"),
          GuardedInstanceMethod<&Workbook::SetText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setText"),
          GuardedInstanceMethod<&Workbook::SheetCount>("getSheetCount"),
          GuardedInstanceMethod<&Workbook::ActiveSheet>("activeSheet"),
          GuardedInstanceMethod<&Workbook::GetDefaultFont>("getDefaultFont"),
          GuardedInstanceMethod<&Workbook::SetActiveSheet, 0x1ULL>("setActiveSheet"),
          GuardedInstanceMethod<&Workbook::SheetName, 0x1ULL>("getSheetName"),
          GuardedInstanceMethod<&Workbook::GetSheetIndex, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getSheetIndex"),
          GuardedInstanceMethod<&Workbook::GetSpill, 0x7ULL>("getSpill"),
          GuardedInstanceMethod<&Workbook::ParseReference, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("parseReference"),
          GuardedInstanceMethod<&Workbook::FormatReference, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("formatReference"),
          GuardedInstanceMethod<&Workbook::GetArrayFormula, 0x7ULL>("getArrayFormula"),
          GuardedInstanceMethod<&Workbook::SetArrayFormula, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setArrayFormula"),
          GuardedInstanceMethod<&Workbook::TableAt, 0x3ULL>("getTable"),
          GuardedInstanceMethod<&Workbook::TableRuntimeId, 0x3ULL>("getTableRuntimeId"),
          GuardedInstanceMethod<&Workbook::ResolveTableRuntimeId, 0x1ULL>("resolveTableRuntimeId"),
          GuardedInstanceMethod<&Workbook::FindTable, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("findTable"),
          GuardedInstanceMethod<&Workbook::CreateTable, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addTable"),
          GuardedInstanceMethod<&Workbook::UpdateTable, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateTable"),
          GuardedInstanceMethod<&Workbook::RenameTable, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("renameTable"),
          GuardedInstanceMethod<&Workbook::GetTableStyleOptions, 0x3ULL>("getTableStyleOptions"),
          GuardedInstanceMethod<&Workbook::UpdateTableStyleOptions, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateTableStyleOptions"),
          GuardedInstanceMethod<&Workbook::RemoveTable, 0x3ULL>("removeTable"),
          GuardedInstanceMethod<&Workbook::RenameTableColumn, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("renameTableColumn"),
          GuardedInstanceMethod<&Workbook::InsertTableRows, 0xFULL>("insertTableRows"),
          GuardedInstanceMethod<&Workbook::InsertTableColumns, 0xFULL>("insertTableColumns"),
          GuardedInstanceMethod<&Workbook::TableCount, 0x1ULL>("getTableCount"),
          GuardedInstanceMethod<&Workbook::XfCount>("getCellXfCount"),
      });
  // clang-format on
  // `using` calls [Symbol.dispose]; it is the same function as dispose().
  const Napi::Value dispose_symbol = env.Global().Get("Symbol").As<Napi::Object>().Get("dispose");
  if (dispose_symbol.IsSymbol()) {
    Napi::Object prototype = constructor.Get("prototype").As<Napi::Object>();
    prototype.Set(dispose_symbol, prototype.Get("dispose"));
  }
  data.workbook_class = Napi::Persistent(constructor);
  return constructor;
}

void RegisterModuleFunctions(Napi::Env env, Napi::Object exports) {
  // clang-format off
  SetGuardedModuleFunction<&CanonicalizeFormula, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>(env, exports, "canonicalizeFormula");
  SetGuardedModuleFunction<&CanonicalizeFunctionName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>(env, exports, "canonicalizeFunctionName");
  SetGuardedModuleFunction<&CompareText, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xBULL>(env, exports, "compareText");
  SetGuardedModuleFunction<&ConvertCivilToSerial, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "convertCivilToSerial");
  SetGuardedModuleFunction<&ConvertSerialToCivil, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "convertSerialToCivil");
  SetGuardedModuleFunction<&CollationClasses, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "getCollationClasses");
  SetGuardedModuleFunction<&FunctionMetadata, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "getFunctionMetadata");
  SetGuardedModuleFunction<&ProbeImage, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "getImageInfo");
  SetGuardedModuleFunction<&LocaleFacts, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>(env, exports, "getLocaleFacts");
  SetGuardedModuleFunction<&FunctionNames>(env, exports, "listFunctionNames");
  SetGuardedModuleFunction<&ListProfileIds>(env, exports, "listProfileIds");
  SetGuardedModuleFunction<&LocalizeFormula, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>(env, exports, "localizeFormula");
  SetGuardedModuleFunction<&LocalizeFunctionName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>(env, exports, "localizeFunctionName");
  // clang-format on
}

}  // namespace formulon_node
