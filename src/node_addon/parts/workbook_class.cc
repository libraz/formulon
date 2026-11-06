// `Workbook` ObjectWrap lifecycle and class registration. Holds the
// ctor / dtor, the shared argument extraction helpers, and the
// `DefineClass` table that wires every per-area method into the JS
// class. The per-area method bodies live in `parts/*.cc`.

#include "node_addon/parts/workbook_class.h"

#include <cmath>
#include <cstdint>
#include <limits>

namespace formulon_node {
namespace {

// Positional arguments are checked by the registration wrapper before a
// part-TU method reads them. The masks are indexed by the JavaScript argument
// position. A typed, `object` or `required` position must be present; an
// unset bit means the position is optional and handled by the method itself.
// `tools/codegen/gen_wasm_positional_guards.py` emits the same table into the
// WASM pre-js, so both surfaces reject the same arguments the same way.
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

// A uint64 argument arrives as a Number, which is exact only up to 2^53 - 1;
// a larger value could not name the cursor the caller holds.
constexpr double kMaxSafeInteger = 9007199254740991.0;

bool RejectArgumentType(const Napi::CallbackInfo& info, const char* expected) {
  Napi::TypeError::New(info.Env(), expected).ThrowAsJavaScriptException();
  return false;
}

bool RejectArgumentRange(const Napi::CallbackInfo& info, const char* expected) {
  Napi::RangeError::New(info.Env(), expected).ThrowAsJavaScriptException();
  return false;
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

bool ValidatePositionalArgs(const Napi::CallbackInfo& info, const PositionalArgSpec& spec) {
  // A number that is an integer in [min_value, max_value].
  const auto validate_integer = [&](std::size_t idx, double min_value, double max_value, const char* range_message) {
    const Napi::Value value = info[idx];
    if (!value.IsNumber()) {
      return RejectArgumentType(info, "positional argument must be a number");
    }
    const double number = NumberValue(value, 0.0);
    if (!std::isfinite(number) || std::trunc(number) != number || number < min_value || number > max_value) {
      return RejectArgumentRange(info, range_message);
    }
    return true;
  };
  // An optional position may be absent or `undefined`; `null` only when its bit is in `optional_null_mask`.
  const auto validate_optional_integer = [&](std::size_t idx, std::uint64_t bit, double max_value,
                                             const char* range_message) {
    if (idx >= info.Length()) {
      return true;
    }
    const Napi::Value value = info[idx];
    if (value.IsNull() && (spec.optional_null_mask & bit) == 0) {
      return RejectArgumentType(info, "positional argument must be a number");
    }
    return value.IsUndefined() || value.IsNull() || validate_integer(idx, 0.0, max_value, range_message);
  };

  // Required positions are checked even past the supplied arguments, where
  // `info[idx]` reads `undefined`.
  const std::size_t required_width = MaskWidth(spec.u32_mask | spec.i32_mask | spec.u64_mask | spec.double_mask |
                                               spec.string_mask | spec.object_mask | spec.required_mask);
  const std::size_t supplied = info.Length() < 64 ? info.Length() : 64;
  const std::size_t positional_count = supplied > required_width ? supplied : required_width;
  for (std::size_t idx = 0; idx < positional_count; ++idx) {
    const std::uint64_t bit = std::uint64_t{1} << idx;
    if ((spec.u32_mask & bit) != 0 &&
        !validate_integer(idx, 0.0, kUint32MaxNumber, "positional argument is outside uint32 range")) {
      return false;
    }
    if ((spec.optional_u32_mask & bit) != 0 &&
        !validate_optional_integer(idx, bit, kUint32MaxNumber, "positional argument is outside uint32 range")) {
      return false;
    }
    if ((spec.i32_mask & bit) != 0 &&
        !validate_integer(idx, static_cast<double>(std::numeric_limits<std::int32_t>::min()),
                          static_cast<double>(std::numeric_limits<std::int32_t>::max()),
                          "positional argument is outside int32 range")) {
      return false;
    }
    if ((spec.u64_mask & bit) != 0 &&
        !validate_integer(idx, 0.0, kMaxSafeInteger, "positional argument is outside the safe-integer range")) {
      return false;
    }
    if ((spec.optional_u64_mask & bit) != 0 &&
        !validate_optional_integer(idx, bit, kMaxSafeInteger,
                                   "positional argument is outside the safe-integer range")) {
      return false;
    }
    if ((spec.double_mask & bit) != 0 && !info[idx].IsNumber()) {
      return RejectArgumentType(info, "positional argument must be a number");
    }
    if ((spec.string_mask & bit) != 0 && !info[idx].IsString()) {
      return RejectArgumentType(info, "positional argument must be a string");
    }
    if ((spec.object_mask & bit) != 0 && !info[idx].IsObject()) {
      return RejectArgumentType(info, "positional argument must be an object");
    }
    if ((spec.required_mask & bit) != 0 && info[idx].IsUndefined()) {
      return RejectArgumentType(info, "positional argument is required");
    }
  }
  return true;
}

using InstanceMethodCallback = Napi::Value (Workbook::*)(const Napi::CallbackInfo&);

template <InstanceMethodCallback method, std::uint64_t u32_mask, std::uint64_t i32_mask, std::uint64_t u64_mask,
          std::uint64_t double_mask, std::uint64_t string_mask, std::uint64_t optional_u32_mask,
          std::uint64_t optional_u64_mask, std::uint64_t optional_null_mask, std::uint64_t object_mask,
          std::uint64_t required_mask>
struct GuardedMethodSpec {
  static constexpr PositionalArgSpec value{u32_mask,    i32_mask,          u64_mask,          double_mask,
                                           string_mask, optional_u32_mask, optional_u64_mask, optional_null_mask,
                                           object_mask, required_mask};
};

template <InstanceMethodCallback method, std::uint64_t u32_mask, std::uint64_t i32_mask, std::uint64_t u64_mask,
          std::uint64_t double_mask, std::uint64_t string_mask, std::uint64_t optional_u32_mask,
          std::uint64_t optional_u64_mask, std::uint64_t optional_null_mask, std::uint64_t object_mask,
          std::uint64_t required_mask>
napi_value GuardedInstanceMethodCallback(napi_env raw_env, napi_callback_info raw_info) {
  return Napi::details::WrapCallback(raw_env, [&] {
    Napi::CallbackInfo info(raw_env, raw_info);
    const auto* const spec = static_cast<const PositionalArgSpec*>(info.Data());
    if (spec == nullptr || !ValidatePositionalArgs(info, *spec)) {
      return info.Env().Undefined();
    }
    Workbook* const instance = Workbook::Unwrap(info.This().As<Napi::Object>());
    return instance == nullptr ? Napi::Value() : instance->RunActive(method, info);
  });
}

template <InstanceMethodCallback method, std::uint64_t u32_mask = 0, std::uint64_t i32_mask = 0,
          std::uint64_t u64_mask = 0, std::uint64_t double_mask = 0, std::uint64_t string_mask = 0,
          std::uint64_t optional_u32_mask = 0, std::uint64_t optional_u64_mask = 0,
          std::uint64_t optional_null_mask = 0, std::uint64_t object_mask = 0, std::uint64_t required_mask = 0>
Napi::ClassPropertyDescriptor<Workbook> GuardedInstanceMethod(const char* name) {
  napi_property_descriptor descriptor{};
  descriptor.utf8name = name;
  descriptor.method =
      &GuardedInstanceMethodCallback<method, u32_mask, i32_mask, u64_mask, double_mask, string_mask, optional_u32_mask,
                                     optional_u64_mask, optional_null_mask, object_mask, required_mask>;
  descriptor.data = const_cast<void*>(static_cast<const void*>(
      &GuardedMethodSpec<method, u32_mask, i32_mask, u64_mask, double_mask, string_mask, optional_u32_mask,
                         optional_u64_mask, optional_null_mask, object_mask, required_mask>::value));
  descriptor.attributes = napi_default;
  return descriptor;
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
  if (idx >= info.Length()) {
    return false;
  }
  return info[idx].ToBoolean().Value();
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
  const uint32_t index = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  return ValueOrThrow(env, fn(handle_, sheet, index, count), env.Undefined());
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
    RejectArgumentType(info, "positional argument must be a string");
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
    RejectArgumentType(info, "positional argument must be an object");
    return env.Undefined();
  }
  Napi::Uint8Array u8;
  if (!BytesArg(info, 0, &u8)) {
    RejectArgumentType(info, "bytes must be a Uint8Array");
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
          GuardedInstanceMethod<&Workbook::AddHyperlink, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x78ULL>("addHyperlink"),
          GuardedInstanceMethod<&Workbook::AddHyperlinkRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3CULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addHyperlinkRange"),
          GuardedInstanceMethod<&Workbook::AddMerge, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addMerge"),
          GuardedInstanceMethod<&Workbook::AddNumFmt, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addNumFmt"),
          GuardedInstanceMethod<&Workbook::AddSheet, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addSheet"),
          GuardedInstanceMethod<&Workbook::AddValidation, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addValidation"),
          GuardedInstanceMethod<&Workbook::AddPerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addPerson"),
          GuardedInstanceMethod<&Workbook::AddThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addThreadedComment"),
          GuardedInstanceMethod<&Workbook::ApplyAutoFilter, 0x1ULL>("applyAutoFilter"),
          GuardedInstanceMethod<&Workbook::ApplyTableAutoFilter, 0x1ULL>("applyTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::ClearAutoFilter, 0x1ULL>("clearAutoFilter"),
          GuardedInstanceMethod<&Workbook::ClearTableAutoFilter, 0x1ULL>("clearTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::EditThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("updateThreadedComment"),
          GuardedInstanceMethod<&Workbook::EvaluateAutoFilter, 0x1ULL>("evaluateAutoFilter"),
          GuardedInstanceMethod<&Workbook::EvaluateTableAutoFilter, 0x1ULL>("evaluateTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetAutoFilter, 0x1ULL>("getAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetAutoFilterXml, 0x1ULL>("getAutoFilterXml"),
          GuardedInstanceMethod<&Workbook::GetPersons>("listPeople"),
          GuardedInstanceMethod<&Workbook::GetTableAutoFilter, 0x1ULL>("getTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetThreadedComments, 0x1ULL>("listThreadedComments"),
          GuardedInstanceMethod<&Workbook::ListInvalidCells, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x2ULL, 0x6ULL>("getInvalidCells"),
          GuardedInstanceMethod<&Workbook::RemoveAutoFilter, 0x1ULL>("removeAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetAutoFilterXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setAutoFilterXml"),
          GuardedInstanceMethod<&Workbook::RemovePerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removePerson"),
          GuardedInstanceMethod<&Workbook::GetPerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getPerson"),
          GuardedInstanceMethod<&Workbook::RemoveTableAutoFilter, 0x1ULL>("removeTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::RemoveThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeThreadedComment"),
          GuardedInstanceMethod<&Workbook::GetThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getThreadedComment"),
          GuardedInstanceMethod<&Workbook::SetAutoFilter, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetTableAutoFilter, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetThreadResolved, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setThreadResolved"),
          GuardedInstanceMethod<&Workbook::ValidateValue, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("validateValue"),
          GuardedInstanceMethod<&Workbook::ProbeImage, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getImageInfo"),
          GuardedInstanceMethod<&Workbook::ListDrawingObjects, 0x1ULL>("listDrawingObjects"),
          GuardedInstanceMethod<&Workbook::GetDrawingObject, 0x3ULL>("getDrawingObject"),
          GuardedInstanceMethod<&Workbook::GetImage, 0x3ULL>("getImage"),
          GuardedInstanceMethod<&Workbook::InsertImage, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addImage"),
          GuardedInstanceMethod<&Workbook::RemoveDrawing, 0x3ULL>("removeDrawingObject"),
          GuardedInstanceMethod<&Workbook::SetDrawingAnchor, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setDrawingAnchor"),
          GuardedInstanceMethod<&Workbook::SetDrawingZOrder, 0x7ULL>("setDrawingZOrder"),
          GuardedInstanceMethod<&Workbook::SnapshotDrawing, 0x3ULL>("snapshotDrawing"),
          GuardedInstanceMethod<&Workbook::RestoreDrawing, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("restoreDrawing"),
          GuardedInstanceMethod<&Workbook::ListShapes, 0x1ULL>("listShapes"),
          GuardedInstanceMethod<&Workbook::InsertShape, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addShape"),
          GuardedInstanceMethod<&Workbook::UpdateShape, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateShape"),
          GuardedInstanceMethod<&Workbook::RemoveShape, 0x3ULL>("removeShape"),
          GuardedInstanceMethod<&Workbook::GetShape, 0x3ULL>("getShape"),
          GuardedInstanceMethod<&Workbook::ListSlicers, 0x1ULL>("listSlicers"),
          GuardedInstanceMethod<&Workbook::GetSlicer, 0x3ULL>("getSlicer"),
          GuardedInstanceMethod<&Workbook::AddSlicer, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("addSlicer"),
          GuardedInstanceMethod<&Workbook::UpdateSlicer, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateSlicer"),
          GuardedInstanceMethod<&Workbook::RemoveSlicer, 0x3ULL>("removeSlicer"),
          GuardedInstanceMethod<&Workbook::AddDataTable, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addDataTable"),
          GuardedInstanceMethod<&Workbook::GetDataTable, 0x7ULL>("getDataTable"),
          GuardedInstanceMethod<&Workbook::RemoveDataTable, 0x7ULL>("removeDataTable"),
          GuardedInstanceMethod<&Workbook::RecalcDataTables>("recalcDataTables"),
          GuardedInstanceMethod<&Workbook::GetSparklineGroupCount, 0x1ULL>("getSparklineGroupCount"),
          GuardedInstanceMethod<&Workbook::GetSparklineGroup, 0x3ULL>("getSparklineGroup"),
          GuardedInstanceMethod<&Workbook::AddSparklineGroup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addSparklineGroup"),
          GuardedInstanceMethod<&Workbook::UpdateSparklineGroup, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateSparklineGroup"),
          GuardedInstanceMethod<&Workbook::SetSparklineGroupSparklines, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setSparklineGroupSparklines"),
          GuardedInstanceMethod<&Workbook::RemoveSparklineGroup, 0x3ULL>("removeSparklineGroup"),
          GuardedInstanceMethod<&Workbook::GetChart, 0x3ULL>("getChart"),
          GuardedInstanceMethod<&Workbook::GetChartData, 0x3ULL>("getChartData"),
          GuardedInstanceMethod<&Workbook::PreviewChartData, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("getChartPreview"),
          GuardedInstanceMethod<&Workbook::InsertChart, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addChart"),
          GuardedInstanceMethod<&Workbook::SetChartType, 0x7ULL, 0x8ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setChartType"),
          GuardedInstanceMethod<&Workbook::SetChartTitle, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setChartTitle"),
          GuardedInstanceMethod<&Workbook::SetChartLegend, 0x3ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setChartLegend"),
          GuardedInstanceMethod<&Workbook::SetChartAxis, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setChartAxis"),
          GuardedInstanceMethod<&Workbook::SetChartDisplay, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setChartDisplay"),
          GuardedInstanceMethod<&Workbook::SetChartSeries, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setChartSeries"),
          GuardedInstanceMethod<&Workbook::SetChartPoint, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("setChartPoint"),
          GuardedInstanceMethod<&Workbook::SetChartPlot, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setChartPlot"),
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
          GuardedInstanceMethod<&Workbook::ClearRowHeight, 0x7ULL>("clearRowHeight"),
          GuardedInstanceMethod<&Workbook::ClearColumnWidth, 0x7ULL>("clearColumnWidth"),
          GuardedInstanceMethod<&Workbook::ClearColumnStyle, 0x7ULL>("clearColumnStyle"),
          GuardedInstanceMethod<&Workbook::ClearRowStyle, 0x7ULL>("clearRowStyle"),
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
          GuardedInstanceMethod<&Workbook::ClearPinnedNow>("clearPinnedNow"),
          GuardedInstanceMethod<&Workbook::CanonicalizeFormula, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("canonicalizeFormula"),
          GuardedInstanceMethod<&Workbook::CanonicalizeFunctionName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("canonicalizeFunctionName"),
          GuardedInstanceMethod<&Workbook::CellAt, 0x3ULL>("getCellAt"),
          GuardedInstanceMethod<&Workbook::CellCount, 0x1ULL>("getCellCount"),
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
          GuardedInstanceMethod<&Workbook::Dependents, 0xFULL>("getDependents"),
          GuardedInstanceMethod<&Workbook::Dispose>("dispose"),
          InstanceAccessor<&Workbook::Disposed>("disposed"),
          GuardedInstanceMethod<&Workbook::DxfCount>("getDxfCount"),
          GuardedInstanceMethod<&Workbook::EvaluateCfRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("evaluateCfRange"),
          GuardedInstanceMethod<&Workbook::EvaluateConditionalFormula, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("evaluateCfFormula"),
          GuardedInstanceMethod<&Workbook::EvaluateFormulaArray, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("evaluateFormulaArray"),
          GuardedInstanceMethod<&Workbook::EvaluateFormulaText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("evaluateFormula"),
          GuardedInstanceMethod<&Workbook::ExcelProfileId>("getExcelProfileId"),
          GuardedInstanceMethod<&Workbook::FileName>("getFileName"),
          GuardedInstanceMethod<&Workbook::FillCount>("getFillCount"),
          GuardedInstanceMethod<&Workbook::FontCount>("getFontCount"),
          GuardedInstanceMethod<&Workbook::FunctionMetadata, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getFunctionMetadata"),
          GuardedInstanceMethod<&Workbook::FunctionNames>("listFunctionNames"),
          GuardedInstanceMethod<&Workbook::GetBorder, 0x1ULL>("getBorder"),
          GuardedInstanceMethod<&Workbook::GetCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getCellStyle"),
          GuardedInstanceMethod<&Workbook::GetCellStyleXf, 0x1ULL>("getCellStyleXf"),
          GuardedInstanceMethod<&Workbook::GetEffectiveStyle, 0x7ULL>("getEffectiveStyle"),
          GuardedInstanceMethod<&Workbook::GetTheme>("getThemeFonts"),
          GuardedInstanceMethod<&Workbook::RemoveCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeCellStyle"),
          GuardedInstanceMethod<&Workbook::ResolveColor, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("resolveColor"),
          GuardedInstanceMethod<&Workbook::SetCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setCellStyle"),
          GuardedInstanceMethod<&Workbook::ListTableStyles>("listTableStyles"),
          GuardedInstanceMethod<&Workbook::GetTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getTableStyle"),
          GuardedInstanceMethod<&Workbook::SetTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setTableStyle"),
          GuardedInstanceMethod<&Workbook::RemoveTableStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeTableStyle"),
          GuardedInstanceMethod<&Workbook::GetTableStyleDefaults>("getTableStyleDefaults"),
          GuardedInstanceMethod<&Workbook::SetTableStyleDefaults, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setTableStyleDefaults"),
          GuardedInstanceMethod<&Workbook::SetThemeColors, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setThemeColors"),
          GuardedInstanceMethod<&Workbook::SetThemeFonts, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setThemeFonts"),
          GuardedInstanceMethod<&Workbook::SetThemeScriptFont, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setThemeScriptFont"),
          GuardedInstanceMethod<&Workbook::GetThemeScriptFont, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getThemeScriptFont"),
          GuardedInstanceMethod<&Workbook::ResetTheme>("clearTheme"),
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
          GuardedInstanceMethod<&Workbook::GetProtectedRange, 0x3ULL>("getProtectedRange"),
          GuardedInstanceMethod<&Workbook::GetCfRuleCount, 0x1ULL>("getCfRuleCount"),
          GuardedInstanceMethod<&Workbook::GetCfRule, 0x3ULL>("getCfRule"),
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
          GuardedInstanceMethod<&Workbook::GetNumFmt, 0x1ULL>("getNumFmtString"),
          GuardedInstanceMethod<&Workbook::GetSheetColumns, 0x1ULL>("listColumnOverrides"),
          GuardedInstanceMethod<&Workbook::GetSheetProtection, 0x1ULL>("getSheetProtection"),
          GuardedInstanceMethod<&Workbook::GetSheetRowOverrides, 0x1ULL>("listRowOverrides"),
          GuardedInstanceMethod<&Workbook::GetSheetSelection, 0x1ULL>("getSelection"),
          GuardedInstanceMethod<&Workbook::GetSheetView, 0x1ULL>("getView"),
          GuardedInstanceMethod<&Workbook::GetValidationCount, 0x1ULL>("getValidationCount"),
          GuardedInstanceMethod<&Workbook::GetValidation, 0x3ULL>("getValidation"),
          GuardedInstanceMethod<&Workbook::GetValue, 0x7ULL>("getValue"),
          GuardedInstanceMethod<&Workbook::InsertCols, 0x7ULL>("insertColumns"),
          GuardedInstanceMethod<&Workbook::InsertRows, 0x7ULL>("insertRows"),
          GuardedInstanceMethod<&Workbook::CollationClasses, 0x2ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getCollationClasses"),
          GuardedInstanceMethod<&Workbook::CompareText, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xBULL>("compareText"),
          GuardedInstanceMethod<&Workbook::LocaleFacts, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getLocaleFacts"),
          GuardedInstanceMethod<&Workbook::ListLocaleErrorNames, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("listLocaleErrorNames"),
          GuardedInstanceMethod<&Workbook::LocalizeFormula, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("localizeFormula"),
          GuardedInstanceMethod<&Workbook::LocalizeFunctionName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("localizeFunctionName"),
          GuardedInstanceMethod<&Workbook::MemoryUsage>("getMemoryUsage"),
          GuardedInstanceMethod<&Workbook::MoveSheet, 0x3ULL>("moveSheet"),
          GuardedInstanceMethod<&Workbook::CopySheet, 0x3ULL>("copySheet"),
          GuardedInstanceMethod<&Workbook::PartialRecalc, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("recalcViewport"),
          GuardedInstanceMethod<&Workbook::Paginate, 0x1ULL>("getPagination"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetupXml, 0x1ULL>("getPageSetupXml"),
          GuardedInstanceMethod<&Workbook::SetSheetFormatDefaults, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setFormatDefaults"),
          GuardedInstanceMethod<&Workbook::SetSheetPageSetupXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageSetupXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMarginsXml, 0x1ULL>("getPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPageMarginsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintOptionsXml, 0x1ULL>("getPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintOptionsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetHeaderFooterXml, 0x1ULL>("getHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::SetSheetHeaderFooterXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::GetSheetSheetPrXml, 0x1ULL>("getSheetPrXml"),
          GuardedInstanceMethod<&Workbook::SetSheetSheetPrXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPrXml"),
          GuardedInstanceMethod<&Workbook::SetSheetFitToPage, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setFitToPage"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintArea, 0x1ULL>("getPrintArea"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintArea, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintArea"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintTitles, 0x1ULL>("getPrintTitles"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintTitles, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL>("setPrintTitles"),
          GuardedInstanceMethod<&Workbook::AddSheetRowBreak, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addRowBreak"),
          GuardedInstanceMethod<&Workbook::AddSheetColBreak, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addColumnBreak"),
          GuardedInstanceMethod<&Workbook::RemoveSheetRowBreak, 0x3ULL>("removeRowBreak"),
          GuardedInstanceMethod<&Workbook::RemoveSheetColBreak, 0x3ULL>("removeColumnBreak"),
          GuardedInstanceMethod<&Workbook::ClearSheetBreaks, 0x1ULL>("clearBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetRowBreaks, 0x1ULL>("listRowBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetColBreaks, 0x1ULL>("listColumnBreaks"),
          GuardedInstanceMethod<&Workbook::SetSheetPageSetup, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageSetup"),
          GuardedInstanceMethod<&Workbook::SetSheetPageMargins, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPageMargins"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintOptions, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPrintOptions"),
          GuardedInstanceMethod<&Workbook::SetSheetHeaderFooter, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setHeaderFooter"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetup, 0x1ULL>("getPageSetup"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMargins, 0x1ULL>("getPageMargins"),
          GuardedInstanceMethod<&Workbook::ListPassthroughs>("listPassthroughs"),
          GuardedInstanceMethod<&Workbook::GetPassthrough, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getPassthrough"),
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
          GuardedInstanceMethod<&Workbook::PivotCacheFieldClearSharedItems, 0x3ULL>("clearPivotCacheFieldSharedItems"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldCount, 0x1ULL>("getPivotCacheFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldName, 0x3ULL>("getPivotCacheFieldName"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemCount, 0x3ULL>("getPivotCacheFieldSharedItemCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemValue, 0x7ULL>("getPivotCacheFieldSharedItemValue"),
          GuardedInstanceMethod<&Workbook::PivotCacheGetWorksheetSource, 0x1ULL>("getPivotCacheWorksheetSource"),
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
          GuardedInstanceMethod<&Workbook::PivotCacheSetWorksheetSource, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPivotCacheWorksheetSource"),
          GuardedInstanceMethod<&Workbook::PivotColFieldOrder, 0x3ULL>("getPivotColumnFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotCount, 0x1ULL>("getPivotCount"),
          GuardedInstanceMethod<&Workbook::PivotCreate, 0x1DULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addPivot"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldAt, 0x7ULL>("getPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldClear, 0x3ULL>("clearPivotDataFields"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldCount, 0x3ULL>("getPivotDataFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldSet, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotDataField"),
          GuardedInstanceMethod<&Workbook::PivotFieldAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotField"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddItem, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("addPivotFieldItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddItemAt, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("addPivotFieldSharedItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddSubtotalFn, 0x7ULL, 0x8ULL>("addPivotFieldSubtotalFn"),
          GuardedInstanceMethod<&Workbook::PivotFieldAt, 0x7ULL>("getPivotField"),
          GuardedInstanceMethod<&Workbook::PivotFieldClear, 0x3ULL>("clearPivotFields"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearDateGroup, 0x7ULL>("clearPivotFieldDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearItems, 0x7ULL>("clearPivotFieldItems"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearSubtotalFns, 0x7ULL>("clearPivotFieldSubtotalFns"),
          GuardedInstanceMethod<&Workbook::PivotFieldCount, 0x3ULL>("getPivotFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotFieldItemAt, 0xFULL>("getPivotFieldItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldItemCount, 0x7ULL>("getPivotFieldItemCount"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetAxis, 0x7ULL, 0x8ULL>("setPivotFieldAxis"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetDateGroup, 0xFULL, 0x10ULL>("setPivotFieldDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetItemVisible, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setPivotFieldItemVisible"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetNumberFormat, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldNumberFormat"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetSort, 0x7ULL, 0x8ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setPivotFieldSort"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetSubtotalTop, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldSubtotalTop"),
          GuardedInstanceMethod<&Workbook::PivotFilterAdd, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("addPivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotFilterAt, 0x7ULL>("getPivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotFilterClear, 0x3ULL>("clearPivotFilters"),
          GuardedInstanceMethod<&Workbook::PivotFilterCount, 0x3ULL>("getPivotFilterCount"),
          GuardedInstanceMethod<&Workbook::PivotFilterRemoveAt, 0x7ULL>("removePivotFilter"),
          GuardedInstanceMethod<&Workbook::PivotLayout, 0x3ULL>("getPivotLayout"),
          GuardedInstanceMethod<&Workbook::PivotRemove, 0x3ULL>("removePivot"),
          GuardedInstanceMethod<&Workbook::PivotSetAnchor, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotAnchor"),
          GuardedInstanceMethod<&Workbook::PivotSetColFieldOrder, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotColumnFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotSetGrandTotals, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xCULL>("setPivotGrandTotals"),
          GuardedInstanceMethod<&Workbook::PivotSetStyleInfo, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotStyleInfo"),
          GuardedInstanceMethod<&Workbook::PivotGetGrandTotals, 0x3ULL>("getPivotGrandTotals"),
          GuardedInstanceMethod<&Workbook::PivotGetStyleInfo, 0x3ULL>("getPivotStyleInfo"),
          GuardedInstanceMethod<&Workbook::PivotGetLayout, 0x3ULL>("getPivotReportLayout"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetDefaultSubtotal, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldDefaultSubtotal"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetInsertBlankRow, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldInsertBlankRow"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetFillDownLabels, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldFillDownLabels"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetShowDetail, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldShowDetail"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetItemShowDetail, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setPivotFieldItemShowDetail"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetItemCaption, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("setPivotFieldItemCaption"),
          GuardedInstanceMethod<&Workbook::PivotGetOptions, 0x3ULL>("getPivotOptions"),
          GuardedInstanceMethod<&Workbook::PivotSetOptions, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotOptions"),
          GuardedInstanceMethod<&Workbook::PivotCacheGetOptions, 0x1ULL>("getPivotCacheOptions"),
          GuardedInstanceMethod<&Workbook::PivotCacheSetOptions, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setPivotCacheOptions"),
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
          GuardedInstanceMethod<&Workbook::PivotFieldSetCustomName, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldCustomName"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetShowAll, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setPivotFieldShowAll"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldDefaultName, 0xFULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("getPivotDataFieldDefaultName"),
          GuardedInstanceMethod<&Workbook::PivotProject, 0x3ULL>("projectPivot"),
          GuardedInstanceMethod<&Workbook::PivotProjectionPending, 0x3ULL>("isPivotProjectionPending"),
          GuardedInstanceMethod<&Workbook::PivotShowDetails, 0xFULL>("showPivotDetails"),
          GuardedInstanceMethod<&Workbook::SetDrawingName, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setDrawingName"),
          GuardedInstanceMethod<&Workbook::PivotName, 0x3ULL>("getPivotName"),
          GuardedInstanceMethod<&Workbook::PivotRowFieldOrder, 0x3ULL>("getPivotRowFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotSetLayout, 0x3ULL, 0x4ULL>("setPivotReportLayout"),
          GuardedInstanceMethod<&Workbook::PivotSetName, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotName"),
          GuardedInstanceMethod<&Workbook::PivotSetRowFieldOrder, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setPivotRowFieldOrder"),
          GuardedInstanceMethod<&Workbook::Precedents, 0xFULL>("getPrecedents"),
          GuardedInstanceMethod<&Workbook::Recalc>("recalc"),
          GuardedInstanceMethod<&Workbook::RecalcParallel>("recalcParallel"),
          GuardedInstanceMethod<&Workbook::RemoveConditionalFormatAt, 0x3ULL>("removeCfRule"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlink, 0x7ULL>("removeHyperlink"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlinkAt, 0x3ULL>("removeHyperlinkAt"),
          GuardedInstanceMethod<&Workbook::RemoveMerge, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeMerge"),
          GuardedInstanceMethod<&Workbook::RemoveMergeAt, 0x3ULL>("removeMergeAt"),
          GuardedInstanceMethod<&Workbook::RemoveSheet, 0x1ULL>("removeSheet"),
          GuardedInstanceMethod<&Workbook::RemoveValidationAt, 0x3ULL>("removeValidation"),
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
          GuardedInstanceMethod<&Workbook::SetPinnedNow, 0x0ULL, 0x3FULL>("setPinnedNow"),
          GuardedInstanceMethod<&Workbook::SetCellXfIndex, 0xFULL>("setCellXfIndex"),
          GuardedInstanceMethod<&Workbook::ClearCellXfIndex, 0x7ULL>("clearCellXfIndex"),
          GuardedInstanceMethod<&Workbook::SetCellPhonetic, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhonetic"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticRuns, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhoneticRuns"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticProperties, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhoneticProperties"),
          GuardedInstanceMethod<&Workbook::SetCellShowPhonetic, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellShowPhonetic"),
          GuardedInstanceMethod<&Workbook::SetRangeXfIndex, 0x5ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setRangeXfIndex"),
          GuardedInstanceMethod<&Workbook::SetColumnHidden, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnHidden"),
          GuardedInstanceMethod<&Workbook::SetColumnOutline, 0xFULL>("setColumnOutline"),
          GuardedInstanceMethod<&Workbook::SetColumnCollapsed, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnCollapsed"),
          GuardedInstanceMethod<&Workbook::SetColumnStyle, 0xFULL>("setColumnStyle"),
          GuardedInstanceMethod<&Workbook::SetColumnWidth, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnWidth"),
          GuardedInstanceMethod<&Workbook::SetComment, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x18ULL>("setComment"),
          GuardedInstanceMethod<&Workbook::SetCommentVisible, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCommentVisible"),
          GuardedInstanceMethod<&Workbook::SetWorkbookProtection, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setWorkbookProtection"),
          GuardedInstanceMethod<&Workbook::AddProtectedRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("addProtectedRange"),
          GuardedInstanceMethod<&Workbook::UpdateProtectedRange, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("updateProtectedRange"),
          GuardedInstanceMethod<&Workbook::RemoveProtectedRange, 0x3ULL>("removeProtectedRange"),
          GuardedInstanceMethod<&Workbook::SetDefinedName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("setDefinedName"),
          GuardedInstanceMethod<&Workbook::SetDefinedNameScoped, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x3ULL>("setDefinedNameScoped"),
          GuardedInstanceMethod<&Workbook::SetDefinedNameAttrs, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setDefinedNameAttrs"),
          GuardedInstanceMethod<&Workbook::SetDefaultFont, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setDefaultFont"),
          GuardedInstanceMethod<&Workbook::SetFont, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setFont"),
          GuardedInstanceMethod<&Workbook::SetError, 0x7ULL, 0x8ULL>("setError"),
          GuardedInstanceMethod<&Workbook::SetExcelProfileId, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setExcelProfileId"),
          GuardedInstanceMethod<&Workbook::SetFileName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setFileName"),
          GuardedInstanceMethod<&Workbook::SetFormula, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setFormula"),
          GuardedInstanceMethod<&Workbook::SetIterative, 0x0ULL, 0x2ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setIterative"),
          GuardedInstanceMethod<&Workbook::SetIterativeProgress, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setIterativeProgress"),
          GuardedInstanceMethod<&Workbook::SetNumber, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setNumber"),
          GuardedInstanceMethod<&Workbook::SetRowAutoHeight, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowAutoHeight"),
          GuardedInstanceMethod<&Workbook::SetRowHeight, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowHeight"),
          GuardedInstanceMethod<&Workbook::SetRowHidden, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowHidden"),
          GuardedInstanceMethod<&Workbook::SetRowOutline, 0xFULL>("setRowOutline"),
          GuardedInstanceMethod<&Workbook::SetRowCollapsed, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setRowCollapsed"),
          GuardedInstanceMethod<&Workbook::SetRowStyle, 0xFULL>("setRowStyle"),
          GuardedInstanceMethod<&Workbook::RangeStyleBases, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("getRangeStyleBases"),
          GuardedInstanceMethod<&Workbook::RangeStyleBasesExcluding, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xEULL>("getRangeStyleBasesExcluding"),
          GuardedInstanceMethod<&Workbook::RemapRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0xEULL>("remapRangeStyle"),
          GuardedInstanceMethod<&Workbook::SnapshotRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("snapshotRangeStyle"),
          GuardedInstanceMethod<&Workbook::RestoreRangeStyle, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("restoreRangeStyle"),
          GuardedInstanceMethod<&Workbook::SetSheetFreeze, 0x7ULL>("setFreeze"),
          GuardedInstanceMethod<&Workbook::SetSheetProtection, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetProtection"),
          GuardedInstanceMethod<&Workbook::SetSheetRightToLeft, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setRightToLeft"),
          GuardedInstanceMethod<&Workbook::SetSheetShowGridLines, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setShowGridLines"),
          GuardedInstanceMethod<&Workbook::SetSheetShowRowColHeaders, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setShowRowColHeaders"),
          GuardedInstanceMethod<&Workbook::SetSheetShowZeros, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setShowZeros"),
          GuardedInstanceMethod<&Workbook::SetSheetShowFormulas, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setShowFormulas"),
          GuardedInstanceMethod<&Workbook::SetSheetShowOutlineSymbols, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setShowOutlineSymbols"),
          GuardedInstanceMethod<&Workbook::SetSheetTabHidden, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setTabHidden"),
          GuardedInstanceMethod<&Workbook::SetSheetPaneTopLeft, 0x7ULL>("setPaneTopLeft"),
          GuardedInstanceMethod<&Workbook::SetSheetSplit, 0x1ULL, 0x0ULL, 0x0ULL, 0x6ULL>("setSplit"),
          GuardedInstanceMethod<&Workbook::SetSheetSelection, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSelection"),
          GuardedInstanceMethod<&Workbook::SetSheetTopLeft, 0x7ULL>("setTopLeft"),
          GuardedInstanceMethod<&Workbook::SetSheetTabSelected, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setTabSelected"),
          GuardedInstanceMethod<&Workbook::SetSheetViewMode, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setViewMode"),
          GuardedInstanceMethod<&Workbook::SetSheetVisibility, 0x1ULL, 0x2ULL>("setVisibility"),
          GuardedInstanceMethod<&Workbook::SetSheetZoom, 0x3ULL>("setZoom"),
          GuardedInstanceMethod<&Workbook::SetText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setText"),
          GuardedInstanceMethod<&Workbook::SheetCount>("getSheetCount"),
          GuardedInstanceMethod<&Workbook::ActiveSheet>("getActiveSheet"),
          GuardedInstanceMethod<&Workbook::SetActiveSheet, 0x1ULL>("setActiveSheet"),
          GuardedInstanceMethod<&Workbook::SheetName, 0x1ULL>("getSheetName"),
          GuardedInstanceMethod<&Workbook::GetSheetIndex, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("getSheetIndex"),
          GuardedInstanceMethod<&Workbook::SpillInfo, 0x7ULL>("getSpillInfo"),
          GuardedInstanceMethod<&Workbook::ParseReference, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("parseReference"),
          GuardedInstanceMethod<&Workbook::FormatReference, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("formatReference"),
          GuardedInstanceMethod<&Workbook::GetArrayFormula, 0x7ULL>("getArrayFormula"),
          GuardedInstanceMethod<&Workbook::SetArrayFormula, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setArrayFormula"),
          GuardedInstanceMethod<&Workbook::TableAt, 0x1ULL>("getTable"),
          GuardedInstanceMethod<&Workbook::CreateTable, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addTable"),
          GuardedInstanceMethod<&Workbook::UpdateTable, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("updateTable"),
          GuardedInstanceMethod<&Workbook::RemoveTable, 0x1ULL>("removeTable"),
          GuardedInstanceMethod<&Workbook::RenameTableColumn, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("renameTableColumn"),
          GuardedInstanceMethod<&Workbook::SetTableColumnTotals, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setTableColumnTotals"),
          GuardedInstanceMethod<&Workbook::TableCount>("getTableCount"),
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

}  // namespace formulon_node
