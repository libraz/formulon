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

struct WorkbookClassState {
  Napi::FunctionReference constructor;
};

// Positional arguments are checked by the registration wrapper before a
// part-TU method gets to call a coercing N-API helper. The masks are indexed
// by the JavaScript argument position; an unset bit means that position is
// either an object/spec or is deliberately handled by the method itself.
struct PositionalArgSpec {
  std::uint64_t u32_mask;
  std::uint64_t i32_mask;
  std::uint64_t u64_mask;
  std::uint64_t double_mask;
  std::uint64_t string_mask;
  std::uint64_t optional_u32_mask;
  std::uint64_t optional_u64_mask;
  std::uint64_t optional_null_mask;
};

// A uint64 argument arrives as a Number, which is exact only up to 2^53 - 1;
// a larger value could not name the cursor the caller holds.
constexpr double kMaxSafeInteger = 9007199254740991.0;

bool RejectArgumentType(const Napi::CallbackInfo& info, std::size_t idx, const char* expected) {
  Napi::TypeError::New(info.Env(), expected).ThrowAsJavaScriptException();
  (void)idx;
  return false;
}

bool RejectArgumentRange(const Napi::CallbackInfo& info, std::size_t idx, const char* expected) {
  Napi::RangeError::New(info.Env(), expected).ThrowAsJavaScriptException();
  (void)idx;
  return false;
}

bool ValidatePositionalArgs(const Napi::CallbackInfo& info, const PositionalArgSpec& spec) {
  // A number that is an integer in [min_value, max_value]; absent passes.
  const auto validate_integer = [&](std::size_t idx, double min_value, double max_value, const char* range_message) {
    if (idx >= info.Length()) {
      return true;
    }
    const Napi::Value value = info[idx];
    if (!value.IsNumber()) {
      return RejectArgumentType(info, idx, "positional argument must be a number");
    }
    const double number = value.As<Napi::Number>().DoubleValue();
    if (!std::isfinite(number) || std::trunc(number) != number || number < min_value || number > max_value) {
      return RejectArgumentRange(info, idx, range_message);
    }
    return true;
  };
  constexpr double kU32Max = static_cast<double>(std::numeric_limits<std::uint32_t>::max());

  const std::size_t positional_count = info.Length() < 64 ? info.Length() : 64;
  for (std::size_t idx = 0; idx < positional_count; ++idx) {
    const std::uint64_t bit = std::uint64_t{1} << idx;
    if ((spec.u32_mask & bit) != 0 &&
        !validate_integer(idx, 0.0, kU32Max, "positional argument is outside uint32 range")) {
      return false;
    }
    if ((spec.optional_u32_mask & bit) != 0 && idx < info.Length()) {
      const Napi::Value value = info[idx];
      if (value.IsNull() && (spec.optional_null_mask & bit) == 0) {
        return RejectArgumentType(info, idx, "positional argument must be a number");
      }
      if (!value.IsUndefined() && !value.IsNull() &&
          !validate_integer(idx, 0.0, kU32Max, "positional argument is outside uint32 range")) {
        return false;
      }
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
    if ((spec.optional_u64_mask & bit) != 0 && idx < info.Length()) {
      const Napi::Value value = info[idx];
      if (value.IsNull() && (spec.optional_null_mask & bit) == 0) {
        return RejectArgumentType(info, idx, "positional argument must be a number");
      }
      if (!value.IsUndefined() && !value.IsNull() &&
          !validate_integer(idx, 0.0, kMaxSafeInteger, "positional argument is outside the safe-integer range")) {
        return false;
      }
    }
    if ((spec.double_mask & bit) != 0 && idx < info.Length() && !info[idx].IsNumber()) {
      return RejectArgumentType(info, idx, "positional argument must be a number");
    }
    if ((spec.string_mask & bit) != 0 && idx < info.Length() && !info[idx].IsString()) {
      return RejectArgumentType(info, idx, "positional argument must be a string");
    }
  }
  return true;
}

using InstanceMethodCallback = Napi::Value (Workbook::*)(const Napi::CallbackInfo&);

template <InstanceMethodCallback method, std::uint64_t u32_mask, std::uint64_t i32_mask, std::uint64_t u64_mask,
          std::uint64_t double_mask, std::uint64_t string_mask, std::uint64_t optional_u32_mask,
          std::uint64_t optional_u64_mask, std::uint64_t optional_null_mask>
struct GuardedMethodSpec {
  static constexpr PositionalArgSpec value{u32_mask,    i32_mask,          u64_mask,          double_mask,
                                           string_mask, optional_u32_mask, optional_u64_mask, optional_null_mask};
};

template <InstanceMethodCallback method, std::uint64_t u32_mask, std::uint64_t i32_mask, std::uint64_t u64_mask,
          std::uint64_t double_mask, std::uint64_t string_mask, std::uint64_t optional_u32_mask,
          std::uint64_t optional_u64_mask, std::uint64_t optional_null_mask>
napi_value GuardedInstanceMethodCallback(napi_env raw_env, napi_callback_info raw_info) {
  return Napi::details::WrapCallback(raw_env, [&] {
    Napi::CallbackInfo info(raw_env, raw_info);
    const auto* const spec = static_cast<const PositionalArgSpec*>(info.Data());
    if (spec == nullptr || !ValidatePositionalArgs(info, *spec)) {
      return info.Env().Undefined();
    }
    Workbook* const instance = Workbook::Unwrap(info.This().As<Napi::Object>());
    return instance == nullptr ? Napi::Value() : (instance->*method)(info);
  });
}

template <InstanceMethodCallback method, std::uint64_t u32_mask = 0, std::uint64_t i32_mask = 0,
          std::uint64_t u64_mask = 0, std::uint64_t double_mask = 0, std::uint64_t string_mask = 0,
          std::uint64_t optional_u32_mask = 0, std::uint64_t optional_u64_mask = 0,
          std::uint64_t optional_null_mask = 0>
Napi::ClassPropertyDescriptor<Workbook> GuardedInstanceMethod(const char* name) {
  napi_property_descriptor descriptor{};
  descriptor.utf8name = name;
  descriptor.method = &GuardedInstanceMethodCallback<method, u32_mask, i32_mask, u64_mask, double_mask, string_mask,
                                                     optional_u32_mask, optional_u64_mask, optional_null_mask>;
  descriptor.data = const_cast<void*>(
      static_cast<const void*>(&GuardedMethodSpec<method, u32_mask, i32_mask, u64_mask, double_mask, string_mask,
                                                  optional_u32_mask, optional_u64_mask, optional_null_mask>::value));
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

void Workbook::DestroyHandle(Napi::Env env) {
  if (handle_ != nullptr) {
    // Clear this workbook's C callback before destroying its user-data
    // wrapper, then release the JS callback reference below.
    (void)fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
  iterative_progress_callback_.Reset();
  // With the handle gone the measurement is zero, so this hands the
  // whole reported amount back to V8. Doing it after the destroy (rather
  // than before) keeps the accounting from ever going negative if the
  // release path is entered twice.
  SyncExternalMemory(env);
}

void Workbook::SyncExternalMemory(Napi::Env env) {
  int64_t current = 0;
  if (handle_ != nullptr) {
    size_t bytes = 0;
    if (fm_workbook_memory_usage(handle_, &bytes) == 0) {
      current = static_cast<int64_t>(bytes);
    }
  }
  const int64_t delta = current - reported_external_bytes_;
  if (delta == 0) {
    return;
  }
  reported_external_bytes_ = current;
  Napi::MemoryManagement::AdjustExternalMemory(env, delta);
}

uint32_t Workbook::ArgU32(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0;
  }
  return info[idx].ToNumber().Uint32Value();
}

int32_t Workbook::ArgI32(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0;
  }
  return info[idx].ToNumber().Int32Value();
}

double Workbook::ArgDouble(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0.0;
  }
  return info[idx].ToNumber().DoubleValue();
}

std::string Workbook::ArgString(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return std::string();
  }
  return info[idx].ToString().Utf8Value();
}

bool Workbook::ArgBool(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return false;
  }
  return info[idx].ToBoolean().Value();
}

Napi::Object Workbook::ArgObjectOrEmpty(const Napi::CallbackInfo& info, size_t idx) {
  return (info.Length() > idx && info[idx].IsObject()) ? info[idx].As<Napi::Object>() : Napi::Object::New(info.Env());
}

Napi::Value Workbook::InvokeRowColEdit(const Napi::CallbackInfo& info, RowColEditFn fn) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const uint32_t sheet = ArgU32(info, 0);
  const uint32_t index = ArgU32(info, 1);
  const uint32_t count = ArgU32(info, 2);
  return MakeStatus(env, fn(handle_, sheet, index, count));
}

// ---- Static factories -----------------------------------------------

Napi::Value Workbook::CreateWith(const Napi::CallbackInfo& info, CreateFn create) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = create(&wb->handle_);
  if (rc != 0) {
    // Even on failure return the wrapper; the caller can inspect
    // `lastErrorMessage()` and the next operation will fail with
    // `kBindingInvalidHandle`. This matches embind's behaviour.
    wb->handle_ = nullptr;
  }
  wb->SyncExternalMemory(env);
  return jsobj;
}

Napi::Value Workbook::CreateDefault(const Napi::CallbackInfo& info) {
  return CreateWith(info, &fm_workbook_create);
}

Napi::Value Workbook::CreateEmpty(const Napi::CallbackInfo& info) {
  return CreateWith(info, &fm_workbook_create_empty);
}

Napi::Value Workbook::LoadBytes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  if (info.Length() < 1 || !info[0].IsTypedArray()) {
    (void)fm_workbook_load(nullptr, 0, &wb->handle_);
    return jsobj;
  }
  Napi::TypedArray ta = info[0].As<Napi::TypedArray>();
  if (ta.TypedArrayType() != napi_uint8_array) {
    (void)fm_workbook_load(nullptr, 0, &wb->handle_);
    return jsobj;
  }
  Napi::Uint8Array u8 = ta.As<Napi::Uint8Array>();
  const uint8_t* data = u8.Data();
  const std::size_t len = u8.ElementLength();
  if (data == nullptr || len == 0) {
    (void)fm_workbook_load(nullptr, 0, &wb->handle_);
    return jsobj;
  }
  fm_status_t rc = fm_workbook_load(data, len, &wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  // A freshly loaded workbook is the one case where the footprint is
  // both large and known up front, so this is the report that matters
  // most for collection pressure.
  wb->SyncExternalMemory(env);
  return jsobj;
}

// ---- Class registration ---------------------------------------------

Napi::Function Workbook::GetClass(Napi::Env env) {
  if (WorkbookClassState* const state = env.GetInstanceData<WorkbookClassState>(); state != nullptr) {
    return state->constructor.Value();
  }

  // clang-format off
  Napi::Function constructor = DefineClass(
      env, "Workbook",
      {
          StaticMethod<&Workbook::CreateDefault>("createDefault"),
          StaticMethod<&Workbook::CreateEmpty>("createEmpty"),
          StaticMethod<&Workbook::LoadBytes>("loadBytes"),
          GuardedInstanceMethod<&Workbook::AddBorder>("addBorder"),
          GuardedInstanceMethod<&Workbook::AddConditionalFormat, 0x1ULL>("addConditionalFormat"),
          GuardedInstanceMethod<&Workbook::AddDxf>("addDxf"),
          GuardedInstanceMethod<&Workbook::AddFill>("addFill"),
          GuardedInstanceMethod<&Workbook::AddFont>("addFont"),
          GuardedInstanceMethod<&Workbook::AddHyperlink, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x78ULL>("addHyperlink"),
          GuardedInstanceMethod<&Workbook::AddHyperlinkRange, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1E0ULL>("addHyperlinkRange"),
          GuardedInstanceMethod<&Workbook::AddMerge, 0x1ULL>("addMerge"),
          GuardedInstanceMethod<&Workbook::AddNumFmt, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addNumFmt"),
          GuardedInstanceMethod<&Workbook::AddSheet, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("addSheet"),
          GuardedInstanceMethod<&Workbook::AddValidation, 0x1ULL>("addValidation"),
          GuardedInstanceMethod<&Workbook::AddPerson>("addPerson"),
          GuardedInstanceMethod<&Workbook::AddThreadedComment, 0x1ULL>("addThreadedComment"),
          GuardedInstanceMethod<&Workbook::ApplyAutoFilter, 0x1ULL>("applyAutoFilter"),
          GuardedInstanceMethod<&Workbook::ApplyTableAutoFilter, 0x1ULL>("applyTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::ClearAutoFilter, 0x1ULL>("clearAutoFilter"),
          GuardedInstanceMethod<&Workbook::ClearTableAutoFilter, 0x1ULL>("clearTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::EditThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL>("editThreadedComment"),
          GuardedInstanceMethod<&Workbook::EvaluateAutoFilter, 0x1ULL>("evaluateAutoFilter"),
          GuardedInstanceMethod<&Workbook::EvaluateTableAutoFilter, 0x1ULL>("evaluateTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetAutoFilter, 0x1ULL>("getAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetPersons>("getPersons"),
          GuardedInstanceMethod<&Workbook::GetTableAutoFilter, 0x1ULL>("getTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::GetThreadedComments, 0x1ULL>("getThreadedComments"),
          GuardedInstanceMethod<&Workbook::ListInvalidCells, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL, 0x2ULL, 0x6ULL>("listInvalidCells"),
          GuardedInstanceMethod<&Workbook::RemoveAutoFilter, 0x1ULL>("removeAutoFilter"),
          GuardedInstanceMethod<&Workbook::RemovePerson, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removePerson"),
          GuardedInstanceMethod<&Workbook::RemoveTableAutoFilter, 0x1ULL>("removeTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::RemoveThreadedComment, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("removeThreadedComment"),
          GuardedInstanceMethod<&Workbook::SetAutoFilter, 0x1ULL>("setAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetTableAutoFilter, 0x1ULL>("setTableAutoFilter"),
          GuardedInstanceMethod<&Workbook::SetThreadResolved, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setThreadResolved"),
          GuardedInstanceMethod<&Workbook::ValidateValue, 0x7ULL>("validateValue"),
          GuardedInstanceMethod<&Workbook::ProbeImage>("probeImage"),
          GuardedInstanceMethod<&Workbook::ListDrawingObjects, 0x1ULL>("listDrawingObjects"),
          GuardedInstanceMethod<&Workbook::GetImage, 0x3ULL>("getImage"),
          GuardedInstanceMethod<&Workbook::InsertImage, 0x1ULL>("insertImage"),
          GuardedInstanceMethod<&Workbook::RemoveImage, 0x3ULL>("removeImage"),
          GuardedInstanceMethod<&Workbook::SetImageAnchor, 0x3ULL>("setImageAnchor"),
          GuardedInstanceMethod<&Workbook::SetImageZOrder, 0x7ULL>("setImageZOrder"),
          GuardedInstanceMethod<&Workbook::SnapshotImage, 0x3ULL>("snapshotImage"),
          GuardedInstanceMethod<&Workbook::RestoreImage, 0x1ULL>("restoreImage"),
          GuardedInstanceMethod<&Workbook::AddXf>("addXf"),
          GuardedInstanceMethod<&Workbook::BorderCount>("borderCount"),
          GuardedInstanceMethod<&Workbook::CalcMode>("calcMode"),
          GuardedInstanceMethod<&Workbook::ClearRowHeight, 0x3ULL>("clearRowHeight"),
          GuardedInstanceMethod<&Workbook::ClearColumnWidth, 0x7ULL>("clearColumnWidth"),
          GuardedInstanceMethod<&Workbook::ColumnCharsToPt, 0x1ULL, 0x2ULL, 0x0ULL, 0x4ULL>("columnCharsToPt"),
          GuardedInstanceMethod<&Workbook::ColumnPtToChars, 0x1ULL, 0x2ULL, 0x0ULL, 0x4ULL>("columnPtToChars"),
          GuardedInstanceMethod<&Workbook::FormatValue, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("formatValue"),
          GuardedInstanceMethod<&Workbook::GetCellRectPt, 0x1ULL, 0x4ULL>("getCellRectPt"),
          GuardedInstanceMethod<&Workbook::GetCellsInRange, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL, 0x4ULL, 0xCULL>("getCellsInRange"),
          GuardedInstanceMethod<&Workbook::GetColumnWidthPt, 0x3ULL, 0x4ULL>("getColumnWidthPt"),
          GuardedInstanceMethod<&Workbook::GetDisplayText, 0x7ULL>("getDisplayText"),
          GuardedInstanceMethod<&Workbook::GetFormula, 0x7ULL>("getFormula"),
          GuardedInstanceMethod<&Workbook::GetFormulaR1C1, 0x7ULL>("getFormulaR1C1"),
          GuardedInstanceMethod<&Workbook::GetMergesInRange, 0x1ULL>("getMergesInRange"),
          GuardedInstanceMethod<&Workbook::GetRowHeightPt, 0x3ULL>("getRowHeightPt"),
          GuardedInstanceMethod<&Workbook::GetSheetFormatDefaults, 0x1ULL>("getSheetFormatDefaults"),
          GuardedInstanceMethod<&Workbook::GetWidthModel, 0x1ULL, 0x2ULL>("getWidthModel"),
          GuardedInstanceMethod<&Workbook::PinnedNow>("pinnedNow"),
          GuardedInstanceMethod<&Workbook::ClearPinnedNow>("clearPinnedNow"),
          GuardedInstanceMethod<&Workbook::CanonicalizeFunctionName, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL>("canonicalizeFunctionName"),
          GuardedInstanceMethod<&Workbook::CellAt, 0x3ULL>("cellAt"),
          GuardedInstanceMethod<&Workbook::CellCount, 0x1ULL>("cellCount"),
          GuardedInstanceMethod<&Workbook::CellStyleCount>("cellStyleCount"),
          GuardedInstanceMethod<&Workbook::CellStyleXfCount>("cellStyleXfCount"),
          GuardedInstanceMethod<&Workbook::ClearConditionalFormats, 0x1ULL>("clearConditionalFormats"),
          GuardedInstanceMethod<&Workbook::ClearHyperlinks, 0x1ULL>("clearHyperlinks"),
          GuardedInstanceMethod<&Workbook::ClearMerges, 0x1ULL>("clearMerges"),
          GuardedInstanceMethod<&Workbook::ClearValidations, 0x1ULL>("clearValidations"),
          GuardedInstanceMethod<&Workbook::DefinedNameAt, 0x1ULL>("definedNameAt"),
          GuardedInstanceMethod<&Workbook::DefinedNameCount>("definedNameCount"),
          GuardedInstanceMethod<&Workbook::DeleteCols, 0x7ULL>("deleteCols"),
          GuardedInstanceMethod<&Workbook::DeleteRows, 0x7ULL>("deleteRows"),
          GuardedInstanceMethod<&Workbook::Dependents, 0xFULL>("dependents"),
          GuardedInstanceMethod<&Workbook::Dispose>("dispose"),
          GuardedInstanceMethod<&Workbook::DxfCount>("dxfCount"),
          GuardedInstanceMethod<&Workbook::EvaluateCfRange, 0x1FULL, 0x0ULL, 0x0ULL, 0x20ULL>("evaluateCfRange"),
          GuardedInstanceMethod<&Workbook::EvaluateConditionalFormula, 0x1FULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x20ULL>("evaluateConditionalFormula"),
          GuardedInstanceMethod<&Workbook::EvaluateFormulaArray, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("evaluateFormulaArray"),
          GuardedInstanceMethod<&Workbook::EvaluateFormulaText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("evaluateFormulaText"),
          GuardedInstanceMethod<&Workbook::ExcelProfileId>("excelProfileId"),
          GuardedInstanceMethod<&Workbook::FillCount>("fillCount"),
          GuardedInstanceMethod<&Workbook::FontCount>("fontCount"),
          GuardedInstanceMethod<&Workbook::FunctionMetadata, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL>("functionMetadata"),
          GuardedInstanceMethod<&Workbook::FunctionNames>("functionNames"),
          GuardedInstanceMethod<&Workbook::GetBorder, 0x1ULL>("getBorder"),
          GuardedInstanceMethod<&Workbook::GetCellStyle, 0x1ULL>("getCellStyle"),
          GuardedInstanceMethod<&Workbook::GetCellStyleXf, 0x1ULL>("getCellStyleXf"),
          GuardedInstanceMethod<&Workbook::GetEffectiveStyle, 0x7ULL>("getEffectiveStyle"),
          GuardedInstanceMethod<&Workbook::GetTheme>("getTheme"),
          GuardedInstanceMethod<&Workbook::RemoveCellStyle, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("removeCellStyle"),
          GuardedInstanceMethod<&Workbook::ResolveColor, 0x0ULL, 0x2ULL>("resolveColor"),
          GuardedInstanceMethod<&Workbook::SetCellStyle>("setCellStyle"),
          GuardedInstanceMethod<&Workbook::SetThemeColors>("setThemeColors"),
          GuardedInstanceMethod<&Workbook::SetThemeFonts>("setThemeFonts"),
          GuardedInstanceMethod<&Workbook::ResetTheme>("resetTheme"),
          GuardedInstanceMethod<&Workbook::GetCellXf, 0x1ULL>("getCellXf"),
          GuardedInstanceMethod<&Workbook::GetCellXfIndex, 0x7ULL>("getCellXfIndex"),
          GuardedInstanceMethod<&Workbook::GetCellPhonetic, 0x7ULL>("getCellPhonetic"),
          GuardedInstanceMethod<&Workbook::GetCellPhoneticRuns, 0x7ULL>("getCellPhoneticRuns"),
          GuardedInstanceMethod<&Workbook::GetCellPhoneticProperties, 0x7ULL>("getCellPhoneticProperties"),
          GuardedInstanceMethod<&Workbook::GetComment, 0x7ULL>("getComment"),
          GuardedInstanceMethod<&Workbook::GetCommentResult, 0x7ULL>("getCommentResult"),
          GuardedInstanceMethod<&Workbook::GetComments, 0x1ULL>("getComments"),
          GuardedInstanceMethod<&Workbook::GetConditionalFormats, 0x1ULL>("getConditionalFormats"),
          GuardedInstanceMethod<&Workbook::GetDxf, 0x1ULL>("getDxf"),
          GuardedInstanceMethod<&Workbook::GetExternalLinks>("getExternalLinks"),
          GuardedInstanceMethod<&Workbook::GetFill, 0x1ULL>("getFill"),
          GuardedInstanceMethod<&Workbook::GetFont, 0x1ULL>("getFont"),
          GuardedInstanceMethod<&Workbook::GetHyperlinks, 0x1ULL>("getHyperlinks"),
          GuardedInstanceMethod<&Workbook::GetIterative>("getIterative"),
          GuardedInstanceMethod<&Workbook::GetLambdaText, 0x7ULL>("getLambdaText"),
          GuardedInstanceMethod<&Workbook::GetMerges, 0x1ULL>("getMerges"),
          GuardedInstanceMethod<&Workbook::GetNumFmt, 0x1ULL>("getNumFmt"),
          GuardedInstanceMethod<&Workbook::GetSheetColumns, 0x1ULL>("getSheetColumns"),
          GuardedInstanceMethod<&Workbook::GetSheetProtection, 0x1ULL>("getSheetProtection"),
          GuardedInstanceMethod<&Workbook::GetSheetRowOverrides, 0x1ULL>("getSheetRowOverrides"),
          GuardedInstanceMethod<&Workbook::GetSheetView, 0x1ULL>("getSheetView"),
          GuardedInstanceMethod<&Workbook::GetValidations, 0x1ULL>("getValidations"),
          GuardedInstanceMethod<&Workbook::GetValue, 0x7ULL>("getValue"),
          GuardedInstanceMethod<&Workbook::InsertCols, 0x7ULL>("insertCols"),
          GuardedInstanceMethod<&Workbook::InsertRows, 0x7ULL>("insertRows"),
          GuardedInstanceMethod<&Workbook::IsValid>("isValid"),
          GuardedInstanceMethod<&Workbook::LocalizeFunctionName, 0x0ULL, 0x2ULL, 0x0ULL, 0x0ULL, 0x1ULL>("localizeFunctionName"),
          GuardedInstanceMethod<&Workbook::MemoryUsage>("memoryUsage"),
          GuardedInstanceMethod<&Workbook::MoveSheet, 0x3ULL>("moveSheet"),
          GuardedInstanceMethod<&Workbook::PartialRecalc>("partialRecalc"),
          GuardedInstanceMethod<&Workbook::Paginate, 0x1ULL>("paginate"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetupXml, 0x1ULL>("getSheetPageSetupXml"),
          GuardedInstanceMethod<&Workbook::SetSheetFormatDefaults, 0x1ULL>("setSheetFormatDefaults"),
          GuardedInstanceMethod<&Workbook::SetSheetPageSetupXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPageSetupXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMarginsXml, 0x1ULL>("getSheetPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPageMarginsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPageMarginsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintOptionsXml, 0x1ULL>("getSheetPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintOptionsXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPrintOptionsXml"),
          GuardedInstanceMethod<&Workbook::GetSheetHeaderFooterXml, 0x1ULL>("getSheetHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::SetSheetHeaderFooterXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetHeaderFooterXml"),
          GuardedInstanceMethod<&Workbook::GetSheetSheetPrXml, 0x1ULL>("getSheetSheetPrXml"),
          GuardedInstanceMethod<&Workbook::SetSheetSheetPrXml, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetSheetPrXml"),
          GuardedInstanceMethod<&Workbook::SetSheetFitToPage, 0x1ULL>("setSheetFitToPage"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintArea, 0x1ULL>("getSheetPrintArea"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintArea, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetPrintArea"),
          GuardedInstanceMethod<&Workbook::GetSheetPrintTitles, 0x1ULL>("getSheetPrintTitles"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintTitles, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x6ULL>("setSheetPrintTitles"),
          GuardedInstanceMethod<&Workbook::AddSheetRowBreak, 0x3ULL>("addSheetRowBreak"),
          GuardedInstanceMethod<&Workbook::AddSheetColBreak, 0x3ULL>("addSheetColBreak"),
          GuardedInstanceMethod<&Workbook::RemoveSheetRowBreak, 0x3ULL>("removeSheetRowBreak"),
          GuardedInstanceMethod<&Workbook::RemoveSheetColBreak, 0x3ULL>("removeSheetColBreak"),
          GuardedInstanceMethod<&Workbook::ClearSheetBreaks, 0x1ULL>("clearSheetBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetRowBreaks, 0x1ULL>("getSheetRowBreaks"),
          GuardedInstanceMethod<&Workbook::GetSheetColBreaks, 0x1ULL>("getSheetColBreaks"),
          GuardedInstanceMethod<&Workbook::SetSheetPageSetup, 0x1ULL>("setSheetPageSetup"),
          GuardedInstanceMethod<&Workbook::SetSheetPageMargins, 0x1ULL>("setSheetPageMargins"),
          GuardedInstanceMethod<&Workbook::SetSheetPrintOptions, 0x1ULL>("setSheetPrintOptions"),
          GuardedInstanceMethod<&Workbook::SetSheetHeaderFooter, 0x1ULL>("setSheetHeaderFooter"),
          GuardedInstanceMethod<&Workbook::GetSheetPageSetup, 0x1ULL>("getSheetPageSetup"),
          GuardedInstanceMethod<&Workbook::GetSheetPageMargins, 0x1ULL>("getSheetPageMargins"),
          GuardedInstanceMethod<&Workbook::PassthroughAt, 0x1ULL>("passthroughAt"),
          GuardedInstanceMethod<&Workbook::PassthroughCount>("passthroughCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheCount>("pivotCacheCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheCreate, 0x1ULL>("pivotCacheCreate"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAdd, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("pivotCacheFieldAdd"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBlank, 0x3ULL>("pivotCacheFieldAddSharedItemBlank"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBool, 0x3ULL>("pivotCacheFieldAddSharedItemBool"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemError, 0x3ULL, 0x4ULL>("pivotCacheFieldAddSharedItemError"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemNumber, 0x3ULL, 0x0ULL, 0x0ULL, 0x4ULL>("pivotCacheFieldAddSharedItemNumber"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldAddSharedItemText, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("pivotCacheFieldAddSharedItemText"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldClear, 0x1ULL>("pivotCacheFieldClear"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldClearSharedItems, 0x3ULL>("pivotCacheFieldClearSharedItems"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldCount, 0x1ULL>("pivotCacheFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldName, 0x3ULL>("pivotCacheFieldName"),
          GuardedInstanceMethod<&Workbook::PivotCacheFieldSharedItemCount, 0x3ULL>("pivotCacheFieldSharedItemCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheGetWorksheetSource, 0x1ULL>("pivotCacheGetWorksheetSource"),
          GuardedInstanceMethod<&Workbook::PivotCacheIdAt, 0x1ULL>("pivotCacheIdAt"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordAdd, 0x1ULL>("pivotCacheRecordAdd"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordClear, 0x1ULL>("pivotCacheRecordClear"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordCount, 0x1ULL>("pivotCacheRecordCount"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetBlank, 0x7ULL>("pivotCacheRecordSetBlank"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetBool, 0x7ULL>("pivotCacheRecordSetBool"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetError, 0x7ULL, 0x8ULL>("pivotCacheRecordSetError"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetNumber, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("pivotCacheRecordSetNumber"),
          GuardedInstanceMethod<&Workbook::PivotCacheRecordSetText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("pivotCacheRecordSetText"),
          GuardedInstanceMethod<&Workbook::PivotCacheRemove, 0x1ULL>("pivotCacheRemove"),
          GuardedInstanceMethod<&Workbook::PivotCacheSetWorksheetSource, 0x1ULL>("pivotCacheSetWorksheetSource"),
          GuardedInstanceMethod<&Workbook::PivotCount, 0x1ULL>("pivotCount"),
          GuardedInstanceMethod<&Workbook::PivotCreate, 0x1DULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("pivotCreate"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldAdd, 0x3ULL>("pivotDataFieldAdd"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldClear, 0x3ULL>("pivotDataFieldClear"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldCount, 0x3ULL>("pivotDataFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotDataFieldSet, 0x7ULL>("pivotDataFieldSet"),
          GuardedInstanceMethod<&Workbook::PivotFieldAdd, 0x3ULL>("pivotFieldAdd"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddItem, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("pivotFieldAddItem"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddItemAt, 0xFULL>("pivotFieldAddItemAt"),
          GuardedInstanceMethod<&Workbook::PivotFieldAddSubtotalFn, 0x7ULL, 0x8ULL>("pivotFieldAddSubtotalFn"),
          GuardedInstanceMethod<&Workbook::PivotFieldClear, 0x3ULL>("pivotFieldClear"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearDateGroup, 0x7ULL>("pivotFieldClearDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearItems, 0x7ULL>("pivotFieldClearItems"),
          GuardedInstanceMethod<&Workbook::PivotFieldClearSubtotalFns, 0x7ULL>("pivotFieldClearSubtotalFns"),
          GuardedInstanceMethod<&Workbook::PivotFieldCount, 0x3ULL>("pivotFieldCount"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetAxis, 0x7ULL, 0x8ULL>("pivotFieldSetAxis"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetDateGroup, 0x87ULL, 0x78ULL, 0x0ULL, 0x300ULL>("pivotFieldSetDateGroup"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetItemVisible, 0xFULL>("pivotFieldSetItemVisible"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetNumberFormat, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("pivotFieldSetNumberFormat"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetSort, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x10ULL>("pivotFieldSetSort"),
          GuardedInstanceMethod<&Workbook::PivotFieldSetSubtotalTop, 0x7ULL>("pivotFieldSetSubtotalTop"),
          GuardedInstanceMethod<&Workbook::PivotFilterAdd, 0x3ULL>("pivotFilterAdd"),
          GuardedInstanceMethod<&Workbook::PivotFilterAt, 0x7ULL>("pivotFilterAt"),
          GuardedInstanceMethod<&Workbook::PivotFilterClear, 0x3ULL>("pivotFilterClear"),
          GuardedInstanceMethod<&Workbook::PivotFilterCount, 0x3ULL>("pivotFilterCount"),
          GuardedInstanceMethod<&Workbook::PivotFilterRemoveAt, 0x7ULL>("pivotFilterRemoveAt"),
          GuardedInstanceMethod<&Workbook::PivotLayout, 0x3ULL>("pivotLayout"),
          GuardedInstanceMethod<&Workbook::PivotRemove, 0x3ULL>("pivotRemove"),
          GuardedInstanceMethod<&Workbook::PivotSetAnchor, 0x3FULL>("pivotSetAnchor"),
          GuardedInstanceMethod<&Workbook::PivotSetColFieldOrder, 0x3ULL>("pivotSetColFieldOrder"),
          GuardedInstanceMethod<&Workbook::PivotSetGrandTotals, 0x3ULL>("pivotSetGrandTotals"),
          GuardedInstanceMethod<&Workbook::PivotGetLayout, 0x3ULL>("pivotGetLayout"),
          GuardedInstanceMethod<&Workbook::PivotSetLayout, 0x3ULL, 0x4ULL>("pivotSetLayout"),
          GuardedInstanceMethod<&Workbook::PivotSetName, 0x3ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x4ULL>("pivotSetName"),
          GuardedInstanceMethod<&Workbook::PivotSetRowFieldOrder, 0x3ULL>("pivotSetRowFieldOrder"),
          GuardedInstanceMethod<&Workbook::Precedents, 0xFULL>("precedents"),
          GuardedInstanceMethod<&Workbook::Recalc>("recalc"),
          GuardedInstanceMethod<&Workbook::RecalcParallel>("recalcParallel"),
          GuardedInstanceMethod<&Workbook::RemoveConditionalFormatAt, 0x3ULL>("removeConditionalFormatAt"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlink, 0x7ULL>("removeHyperlink"),
          GuardedInstanceMethod<&Workbook::RemoveHyperlinkAt, 0x3ULL>("removeHyperlinkAt"),
          GuardedInstanceMethod<&Workbook::RemoveMerge, 0x1ULL>("removeMerge"),
          GuardedInstanceMethod<&Workbook::RemoveMergeAt, 0x3ULL>("removeMergeAt"),
          GuardedInstanceMethod<&Workbook::RemoveSheet, 0x1ULL>("removeSheet"),
          GuardedInstanceMethod<&Workbook::RemoveValidationAt, 0x3ULL>("removeValidationAt"),
          GuardedInstanceMethod<&Workbook::RenameSheet, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("renameSheet"),
          GuardedInstanceMethod<&Workbook::Save>("save"),
          GuardedInstanceMethod<&Workbook::SaveAs, 0x0ULL, 0x1ULL>("saveAs"),
          GuardedInstanceMethod<&Workbook::SaveWithDiagnostics, 0x0ULL, 0x1ULL>("saveWithDiagnostics"),
          GuardedInstanceMethod<&Workbook::ReadDiagnostics>("readDiagnostics"),
          GuardedInstanceMethod<&Workbook::SetBlank, 0x7ULL>("setBlank"),
          GuardedInstanceMethod<&Workbook::SetBool, 0x7ULL>("setBool"),
          GuardedInstanceMethod<&Workbook::SetCalcMode, 0x0ULL, 0x1ULL>("setCalcMode"),
          GuardedInstanceMethod<&Workbook::SetPinnedNow, 0x0ULL, 0x3FULL>("setPinnedNow"),
          GuardedInstanceMethod<&Workbook::SetCellXfIndex, 0xFULL>("setCellXfIndex"),
          GuardedInstanceMethod<&Workbook::SetCellPhonetic, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setCellPhonetic"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticRuns, 0x7ULL>("setCellPhoneticRuns"),
          GuardedInstanceMethod<&Workbook::SetCellPhoneticProperties, 0x7ULL>("setCellPhoneticProperties"),
          GuardedInstanceMethod<&Workbook::SetRangeXfIndex, 0x3FULL>("setRangeXfIndex"),
          GuardedInstanceMethod<&Workbook::SetColumnHidden, 0x7ULL>("setColumnHidden"),
          GuardedInstanceMethod<&Workbook::SetColumnOutline, 0xFULL>("setColumnOutline"),
          GuardedInstanceMethod<&Workbook::SetColumnWidth, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setColumnWidth"),
          GuardedInstanceMethod<&Workbook::SetComment, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x18ULL>("setComment"),
          GuardedInstanceMethod<&Workbook::SetDefinedName, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x3ULL>("setDefinedName"),
          GuardedInstanceMethod<&Workbook::SetDefinedNameScoped, 0x0ULL, 0x4ULL, 0x0ULL, 0x0ULL, 0x3ULL>("setDefinedNameScoped"),
          GuardedInstanceMethod<&Workbook::SetDefaultFont>("setDefaultFont"),
          GuardedInstanceMethod<&Workbook::SetFont, 0x1ULL>("setFont"),
          GuardedInstanceMethod<&Workbook::SetError, 0x7ULL, 0x8ULL>("setError"),
          GuardedInstanceMethod<&Workbook::SetExcelProfileId, 0x0ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x1ULL>("setExcelProfileId"),
          GuardedInstanceMethod<&Workbook::SetFormula, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setFormula"),
          GuardedInstanceMethod<&Workbook::SetIterative, 0x2ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setIterative"),
          GuardedInstanceMethod<&Workbook::SetIterativeProgress>("setIterativeProgress"),
          GuardedInstanceMethod<&Workbook::SetNumber, 0x7ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setNumber"),
          GuardedInstanceMethod<&Workbook::SetRowHeight, 0x3ULL, 0x0ULL, 0x0ULL, 0x4ULL>("setRowHeight"),
          GuardedInstanceMethod<&Workbook::SetRowHidden, 0x3ULL>("setRowHidden"),
          GuardedInstanceMethod<&Workbook::SetRowOutline, 0x7ULL>("setRowOutline"),
          GuardedInstanceMethod<&Workbook::SetSheetFreeze, 0x7ULL>("setSheetFreeze"),
          GuardedInstanceMethod<&Workbook::SetSheetProtection, 0x1ULL>("setSheetProtection"),
          GuardedInstanceMethod<&Workbook::SetSheetRightToLeft, 0x1ULL>("setSheetRightToLeft"),
          GuardedInstanceMethod<&Workbook::SetSheetShowGridLines, 0x1ULL>("setSheetShowGridLines"),
          GuardedInstanceMethod<&Workbook::SetSheetShowRowColHeaders, 0x1ULL>("setSheetShowRowColHeaders"),
          GuardedInstanceMethod<&Workbook::SetSheetShowZeros, 0x1ULL>("setSheetShowZeros"),
          GuardedInstanceMethod<&Workbook::SetSheetTabHidden, 0x1ULL>("setSheetTabHidden"),
          GuardedInstanceMethod<&Workbook::SetSheetTabSelected, 0x1ULL>("setSheetTabSelected"),
          GuardedInstanceMethod<&Workbook::SetSheetViewMode, 0x1ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x2ULL>("setSheetViewMode"),
          GuardedInstanceMethod<&Workbook::SetSheetVisibility, 0x1ULL, 0x2ULL>("setSheetVisibility"),
          GuardedInstanceMethod<&Workbook::SetSheetZoom, 0x3ULL>("setSheetZoom"),
          GuardedInstanceMethod<&Workbook::SetText, 0x7ULL, 0x0ULL, 0x0ULL, 0x0ULL, 0x8ULL>("setText"),
          GuardedInstanceMethod<&Workbook::SheetCount>("sheetCount"),
          GuardedInstanceMethod<&Workbook::SheetName, 0x1ULL>("sheetName"),
          GuardedInstanceMethod<&Workbook::SpillInfo, 0x7ULL>("spillInfo"),
          GuardedInstanceMethod<&Workbook::TableAt, 0x1ULL>("tableAt"),
          GuardedInstanceMethod<&Workbook::TableCount>("tableCount"),
          GuardedInstanceMethod<&Workbook::XfCount>("xfCount"),
      });
  // clang-format on
  auto* const state = new WorkbookClassState{};
  state->constructor = Napi::Persistent(constructor);
  env.SetInstanceData(state);
  return constructor;
}

}  // namespace formulon_node
