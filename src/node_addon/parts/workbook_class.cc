// Copyright 2026 libraz. Licensed under the MIT License.
//
// `Workbook` ObjectWrap lifecycle and class registration. Holds the
// ctor / dtor, the shared argument extraction helpers, and the
// `DefineClass` table that wires every per-area method into the JS
// class. The per-area method bodies live in `parts/*.cc`.

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

Workbook::Workbook(const Napi::CallbackInfo& info) : Napi::ObjectWrap<Workbook>(info) {
  // Default constructor used by the static factories. They populate
  // `handle_` after construction via `wb->handle_ = ...`.
  //
  // External JS callers should NOT invoke `new Workbook()` directly;
  // the JS-side index.mjs only re-exports the static factories.
  (void)info;
}

Workbook::~Workbook() {
  if (handle_ != nullptr) {
    // If this workbook had the global progress callback installed
    // we tear that down too -- otherwise the dangling C-callback
    // pointer would be invoked on a destroyed handle if recalc is
    // somehow re-driven against the same workbook. The slot itself
    // keeps the JS function reference alive across the addon's
    // lifetime; only the per-handle registration is unwound here.
    (void)fm_workbook_set_iterative_progress(handle_, nullptr, nullptr);
    fm_workbook_destroy(handle_);
    handle_ = nullptr;
  }
}

uint32_t Workbook::ArgU32(const Napi::CallbackInfo& info, size_t idx) {
  if (idx >= info.Length()) {
    return 0;
  }
  return info[idx].ToNumber().Uint32Value();
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

Napi::Value Workbook::CreateDefault(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create(&wb->handle_);
  if (rc != 0) {
    // Even on failure return the wrapper; the caller can inspect
    // `lastErrorMessage()` and the next operation will fail with
    // `kBindingNullPointer`. This matches embind's behaviour.
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::CreateEmpty(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  fm_status_t rc = fm_workbook_create_empty(&wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

Napi::Value Workbook::LoadBytes(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Function ctor = GetClass(env);
  Napi::Object jsobj = ctor.New({});
  Workbook* wb = Napi::ObjectWrap<Workbook>::Unwrap(jsobj);
  if (info.Length() < 1 || !info[0].IsTypedArray()) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::TypedArray ta = info[0].As<Napi::TypedArray>();
  if (ta.TypedArrayType() != napi_uint8_array) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  Napi::Uint8Array u8 = ta.As<Napi::Uint8Array>();
  const uint8_t* data = u8.Data();
  const std::size_t len = u8.ElementLength();
  if (data == nullptr || len == 0) {
    wb->handle_ = nullptr;
    return jsobj;
  }
  fm_status_t rc = fm_workbook_load(data, len, &wb->handle_);
  if (rc != 0) {
    wb->handle_ = nullptr;
  }
  return jsobj;
}

// ---- Class registration ---------------------------------------------

Napi::Function Workbook::GetClass(Napi::Env env) {
  return DefineClass(
      env, "Workbook",
      {
          StaticMethod<&Workbook::CreateDefault>("createDefault"),
          StaticMethod<&Workbook::CreateEmpty>("createEmpty"),
          StaticMethod<&Workbook::LoadBytes>("loadBytes"),
          InstanceMethod<&Workbook::AddBorder>("addBorder"),
          InstanceMethod<&Workbook::AddConditionalFormat>("addConditionalFormat"),
          InstanceMethod<&Workbook::AddFill>("addFill"),
          InstanceMethod<&Workbook::AddFont>("addFont"),
          InstanceMethod<&Workbook::AddHyperlink>("addHyperlink"),
          InstanceMethod<&Workbook::AddMerge>("addMerge"),
          InstanceMethod<&Workbook::AddNumFmt>("addNumFmt"),
          InstanceMethod<&Workbook::AddSheet>("addSheet"),
          InstanceMethod<&Workbook::AddValidation>("addValidation"),
          InstanceMethod<&Workbook::AddXf>("addXf"),
          InstanceMethod<&Workbook::BorderCount>("borderCount"),
          InstanceMethod<&Workbook::CalcMode>("calcMode"),
          InstanceMethod<&Workbook::CanonicalizeFunctionName>("canonicalizeFunctionName"),
          InstanceMethod<&Workbook::CellAt>("cellAt"),
          InstanceMethod<&Workbook::CellCount>("cellCount"),
          InstanceMethod<&Workbook::CellStyleCount>("cellStyleCount"),
          InstanceMethod<&Workbook::CellStyleXfCount>("cellStyleXfCount"),
          InstanceMethod<&Workbook::ClearConditionalFormats>("clearConditionalFormats"),
          InstanceMethod<&Workbook::ClearHyperlinks>("clearHyperlinks"),
          InstanceMethod<&Workbook::ClearMerges>("clearMerges"),
          InstanceMethod<&Workbook::ClearValidations>("clearValidations"),
          InstanceMethod<&Workbook::DefinedNameAt>("definedNameAt"),
          InstanceMethod<&Workbook::DefinedNameCount>("definedNameCount"),
          InstanceMethod<&Workbook::DeleteCols>("deleteCols"),
          InstanceMethod<&Workbook::DeleteRows>("deleteRows"),
          InstanceMethod<&Workbook::Dependents>("dependents"),
          InstanceMethod<&Workbook::DxfCount>("dxfCount"),
          InstanceMethod<&Workbook::EvaluateCfRange>("evaluateCfRange"),
          InstanceMethod<&Workbook::ExcelProfileId>("excelProfileId"),
          InstanceMethod<&Workbook::FillCount>("fillCount"),
          InstanceMethod<&Workbook::FontCount>("fontCount"),
          InstanceMethod<&Workbook::FunctionMetadata>("functionMetadata"),
          InstanceMethod<&Workbook::FunctionNames>("functionNames"),
          InstanceMethod<&Workbook::GetBorder>("getBorder"),
          InstanceMethod<&Workbook::GetCellStyle>("getCellStyle"),
          InstanceMethod<&Workbook::GetCellStyleXf>("getCellStyleXf"),
          InstanceMethod<&Workbook::GetCellXf>("getCellXf"),
          InstanceMethod<&Workbook::GetCellXfIndex>("getCellXfIndex"),
          InstanceMethod<&Workbook::GetComment>("getComment"),
          InstanceMethod<&Workbook::GetConditionalFormats>("getConditionalFormats"),
          InstanceMethod<&Workbook::GetDxf>("getDxf"),
          InstanceMethod<&Workbook::GetExternalLinks>("getExternalLinks"),
          InstanceMethod<&Workbook::GetFill>("getFill"),
          InstanceMethod<&Workbook::GetFont>("getFont"),
          InstanceMethod<&Workbook::GetHyperlinks>("getHyperlinks"),
          InstanceMethod<&Workbook::GetLambdaText>("getLambdaText"),
          InstanceMethod<&Workbook::GetMerges>("getMerges"),
          InstanceMethod<&Workbook::GetNumFmt>("getNumFmt"),
          InstanceMethod<&Workbook::GetSheetColumns>("getSheetColumns"),
          InstanceMethod<&Workbook::GetSheetProtection>("getSheetProtection"),
          InstanceMethod<&Workbook::GetSheetRowOverrides>("getSheetRowOverrides"),
          InstanceMethod<&Workbook::GetSheetView>("getSheetView"),
          InstanceMethod<&Workbook::GetValidations>("getValidations"),
          InstanceMethod<&Workbook::GetValue>("getValue"),
          InstanceMethod<&Workbook::InsertCols>("insertCols"),
          InstanceMethod<&Workbook::InsertRows>("insertRows"),
          InstanceMethod<&Workbook::IsValid>("isValid"),
          InstanceMethod<&Workbook::LocalizeFunctionName>("localizeFunctionName"),
          InstanceMethod<&Workbook::MoveSheet>("moveSheet"),
          InstanceMethod<&Workbook::PartialRecalc>("partialRecalc"),
          InstanceMethod<&Workbook::PassthroughAt>("passthroughAt"),
          InstanceMethod<&Workbook::PassthroughCount>("passthroughCount"),
          InstanceMethod<&Workbook::PivotCacheCount>("pivotCacheCount"),
          InstanceMethod<&Workbook::PivotCacheCreate>("pivotCacheCreate"),
          InstanceMethod<&Workbook::PivotCacheFieldAdd>("pivotCacheFieldAdd"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBlank>("pivotCacheFieldAddSharedItemBlank"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemBool>("pivotCacheFieldAddSharedItemBool"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemError>("pivotCacheFieldAddSharedItemError"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemNumber>("pivotCacheFieldAddSharedItemNumber"),
          InstanceMethod<&Workbook::PivotCacheFieldAddSharedItemText>("pivotCacheFieldAddSharedItemText"),
          InstanceMethod<&Workbook::PivotCacheFieldClear>("pivotCacheFieldClear"),
          InstanceMethod<&Workbook::PivotCacheFieldClearSharedItems>("pivotCacheFieldClearSharedItems"),
          InstanceMethod<&Workbook::PivotCacheFieldCount>("pivotCacheFieldCount"),
          InstanceMethod<&Workbook::PivotCacheFieldName>("pivotCacheFieldName"),
          InstanceMethod<&Workbook::PivotCacheFieldSharedItemCount>("pivotCacheFieldSharedItemCount"),
          InstanceMethod<&Workbook::PivotCacheIdAt>("pivotCacheIdAt"),
          InstanceMethod<&Workbook::PivotCacheRecordAdd>("pivotCacheRecordAdd"),
          InstanceMethod<&Workbook::PivotCacheRecordClear>("pivotCacheRecordClear"),
          InstanceMethod<&Workbook::PivotCacheRecordCount>("pivotCacheRecordCount"),
          InstanceMethod<&Workbook::PivotCacheRecordSetBlank>("pivotCacheRecordSetBlank"),
          InstanceMethod<&Workbook::PivotCacheRecordSetBool>("pivotCacheRecordSetBool"),
          InstanceMethod<&Workbook::PivotCacheRecordSetError>("pivotCacheRecordSetError"),
          InstanceMethod<&Workbook::PivotCacheRecordSetNumber>("pivotCacheRecordSetNumber"),
          InstanceMethod<&Workbook::PivotCacheRecordSetText>("pivotCacheRecordSetText"),
          InstanceMethod<&Workbook::PivotCacheRemove>("pivotCacheRemove"),
          InstanceMethod<&Workbook::PivotCount>("pivotCount"),
          InstanceMethod<&Workbook::PivotCreate>("pivotCreate"),
          InstanceMethod<&Workbook::PivotDataFieldAdd>("pivotDataFieldAdd"),
          InstanceMethod<&Workbook::PivotDataFieldClear>("pivotDataFieldClear"),
          InstanceMethod<&Workbook::PivotDataFieldCount>("pivotDataFieldCount"),
          InstanceMethod<&Workbook::PivotDataFieldSet>("pivotDataFieldSet"),
          InstanceMethod<&Workbook::PivotFieldAdd>("pivotFieldAdd"),
          InstanceMethod<&Workbook::PivotFieldAddAggregation>("pivotFieldAddAggregation"),
          InstanceMethod<&Workbook::PivotFieldAddItem>("pivotFieldAddItem"),
          InstanceMethod<&Workbook::PivotFieldAddSubtotalFn>("pivotFieldAddSubtotalFn"),
          InstanceMethod<&Workbook::PivotFieldClear>("pivotFieldClear"),
          InstanceMethod<&Workbook::PivotFieldClearAggregations>("pivotFieldClearAggregations"),
          InstanceMethod<&Workbook::PivotFieldClearDateGroup>("pivotFieldClearDateGroup"),
          InstanceMethod<&Workbook::PivotFieldClearItems>("pivotFieldClearItems"),
          InstanceMethod<&Workbook::PivotFieldClearSubtotalFns>("pivotFieldClearSubtotalFns"),
          InstanceMethod<&Workbook::PivotFieldCount>("pivotFieldCount"),
          InstanceMethod<&Workbook::PivotFieldSetAxis>("pivotFieldSetAxis"),
          InstanceMethod<&Workbook::PivotFieldSetDateGroup>("pivotFieldSetDateGroup"),
          InstanceMethod<&Workbook::PivotFieldSetItemVisible>("pivotFieldSetItemVisible"),
          InstanceMethod<&Workbook::PivotFieldSetNumberFormat>("pivotFieldSetNumberFormat"),
          InstanceMethod<&Workbook::PivotFieldSetSort>("pivotFieldSetSort"),
          InstanceMethod<&Workbook::PivotFieldSetSubtotalTop>("pivotFieldSetSubtotalTop"),
          InstanceMethod<&Workbook::PivotFilterAdd>("pivotFilterAdd"),
          InstanceMethod<&Workbook::PivotFilterClear>("pivotFilterClear"),
          InstanceMethod<&Workbook::PivotFilterCount>("pivotFilterCount"),
          InstanceMethod<&Workbook::PivotFilterRemoveAt>("pivotFilterRemoveAt"),
          InstanceMethod<&Workbook::PivotLayout>("pivotLayout"),
          InstanceMethod<&Workbook::PivotRemove>("pivotRemove"),
          InstanceMethod<&Workbook::PivotSetAnchor>("pivotSetAnchor"),
          InstanceMethod<&Workbook::PivotSetColFieldOrder>("pivotSetColFieldOrder"),
          InstanceMethod<&Workbook::PivotSetGrandTotals>("pivotSetGrandTotals"),
          InstanceMethod<&Workbook::PivotSetName>("pivotSetName"),
          InstanceMethod<&Workbook::PivotSetRowFieldOrder>("pivotSetRowFieldOrder"),
          InstanceMethod<&Workbook::Precedents>("precedents"),
          InstanceMethod<&Workbook::Recalc>("recalc"),
          InstanceMethod<&Workbook::RemoveConditionalFormatAt>("removeConditionalFormatAt"),
          InstanceMethod<&Workbook::RemoveHyperlink>("removeHyperlink"),
          InstanceMethod<&Workbook::RemoveHyperlinkAt>("removeHyperlinkAt"),
          InstanceMethod<&Workbook::RemoveMerge>("removeMerge"),
          InstanceMethod<&Workbook::RemoveMergeAt>("removeMergeAt"),
          InstanceMethod<&Workbook::RemoveSheet>("removeSheet"),
          InstanceMethod<&Workbook::RemoveValidationAt>("removeValidationAt"),
          InstanceMethod<&Workbook::RenameSheet>("renameSheet"),
          InstanceMethod<&Workbook::Save>("save"),
          InstanceMethod<&Workbook::SetBlank>("setBlank"),
          InstanceMethod<&Workbook::SetBool>("setBool"),
          InstanceMethod<&Workbook::SetCalcMode>("setCalcMode"),
          InstanceMethod<&Workbook::SetCellXfIndex>("setCellXfIndex"),
          InstanceMethod<&Workbook::SetColumnHidden>("setColumnHidden"),
          InstanceMethod<&Workbook::SetColumnOutline>("setColumnOutline"),
          InstanceMethod<&Workbook::SetColumnWidth>("setColumnWidth"),
          InstanceMethod<&Workbook::SetComment>("setComment"),
          InstanceMethod<&Workbook::SetDefinedName>("setDefinedName"),
          InstanceMethod<&Workbook::SetDefinedNameScoped>("setDefinedNameScoped"),
          InstanceMethod<&Workbook::SetError>("setError"),
          InstanceMethod<&Workbook::SetExcelProfileId>("setExcelProfileId"),
          InstanceMethod<&Workbook::SetFormula>("setFormula"),
          InstanceMethod<&Workbook::SetIterative>("setIterative"),
          InstanceMethod<&Workbook::SetIterativeProgress>("setIterativeProgress"),
          InstanceMethod<&Workbook::SetNumber>("setNumber"),
          InstanceMethod<&Workbook::SetRowHeight>("setRowHeight"),
          InstanceMethod<&Workbook::SetRowHidden>("setRowHidden"),
          InstanceMethod<&Workbook::SetRowOutline>("setRowOutline"),
          InstanceMethod<&Workbook::SetSheetFreeze>("setSheetFreeze"),
          InstanceMethod<&Workbook::SetSheetProtection>("setSheetProtection"),
          InstanceMethod<&Workbook::SetSheetTabHidden>("setSheetTabHidden"),
          InstanceMethod<&Workbook::SetSheetZoom>("setSheetZoom"),
          InstanceMethod<&Workbook::SetText>("setText"),
          InstanceMethod<&Workbook::SheetCount>("sheetCount"),
          InstanceMethod<&Workbook::SheetName>("sheetName"),
          InstanceMethod<&Workbook::SpillInfo>("spillInfo"),
          InstanceMethod<&Workbook::TableAt>("tableAt"),
          InstanceMethod<&Workbook::TableCount>("tableCount"),
          InstanceMethod<&Workbook::XfCount>("xfCount"),
      });
}

}  // namespace formulon_node
