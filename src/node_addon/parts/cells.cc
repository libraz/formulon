// Workbook cell bindings: cell mutation and read (including phonetic
// metadata), ad-hoc formula evaluation, and lambda text.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "node_addon/parts/workbook_class.h"

namespace formulon_node {

namespace {

struct CellAddressArgs {
  std::size_t sheet;
  uint32_t row;
  uint32_t col;
};

CellAddressArgs ReadCellAddress(const Napi::CallbackInfo& info) {
  return {static_cast<std::size_t>(Workbook::ArgU32(info, 0)), Workbook::ArgU32(info, 1), Workbook::ArgU32(info, 2)};
}

// Shared body of the `(sheet, row, col, string)` cell setters.
using CellTextFn = fm_status_t (*)(fm_workbook_t*, size_t, uint32_t, uint32_t, const char*);

Napi::Value InvokeCellText(const Napi::CallbackInfo& info, fm_workbook_t* handle, CellTextFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string text = Workbook::ArgString(info, 3);
  fm_status_t rc = fn(handle, address.sheet, address.row, address.col, text.c_str());
  return ValueOrThrow(env, rc, env.Undefined());
}

}  // namespace

// ---- Cell mutation --------------------------------------------------

Napi::Value Workbook::SetNumber(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_number);
}

Napi::Value Workbook::SetBool(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CellAddressArgs address = ReadCellAddress(info);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_bool(handle, address.sheet, address.row, address.col, value);
  });
}

Napi::Value Workbook::SetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // The registration guard requires `errorCode`: a zero default would
  // silently write `#NULL!`.
  const CellAddressArgs address = ReadCellAddress(info);
  const int32_t error_code = ArgI32(info, 3);
  fm_status_t rc =
      fm_sheet_set_error(handle_, address.sheet, address.row, address.col, static_cast<fm_error_code_t>(error_code));
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::SetText(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_sheet_set_text);
}

Napi::Value Workbook::SetCellPhonetic(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_sheet_set_cell_phonetic);
}

Napi::Value Workbook::SetCellPhoneticRuns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  Napi::Array runs;
  if (!ArrayArg(info, 3, &runs)) {
    ThrowArgumentTypeError(env, "runs", "must be an array");
    return env.Undefined();
  }
  const uint32_t count = runs.Length();
  CheckedSpecReader reader(env);
  // Two passes so no `c_str()` is taken before `texts` has finished growing.
  std::vector<std::string> texts;
  texts.reserve(count);
  std::vector<fm_phonetic_run_t> records;
  records.reserve(count);
  for (uint32_t i = 0; i < count; ++i) {
    Napi::Value element;
    if (!reader.ArrayElement(runs, i, &element)) {
      if (!reader.ok()) {
        return env.Undefined();
      }
      Napi::TypeError::New(env, "Sheet.setCellPhoneticRuns runs[] must be an object").ThrowAsJavaScriptException();
      return env.Undefined();
    }
    Napi::Object run;
    if (!reader.Object(element, "run", &run)) {
      if (!reader.ok()) {
        return env.Undefined();
      }
      Napi::TypeError::New(env, "Sheet.setCellPhoneticRuns runs[] must be an object").ThrowAsJavaScriptException();
      return env.Undefined();
    }
    std::string text;
    (void)reader.String(run, "text", &text);
    texts.push_back(std::move(text));
    records.push_back(fm_phonetic_run_t{reader.U32(run, "start", 0U), reader.U32(run, "end", 0U), 0U, 0U, nullptr});
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_cell_phonetic_runs(handle_, address.sheet, address.row, address.col,
                                                   FM_TEXT_OFFSET_UNIT_UTF16, records.data(), records.size());
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::SetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // `null` resets the cell to the unstated triple; `stated` is read-only and ignored.
  const CellAddressArgs address = ReadCellAddress(info);
  CheckedSpecReader reader(env);
  Napi::Object props;
  const bool present = reader.Object(info[3], "", &props);
  fm_phonetic_properties_t record = FM_STRUCT_INIT(fm_phonetic_properties_t);
  if (present) {
    record.font_index = reader.U32(props, "fontIndex", 0U);
    record.type = reader.I32(props, "type", 0);
    record.alignment = reader.I32(props, "alignment", 0);
  }
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_sheet_set_cell_phonetic_properties(handle_, address.sheet, address.row, address.col,
                                                         present ? &record : nullptr);
  return ValueOrThrow(env, rc, env.Undefined());
}

Napi::Value Workbook::SetBlank(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_blank);
}

Napi::Value Workbook::SetFormula(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_sheet_set_formula);
}

// ---- Cell read ------------------------------------------------------

Napi::Value Workbook::GetValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_value_t v{};
  fm_status_t rc = fm_sheet_get_value(handle_, address.sheet, address.row, address.col, &v);
  return ValueOrThrow(env, rc, ValueToJs(env, v));
}

Napi::Value Workbook::GetCellPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  int32_t present = 0;
  const char* text = nullptr;
  fm_status_t rc = fm_sheet_get_cell_phonetic(handle_, address.sheet, address.row, address.col, &present, &text);
  return ValueOrThrow(env, rc, present != 0 ? JsString(env, text) : env.Null());
}

Napi::Value Workbook::GetCellPhoneticRunCount(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  std::size_t count = 0;
  const fm_status_t rc = fm_sheet_get_cell_phonetic_run_count(handle_, address.sheet, address.row, address.col, &count);
  return ValueOrThrow(env, rc, Napi::Number::New(env, static_cast<double>(count)));
}

Napi::Value Workbook::GetCellPhoneticRun(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_phonetic_run_t run{};
  const fm_status_t rc =
      fm_sheet_get_cell_phonetic_run_at(handle_, address.sheet, address.row, address.col, ArgU32(info, 3), &run);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object item = Napi::Object::New(env);
  item.Set("start", Napi::Number::New(env, run.start));
  item.Set("end", Napi::Number::New(env, run.end));
  item.Set("text", JsString(env, run.text));
  return item;
}

Napi::Value Workbook::GetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_phonetic_properties_t record = FM_STRUCT_INIT(fm_phonetic_properties_t);
  const fm_status_t rc =
      fm_sheet_get_cell_phonetic_properties(handle_, address.sheet, address.row, address.col, &record);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("fontIndex", Napi::Number::New(env, record.font_index));
  out.Set("type", Napi::Number::New(env, record.type));
  out.Set("alignment", Napi::Number::New(env, record.alignment));
  out.Set("stated", Napi::Boolean::New(env, record.stated != 0));
  return out;
}

Napi::Value Workbook::GetCellShowPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  int32_t show = 0;
  const fm_status_t rc = fm_sheet_get_cell_show_phonetic(handle_, address.sheet, address.row, address.col, &show);
  return ValueOrThrow(env, rc, JsBool(env, show));
}

Napi::Value Workbook::SetCellShowPhonetic(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CellAddressArgs address = ReadCellAddress(info);
    const int32_t value = ArgFlag(info, 3);
    if (info.Env().IsExceptionPending()) {
      return FM_STATUS_INVALID_ARGUMENT;
    }
    return fm_sheet_set_cell_show_phonetic(handle, address.sheet, address.row, address.col, value);
  });
}

Napi::Value Workbook::EvaluateFormulaArray(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string formula = ArgString(info, 3);

  fm_value_array_t* array = nullptr;
  fm_status_t rc =
      fm_sheet_evaluate_formula_array(handle_, address.sheet, address.row, address.col, formula.c_str(), &array);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }

  const Napi::Object out = ValueArrayToJs(env, array);
  fm_value_array_destroy(array);
  return out;
}

Napi::Value Workbook::EvaluateConditionalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const uint32_t anchor_row = ArgU32(info, 3);
  const uint32_t anchor_col = ArgU32(info, 4);
  const std::string formula = ArgString(info, 5);
  fm_value_t v{};
  fm_status_t rc = fm_sheet_evaluate_cf_formula(handle_, address.sheet, address.row, address.col, anchor_row,
                                                anchor_col, formula.c_str(), &v);
  return ValueOrThrow(env, rc, Napi::Boolean::New(env, v.u.boolean != 0));
}

Napi::Value Workbook::GetLambdaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  int32_t present = 0;
  const char* text = nullptr;
  fm_status_t rc = fm_sheet_get_lambda_text(handle_, address.sheet, address.row, address.col, &present, &text);
  return OptionalOrNull(env, rc, present, JsString(env, text));
}

// ---- Formula text, range enumeration, display text ------------------

namespace {

// Shared body of the A1 / R1C1 formula-text getters; `null` when the cell has no formula.
using FormulaTextFn = fm_status_t (*)(const fm_workbook_t*, size_t, uint32_t, uint32_t, int32_t*, const char**);

Napi::Value FormulaTextResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, FormulaTextFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* formula = nullptr;
  int32_t present = 0;
  const fm_status_t rc = fn(handle, address.sheet, address.row, address.col, &present, &formula);
  return OptionalOrNull(env, rc, present, JsString(env, formula));
}

// Builds `{ text, displayStatus }`.
Napi::Value DisplayResult(Napi::Env env, fm_status_t rc, const char* text, int32_t display_status) {
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("text", JsString(env, text));
  out.Set("displayStatus", Napi::Number::New(env, display_status));
  return out;
}

}  // namespace

Napi::Value Workbook::GetFormula(const Napi::CallbackInfo& info) {
  return FormulaTextResult(info, handle_, &fm_sheet_get_formula);
}

Napi::Value Workbook::GetFormulaR1C1(const Napi::CallbackInfo& info) {
  return FormulaTextResult(info, handle_, &fm_sheet_get_formula_r1c1);
}

namespace {

// Builds `{ cells, nextCursor }` from a cell-range page, which it consumes.
// `rc` is the status of the call that produced `page`.
Napi::Value CellsPageResult(Napi::Env env, fm_cell_range_t* page, fm_status_t rc) {
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Array cells = Napi::Array::New(env);
  size_t count = 0;
  rc = fm_cell_range_get_cell_count(page, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    rc = fm_cell_range_get_cell_at(page, i, &row, &col, &formula, &value);
    if (rc != 0) {
      break;
    }
    Napi::Object cell = Napi::Object::New(env);
    cell.Set("row", Napi::Number::New(env, row));
    cell.Set("col", Napi::Number::New(env, col));
    cell.Set("formula", formula != nullptr ? static_cast<Napi::Value>(Napi::String::New(env, formula)) : env.Null());
    cell.Set("value", ValueToJs(env, value));
    cells.Set(static_cast<uint32_t>(i), cell);
  }
  uint64_t next = UINT64_MAX;
  if (rc == 0) {
    rc = fm_cell_range_next_cursor(page, &next);
  }
  fm_cell_range_destroy(page);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("cells", cells);
  out.Set("nextCursor",
          next != UINT64_MAX ? Napi::Value(Napi::Number::New(env, static_cast<double>(next))) : env.Null());
  return out;
}

}  // namespace

Napi::Value Workbook::GetCellsInRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  // The guard checked `cursor` as a safe integer and `limit` as uint32 when present.
  const uint64_t cursor = static_cast<uint64_t>(NumberArg(info, 2, 0.0));
  const uint32_t limit = ArgU32(info, 3);
  CheckedSpecReader reader(env);
  fm_cell_rect_t range{};
  if (!CellRectArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  fm_cell_range_t* page = nullptr;
  fm_status_t rc = fm_sheet_get_cells_in_range(handle_, ArgU32(info, 0), &range, cursor, limit, &page);
  return CellsPageResult(env, page, rc);
}

Napi::Value Workbook::ListInvalidCells(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const uint64_t cursor = static_cast<uint64_t>(NumberArg(info, 1, 0.0));
  const uint32_t limit = ArgU32(info, 2);
  fm_cell_range_t* page = nullptr;
  const fm_status_t rc = fm_sheet_get_invalid_cells(handle_, ArgU32(info, 0), cursor, limit, &page);
  return CellsPageResult(env, page, rc);
}

Napi::Value Workbook::ValidateValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::string text;
  fm_value_t value{};
  if (!JsToValue(env, info[3], &value, &text)) {
    return env.Undefined();
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_validation_outcome_info_t outcome = FM_STRUCT_INIT(fm_validation_outcome_info_t);
  const fm_status_t rc = fm_sheet_validate_value(handle_, address.sheet, address.row, address.col, &value, &outcome);
  if (rc != 0) {
    return ThrowStatus(env, rc);
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("valid", Napi::Boolean::New(env, outcome.valid != 0));
  out.Set("ruleId", outcome.has_rule != 0 ? Napi::Value(JsString(env, outcome.rule_id)) : env.Null());
  out.Set("errorStyle", JsNumberOrNull(env, outcome.has_rule, static_cast<uint32_t>(outcome.error_style)));
  return out;
}

Napi::Value Workbook::GetDisplayText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* text = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc =
      fm_sheet_get_display_text(handle_, address.sheet, address.row, address.col, &text, &display_status);
  return DisplayResult(env, rc, text, display_status);
}

Napi::Value Workbook::FormatValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return ThrowInvalidHandle(env);
  }
  std::string text;
  fm_value_t value{};
  if (!JsToValue(env, info[0], &value, &text)) {
    return env.Undefined();
  }
  const std::string format_code = ArgString(info, 1);
  const char* out_text = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc = fm_workbook_format_value(handle_, &value, format_code.c_str(), &out_text, &display_status);
  return DisplayResult(env, rc, out_text, display_status);
}

}  // namespace formulon_node
