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
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string text = Workbook::ArgString(info, 3);
  fm_status_t rc = fn(handle, address.sheet, address.row, address.col, text.c_str());
  return MakeStatus(env, rc);
}

// Builds `{ status, rows: 0, cols: 0, cells: [] }` for a failed array evaluation.
Napi::Object EmptyFormulaArrayResult(Napi::Env env, Napi::Object status) {
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", status);
  out.Set("rows", Napi::Number::New(env, 0));
  out.Set("cols", Napi::Number::New(env, 0));
  out.Set("cells", Napi::Array::New(env, 0));
  return out;
}

}  // namespace

// ---- Cell mutation --------------------------------------------------

Napi::Value Workbook::SetNumber(const Napi::CallbackInfo& info) {
  return InvokeSheetCall(info, &fm_sheet_set_number);
}

Napi::Value Workbook::SetBool(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CellAddressArgs address = ReadCellAddress(info);
    return fm_sheet_set_bool(handle, address.sheet, address.row, address.col, ArgFlag(info, 3));
  });
}

Napi::Value Workbook::SetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  // The registration guard requires `errorCode`: a zero default would
  // silently write `#NULL!`.
  const CellAddressArgs address = ReadCellAddress(info);
  const int32_t error_code = ArgI32(info, 3);
  fm_status_t rc =
      fm_sheet_set_error(handle_, address.sheet, address.row, address.col, static_cast<fm_error_code_t>(error_code));
  return MakeStatus(env, rc);
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
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  Napi::Array runs;
  if (!ArrayArg(info, 3, &runs)) {
    return MakeBindingArgumentError(env, "setCellPhoneticRuns: `runs` must be an array of { sb, eb, text }");
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
      return MakeBindingArgumentError(env, "setCellPhoneticRuns: each run must be an object { sb, eb, text }");
    }
    Napi::Object run;
    if (!reader.Object(element, "run", &run)) {
      if (!reader.ok()) {
        return env.Undefined();
      }
      return MakeBindingArgumentError(env, "setCellPhoneticRuns: each run must be an object { sb, eb, text }");
    }
    std::string text;
    (void)reader.String(run, "text", &text);
    texts.push_back(std::move(text));
    records.push_back(fm_phonetic_run_t{reader.U32(run, "sb", 0U), reader.U32(run, "eb", 0U), nullptr});
  }
  for (uint32_t i = 0; i < count; ++i) {
    records[i].text = texts[i].c_str();
  }
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc =
      fm_sheet_set_cell_phonetic_runs(handle_, address.sheet, address.row, address.col, records.data(), records.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const Napi::Object props = ArgObject(info, 3);
  CheckedSpecReader reader(env);
  const uint32_t font_id = reader.U32(props, "fontId", 0U);
  const uint32_t type = reader.U32(props, "type", 0U);
  const uint32_t alignment = reader.U32(props, "alignment", 0U);
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc =
      fm_sheet_set_cell_phonetic_properties(handle_, address.sheet, address.row, address.col, font_id, type, alignment);
  return MakeStatus(env, rc);
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
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_value_t v{};
  fm_status_t rc = fm_sheet_get_value(handle_, address.sheet, address.row, address.col, &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetCellPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* text = nullptr;
  fm_status_t rc = fm_sheet_get_cell_phonetic(handle_, address.sheet, address.row, address.col, &text);
  return MakeStringResult(env, rc, text);
}

Napi::Value Workbook::GetCellPhoneticRunCount(const Napi::CallbackInfo& info) {
  const CellAddressArgs address = ReadCellAddress(info);
  std::size_t count = 0;
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_cell_phonetic_run_count(handle_, address.sheet, address.row,
                                                                                   address.col, &count)
                                            : kBindingInvalidHandle;
  return MakeNumberResult(info.Env(), rc, static_cast<double>(count));
}

Napi::Value Workbook::GetCellPhoneticRun(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const CellAddressArgs address = ReadCellAddress(info);
  fm_phonetic_run_t run{};
  const fm_status_t rc = handle_ != nullptr ? fm_sheet_get_cell_phonetic_run(handle_, address.sheet, address.row,
                                                                             address.col, ArgU32(info, 3), &run)
                                            : kBindingInvalidHandle;
  if (rc != 0) {
    run = fm_phonetic_run_t{};
  }
  Napi::Object item = Napi::Object::New(env);
  item.Set("sb", Napi::Number::New(env, run.sb));
  item.Set("eb", Napi::Number::New(env, run.eb));
  item.Set("text", JsString(env, run.text));
  return FinishItemResult(env, item, rc);
}

Napi::Value Workbook::GetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const CellAddressArgs address = ReadCellAddress(info);
  uint32_t font_id = 0;
  uint32_t type = 0;
  uint32_t alignment = 0;
  int32_t present = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_cell_phonetic_properties(handle_, address.sheet, address.row, address.col,
                                                                     &font_id, &type, &alignment)
                             : kBindingInvalidHandle;
  if (rc == 0) {
    (void)fm_sheet_has_cell_phonetic_properties(handle_, address.sheet, address.row, address.col, &present);
  }
  // The payload keys are declared unconditionally, so a failure reports the
  // defaults beside the status rather than dropping them.
  if (rc != 0) {
    font_id = 0;
    type = 0;
    alignment = 0;
  }
  Napi::Object out = Napi::Object::New(env);
  out.Set("fontId", Napi::Number::New(env, font_id));
  out.Set("type", Napi::Number::New(env, type));
  out.Set("alignment", Napi::Number::New(env, alignment));
  out.Set("present", Napi::Boolean::New(env, present != 0));
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::GetCellShowPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const CellAddressArgs address = ReadCellAddress(info);
  int32_t show = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_sheet_get_cell_show_phonetic(handle_, address.sheet, address.row, address.col, &show)
                             : kBindingInvalidHandle;
  return MakeBoolResult(env, rc, show);
}

Napi::Value Workbook::SetCellShowPhonetic(const Napi::CallbackInfo& info) {
  return StatusCall(info, [&](fm_workbook_t* handle) {
    const CellAddressArgs address = ReadCellAddress(info);
    return fm_sheet_set_cell_show_phonetic(handle, address.sheet, address.row, address.col, ArgFlag(info, 3));
  });
}

Napi::Value Workbook::EvaluateFormulaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string formula = ArgString(info, 3);
  fm_value_t v{};
  fm_status_t rc = fm_sheet_evaluate_formula(handle_, address.sheet, address.row, address.col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::EvaluateFormulaArray(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyFormulaArrayResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string formula = ArgString(info, 3);

  fm_value_array_t* array = nullptr;
  fm_status_t rc =
      fm_sheet_evaluate_formula_array(handle_, address.sheet, address.row, address.col, formula.c_str(), &array);
  if (rc != 0) {
    return EmptyFormulaArrayResult(env, MakeErrorStatus(env, rc));
  }

  // Build a rows x cols nested array of Value objects, reading each cell by
  // its row-major index (r * cols + c).
  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_value_array_get_dimensions(array, &rows, &cols);
  Napi::Array cells = Napi::Array::New(env, rows);
  for (uint32_t r = 0; r < rows; ++r) {
    Napi::Array js_row = Napi::Array::New(env, cols);
    for (uint32_t c = 0; c < cols; ++c) {
      fm_value_t v{};
      fm_value_array_get_value_at(array, static_cast<std::size_t>(r) * cols + c, &v);
      js_row.Set(c, TranslateValue(env, v));
    }
    cells.Set(r, js_row);
  }
  fm_value_array_destroy(array);

  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeOkStatus(env));
  out.Set("rows", Napi::Number::New(env, rows));
  out.Set("cols", Napi::Number::New(env, cols));
  out.Set("cells", cells);
  return out;
}

Napi::Value Workbook::EvaluateConditionalFormula(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const uint32_t anchor_row = ArgU32(info, 3);
  const uint32_t anchor_col = ArgU32(info, 4);
  const std::string formula = ArgString(info, 5);
  fm_value_t v{};
  fm_status_t rc = fm_sheet_evaluate_cf_formula(handle_, address.sheet, address.row, address.col, anchor_row,
                                                anchor_col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetLambdaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeOptionalResult(env, NullHandleError(env), env.Null());
  }
  const CellAddressArgs address = ReadCellAddress(info);
  int32_t present = 0;
  const char* text = nullptr;
  fm_status_t rc = fm_sheet_get_lambda_text(handle_, address.sheet, address.row, address.col, &present, &text);
  return MakeOptionalResult(env, rc, present, rc == 0 && present != 0 ? Napi::Value(JsString(env, text)) : env.Null());
}

// ---- Formula text, range enumeration, display text ------------------

namespace {

// Builds the `{ status, value }` OptionalResult; an empty C string is "no formula".
Napi::Object MakeFormulaResult(Napi::Env env, fm_status_t code, const char* formula) {
  const bool present = code == 0 && formula != nullptr && formula[0] != '\0';
  return MakeOptionalResult(env, code, present ? 1 : 0, present ? Napi::String::New(env, formula) : env.Null());
}

// Shared body of the A1 / R1C1 formula-text getters.
using FormulaTextFn = fm_status_t (*)(const fm_workbook_t*, size_t, uint32_t, uint32_t, const char**);

Napi::Value FormulaTextResult(const Napi::CallbackInfo& info, const fm_workbook_t* handle, FormulaTextFn fn) {
  Napi::Env env = info.Env();
  if (handle == nullptr) {
    return MakeFormulaResult(env, kBindingInvalidHandle, nullptr);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* formula = nullptr;
  const fm_status_t rc = fn(handle, address.sheet, address.row, address.col, &formula);
  return MakeFormulaResult(env, rc, formula);
}

// Builds `{ status, text, displayStatus }`; `text` is "" and the display
// status OK when the call failed.
Napi::Object MakeDisplayResult(Napi::Env env, fm_status_t code, const char* text, int32_t display_status) {
  const bool ok = code == 0;
  Napi::Object out = Napi::Object::New(env);
  out.Set("status", MakeStatus(env, code));
  out.Set("text", JsString(env, ok ? text : nullptr));
  out.Set("displayStatus", Napi::Number::New(env, ok ? display_status : 0));
  return out;
}

}  // namespace

Napi::Value Workbook::GetFormula(const Napi::CallbackInfo& info) {
  return FormulaTextResult(info, handle_, &fm_sheet_get_formula);
}

Napi::Value Workbook::GetFormulaR1C1(const Napi::CallbackInfo& info) {
  return FormulaTextResult(info, handle_, &fm_sheet_get_formula_r1c1);
}

// Reads a `Value`-shaped JS object into `value`; `text` backs a text payload
// and must outlive the C call.
bool ReadValueSpec(CheckedSpecReader& reader, const Napi::Object& spec, std::string& text, fm_value_t& value) {
  (void)reader.String(spec, "text", &text);
  value.kind = static_cast<fm_value_kind_t>(reader.I32(spec, "kind", FM_VAL_BLANK));
  switch (value.kind) {
    case FM_VAL_NUMBER:
      value.u.number = reader.Double(spec, "number", 0.0);
      break;
    case FM_VAL_BOOL:
      value.u.boolean = reader.I32(spec, "boolean", 0) != 0 ? 1 : 0;
      break;
    case FM_VAL_TEXT:
      value.u.text = text.c_str();
      break;
    case FM_VAL_ERROR:
      value.u.error_code = reader.I32(spec, "errorCode", 0);
      break;
    default:
      break;
  }
  return reader.ok();
}

namespace {

// Fills `out` ({status, cells, nextCursor}) from a cell-range page, which it
// consumes. `rc` is the status of the call that produced `page`.
void FillCellsPage(Napi::Env env, Napi::Object out, fm_cell_range_t* page, fm_status_t rc) {
  Napi::Array cells = out.Get("cells").As<Napi::Array>();
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return;
  }
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
    cell.Set("value", TranslateValue(env, value));
    cells.Set(static_cast<uint32_t>(i), cell);
  }
  uint64_t next = UINT64_MAX;
  if (rc == 0) {
    rc = fm_cell_range_next_cursor(page, &next);
  }
  fm_cell_range_destroy(page);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return;
  }
  out.Set("status", MakeOkStatus(env));
  if (next != UINT64_MAX) {
    out.Set("nextCursor", Napi::Number::New(env, static_cast<double>(next)));
  }
}

}  // namespace

Napi::Value Workbook::GetCellsInRange(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  Napi::Array cells = Napi::Array::New(env);
  out.Set("cells", cells);
  out.Set("nextCursor", env.Null());
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
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
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  FillCellsPage(env, out, page, rc);
  return out;
}

Napi::Value Workbook::ListInvalidCells(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("cells", Napi::Array::New(env));
  out.Set("nextCursor", env.Null());
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  const uint64_t cursor = static_cast<uint64_t>(NumberArg(info, 1, 0.0));
  const uint32_t limit = ArgU32(info, 2);
  fm_cell_range_t* page = nullptr;
  const fm_status_t rc = fm_sheet_get_invalid_cells(handle_, ArgU32(info, 0), cursor, limit, &page);
  FillCellsPage(env, out, page, rc);
  return out;
}

Napi::Value Workbook::ValidateValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  Napi::Object out = Napi::Object::New(env);
  out.Set("hasRule", Napi::Boolean::New(env, false));
  out.Set("valid", Napi::Boolean::New(env, false));
  out.Set("ruleIndex", Napi::Number::New(env, 0));
  out.Set("errorStyle", Napi::Number::New(env, 0));
  if (handle_ == nullptr) {
    out.Set("status", NullHandleError(env));
    return out;
  }
  std::string text;
  fm_value_t value{};
  const Napi::Object spec = ArgObject(info, 3);
  if (!spec.Has("kind")) {
    out.Set("status", MakeBindingError(env, kInvalidArgument, "validateValue: `value` requires a `kind`"));
    return out;
  }
  CheckedSpecReader reader(env);
  if (!ReadValueSpec(reader, spec, text, value)) {
    return env.Undefined();
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_validation_outcome_t outcome{};
  const fm_status_t rc = fm_sheet_validate_value(handle_, address.sheet, address.row, address.col, &value, &outcome);
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return out;
  }
  out.Set("status", MakeOkStatus(env));
  out.Set("hasRule", Napi::Boolean::New(env, outcome.has_rule != 0));
  out.Set("valid", Napi::Boolean::New(env, outcome.valid != 0));
  out.Set("ruleIndex", Napi::Number::New(env, outcome.rule_index));
  out.Set("errorStyle", Napi::Number::New(env, outcome.error_style));
  return out;
}

Napi::Value Workbook::GetDisplayText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeDisplayResult(env, kBindingInvalidHandle, nullptr, 0);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* text = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc =
      fm_sheet_get_display_text(handle_, address.sheet, address.row, address.col, &text, &display_status);
  return MakeDisplayResult(env, rc, text, display_status);
}

Napi::Value Workbook::FormatValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeDisplayResult(env, kBindingInvalidHandle, nullptr, 0);
  }
  const Napi::Object spec = ArgObject(info, 0);
  if (!spec.Has("kind")) {
    Napi::Object out = MakeDisplayResult(env, kInvalidArgument, nullptr, 0);
    out.Set("status", MakeBindingError(env, kInvalidArgument, "formatValue: `value` requires a `kind`"));
    return out;
  }
  const std::string format_code = ArgString(info, 1);
  std::string text;
  fm_value_t value{};
  CheckedSpecReader reader(env);
  if (!ReadValueSpec(reader, spec, text, value)) {
    return env.Undefined();
  }
  const char* out_text = nullptr;
  int32_t display_status = 0;
  const fm_status_t rc = fm_workbook_format_value(handle_, &value, format_code.c_str(), &out_text, &display_status);
  return MakeDisplayResult(env, rc, out_text, display_status);
}

}  // namespace formulon_node
