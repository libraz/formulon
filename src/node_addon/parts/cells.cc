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
    return MakeErrorStatus(env, kBindingInvalidHandle);
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
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const double value = ArgDouble(info, 3);
  fm_status_t rc = fm_workbook_set_number(handle_, address.sheet, address.row, address.col, value);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetBool(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  bool value = false;
  if (info.Length() > 3) {
    value = info[3].ToBoolean().Value();
  }
  fm_status_t rc = fm_workbook_set_bool(handle_, address.sheet, address.row, address.col, value ? 1 : 0);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetError(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (info.Length() < 4) {
    // Unlike the other cell-mutation setters, a missing 4th argument
    // here has no sane zero-value default: `error_code = 0` silently
    // writes `#NULL!`, masking a caller bug instead of surfacing it.
    // Reject it the same way the WASM (embind arity check) and Python
    // (required positional parameter) bindings already do.
    Napi::TypeError::New(env, "setError requires 4 arguments (sheet, row, col, errorCode)")
        .ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (!info[3].IsNumber()) {
    // `.As<Napi::Number>()` below is an unchecked cast: on a non-number
    // 4th argument it does not itself throw, so without this check the
    // mutation below would run first and any resulting exception would
    // only surface afterwards, on return to JS.
    Napi::TypeError::New(env, "setError: `errorCode` must be a number").ThrowAsJavaScriptException();
    return env.Undefined();
  }
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const int32_t error_code = info[3].As<Napi::Number>().Int32Value();
  fm_status_t rc =
      fm_workbook_set_error(handle_, address.sheet, address.row, address.col, static_cast<fm_error_code_t>(error_code));
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetText(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_workbook_set_text);
}

Napi::Value Workbook::SetCellPhonetic(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_workbook_set_cell_phonetic);
}

Napi::Value Workbook::SetCellPhoneticRuns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  if (info.Length() <= 3 || !info[3].IsArray()) {
    return MakeBindingArgumentError(env, "setCellPhoneticRuns: `runs` must be an array of { sb, eb, text }");
  }
  const Napi::Array runs = info[3].As<Napi::Array>();
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
  fm_status_t rc = fm_workbook_set_cell_phonetic_runs(handle_, address.sheet, address.row, address.col, records.data(),
                                                      records.size());
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  if (info.Length() <= 3 || !info[3].IsObject()) {
    return MakeBindingArgumentError(
        env, "setCellPhoneticProperties: `properties` must be an object { fontId, type, alignment }");
  }
  const Napi::Object props = info[3].As<Napi::Object>();
  CheckedSpecReader reader(env);
  const uint32_t font_id = reader.U32(props, "fontId", 0U);
  const uint32_t type = reader.U32(props, "type", 0U);
  const uint32_t alignment = reader.U32(props, "alignment", 0U);
  if (!reader.ok()) {
    // A malformed field left a pending JS exception: stop before the C ABI
    // call commits a default value for it.
    return env.Undefined();
  }
  fm_status_t rc = fm_workbook_set_cell_phonetic_properties(handle_, address.sheet, address.row, address.col, font_id,
                                                            type, alignment);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetBlank(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return NullHandleError(env);
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_status_t rc = fm_workbook_set_blank(handle_, address.sheet, address.row, address.col);
  return MakeStatus(env, rc);
}

Napi::Value Workbook::SetFormula(const Napi::CallbackInfo& info) {
  return InvokeCellText(info, handle_, &fm_workbook_set_formula);
}

// ---- Cell read ------------------------------------------------------

Napi::Value Workbook::GetValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_value_t v{};
  fm_status_t rc = fm_workbook_get_value(handle_, address.sheet, address.row, address.col, &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetCellPhonetic(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "value", "");
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_get_cell_phonetic(handle_, address.sheet, address.row, address.col, &text);
  return MakeStringResult(env, rc, text);
}

Napi::Value Workbook::GetCellPhoneticRuns(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeFieldResult(env, NullHandleError(env), "runs", Napi::Array::New(env, 0));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  uint32_t count = 0;
  fm_status_t rc = fm_workbook_get_cell_phonetic_run_count(handle_, address.sheet, address.row, address.col, &count);
  Napi::Array out = Napi::Array::New(env, rc == 0 ? count : 0);
  for (uint32_t i = 0; rc == 0 && i < count; ++i) {
    fm_phonetic_run_t run{};
    rc = fm_workbook_get_cell_phonetic_run(handle_, address.sheet, address.row, address.col, i, &run);
    if (rc != 0) {
      break;
    }
    Napi::Object entry = Napi::Object::New(env);
    entry.Set("sb", Napi::Number::New(env, run.sb));
    entry.Set("eb", Napi::Number::New(env, run.eb));
    // Copied immediately: each read refreshes the handle's scratch, so the
    // previous run's pointer is dead by the time the next one lands.
    entry.Set("text", JsString(env, run.text));
    out.Set(i, entry);
  }
  if (rc != 0) {
    return MakeFieldResult(env, MakeErrorStatus(env, rc), "runs", Napi::Array::New(env, 0));
  }
  return MakeFieldResult(env, MakeOkStatus(env), "runs", out);
}

Napi::Value Workbook::GetCellPhoneticProperties(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  const CellAddressArgs address = ReadCellAddress(info);
  uint32_t font_id = 0;
  uint32_t type = 0;
  uint32_t alignment = 0;
  const fm_status_t rc = handle_ != nullptr
                             ? fm_workbook_get_cell_phonetic_properties(handle_, address.sheet, address.row,
                                                                        address.col, &font_id, &type, &alignment)
                             : kBindingInvalidHandle;
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
  out.Set("status", MakeStatus(env, rc));
  return out;
}

Napi::Value Workbook::EvaluateFormulaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeEmptyValueResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string formula = ArgString(info, 3);
  fm_value_t v{};
  fm_status_t rc = fm_workbook_evaluate_formula(handle_, address.sheet, address.row, address.col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::EvaluateFormulaArray(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return EmptyFormulaArrayResult(env, NullHandleError(env));
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const std::string formula = ArgString(info, 3);

  uint32_t rows = 0;
  uint32_t cols = 0;
  fm_status_t rc = fm_workbook_evaluate_formula_array(handle_, address.sheet, address.row, address.col, formula.c_str(),
                                                      &rows, &cols);
  if (rc != 0) {
    return EmptyFormulaArrayResult(env, MakeErrorStatus(env, rc));
  }

  // Build a rows x cols nested array of Value objects, reading each stashed
  // cell by its row-major index (r * cols + c).
  Napi::Array cells = Napi::Array::New(env, rows);
  for (uint32_t r = 0; r < rows; ++r) {
    Napi::Array js_row = Napi::Array::New(env, cols);
    for (uint32_t c = 0; c < cols; ++c) {
      const std::size_t index = static_cast<std::size_t>(r) * cols + c;
      fm_value_t v{};
      fm_status_t cell_rc = fm_workbook_evaluate_formula_array_cell(handle_, index, &v);
      if (cell_rc != 0) {
        return EmptyFormulaArrayResult(env, MakeErrorStatus(env, cell_rc));
      }
      js_row.Set(c, TranslateValue(env, v));
    }
    cells.Set(r, js_row);
  }

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
  fm_status_t rc = fm_workbook_evaluate_cf_formula(handle_, address.sheet, address.row, address.col, anchor_row,
                                                   anchor_col, formula.c_str(), &v);
  return MakeValueResult(env, rc, v);
}

Napi::Value Workbook::GetLambdaText(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeStringFieldResult(env, NullHandleError(env), "text", "");
  }
  const CellAddressArgs address = ReadCellAddress(info);
  const char* text = nullptr;
  fm_status_t rc = fm_workbook_lambda_text_at(handle_, address.sheet, address.row, address.col, &text);
  return MakeStringFieldResult(env, rc, "text", text);
}

// ---- Formula text, range enumeration, display text ------------------

namespace {

// Builds `{ status, <field>: string | null }`; an empty C string is "no formula".
Napi::Object MakeFormulaResult(Napi::Env env, fm_status_t code, const char* formula) {
  const bool present = code == 0 && formula != nullptr && formula[0] != '\0';
  return MakeFieldResult(env, MakeStatus(env, code), "formula",
                         present ? static_cast<Napi::Value>(Napi::String::New(env, formula)) : env.Null());
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
  return FormulaTextResult(info, handle_, &fm_workbook_get_formula);
}

Napi::Value Workbook::GetFormulaR1C1(const Napi::CallbackInfo& info) {
  return FormulaTextResult(info, handle_, &fm_workbook_get_formula_r1c1);
}

namespace {

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

// Fills `out` ({status, cells, nextCursor}) from a cell-range page, which it
// consumes. `rc` is the status of the call that produced `page`.
void FillCellsPage(Napi::Env env, Napi::Object out, fm_cell_range_t* page, fm_status_t rc) {
  Napi::Array cells = out.Get("cells").As<Napi::Array>();
  if (rc != 0) {
    out.Set("status", MakeErrorStatus(env, rc));
    return;
  }
  size_t count = 0;
  rc = fm_cell_range_count(page, &count);
  for (size_t i = 0; rc == 0 && i < count; ++i) {
    uint32_t row = 0;
    uint32_t col = 0;
    const char* formula = nullptr;
    fm_value_t value{};
    rc = fm_cell_range_at(page, i, &row, &col, &formula, &value);
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
  if (info.Length() < 2) {
    out.Set("status",
            MakeBindingArgumentError(env, "getCellsInRange expects (sheet:number, range:object, cursor?, limit?)"));
    return out;
  }
  const bool has_cursor = info.Length() > 2 && info[2].IsNumber();
  const uint64_t cursor = has_cursor ? static_cast<uint64_t>(info[2].As<Napi::Number>().DoubleValue()) : 0U;
  const uint32_t limit = info.Length() > 3 && info[3].IsNumber() ? ArgU32(info, 3) : 0U;
  CheckedSpecReader reader(env);
  fm_merge_range range{};
  if (!MergeRangeArg(reader, info, 1, &range)) {
    return env.Undefined();
  }
  fm_cell_range_t* page = nullptr;
  fm_status_t rc = fm_sheet_cells_in_range(handle_, ArgU32(info, 0), range.first_row, range.first_col, range.last_row,
                                           range.last_col, cursor, limit, &page);
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
  const bool has_cursor = info.Length() > 1 && info[1].IsNumber();
  const uint64_t cursor = has_cursor ? static_cast<uint64_t>(info[1].As<Napi::Number>().DoubleValue()) : 0U;
  const uint32_t limit = info.Length() > 2 && info[2].IsNumber() ? ArgU32(info, 2) : 0U;
  fm_cell_range_t* page = nullptr;
  const fm_status_t rc = fm_sheet_list_invalid_cells(handle_, ArgU32(info, 0), cursor, limit, &page);
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
  if (info.Length() < 4 || !info[3].IsObject()) {
    out.Set("status",
            MakeBindingArgumentError(env, "validateValue expects (sheet:number, row:number, col:number, value:Value)"));
    return out;
  }
  std::string text;
  fm_value_t value{};
  CheckedSpecReader reader(env);
  if (!ReadValueSpec(reader, info[3].As<Napi::Object>(), text, value)) {
    return env.Undefined();
  }
  const CellAddressArgs address = ReadCellAddress(info);
  fm_validation_outcome outcome{};
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
      fm_workbook_get_display_text(handle_, address.sheet, address.row, address.col, &text, &display_status);
  return MakeDisplayResult(env, rc, text, display_status);
}

Napi::Value Workbook::FormatValue(const Napi::CallbackInfo& info) {
  Napi::Env env = info.Env();
  if (handle_ == nullptr) {
    return MakeDisplayResult(env, kBindingInvalidHandle, nullptr, 0);
  }
  if (info.Length() < 1 || !info[0].IsObject()) {
    Napi::Object out = MakeDisplayResult(env, kBindingNullPointer, nullptr, 0);
    out.Set("status", MakeBindingArgumentError(env, "formatValue expects (value:Value, formatCode:string)"));
    return out;
  }
  const Napi::Object spec = info[0].As<Napi::Object>();
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
