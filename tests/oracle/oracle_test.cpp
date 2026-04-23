// Copyright 2026 libraz. Licensed under the MIT License.
//
// Parameterized gtest that verifies Formulon's tree-walk evaluator against
// golden JSON files produced by the xlwings-driven oracle-gen pipeline.
//
// Each parameter is an (suite, case_id) pair loaded via
// `oracle_runner::load_oracle_cases` from `tests/oracle/golden/*.golden.json`.
// The build wires the directory through the compile-time define
// `FORMULON_ORACLE_GOLDEN_DIR`. When the directory is empty (the expected
// state before oracle-gen has been run on a Mac) the parameter vector is
// empty and gtest registers the suite with zero instantiations, so the
// build stays green.
//
// Divergences are intentionally reported as test failures here: the
// accepted-divergence policy lives in `tests/divergence.yaml` and is
// consulted by the Python generator (it can skip or widen tolerance) — the
// C++ side verifies what the goldens actually commit to.

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "gtest/gtest.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "tests/oracle/json_reader.h"
#include "tests/oracle/oracle_runner.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace tests {
namespace oracle {
namespace {

using formulon::parser::AstNode;
using formulon::parser::Parser;

// Maps the golden JSON's `"#DIV/0!"` style display name to the matching
// ErrorCode enum. Returns `false` on an unknown code.
bool display_name_to_code(std::string_view name, ErrorCode* out) {
  struct Entry {
    std::string_view display;
    ErrorCode code;
  };
  static constexpr Entry kTable[] = {
      {"#NULL!", ErrorCode::Null},    {"#DIV/0!", ErrorCode::Div0},
      {"#VALUE!", ErrorCode::Value},  {"#REF!", ErrorCode::Ref},
      {"#NAME?", ErrorCode::Name},    {"#NUM!", ErrorCode::Num},
      {"#N/A", ErrorCode::NA},        {"#SPILL!", ErrorCode::Spill},
      {"#CALC!", ErrorCode::Calc},    {"#FIELD!", ErrorCode::Field},
      {"#BLOCKED!", ErrorCode::Blocked},
      {"#CONNECT!", ErrorCode::Connect},
      {"#EXTERNAL!", ErrorCode::External},
      {"#BUSY!", ErrorCode::Busy},   {"#PYTHON!", ErrorCode::Python},
      {"#UNKNOWN!", ErrorCode::Unknown},
  };
  for (const auto& e : kTable) {
    if (e.display == name) {
      *out = e.code;
      return true;
    }
  }
  return false;
}

// Applies a single {kind, value} JSON record to a cell. Returns nullptr on
// success or a short error string on failure (unknown kind, missing value,
// etc.) that the caller folds into the test failure message.
const char* apply_cell_value(const JsonValue& spec, Sheet& sheet,
                             std::uint32_t row, std::uint32_t col,
                             Arena& text_arena) {
  if (!spec.is_object()) return "setup cell is not an object";
  const JsonValue* kind_v = spec.find("kind");
  if (kind_v == nullptr || !kind_v->is_string()) return "missing 'kind'";
  const std::string& kind = kind_v->as_string();

  if (kind == "formula") {
    const JsonValue* f = spec.find("formula");
    if (f == nullptr || !f->is_string()) return "formula cell missing 'formula'";
    sheet.set_cell_formula(row, col, f->as_string());
    return nullptr;
  }

  const JsonValue* val_v = spec.find("value");
  if (kind == "blank") {
    // Nothing to do: absence from storage already means blank.
    return nullptr;
  }
  if (kind == "number") {
    if (val_v == nullptr || !val_v->is_number()) return "number missing 'value'";
    sheet.set_cell_value(row, col, Value::number(val_v->as_number()));
    return nullptr;
  }
  if (kind == "bool") {
    if (val_v == nullptr || !val_v->is_bool()) return "bool missing 'value'";
    sheet.set_cell_value(row, col, Value::boolean(val_v->as_bool()));
    return nullptr;
  }
  if (kind == "text") {
    if (val_v == nullptr || !val_v->is_string()) return "text missing 'value'";
    // Intern the string into the arena so the Value's string_view remains
    // valid for the lifetime of the test.
    const std::string& payload = val_v->as_string();
    std::string_view view = text_arena.intern(payload);
    sheet.set_cell_value(row, col, Value::text(view));
    return nullptr;
  }
  if (kind == "error") {
    const JsonValue* code_v = spec.find("code");
    if (code_v == nullptr || !code_v->is_string()) return "error missing 'code'";
    ErrorCode code;
    if (!display_name_to_code(code_v->as_string(), &code)) {
      return "error has unknown 'code'";
    }
    sheet.set_cell_value(row, col, Value::error(code));
    return nullptr;
  }
  return "setup cell has unknown 'kind'";
}

// Renders a Formulon Value as a short display string for assertion failure
// messages. Mirrors the debug_to_string helper but keeps error / text
// formatting consistent with the JSON schema.
std::string format_value(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return "blank";
    case ValueKind::Number:
      return "number(" + std::to_string(v.as_number()) + ")";
    case ValueKind::Bool:
      return v.as_boolean() ? "bool(TRUE)" : "bool(FALSE)";
    case ValueKind::Text:
      return std::string("text(\"") + std::string(v.as_text()) + "\")";
    case ValueKind::Error:
      return std::string("error(") + display_name(v.as_error()) + ")";
    default:
      return "<unsupported>";
  }
}

// Compares `actual` to the golden `expect` JSON record under the given
// tolerance. Returns an empty string on match; otherwise a human-readable
// diff message.
std::string compare_value(const JsonValue& expect, const Value& actual,
                          double tol_abs, double tol_rel) {
  if (!expect.is_object()) return "golden 'expect' is not an object";
  const JsonValue* kind_v = expect.find("kind");
  if (kind_v == nullptr || !kind_v->is_string()) {
    return "golden 'expect' missing 'kind'";
  }
  const std::string& kind = kind_v->as_string();

  if (kind == "blank") {
    if (actual.is_blank()) return {};
    return "expected blank, got " + format_value(actual);
  }
  if (kind == "number") {
    if (!actual.is_number()) return "expected number, got " + format_value(actual);
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_number()) return "golden missing 'value'";
    double want = val_v->as_number();
    double got = actual.as_number();
    // Exact equality is the strict path. When tolerances are non-zero, we
    // accept a match if either absolute or relative diff fits. NaN matches
    // NaN (Excel treats NaN as #NUM!, so this usually doesn't arise).
    if (std::isnan(want) && std::isnan(got)) return {};
    double diff = std::abs(want - got);
    double scale = std::max(std::abs(want), std::abs(got));
    if (diff == 0.0) return {};
    if (tol_abs > 0.0 && diff <= tol_abs) return {};
    if (tol_rel > 0.0 && scale > 0.0 && diff / scale <= tol_rel) return {};
    return "number mismatch: expected " + std::to_string(want) + ", got " +
           std::to_string(got);
  }
  if (kind == "bool") {
    if (!actual.is_boolean()) return "expected bool, got " + format_value(actual);
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_bool()) return "golden missing 'value'";
    if (actual.as_boolean() == val_v->as_bool()) return {};
    return std::string("bool mismatch: expected ") +
           (val_v->as_bool() ? "TRUE" : "FALSE") + ", got " +
           (actual.as_boolean() ? "TRUE" : "FALSE");
  }
  if (kind == "text") {
    if (!actual.is_text()) return "expected text, got " + format_value(actual);
    const JsonValue* val_v = expect.find("value");
    if (val_v == nullptr || !val_v->is_string()) return "golden missing 'value'";
    if (std::string_view(actual.as_text()) ==
        std::string_view(val_v->as_string())) {
      return {};
    }
    return "text mismatch: expected \"" + val_v->as_string() + "\", got \"" +
           std::string(actual.as_text()) + "\"";
  }
  if (kind == "error") {
    if (!actual.is_error()) return "expected error, got " + format_value(actual);
    const JsonValue* code_v = expect.find("code");
    if (code_v == nullptr || !code_v->is_string()) return "golden missing 'code'";
    ErrorCode want_code;
    if (!display_name_to_code(code_v->as_string(), &want_code)) {
      return "golden has unknown error 'code': " + code_v->as_string();
    }
    if (actual.as_error() == want_code) return {};
    return std::string("error mismatch: expected ") + code_v->as_string() +
           ", got " + display_name(actual.as_error());
  }
  return "unknown expect kind: " + kind;
}

// ---------------------------------------------------------------------------
// Parameter provider
// ---------------------------------------------------------------------------

const std::vector<OracleCase>& oracle_cases() {
  // Loaded once at first call; `load_oracle_cases` is safe to call multiple
  // times but we cache to keep test discovery deterministic even if the
  // directory is mutated mid-run (it isn't, but it's cheap insurance).
  static const std::vector<OracleCase> cached =
      load_oracle_cases(configured_golden_dir());
  return cached;
}

// ---------------------------------------------------------------------------
// Fixture
// ---------------------------------------------------------------------------

class OracleTest : public ::testing::TestWithParam<OracleCase> {};

TEST_P(OracleTest, Matches) {
  const OracleCase& param = GetParam();
  if (param.case_id == "<load-error>") {
    const JsonValue* detail = param.raw_case.find("error");
    FAIL() << "failed to load " << param.source_file << ": "
           << (detail && detail->is_string() ? detail->as_string()
                                              : std::string("unknown"));
    return;
  }

  // Cases marked skip-oracle in tests/divergence.yaml flow through oracle-gen
  // with a bare `"skipped": "<reason>"` field (no `expect`). The verifier
  // surfaces them as gtest-skipped so the pass-rate math still reflects them
  // as "known non-verified" rather than hiding the gap.
  if (const JsonValue* skipped = param.raw_case.find("skipped");
      skipped && skipped->is_string()) {
    GTEST_SKIP() << "divergence.yaml skip-oracle: " << skipped->as_string();
  }

  const JsonValue* formula_v = param.raw_case.find("formula");
  const JsonValue* expect_v = param.raw_case.find("expect");
  if (formula_v == nullptr || !formula_v->is_string() ||
      expect_v == nullptr || !expect_v->is_object()) {
    FAIL() << "case " << param.suite << "." << param.case_id
           << " is missing 'formula' or 'expect'";
    return;
  }
  const std::string& formula_src = formula_v->as_string();

  // Build an in-memory workbook seeded with the case's setup cells.
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  Arena text_arena;

  if (const JsonValue* setup = param.raw_case.find("setup");
      setup && setup->is_object()) {
    for (const auto& entry : setup->as_object()) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      if (!a1_to_row_col(entry.first, &row, &col)) {
        FAIL() << param.suite << "." << param.case_id << ": invalid A1 address '"
               << entry.first << "'";
        return;
      }
      const char* err_msg = apply_cell_value(entry.second, sheet, row, col,
                                              text_arena);
      if (err_msg != nullptr) {
        FAIL() << param.suite << "." << param.case_id << ": setup[" << entry.first
               << "]: " << err_msg;
        return;
      }
    }
  }

  // Parse and evaluate the formula through the default registry. Leading '='
  // is stripped to match Formulon's parser expectation (the tokenizer treats
  // the formula body as starting at the first token after '=').
  std::string_view body = formula_src;
  if (!body.empty() && body.front() == '=') body.remove_prefix(1);

  Arena parse_arena;
  Arena eval_arena;
  Parser p(body, parse_arena);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr) << param.suite << "." << param.case_id
                            << ": parse failed for '" << formula_src << "'";

  // Use the full evaluator entry point so recursive cell refs, cycle
  // detection, and the default function registry all kick in — matching
  // what a real Formulon calc would do.
  const eval::FunctionRegistry& registry = eval::default_registry();
  eval::EvalState state;
  eval::EvalContext ctx(wb, sheet, state);
  Value actual = eval::evaluate(*root, eval_arena, registry, ctx);

  std::string diff =
      compare_value(*expect_v, actual, param.tolerance_abs, param.tolerance_rel);
  if (!diff.empty()) {
    FAIL() << param.suite << "." << param.case_id << ": " << diff << "\n"
           << "  formula: " << formula_src << "\n"
           << "  golden file: " << param.source_file;
  }
}

// Human-readable gtest parameter names so failures show up as
// `OracleTest.Matches/<suite>_<case_id>` instead of a numeric index.
std::string PrintParamName(
    const ::testing::TestParamInfo<OracleCase>& info) {
  std::string name = info.param.suite + "_" + info.param.case_id;
  // gtest requires [A-Za-z0-9_]; fold everything else to '_'.
  for (char& c : name) {
    if (!((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') ||
          (c >= '0' && c <= '9') || c == '_')) {
      c = '_';
    }
  }
  return name;
}

INSTANTIATE_TEST_SUITE_P(Oracle, OracleTest,
                         ::testing::ValuesIn(oracle_cases()), PrintParamName);

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
