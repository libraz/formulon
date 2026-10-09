// Parameterized gtest that verifies Formulon's tree-walk evaluator against
// golden JSON files produced by the xlwings-driven oracle-gen pipeline.
//
// Fixture, parameter-name formatter, instantiation, and the configured
// directory guard stay in this translation unit so every oracle target keeps
// the same registration topology. Comparison/setup/provider helpers live in
// oracle_test_support.h.

#include <filesystem>
#include <fstream>
#include <map>
#include <random>
#include <string>

#include "gtest/gtest.h"
#include "tests/oracle/oracle_test_support.h"

#ifndef FORMULON_ORACLE_PRIMARY_PROFILE_ID
#define FORMULON_ORACLE_PRIMARY_PROFILE_ID "mac-365-ja_JP"
#endif

namespace formulon {
namespace tests {
namespace oracle {
namespace {

using namespace support;
using formulon::parser::AstNode;
using formulon::parser::Parser;

class OracleTest : public ::testing::TestWithParam<OracleCase> {};

TEST_P(OracleTest, Matches) {
  const OracleCase& param = GetParam();
  if (param.case_id == "<load-error>") {
    const JsonValue* detail = param.raw_case.find("error");
    FAIL() << "failed to load " << param.source_file << ": "
           << (detail && detail->is_string() ? detail->as_string() : std::string("unknown"));
    return;
  }

  // Cases marked skip-oracle in tests/divergence.yaml reach the verifier
  // either with a bare `"skipped": "<reason>"` field baked into the golden
  // (no `expect`) or, for a golden captured before the entry was
  // adjudicated, through the divergence registry the loader consults. Both
  // land in `skipped_reason`. Surface them as gtest-skipped so the pass-rate
  // math still reflects them as "known non-verified" rather than hiding the
  // gap.
  if (!param.skipped_reason.empty()) {
    GTEST_SKIP() << "divergence.yaml skip-oracle: " << param.skipped_reason;
  }

  const JsonValue* formula_v = param.raw_case.find("formula");
  const JsonValue* expect_v = param.raw_case.find("expect");
  if (formula_v == nullptr || !formula_v->is_string() || expect_v == nullptr || !expect_v->is_object()) {
    FAIL() << "case " << param.suite << "." << param.case_id << " is missing 'formula' or 'expect'";
    return;
  }
  const std::string& formula_src = formula_v->as_string();

  // Build an in-memory workbook seeded with the case's setup cells.
  Workbook wb = Workbook::create();
  ExcelProfile profile;
  // A golden may name its own evaluation profile (IronCalc corpus: the profile
  // of the host that saved the workbook); otherwise the binary's primary applies.
  const JsonValue* golden_profile_v = param.environment.find("profile_id");
  if (param.variant.empty() && golden_profile_v != nullptr && golden_profile_v->is_string()) {
    ASSERT_TRUE(parse_excel_profile_id(golden_profile_v->as_string(), &profile))
        << "invalid golden profile_id: " << golden_profile_v->as_string();
    wb.set_excel_profile(profile);
  } else if (param.variant.empty()) {
    ASSERT_TRUE(parse_excel_profile_id(FORMULON_ORACLE_PRIMARY_PROFILE_ID, &profile))
        << "invalid primary oracle profile: " << FORMULON_ORACLE_PRIMARY_PROFILE_ID;
    wb.set_excel_profile(profile);
  } else if (parse_excel_profile_id(param.variant, &profile)) {
    wb.set_excel_profile(profile);
  } else if (param.variant.rfind("win-", 0) == 0) {
    wb.set_excel_profile(win_365_ja_jp_profile());
  } else {
    wb.set_excel_profile(mac_365_ja_jp_profile());
  }
  // The oracle generator records the suite-level date system in the
  // environment record. Keep the native workbook and every evaluator context
  // on the same epoch; absent or non-boolean values preserve the 1900 default.
  // Coverage is intentionally through the parameterized oracle path: the H3
  // golden's date1904=true environment and serial anchor exercise parsing,
  // workbook setup, and formula evaluation together.
  if (const JsonValue* date1904_v = param.environment.find("date1904");
      date1904_v != nullptr && date1904_v->is_bool()) {
    wb.set_date1904(date1904_v->as_bool());
  }
  // Honour the suite-level iterative-calc flag. The Python generator stamps
  // `environment.iterative` from the suite's `options.iterative`; when true
  // the workbook resolves circular formulas via fixed-point iteration
  // (Excel's "Enable iterative calculation" option) instead of surfacing a
  // circular-reference error. Defaults match Excel: 100 iterations,
  // max-change 0.001.
  if (const JsonValue* iter_v = param.environment.find("iterative");
      iter_v != nullptr && iter_v->is_bool() && iter_v->as_bool()) {
    IterativeOptions iopts;
    iopts.enabled = true;
    wb.set_iterative_options(iopts);
  }

  // Deliberately not hoisted into a `Sheet&`: a sheet-qualified setup key adds
  // a sheet, and `Workbook` stores sheets by value, so any reference taken
  // before that point dangles afterwards. Re-fetching by index at each use
  // keeps the default sheet valid across the whole case.
  Arena text_arena;
  std::vector<SetupFormulaCell> setup_formulas;

  // Determine the formula-under-test cell up-front so we can pre-register the
  // formula text on the target cell before applying setup. This lets setup
  // formulas that reference the test cell via FORMULATEXT / ISFORMULA observe
  // the test cell as a formula cell. Setup overrides written below win because
  // they execute after this pre-registration.
  // Placement precedence, highest first:
  //   1. the case's explicit `formula_cell` (the drivers write the formula
  //      there, so the native run must evaluate there too),
  //   2. an `id` that is itself an A1 address,
  //   3. Z1 = (row=0, col=25), the default driver placement.
  std::uint32_t case_row = 0;
  std::uint32_t case_col = 25;
  const JsonValue* formula_cell_v = param.raw_case.find("formula_cell");
  if (formula_cell_v != nullptr && formula_cell_v->is_string()) {
    if (!a1_to_row_col(formula_cell_v->as_string(), &case_row, &case_col)) {
      FAIL() << param.suite << "." << param.case_id << ": invalid formula_cell address '" << formula_cell_v->as_string()
             << "'";
      return;
    }
  } else if (!a1_to_row_col(param.case_id, &case_row, &case_col)) {
    case_row = 0U;
    case_col = 25U;
  }
  // Pre-register the formula-under-test on its target cell so cross-references
  // (e.g. setup formulas calling FORMULATEXT(<test-cell>)) see the cell as a
  // formula cell. The actual value comes from `eval::evaluate` below; the
  // formula_text stored here is only consulted by FORMULATEXT / ISFORMULA.
  wb.sheet(0).set_cell_formula(case_row, case_col, formula_src);

  if (const JsonValue* merges = param.raw_case.find("merges"); merges != nullptr) {
    const char* err_msg = apply_merge_ranges(*merges, wb.sheet(0));
    if (err_msg != nullptr) {
      FAIL() << param.suite << "." << param.case_id << ": " << err_msg;
      return;
    }
  }

  if (const JsonValue* setup = param.raw_case.find("setup"); setup && setup->is_object()) {
    for (const auto& entry : setup->as_object()) {
      // Setup keys may be sheet-qualified ("Sheet2!A1", "'My Sheet'!B5")
      // or bare A1; the qualifier targets a secondary sheet that we add
      // to the workbook on first reference. Bare keys still land on the
      // default Sheet1, preserving the historical default behaviour.
      auto [sheet_name, bare_addr] = split_sheet_qualified_addr(entry.first);
      Sheet* target_sheet = &wb.sheet(0);
      std::size_t target_sheet_index = 0;
      if (!sheet_name.empty()) {
        std::size_t idx = wb.sheet_index_by_name(sheet_name);
        if (idx == static_cast<std::size_t>(-1)) {
          target_sheet_index = wb.sheet_count();
          target_sheet = &wb.sheet(wb.add_sheet(sheet_name));
        } else {
          target_sheet_index = idx;
          target_sheet = &wb.sheet(idx);
        }
      }
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      if (!a1_to_row_col(bare_addr, &row, &col)) {
        FAIL() << param.suite << "." << param.case_id << ": invalid A1 address '" << entry.first << "'";
        return;
      }
      const char* err_msg = apply_cell_value(entry.second, *target_sheet, row, col, text_arena);
      if (err_msg != nullptr) {
        FAIL() << param.suite << "." << param.case_id << ": setup[" << entry.first << "]: " << err_msg;
        return;
      }
      const JsonValue* kind_v = entry.second.find("kind");
      const JsonValue* setup_formula_v = entry.second.find("formula");
      if (kind_v != nullptr && kind_v->is_string() && kind_v->as_string() == "formula" && setup_formula_v != nullptr &&
          setup_formula_v->is_string()) {
        setup_formulas.push_back({target_sheet_index, row, col, setup_formula_v->as_string()});
      }
    }
  }

  // Materialise setup formulas before evaluating the formula under test.
  // This matters for dynamic-array setup like `A1=SEQUENCE(3)`: Excel has
  // already committed the A1:A3 spill by the time `=SUM(A1#)` or
  // `=_xlfn.ANCHORARRAY(A1)` runs. The oracle harness previously stored the
  // formula text only, so spill-aware consumers saw no committed region.
  const eval::FunctionRegistry& registry = eval::default_registry();
  eval::EvalState setup_state;
  for (const SetupFormulaCell& setup_formula : setup_formulas) {
    Sheet& setup_sheet = wb.sheet(setup_formula.sheet_index);
    std::string_view setup_body = setup_formula.formula;
    if (!setup_body.empty() && setup_body.front() == '=') {
      setup_body.remove_prefix(1);
    }
    Arena setup_parse_arena;
    Arena setup_eval_arena;
    Parser setup_parser(setup_body, setup_parse_arena);
    AstNode* setup_root = setup_parser.parse();
    ASSERT_NE(setup_root, nullptr) << param.suite << "." << param.case_id << ": setup formula parse failed for '"
                                   << setup_formula.formula << "'";
    eval::EvalContext setup_ctx = eval::EvalContext(wb, setup_sheet, setup_state)
                                      .with_excel_profile(wb.excel_profile())
                                      .with_date1904(wb.date1904())
                                      .with_mutable_sheet(setup_sheet)
                                      .with_formula_cell(setup_formula.row, setup_formula.col);
    Value setup_value = eval::evaluate(*setup_root, setup_eval_arena, registry, setup_ctx);
    setup_value = setup_ctx.dispatch_array_result(setup_value);
    setup_sheet.set_cell_cached_value(setup_formula.row, setup_formula.col, setup_value);
  }

  // Parse and evaluate the formula through the default registry. Leading '='
  // is stripped to match Formulon's parser expectation (the tokenizer treats
  // the formula body as starting at the first token after '=').
  std::string_view body = formula_src;
  if (!body.empty() && body.front() == '=')
    body.remove_prefix(1);

  Arena parse_arena;
  Arena eval_arena;
  Parser p(body, parse_arena);
  AstNode* root = p.parse();
  ASSERT_NE(root, nullptr) << param.suite << "." << param.case_id << ": parse failed for '" << formula_src << "'";

  // Use the full evaluator entry point so recursive cell refs, cycle
  // detection, and the default function registry all kick in — matching
  // what a real Formulon calc would do.
  eval::EvalState state;
  eval::EvalContext ctx =
      eval::EvalContext(wb, wb.sheet(0), state).with_excel_profile(wb.excel_profile()).with_date1904(wb.date1904());
  // Anchor the formula at its own cell (resolved above) so zero-arg ROW() /
  // COLUMN() return the case's row / column. The anchor is what
  // differentiates spill (top-left) from implicit intersection (row/col
  // projection) for the `implicit_intersection` suite and for any other suite
  // that uses `@`-prefixed range arguments, and it decides whether a
  // whole-axis spill footprint still fits inside the grid.
  ctx = ctx.with_formula_cell(case_row, case_col);
  Value actual = eval::evaluate(*root, eval_arena, registry, ctx);

  // A shape-captured golden needs the formula cell to resolve its sample
  // addresses, which the value comparator does not carry.
  const JsonValue* expect_kind = expect_v->find("kind");
  const bool shape_capture =
      expect_kind != nullptr && expect_kind->is_string() && expect_kind->as_string() == "array_shape";
  std::string diff =
      shape_capture ? compare_array_shape(*expect_v, actual, case_row, case_col, param.tolerance_abs,
                                          param.tolerance_rel, param.compare_mode)
                    : compare_value(*expect_v, actual, param.tolerance_abs, param.tolerance_rel, param.compare_mode);
  if (!diff.empty()) {
    FAIL() << param.suite << "." << param.case_id << ": " << diff << "\n"
           << "  formula: " << formula_src << "\n"
           << "  golden file: " << param.source_file;
  }
}

// Human-readable gtest parameter names so failures show up as
// `OracleTest.Matches/<suite>_<case_id>` instead of a numeric index.
std::string PrintParamName(const ::testing::TestParamInfo<OracleCase>& info) {
  std::string name = info.param.suite + "_" + info.param.case_id;
  // Variant suffix is omitted for primary cases so existing test names
  // stay byte-identical to the pre-variant build. Non-empty tags get a
  // `__<tag>` suffix that disambiguates parameter names across binaries
  // and inside a single discovery pass.
  if (!info.param.variant.empty()) {
    name += "__" + info.param.variant;
  }
  return fold_param_name(std::move(name));
}

INSTANTIATE_TEST_SUITE_P(Oracle, OracleTest, ::testing::ValuesIn(oracle_cases()), PrintParamName);

// The variant oracle binary loads only `tests/oracle/targets/<tag>/golden/`
// and is expected to register zero parameters when no variants have been
// scanned in (the default empty-target state). The instantiation above
// then expands to nothing and gtest would otherwise fail the suite with
// `GoogleTestVerification.UninstantiatedParameterizedTestSuite`. Allow that
// state explicitly so an empty variant tree builds and runs cleanly; the
// primary binary still exercises the same instantiation with 2k+ cases, so
// no real coverage is lost by relaxing the check here.
GTEST_ALLOW_UNINSTANTIATED_PARAMETERIZED_TEST(OracleTest);

// The allowance above is suite-wide (it applies to every binary this .cpp is
// compiled into, not just the variant one), so a primary or IronCalc build
// pointed at a golden directory that is misconfigured, moved, or simply
// missing would register zero OracleTest cases and still exit 0 -- the
// "2k+ cases" the comment above promises is nowhere actually asserted. This
// plain, non-parameterized TEST closes that gap without touching the
// variant binary's legitimately-empty state: `configured_golden_dir()` is
// only non-empty for the primary and IronCalc binaries.
TEST(OracleGoldenDirectory, NonEmptyConfiguredDirYieldsCases) {
  const std::string dir = configured_golden_dir();
  if (dir.empty()) {
    GTEST_SKIP() << "no primary golden directory configured for this binary (e.g. the variant oracle build)";
  }
  EXPECT_FALSE(load_oracle_cases(dir, "").empty())
      << "configured golden directory " << dir << " loaded zero oracle cases";
}

}  // namespace
}  // namespace oracle
}  // namespace tests
}  // namespace formulon
