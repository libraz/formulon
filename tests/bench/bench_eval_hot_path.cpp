//
// Microbenchmark: hot-path evaluation throughput for small expressions.
//
// Each scenario parses + tree-walk-evaluates a single representative
// formula in a tight loop. nanobench reports the per-iteration wall-clock
// time and a derived ops-per-second number. There is no hard target — the
// bench tracks regression only — but the regression gate fails the run
// if any scenario exceeds the committed baseline by more than 20%.
//
// Scenarios cover the common shapes seen in real-world workbooks:
//   * scalar arithmetic    `=A1+B1`
//   * range aggregation    `=SUM(A1:A1000)`
//   * conditional dispatch `=IF(A1>0, A1*2, 0)`
//   * lookup               `=VLOOKUP(...)`
//   * conditional aggregate `=SUMIFS(...)`

#define ANKERL_NANOBENCH_IMPLEMENT
#include <nanobench.h>

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <string>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/tree_walker.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace {

// Helper: parse `src` once, evaluate `iters` times against `ctx`. Each
// evaluation uses a fresh per-iteration arena so transient text values
// (concat results, etc.) do not accumulate. Returns the last evaluated
// value to defeat dead-code elimination.
formulon::Value RunScalarLoop(std::string_view src, const formulon::eval::EvalContext& ctx, std::uint64_t iters) {
  formulon::Value last = formulon::Value::blank();
  formulon::Arena parse_arena;
  formulon::parser::Parser p(src, parse_arena);
  formulon::parser::AstNode* root = p.parse();
  if (root == nullptr) {
    std::fprintf(stderr, "bench_eval_hot_path: parse failed for: %.*s\n", static_cast<int>(src.size()), src.data());
    std::abort();
  }
  for (std::uint64_t i = 0; i < iters; ++i) {
    formulon::Arena eval_arena;
    last = formulon::eval::evaluate(*root, eval_arena, formulon::eval::default_registry(), ctx);
    ankerl::nanobench::doNotOptimizeAway(last);
  }
  return last;
}

// Build a workbook that backs every scenario: A1..A1000 = 1..1000 in
// column A, B1..B1000 = 1001..2000 in column B, plus a 5-row lookup
// table at D1:E5 used by the VLOOKUP scenario.
formulon::Workbook BuildFixture() {
  using formulon::Sheet;
  using formulon::Value;
  using formulon::Workbook;

  Workbook wb = Workbook::create();
  Sheet& sh = wb.sheet(0);
  for (std::uint32_t r = 0; r < 1000; ++r) {
    sh.set_cell_value(r, 0U, Value::number(static_cast<double>(r + 1)));
    sh.set_cell_value(r, 1U, Value::number(static_cast<double>(r + 1001)));
  }
  // Text columns for the XLOOKUP scenarios: F is ASCII, G is kana; the last row is the one looked up.
  for (std::uint32_t r = 0; r < 1000; ++r) {
    sh.set_cell_value(r, 5U, Value::text("item" + std::to_string(r)));
    sh.set_cell_value(r, 6U, Value::text("かな" + std::to_string(r)));
  }
  // Lookup table: D1:E5 = (1, "one"), (2, "two"), ... (5, "five"). The
  // VLOOKUP scenario reads D2 -> E2 -> ... so the table must be sorted.
  static const char* kLabels[5] = {"one", "two", "three", "four", "five"};
  for (std::uint32_t r = 0; r < 5; ++r) {
    sh.set_cell_value(r, 3U, Value::number(static_cast<double>(r + 1)));
    sh.set_cell_value(r, 4U, Value::text(kLabels[r]));
  }
  return wb;
}

// Drive nanobench for one scenario. Each call configures a single named
// run on `bench` and writes its iteration timing into the shared output.
void RunScenario(ankerl::nanobench::Bench& bench, const std::string& name, std::string_view src,
                 const formulon::eval::EvalContext& ctx) {
  bench.run(name, [&]() {
    // Inner loop: evaluate the formula once per nanobench iteration. We
    // do NOT batch internally — nanobench runs us many times to settle
    // on a stable per-iteration number, so a single call here gives the
    // most direct M-eval/sec figure.
    formulon::Arena parse_arena;
    formulon::parser::Parser p(src, parse_arena);
    formulon::parser::AstNode* root = p.parse();
    if (root == nullptr) {
      std::fprintf(stderr, "parse failed for: %.*s\n", static_cast<int>(src.size()), src.data());
      std::abort();
    }
    formulon::Arena eval_arena;
    auto result = formulon::eval::evaluate(*root, eval_arena, formulon::eval::default_registry(), ctx);
    ankerl::nanobench::doNotOptimizeAway(result);
  });
}

// Twenty-sheet workbook for cross-sheet scenarios: Sheet12!A1:A1000 are formulas `=Sheet17!B1`, and
// Sheet17!B1:B1000 hold numbers.
formulon::Workbook BuildCrossSheetFixture() {
  using formulon::Value;
  formulon::Workbook wb = formulon::Workbook::create();
  for (int i = 2; i <= 20; ++i) {
    wb.add_sheet("Sheet" + std::to_string(i));
  }
  formulon::Sheet& src = wb.sheet(11);
  formulon::Sheet& dst = wb.sheet(16);
  for (std::uint32_t r = 0; r < 1000; ++r) {
    dst.set_cell_value(r, 1U, Value::number(static_cast<double>(r + 1)));
    src.set_cell_formula(r, 0U, "=Sheet17!B" + std::to_string(r + 1));
  }
  return wb;
}

// Evaluates `src` against a fresh recursion state each iteration so formula cells are not served from the memo.
void RunStatefulScenario(ankerl::nanobench::Bench& bench, const std::string& name, std::string_view src,
                         const formulon::Workbook& wb, const formulon::Sheet& current) {
  bench.run(name, [&]() {
    formulon::Arena parse_arena;
    formulon::parser::Parser p(src, parse_arena);
    formulon::parser::AstNode* root = p.parse();
    if (root == nullptr) {
      std::fprintf(stderr, "parse failed for: %.*s\n", static_cast<int>(src.size()), src.data());
      std::abort();
    }
    formulon::Arena eval_arena;
    formulon::eval::EvalState state;
    const formulon::eval::EvalContext ctx(wb, current, state);
    auto result = formulon::eval::evaluate(*root, eval_arena, formulon::eval::default_registry(), ctx);
    ankerl::nanobench::doNotOptimizeAway(result);
  });
}

}  // namespace

int main(int argc, char* argv[]) {
  std::string json_path;
  for (int i = 1; i < argc; ++i) {
    if (std::strcmp(argv[i], "--json") == 0 && i + 1 < argc) {
      json_path = argv[i + 1];
      ++i;
    }
  }

  formulon::Workbook wb = BuildFixture();
  formulon::eval::EvalContext ctx(wb.sheet(0));

  // Warm the scenarios: the first parse / dispatch in a process pulls
  // function-registry tables and other one-shot caches into memory.
  // Without warming, the first scenario consistently reports inflated
  // numbers and corrupts the baseline.
  (void)RunScalarLoop("=A1+B1", ctx, 1);

  ankerl::nanobench::Bench bench;
  bench.title("eval_hot_path").unit("eval").warmup(50).relative(true).minEpochIterations(50);

  RunScenario(bench, "scalar arithmetic A1+B1", "=A1+B1", ctx);
  RunScenario(bench, "range sum SUM(A1:A1000)", "=SUM(A1:A1000)", ctx);
  RunScenario(bench, "branching IF(A1>0,A1*2,0)", "=IF(A1>0,A1*2,0)", ctx);
  RunScenario(bench, "VLOOKUP D2 against D1:E5", "=VLOOKUP(2,D1:E5,2,FALSE)", ctx);
  RunScenario(bench, "SUMIFS A1:A1000 by B>1500", "=SUMIFS(A1:A1000,B1:B1000,\">1500\")", ctx);

  RunScenario(bench, "TEXT date format over A1:A1000", "=TEXT(A1:A1000,\"yyyy-mm-dd\")", ctx);
  RunScenario(bench, "TEXT number format over A1:A1000", "=TEXT(A1:A1000,\"#,##0.00\")", ctx);

  RunScenario(bench, "XLOOKUP ascii text over F1:F1000", "=XLOOKUP(\"item999\",F1:F1000,A1:A1000)", ctx);
  RunScenario(bench, "XLOOKUP kana text over G1:G1000", "=XLOOKUP(\"かな999\",G1:G1000,A1:A1000)", ctx);

  formulon::Workbook cross_wb = BuildCrossSheetFixture();
  RunStatefulScenario(bench, "cross-sheet SUM(Sheet12!A1:A1000) of =Sheet17!B", "=SUM(Sheet12!A1:A1000)", cross_wb,
                      cross_wb.sheet(0));
  RunStatefulScenario(bench, "cross-sheet SUM(Sheet17!B1:B1000) literal", "=SUM(Sheet17!B1:B1000)", cross_wb,
                      cross_wb.sheet(0));

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }

  return 0;
}
