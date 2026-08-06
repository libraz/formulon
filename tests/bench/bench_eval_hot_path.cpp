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

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }

  return 0;
}
