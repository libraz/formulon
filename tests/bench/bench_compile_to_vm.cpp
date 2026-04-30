// Copyright 2026 libraz. Licensed under the MIT License.
//
// Microbenchmark: bytecode pipeline (compile + optimise + execute) vs the
// tree-walk evaluator on the same set of small expressions.
//
// The VM is currently a feature-parity replacement for the tree-walker;
// it does not yet outperform the walker because none of the planned
// specialised opcodes have landed (range-aware aggregator support, fast
// numeric paths, etc.). The point of this bench is to keep the two
// backends within a constant factor of each other while the optimiser
// is still maturing — the regression gate flags any case where the VM
// drifts beyond 2x the walker's wall clock.
//
// Each scenario reports two timings:
//   * tree-walk (parse + tree_walk_evaluate)
//   * bytecode  (parse + compile + optimize + execute)
//
// Both share the same fixture (a tiny workbook with A1..A10 = 1..10).

#define ANKERL_NANOBENCH_IMPLEMENT
#include <nanobench.h>

#include <cstdint>
#include <cstdio>
#include <cstring>
#include <fstream>
#include <string>
#include <string_view>

#include "eval/bytecode.h"
#include "eval/compiler.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/optimizer.h"
#include "eval/tree_walker.h"
#include "eval/vm.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "value.h"
#include "workbook.h"

namespace {

formulon::Workbook BuildFixture() {
  using formulon::Sheet;
  using formulon::Value;
  using formulon::Workbook;

  Workbook wb = Workbook::create();
  Sheet& sh = wb.sheet(0);
  for (std::uint32_t r = 0; r < 10; ++r) {
    sh.set_cell_value(r, 0U, Value::number(static_cast<double>(r + 1)));
  }
  return wb;
}

void RunTreeWalk(ankerl::nanobench::Bench& bench, const std::string& name, std::string_view src,
                 const formulon::eval::EvalContext& ctx) {
  bench.run(name + " [tree-walk]", [&]() {
    formulon::Arena arena;
    formulon::parser::Parser p(src, arena);
    formulon::parser::AstNode* root = p.parse();
    if (root == nullptr)
      std::abort();
    auto result = formulon::eval::evaluate(*root, arena, formulon::eval::default_registry(), ctx);
    ankerl::nanobench::doNotOptimizeAway(result);
  });
}

void RunVm(ankerl::nanobench::Bench& bench, const std::string& name, std::string_view src,
           const formulon::eval::EvalContext& ctx) {
  bench.run(name + " [vm]", [&]() {
    formulon::Arena arena;
    formulon::parser::Parser p(src, arena);
    formulon::parser::AstNode* root = p.parse();
    if (root == nullptr)
      std::abort();
    auto bc = formulon::eval::compile(*root, arena);
    if (!bc.has_value())
      std::abort();
    formulon::eval::OptimizerStats opt_stats;
    auto opt = formulon::eval::optimize(std::move(bc.value()), arena, &opt_stats);
    if (!opt.has_value())
      std::abort();
    auto result = formulon::eval::execute(opt.value(), arena, formulon::eval::default_registry(), ctx);
    if (!result.has_value())
      std::abort();
    ankerl::nanobench::doNotOptimizeAway(result.value());
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

  // Warm both backends so the first scenario does not absorb one-shot
  // initialisation costs (function-registry construction, JIT-style
  // first-touch caches).
  {
    formulon::Arena arena;
    formulon::parser::Parser p("=A1+1", arena);
    formulon::parser::AstNode* root = p.parse();
    (void)formulon::eval::evaluate(*root, arena, formulon::eval::default_registry(), ctx);
    auto bc = formulon::eval::compile(*root, arena);
    if (bc.has_value()) {
      auto opt = formulon::eval::optimize(std::move(bc.value()), arena, nullptr);
      if (opt.has_value()) {
        (void)formulon::eval::execute(opt.value(), arena, formulon::eval::default_registry(), ctx);
      }
    }
  }

  ankerl::nanobench::Bench bench;
  bench.title("compile_to_vm").unit("eval").warmup(50).relative(true).minEpochIterations(50);

  // Same five scenarios for both backends so the JSON output can be
  // diffed pairwise by `tools/bench/check_regression.py`.
  static const std::pair<std::string, std::string_view> kCases[] = {
      {"const arithmetic 1+2", "=1+2"}, {"ref arithmetic A1+A2", "=A1+A2"},    {"unary +-* MIN/MAX", "=A1+A2*A3-A4"},
      {"comparison chain", "=A1>0"},    {"branching IF", "=IF(A1>0, A2, A3)"},
  };

  for (const auto& [name, src] : kCases) {
    RunTreeWalk(bench, name, src, ctx);
    RunVm(bench, name, src, ctx);
  }

  if (!json_path.empty()) {
    std::ofstream out(json_path);
    bench.render(ankerl::nanobench::templates::json(), out);
  }

  return 0;
}
