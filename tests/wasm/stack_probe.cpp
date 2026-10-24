// Guards the WASM shadow-stack budget against the recursive formula walkers.
//
// The parser refuses a formula past Excel's nesting limits
// (`kMaxFunctionNesting`, `kMaxGroupingParenDepth`, `kMaxOperatorNesting`)
// and bounds its own recursion, and every walker's tree, by
// `kMaxFormulaAstDepth`. Those bounds are only a stack-safety guarantee if
// the stack they are sized against is known, and under Emscripten the stack
// is a link-time constant. This program measures the deepest shape of each
// kind the parser accepts, through parse, evaluate, shift and format, plus
// the runaway nesting the parser recurses into before refusing it, and fails
// when the worst exceeds half the linked stack. Each deep shape is built at
// its limit and checked to parse cleanly while one level more is refused, so
// a limit that moves cannot leave the probe measuring a rejected stub.
//
// It is a separate executable because peak stack use is not observable
// through the binding surface: it paints the unused shadow stack, runs one
// walker and scans for the lowest byte overwritten. A flat operator chain is
// one node however long, so it also fails if any walker grows with the
// chain's length. Built and run by `make wasm` / `make test-wasm`; see
// cmake/FormulonWasm.cmake for the link.

#include <emscripten/stack.h>

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>

#include "c_api/formulon_c.h"
#include "eval/tree_walker/depth_guard.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "utils/arena.h"

#ifndef FORMULON_WASM_EXPECTED_STACK_SIZE
#error "FORMULON_WASM_EXPECTED_STACK_SIZE must name the -sSTACK_SIZE the link requested"
#endif

namespace {

using formulon::Arena;
using formulon::parser::AstNode;
using formulon::parser::kMaxFormulaAstDepth;
using formulon::parser::kMaxFunctionNesting;
using formulon::parser::kMaxGroupingParenDepth;
using formulon::parser::kMaxOperatorNesting;
using formulon::parser::Parser;

// The measured worst case may claim at most this fraction of the linked
// stack: the guard then fails well before a real overflow, while still
// tolerating the frame-layout noise a compiler bump produces.
constexpr std::size_t kStackBudgetDivisor = 2;

constexpr unsigned char kPaint = 0xA5;

// Bytes left unpainted immediately below the stack pointer, so the painting
// loop's own frame survives the paint.
constexpr std::uintptr_t kPaintMargin = 512;

std::uintptr_t g_paint_low = 0;
std::uintptr_t g_paint_high = 0;

__attribute__((noinline)) void paint_free_stack() {
  const std::uintptr_t cur = emscripten_stack_get_current();
  g_paint_low = emscripten_stack_get_end();
  g_paint_high = cur - kPaintMargin;
  volatile unsigned char* p = reinterpret_cast<volatile unsigned char*>(g_paint_low);
  for (std::uintptr_t a = g_paint_low; a < g_paint_high; ++a, ++p) {
    *p = kPaint;
  }
}

// Total shadow stack in use at the deepest point reached since the paint.
__attribute__((noinline)) std::size_t peak_since_paint() {
  const volatile unsigned char* p = reinterpret_cast<const volatile unsigned char*>(g_paint_low);
  std::uintptr_t a = g_paint_low;
  for (; a < g_paint_high; ++a, ++p) {
    if (*p != kPaint) {
      break;
    }
  }
  return static_cast<std::size_t>(emscripten_stack_get_base() - a);
}

// `A1+A1+...` with `terms` operands. A flat chain is one node however long it
// is, so every walker must cost the same at 2,700 terms (the most an
// 8,192-character formula holds) as at 150.
std::string flat_chain(int terms) {
  std::string s = "A1";
  for (int i = 1; i < terms; ++i) {
    s += "+A1";
  }
  return s;
}

constexpr int kShortChain = 150;
constexpr int kLongChain = 2700;

// The most a walker's peak may grow between the short and the long chain.
// A walker that recursed per operand would grow by well over 100 KiB.
constexpr std::size_t kChainGrowthAllowance = 4096;

std::string repeated(const char* open, int n, const char* leaf, const char* close) {
  std::string s;
  for (int i = 0; i < n; ++i) {
    s += open;
  }
  s += leaf;
  for (int i = 0; i < n; ++i) {
    s += close;
  }
  return s;
}

// `1+(1+(...A1))`: one grouping pair and one binary level per step.
std::string sum_parens(int n) {
  return repeated("1+(", n, "A1", ")");
}

// `calls` nested IFs around `parens` grouping pairs around `A1`; A1 holds 7,
// so every IF takes its else branch and the value is 7.
std::string ifs_around_parens(int calls, int parens) {
  return repeated("IF(A1=0,0,", calls, repeated("(", parens, "A1", ")").c_str(), ")");
}

// `--...-A1`: a prefix run, one UnaryOp per sign.
std::string prefix_signs(int n) {
  return repeated("-", n, "A1", "");
}

// Peak usage of one walker at a shallow and a deep nesting level.
struct Walker {
  const char* label;
  int shallow_depth;
  int deep_depth;
  std::size_t shallow_peak;
  std::size_t deep_peak;
};

int g_failures = 0;

void fail(const char* what) {
  std::printf("FAIL: %s\n", what);
  ++g_failures;
}

enum class Expect { kClean, kRefused };

__attribute__((noinline)) std::size_t measure_parse(const std::string& src, Expect expect) {
  Arena arena;
  paint_free_stack();
  Parser p(src, arena);
  AstNode* root = p.parse();
  const std::size_t peak = peak_since_paint();
  const bool clean = root != nullptr && p.errors().empty();
  if (expect == Expect::kClean && !clean) {
    // The downstream walkers are measured over the tree this parse builds. A
    // rejected parse yields a stub, which would under-report them silently.
    fail("a shape at its nesting limit no longer parses cleanly; the walkers would measure a stub");
  }
  if (expect == Expect::kRefused && clean) {
    fail("a shape one level past its nesting limit parsed; the probe is not measuring the limit");
  }
  return peak;
}

__attribute__((noinline)) std::size_t measure_shift(const std::string& src) {
  Arena arena;
  Parser p(src, arena);
  const AstNode* root = p.parse();
  if (root == nullptr) {
    return 0;
  }
  paint_free_stack();
  const AstNode* out = formulon::parser::shift_relative_refs(*root, arena, 1, 1);
  const std::size_t peak = peak_since_paint();
  if (out == nullptr) {
    fail("shift_relative_refs exhausted the arena; its measurement is not usable");
  }
  return peak;
}

__attribute__((noinline)) std::size_t measure_format(const std::string& src) {
  Arena arena;
  Parser p(src, arena);
  const AstNode* root = p.parse();
  if (root == nullptr) {
    return 0;
  }
  paint_free_stack();
  const std::string text = formulon::parser::format_formula(*root);
  const std::size_t peak = peak_since_paint();
  if (text.empty()) {
    fail("format_formula produced nothing; its measurement is not usable");
  }
  return peak;
}

// The parse-and-evaluate path a host reaches through the binding surface.
// Evaluation recurses over the tree the parse just built, so this covers the
// consumer that no unit-level walker does.
// `expected` is the formula's value, checked when non-negative: a wrong
// answer is how a stack that ran into the data segment below it shows up.
__attribute__((noinline)) std::size_t measure_evaluate(const std::string& formula, fm_workbook_t* wb, double expected) {
  const std::string src = "=" + formula;
  fm_value_t v{};
  paint_free_stack();
  const fm_status_t rc = fm_workbook_evaluate_formula(wb, 0, 1, 0, src.c_str(), &v);
  const std::size_t peak = peak_since_paint();
  if (expected >= 0.0 && (rc != 0 || v.kind != FM_VAL_NUMBER || v.u.number != expected)) {
    fail("a formula at a nesting limit did not evaluate to its value");
  }
  return peak;
}

void print_walker(const Walker& w) {
  const int span = w.deep_depth - w.shallow_depth;
  const std::size_t per_level =
      (span > 0 && w.deep_peak > w.shallow_peak) ? (w.deep_peak - w.shallow_peak) / static_cast<std::size_t>(span) : 0;
  std::printf("  %-28s depth %4d: %8zu B   depth %4d: %8zu B   %5zu B/level\n", w.label, w.shallow_depth,
              w.shallow_peak, w.deep_depth, w.deep_peak, per_level);
  // A probe that overruns the stack dies mid-run; what was measured so far still reaches the log.
  std::fflush(stdout);
}

}  // namespace

int main() {
  const std::size_t stack_size = static_cast<std::size_t>(emscripten_stack_get_base() - emscripten_stack_get_end());
  const std::size_t expected_stack_size = static_cast<std::size_t>(FORMULON_WASM_EXPECTED_STACK_SIZE);

  std::printf("shadow stack: %zu bytes linked, %zu requested\n", stack_size, expected_stack_size);
  std::printf("nesting caps: %u calls, %u parentheses, %u signs, %u nodes\n",
              static_cast<unsigned>(kMaxFunctionNesting), static_cast<unsigned>(kMaxGroupingParenDepth),
              static_cast<unsigned>(kMaxOperatorNesting), static_cast<unsigned>(kMaxFormulaAstDepth));

  if (stack_size != expected_stack_size) {
    // Either the flag did not reach the link or the toolchain no longer
    // honours it. Every number below would then describe the wrong binary.
    fail("the linked shadow stack is not the size the build asked for");
  }

  fm_workbook_t* wb = nullptr;
  if (fm_workbook_create(&wb) != 0) {
    fail("could not create a workbook to drive the evaluate path");
    return 1;
  }
  fm_workbook_set_number(wb, 0, 0, 0, 7.0);

  const int kOps = static_cast<int>(kMaxOperatorNesting);
  const int kParens = static_cast<int>(kMaxGroupingParenDepth);
  const int kCalls = static_cast<int>(kMaxFunctionNesting);
  const int kRunaway = static_cast<int>(kMaxFormulaAstDepth) + 1;

  // Every deep shape sits exactly at its limit: it parses, one more level does not.
  measure_parse(sum_parens(kParens + 1), Expect::kRefused);
  measure_parse(ifs_around_parens(kCalls + 1, kParens), Expect::kRefused);
  measure_parse(ifs_around_parens(kCalls, kParens + 1), Expect::kRefused);
  measure_parse(prefix_signs(kOps + 1), Expect::kRefused);

  Walker walkers[] = {
      {"parse: runaway (", 64, kRunaway, 0, 0},  {"parse: 1+( run", 64, kParens, 0, 0},
      {"parse: IFs around (", 16, kCalls, 0, 0}, {"parse: sign run", 256, kOps, 0, 0},
      {"evaluate: 1+( run", 64, kParens, 0, 0},  {"evaluate: IFs around (", 16, kCalls, 0, 0},
      {"evaluate: sign run", 256, kOps, 0, 0},   {"shift: 1+( run", 64, kParens, 0, 0},
      {"shift: IFs around (", 16, kCalls, 0, 0}, {"shift: sign run", 256, kOps, 0, 0},
      {"format: 1+( run", 64, kParens, 0, 0},    {"format: IFs around (", 16, kCalls, 0, 0},
      {"format: sign run", 256, kOps, 0, 0},
  };
  // The IF shape's shallow and deep cases scale its calls and parentheses together.
  const std::string sums[] = {sum_parens(64), sum_parens(kParens)};
  const std::string ifs[] = {ifs_around_parens(16, 64), ifs_around_parens(kCalls, kParens)};
  const std::string signs[] = {prefix_signs(256), prefix_signs(kOps)};
  const double sum_values[] = {64.0 + 7.0, kParens + 7.0};
  const std::string* shapes[] = {sums, ifs, signs};

  std::printf("peak shadow stack by walker:\n");
  walkers[0].shallow_peak = measure_parse(repeated("(", 64, "A1", ")"), Expect::kClean);
  walkers[0].deep_peak = measure_parse(repeated("(", kRunaway, "A1", ")"), Expect::kRefused);
  print_walker(walkers[0]);
  for (int k = 0; k < 3; ++k) {
    Walker& w = walkers[1 + k];
    w.shallow_peak = measure_parse(shapes[k][0], Expect::kClean);
    w.deep_peak = measure_parse(shapes[k][1], Expect::kClean);
    print_walker(w);
  }
  for (int k = 0; k < 3; ++k) {
    Walker& w = walkers[4 + k];
    // An even sign run and every IF branch leave A1's 7 unchanged.
    w.shallow_peak = measure_evaluate(shapes[k][0], wb, k == 0 ? sum_values[0] : 7.0);
    w.deep_peak = measure_evaluate(shapes[k][1], wb, k == 0 ? sum_values[1] : 7.0);
    print_walker(w);
  }
  for (int k = 0; k < 3; ++k) {
    Walker& w = walkers[7 + k];
    w.shallow_peak = measure_shift(shapes[k][0]);
    w.deep_peak = measure_shift(shapes[k][1]);
    print_walker(w);
  }
  for (int k = 0; k < 3; ++k) {
    Walker& w = walkers[10 + k];
    w.shallow_peak = measure_format(shapes[k][0]);
    w.deep_peak = measure_format(shapes[k][1]);
    print_walker(w);
  }

  // The same walkers over a flat chain, where depth is the operand count.
  Walker chain_walkers[] = {
      {"chain parse", kShortChain, kLongChain, 0, 0},
      {"chain parse + evaluate", kShortChain, kLongChain, 0, 0},
      {"chain shift", kShortChain, kLongChain, 0, 0},
      {"chain format", kShortChain, kLongChain, 0, 0},
  };
  chain_walkers[0].shallow_peak = measure_parse(flat_chain(kShortChain), Expect::kClean);
  chain_walkers[0].deep_peak = measure_parse(flat_chain(kLongChain), Expect::kClean);
  chain_walkers[1].shallow_peak = measure_evaluate(flat_chain(kShortChain), wb, 7.0 * kShortChain);
  chain_walkers[1].deep_peak = measure_evaluate(flat_chain(kLongChain), wb, 7.0 * kLongChain);
  chain_walkers[2].shallow_peak = measure_shift(flat_chain(kShortChain));
  chain_walkers[2].deep_peak = measure_shift(flat_chain(kLongChain));
  chain_walkers[3].shallow_peak = measure_format(flat_chain(kShortChain));
  chain_walkers[3].deep_peak = measure_format(flat_chain(kLongChain));

  fm_workbook_destroy(wb);

  std::size_t worst = 0;
  const char* worst_label = "";
  for (const Walker& w : walkers) {
    if (w.deep_peak > worst) {
      worst = w.deep_peak;
      worst_label = w.label;
    }
  }

  for (const Walker& w : chain_walkers) {
    print_walker(w);
    if (w.deep_peak > worst) {
      worst = w.deep_peak;
      worst_label = w.label;
    }
    if (w.deep_peak > w.shallow_peak + kChainGrowthAllowance) {
      fail("a walker's stack use grows with the length of a flat operator chain");
    }
  }

  // Evaluation may recurse to `kMaxEvalDepth` levels, through a re-entrant
  // cell chain, past the deepest single formula; extrapolate the sign run's
  // per-level cost to it.
  const Walker& sign_eval = walkers[6];
  const std::size_t eval_per_level =
      sign_eval.deep_peak > sign_eval.shallow_peak
          ? (sign_eval.deep_peak - sign_eval.shallow_peak) / static_cast<std::size_t>(kOps - sign_eval.shallow_depth)
          : 0;
  const std::size_t eval_projected =
      sign_eval.deep_peak + eval_per_level * (formulon::eval::kMaxEvalDepth - static_cast<std::size_t>(kOps));
  std::printf("  %-28s projected to depth %u: %8zu B\n", "evaluate", formulon::eval::kMaxEvalDepth, eval_projected);
  if (eval_projected > worst) {
    worst = eval_projected;
    worst_label = "evaluate projected to kMaxEvalDepth";
  }

  const std::size_t budget = stack_size / kStackBudgetDivisor;
  std::printf("worst case:   %zu B (%s), budget %zu B (1/%zu of the linked stack)\n", worst, worst_label, budget,
              kStackBudgetDivisor);

  if (worst > budget) {
    fail(
        "the nesting limits now cost more stack than the link reserves for them; "
        "re-measure and raise _FM_WASM_STACK_SIZE in cmake/FormulonWasm.cmake");
  }

  if (g_failures != 0) {
    std::printf("stack probe: %d failure(s)\n", g_failures);
    return 1;
  }
  std::printf("stack probe: ok\n");
  return 0;
}
