// Guards the WASM shadow-stack budget against the recursive formula walkers.
//
// `parser::kMaxFormulaAstDepth` is the engine's single nesting bound: the
// Pratt parser counts its own recursion against it, and `ast_shift`,
// `ast_format` and the XLSB `ptg_reader` re-check the finished tree against
// the same number before walking it. That bound is only a stack-safety
// guarantee if the stack it is sized against is known, and under Emscripten
// the stack is a link-time constant. This program measures what the bound
// actually costs and fails if the two have drifted apart.
//
// It exists as a separate executable rather than a case inside the Node
// smoke suite because peak stack usage is not observable through the binding
// surface. The measurement paints the unused shadow stack with a known byte,
// runs one walker, then scans upward from the stack limit for the first byte
// the walker overwrote; usage is `base - lowest_touched`.
//
// Three ways this can fail, in the order they are worth catching:
//
//   * the link stopped applying `-sSTACK_SIZE`, or a toolchain renamed it,
//     so the binary silently fell back to Emscripten's 64 KiB default;
//   * a walker's per-frame cost grew, eating the margin between the measured
//     worst case and the linked size;
//   * `kMaxFormulaAstDepth` was raised without re-pinning the link flag.
//
// Nothing here hard-codes a byte count. The depth comes from the engine's own
// constant and the budget from the linked stack, so both sides of the
// comparison move with the code they describe.
//
// Built and run by `make wasm` / `make test-wasm`; see
// cmake/FormulonWasm.cmake for the link, which mirrors the shipped one so the
// numbers reflect shipped codegen.

#include <emscripten/stack.h>

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>

#include "c_api/formulon_c.h"
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
using formulon::parser::Parser;

// Deepest nesting the parser accepts. One `SUM(` per level plus the `A1`
// leaf puts the tree at exactly `kMaxFormulaAstDepth` nodes deep.
constexpr int kDeepestAccepted = static_cast<int>(kMaxFormulaAstDepth) - 1;

// One level past it. The parser recurses to the cap before emitting its
// diagnostic and unwinding, so this -- not `kDeepestAccepted` -- is the
// parser's true worst case.
constexpr int kRejected = static_cast<int>(kMaxFormulaAstDepth);

// A shallow depth measured alongside the deep one, so a failure can report
// bytes per nesting level instead of a bare total.
constexpr int kShallow = kDeepestAccepted / 4;

// The measured worst case may claim at most this fraction of the linked
// stack. The link is sized at a little over 3x the worst case; requiring 2x
// leaves the guard failing well before a real overflow, while still tolerating
// the frame-layout noise a compiler bump produces.
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

std::string nested_calls(int n) {
  std::string s;
  for (int i = 0; i < n; ++i) {
    s += "SUM(";
  }
  s += "A1";
  s.append(static_cast<std::size_t>(n), ')');
  return s;
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

__attribute__((noinline)) std::size_t measure_parse(int depth, bool expect_clean) {
  const std::string src = nested_calls(depth);
  Arena arena;
  paint_free_stack();
  Parser p(src, arena);
  AstNode* root = p.parse();
  const std::size_t peak = peak_since_paint();
  if (expect_clean && (root == nullptr || !p.errors().empty())) {
    // The downstream walkers are measured over the tree this parse builds. A
    // rejected parse yields a stub, which would under-report them silently.
    fail(
        "the deepest accepted nesting no longer parses cleanly; the walker "
        "measurements below would be taken over a truncated tree");
  }
  return peak;
}

__attribute__((noinline)) std::size_t measure_shift(int depth) {
  const std::string src = nested_calls(depth);
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

__attribute__((noinline)) std::size_t measure_format(int depth) {
  const std::string src = nested_calls(depth);
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
__attribute__((noinline)) std::size_t measure_evaluate(int depth, fm_workbook_t* wb, bool check_value) {
  const std::string src = "=" + nested_calls(depth);
  fm_value_t v{};
  paint_free_stack();
  const fm_status_t rc = fm_workbook_evaluate_formula(wb, 0, 1, 0, src.c_str(), &v);
  const std::size_t peak = peak_since_paint();
  if (check_value) {
    // Nested SUM over a single cell is the identity. A wrong answer here is
    // how a stack that ran into the data segment below it shows up.
    if (rc != 0 || v.kind != FM_VAL_NUMBER || v.u.number != 7.0) {
      fail("the deepest accepted nesting did not evaluate to its input value");
    }
  }
  return peak;
}

void print_walker(const Walker& w) {
  const int span = w.deep_depth - w.shallow_depth;
  const std::size_t per_level =
      (span > 0 && w.deep_peak > w.shallow_peak) ? (w.deep_peak - w.shallow_peak) / static_cast<std::size_t>(span) : 0;
  std::printf("  %-22s depth %3d: %7zu B   depth %3d: %7zu B   %5zu B/level\n", w.label, w.shallow_depth,
              w.shallow_peak, w.deep_depth, w.deep_peak, per_level);
}

}  // namespace

int main() {
  const std::size_t stack_size = static_cast<std::size_t>(emscripten_stack_get_base() - emscripten_stack_get_end());
  const std::size_t expected_stack_size = static_cast<std::size_t>(FORMULON_WASM_EXPECTED_STACK_SIZE);

  std::printf("shadow stack: %zu bytes linked, %zu requested\n", stack_size, expected_stack_size);
  std::printf("nesting cap:  %u (kMaxFormulaAstDepth)\n", static_cast<unsigned>(kMaxFormulaAstDepth));

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

  Walker walkers[] = {
      {"parse", kShallow, kRejected, 0, 0},
      {"parse + evaluate", kShallow, kDeepestAccepted, 0, 0},
      {"shift_relative_refs", kShallow, kDeepestAccepted, 0, 0},
      {"format_formula", kShallow, kDeepestAccepted, 0, 0},
  };

  walkers[0].shallow_peak = measure_parse(kShallow, true);
  walkers[0].deep_peak = measure_parse(kRejected, false);
  measure_parse(kDeepestAccepted, true);
  walkers[1].shallow_peak = measure_evaluate(kShallow, wb, false);
  walkers[1].deep_peak = measure_evaluate(kDeepestAccepted, wb, true);
  walkers[2].shallow_peak = measure_shift(kShallow);
  walkers[2].deep_peak = measure_shift(kDeepestAccepted);
  walkers[3].shallow_peak = measure_format(kShallow);
  walkers[3].deep_peak = measure_format(kDeepestAccepted);

  fm_workbook_destroy(wb);

  std::printf("peak shadow stack by walker:\n");
  std::size_t worst = 0;
  const char* worst_label = "";
  for (const Walker& w : walkers) {
    print_walker(w);
    if (w.deep_peak > worst) {
      worst = w.deep_peak;
      worst_label = w.label;
    }
  }

  const std::size_t budget = stack_size / kStackBudgetDivisor;
  std::printf("worst case:   %zu B (%s), budget %zu B (1/%zu of the linked stack)\n", worst, worst_label, budget,
              kStackBudgetDivisor);

  if (worst > budget) {
    fail(
        "the nesting cap now costs more stack than the link reserves for it; "
        "re-measure and raise _FM_WASM_STACK_SIZE in cmake/FormulonWasm.cmake, "
        "or lower parser::kMaxFormulaAstDepth");
  }

  if (g_failures != 0) {
    std::printf("stack probe: %d failure(s)\n", g_failures);
    return 1;
  }
  std::printf("stack probe: ok\n");
  return 0;
}
