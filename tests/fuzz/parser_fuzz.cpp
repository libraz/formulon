//
// libFuzzer harness for the Formulon Pratt parser.
//
// Fuzz goal: feed arbitrary bytes as Excel formula text into the parser
// and assert no crash, ASan violation, UBSan violation, or infinite loop.
// The parser must reject malformed input via the diagnostics list rather
// than aborting.
//
// On top of that, every input that parses cleanly is checked against the
// formatter fixpoint property: formatting an AST that came out of a
// diagnostic-free parse, re-parsing that text and formatting again must
// reproduce the same text. A divergence means the tokenizer, the parser and
// the formatter disagree about some surface form -- exactly the class of
// defect that only shows up on a save/reload or on an AST decoded from
// storage rather than from source.
//
// The property is on the emitted text, not on AST identity, and a re-parse
// that itself reports diagnostics is skipped rather than reported. Both
// limits have the same cause: the parser accepts range endpoints that only
// the evaluator can reject (a value-returning call, for one), so some
// reachable ASTs have no surface form that reads back unambiguously. Pinning
// AST identity or re-parse cleanliness here would report that grammar
// looseness as a formatter crash on every run.
//
// The same clean parse then drives the reference-rewriting walker, which is
// the layer row/column insert-delete, sheet removal and defined-name rename
// all route through. Its input is an AST built from file-supplied formula
// text and its parameters are host-supplied indices, so the two arrive from
// different trust domains and only their combination is interesting. The
// shift deltas are derived from a hash of the whole input rather than from a
// prefix, which keeps an input's meaning as formula text unchanged while
// still making the delta reproducible from the input file alone.
//
// Build: clang only, with `-DFM_BUILD_FUZZ=ON`.
// Run: `./build/bin/parser_fuzz -runs=1000` (smoke), or the same binary with
// `-max_total_time=<seconds>` for a longer session.

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string>
#include <string_view>

#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "utils/arena.h"

namespace {

// The build has no exceptions, so a property violation is reported the only
// way libFuzzer understands: an abnormal termination that leaves the input
// on disk as a reproducer. The diverging pair goes to stderr first --
// libFuzzer records the input bytes but not what the formatter did with
// them, and that is the whole content of the finding.
[[noreturn]] void ReportFixpointViolation(std::string_view once, std::string_view twice) {
  std::fprintf(stderr, "format fixpoint violated\n  first:  %.*s\n  second: %.*s\n", static_cast<int>(once.size()),
               once.data(), static_cast<int>(twice.size()), twice.data());
  std::abort();
}

/// FNV-1a over the whole input. Used only to pick a shift delta, so the
/// requirement is determinism and spread, not cryptographic strength.
std::uint32_t HashInput(const uint8_t* data, size_t size) {
  std::uint32_t hash = 2166136261U;
  for (size_t i = 0; i < size; ++i) {
    hash ^= data[i];
    hash *= 16777619U;
  }
  return hash;
}

// Deltas worth spending executions on. The interesting behaviour of the
// walker is at the edges: no movement at all (the identity walk, which must
// not allocate), one step either way, a full Excel axis, and the values
// where a naive `coord + delta` would overflow rather than collapse to
// `#REF!`.
constexpr std::int32_t kShiftDeltas[] = {
    0, 1, -1, 16384, -16384, 1048576, -1048576, 2147483647, -2147483647 - 1,
};

/// Formats `node`, and reports a violation if the text does not survive a
/// re-parse unchanged. A re-parse that itself reports diagnostics is skipped
/// for the reason the file header gives: the grammar admits shapes with no
/// unambiguous surface form, and pinning them here would report grammar
/// looseness as a formatter defect.
void ExpectFormatFixpoint(const formulon::parser::AstNode& node) {
  const std::string once = formulon::parser::format_formula(node);
  formulon::Arena arena;
  formulon::parser::Parser parser(once, arena);
  const formulon::parser::AstNode* reparsed = parser.parse();
  if (reparsed == nullptr || !parser.errors().empty()) {
    return;
  }
  const std::string twice = formulon::parser::format_formula(*reparsed);
  if (twice != once) {
    ReportFixpointViolation(once, twice);
  }
}

}  // namespace

extern "C" int LLVMFuzzerTestOneInput(const uint8_t* data, size_t size) {
  if (size > 65536) {
    return 0;
  }
  formulon::Arena arena;
  std::string_view src(reinterpret_cast<const char*>(data), size);
  formulon::parser::Parser parser(src, arena);
  const formulon::parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    // Malformed input: the diagnostics are the contract, and an AST built
    // out of recovery placeholders carries no round-trip obligation.
    return 0;
  }

  ExpectFormatFixpoint(*root);

  // The reference-rewriting walker, on a tree that is known to parse
  // cleanly. A rewritten tree is still a tree the engine will format and
  // store, so it owes the same surface-form obligation the source AST does:
  // an insert/delete that produces a reference with no readable spelling
  // corrupts the formula on the next save just as surely as a crash would.
  const std::uint32_t hash = HashInput(data, size);
  constexpr std::size_t kDeltaCount = sizeof(kShiftDeltas) / sizeof(kShiftDeltas[0]);
  const std::int32_t row_delta = kShiftDeltas[hash % kDeltaCount];
  const std::int32_t col_delta = kShiftDeltas[(hash / kDeltaCount) % kDeltaCount];

  formulon::Arena shift_arena;
  const formulon::parser::AstNode* shifted =
      formulon::parser::shift_relative_refs(*root, shift_arena, row_delta, col_delta);
  if (shifted == nullptr) {
    // Documented as reachable on arena exhaustion, which is a resource
    // outcome rather than a defect.
    return 0;
  }
  if (shifted == root) {
    // An identity walk returns the input pointer rather than a copy, and
    // most formulas hold no relative reference, so every delta is an
    // identity walk for them. Re-checking the same pointer would repeat the
    // assertion made a few lines above, byte for byte, on the majority of
    // executions; skipping it costs no coverage and returns the budget to
    // the mutator.
    return 0;
  }
  ExpectFormatFixpoint(*shifted);
  return 0;
}
