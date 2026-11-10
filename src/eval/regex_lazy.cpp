//
// REGEXTEST / REGEXEXTRACT / REGEXREPLACE backed by PCRE2 (8-bit, no JIT,
// UTF + UCP). All three functions share the `regex_kernel` helper to
// guarantee a single compile + match pipeline; this is a hard
// architectural requirement (one `pcre2_compile` call site, not three).
//
// Compilation is also hoisted out of the per-subject loop: each impl builds
// one `CompiledPattern` for its (pattern, case-sensitivity) pair and hands
// it to every subject cell, so a range-shaped `text` argument compiles once
// rather than once per cell. `regex_compile_count()` publishes the running
// compile total so that contract stays observable.
//
// Spec / oracle references:
//   * Excel REGEX function family signatures and shape semantics:
//     Mac Excel oracle observations.
//   * Resource limits: match_limit = 1_000_000, depth_limit = 10_000;
//     exhausting the match limit is #VALUE!. PCRE2 applies the match
//     limit per start position, and nothing bounds a call's total work,
//     which grows linearly with subject length -- as Excel's does.
//   * Pattern length cap (32_767 bytes): matches Excel's worksheet
//     formula text limit; longer patterns yield #VALUE! before
//     pcre2_compile is invoked.
//
// Compile flags applied to every pattern:
//   PCRE2_UTF | PCRE2_UCP                       (always)
//   PCRE2_CASELESS                              (when case_sensitivity = 1)
// Notably NOT set:
//   PCRE2_MULTILINE — `^` / `$` match the whole subject only, not each
//                     line. Excel observed behaviour aligns.
//   PCRE2_DOTALL    — `.` does not match `\n`.

#include "eval/regex_lazy.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#define PCRE2_CODE_UNIT_WIDTH 8
#include <pcre2.h>

#include "eval/array_alloc.h"
#include "eval/array_context.h"
#include "eval/coerce.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/regex_program.h"
#include "eval/regex_replacement.h"
#include "eval/scalar_args.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

using detail::CompiledPattern;
using detail::expand_replacement;
using detail::KernelResult;
using detail::kMatchLimitError;
using detail::kResourceError;
using detail::MatchSpan;
using detail::parse_replacement;
using detail::regex_kernel;
using detail::ReplToken;
using detail::scan_matches;
using detail::ScanEnd;

// --- Argument coercion helpers -------------------------------------------
//
// Every REGEX* function shares the same trailing-arg vocabulary (case
// sensitivity, return mode, occurrence). Coercion always: number coerce
// -> truncate -> bound check. A non-coercible value or out-of-domain
// integer surfaces #VALUE! (or the upstream error code).

// Largest magnitude a coerced integer argument may take.
constexpr double kMaxIntArgMagnitude = 2147483647.0;

// Shared helper: evaluates `node`, propagates errors, coerces to a
// truncated integer. On success writes the int into `out` and returns
// true. On failure writes the error sentinel into `out_err` and returns
// false.
bool coerce_int_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    long long& out, Value& out_err) {
  double d = 0.0;
  if (!eval_truncated_number_arg(node, arena, registry, ctx, d, out_err)) {
    return false;
  }
  // Defend against NaN / Inf surviving truncation.
  if (std::isnan(d) || d > kMaxIntArgMagnitude || d < -kMaxIntArgMagnitude) {
    out_err = Value::error(ErrorCode::Value);
    return false;
  }
  out = static_cast<long long>(d);
  return true;
}

// Reads a case_sensitivity argument. Mac Excel 365 convention (matching
// MS docs): 0/FALSE = case-sensitive, 1/TRUE = case-insensitive; any other
// integer -> #VALUE!.
bool read_case_insensitive(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                           const EvalContext& ctx, bool& out, Value& out_err) {
  long long cs = 0;
  if (!coerce_int_arg(node, arena, registry, ctx, cs, out_err)) {
    return false;
  }
  if (cs != 0 && cs != 1) {
    out_err = Value::error(ErrorCode::Value);
    return false;
  }
  out = (cs == 1);
  return true;
}

// Materialises a substring of `subject` into the arena and returns a
// `Value::text` view of the copy. Used when handing match spans back
// to the caller — the subject buffer comes from coerce_to_text's
// std::string and its lifetime ends with the impl's stack frame.
Value text_from_span(std::string_view subject, std::size_t start, std::size_t end, Arena& arena) {
  if (end <= start) {
    return Value::text(arena.intern(""));
  }
  const std::string_view slice = subject.substr(start, end - start);
  return Value::text(arena.intern(slice));
}

// --- Text argument resolution --------------------------------------------
//
// The result shape of a REGEX* call follows the *value* of its `text`
// argument, not the AST kind that produced it: a range, an array literal,
// an arithmetic broadcast and a function call that returns an array all
// broadcast cellwise. This helper materialises the arg as either a scalar
// string or an Array of cells.
//
// On error, sets `out_err` to the error sentinel and returns false.

struct TextArg {
  // Exactly one of `scalar` / `array` is populated. When `is_array` is
  // false the kernel runs once with `scalar` as subject; when true it
  // runs once per cell, broadcasting cellwise and packing results
  // into an output array of the same shape.
  bool is_array = false;
  std::string scalar;
  const ArrayValue* array = nullptr;
};

bool resolve_text_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, TextArg& out, Value& out_err) {
  // `eval_node_as_array` is the shape-preserving seam shared with the other
  // broadcast-aware lazy impls (TEXTSPLIT / FILTER / TRIMRANGE). It handles
  // every AST kind: ranges and array literals keep their (rows, cols), an
  // arithmetic broadcast is evaluated cellwise, a call that returns an array
  // (UNIQUE / SORT / FILTER ...) is forwarded, and anything scalar comes back
  // wrapped as 1x1. Gating on the node kind here would make the result shape
  // depend on how the value was spelled rather than on the value itself.
  const Value v = eval_node_as_array(node, arena, registry, ctx);
  if (v.is_error()) {
    out_err = v;
    return false;
  }
  if (!v.is_array()) {
    // Defensive fallback: eval_node_as_array always wraps to at
    // least 1x1, but if it ever returned a raw scalar we'd treat
    // that as the scalar path.
    auto t = coerce_to_text(v);
    if (!t) {
      out_err = Value::error(t.error());
      return false;
    }
    out.is_array = false;
    out.scalar = std::move(t.value());
    return true;
  }
  out.is_array = true;
  out.array = v.as_array();
  // 1x1 degenerates to a scalar so the result is also scalar (matches
  // Mac Excel: a single-cell range broadcast collapses to a scalar
  // outcome).
  if (out.array->rows == 1U && out.array->cols == 1U) {
    const Value& cell = out.array->cells[0];
    if (cell.is_error()) {
      out_err = cell;
      return false;
    }
    auto t = coerce_to_text(cell);
    if (!t) {
      out_err = Value::error(t.error());
      return false;
    }
    out.is_array = false;
    out.scalar = std::move(t.value());
    out.array = nullptr;
  }
  return true;
}

// Resolves the pattern argument (scalar text, errors propagate). An empty
// pattern is valid and matches at every position.
bool resolve_pattern(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx, std::string& out, Value& out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    out_err = v;
    return false;
  }
  auto t = coerce_to_text(v);
  if (!t) {
    out_err = Value::error(t.error());
    return false;
  }
  out = std::move(t.value());
  return true;
}

// --- REGEXEXTRACT result builders ----------------------------------------

// mode 0: scalar text of the first whole match. `#N/A` on no match.
Value extract_mode0(const KernelResult& kr, std::string_view subject, Arena& arena) {
  if (kr.matches.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const MatchSpan& m = kr.matches.front();
  return text_from_span(subject, m.whole_start, m.whole_end, arena);
}

// mode 1: row array (1 x N) of all whole matches.
Value extract_mode1(const KernelResult& kr, std::string_view subject, Arena& arena) {
  if (kr.matches.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const std::uint32_t n = static_cast<std::uint32_t>(kr.matches.size());
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(1U, n, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    const MatchSpan& m = kr.matches[i];
    cells[i] = text_from_span(subject, m.whole_start, m.whole_end, arena);
  }
  return Value::array(out);
}

// mode 2: row array (1 x G) of capture groups from the FIRST match.
// #N/A on no match; #VALUE! when matches exist but the pattern has no
// capture groups (Excel observed: the user expected at least one
// `(...)` so an empty group list is a definitional mistake).
Value extract_mode2(const KernelResult& kr, std::string_view subject, Arena& arena) {
  if (kr.matches.empty()) {
    return Value::error(ErrorCode::NA);
  }
  if (kr.capture_count == 0U) {
    return Value::error(ErrorCode::Value);
  }
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(1U, kr.capture_count, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const MatchSpan& m = kr.matches.front();
  for (std::uint32_t g = 0; g < kr.capture_count; ++g) {
    if (!m.participated[g]) {
      // Non-participating group: Excel surfaces an empty string for
      // optional groups that didn't capture (e.g. `(a)?b` matching
      // "b").
      cells[g] = Value::text(arena.intern(""));
    } else {
      cells[g] = text_from_span(subject, m.groups[g].first, m.groups[g].second, arena);
    }
  }
  return Value::array(out);
}

Value extract_dispatch(const KernelResult& kr, std::string_view subject, long long mode, Arena& arena) {
  switch (mode) {
    case 0:
      return extract_mode0(kr, subject, arena);
    case 1:
      return extract_mode1(kr, subject, arena);
    case 2:
      return extract_mode2(kr, subject, arena);
    default:
      return Value::error(ErrorCode::Value);
  }
}

// Finishes a replacement result: one past the Excel text cap is `#VALUE!`.
Value finish_replacement(const std::string& out, Arena& arena) {
  if (static_cast<std::uint64_t>(utf16_units_in(out)) > kExcelTextCapUnits) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

// Replaces every match, streaming through the subject.
Value substitute_global(const CompiledPattern& program, std::string_view subject, const std::vector<ReplToken>& tokens,
                        Arena& arena) {
  if (!program.ok()) {
    return Value::error(ErrorCode::Value);
  }
  std::string out;
  std::size_t copied = 0;
  bool expansion_failed = false;
  const ScanEnd end = scan_matches(program, subject, /*find_all=*/true, [&](const PCRE2_SIZE* ovec) {
    const std::size_t start = std::max(static_cast<std::size_t>(ovec[0]), copied);
    const std::size_t stop = std::max(static_cast<std::size_t>(ovec[1]), start);
    out.append(subject.substr(copied, start - copied));
    if (!expand_replacement(tokens, program, subject, ovec, out)) {
      expansion_failed = true;
      return false;
    }
    copied = stop;
    return true;
  });
  if (expansion_failed) {
    return Value::error(ErrorCode::Value);
  }
  if (end == ScanEnd::kLimitHit) {
    return Value::error(kMatchLimitError);
  }
  if (end == ScanEnd::kError) {
    note_unknown_error_origin();
    return Value::error(kResourceError);
  }
  out.append(subject.substr(copied));
  return finish_replacement(out, arena);
}

// Replaces only the N-th match (occurrence == N != 0). A positive N counts
// matches from the start (1 = first); a negative N counts from the end
// (-1 = last), matching Excel's documented REGEXREPLACE occurrence
// semantics. The kernel finds the matches; an occurrence beyond their count
// leaves the text unchanged.
Value substitute_nth(const CompiledPattern& program, std::string_view subject, const std::vector<ReplToken>& tokens,
                     long long occurrence, Arena& arena) {
  KernelResult kr = regex_kernel(program, subject, /*find_all=*/true,
                                 /*on_match_limit_returns_no_match=*/false);
  if (!kr.ok) {
    return kr.err;
  }
  const std::size_t count = kr.matches.size();
  const long long index = occurrence > 0 ? occurrence - 1 : static_cast<long long>(count) + occurrence;
  if (index < 0 || static_cast<std::size_t>(index) >= count) {
    return Value::text(arena.intern(subject));
  }
  const MatchSpan& m = kr.matches[static_cast<std::size_t>(index)];
  std::vector<PCRE2_SIZE> ovec(2U * (static_cast<std::size_t>(kr.capture_count) + 1U), PCRE2_UNSET);
  ovec[0] = m.whole_start;
  ovec[1] = m.whole_end;
  for (std::uint32_t g = 0; g < kr.capture_count; ++g) {
    if (m.participated[g]) {
      ovec[2U * (g + 1U)] = m.groups[g].first;
      ovec[2U * (g + 1U) + 1U] = m.groups[g].second;
    }
  }
  std::string out(subject.substr(0, m.whole_start));
  if (!expand_replacement(tokens, program, subject, ovec.data(), out)) {
    return Value::error(ErrorCode::Value);
  }
  out.append(subject.substr(std::max(m.whole_end, m.whole_start)));
  return finish_replacement(out, arena);
}

// Array broadcast: runs `per_cell` once per text-coerced cell of `in`,
// output shape = input shape. Error cells and failed coercions pass
// through as the cell's error.
template <typename PerCell>
Value broadcast_over_text(const ArrayValue* in, Arena& arena, PerCell&& per_cell) {
  Value* cells = nullptr;
  ArrayValue* out = allocate_derived_array_value(in->rows, in->cols, arena, cells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  // Results are interned into the arena, so a subject only lives for its own cell.
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    if (cell.is_error()) {
      cells[i] = cell;
      continue;
    }
    auto t = coerce_to_text(cell);
    if (!t) {
      cells[i] = Value::error(t.error());
      continue;
    }
    const std::string subject = std::move(t.value());
    cells[i] = per_cell(std::string_view(subject));
  }
  return Value::array(out);
}

// Resolves the leading (text, pattern) pair every REGEX* function takes, in
// signature order so the leftmost error wins.
bool resolve_text_and_pattern(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                              const EvalContext& ctx, TextArg& text_arg, std::string& pattern, Value& err) {
  return resolve_text_arg(call.as_call_arg(0), arena, registry, ctx, text_arg, err) &&
         resolve_pattern(call.as_call_arg(1), arena, registry, ctx, pattern, err);
}

}  // namespace

struct RegexSearch::Impl {
  explicit Impl(std::string_view pattern) : program(pattern, /*case_insensitive=*/false) {}
  CompiledPattern program;
};

RegexSearch::RegexSearch(std::string_view pattern) : impl_(new Impl(pattern)) {}

RegexSearch::~RegexSearch() = default;

bool RegexSearch::ok() const noexcept {
  return impl_->program.ok();
}

Expected<bool, ErrorCode> RegexSearch::matches(std::string_view subject) const {
  const KernelResult result = regex_kernel(impl_->program, subject, /*find_all=*/false,
                                           /*on_match_limit_returns_no_match=*/true);
  if (!result.ok) {
    return result.err.as_error();
  }
  return !result.matches.empty();
}

// ---------------------------------------------------------------------------
// REGEXTEST
// ---------------------------------------------------------------------------
Value eval_regextest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2 || arity > 3) {
    return Value::error(ErrorCode::Value);
  }

  // Arguments are evaluated in signature order so the leftmost error wins,
  // as the rest of the lazy families do. The option arguments feed the
  // compiled program but sit rightmost in the signature, so they are
  // evaluated last and cannot pre-empt an error in text or pattern.
  Value err = Value::error(ErrorCode::Value);

  // Text may be an array (per-cell broadcast); the pattern is scalar text.
  TextArg text_arg;
  std::string pattern;
  if (!resolve_text_and_pattern(call, arena, registry, ctx, text_arg, pattern, err)) {
    return err;
  }

  // Optional case_sensitivity (third arg, default case-sensitive).
  bool case_insensitive = false;
  if (arity == 3 && !read_case_insensitive(call.as_call_arg(2), arena, registry, ctx, case_insensitive, err)) {
    return err;
  }

  // Compile once for the whole call: the pattern and the case flag are
  // loop-invariant, the subjects are not.
  const CompiledPattern program(pattern, case_insensitive);

  auto per_subject = [&](std::string_view subject) -> Value {
    KernelResult kr = regex_kernel(program, subject, /*find_all=*/false,
                                   /*on_match_limit_returns_no_match=*/true);
    if (!kr.ok) {
      return kr.err;
    }
    return Value::boolean(!kr.matches.empty());
  };

  if (!text_arg.is_array) {
    return per_subject(text_arg.scalar);
  }

  return broadcast_over_text(text_arg.array, arena, per_subject);
}

// ---------------------------------------------------------------------------
// REGEXEXTRACT
// ---------------------------------------------------------------------------
Value eval_regexextract_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2 || arity > 4) {
    return Value::error(ErrorCode::Value);
  }

  // Signature order, so the leftmost error wins (see eval_regextest_lazy).
  Value err = Value::error(ErrorCode::Value);

  TextArg text_arg;
  std::string pattern;
  if (!resolve_text_and_pattern(call, arena, registry, ctx, text_arg, pattern, err)) {
    return err;
  }

  // return_mode (third arg, default 0).
  long long mode = 0;
  if (arity >= 3) {
    if (!coerce_int_arg(call.as_call_arg(2), arena, registry, ctx, mode, err)) {
      return err;
    }
    if (mode < 0 || mode > 2) {
      return Value::error(ErrorCode::Value);
    }
  }

  // case_sensitivity (fourth arg, default case-sensitive).
  bool case_insensitive = false;
  if (arity == 4 && !read_case_insensitive(call.as_call_arg(3), arena, registry, ctx, case_insensitive, err)) {
    return err;
  }

  // For mode 1 a single match already yields an array — so
  // running them in array context (per-cell broadcast) produces an
  // array of arrays which Excel/Formulon do not represent.
  // Mac Excel surfaces #CALC! in that situation; we follow suit by
  // collapsing the broadcast to per-cell scalars only when the mode
  // is 0 or 2 (both produce a single scalar / row), and rejecting
  // arrays for mode 1.
  const bool mode_yields_array = (mode == 1);

  // Compile once for the whole call; every subject reuses the program.
  const CompiledPattern program(pattern, case_insensitive);

  // find_all is needed for mode 1; modes 0 and 2 want only the first match.
  auto per_subject = [&](std::string_view subject) -> Value {
    KernelResult kr = regex_kernel(program, subject, /*find_all=*/mode_yields_array,
                                   /*on_match_limit_returns_no_match=*/false);
    if (!kr.ok) {
      return kr.err;
    }
    return extract_dispatch(kr, subject, mode, arena);
  };

  if (!text_arg.is_array) {
    return per_subject(text_arg.scalar);
  }

  // Array broadcast (only meaningful for scalar-yielding modes).
  if (mode_yields_array) {
    note_unknown_error_origin();
    return Value::error(ErrorCode::Calc);
  }

  return broadcast_over_text(text_arg.array, arena, per_subject);
}

// ---------------------------------------------------------------------------
// REGEXREPLACE
// ---------------------------------------------------------------------------
Value eval_regexreplace_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                             const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 3 || arity > 5) {
    return Value::error(ErrorCode::Value);
  }

  // Signature order, so the leftmost error wins (see eval_regextest_lazy).
  Value err = Value::error(ErrorCode::Value);

  TextArg text_arg;
  std::string pattern;
  if (!resolve_text_and_pattern(call, arena, registry, ctx, text_arg, pattern, err)) {
    return err;
  }
  std::string replacement;
  if (!resolve_pattern(call.as_call_arg(2), arena, registry, ctx, replacement, err)) {
    return err;
  }
  const std::vector<ReplToken> replacement_tokens = parse_replacement(replacement);

  // occurrence (fourth arg, default 0). 0 replaces every match; a positive
  // N replaces only the N-th match counting from the start; a negative N
  // replaces only the |N|-th match counting from the end (Excel's
  // documented REGEXREPLACE occurrence semantics).
  long long occurrence = 0;
  if (arity >= 4) {
    if (!coerce_int_arg(call.as_call_arg(3), arena, registry, ctx, occurrence, err)) {
      return err;
    }
  }

  // case_sensitivity (fifth arg, default case-sensitive).
  bool case_insensitive = false;
  if (arity == 5 && !read_case_insensitive(call.as_call_arg(4), arena, registry, ctx, case_insensitive, err)) {
    return err;
  }

  // Compile once for the whole call; every subject reuses the program,
  // including the kernel scan an occurrence-N replace runs first.
  const CompiledPattern program(pattern, case_insensitive);

  auto run_one = [&](std::string_view subject) -> Value {
    if (occurrence == 0) {
      return substitute_global(program, subject, replacement_tokens, arena);
    }
    return substitute_nth(program, subject, replacement_tokens, occurrence, arena);
  };

  if (!text_arg.is_array) {
    return run_one(text_arg.scalar);
  }

  return broadcast_over_text(text_arg.array, arena, run_one);
}

}  // namespace eval
}  // namespace formulon
