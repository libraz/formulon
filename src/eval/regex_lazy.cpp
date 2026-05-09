// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// REGEXTEST / REGEXEXTRACT / REGEXREPLACE backed by PCRE2 (8-bit, no JIT,
// UTF + UCP). All three functions share the `regex_kernel` helper to
// guarantee a single compile + match pipeline; this is a hard
// architectural requirement (one `pcre2_compile` call site, not three).
//
// Spec / oracle references:
//   * Excel REGEX function family signatures and shape semantics:
//     Mac Excel oracle observations.
//   * Resource limits: match_limit = 1_000_000, depth_limit = 10_000.
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

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#define PCRE2_CODE_UNIT_WIDTH 8
#include <pcre2.h>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Worksheet formula text length cap; matches Excel's 32 767-character
// limit on a single formula token. PCRE2 itself can compile much longer
// patterns, but Excel's REGEX argument cannot in practice exceed this,
// and capping early gives a clean #VALUE! instead of a slow compile.
constexpr std::size_t kMaxPatternBytes = 32767U;

// The match limit bounds backtracking iterations; the depth limit bounds
// recursion in the regex VM. Both are checked by pcre2_match itself; on
// overflow it returns PCRE2_ERROR_MATCHLIMIT or PCRE2_ERROR_DEPTHLIMIT.
constexpr std::uint32_t kMatchLimit = 1000000U;
constexpr std::uint32_t kDepthLimit = 10000U;

// --- Argument coercion helpers -------------------------------------------
//
// Every REGEX* function shares the same trailing-arg vocabulary (case
// sensitivity, return mode, occurrence). Coercion always: number coerce
// -> truncate -> bound check. A non-coercible value or out-of-domain
// integer surfaces #VALUE! (or the upstream error code).

// Shared helper: evaluates `node`, propagates errors, coerces to a
// truncated integer. On success writes the int into `out` and returns
// true. On failure writes the error sentinel into `out_err` and returns
// false.
bool coerce_int_arg(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    long long& out, Value& out_err) {
  const Value v = eval_node(node, arena, registry, ctx);
  if (v.is_error()) {
    out_err = v;
    return false;
  }
  auto num = coerce_to_number(v);
  if (!num) {
    out_err = Value::error(num.error());
    return false;
  }
  const double d = std::trunc(num.value());
  // Defend against NaN / Inf surviving truncation.
  if (!(d == d) || d > static_cast<double>(0x7FFFFFFF) || d < -static_cast<double>(0x7FFFFFFF)) {
    out_err = Value::error(ErrorCode::Value);
    return false;
  }
  out = static_cast<long long>(d);
  return true;
}

// --- Match-result accumulator --------------------------------------------
//
// The kernel is a generator: it yields one match at a time to the
// per-function consumer. We model this by having the kernel iterate
// internally and store offsets into a small_vector-equivalent
// (arena-backed std::vector is fine here because the impl is a leaf
// TU and the sizes are bounded by the subject length).

struct MatchSpan {
  std::size_t whole_start;
  std::size_t whole_end;
  // Per match: offsets for each capture group (group 0 is the whole
  // match, group i = i-th `(...)`). PCRE2 uses (PCRE2_SIZE)-1 to mean
  // "did not participate"; we map that to {0, 0} and the caller
  // distinguishes via the `participated` bitset.
  std::vector<std::pair<std::size_t, std::size_t>> groups;
  std::vector<bool> participated;
};

// Result of `regex_kernel`. `compiled` and `match_data` are owned by
// the kernel and freed before return; the caller only sees the matches
// vector and the capture-group count. `error` is set on any failure
// (compile failure, match-limit hit, allocation failure).
struct KernelResult {
  std::vector<MatchSpan> matches;
  std::uint32_t capture_count = 0;  // # of `(...)` groups, EXCLUDING group 0
  bool ok = true;
  // Error to return when ok == false. Distinct from a "no match" (which
  // is ok = true with matches empty); the distinction matters because
  // REGEXEXTRACT mode 2/3 needs to surface #N/A on no match versus
  // #VALUE! on no capture groups.
  Value err = Value::error(ErrorCode::Value);
};

// Single shared compile + match pipeline.
//
// `subject` and `pattern` are passed as string_view; both must remain
// valid for the duration of the call. `case_insensitive` controls the
// PCRE2_CASELESS compile flag. `find_all` decides whether the kernel
// iterates after the first match (REGEXTEST and the occurrence-N
// variant of REGEXREPLACE only need the first hit; everything else
// wants every hit).
//
// `on_match_limit_returns_no_match`: when true (REGEXTEST), a hit on
// match_limit / depth_limit is converted to "no match" (kernel still
// returns ok=true with matches empty). When false (the extract /
// replace pair), it surfaces as #CALC!.
KernelResult regex_kernel(std::string_view subject, std::string_view pattern, bool case_insensitive, bool find_all,
                          bool on_match_limit_returns_no_match) {
  KernelResult result;

  // Pattern guards before pcre2_compile so pathological inputs cannot
  // burn time inside the regex compiler.
  if (pattern.empty()) {
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }
  if (pattern.size() > kMaxPatternBytes) {
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }

  std::uint32_t compile_flags = PCRE2_UTF | PCRE2_UCP;
  if (case_insensitive) {
    compile_flags |= PCRE2_CASELESS;
  }

  int err_code = 0;
  PCRE2_SIZE err_offset = 0;
  pcre2_code* code =
      pcre2_compile(reinterpret_cast<PCRE2_SPTR>(pattern.data()), static_cast<PCRE2_SIZE>(pattern.size()),
                    compile_flags, &err_code, &err_offset, nullptr);
  if (code == nullptr) {
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }

  // Inspect capture group count. PCRE2_INFO_CAPTURECOUNT excludes the
  // whole-match group, which is exactly what we want to return.
  std::uint32_t capture_count = 0;
  pcre2_pattern_info(code, PCRE2_INFO_CAPTURECOUNT, &capture_count);
  result.capture_count = capture_count;

  pcre2_match_data* match_data = pcre2_match_data_create_from_pattern(code, nullptr);
  if (match_data == nullptr) {
    pcre2_code_free(code);
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }

  pcre2_match_context* mctx = pcre2_match_context_create(nullptr);
  if (mctx == nullptr) {
    pcre2_match_data_free(match_data);
    pcre2_code_free(code);
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }
  pcre2_set_match_limit(mctx, kMatchLimit);
  pcre2_set_depth_limit(mctx, kDepthLimit);

  PCRE2_SIZE start_offset = 0;
  while (start_offset <= static_cast<PCRE2_SIZE>(subject.size())) {
    const int rc = pcre2_match(code, reinterpret_cast<PCRE2_SPTR>(subject.data()),
                               static_cast<PCRE2_SIZE>(subject.size()), start_offset, 0, match_data, mctx);
    if (rc == PCRE2_ERROR_NOMATCH) {
      break;
    }
    if (rc < 0) {
      // Resource exhaustion: PCRE2_ERROR_MATCHLIMIT / DEPTHLIMIT /
      // HEAPLIMIT. Treat as a graceful no-match for REGEXTEST (the
      // predicate semantics make a soft FALSE more useful than an
      // error), and as #CALC! for the extract / replace pair.
      if (on_match_limit_returns_no_match) {
        break;
      }
      result.ok = false;
      result.err = Value::error(ErrorCode::Calc);
      pcre2_match_context_free(mctx);
      pcre2_match_data_free(match_data);
      pcre2_code_free(code);
      return result;
    }

    // rc > 0: rc-1 capture groups participated in addition to the whole
    // match. PCRE2's ovector layout: pairs of (start, end) for groups
    // 0..capture_count.
    const PCRE2_SIZE* ovec = pcre2_get_ovector_pointer(match_data);
    MatchSpan ms;
    ms.whole_start = static_cast<std::size_t>(ovec[0]);
    ms.whole_end = static_cast<std::size_t>(ovec[1]);
    ms.groups.reserve(capture_count);
    ms.participated.reserve(capture_count);
    for (std::uint32_t g = 1; g <= capture_count; ++g) {
      const PCRE2_SIZE gs = ovec[2 * g];
      const PCRE2_SIZE ge = ovec[2 * g + 1];
      if (gs == PCRE2_UNSET) {
        ms.groups.emplace_back(0U, 0U);
        ms.participated.push_back(false);
      } else {
        ms.groups.emplace_back(static_cast<std::size_t>(gs), static_cast<std::size_t>(ge));
        ms.participated.push_back(true);
      }
    }
    result.matches.push_back(std::move(ms));

    if (!find_all) {
      break;
    }

    // Advance past the match. Zero-length matches at `start_offset`
    // would loop forever; bump by one byte (UTF-8 safe because PCRE2
    // refuses to match in the middle of a multi-byte sequence under
    // PCRE2_UTF, so the next byte is a valid restart).
    const std::size_t new_start = static_cast<std::size_t>(ovec[1]);
    if (new_start <= start_offset) {
      if (start_offset >= subject.size()) {
        break;
      }
      start_offset = start_offset + 1;
    } else {
      start_offset = new_start;
    }
  }

  pcre2_match_context_free(mctx);
  pcre2_match_data_free(match_data);
  pcre2_code_free(code);
  return result;
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
// REGEX* functions accept Range / Ref / ArrayLiteral as `text`, in
// which case the result is array-shaped (one regex evaluation per
// cell). This helper materialises the arg as either a scalar string
// (when it's a literal / arithmetic / function call producing a
// scalar) or an Array of cells.
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
  // Inspect the AST shape to decide between scalar and array contexts.
  // Range / Ref / ArrayLiteral go through `eval_node_as_array` so the
  // (rows, cols) shape is preserved; everything else evaluates as a
  // scalar via `eval_node`. This mirrors the convention used by the
  // shape-preserving lazy impls (TEXTSPLIT / FILTER / TRIMRANGE).
  const parser::NodeKind k = node.kind();
  const bool array_shape =
      k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp || k == parser::NodeKind::ArrayLiteral;
  if (array_shape) {
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
  out.is_array = false;
  out.scalar = std::move(t.value());
  return true;
}

// Resolves the pattern argument (scalar text, errors propagate). Empty
// pattern is enforced inside `regex_kernel` so callers don't have to
// duplicate the early guard.
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

// mode 1: column array (N x 1) of all whole matches.
Value extract_mode1(const KernelResult& kr, std::string_view subject, Arena& arena) {
  if (kr.matches.empty()) {
    return Value::error(ErrorCode::NA);
  }
  const std::uint32_t n = static_cast<std::uint32_t>(kr.matches.size());
  Value* cells = arena.create_array<Value>(n);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    const MatchSpan& m = kr.matches[i];
    cells[i] = text_from_span(subject, m.whole_start, m.whole_end, arena);
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = n;
  out->cols = 1U;
  out->cells = cells;
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
  Value* cells = arena.create_array<Value>(kr.capture_count);
  if (cells == nullptr) {
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
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = 1U;
  out->cols = kr.capture_count;
  out->cells = cells;
  return Value::array(out);
}

// mode 3: 2D array (N x G). Same #N/A / #VALUE! split as mode 2.
Value extract_mode3(const KernelResult& kr, std::string_view subject, Arena& arena) {
  if (kr.matches.empty()) {
    return Value::error(ErrorCode::NA);
  }
  if (kr.capture_count == 0U) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t n = static_cast<std::uint32_t>(kr.matches.size());
  const std::size_t total = static_cast<std::size_t>(n) * static_cast<std::size_t>(kr.capture_count);
  Value* cells = arena.create_array<Value>(total);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t i = 0; i < n; ++i) {
    const MatchSpan& m = kr.matches[i];
    for (std::uint32_t g = 0; g < kr.capture_count; ++g) {
      const std::size_t idx = static_cast<std::size_t>(i) * static_cast<std::size_t>(kr.capture_count) + g;
      if (!m.participated[g]) {
        cells[idx] = Value::text(arena.intern(""));
      } else {
        cells[idx] = text_from_span(subject, m.groups[g].first, m.groups[g].second, arena);
      }
    }
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = n;
  out->cols = kr.capture_count;
  out->cells = cells;
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
    case 3:
      return extract_mode3(kr, subject, arena);
    default:
      return Value::error(ErrorCode::Value);
  }
}

// --- REGEXREPLACE substitution -------------------------------------------
//
// pcre2_substitute does the heavy lifting. We always pass
// PCRE2_SUBSTITUTE_EXTENDED so Excel's `$1`, `${name}`, `$$`, and `\n`
// escapes work; we add PCRE2_SUBSTITUTE_GLOBAL only when occurrence == 0.

// Runs one pcre2_substitute call with caller-selected global/single flags.
Value substitute_with_flags(std::string_view subject, std::string_view pattern, std::string_view replacement,
                            bool case_insensitive, std::size_t start_offset, std::uint32_t sub_flags, Arena& arena) {
  // Pattern guards mirror regex_kernel.
  if (pattern.empty() || pattern.size() > kMaxPatternBytes) {
    return Value::error(ErrorCode::Value);
  }

  std::uint32_t compile_flags = PCRE2_UTF | PCRE2_UCP;
  if (case_insensitive) {
    compile_flags |= PCRE2_CASELESS;
  }
  int err_code = 0;
  PCRE2_SIZE err_offset = 0;
  pcre2_code* code =
      pcre2_compile(reinterpret_cast<PCRE2_SPTR>(pattern.data()), static_cast<PCRE2_SIZE>(pattern.size()),
                    compile_flags, &err_code, &err_offset, nullptr);
  if (code == nullptr) {
    return Value::error(ErrorCode::Value);
  }

  pcre2_match_context* mctx = pcre2_match_context_create(nullptr);
  if (mctx == nullptr) {
    pcre2_code_free(code);
    return Value::error(ErrorCode::Value);
  }
  pcre2_set_match_limit(mctx, kMatchLimit);
  pcre2_set_depth_limit(mctx, kDepthLimit);

  // Sizing: start with subject + replacement * 2 as a guess. PCRE2 will
  // tell us the required size via PCRE2_ERROR_NOMEMORY if the buffer
  // is too small; we grow once on that signal.
  std::size_t bufsize = subject.size() + replacement.size() * 2 + 16;
  std::vector<unsigned char> buffer(bufsize);
  PCRE2_SIZE outlen = bufsize;

  int rc = pcre2_substitute(code, reinterpret_cast<PCRE2_SPTR>(subject.data()), static_cast<PCRE2_SIZE>(subject.size()),
                            static_cast<PCRE2_SIZE>(start_offset), sub_flags, nullptr, mctx,
                            reinterpret_cast<PCRE2_SPTR>(replacement.data()),
                            static_cast<PCRE2_SIZE>(replacement.size()), buffer.data(), &outlen);
  if (rc == PCRE2_ERROR_NOMEMORY) {
    // outlen now contains the required size; reallocate and retry.
    buffer.resize(outlen);
    bufsize = outlen;
    outlen = bufsize;
    rc = pcre2_substitute(code, reinterpret_cast<PCRE2_SPTR>(subject.data()), static_cast<PCRE2_SIZE>(subject.size()),
                          static_cast<PCRE2_SIZE>(start_offset), sub_flags, nullptr, mctx,
                          reinterpret_cast<PCRE2_SPTR>(replacement.data()), static_cast<PCRE2_SIZE>(replacement.size()),
                          buffer.data(), &outlen);
  }

  pcre2_match_context_free(mctx);
  pcre2_code_free(code);

  if (rc < 0) {
    // Match-limit / depth-limit / malformed replacement.
    if (rc == PCRE2_ERROR_MATCHLIMIT || rc == PCRE2_ERROR_DEPTHLIMIT || rc == PCRE2_ERROR_HEAPLIMIT) {
      return Value::error(ErrorCode::Calc);
    }
    return Value::error(ErrorCode::Value);
  }
  // rc == 0 means no matches were found — return original text unchanged.
  std::string_view out_view(reinterpret_cast<const char*>(buffer.data()), static_cast<std::size_t>(outlen));
  return Value::text(arena.intern(out_view));
}

Value substitute_global(std::string_view subject, std::string_view pattern, std::string_view replacement,
                        bool case_insensitive, Arena& arena) {
  return substitute_with_flags(subject, pattern, replacement, case_insensitive, /*start_offset=*/0U,
                               PCRE2_SUBSTITUTE_EXTENDED | PCRE2_SUBSTITUTE_GLOBAL, arena);
}

// Replaces only the N-th match (occurrence == N > 0). Strategy: run the
// kernel with find_all = true (subject to N matches), then substitute
// at exactly the chosen match by passing `start_offset = match_start`
// and NOT setting GLOBAL.
Value substitute_nth(std::string_view subject, std::string_view pattern, std::string_view replacement,
                     bool case_insensitive, long long occurrence, Arena& arena) {
  KernelResult kr = regex_kernel(subject, pattern, case_insensitive, /*find_all=*/true,
                                 /*on_match_limit_returns_no_match=*/false);
  if (!kr.ok) {
    return kr.err;
  }
  if (occurrence < 1 || static_cast<std::size_t>(occurrence) > kr.matches.size()) {
    // N exceeds total matches OR no matches found — return original.
    return Value::text(arena.intern(subject));
  }

  const std::size_t target_start = kr.matches[static_cast<std::size_t>(occurrence - 1)].whole_start;
  return substitute_with_flags(subject, pattern, replacement, case_insensitive, target_start, PCRE2_SUBSTITUTE_EXTENDED,
                               arena);
}

}  // namespace

// ---------------------------------------------------------------------------
// REGEXTEST
// ---------------------------------------------------------------------------
Value eval_regextest_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2 || arity > 3) {
    return Value::error(ErrorCode::Value);
  }

  // Optional case_sensitivity (third arg). Coerce + bound to {0, 1};
  // out-of-range -> #VALUE!. Mac Excel 365 convention (matching MS docs):
  // 0/FALSE (default) = case-sensitive; 1/TRUE = case-insensitive.
  bool case_insensitive = false;
  if (arity == 3) {
    long long cs = 0;
    Value err = Value::error(ErrorCode::Value);
    if (!coerce_int_arg(call.as_call_arg(2), arena, registry, ctx, cs, err)) {
      return err;
    }
    if (cs != 0 && cs != 1) {
      return Value::error(ErrorCode::Value);
    }
    case_insensitive = (cs == 1);
  }

  // Pattern (scalar text).
  std::string pattern;
  Value err = Value::error(ErrorCode::Value);
  if (!resolve_pattern(call.as_call_arg(1), arena, registry, ctx, pattern, err)) {
    return err;
  }

  // Text (scalar or array). Array shape -> per-cell broadcast.
  TextArg text_arg;
  if (!resolve_text_arg(call.as_call_arg(0), arena, registry, ctx, text_arg, err)) {
    return err;
  }

  if (!text_arg.is_array) {
    KernelResult kr = regex_kernel(text_arg.scalar, pattern, case_insensitive, /*find_all=*/false,
                                   /*on_match_limit_returns_no_match=*/true);
    if (!kr.ok) {
      return kr.err;
    }
    return Value::boolean(!kr.matches.empty());
  }

  // Array broadcast: one regex evaluation per cell, output shape =
  // input shape.
  const ArrayValue* in = text_arg.array;
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  Value* cells = arena.create_array<Value>(n);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }
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
    KernelResult kr = regex_kernel(t.value(), pattern, case_insensitive, /*find_all=*/false,
                                   /*on_match_limit_returns_no_match=*/true);
    if (!kr.ok) {
      cells[i] = kr.err;
      continue;
    }
    cells[i] = Value::boolean(!kr.matches.empty());
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = in->rows;
  out->cols = in->cols;
  out->cells = cells;
  return Value::array(out);
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

  // return_mode (third arg, default 0).
  long long mode = 0;
  Value err = Value::error(ErrorCode::Value);
  if (arity >= 3) {
    if (!coerce_int_arg(call.as_call_arg(2), arena, registry, ctx, mode, err)) {
      return err;
    }
    if (mode < 0 || mode > 3) {
      return Value::error(ErrorCode::Value);
    }
  }

  // case_sensitivity (fourth arg, default 0). Mac Excel 365 convention:
  // 0 = case-sensitive (default), 1 = case-insensitive.
  bool case_insensitive = false;
  if (arity == 4) {
    long long cs = 0;
    if (!coerce_int_arg(call.as_call_arg(3), arena, registry, ctx, cs, err)) {
      return err;
    }
    if (cs != 0 && cs != 1) {
      return Value::error(ErrorCode::Value);
    }
    case_insensitive = (cs == 1);
  }

  std::string pattern;
  if (!resolve_pattern(call.as_call_arg(1), arena, registry, ctx, pattern, err)) {
    return err;
  }

  TextArg text_arg;
  if (!resolve_text_arg(call.as_call_arg(0), arena, registry, ctx, text_arg, err)) {
    return err;
  }

  // For modes 1 and 3 a single match already yields an array — so
  // running them in array context (per-cell broadcast) produces an
  // array of arrays which Excel/Formulon do not represent.
  // Mac Excel surfaces #CALC! in that situation; we follow suit by
  // collapsing the broadcast to per-cell scalars only when the mode
  // is 0 or 2 (both produce a single scalar / row), and rejecting
  // arrays for modes 1 / 3.
  const bool mode_yields_array = (mode == 1 || mode == 3);

  if (!text_arg.is_array) {
    // find_all is needed for modes 1 and 3; modes 0 and 2 want only the
    // first match. Cheap to over-collect, so we always find_all. The
    // kernel iterates internally — bounded by subject length.
    const bool find_all = (mode == 1 || mode == 3);
    KernelResult kr = regex_kernel(text_arg.scalar, pattern, case_insensitive, find_all,
                                   /*on_match_limit_returns_no_match=*/false);
    if (!kr.ok) {
      return kr.err;
    }
    return extract_dispatch(kr, text_arg.scalar, mode, arena);
  }

  // Array broadcast (only meaningful for scalar-yielding modes).
  if (mode_yields_array) {
    return Value::error(ErrorCode::Calc);
  }

  const ArrayValue* in = text_arg.array;
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  Value* cells = arena.create_array<Value>(n);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  // Per-cell subject buffers must outlive the text_from_span calls
  // below; keep them around for the whole loop.
  std::vector<std::string> subjects;
  subjects.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    if (cell.is_error()) {
      cells[i] = cell;
      subjects.emplace_back();
      continue;
    }
    auto t = coerce_to_text(cell);
    if (!t) {
      cells[i] = Value::error(t.error());
      subjects.emplace_back();
      continue;
    }
    subjects.push_back(std::move(t.value()));
    KernelResult kr = regex_kernel(subjects.back(), pattern, case_insensitive, /*find_all=*/(mode == 1 || mode == 3),
                                   /*on_match_limit_returns_no_match=*/false);
    if (!kr.ok) {
      cells[i] = kr.err;
      continue;
    }
    cells[i] = extract_dispatch(kr, subjects.back(), mode, arena);
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = in->rows;
  out->cols = in->cols;
  out->cells = cells;
  return Value::array(out);
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

  // occurrence (fourth arg, default 0). Mac Excel 365 is permissive on
  // negative values: REGEXREPLACE("abc", "a", "x", -1) returns "xbc"
  // (one substitution), matching the global behavior on this single-
  // match input. Clamp negative occurrence to 0 (global) to match Mac.
  long long occurrence = 0;
  Value err = Value::error(ErrorCode::Value);
  if (arity >= 4) {
    if (!coerce_int_arg(call.as_call_arg(3), arena, registry, ctx, occurrence, err)) {
      return err;
    }
    if (occurrence < 0) {
      occurrence = 0;
    }
  }

  // case_sensitivity (fifth arg, default 0). Mac Excel 365 convention:
  // 0 = case-sensitive (default), 1 = case-insensitive.
  bool case_insensitive = false;
  if (arity == 5) {
    long long cs = 0;
    if (!coerce_int_arg(call.as_call_arg(4), arena, registry, ctx, cs, err)) {
      return err;
    }
    if (cs != 0 && cs != 1) {
      return Value::error(ErrorCode::Value);
    }
    case_insensitive = (cs == 1);
  }

  std::string pattern;
  if (!resolve_pattern(call.as_call_arg(1), arena, registry, ctx, pattern, err)) {
    return err;
  }
  std::string replacement;
  if (!resolve_pattern(call.as_call_arg(2), arena, registry, ctx, replacement, err)) {
    return err;
  }

  TextArg text_arg;
  if (!resolve_text_arg(call.as_call_arg(0), arena, registry, ctx, text_arg, err)) {
    return err;
  }

  auto run_one = [&](std::string_view subject) -> Value {
    if (occurrence == 0) {
      return substitute_global(subject, pattern, replacement, case_insensitive, arena);
    }
    return substitute_nth(subject, pattern, replacement, case_insensitive, occurrence, arena);
  };

  if (!text_arg.is_array) {
    return run_one(text_arg.scalar);
  }

  const ArrayValue* in = text_arg.array;
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  Value* cells = arena.create_array<Value>(n);
  if (cells == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::vector<std::string> subjects;
  subjects.reserve(n);
  for (std::size_t i = 0; i < n; ++i) {
    const Value& cell = in->cells[i];
    if (cell.is_error()) {
      cells[i] = cell;
      subjects.emplace_back();
      continue;
    }
    auto t = coerce_to_text(cell);
    if (!t) {
      cells[i] = Value::error(t.error());
      subjects.emplace_back();
      continue;
    }
    subjects.push_back(std::move(t.value()));
    cells[i] = run_one(subjects.back());
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  out->rows = in->rows;
  out->cols = in->cols;
  out->cells = cells;
  return Value::array(out);
}

}  // namespace eval
}  // namespace formulon
