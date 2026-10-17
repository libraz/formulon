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
//     exhausting the match limit is #VALUE!.
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
#include <atomic>
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
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/shape_ops_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Worksheet formula text length cap; matches Excel's 32 767-character
// limit on a single formula token. PCRE2 itself can compile much longer
// patterns, but Excel's REGEX argument cannot in practice exceed this,
// and capping early gives a clean #VALUE! instead of a slow compile.
constexpr std::size_t kMaxPatternBytes = 32767U;

// Conservative byte bound for the cap: a UTF-16 unit never needs more than
// 4 UTF-8 bytes, so a required substitution output beyond this many bytes
// always exceeds the cap and can be rejected before any buffer growth.
constexpr std::size_t kMaxSubstituteOutputBytes = static_cast<std::size_t>(kExcelTextCapUnits) * 4U;

// The match limit bounds backtracking iterations; the depth limit bounds
// recursion in the regex VM. Both are checked by pcre2_match itself; on
// overflow it returns PCRE2_ERROR_MATCHLIMIT or PCRE2_ERROR_DEPTHLIMIT.
constexpr std::uint32_t kMatchLimit = 1000000U;
constexpr std::uint32_t kDepthLimit = 10000U;

// Published through `regex_compile_count()`. Atomic because two worker
// threads of a parallel recalc may evaluate REGEX* cells concurrently;
// relaxed ordering is enough for a pure counter.
std::atomic<std::uint64_t> g_compile_count{0U};

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
  // REGEXEXTRACT mode 2 needs to surface #N/A on no match versus
  // #VALUE! on no capture groups.
  Value err = Value::error(ErrorCode::Value);
};

// One compiled pattern plus the match data and match context that go with
// it, owned for the lifetime of a single REGEX* call.
//
// Compilation is loop-invariant work: it costs an order of magnitude more
// than matching a short cell string, so a REGEX* call whose `text` argument
// is an N-cell range builds exactly one of these and reuses it for every
// subject. The constructor applies the same pattern guards the kernel used
// to apply inline; a rejected or uncompilable pattern leaves `ok()` false
// and every subject then surfaces the same `#VALUE!` it did before, which
// keeps a broadcast's result shape unchanged.
class CompiledPattern {
 public:
  CompiledPattern(std::string_view pattern, bool case_insensitive) {
    // Pattern guards before pcre2_compile so pathological inputs cannot
    // burn time inside the regex compiler.
    if (pattern.size() > kMaxPatternBytes) {
      return;
    }

    std::uint32_t compile_flags = PCRE2_UTF | PCRE2_UCP;
    if (case_insensitive) {
      compile_flags |= PCRE2_CASELESS;
    }

    int err_code = 0;
    PCRE2_SIZE err_offset = 0;
    code_ = pcre2_compile(reinterpret_cast<PCRE2_SPTR>(pattern.empty() ? "" : pattern.data()),
                          static_cast<PCRE2_SIZE>(pattern.size()), compile_flags, &err_code, &err_offset, nullptr);
    g_compile_count.fetch_add(1U, std::memory_order_relaxed);
    if (code_ == nullptr) {
      return;
    }

    // Inspect capture group count. PCRE2_INFO_CAPTURECOUNT excludes the
    // whole-match group, which is exactly what the kernel reports.
    pcre2_pattern_info(code_, PCRE2_INFO_CAPTURECOUNT, &capture_count_);

    match_data_ = pcre2_match_data_create_from_pattern(code_, nullptr);
    if (match_data_ == nullptr) {
      return;
    }
    context_ = pcre2_match_context_create(nullptr);
    if (context_ == nullptr) {
      return;
    }
    pcre2_set_match_limit(context_, kMatchLimit);
    pcre2_set_depth_limit(context_, kDepthLimit);
    ok_ = true;
  }

  ~CompiledPattern() {
    if (context_ != nullptr) {
      pcre2_match_context_free(context_);
    }
    if (match_data_ != nullptr) {
      pcre2_match_data_free(match_data_);
    }
    if (code_ != nullptr) {
      pcre2_code_free(code_);
    }
  }

  CompiledPattern(const CompiledPattern&) = delete;
  CompiledPattern& operator=(const CompiledPattern&) = delete;

  /// True once the pattern compiled and both PCRE2 side structures were
  /// obtained. A false value is always the caller's `#VALUE!`.
  bool ok() const noexcept { return ok_; }

  pcre2_code* code() const noexcept { return code_; }
  pcre2_match_data* match_data() const noexcept { return match_data_; }
  pcre2_match_context* context() const noexcept { return context_; }
  std::uint32_t capture_count() const noexcept { return capture_count_; }

 private:
  pcre2_code* code_ = nullptr;
  pcre2_match_data* match_data_ = nullptr;
  pcre2_match_context* context_ = nullptr;
  std::uint32_t capture_count_ = 0;
  bool ok_ = false;
};

// How a `scan_matches` walk ended.
enum class ScanEnd : std::uint8_t {
  kExhausted,   // no further match
  kStopped,     // the visitor declined to continue
  kMatchLimit,  // the match limit was hit
  kError,       // any other PCRE2 error (depth / heap limit, bad offset)
};

// Walks the matches of `program` over `subject` left to right, handing each
// one's ovector to `on_match`; a visitor returning false stops the walk. With
// `find_all` false the walk ends after the first match.
template <typename OnMatch>
ScanEnd scan_matches(const CompiledPattern& program, std::string_view subject, bool find_all, OnMatch&& on_match) {
  PCRE2_SIZE start_offset = 0;
  while (start_offset <= static_cast<PCRE2_SIZE>(subject.size())) {
    const int rc =
        pcre2_match(program.code(), reinterpret_cast<PCRE2_SPTR>(subject.data()),
                    static_cast<PCRE2_SIZE>(subject.size()), start_offset, 0, program.match_data(), program.context());
    if (rc == PCRE2_ERROR_NOMATCH) {
      break;
    }
    if (rc == PCRE2_ERROR_MATCHLIMIT) {
      return ScanEnd::kMatchLimit;
    }
    if (rc < 0) {
      return ScanEnd::kError;
    }
    const PCRE2_SIZE* ovec = pcre2_get_ovector_pointer(program.match_data());
    if (!on_match(ovec)) {
      return ScanEnd::kStopped;
    }
    if (!find_all) {
      break;
    }

    // Advance past the match. Zero-length matches at `start_offset` would
    // loop forever; bump by one whole UTF-8 codepoint rather than one byte.
    // Under PCRE2_UTF, `pcre2_match` rejects a `start_offset` that lands
    // inside a multi-byte sequence with PCRE2_ERROR_BADUTFOFFSET, so a raw
    // one-byte bump is only safe before an ASCII byte.
    const std::size_t new_start = static_cast<std::size_t>(ovec[1]);
    if (new_start <= start_offset) {
      if (start_offset >= subject.size()) {
        break;
      }
      std::size_t step = 0;
      decode_utf8_step(subject, start_offset, &step);
      start_offset = start_offset + static_cast<PCRE2_SIZE>(step == 0 ? 1 : step);
    } else {
      start_offset = new_start;
    }
  }
  return ScanEnd::kExhausted;
}

// Single shared match pipeline.
//
// `subject` must remain valid for the duration of the call. `find_all`
// decides whether the kernel iterates after the first match (REGEXTEST and
// REGEXEXTRACT modes 0 / 2 only need the first hit; everything else wants
// every hit).
//
// `on_match_limit_returns_no_match`: when true (REGEXTEST), a hit on
// depth_limit is converted to "no match" (kernel still returns ok=true with
// matches empty). When false (the extract / replace pair), it surfaces as
// #CALC!. Exhausting the match limit is #VALUE! either way.
KernelResult regex_kernel(const CompiledPattern& program, std::string_view subject, bool find_all,
                          bool on_match_limit_returns_no_match) {
  KernelResult result;

  if (!program.ok()) {
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }

  const std::uint32_t capture_count = program.capture_count();
  result.capture_count = capture_count;

  // PCRE2's limits bound one match attempt, not how much a find-all scan
  // retains across attempts. Charge each retained match its whole-match slot
  // plus one per capture group so a many-group pattern over a long subject
  // cannot accumulate without bound.
  ResourceBudget match_slots(kMaxRegexMatchSlots);
  bool budget_exhausted = false;

  const ScanEnd end = scan_matches(program, subject, find_all, [&](const PCRE2_SIZE* ovec) {
    if (!match_slots.consume(static_cast<std::uint64_t>(capture_count) + 1U)) {
      budget_exhausted = true;
      return false;
    }
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
    return true;
  });

  if (end == ScanEnd::kMatchLimit) {
    result.ok = false;
    result.err = Value::error(ErrorCode::Value);
    return result;
  }
  if (budget_exhausted || end == ScanEnd::kError) {
    // Resource exhaustion (match-slot budget, depth / heap limit): REGEXTEST
    // degrades to "no match", the extract / replace pair reports #CALC!.
    if (!on_match_limit_returns_no_match) {
      result.ok = false;
      result.err = Value::error(ErrorCode::Calc);
    }
  }
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
  ArrayValue* out = allocate_array_value(1U, n, arena, cells, kMaxDerivedArrayCells);
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
  ArrayValue* out = allocate_array_value(1U, kr.capture_count, arena, cells, kMaxDerivedArrayCells);
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

// --- REGEXREPLACE substitution -------------------------------------------
//
// The replacement string is parsed once per call into a token list and
// expanded per match. Two layers:
//
//   `$` layer       `$n` (one or two digits, greedy), `${n}`, `${name}`,
//                   `$name`, `$<name>`, `$&` (whole match), `` $` `` / `$'` (text before /
//                   after the match), `$_` (whole subject), `$$`.
//   backslash layer `\1`..`\99` and `\g<n>` / `\g<name>` (group references),
//                   `\n \t \r \e \a \f \v \b`, `\xHH`, `\x{H..}`, octal `\0HH`
//                   / `\ooo`, `\\ \$ \&`, `\u` (upper-case the next
//                   character), `\l` (lower-case the next), `\U`..`\E` /
//                   `\L`..`\E` (upper- / lower-case a run), `\Q`..`\E`
//                   (literal span).
//
// Anything else (a bare `$`, `$+`, `\0`, `\g1`, `\k<name>`, an
// unknown escape, a trailing `\`) is `#VALUE!`, as is a reference to a group
// the pattern lacks. Errors surface when a match is expanded, so a pattern
// that matches nothing leaves the subject unchanged. A group that did not
// participate expands to nothing.

enum class ReplTokenKind : std::uint8_t {
  kLiteral,
  kGroup,    // `number`
  kNamed,    // `text` is the group name
  kBefore,   // text before the match
  kAfter,    // text after the match
  kSubject,  // the whole subject
  kUpperNext,
  kLowerNext,
  kUpperRun,
  kLowerRun,
  kCaseEnd,
  kInvalid,
};

struct ReplToken {
  ReplTokenKind kind = ReplTokenKind::kLiteral;
  std::uint32_t number = 0;
  std::string text;
};

bool is_name_byte(char c) noexcept {
  return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_';
}

bool is_digit_byte(char c) noexcept {
  return c >= '0' && c <= '9';
}

bool is_octal_byte(char c) noexcept {
  return c >= '0' && c <= '7';
}

int hex_value(char c) noexcept {
  if (c >= '0' && c <= '9') {
    return c - '0';
  }
  if (c >= 'a' && c <= 'f') {
    return c - 'a' + 10;
  }
  if (c >= 'A' && c <= 'F') {
    return c - 'A' + 10;
  }
  return -1;
}

// Appends the UTF-8 encoding of `cp`; false for a surrogate or a value past
// U+10FFFF.
bool append_code_point(std::string& out, std::uint32_t cp) {
  if (cp > 0x10FFFFU || (cp >= 0xD800U && cp <= 0xDFFFU)) {
    return false;
  }
  if (cp < 0x80U) {
    out.push_back(static_cast<char>(cp));
  } else if (cp < 0x800U) {
    out.push_back(static_cast<char>(0xC0U | (cp >> 6)));
    out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
  } else if (cp < 0x10000U) {
    out.push_back(static_cast<char>(0xE0U | (cp >> 12)));
    out.push_back(static_cast<char>(0x80U | ((cp >> 6) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
  } else {
    out.push_back(static_cast<char>(0xF0U | (cp >> 18)));
    out.push_back(static_cast<char>(0x80U | ((cp >> 12) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | ((cp >> 6) & 0x3FU)));
    out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
  }
  return true;
}

// Reads `[0-9]{1,max_digits}` at `i`, advancing it.
std::uint32_t read_digits(std::string_view s, std::size_t& i, std::size_t max_digits) {
  std::uint32_t n = 0;
  for (std::size_t k = 0; k < max_digits && i < s.size() && is_digit_byte(s[i]); ++k, ++i) {
    n = n * 10U + static_cast<std::uint32_t>(s[i] - '0');
  }
  return n;
}

// A group reference spelled with a name or digits (`${..}`, `\g<..>`): all
// digits is a number, otherwise a name. Empty or non-word text is invalid.
ReplToken group_reference_token(std::string_view body) {
  ReplToken tok;
  tok.kind = ReplTokenKind::kInvalid;
  if (body.empty() || body.size() > 32U) {
    return tok;
  }
  bool all_digits = true;
  for (const char c : body) {
    if (!is_name_byte(c)) {
      return tok;
    }
    all_digits = all_digits && is_digit_byte(c);
  }
  if (all_digits) {
    std::size_t i = 0;
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(body, i, body.size());
    return tok;
  }
  tok.kind = ReplTokenKind::kNamed;
  tok.text = std::string(body);
  return tok;
}

// Parses the `$` token starting at `i` (pointing at the `$`), advancing past it.
ReplToken parse_dollar(std::string_view r, std::size_t& i) {
  ReplToken tok;
  ++i;
  if (i >= r.size()) {
    tok.kind = ReplTokenKind::kInvalid;
    return tok;
  }
  const char c = r[i];
  if (is_digit_byte(c)) {
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(r, i, 2U);
    return tok;
  }
  if (c == '{') {
    const std::size_t close = r.find('}', i);
    if (close == std::string_view::npos) {
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
    i = close + 1U;
    return tok;
  }
  if (c == '<') {
    const std::size_t close = r.find('>', i);
    if (close == std::string_view::npos) {
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
    if (tok.kind == ReplTokenKind::kGroup) {
      tok.kind = ReplTokenKind::kInvalid;  // `$<n>` names groups only.
    }
    i = close + 1U;
    return tok;
  }
  if (c == '_') {
    ++i;
    tok.kind = ReplTokenKind::kSubject;
    return tok;
  }
  if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_') {
    std::size_t end = i;
    while (end < r.size() && is_name_byte(r[end])) {
      ++end;
    }
    tok = group_reference_token(r.substr(i, end - i));
    i = end;
    return tok;
  }
  tok.kind = ReplTokenKind::kLiteral;
  ++i;
  switch (c) {
    case '$':
      tok.text = "$";
      return tok;
    case '&':
      tok.kind = ReplTokenKind::kGroup;
      return tok;
    case '`':
      tok.kind = ReplTokenKind::kBefore;
      return tok;
    case '\'':
      tok.kind = ReplTokenKind::kAfter;
      return tok;
    default:
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
  }
}

// Parses the backslash token starting at `i` (pointing at the `\`), advancing
// past it.
ReplToken parse_backslash(std::string_view r, std::size_t& i) {
  ReplToken tok;
  ++i;
  if (i >= r.size()) {
    tok.kind = ReplTokenKind::kInvalid;
    return tok;
  }
  const char c = r[i++];
  const auto literal = [&tok](char ch) {
    tok.kind = ReplTokenKind::kLiteral;
    tok.text.assign(1U, ch);
  };
  switch (c) {
    case '\\':
    case '$':
    case '&':
      literal(c);
      return tok;
    case 'n':
      literal('\n');
      return tok;
    case 't':
      literal('\t');
      return tok;
    case 'r':
      literal('\r');
      return tok;
    case 'e':
      literal('\x1B');
      return tok;
    case 'a':
      literal('\a');
      return tok;
    case 'f':
      literal('\f');
      return tok;
    case 'v':
      literal('\v');
      return tok;
    case 'b':
      literal('\b');
      return tok;
    case 'u':
      tok.kind = ReplTokenKind::kUpperNext;
      return tok;
    case 'l':
      tok.kind = ReplTokenKind::kLowerNext;
      return tok;
    case 'U':
      tok.kind = ReplTokenKind::kUpperRun;
      return tok;
    case 'L':
      tok.kind = ReplTokenKind::kLowerRun;
      return tok;
    case 'E':
      tok.kind = ReplTokenKind::kCaseEnd;
      return tok;
    case 'Q': {
      const std::size_t end = r.find("\\E", i);
      tok.kind = ReplTokenKind::kLiteral;
      tok.text = std::string(r.substr(i, end == std::string_view::npos ? std::string_view::npos : end - i));
      i = end == std::string_view::npos ? r.size() : end + 2U;
      return tok;
    }
    case 'x': {
      std::uint32_t cp = 0;
      bool ok = false;
      if (i < r.size() && r[i] == '{') {
        std::size_t k = i + 1U;
        std::size_t digits = 0;
        while (k < r.size() && hex_value(r[k]) >= 0 && digits < 8U) {
          cp = cp * 16U + static_cast<std::uint32_t>(hex_value(r[k]));
          ++k;
          ++digits;
        }
        ok = digits > 0U && k < r.size() && r[k] == '}';
        i = ok ? k + 1U : i;
      } else {
        std::size_t digits = 0;
        while (digits < 2U && i < r.size() && hex_value(r[i]) >= 0) {
          cp = cp * 16U + static_cast<std::uint32_t>(hex_value(r[i]));
          ++i;
          ++digits;
        }
        ok = digits > 0U;
      }
      tok.kind = ReplTokenKind::kLiteral;
      if (!ok || !append_code_point(tok.text, cp)) {
        tok.kind = ReplTokenKind::kInvalid;
      }
      return tok;
    }
    case 'g': {
      if (i < r.size() && r[i] == '<') {
        const std::size_t close = r.find('>', i);
        if (close != std::string_view::npos) {
          tok = group_reference_token(r.substr(i + 1U, close - i - 1U));
          i = close + 1U;
          return tok;
        }
      }
      tok.kind = ReplTokenKind::kInvalid;
      return tok;
    }
    default:
      break;
  }
  if (is_digit_byte(c)) {
    --i;  // Back onto the first digit.
    const bool octal3 = is_octal_byte(c) && i + 2U < r.size() && is_octal_byte(r[i + 1U]) && is_octal_byte(r[i + 2U]);
    if (c == '0' || octal3) {
      std::uint32_t cp = 0;
      std::size_t digits = 0;
      while (digits < 3U && i < r.size() && is_octal_byte(r[i])) {
        cp = cp * 8U + static_cast<std::uint32_t>(r[i] - '0');
        ++i;
        ++digits;
      }
      tok.kind = ReplTokenKind::kLiteral;
      if (digits < 2U || !append_code_point(tok.text, cp)) {
        tok.kind = ReplTokenKind::kInvalid;  // `\0` alone is not a reference.
      }
      return tok;
    }
    tok.kind = ReplTokenKind::kGroup;
    tok.number = read_digits(r, i, 2U);
    return tok;
  }
  tok.kind = ReplTokenKind::kInvalid;
  return tok;
}

std::vector<ReplToken> parse_replacement(std::string_view r) {
  std::vector<ReplToken> tokens;
  std::string pending;
  const auto flush = [&] {
    if (!pending.empty()) {
      ReplToken lit;
      lit.text = std::move(pending);
      pending.clear();
      tokens.push_back(std::move(lit));
    }
  };
  for (std::size_t i = 0; i < r.size();) {
    if (r[i] != '$' && r[i] != '\\') {
      pending.push_back(r[i++]);
      continue;
    }
    ReplToken tok = r[i] == '$' ? parse_dollar(r, i) : parse_backslash(r, i);
    if (tok.kind == ReplTokenKind::kLiteral) {
      pending += tok.text;
      continue;
    }
    flush();
    tokens.push_back(std::move(tok));
  }
  flush();
  return tokens;
}

// Output of one expansion, with the pending `\u` / `\l` and the `\U` / `\L`
// run applied to text as it is appended; a one-shot mapping beats the run.
class ReplacementWriter {
 public:
  explicit ReplacementWriter(std::string& out) : out_(out) {}

  void append(std::string_view text) {
    for (const char c : text) {
      CaseMode mode = run_;
      if (next_ != CaseMode::kNone) {
        mode = next_;
        next_ = CaseMode::kNone;
      }
      out_.push_back(mode == CaseMode::kUpper   ? strings::ascii_to_upper(c)
                     : mode == CaseMode::kLower ? strings::ascii_to_lower(c)
                                                : c);
    }
  }
  void next(bool upper) { next_ = upper ? CaseMode::kUpper : CaseMode::kLower; }
  void run(bool upper) { run_ = upper ? CaseMode::kUpper : CaseMode::kLower; }
  void end_run() { run_ = CaseMode::kNone; }

 private:
  enum class CaseMode : std::uint8_t { kNone, kUpper, kLower };
  std::string& out_;
  CaseMode run_ = CaseMode::kNone;
  CaseMode next_ = CaseMode::kNone;
};

// Resolves a group number from a name; -1 when absent or ambiguous.
long long group_number_for_name(const CompiledPattern& program, const std::string& name) {
  const int n = pcre2_substring_number_from_name(program.code(), reinterpret_cast<PCRE2_SPTR>(name.c_str()));
  return n < 0 ? -1 : n;
}

// Appends the expansion of `tokens` for one match to `out`. `ovec` is the
// PCRE2 ovector of the match. False when a token is invalid or names a group
// the pattern lacks.
bool expand_replacement(const std::vector<ReplToken>& tokens, const CompiledPattern& program, std::string_view subject,
                        const PCRE2_SIZE* ovec, std::string& out) {
  ReplacementWriter writer(out);
  const auto group_text = [&](std::size_t group) -> std::string_view {
    if (ovec[2 * group] == PCRE2_UNSET) {
      return {};
    }
    return subject.substr(static_cast<std::size_t>(ovec[2 * group]),
                          static_cast<std::size_t>(ovec[2 * group + 1] - ovec[2 * group]));
  };
  for (const ReplToken& tok : tokens) {
    switch (tok.kind) {
      case ReplTokenKind::kLiteral:
        writer.append(tok.text);
        break;
      case ReplTokenKind::kGroup:
        if (tok.number > program.capture_count()) {
          return false;
        }
        writer.append(group_text(tok.number));
        break;
      case ReplTokenKind::kNamed: {
        const long long group = group_number_for_name(program, tok.text);
        if (group < 0) {
          return false;
        }
        writer.append(group_text(static_cast<std::size_t>(group)));
        break;
      }
      case ReplTokenKind::kBefore:
        writer.append(subject.substr(0, static_cast<std::size_t>(ovec[0])));
        break;
      case ReplTokenKind::kAfter:
        writer.append(subject.substr(static_cast<std::size_t>(ovec[1])));
        break;
      case ReplTokenKind::kSubject:
        writer.append(subject);
        break;
      case ReplTokenKind::kUpperNext:
        writer.next(/*upper=*/true);
        break;
      case ReplTokenKind::kLowerNext:
        writer.next(/*upper=*/false);
        break;
      case ReplTokenKind::kUpperRun:
        writer.run(/*upper=*/true);
        break;
      case ReplTokenKind::kLowerRun:
        writer.run(/*upper=*/false);
        break;
      case ReplTokenKind::kCaseEnd:
        writer.end_run();
        break;
      case ReplTokenKind::kInvalid:
        return false;
    }
    if (out.size() > kMaxSubstituteOutputBytes) {
      return false;
    }
  }
  return true;
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
  if (expansion_failed || end == ScanEnd::kMatchLimit) {
    return Value::error(ErrorCode::Value);
  }
  if (end == ScanEnd::kError) {
    return Value::error(ErrorCode::Calc);
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
  ArrayValue* out = allocate_array_value(in->rows, in->cols, arena, cells, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  const std::size_t n = static_cast<std::size_t>(in->rows) * static_cast<std::size_t>(in->cols);
  // Per-cell subject buffers must outlive any result built from a span of
  // them; keep them around for the whole loop.
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
    cells[i] = per_cell(std::string_view(subjects.back()));
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

bool RegexSearch::matches(std::string_view subject) const {
  const KernelResult result = regex_kernel(impl_->program, subject, /*find_all=*/false,
                                           /*on_match_limit_returns_no_match=*/true);
  return result.ok && !result.matches.empty();
}

std::uint64_t regex_compile_count() noexcept {
  return g_compile_count.load(std::memory_order_relaxed);
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
