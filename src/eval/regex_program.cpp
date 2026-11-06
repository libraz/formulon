//
// PCRE2 program wrapper for the REGEX* functions: compile options (UTF + UCP,
// CASELESS on request), the pattern length cap, match / depth limits, and the
// match-collecting kernel. See `regex_program.h`.

#include "eval/regex_program.h"

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/regex_lazy.h"
#include "utils/resource_budget.h"

namespace formulon::eval {
namespace detail {
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

// Published through `regex_compile_count()`. Atomic because two worker
// threads of a parallel recalc may evaluate REGEX* cells concurrently;
// relaxed ordering is enough for a pure counter.
std::atomic<std::uint64_t> g_compile_count{0U};

}  // namespace

CompiledPattern::CompiledPattern(std::string_view pattern, bool case_insensitive) {
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

CompiledPattern::~CompiledPattern() {
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

  if (end == ScanEnd::kLimitHit) {
    result.ok = false;
    result.err = Value::error(kMatchLimitError);
    return result;
  }
  if (budget_exhausted || end == ScanEnd::kError) {
    // Resource exhaustion (match-slot budget, depth / heap limit): REGEXTEST
    // degrades to "no match", the extract / replace pair reports #CALC!.
    if (!on_match_limit_returns_no_match) {
      result.ok = false;
      result.err = Value::error(kResourceError);
    }
  }
  return result;
}

}  // namespace detail

std::uint64_t regex_compile_count() noexcept {
  return detail::g_compile_count.load(std::memory_order_relaxed);
}

}  // namespace formulon::eval
