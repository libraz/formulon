//
// PCRE2 program wrapper shared by the REGEX* functions: a compiled pattern with
// its match data and context, a left-to-right match walk, and the kernel that
// collects matches. Compile options and resource limits live in
// `regex_program.cpp`.

#ifndef FORMULON_EVAL_REGEX_PROGRAM_H_
#define FORMULON_EVAL_REGEX_PROGRAM_H_

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <utility>
#include <vector>

#define PCRE2_CODE_UNIT_WIDTH 8
#include <pcre2.h>

#include "utils/expected.h"
#include "utils/text_ops.h"
#include "value.h"

namespace formulon::eval::detail {

// What a REGEX* call returns when a match attempt exhausts `kMatchLimit`, and
// what the extract / replace pair returns for any other resource exhaustion
// (depth / heap limit, retained-match budget).
constexpr ErrorCode kMatchLimitError = ErrorCode::Value;
constexpr ErrorCode kResourceError = ErrorCode::Calc;

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
  CompiledPattern(std::string_view pattern, bool case_insensitive);
  ~CompiledPattern();

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
  kExhausted,  // no further match
  kStopped,    // the visitor declined to continue
  kLimitHit,   // the match limit was hit
  kError,      // any other PCRE2 error (depth / heap limit, bad offset)
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
      return ScanEnd::kLimitHit;
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

KernelResult regex_kernel(const CompiledPattern& program, std::string_view subject, bool find_all,
                          bool on_match_limit_returns_no_match);

}  // namespace formulon::eval::detail

#endif  // FORMULON_EVAL_REGEX_PROGRAM_H_
