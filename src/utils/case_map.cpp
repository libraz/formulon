//
// Case mapping over the generated run tables. See `utils/case_map.h`.

#include "utils/case_map.h"

#include <cstddef>
#include <cstdint>

#include "utils/case_map_table.h"
#include "utils/text_ops.h"

namespace formulon {

namespace {

constexpr std::uint32_t kCapitalSigma = 0x03A3U;
constexpr std::uint32_t kSmallSigma = 0x03C3U;
constexpr std::uint32_t kFinalSigma = 0x03C2U;
constexpr std::uint32_t kLastBmp = 0xFFFFU;

// Looks `cp` up in `runs`; true with `*out` set when a run covers it.
bool lookup(case_map::CaseRuns runs, std::uint32_t cp, std::uint32_t* out) noexcept {
  if (cp > kLastBmp) {
    return false;
  }
  std::size_t lo = 0;
  std::size_t hi = runs.count;
  while (lo < hi) {
    const std::size_t mid = lo + (hi - lo) / 2U;
    const case_map::CaseRun& run = runs.runs[mid];
    if (cp < run.first) {
      hi = mid;
    } else if (cp > run.last) {
      lo = mid + 1U;
    } else {
      if ((cp - run.first) % run.stride != 0U) {
        return false;
      }
      *out = static_cast<std::uint32_t>(static_cast<std::int32_t>(cp) + run.delta);
      return true;
    }
  }
  return false;
}

// The mapped code point of `cp` in `runs`, or `cp` itself when no run covers it.
std::uint32_t map_with(case_map::CaseRuns runs, std::uint32_t cp) noexcept {
  std::uint32_t mapped = cp;
  lookup(runs, cp, &mapped);
  return mapped;
}

bool is_letter(std::uint32_t cp) noexcept {
  if (cp > kLastBmp) {
    return false;
  }
  const case_map::LetterRanges ranges = case_map::letter_ranges();
  std::size_t lo = 0;
  std::size_t hi = ranges.count;
  while (lo < hi) {
    const std::size_t mid = lo + (hi - lo) / 2U;
    if (cp < ranges.pairs[2U * mid]) {
      hi = mid;
    } else if (cp > ranges.pairs[2U * mid + 1U]) {
      lo = mid + 1U;
    } else {
      return true;
    }
  }
  return false;
}

// True when a capital sigma before `next` takes its final form: the end of the
// text, a control space, or ASCII punctuation other than `( ) * ^ ` ~`.
bool sigma_is_final(std::string_view text, std::size_t next_index) noexcept {
  if (next_index >= text.size()) {
    return true;
  }
  std::size_t bytes = 0;
  const std::uint32_t next = decode_utf8_step(text, next_index, &bytes);
  if (next == 0x09U || next == 0x0AU || next == 0x0DU) {
    return true;
  }
  if (next < 0x20U || next > 0x7DU) {
    return false;
  }
  return !((next >= '0' && next <= '9') || (next >= 'A' && next <= 'Z') || (next >= 'a' && next <= 'z') ||
           next == '(' || next == ')' || next == '*' || next == '^' || next == '`');
}

// Lower-case form of `cp` inside `text`, whose code point ends at `end_index`.
std::uint32_t lower_in_context(std::string_view text, std::uint32_t cp, std::size_t end_index) noexcept {
  if (cp == kCapitalSigma) {
    return sigma_is_final(text, end_index) ? kFinalSigma : kSmallSigma;
  }
  return map_with(case_map::lower_runs(), cp);
}

void append_code_point(std::string& out, std::string_view text, std::size_t begin, std::size_t bytes,
                       std::uint32_t original, std::uint32_t mapped) {
  if (mapped == original) {
    out.append(text.substr(begin, bytes));
  } else {
    out.append(encode_utf8_codepoint(mapped));
  }
}

}  // namespace

std::uint32_t case_map_upper(std::uint32_t cp) noexcept {
  return map_with(case_map::upper_runs(), cp);
}

std::uint32_t case_map_lower_isolated(std::uint32_t cp) noexcept {
  return cp == kCapitalSigma ? kFinalSigma : map_with(case_map::lower_runs(), cp);
}

std::string case_map_to_upper(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size();) {
    std::size_t bytes = 0;
    const std::uint32_t cp = decode_utf8_step(text, i, &bytes);
    const bool valid = cp != 0xFFFDU || bytes != 1U;
    append_code_point(out, text, i, bytes, cp, valid ? case_map_upper(cp) : cp);
    i += bytes;
  }
  return out;
}

std::string case_map_to_lower(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  for (std::size_t i = 0; i < text.size();) {
    std::size_t bytes = 0;
    const std::uint32_t cp = decode_utf8_step(text, i, &bytes);
    const bool valid = cp != 0xFFFDU || bytes != 1U;
    append_code_point(out, text, i, bytes, cp, valid ? lower_in_context(text, cp, i + bytes) : cp);
    i += bytes;
  }
  return out;
}

std::string case_map_to_proper(std::string_view text) {
  std::string out;
  out.reserve(text.size());
  bool word_start = true;
  for (std::size_t i = 0; i < text.size();) {
    std::size_t bytes = 0;
    const std::uint32_t cp = decode_utf8_step(text, i, &bytes);
    if (cp == 0xFFFDU && bytes == 1U) {
      out.push_back(text[i]);
      ++i;
      word_start = true;
      continue;
    }
    const bool letter = is_letter(cp);
    const case_map::CaseRuns exceptions = word_start ? case_map::proper_start_runs() : case_map::proper_mid_runs();
    std::uint32_t mapped = cp;
    if (!lookup(exceptions, cp, &mapped) && letter) {
      mapped = word_start ? case_map_upper(cp) : lower_in_context(text, cp, i + bytes);
    }
    append_code_point(out, text, i, bytes, cp, mapped);
    word_start = !letter;
    i += bytes;
  }
  return out;
}

}  // namespace formulon
