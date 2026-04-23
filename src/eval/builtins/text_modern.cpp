// Copyright 2026 libraz. Licensed under the MIT License.
//
// Modern text accessor family: TEXTBEFORE / TEXTAFTER. Extracted from the
// core text TU to keep the catalog's translation units roughly balanced.
// The shared match-locator (`find_text_instance`) and the trailing-argument
// reader (`read_tba_opts`) live in this file's anonymous namespace because
// they are referenced only by TEXTBEFORE and TEXTAFTER.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtins/text_detail.h"
#include "eval/coerce.h"
#include "eval/text_ops.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// --- TEXTBEFORE / TEXTAFTER shared core ---------------------------------
//
// Excel's TEXTBEFORE / TEXTAFTER locate the Nth occurrence of a delimiter in
// `text` and return everything before or after it. Semantics (Excel 365):
//
//   * `instance_num` defaults to 1. 0 -> `#VALUE!`. Negative counts from
//     the end: -1 is the last occurrence, -2 the second-to-last, etc.
//   * `match_mode` 0 (default) = case-sensitive, 1 = case-insensitive; any
//     other value -> `#VALUE!`.
//   * `match_end` 0 (default) = only literal delimiter hits count; 1 = the
//     start and end of `text` also participate as virtual matches, so
//     `TEXTBEFORE("abc", "-", 1, 0, 1) = "abc"` (end-of-text virtual hit).
//     Any other value -> `#VALUE!`.
//   * `if_not_found` optional: when the requested instance does not exist,
//     return this value instead of `#N/A`.
//   * Empty `delimiter` -> `#VALUE!` (matches Excel).
//
// The shared helper locates the match window for the Nth hit and returns
// its byte range. The caller (TEXTBEFORE or TEXTAFTER) substrings around
// it. For virtual matches under `match_end=1`, TEXTBEFORE matching the
// start sentinel yields `""`, and TEXTAFTER matching the end sentinel
// yields `""`.

struct TextMatchResult {
  bool found = false;
  std::size_t match_start_byte = 0;
  std::size_t match_end_byte = 0;  // one past the last byte of the match
};

// Searches `text` for the Nth occurrence of `delimiter`, honouring
// `match_mode` and `match_end`. Sets `out_err` if the arguments are
// malformed. Returns `{found=false}` when the instance does not exist.
TextMatchResult find_text_instance(std::string_view text, std::string_view delimiter, int instance_num, int match_mode,
                                   int match_end, bool* out_err_value) {
  if (delimiter.empty() || instance_num == 0 || (match_mode != 0 && match_mode != 1) ||
      (match_end != 0 && match_end != 1)) {
    *out_err_value = true;
    return {};
  }
  *out_err_value = false;
  // Collect all literal delimiter byte offsets.
  std::vector<std::pair<std::size_t, std::size_t>> hits;  // [start, end)
  {
    std::string lowered_text;
    std::string lowered_delim;
    std::string_view hay = text;
    std::string_view needle = delimiter;
    if (match_mode == 1) {
      lowered_text = to_lower_ascii(text);
      lowered_delim = to_lower_ascii(delimiter);
      hay = lowered_text;
      needle = lowered_delim;
    }
    std::size_t i = 0;
    while (i <= hay.size()) {
      const std::size_t pos = hay.find(needle, i);
      if (pos == std::string_view::npos) {
        break;
      }
      hits.emplace_back(pos, pos + needle.size());
      // Advance past the match to avoid overlapping counts.
      i = pos + needle.size();
      if (needle.empty()) {
        break;  // defensive; delimiter.empty() is rejected above
      }
    }
  }
  // Virtual start/end sentinels when match_end=1.
  const bool use_virtual = match_end == 1;
  const std::size_t virtual_start_start = 0;
  const std::size_t virtual_start_end = 0;
  const std::size_t virtual_end_start = text.size();
  const std::size_t virtual_end_end = text.size();
  // Build the ordered list of candidate (start, end) pairs.
  std::vector<std::pair<std::size_t, std::size_t>> ordered;
  ordered.reserve(hits.size() + (use_virtual ? 2u : 0u));
  if (use_virtual) {
    ordered.emplace_back(virtual_start_start, virtual_start_end);
  }
  for (const auto& h : hits) {
    ordered.push_back(h);
  }
  if (use_virtual) {
    ordered.emplace_back(virtual_end_start, virtual_end_end);
  }
  if (ordered.empty()) {
    return {};
  }
  std::size_t idx = 0;
  if (instance_num > 0) {
    if (static_cast<std::size_t>(instance_num) > ordered.size()) {
      return {};
    }
    idx = static_cast<std::size_t>(instance_num - 1);
  } else {
    const auto k = static_cast<std::size_t>(-instance_num);
    if (k > ordered.size()) {
      return {};
    }
    idx = ordered.size() - k;
  }
  TextMatchResult r;
  r.found = true;
  r.match_start_byte = ordered[idx].first;
  r.match_end_byte = ordered[idx].second;
  return r;
}

// Reads the optional (instance_num, match_mode, match_end, if_not_found)
// trailing args common to TEXTBEFORE and TEXTAFTER. Writes the three int
// params; returns the not-found fallback Value if the caller supplied one
// (indicated by `has_if_not_found`).
struct TextBeforeAfterOpts {
  int instance_num = 1;
  int match_mode = 0;
  int match_end = 0;
  bool has_if_not_found = false;
  Value if_not_found = Value::error(ErrorCode::NA);
};

Expected<TextBeforeAfterOpts, ErrorCode> read_tba_opts(const Value* args, std::uint32_t arity) {
  TextBeforeAfterOpts opts;
  if (arity >= 3) {
    auto parsed = text_detail::read_int_arg(args[2]);
    if (!parsed) {
      return parsed.error();
    }
    opts.instance_num = parsed.value();
  }
  if (arity >= 4) {
    auto parsed = text_detail::read_int_arg(args[3]);
    if (!parsed) {
      return parsed.error();
    }
    opts.match_mode = parsed.value();
  }
  if (arity >= 5) {
    auto parsed = text_detail::read_int_arg(args[4]);
    if (!parsed) {
      return parsed.error();
    }
    opts.match_end = parsed.value();
  }
  if (arity >= 6) {
    opts.has_if_not_found = true;
    opts.if_not_found = args[5];
    if (opts.if_not_found.is_error()) {
      return opts.if_not_found.as_error();
    }
  }
  return opts;
}

}  // namespace

namespace text_detail {

// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
Value TextBefore_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto delimiter = coerce_to_text(args[1]);
  if (!delimiter) {
    return Value::error(delimiter.error());
  }
  auto opts = read_tba_opts(args, arity);
  if (!opts) {
    return Value::error(opts.error());
  }
  bool arg_err = false;
  const TextMatchResult hit = find_text_instance(text.value(), delimiter.value(), opts.value().instance_num,
                                                 opts.value().match_mode, opts.value().match_end, &arg_err);
  if (arg_err) {
    return Value::error(ErrorCode::Value);
  }
  if (!hit.found) {
    if (opts.value().has_if_not_found) {
      return opts.value().if_not_found;
    }
    return Value::error(ErrorCode::NA);
  }
  // Everything before the match's starting byte.
  return Value::text(arena.intern(text.value().substr(0, hit.match_start_byte)));
}

// TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
Value TextAfter_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto delimiter = coerce_to_text(args[1]);
  if (!delimiter) {
    return Value::error(delimiter.error());
  }
  auto opts = read_tba_opts(args, arity);
  if (!opts) {
    return Value::error(opts.error());
  }
  bool arg_err = false;
  const TextMatchResult hit = find_text_instance(text.value(), delimiter.value(), opts.value().instance_num,
                                                 opts.value().match_mode, opts.value().match_end, &arg_err);
  if (arg_err) {
    return Value::error(ErrorCode::Value);
  }
  if (!hit.found) {
    if (opts.value().has_if_not_found) {
      return opts.value().if_not_found;
    }
    return Value::error(ErrorCode::NA);
  }
  return Value::text(arena.intern(text.value().substr(hit.match_end_byte)));
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
