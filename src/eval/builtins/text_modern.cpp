//
// Modern text accessor family: TEXTBEFORE / TEXTAFTER. Extracted from the
// core text TU to keep the catalog's translation units roughly balanced.
// The shared match-locator (`find_text_instance`) and the trailing-argument
// reader (`read_tba_opts`) live in this file's anonymous namespace because
// they are referenced only by TEXTBEFORE and TEXTAFTER.

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/builtins/text_detail.h"
#include "eval/coerce.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
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
//   * Empty `delimiter` matches at the start of `text` for any positive
//     `instance_num` and at the end for any negative one, so it never falls
//     through to not-found; `match_end` has no effect on it.
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
  if (instance_num == 0 || (match_mode != 0 && match_mode != 1) || (match_end != 0 && match_end != 1)) {
    *out_err_value = true;
    return {};
  }
  *out_err_value = false;
  if (delimiter.empty()) {
    const std::size_t at = instance_num > 0 ? 0u : text.size();
    return TextMatchResult{true, at, at};
  }

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

  // Virtual sentinels are directional: positive instances add the end of
  // text, while negative instances add the start. Keeping that distinction
  // here lets the ordinary forward scan retain Excel's greedy, non-overlap
  // matching without materialising every hit.
  const bool use_virtual_end = match_end == 1 && instance_num > 0;
  const bool use_virtual_start = match_end == 1 && instance_num < 0;
  const auto hit_at = [&needle](std::size_t pos) { return TextMatchResult{true, pos, pos + needle.size()}; };

  if (instance_num > 0) {
    const std::size_t wanted = static_cast<std::size_t>(instance_num);
    std::size_t ordinal = 0;
    std::size_t scan = 0;
    while (scan <= hay.size()) {
      const std::size_t pos = hay.find(needle, scan);
      if (pos == std::string_view::npos) {
        break;
      }
      ++ordinal;
      if (ordinal == wanted) {
        return hit_at(pos);
      }
      // Advance past the match to preserve the existing greedy non-overlap
      // rule (for example, "aa" in "aaaa" has hits at 0 and 2).
      scan = pos + needle.size();
    }
    if (use_virtual_end && ordinal + 1u == wanted) {
      return TextMatchResult{true, text.size(), text.size()};
    }
    return {};
  }

  // Convert through a wider signed type before negating. Unary minus on
  // INT_MIN is undefined; the snapped numeric argument can still produce it.
  const std::size_t wanted_from_end = static_cast<std::size_t>(-(static_cast<std::int64_t>(instance_num)));
  if (wanted_from_end == 1u) {
    TextMatchResult last;
    std::size_t scan = 0;
    while (scan <= hay.size()) {
      const std::size_t pos = hay.find(needle, scan);
      if (pos == std::string_view::npos) {
        break;
      }
      last = hit_at(pos);
      scan = last.match_end_byte;
    }
    return last.found ? last : (use_virtual_start ? TextMatchResult{true, 0u, 0u} : TextMatchResult{});
  }

  // Keep only start offsets for general negative instances. A second search
  // would repeat nearly the entire scan for -2; storing offsets preserves the
  // forward non-overlapping sequence without copying a second hit array.
  std::vector<std::size_t> starts;
  std::size_t scan = 0;
  while (scan <= hay.size()) {
    const std::size_t pos = hay.find(needle, scan);
    if (pos == std::string_view::npos) {
      break;
    }
    starts.push_back(pos);
    scan = pos + needle.size();
  }
  if (wanted_from_end <= starts.size()) {
    return hit_at(starts[starts.size() - wanted_from_end]);
  }
  if (use_virtual_start && wanted_from_end - 1u == starts.size()) {
    return TextMatchResult{true, 0u, 0u};
  }
  return {};
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
    // instance_num snaps to a near integer, then uses floor-toward-negative-
    // infinity (Excel INT semantics), not truncation: -1.6 -> -2 selects the
    // second-to-last occurrence.
    auto coerced = coerce_to_number(args[2]);
    if (!coerced) {
      return std::move(coerced.error());
    }
    const double d = coerced.value();
    if (std::isnan(d) || std::isinf(d)) {
      return ErrorCode::Num;
    }
    // Saturated: find_text_instance treats an out-of-range instance as not found.
    opts.instance_num = truncate_saturated_int(std::floor(snap_near_integer(d)));
  }
  auto match_mode = text_detail::read_optional_int_arg(args, arity, 3u, 0);
  if (!match_mode) {
    return std::move(match_mode.error());
  }
  opts.match_mode = match_mode.value();
  auto match_end = text_detail::read_optional_int_arg(args, arity, 4u, 0);
  if (!match_end) {
    return std::move(match_end.error());
  }
  opts.match_end = match_end.value();
  if (arity >= 6) {
    // if_not_found is a passthrough fallback value: an error value here is
    // only surfaced when the match actually fails. With propagate_errors=false
    // on the registration, the dispatcher delivers it as a raw Value.
    opts.has_if_not_found = true;
    opts.if_not_found = args[5];
  }
  return opts;
}

// Shared frame of TEXTBEFORE / TEXTAFTER. A boolean keeps the two public
// entry points on one non-templated implementation, avoiding duplicate code
// generation for the otherwise identical argument and matching paths.
Value text_before_after(const Value* args, std::uint32_t arity, Arena& arena, bool after) {
  // Registered with propagate_errors=false so if_not_found (args[5]) can pass
  // through as an error fallback. Propagate errors manually for args 0-4.
  for (std::uint32_t i = 0; i < arity && i < 5; ++i) {
    if (args[i].is_error()) {
      return args[i];
    }
  }
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
  // Domain error: delimiter longer than text is a #VALUE! — but ONLY when
  // the caller passed any optional argument. The 2-arg minimum form
  // (`TEXTBEFORE(text, delim)`) treats an impossible delimiter as a
  // regular "not found" and returns #N/A. Observed in Excel 365 /
  // IronCalc; K19 / J19 in TEXTBEFORE_TEXTAFTER_Sheet1 pin this split.
  //
  // The comparison is a length test, not a shape test, so it uses Excel's
  // own length unit. A raw byte comparison makes any non-ASCII delimiter
  // look impossible against a shorter-in-bytes ASCII text — `TEXTBEFORE(
  // "ab", "あ", 1)` would be `#VALUE!` instead of the not-found `#N/A`
  // that `if_not_found` / `IFNA` are meant to catch.
  if (arity >= 3 && utf16_units_in(delimiter.value()) > utf16_units_in(text.value())) {
    return Value::error(ErrorCode::Value);
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
  const std::string_view text_view = text.value();
  const std::string_view result =
      after ? text_view.substr(hit.match_end_byte) : text_view.substr(0, hit.match_start_byte);
  return Value::text(arena.intern(result));
}

}  // namespace

namespace text_detail {

// TEXTBEFORE(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
Value TextBefore_(const Value* args, std::uint32_t arity, Arena& arena) {
  return text_before_after(args, arity, arena, false);
}

// TEXTAFTER(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])
Value TextAfter_(const Value* args, std::uint32_t arity, Arena& arena) {
  return text_before_after(args, arity, arena, true);
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
