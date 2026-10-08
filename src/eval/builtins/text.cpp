//
// Implementation of Formulon's text built-in functions: UPPER, LOWER, TRIM,
// LEFT, RIGHT, MID, REPT, SUBSTITUTE, FIND, SEARCH, EXACT, TEXTJOIN,
// UNICHAR, UNICODE, CLEAN, PROPER. VALUE (and its sibling converters TEXT /
// NUMBERVALUE) live under `eval/builtins/text_format.{h,cpp}` alongside the
// format-string engine they share with TEXT.
//
// Every text builtin coerces its inputs via `coerce_to_text` /
// `coerce_to_number`. Errors among the inputs already short-circuit through
// the dispatcher's left-most-error rule before we get here. Length and
// position arithmetic uses Excel's UTF-16 unit semantics via
// `eval/text_ops.h`; the result text (when any) is interned into the
// caller's arena so the returned `Value::text` payload is readable.

#include "eval/builtins/text.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/builtins/registration_helpers.h"
#include "eval/builtins/text_detail.h"
#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/jis0208_table.h"
#include "excel_locale.h"
#include "sbcs_codepage.h"
#include "utils/arena.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

using text_detail::read_int_arg;
using text_detail::read_snapped_int_arg;
using text_detail::read_text_window_args;

// UPPER(text) / LOWER(text) - ASCII case fold. Multi-byte UTF-8 bytes are
// preserved verbatim (see `text_ops::to_upper_ascii` for the contract).
Value Upper(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(to_upper_ascii(text.value())));
}

Value Lower(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(to_lower_ascii(text.value())));
}

// TRIM(text) - removes leading and trailing ASCII spaces (0x20), plus the
// ideographic space U+3000 (encoded UTF-8 as `E3 80 80`) for DBCS profiles.
// It collapses runs of those same characters internally. Other
// whitespace-like bytes (tabs, newlines, NBSP U+00A0, etc.) are preserved
// verbatim.
//
// A collapsed run keeps the character that *started* it rather than
// normalising to an ASCII space: `"a　　b"` trims to `"a　b"`, and a mixed
// run takes its first member, so `" 　 "` collapses to `" "` while `"　 "`
// collapses to `"　"`. Emitting an ASCII space for every run reads as the
// obvious simplification and is what this did; it silently rewrote
// ja-JP text, since a full-width space between two words is typography
// rather than padding.
Value Trim(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  const bool dbcs = locale_facts(current_eval_profile()).dbcs;
  // Detects U+3000 (UTF-8: 0xE3 0x80 0x80) starting at byte index `i` in
  // src. The bound is written `i + 3 <= src.size()` rather than the
  // equivalent `i + 2 < src.size()` so the "we need to read three bytes
  // starting at `i`" intent is visible at a glance and so a future
  // refactor that changes the sequence length is harder to get wrong.
  auto is_ideographic_space_at = [&src](std::size_t i) -> bool {
    return i + 3 <= src.size() && static_cast<unsigned char>(src[i]) == 0xE3u &&
           static_cast<unsigned char>(src[i + 1]) == 0x80u && static_cast<unsigned char>(src[i + 2]) == 0x80u;
  };
  std::string out;
  out.reserve(src.size());
  // The pending run's first member, held as the bytes to emit if the run
  // turns out to be interior. Empty means no run is open, which is why a
  // separate `pending_space` flag is not needed.
  std::string_view pending_space;
  bool seen_non_space = false;
  for (std::size_t i = 0; i < src.size();) {
    if (src[i] == ' ') {
      if (seen_non_space && pending_space.empty()) {
        pending_space = " ";
      }
      ++i;
      continue;
    }
    if (dbcs && is_ideographic_space_at(i)) {
      if (seen_non_space && pending_space.empty()) {
        pending_space = "\xE3\x80\x80";
      }
      i += 3;
      continue;
    }
    if (!pending_space.empty()) {
      out.append(pending_space);
      pending_space = std::string_view();
    }
    out.push_back(src[i]);
    ++i;
    seen_non_space = true;
  }
  return Value::text(arena.intern(out));
}

// The `(text, [n])` arguments of LEFT / RIGHT; `n` defaults to 1, snaps to
// a near integer, and a negative raw `n` is `#VALUE!`.
struct TextAndCount {
  std::string text;
  int count;
};

Expected<TextAndCount, ErrorCode> read_text_and_count(const Value* args, std::uint32_t arity) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return std::move(text.error());
  }
  auto parsed = arity > 1u ? read_snapped_int_arg(args[1], 0.0) : Expected<int, ErrorCode>(1);
  if (!parsed) {
    return std::move(parsed.error());
  }
  return TextAndCount{std::string(text.value()), parsed.value()};
}

// LEFT(text, [n]) - first `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Left(const Value* args, std::uint32_t arity, Arena& arena) {
  auto in = read_text_and_count(args, arity);
  if (!in) {
    return Value::error(in.error());
  }
  const std::string& text = in.value().text;
  const int n = in.value().count;
  if (n == 0) {
    return Value::text({});
  }
  return Value::text(arena.intern(utf16_substring(text, 0u, static_cast<std::uint32_t>(n))));
}

// RIGHT(text, [n]) - last `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Right(const Value* args, std::uint32_t arity, Arena& arena) {
  auto in = read_text_and_count(args, arity);
  if (!in) {
    return Value::error(in.error());
  }
  const std::string& text = in.value().text;
  const int n = in.value().count;
  if (n == 0) {
    return Value::text({});
  }
  const std::uint32_t total = utf16_units_in(text);
  const auto take = static_cast<std::uint32_t>(n);
  const std::uint32_t start = take >= total ? 0u : total - take;
  return Value::text(arena.intern(utf16_substring(text, start, take)));
}

// MID(text, start_num, num_chars) - 1-based slice in UTF-16 units. Excel
// returns `""` when `start_num` is past the end. `start_num<1` or
// `num_chars<0` -> `#VALUE!`.
Value Mid(const Value* args, std::uint32_t arity, Arena& arena) {
  auto parsed = read_text_window_args(args, arity, /*snap_start=*/false);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const std::string& text = parsed.value().text;
  const int start = parsed.value().start;
  const int length = parsed.value().count;
  const std::uint32_t total = utf16_units_in(text);
  const auto start_unit = static_cast<std::uint32_t>(start - 1);
  if (start_unit >= total) {
    return Value::text({});
  }
  if (length == 0) {
    return Value::text({});
  }
  return Value::text(arena.intern(utf16_substring(text, start_unit, static_cast<std::uint32_t>(length))));
}

// REPT(text, n) - repeat. n<0 -> `#VALUE!`. Excel caps the result length at
// 32,767 UTF-16 units; exceeding the cap also surfaces as `#VALUE!`.
Value Rept(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto count = read_int_arg(args[1]);
  if (!count) {
    return Value::error(count.error());
  }
  if (count.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (count.value() == 0 || text.value().empty()) {
    return Value::text({});
  }
  const auto unit_len = static_cast<std::uint64_t>(utf16_units_in(text.value()));
  const auto reps = static_cast<std::uint64_t>(count.value());
  if (unit_len > 0 && reps > kExcelTextCapUnits / unit_len) {
    return Value::error(ErrorCode::Value);
  }
  std::string out;
  // The `unit_len * reps` overflow check above guarantees the product
  // fits within Excel's text-cap and therefore within `size_t` on every
  // supported target (including 32-bit wasm32). Cast explicitly to
  // silence `-Wshorten-64-to-32` under wasm32 where `size_t` is 32-bit.
  out.reserve(static_cast<std::size_t>(text.value().size() * reps));
  for (std::uint64_t i = 0; i < reps; ++i) {
    out.append(text.value());
  }
  return Value::text(arena.intern(out));
}

// SUBSTITUTE(text, old_text, new_text, [instance_num]) - case-sensitive,
// byte-exact replace. Without `instance_num`, every occurrence is replaced.
// With `instance_num`, only the Nth (1-based) occurrence. Empty `old_text`
// and `instance_num` greater than the number of occurrences both return
// `text` unchanged. `instance_num < 1` -> `#VALUE!`.
Value Substitute(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto old_text = coerce_to_text(args[1]);
  if (!old_text) {
    return Value::error(old_text.error());
  }
  auto new_text = coerce_to_text(args[2]);
  if (!new_text) {
    return Value::error(new_text.error());
  }
  bool nth_only = false;
  int instance = 0;
  if (arity >= 4) {
    auto parsed = read_snapped_int_arg(args[3], 1.0);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    nth_only = true;
    instance = parsed.value();
  }
  const std::string& haystack = text.value();
  const std::string& needle = old_text.value();
  if (needle.empty()) {
    return Value::text(arena.intern(haystack));
  }
  std::string out;
  out.reserve(haystack.size());
  std::uint64_t out_units = 0;
  const auto append_capped = [&out, &out_units](std::string_view piece) {
    const std::uint64_t units = utf16_units_in(piece);
    if (units > kExcelTextCapUnits - out_units) {
      return false;
    }
    out.append(piece);
    out_units += units;
    return true;
  };
  std::size_t i = 0;
  int hits = 0;
  while (i < haystack.size()) {
    const std::size_t pos = haystack.find(needle, i);
    if (pos == std::string::npos) {
      if (!append_capped(std::string_view(haystack).substr(i))) {
        return Value::error(ErrorCode::Value);
      }
      break;
    }
    if (!append_capped(std::string_view(haystack).substr(i, pos - i))) {
      return Value::error(ErrorCode::Value);
    }
    ++hits;
    if (!nth_only || hits == instance) {
      if (!append_capped(new_text.value())) {
        return Value::error(ErrorCode::Value);
      }
    } else {
      if (!append_capped(needle)) {
        return Value::error(ErrorCode::Value);
      }
    }
    i = pos + needle.size();
  }
  return Value::text(arena.intern(out));
}

// REPLACE(old_text, start_num, num_chars, new_text) - UTF-16-unit replace.
// Replaces `num_chars` UTF-16 units of `old_text` starting at 1-based
// `start_num` with `new_text`. `start_num > total_units + 1` clamps to
// append (the suffix is empty). `start_num < 1` or `num_chars < 0` surface
// `#VALUE!`. The result is capped at Excel's 32,767-unit text limit.
Value Replace_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto parsed = read_text_window_args(args, arity, /*snap_start=*/true);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const std::string& old_text = parsed.value().text;
  const int start = parsed.value().start;
  const int num_chars = parsed.value().count;
  const std::string& new_text = parsed.value().new_text;
  const std::uint32_t total = utf16_units_in(old_text);
  auto start_unit = static_cast<std::uint32_t>(start - 1);
  if (start_unit > total) {
    start_unit = total;
  }
  const std::uint32_t remaining = total - start_unit;
  const auto delete_units =
      static_cast<std::uint32_t>(num_chars) > remaining ? remaining : static_cast<std::uint32_t>(num_chars);
  const std::string prefix = utf16_substring(old_text, 0u, start_unit);
  const std::string suffix = utf16_substring(old_text, start_unit + delete_units, remaining - delete_units);
  std::string out;
  out.reserve(prefix.size() + new_text.size() + suffix.size());
  out.append(prefix);
  out.append(new_text);
  out.append(suffix);
  if (static_cast<std::uint64_t>(utf16_units_in(out)) > kExcelTextCapUnits) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

// FIND(find_text, within_text, [start_num]) - case-sensitive, no wildcards.
// 1-based UTF-16-unit position of the first occurrence at or after
// `start_num` (default 1). Not found -> `#VALUE!`. Out-of-range `start_num`
// -> `#VALUE!`. Empty `find_text` returns `start_num` (Excel quirk).
Value Find(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  text_detail::SearchArgs sargs;
  Value early = Value::blank();
  if (!text_detail::read_search_args(args, arity, text_detail::SearchUnit::Utf16, &sargs, &early)) {
    return early;
  }
  const int start = sargs.start;
  const std::size_t start_byte = utf16_to_byte_offset(sargs.haystack, static_cast<std::uint32_t>(start - 1));
  const std::size_t pos = sargs.haystack.find(sargs.needle, start_byte);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  // Convert byte offset back to a 1-based UTF-16 unit position.
  const std::uint32_t units = utf16_units_in(std::string_view(sargs.haystack).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

// SEARCH(find_text, within_text, [start_num]) - case-insensitive substring
// search with DOS-style wildcards in `find_text`: `?` matches any single
// UTF-16 unit, `*` matches zero or more units, and `~?` / `~*` match the
// literal metacharacter. Otherwise mirrors FIND (which is strictly literal
// and retains that contract).
Value Search(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  text_detail::SearchArgs sargs;
  Value early = Value::blank();
  if (!text_detail::read_search_args(args, arity, text_detail::SearchUnit::Utf16, &sargs, &early)) {
    return early;
  }
  const std::size_t start_byte = utf16_to_byte_offset(sargs.haystack, static_cast<std::uint32_t>(sargs.start - 1));
  const std::size_t pos =
      text_detail::find_folded(sargs.haystack, sargs.needle, start_byte, text_detail::SearchUnit::Utf16);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t units = utf16_units_in(std::string_view(sargs.haystack).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

// EXACT(text1, text2) - byte-wise (case-sensitive) equality.
Value Exact(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto a = coerce_to_text(args[0]);
  if (!a) {
    return Value::error(a.error());
  }
  auto b = coerce_to_text(args[1]);
  if (!b) {
    return Value::error(b.error());
  }
  return Value::boolean(a.value() == b.value());
}

// --- Text manipulation, second batch ------------------------------------
//
// UNICHAR, UNICODE, CLEAN, PROPER. The same conventions as the first text
// batch apply: argument coercion via the helpers in `eval/coerce.h`, error
// propagation through the dispatcher's left-most rule, results interned
// into the call's arena.
//
// TEXTJOIN moved to `eval/textjoin_lazy.{h,cpp}` and the lazy dispatch
// table: its `delimiter` / `ignore_empty` positions must not be flattened
// by the generic `accepts_ranges` dispatcher the way `text1, [text2], ...`
// is, which an eager FunctionDef registration cannot express.

// UNICHAR(number) - returns the Unicode character whose codepoint is
// `number`. Truncates the input to an integer. Out-of-range and surrogate
// codepoints surface `#VALUE!`. Result is encoded as UTF-8 bytes.
Value Unichar(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto parsed = read_int_arg(args[0]);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const int n = parsed.value();
  if (n < 1 || n > 0x10FFFF) {
    return Value::error(ErrorCode::Value);
  }
  if (n >= 0xD800 && n <= 0xDFFF) {
    // UTF-16 surrogate halves do not represent characters on their own.
    return Value::error(ErrorCode::Value);
  }
  const std::string encoded = encode_utf8_codepoint(static_cast<std::uint32_t>(n));
  if (encoded.empty()) {
    // Defensive: encoder validates internally; an empty string here would
    // mean the helper rejected our codepoint despite the range checks.
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(encoded));
}

// UNICODE(text) - returns the Unicode codepoint of the first character in
// `text`. Empty text yields `#VALUE!`. The returned value is the actual
// codepoint, not a UTF-16 code unit: supplementary-plane characters return
// values above 0xFFFF (e.g. `UNICODE("😀")` = 128512, not the high surrogate).
Value Unicode_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  if (text.value().empty()) {
    return Value::error(ErrorCode::Value);
  }
  const Utf8DecodeResult decoded = decode_first_utf8_codepoint(text.value());
  if (!decoded.valid) {
    return Value::error(ErrorCode::Value);
  }
  return Value::number(static_cast<double>(decoded.codepoint));
}

// CLEAN(text) - strips ASCII control characters (0x00..0x1F) from `text`.
// Bytes >= 0x20 (including 0x7F DEL and the entire UTF-8 multi-byte range
// 0x80..0xFF) are preserved verbatim. Embedded NUL is NOT a string
// terminator here: the input is a `string_view` and we copy through every
// non-control byte.
Value Clean(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  for (char c : src) {
    if (static_cast<unsigned char>(c) >= 0x20u) {
      out.push_back(c);
    }
  }
  return Value::text(arena.intern(out));
}

// --- CHAR / CODE ---------------------------------------------------------
//
// Excel ja-JP measures "byte" length against the Shift-JIS (CP932) single-
// byte region: ASCII and half-width katakana are 1 byte, and every other
// character -- including supplementary-plane codepoints such as emoji --
// counts as 2 bytes (matching observed Mac Excel ja-JP behaviour). See
// `src/eval/utf8_length.h::byte_count_jajp` for the classifier.
//
// The byte-oriented slicing family (LENB / LEFTB / RIGHTB / MIDB /
// REPLACEB / FINDB / SEARCHB) lives in `text_dbcs.cpp`.

// CHAR(number) - returns the single-character text whose active profile's
// single-byte codepage value is `number`. DBCS profiles retain the CP932/JIS
// extension below; single-byte profiles accept 1..255 only.
//
// Mac Excel ja-JP uses CP932 / JIS X 0208. The valid argument space is the
// disjoint union of two ranges:
//
//   1..255       SBCS region. ASCII (1..127), the CP1252 high table for
//                lead-byte slots (0x80..0x9F), half-width katakana
//                (0xA1..0xDF -> U+FF61..U+FF9F), and the 0xA0..0xFF
//                passthrough are kept exactly as before; ASC/JIS round-trip
//                relies on the half-width katakana mapping.
//   8481..32382  DBCS region: `n` is split into `(hi << 8) | lo` where
//                hi, lo each live in [0x21, 0x7E] (i.e. JIS X 0208 row,
//                cell in [1, 94]). The Unicode codepoint comes from the
//                JIS X 0208 reverse table.
//
// Anything else (256..8480, unmapped DBCS slots, bytes outside [0x21, 0x7E],
// n < 1, n > 65535) yields `#VALUE!`. Mac probe golden:
// tests/oracle/targets/mac-365-ja_JP/golden/code_char_jp_probes.golden.json.
Value Char_(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  const ExcelProfile profile = current_eval_profile();
  const LocaleFacts& facts = locale_facts(profile);
  auto parsed = facts.char_snaps_near_integer ? read_snapped_int_arg(args[0], 0.0) : read_int_arg(args[0]);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const int n = parsed.value();
  if (n < 1) {
    return Value::error(ErrorCode::Value);
  }
  if (!facts.dbcs) {
    if (n > 0xFF) {
      return Value::error(ErrorCode::Value);
    }
    const std::uint32_t cp = sbcs_decode_byte(sbcs_codepage(profile), static_cast<std::uint8_t>(n));
    const std::string encoded = encode_utf8_codepoint(cp);
    if (encoded.empty()) {
      return Value::error(ErrorCode::Value);
    }
    return Value::text(arena.intern(encoded));
  }
  std::uint32_t cp = 0;
  if (n <= 0xFF) {
    if (n < 0x80) {
      cp = static_cast<std::uint32_t>(n);
    } else if (n >= 0xA1 && n <= 0xDF) {
      // Half-width katakana mapping: 0xA1 -> U+FF61, 0xDF -> U+FF9F.
      cp = 0xFF61u + static_cast<std::uint32_t>(n - 0xA1);
    } else if ((n >= 0x81 && n <= 0x9F) || (n >= 0xE0 && n <= 0xFC)) {
      // A CP932 lead byte on its own is a space (measured on Mac Excel ja-JP).
      cp = 0x20u;
    } else {
      // 0x80, 0xA0 and 0xFD-0xFF map to the same-numbered code point.
      cp = static_cast<std::uint32_t>(n);
    }
  } else if (n < 0x2121 || n > 0xFFFF) {
    // Gap between SBCS and DBCS, or wider than two bytes. Mac returns
    // #VALUE! (probe `char_above_255` in code_char_jp_probes).
    return Value::error(ErrorCode::Value);
  } else {
    const auto hi = static_cast<std::uint8_t>((n >> 8) & 0xFF);
    const auto lo = static_cast<std::uint8_t>(n & 0xFF);
    const std::uint16_t mapped = lookup_jis0208_to_unicode(hi, lo);
    if (mapped == 0u) {
      return Value::error(ErrorCode::Value);
    }
    cp = static_cast<std::uint32_t>(mapped);
  }
  const std::string encoded = encode_utf8_codepoint(cp);
  if (encoded.empty()) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(encoded));
}

// CODE(text) - returns the active profile's codepage value for the first
// character in `text`. Empty text yields `#VALUE!`.
//
//   * ASCII (codepoint < 0x80): codepoint itself.
//   * Half-width katakana (U+FF61..U+FF9F): 0xA1..0xDF.
//   * Anything in JIS X 0208: the row-cell encoding from
//     `lookup_unicode_to_jis0208` (e.g. CODE("あ") = 0x2422 = 9250).
//   * Anything else (NEC extensions, emoji, supplementary plane, etc.):
//     95 (ASCII underscore), the empirically confirmed Mac fallback.
//
// Mac probe golden: tests/oracle/targets/mac-365-ja_JP/golden/code_char_jp_probes.golden.json.
//
// This is the Mac-profile behavior. Under the runtime-default win-365-ja_JP
// profile, `eval_code_lazy` (src/eval/info_lazy.cpp) overrides the fallback
// case before it ever reaches this eager impl: a supplementary-plane
// codepoint (> U+FFFF) returns 63, and U+9AD9 specifically returns 38526,
// instead of the 95 fallback above.
Value Code_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const ExcelProfile profile = current_eval_profile();
  const LocaleFacts& facts = locale_facts(profile);
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  if (text.value().empty()) {
    return Value::error(ErrorCode::Value);
  }
  const Utf8DecodeResult decoded = decode_first_utf8_codepoint(text.value());
  if (!decoded.valid) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t cp = decoded.codepoint;
  if (!facts.dbcs) {
    const int encoded = sbcs_encode_codepoint(sbcs_codepage(profile), cp);
    return Value::number(static_cast<double>(encoded < 0 ? 95 : encoded));
  }
  if (cp < 0x80u) {
    return Value::number(static_cast<double>(cp));
  }
  if (cp >= 0xFF61u && cp <= 0xFF9Fu) {
    return Value::number(static_cast<double>(0xA1u + (cp - 0xFF61u)));
  }
  const std::uint16_t mapped = lookup_unicode_to_jis0208(cp);
  if (mapped != 0u) {
    return Value::number(static_cast<double>(mapped));
  }
  // Mac fallback for anything outside JIS X 0208 (NEC ext, emoji, etc.).
  return Value::number(95.0);
}

// PROPER(text) - title-case `text`. ASCII letters that begin a "word" are
// uppercased; ASCII letters that follow another ASCII letter are lowercased.
// A "word boundary" is any byte that is NOT an ASCII letter, including
// digits, punctuation, whitespace, and any byte >= 0x80 (so a Japanese
// character followed by an ASCII letter starts a new word). Non-ASCII bytes
// pass through unchanged - matching the existing UPPER / LOWER policy.
Value Proper(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  bool start_of_word = true;
  for (char c : src) {
    const auto u = static_cast<unsigned char>(c);
    const bool is_lower = (u >= 'a' && u <= 'z');
    const bool is_upper = (u >= 'A' && u <= 'Z');
    if (is_lower || is_upper) {
      if (start_of_word) {
        out.push_back(is_upper ? c : static_cast<char>(c - 32));
      } else {
        out.push_back(is_lower ? c : static_cast<char>(c + 32));
      }
      start_of_word = false;
    } else {
      out.push_back(c);
      start_of_word = true;
    }
  }
  return Value::text(arena.intern(out));
}

// HYPERLINK(link_location, [friendly_name]) - Excel's hyperlink function
// returns `friendly_name` as-is (preserving its Value variant) when the
// 2-arg form is used; otherwise it returns `link_location` coerced to
// text (the link argument is documented as a URL, so the cell value is a
// string regardless of whether the input was a number). Formulon has no
// concept of clickable cells, so the function exists purely to mirror
// Excel's pure-value behaviour. Error arguments propagate through the
// dispatcher's default short-circuit. Mac Excel preserves the original
// type for friendly_name: `=HYPERLINK("x", 42)` yields the number 42,
// not the text "42"; booleans stay booleans; blanks render as 0 (Excel's
// usual blank-as-zero rule for non-text-context outputs).
Value Hyperlink(const Value* args, std::uint32_t arity, Arena& arena) {
  if (arity >= 2) {
    const Value& friendly = args[1];
    // Blank friendly_name renders as 0 in Mac Excel (the cell shows 0
    // rather than empty) - mirror by coercing to a Number.
    if (friendly.kind() == ValueKind::Blank) {
      return Value::number(0.0);
    }
    // Number, Bool, Text, Error: pass through unchanged.
    return friendly;
  }
  auto link = coerce_to_text(args[0]);
  if (!link) {
    return Value::error(link.error());
  }
  return Value::text(arena.intern(link.value()));
}

}  // namespace

void register_text_builtins(FunctionRegistry& registry) {
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"UPPER", 1u, 1u, &Upper},
      {"LOWER", 1u, 1u, &Lower},
      {"TRIM", 1u, 1u, &Trim},
      {"LEFT", 1u, 2u, &Left},
      {"RIGHT", 1u, 2u, &Right},
      {"MID", 3u, 3u, &Mid},
      {"REPT", 2u, 2u, &Rept},
      {"SUBSTITUTE", 3u, 4u, &Substitute},
      {"FIND", 2u, 3u, &Find},
      {"SEARCH", 2u, 3u, &Search},
      {"EXACT", 2u, 2u, &Exact},
      {"REPLACE", 4u, 4u, &Replace_},
      {"FINDB", 2u, 3u, &text_detail::FindB_},
      {"SEARCHB", 2u, 3u, &text_detail::SearchB_},
      {"TEXTBEFORE", 2u, 6u, &text_detail::TextBefore_, false},
      {"TEXTAFTER", 2u, 6u, &text_detail::TextAfter_, false},
      // TEXTJOIN is lazy (see eval/textjoin_lazy.h) so its delimiter /
      // ignore_empty positions can be told apart from the flattened
      // text1, [text2], ... positions; no FunctionDef entry here.
      {"UNICHAR", 1u, 1u, &Unichar},
      {"UNICODE", 1u, 1u, &Unicode_},
      {"CLEAN", 1u, 1u, &Clean},
      {"PROPER", 1u, 1u, &Proper},
      {"LENB", 1u, 1u, &text_detail::Lenb},
      {"LEFTB", 1u, 2u, &text_detail::Leftb},
      {"RIGHTB", 1u, 2u, &text_detail::Rightb},
      {"MIDB", 3u, 3u, &text_detail::Midb},
      {"REPLACEB", 4u, 4u, &text_detail::ReplaceB_},
      {"CHAR", 1u, 1u, &Char_},
      {"CODE", 1u, 1u, &Code_},
      {"HYPERLINK", 1u, 2u, &Hyperlink},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
