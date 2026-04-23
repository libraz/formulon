// Copyright 2026 libraz. Licensed under the MIT License.
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

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/function_registry.h"
#include "eval/text_ops.h"
#include "eval/utf8_length.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Excel caps the result of REPT (and a handful of related text functions)
// at 32,767 UTF-16 units. We reuse the same constant in REPT's overflow
// guard.
constexpr std::uint64_t kExcelTextCapUnits = 32767u;

// Helper: read a numeric arg as an `int` via `std::trunc`. Returns `#VALUE!`
// on coercion failure, `#NUM!` on non-finite input. Used by LEFT/RIGHT/MID/
// REPT/FIND/SEARCH/SUBSTITUTE for their integer-typed parameters.
Expected<int, ErrorCode> read_int_arg(const Value& v) {
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double d = coerced.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d));
}

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

// TRIM(text) - removes leading and trailing ASCII spaces (0x20) and
// collapses runs of internal ASCII spaces to a single space. Other
// whitespace-like bytes (tabs, newlines, Unicode whitespace) are preserved
// verbatim - this matches Excel exactly.
Value Trim(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string& src = text.value();
  std::string out;
  out.reserve(src.size());
  bool pending_space = false;
  bool seen_non_space = false;
  for (char c : src) {
    if (c == ' ') {
      if (seen_non_space) {
        pending_space = true;
      }
      continue;
    }
    if (pending_space) {
      out.push_back(' ');
      pending_space = false;
    }
    out.push_back(c);
    seen_non_space = true;
  }
  return Value::text(arena.intern(out));
}

// LEFT(text, [n]) - first `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Left(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  return Value::text(arena.intern(utf16_substring(text.value(), 0u, static_cast<std::uint32_t>(n))));
}

// RIGHT(text, [n]) - last `n` UTF-16 units. Default n=1. n<0 -> `#VALUE!`.
Value Right(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  const std::uint32_t total = utf16_units_in(text.value());
  const auto take = static_cast<std::uint32_t>(n);
  const std::uint32_t start = take >= total ? 0u : total - take;
  return Value::text(arena.intern(utf16_substring(text.value(), start, take)));
}

// MID(text, start_num, num_chars) - 1-based slice in UTF-16 units. Excel
// returns `""` when `start_num` is past the end. `start_num<1` or
// `num_chars<0` -> `#VALUE!`.
Value Mid(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto start = read_int_arg(args[1]);
  if (!start) {
    return Value::error(start.error());
  }
  auto length = read_int_arg(args[2]);
  if (!length) {
    return Value::error(length.error());
  }
  if (start.value() < 1 || length.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t total = utf16_units_in(text.value());
  const auto start_unit = static_cast<std::uint32_t>(start.value() - 1);
  if (start_unit >= total) {
    return Value::text({});
  }
  if (length.value() == 0) {
    return Value::text({});
  }
  return Value::text(
      arena.intern(utf16_substring(text.value(), start_unit, static_cast<std::uint32_t>(length.value()))));
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
  out.reserve(text.value().size() * reps);
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
    auto parsed = read_int_arg(args[3]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    if (parsed.value() < 1) {
      return Value::error(ErrorCode::Value);
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
  std::size_t i = 0;
  int hits = 0;
  while (i < haystack.size()) {
    const std::size_t pos = haystack.find(needle, i);
    if (pos == std::string::npos) {
      out.append(haystack, i, std::string::npos);
      break;
    }
    out.append(haystack, i, pos - i);
    ++hits;
    if (!nth_only || hits == instance) {
      out.append(new_text.value());
    } else {
      out.append(needle);
    }
    i = pos + needle.size();
  }
  return Value::text(arena.intern(out));
}

// FIND(find_text, within_text, [start_num]) - case-sensitive, no wildcards.
// 1-based UTF-16-unit position of the first occurrence at or after
// `start_num` (default 1). Not found -> `#VALUE!`. Out-of-range `start_num`
// -> `#VALUE!`. Empty `find_text` returns `start_num` (Excel quirk).
Value Find(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    return Value::error(needle.error());
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    return Value::error(haystack.error());
  }
  int start = 1;
  if (arity >= 3) {
    auto parsed = read_int_arg(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    start = parsed.value();
  }
  const std::uint32_t total = utf16_units_in(haystack.value());
  if (start < 1 || static_cast<std::uint32_t>(start) > total + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::size_t start_byte = utf16_to_byte_offset(haystack.value(), static_cast<std::uint32_t>(start - 1));
  const std::size_t pos = haystack.value().find(needle.value(), start_byte);
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  // Convert byte offset back to a 1-based UTF-16 unit position.
  const std::uint32_t units = utf16_units_in(std::string_view(haystack.value()).substr(0, pos));
  return Value::number(static_cast<double>(units + 1));
}

// SEARCH(find_text, within_text, [start_num]) - case-insensitive substring
// search with DOS-style wildcards in `find_text`: `?` matches any single
// UTF-16 unit, `*` matches zero or more units, and `~?` / `~*` match the
// literal metacharacter. Otherwise mirrors FIND (which is strictly literal
// and retains that contract).
Value Search(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto needle = coerce_to_text(args[0]);
  if (!needle) {
    return Value::error(needle.error());
  }
  auto haystack = coerce_to_text(args[1]);
  if (!haystack) {
    return Value::error(haystack.error());
  }
  int start = 1;
  if (arity >= 3) {
    auto parsed = read_int_arg(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    start = parsed.value();
  }
  const std::uint32_t total = utf16_units_in(haystack.value());
  if (start < 1 || static_cast<std::uint32_t>(start) > total + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::string lowered_haystack = to_lower_ascii(haystack.value());
  const std::string lowered_needle = to_lower_ascii(needle.value());
  const std::size_t start_byte = utf16_to_byte_offset(haystack.value(), static_cast<std::uint32_t>(start - 1));
  // Fast path: when the pattern carries no metacharacter at all (no `*`,
  // `?`, or `~`), use a plain substring search on the lower-cased buffers.
  // This keeps the common wildcard-free case allocation-free beyond the
  // two `to_lower_ascii` copies that case-insensitive matching already
  // requires. A bare `~` still needs the wildcard path because `~?` / `~*`
  // must be un-escaped when comparing.
  std::size_t pos = std::string::npos;
  const bool has_metachar = lowered_needle.find_first_of("*?~") != std::string::npos;
  if (!has_metachar) {
    pos = lowered_haystack.find(lowered_needle, start_byte);
  } else {
    // Wildcard path: scan the haystack suffix starting at `start_byte`. The
    // byte offset returned by `wildcard_find` is relative to that suffix, so
    // we add `start_byte` back to obtain an absolute byte position.
    const std::string_view suffix = std::string_view(lowered_haystack).substr(start_byte);
    const std::size_t rel = wildcard_find(lowered_needle, suffix);
    if (rel != std::string_view::npos) {
      pos = start_byte + rel;
    }
  }
  if (pos == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  const std::uint32_t units = utf16_units_in(std::string_view(haystack.value()).substr(0, pos));
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
// TEXTJOIN, UNICHAR, UNICODE, CLEAN, PROPER. The same conventions as the
// first text batch apply: argument coercion via the helpers in
// `eval/coerce.h`, error propagation through the dispatcher's left-most
// rule, results interned into the call's arena.

// TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)
//
// Joins every text argument with `delimiter`. When `ignore_empty` is TRUE
// (the typical case), arguments whose text representation is the empty
// string are skipped, so two consecutive empty inputs do NOT produce a
// double delimiter. With `ignore_empty` FALSE every argument participates
// even if empty (yielding consecutive delimiters). Result length is capped
// at Excel's 32,767-unit limit; exceeding it surfaces `#VALUE!`.
Value TextJoin(const Value* args, std::uint32_t arity, Arena& arena) {
  auto delimiter = coerce_to_text(args[0]);
  if (!delimiter) {
    return Value::error(delimiter.error());
  }
  auto ignore_empty = coerce_to_bool(args[1]);
  if (!ignore_empty) {
    return Value::error(ignore_empty.error());
  }
  std::string out;
  bool first = true;
  for (std::uint32_t i = 2; i < arity; ++i) {
    auto piece = coerce_to_text(args[i]);
    if (!piece) {
      return Value::error(piece.error());
    }
    if (ignore_empty.value() && piece.value().empty()) {
      continue;
    }
    if (!first) {
      out.append(delimiter.value());
    }
    out.append(piece.value());
    first = false;
    // Early-out cap check: once the byte length exceeds the byte upper bound
    // for the cap (4 bytes per UTF-16 unit pessimistically), the UTF-16 unit
    // count must also exceed the cap. Definitive check below.
    if (utf16_units_in(out) > kExcelTextCapUnits) {
      return Value::error(ErrorCode::Value);
    }
  }
  return Value::text(arena.intern(out));
}

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

// --- Byte-oriented text family (LENB / LEFTB / RIGHTB / MIDB / CHAR / CODE)
//
// Excel ja-JP measures "byte" length against the Shift-JIS (CP932) single-
// byte region: ASCII and half-width katakana are 1 byte, every other BMP
// character is 2 bytes, and supplementary-plane characters are 4 bytes
// (two UTF-16 surrogate halves, each classified as 2-byte DBCS). See
// `src/eval/utf8_length.h::byte_count_jajp` for the classifier.
//
// When a byte-budgeted slice (LEFTB/RIGHTB/MIDB) would land in the middle
// of a 2-byte character (i.e. we have exactly 1 byte of budget left and the
// next character costs 2), Excel substitutes a single ASCII space (0x20)
// and stops. MIDB additionally substitutes a space when its `start_byte`
// falls inside a 2-byte character.

// Decode a single UTF-8 codepoint starting at `src[i]`. On malformed input
// returns a 1-byte "raw" step so the caller can still make progress; such
// steps are classified as 1 DBCS byte per `byte_count_jajp` convention.
struct Utf8Step {
  std::uint32_t codepoint;  // Decoded codepoint (0xFFFD on malformed input).
  std::size_t byte_len;     // Number of UTF-8 bytes consumed (>= 1).
  int dbcs_bytes;           // DBCS byte cost under the ja-JP rule.
};

Utf8Step next_utf8_step(std::string_view src, std::size_t i) noexcept {
  const auto c0 = static_cast<unsigned char>(src[i]);
  if (c0 < 0x80u) {
    return {static_cast<std::uint32_t>(c0), 1u, 1};
  }
  std::size_t need = 0;
  std::uint32_t value = 0;
  if ((c0 & 0xE0u) == 0xC0u) {
    need = 1;
    value = c0 & 0x1Fu;
  } else if ((c0 & 0xF0u) == 0xE0u) {
    need = 2;
    value = c0 & 0x0Fu;
  } else if ((c0 & 0xF8u) == 0xF0u) {
    need = 3;
    value = c0 & 0x07u;
  } else {
    return {0xFFFDu, 1u, 1};
  }
  if (i + need >= src.size()) {
    return {0xFFFDu, 1u, 1};
  }
  for (std::size_t k = 0; k < need; ++k) {
    const auto ck = static_cast<unsigned char>(src[i + 1 + k]);
    if ((ck & 0xC0u) != 0x80u) {
      return {0xFFFDu, 1u, 1};
    }
    value = (value << 6) | (ck & 0x3Fu);
  }
  return {value, need + 1, byte_count_jajp(value)};
}

// LENB(text) - returns the ja-JP DBCS byte length of the coerced text.
Value Lenb(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::number(static_cast<double>(bytes_in_jajp(text.value())));
}

// LEFTB(text, [num_bytes=1]) - leftmost `num_bytes` bytes. When a 2-byte
// character would straddle the budget (1 byte remaining), emit an ASCII
// space in its place and stop.
Value Leftb(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  const std::string& src = text.value();
  const auto budget = static_cast<std::uint64_t>(n);
  std::uint64_t used = 0;
  std::string out;
  out.reserve(src.size());
  std::size_t i = 0;
  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i);
    const auto cost = static_cast<std::uint64_t>(step.dbcs_bytes);
    if (used + cost <= budget) {
      out.append(src, i, step.byte_len);
      used += cost;
      i += step.byte_len;
      if (used == budget) {
        break;
      }
    } else {
      // Character would overflow the budget. If we have exactly 1 byte
      // remaining and the character is a multi-byte (cost >= 2) char, pad
      // with a single ASCII space. Otherwise stop cleanly.
      if (used + 1 == budget && cost >= 2) {
        out.push_back(' ');
      }
      break;
    }
  }
  return Value::text(arena.intern(out));
}

// RIGHTB(text, [num_bytes=1]) - mirror of LEFTB. Walk the byte stream once
// forward to record the per-character byte cost and byte offset, then take
// the rightmost window that fits the budget. If a 2-byte character would
// straddle the window boundary on the left edge, emit an ASCII space.
Value Rightb(const Value* args, std::uint32_t arity, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  int n = 1;
  if (arity >= 2) {
    auto parsed = read_int_arg(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    n = parsed.value();
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (n == 0) {
    return Value::text({});
  }
  const std::string& src = text.value();
  const auto budget = static_cast<std::uint64_t>(n);

  // Collect (byte_offset, byte_len, dbcs_cost) for each character.
  struct CharRec {
    std::size_t offset;
    std::size_t byte_len;
    int dbcs_bytes;
  };
  std::vector<CharRec> chars;
  chars.reserve(src.size());
  {
    std::size_t i = 0;
    while (i < src.size()) {
      const Utf8Step step = next_utf8_step(src, i);
      chars.push_back({i, step.byte_len, step.dbcs_bytes});
      i += step.byte_len;
    }
  }

  // Walk right-to-left accumulating cost until we can't fit the next char.
  std::uint64_t used = 0;
  std::size_t first_full_idx = chars.size();
  bool pad_space = false;
  for (std::size_t k = chars.size(); k > 0; --k) {
    const CharRec& rec = chars[k - 1];
    const auto cost = static_cast<std::uint64_t>(rec.dbcs_bytes);
    if (used + cost <= budget) {
      used += cost;
      first_full_idx = k - 1;
      if (used == budget) {
        break;
      }
    } else {
      if (used + 1 == budget && cost >= 2) {
        pad_space = true;
      }
      break;
    }
  }

  std::string out;
  out.reserve(src.size());
  if (pad_space) {
    out.push_back(' ');
  }
  if (first_full_idx < chars.size()) {
    const std::size_t start_byte = chars[first_full_idx].offset;
    out.append(src, start_byte, std::string::npos);
  }
  return Value::text(arena.intern(out));
}

// MIDB(text, start_byte, num_bytes) - byte-window slice.
// `start_byte<1` or `num_bytes<0` -> `#VALUE!`. If `start_byte > LENB(text)`,
// return "". If `start_byte` lands inside a 2-byte character, emit a space
// pad and consume 1 byte of budget; from there walk normally, padding on
// the trailing edge with the same 1-byte-overflow rule as LEFTB.
Value Midb(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  auto start = read_int_arg(args[1]);
  if (!start) {
    return Value::error(start.error());
  }
  auto length = read_int_arg(args[2]);
  if (!length) {
    return Value::error(length.error());
  }
  if (start.value() < 1 || length.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  if (length.value() == 0) {
    return Value::text({});
  }
  const std::string& src = text.value();
  const auto start_byte = static_cast<std::uint64_t>(start.value());  // 1-based
  const auto budget = static_cast<std::uint64_t>(length.value());

  // Walk characters and track byte cursor.
  std::uint64_t cursor = 0;  // Bytes consumed so far (DBCS byte index).
  std::size_t i = 0;
  std::string out;
  out.reserve(src.size());
  std::uint64_t used = 0;
  bool started = false;
  bool head_pad = false;

  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i);
    const auto cost = static_cast<std::uint64_t>(step.dbcs_bytes);
    const std::uint64_t char_start = cursor + 1;   // 1-based byte position of this char.
    const std::uint64_t char_end = cursor + cost;  // 1-based byte position of final byte.
    if (!started) {
      if (char_end < start_byte) {
        // Entire character is before the window.
        cursor += cost;
        i += step.byte_len;
        continue;
      }
      // This character either starts at `start_byte` or straddles it.
      if (char_start == start_byte) {
        started = true;
        // Fall through to the "consume within budget" path below.
      } else {
        // Start falls strictly inside a multi-byte character. Emit a head
        // space pad consuming 1 byte of budget, skip the remainder of this
        // character, and continue.
        head_pad = true;
        out.push_back(' ');
        used += 1;
        cursor += cost;
        i += step.byte_len;
        started = true;
        if (used == budget) {
          return Value::text(arena.intern(out));
        }
        continue;
      }
    }
    // Inside the window.
    if (used + cost <= budget) {
      out.append(src, i, step.byte_len);
      used += cost;
      cursor += cost;
      i += step.byte_len;
      if (used == budget) {
        break;
      }
    } else {
      if (used + 1 == budget && cost >= 2) {
        out.push_back(' ');
      }
      break;
    }
  }
  // If the start was past the end of the text we emit nothing.
  if (!started && !head_pad) {
    return Value::text({});
  }
  return Value::text(arena.intern(out));
}

// CHAR(number) - returns the single-character text whose Mac Excel ja-JP
// codepage value is `number`. `number` is truncated to an integer; values
// outside 1..255 surface `#VALUE!`.
//
// Mac Excel ja-JP uses the Shift-JIS (CP932) codepage for this function.
// CP932 partitions 0x00..0xFF as follows:
//
//   0x00..0x7F  - ASCII (identical to Unicode)
//   0x81..0x9F  - DBCS lead byte (not a valid single-byte char)
//   0xA1..0xDF  - Half-width katakana (U+FF61..U+FF9F)
//   0xE0..0xFC  - DBCS lead byte
//   0x80, 0xA0, 0xFD..0xFF - unassigned
//
// We implement the two single-byte regions (ASCII and half-width katakana)
// exactly. Lead-byte values are echoed back as their CP1252 equivalent:
// Mac Excel returns a variety of fallback glyphs for these and we capture
// specific mismatches in `tests/divergence.yaml`.
Value Char_(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto parsed = read_int_arg(args[0]);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const int n = parsed.value();
  if (n < 1 || n > 255) {
    return Value::error(ErrorCode::Value);
  }
  std::uint32_t cp = 0;
  if (n < 0x80) {
    cp = static_cast<std::uint32_t>(n);
  } else if (n >= 0xA1 && n <= 0xDF) {
    // Half-width katakana mapping: 0xA1 -> U+FF61, 0xDF -> U+FF9F.
    cp = 0xFF61u + static_cast<std::uint32_t>(n - 0xA1);
  } else {
    // CP932 DBCS lead-byte or unassigned slot. Fall back to the CP1252
    // value so the character is at least printable; any divergence from
    // Mac Excel ja-JP is captured per-case in tests/divergence.yaml.
    static constexpr std::uint32_t kCp1252HighTable[32] = {
        0x20ACu, 0xFFFDu, 0x201Au, 0x0192u, 0x201Eu, 0x2026u, 0x2020u, 0x2021u, 0x02C6u, 0x2030u, 0x0160u,
        0x2039u, 0x0152u, 0xFFFDu, 0x017Du, 0xFFFDu, 0xFFFDu, 0x2018u, 0x2019u, 0x201Cu, 0x201Du, 0x2022u,
        0x2013u, 0x2014u, 0x02DCu, 0x2122u, 0x0161u, 0x203Au, 0x0153u, 0xFFFDu, 0x017Eu, 0x0178u,
    };
    if (n < 0xA0) {
      cp = kCp1252HighTable[n - 0x80];
    } else {
      cp = static_cast<std::uint32_t>(n);
    }
  }
  const std::string encoded = encode_utf8_codepoint(cp);
  if (encoded.empty()) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(encoded));
}

// CODE(text) - returns the ANSI/Mac codepage value of the first character
// in `text`. Empty text yields `#VALUE!`. For ASCII (<= 0x7F) the result is
// the codepoint; for non-ASCII we return the first UTF-8 codepoint as a
// number. Mac Excel ja-JP returns CP932 values for non-ASCII; any mismatch
// is documented in `tests/divergence.yaml`.
Value Code_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
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

}  // namespace

void register_text_builtins(FunctionRegistry& registry) {
  // Text manipulation.
  registry.register_function(FunctionDef{"UPPER", 1u, 1u, &Upper});
  registry.register_function(FunctionDef{"LOWER", 1u, 1u, &Lower});
  registry.register_function(FunctionDef{"TRIM", 1u, 1u, &Trim});
  registry.register_function(FunctionDef{"LEFT", 1u, 2u, &Left});
  registry.register_function(FunctionDef{"RIGHT", 1u, 2u, &Right});
  registry.register_function(FunctionDef{"MID", 3u, 3u, &Mid});
  registry.register_function(FunctionDef{"REPT", 2u, 2u, &Rept});
  registry.register_function(FunctionDef{"SUBSTITUTE", 3u, 4u, &Substitute});
  registry.register_function(FunctionDef{"FIND", 2u, 3u, &Find});
  registry.register_function(FunctionDef{"SEARCH", 2u, 3u, &Search});
  registry.register_function(FunctionDef{"EXACT", 2u, 2u, &Exact});

  // Text manipulation, second batch.
  {
    FunctionDef def{"TEXTJOIN", 3u, kVariadic, &TextJoin};
    // TEXTJOIN walks every cell in a range argument (row-major) and
    // concatenates their text projections. `accepts_ranges` asks the
    // dispatcher to flatten `Ref:Ref` arguments; blank cells coerce to "",
    // which the `ignore_empty=TRUE` branch skips, matching Excel.
    def.accepts_ranges = true;
    registry.register_function(def);
  }
  registry.register_function(FunctionDef{"UNICHAR", 1u, 1u, &Unichar});
  registry.register_function(FunctionDef{"UNICODE", 1u, 1u, &Unicode_});
  registry.register_function(FunctionDef{"CLEAN", 1u, 1u, &Clean});
  registry.register_function(FunctionDef{"PROPER", 1u, 1u, &Proper});

  // Byte-oriented text family (ja-JP DBCS).
  registry.register_function(FunctionDef{"LENB", 1u, 1u, &Lenb});
  registry.register_function(FunctionDef{"LEFTB", 1u, 2u, &Leftb});
  registry.register_function(FunctionDef{"RIGHTB", 1u, 2u, &Rightb});
  registry.register_function(FunctionDef{"MIDB", 3u, 3u, &Midb});
  registry.register_function(FunctionDef{"CHAR", 1u, 1u, &Char_});
  registry.register_function(FunctionDef{"CODE", 1u, 1u, &Code_});
}

}  // namespace eval
}  // namespace formulon
