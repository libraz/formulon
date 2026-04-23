// Copyright 2026 libraz. Licensed under the MIT License.
//
// Byte-oriented text family (ja-JP DBCS): LENB, LEFTB, RIGHTB, MIDB,
// REPLACEB, FINDB, SEARCHB. See `text.cpp` for the UTF-16-unit-based core
// family (LEN / LEFT / RIGHT / MID / FIND / SEARCH / REPLACE). The DBCS
// rule is documented near the LENB block below.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/builtins/text_detail.h"
#include "eval/coerce.h"
#include "eval/criteria.h"
#include "eval/text_ops.h"
#include "eval/utf8_length.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace text_detail {

// --- FINDB / SEARCHB shared helpers -------------------------------------

// FINDB(find_text, within_text, [start_num]) - case-sensitive, no wildcards.
// Returns the 1-based ja-JP DBCS byte position of the first occurrence of
// `find_text` in `within_text` at or after `start_num` (default 1, also in
// DBCS bytes). Not found / out-of-range `start_num` / `start_num < 1`
// surface `#VALUE!`. Empty `find_text` returns `start_num` (FIND's quirk).
Value FindB_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  const std::uint64_t total_dbcs = bytes_in_jajp(haystack.value());
  if (start < 1 || static_cast<std::uint64_t>(start) > total_dbcs + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::vector<DbcsCharRec> chars = build_dbcs_char_map(haystack.value());
  for (const DbcsCharRec& rec : chars) {
    if (rec.dbcs_position < static_cast<std::uint64_t>(start)) {
      continue;
    }
    if (haystack.value().compare(rec.byte_offset, needle.value().size(), needle.value()) == 0) {
      return Value::number(static_cast<double>(rec.dbcs_position));
    }
  }
  return Value::error(ErrorCode::Value);
}

// SEARCHB(find_text, within_text, [start_num]) - case-insensitive with
// DOS-style wildcards. Mirrors SEARCH semantically but position values are
// in DBCS bytes. `start_num` is also in DBCS bytes.
Value SearchB_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
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
  const std::uint64_t total_dbcs = bytes_in_jajp(haystack.value());
  if (start < 1 || static_cast<std::uint64_t>(start) > total_dbcs + 1) {
    return Value::error(ErrorCode::Value);
  }
  if (needle.value().empty()) {
    return Value::number(static_cast<double>(start));
  }
  const std::vector<DbcsCharRec> chars = build_dbcs_char_map(haystack.value());
  // Find the byte offset that corresponds to `start`. If `start` lands
  // strictly inside a 2-byte character, round up to the next character
  // boundary (matches SEARCH's rounding-up convention on UTF-16 splits).
  std::size_t start_byte = haystack.value().size();
  for (const DbcsCharRec& rec : chars) {
    if (rec.dbcs_position >= static_cast<std::uint64_t>(start)) {
      start_byte = rec.byte_offset;
      break;
    }
  }
  // Case-fold on both sides (ASCII only; matches SEARCH's policy).
  const std::string lowered_haystack = to_lower_ascii(haystack.value());
  const std::string lowered_needle = to_lower_ascii(needle.value());
  std::size_t match_byte = std::string::npos;
  const bool has_metachar = lowered_needle.find_first_of("*?~") != std::string::npos;
  if (!has_metachar) {
    match_byte = lowered_haystack.find(lowered_needle, start_byte);
  } else {
    const std::string_view suffix = std::string_view(lowered_haystack).substr(start_byte);
    const std::size_t rel = wildcard_find(lowered_needle, suffix);
    if (rel != std::string_view::npos) {
      match_byte = start_byte + rel;
    }
  }
  if (match_byte == std::string::npos) {
    return Value::error(ErrorCode::Value);
  }
  // Walk the character map to translate the byte offset into a 1-based
  // DBCS position. If the match lands mid-character (should not happen for
  // UTF-8-aligned needles) we fall back to the containing character's DBCS
  // position.
  for (const DbcsCharRec& rec : chars) {
    if (rec.byte_offset == match_byte) {
      return Value::number(static_cast<double>(rec.dbcs_position));
    }
    if (rec.byte_offset > match_byte) {
      break;
    }
  }
  // Defensive: if the match happens to land exactly at haystack.size()
  // (possible with wildcard on empty trailing), return the final position
  // after the last character.
  return Value::number(static_cast<double>(total_dbcs + 1));
}

}  // namespace text_detail

namespace {

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

}  // namespace

namespace text_detail {

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

// Builds the per-character DBCS map for a UTF-8 byte sequence. Each entry
// stores the character's UTF-8 byte offset and length, its DBCS cost, and
// its 1-based DBCS position. Used by FINDB / SEARCHB / REPLACEB to
// translate between UTF-8 byte offsets and ja-JP DBCS byte positions.
std::vector<DbcsCharRec> build_dbcs_char_map(std::string_view src) {
  std::vector<DbcsCharRec> out;
  out.reserve(src.size());
  std::uint64_t dbcs_cursor = 1;  // 1-based DBCS position of the next character
  std::size_t i = 0;
  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i);
    DbcsCharRec rec;
    rec.byte_offset = i;
    rec.byte_len = step.byte_len;
    rec.dbcs_position = dbcs_cursor;
    rec.dbcs_bytes = step.dbcs_bytes;
    out.push_back(rec);
    dbcs_cursor += static_cast<std::uint64_t>(step.dbcs_bytes);
    i += step.byte_len;
  }
  return out;
}

// REPLACEB(old_text, start_num, num_bytes, new_text) - ja-JP DBCS byte
// replace. Mirrors REPLACE but all offsets are in DBCS bytes. If
// `start_num` falls strictly inside a 2-byte character, a single ASCII
// space is emitted in place of the character's leading byte (MIDB's head-
// pad convention). Similarly, if the final byte to remove lands mid-
// character, the trailing byte of deletion is substituted with an ASCII
// space before `new_text` is appended. `start_num < 1` or `num_bytes < 0`
// -> `#VALUE!`. Result capped at Excel's 32,767-unit text limit.
Value ReplaceB_(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  // Excel caps the result of REPT (and a handful of related text functions)
  // at 32,767 UTF-16 units. We reuse the same constant in REPT's overflow
  // guard.
  constexpr std::uint64_t kExcelTextCapUnits = 32767u;

  auto old_text = coerce_to_text(args[0]);
  if (!old_text) {
    return Value::error(old_text.error());
  }
  auto start = read_int_arg(args[1]);
  if (!start) {
    return Value::error(start.error());
  }
  auto num_bytes = read_int_arg(args[2]);
  if (!num_bytes) {
    return Value::error(num_bytes.error());
  }
  auto new_text = coerce_to_text(args[3]);
  if (!new_text) {
    return Value::error(new_text.error());
  }
  if (start.value() < 1 || num_bytes.value() < 0) {
    return Value::error(ErrorCode::Value);
  }
  const std::string& src = old_text.value();
  const auto start_byte = static_cast<std::uint64_t>(start.value());  // 1-based
  const auto budget = static_cast<std::uint64_t>(num_bytes.value());  // bytes to remove
  const std::uint64_t end_byte_inclusive = start_byte + budget - 1;   // final byte to delete

  std::string prefix;
  std::string suffix;
  prefix.reserve(src.size());
  suffix.reserve(src.size());
  bool head_pad = false;
  bool tail_pad = false;
  std::uint64_t cursor = 0;  // DBCS bytes consumed so far (0-based, next char's first byte is cursor+1)
  std::size_t i = 0;
  bool in_deletion = false;  // we've emitted either the head_pad or started consuming the deletion window

  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i);
    const auto cost = static_cast<std::uint64_t>(step.dbcs_bytes);
    const std::uint64_t char_start = cursor + 1;   // 1-based first byte
    const std::uint64_t char_end = cursor + cost;  // 1-based last byte

    if (!in_deletion) {
      // Still accumulating the prefix or about to enter the deletion window.
      if (char_end < start_byte) {
        prefix.append(src, i, step.byte_len);
        cursor += cost;
        i += step.byte_len;
        continue;
      }
      if (char_start == start_byte) {
        // Deletion window starts cleanly at this character.
        in_deletion = true;
        // Fall through to the deletion handler below.
      } else {
        // Deletion window starts strictly inside this multi-byte character.
        // Emit a head-pad space in the prefix; conceptually the first byte
        // of this character has been "consumed" from the deletion budget.
        head_pad = true;
        prefix.push_back(' ');
        cursor += cost;
        i += step.byte_len;
        // The remaining (cost - (start_byte - char_start)) bytes of this
        // character are also deleted from the budget. Since the entire
        // character is skipped, `budget` effectively decreases by
        // `char_end - start_byte + 1` (the bytes from start_byte through
        // char_end). We advance directly past the character.
        const std::uint64_t consumed_from_budget = char_end - start_byte + 1;
        if (consumed_from_budget >= budget) {
          // Budget already exhausted inside this one character; nothing
          // more to delete. The tail of the deletion window also fell
          // within this character, so we do NOT emit a tail pad separately.
          in_deletion = true;
          // The rest of the loop walks suffix characters; mark we're past
          // the deletion window.
          break;
        }
        in_deletion = true;
        // Remaining budget after this character.
        const std::uint64_t remaining_budget = budget - consumed_from_budget;
        // Continue from the next character with the remaining budget.
        // We reset the effective window: simulate an in-progress deletion
        // by recomputing `end_byte_inclusive_local`.
        // To keep the logic simple, re-enter the main loop with adjusted
        // bookkeeping via a helper lambda-free approach: consume characters
        // fully or emit tail pad on partial overflow.
        std::uint64_t budget_left = remaining_budget;
        while (i < src.size() && budget_left > 0) {
          const Utf8Step s2 = next_utf8_step(src, i);
          const auto c2 = static_cast<std::uint64_t>(s2.dbcs_bytes);
          if (c2 <= budget_left) {
            cursor += c2;
            i += s2.byte_len;
            budget_left -= c2;
          } else {
            // Partial overlap: final byte of deletion lands inside this
            // character (only possible with c2 == 2, budget_left == 1).
            tail_pad = true;
            cursor += c2;
            i += s2.byte_len;
            budget_left = 0;
          }
        }
        break;
      }
    }
    // in_deletion is true: consume `budget` bytes starting from `start_byte`.
    // At this point `char_start == start_byte` on entry.
    {
      std::uint64_t budget_left = budget;
      while (i < src.size() && budget_left > 0) {
        const Utf8Step s2 = next_utf8_step(src, i);
        const auto c2 = static_cast<std::uint64_t>(s2.dbcs_bytes);
        if (c2 <= budget_left) {
          cursor += c2;
          i += s2.byte_len;
          budget_left -= c2;
        } else {
          tail_pad = true;
          cursor += c2;
          i += s2.byte_len;
          budget_left = 0;
        }
      }
      break;
    }
  }
  // Copy any remaining characters into the suffix.
  if (i < src.size()) {
    suffix.append(src, i, std::string::npos);
  }
  (void)end_byte_inclusive;
  (void)head_pad;

  std::string out;
  out.reserve(prefix.size() + new_text.value().size() + (tail_pad ? 1u : 0u) + suffix.size());
  out.append(prefix);
  out.append(new_text.value());
  if (tail_pad) {
    out.push_back(' ');
  }
  out.append(suffix);
  if (static_cast<std::uint64_t>(utf16_units_in(out)) > kExcelTextCapUnits) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

}  // namespace text_detail
}  // namespace eval
}  // namespace formulon
