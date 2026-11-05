//
// Byte-oriented text family: LENB, LEFTB, RIGHTB, MIDB, REPLACEB, FINDB,
// SEARCHB. DBCS profiles count code-page bytes (`dbcs_char_bytes`); other
// profiles use Excel's UTF-16-unit rule. See `text.cpp` for the corresponding non-B family.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/builtins/text_detail.h"
#include "eval/coerce.h"
#include "eval/eval_profile_scope.h"
#include "excel_locale.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

bool uses_dbcs_profile() noexcept {
  return locale_facts(current_eval_profile()).dbcs_codepage != DbcsCodepage::kNone;
}

DbcsCodepage dbcs_codepage() noexcept {
  return locale_facts(current_eval_profile()).dbcs_codepage;
}

}  // namespace

namespace text_detail {

// --- FINDB / SEARCHB shared helpers -------------------------------------

// FINDB(find_text, within_text, [start_num]) - case-sensitive, no wildcards.
// Returns the 1-based profile-specific byte/unit position of the first
// occurrence of `find_text` in `within_text` at or after `start_num`.
// Not found / out-of-range `start_num` / `start_num < 1` surface `#VALUE!`.
// Empty `find_text` returns `start_num` (FIND's quirk).
Value FindB_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const bool dbcs = uses_dbcs_profile();
  SearchArgs sargs;
  Value early = Value::blank();
  if (!read_search_args(args, arity, dbcs ? SearchUnit::DbcsByte : SearchUnit::Utf16, &sargs, &early)) {
    return early;
  }
  if (!dbcs) {
    return find_utf16_exact(sargs);
  }
  const std::vector<DbcsCharRec> chars = build_dbcs_char_map(sargs.haystack, dbcs_codepage());
  for (const DbcsCharRec& rec : chars) {
    if (rec.dbcs_position < static_cast<std::uint64_t>(sargs.start)) {
      continue;
    }
    if (sargs.haystack.compare(rec.byte_offset, sargs.needle.size(), sargs.needle) == 0) {
      return Value::number(static_cast<double>(rec.dbcs_position));
    }
  }
  return Value::error(ErrorCode::Value);
}

// SEARCHB(find_text, within_text, [start_num]) - case-insensitive with
// DOS-style wildcards. Mirrors SEARCH semantically in the active profile's
// byte/unit position space.
Value SearchB_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  const bool dbcs = uses_dbcs_profile();
  SearchArgs sargs;
  Value early = Value::blank();
  if (!read_search_args(args, arity, dbcs ? SearchUnit::DbcsByte : SearchUnit::Utf16, &sargs, &early)) {
    return early;
  }
  if (!dbcs) {
    return find_utf16_folded(sargs);
  }
  const DbcsCodepage codepage = dbcs_codepage();
  const std::vector<DbcsCharRec> chars = build_dbcs_char_map(sargs.haystack, codepage);
  // Find the byte offset that corresponds to `start`. If `start` lands
  // strictly inside a 2-byte character, round up to the next character
  // boundary (matches SEARCH's rounding-up convention on UTF-16 splits).
  std::size_t start_byte = sargs.haystack.size();
  for (const DbcsCharRec& rec : chars) {
    if (rec.dbcs_position >= static_cast<std::uint64_t>(sargs.start)) {
      start_byte = rec.byte_offset;
      break;
    }
  }
  // In ja / ko SEARCHB's `?` matches only SBCS codepoints, so
  // `=SEARCHB("?","漢ABC")` skips the kanji and lands on the 'A'.
  const std::size_t match_byte = find_folded(sargs.haystack, sargs.needle, start_byte, SearchUnit::DbcsByte);
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
  return Value::number(static_cast<double>(dbcs_bytes_in(sargs.haystack, codepage) + 1));
}

}  // namespace text_detail

namespace {

// --- Byte-oriented text family (LENB / LEFTB / RIGHTB / MIDB / CHAR / CODE)
//
// A DBCS profile counts "byte" length with `dbcs_char_bytes`: ASCII is 1
// byte, a BMP character takes the locale's width (1 or 2), and a
// supplementary-plane code point such as an emoji counts as 2 bytes.
//
// When a byte-budgeted slice (LEFTB/RIGHTB/MIDB) would land in the middle
// of a 2-byte character (i.e. we have exactly 1 byte of budget left and the
// next character costs 2), Excel substitutes a single ASCII space (0x20)
// and stops. MIDB additionally substitutes a space when its `start_byte`
// falls inside a 2-byte character.

// Decode a single UTF-8 codepoint starting at `src[i]`. On malformed input
// returns a 1-byte "raw" step so the caller can still make progress; such
// steps are classified as 1 DBCS byte.
struct Utf8Step {
  std::uint32_t codepoint;  // Decoded codepoint (0xFFFD on malformed input).
  std::size_t byte_len;     // Number of UTF-8 bytes consumed (>= 1).
  int dbcs_bytes;           // DBCS byte cost under the active profile.
};

Utf8Step next_utf8_step(std::string_view src, std::size_t i, DbcsCodepage codepage) noexcept {
  std::size_t bytes = 0;
  const std::uint32_t cp = decode_utf8_step(src, i, &bytes);
  // Lenient decoder reports U+FFFD on malformed input with a 1-byte
  // advance. The DBCS cost for replacement / single-byte fallbacks is
  // 1 (a single ASCII-equivalent byte), matching the prior contract.
  if (bytes == 1 && cp == 0xFFFDu) {
    return {0xFFFDu, 1u, 1};
  }
  return {cp, bytes, text_detail::dbcs_char_bytes(cp, codepage)};
}

// True when a character of `cost` bytes would straddle a budget with exactly one byte left.
inline bool straddles_last_byte(std::uint64_t used, std::uint64_t budget, std::uint64_t cost) noexcept {
  return used + 1 == budget && cost >= 2;
}

// Reads the `text, [num_bytes=1]` prologue LEFTB and RIGHTB share.
// Returns `false` when the caller must return `*out_early` instead of
// slicing: a coercion error, a negative budget (`#VALUE!`), or a zero
// budget (the empty string).
bool read_byte_budget_args(const Value* args, std::uint32_t arity, std::string* out_text, std::uint64_t* out_budget,
                           Value* out_early) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    *out_early = Value::error(text.error());
    return false;
  }
  auto byte_count = arity > 1u ? text_detail::read_snapped_int_arg(args[1], 0.0) : Expected<int, ErrorCode>(1);
  if (!byte_count) {
    *out_early = Value::error(byte_count.error());
    return false;
  }
  const int budget = byte_count.value();
  if (budget == 0) {
    *out_early = Value::text({});
    return false;
  }
  *out_text = std::move(text.value());
  *out_budget = static_cast<std::uint64_t>(budget);
  return true;
}

Value leftb_utf16(const Value* args, std::uint32_t arity, Arena& arena) {
  std::string src;
  std::uint64_t budget = 0;
  Value early = Value::blank();
  if (!read_byte_budget_args(args, arity, &src, &budget, &early)) {
    return early;
  }
  return Value::text(arena.intern(utf16_substring(src, 0u, static_cast<std::uint32_t>(budget))));
}

Value rightb_utf16(const Value* args, std::uint32_t arity, Arena& arena) {
  std::string src;
  std::uint64_t budget = 0;
  Value early = Value::blank();
  if (!read_byte_budget_args(args, arity, &src, &budget, &early)) {
    return early;
  }
  const std::uint32_t total = utf16_units_in(src);
  const std::uint32_t take = budget >= total ? total : static_cast<std::uint32_t>(budget);
  return Value::text(arena.intern(utf16_substring(src, total - take, take)));
}

}  // namespace

namespace text_detail {

// LENB(text) - returns the active profile's byte-oriented length. Non-DBCS
// profiles use Excel's UTF-16 code-unit count for the *B family.
Value Lenb(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const double length = uses_dbcs_profile() ? static_cast<double>(dbcs_bytes_in(text.value(), dbcs_codepage()))
                                            : static_cast<double>(utf16_units_in(text.value()));
  return Value::number(length);
}

// LEFTB(text, [num_bytes=1]) - leftmost `num_bytes` bytes. When a 2-byte
// character would straddle the budget (1 byte remaining), emit an ASCII
// space in its place and stop.
Value Leftb(const Value* args, std::uint32_t arity, Arena& arena) {
  if (!uses_dbcs_profile()) {
    return leftb_utf16(args, arity, arena);
  }
  const DbcsCodepage codepage = dbcs_codepage();
  std::string src;
  std::uint64_t budget = 0;
  Value early = Value::blank();
  if (!read_byte_budget_args(args, arity, &src, &budget, &early)) {
    return early;
  }
  std::uint64_t used = 0;
  std::string out;
  out.reserve(src.size());
  std::size_t i = 0;
  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i, codepage);
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
      if (straddles_last_byte(used, budget, cost)) {
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
  if (!uses_dbcs_profile()) {
    return rightb_utf16(args, arity, arena);
  }
  const DbcsCodepage codepage = dbcs_codepage();
  std::string src;
  std::uint64_t budget = 0;
  Value early = Value::blank();
  if (!read_byte_budget_args(args, arity, &src, &budget, &early)) {
    return early;
  }

  const std::vector<DbcsCharRec> chars = build_dbcs_char_map(src, codepage);

  // Walk right-to-left accumulating cost until we can't fit the next char.
  std::uint64_t used = 0;
  std::size_t first_full_idx = chars.size();
  bool pad_space = false;
  for (std::size_t k = chars.size(); k > 0; --k) {
    const DbcsCharRec& rec = chars[k - 1];
    const auto cost = static_cast<std::uint64_t>(rec.dbcs_bytes);
    if (used + cost <= budget) {
      used += cost;
      first_full_idx = k - 1;
      if (used == budget) {
        break;
      }
    } else {
      if (straddles_last_byte(used, budget, cost)) {
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
    const std::size_t start_byte = chars[first_full_idx].byte_offset;
    out.append(src, start_byte, std::string::npos);
  }
  return Value::text(arena.intern(out));
}

// MIDB(text, start_byte, num_bytes) - byte-window slice.
// `start_byte<1` or `num_bytes<0` -> `#VALUE!`. If `start_byte > LENB(text)`,
// return "". If `start_byte` lands inside a 2-byte character, emit a space
// pad and consume 1 byte of budget; from there walk normally, padding on
// the trailing edge with the same 1-byte-overflow rule as LEFTB.
Value Midb(const Value* args, std::uint32_t arity, Arena& arena) {
  if (!uses_dbcs_profile()) {
    return MidUtf16(args, arity, arena);
  }
  const DbcsCodepage codepage = dbcs_codepage();
  auto parsed = read_text_window_args(args, arity, /*snap_start=*/false);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const std::string& src = parsed.value().text;
  const int start = parsed.value().start;
  const int length = parsed.value().count;
  if (length == 0) {
    return Value::text({});
  }
  const auto start_byte = static_cast<std::uint64_t>(start);  // 1-based
  const auto budget = static_cast<std::uint64_t>(length);

  // Walk characters and track byte cursor.
  std::uint64_t cursor = 0;  // Bytes consumed so far (DBCS byte index).
  std::size_t i = 0;
  std::string out;
  out.reserve(src.size());
  std::uint64_t used = 0;
  bool started = false;
  bool head_pad = false;

  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i, codepage);
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
      if (straddles_last_byte(used, budget, cost)) {
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
// translate between UTF-8 byte offsets and DBCS byte positions.
std::vector<DbcsCharRec> build_dbcs_char_map(std::string_view src, DbcsCodepage codepage) {
  std::vector<DbcsCharRec> out;
  out.reserve(src.size());
  std::uint64_t dbcs_cursor = 1;  // 1-based DBCS position of the next character
  std::size_t i = 0;
  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i, codepage);
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

// REPLACEB(old_text, start_num, num_bytes, new_text) - DBCS byte
// replace. Mirrors REPLACE but all offsets are in DBCS bytes. If
// `start_num` falls strictly inside a 2-byte character, a single ASCII
// space is emitted in place of the character's leading byte (MIDB's head-
// pad convention). Similarly, if the final byte to remove lands mid-
// character, the trailing byte of deletion is substituted with an ASCII
// space before `new_text` is appended. `start_num < 1` or `num_bytes < 0`
// -> `#VALUE!`. Result capped at Excel's 32,767-unit text limit.
Value ReplaceB_(const Value* args, std::uint32_t arity, Arena& arena) {
  if (!uses_dbcs_profile()) {
    return ReplaceUtf16(args, arity, arena);
  }
  const DbcsCodepage codepage = dbcs_codepage();
  auto parsed = read_text_window_args(args, arity, /*snap_start=*/true);
  if (!parsed) {
    return Value::error(parsed.error());
  }
  const std::string& src = parsed.value().text;
  const std::string& new_text = parsed.value().new_text;
  const auto start_byte = static_cast<std::uint64_t>(parsed.value().start);  // 1-based
  const auto budget = static_cast<std::uint64_t>(parsed.value().count);      // bytes to remove

  std::string prefix;
  std::string suffix;
  prefix.reserve(src.size());
  suffix.reserve(src.size());
  bool tail_pad = false;
  std::size_t i = 0;

  // Drops whole characters from `i` until `left` bytes are used; a character straddling the end of the
  // window is dropped whole and replaced by a tail pad.
  const auto consume_budget = [&](std::uint64_t left) {
    while (i < src.size() && left > 0) {
      const Utf8Step step = next_utf8_step(src, i, codepage);
      const auto cost = static_cast<std::uint64_t>(step.dbcs_bytes);
      if (cost > left) {
        tail_pad = true;
        left = 0;
      } else {
        left -= cost;
      }
      i += step.byte_len;
    }
  };

  std::uint64_t cursor = 0;  // DBCS bytes consumed so far
  while (i < src.size()) {
    const Utf8Step step = next_utf8_step(src, i, codepage);
    const auto cost = static_cast<std::uint64_t>(step.dbcs_bytes);
    const std::uint64_t char_start = cursor + 1;
    const std::uint64_t char_end = cursor + cost;
    if (char_end < start_byte) {
      prefix.append(src, i, step.byte_len);
      cursor += cost;
      i += step.byte_len;
      continue;
    }
    if (char_start == start_byte) {
      consume_budget(budget);
    } else {
      // The window starts inside this character: it becomes a head pad and its remaining bytes count
      // against the budget.
      prefix.push_back(' ');
      i += step.byte_len;
      const std::uint64_t consumed = char_end - start_byte + 1;
      if (consumed < budget) {
        consume_budget(budget - consumed);
      }
    }
    break;
  }
  if (i < src.size()) {
    suffix.append(src, i, std::string::npos);
  }

  std::string out;
  out.reserve(prefix.size() + new_text.size() + (tail_pad ? 1u : 0u) + suffix.size());
  out.append(prefix);
  out.append(new_text);
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
