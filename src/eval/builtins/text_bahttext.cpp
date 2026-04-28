// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Excel's BAHTTEXT function: spells out a numeric value in
// Thai script as Thai baht and satang. The output is locale-independent: Excel
// always emits Thai script regardless of the active UI locale.
//
// Output format: `[ลบ]<integer-baht-words>บาท<satang-clause>` where the
// satang-clause is `ถ้วน` ("exactly") for an integer amount or
// `<satang-words>สตางค์` for a fractional amount. Negative amounts are
// prefixed with `ลบ`.
//
// The number is rounded to two decimal places (away from zero) before
// processing; absolute values that reach `1e17` or beyond surface `#VALUE!`.
// Older Microsoft documentation specified a `1e15` ceiling, but Mac Excel 365
// (16.108.1) accepts inputs up through the chained-`ล้าน` range — e.g.
// `=BAHTTEXT(1E15)` evaluates to `หนึ่งพันล้านล้านบาทถ้วน` rather than
// `#VALUE!`. The `1e17` cutoff stays safely below the u64 overflow boundary
// in the satang scaling step.
//
// Reading rules (all rendered into UTF-8 byte sequences below):
//   * The integer part is read in chunks of 6 digits separated by `ล้าน`
//     (laan = "million"). Chained `ล้าน` are emitted as-is for
//     `>= 10^12`, e.g. `1e12` reads as `หนึ่งล้านล้าน`.
//   * Within a chunk, place-value suffixes (`สิบ`, `ร้อย`, `พัน`, `หมื่น`,
//     `แสน`) follow each non-zero digit; zero digits are skipped.
//   * The tens digit `1` is rendered as bare `สิบ` (the `หนึ่ง` digit-word
//     is dropped). The tens digit `2` is `ยี่` instead of `สอง`.
//   * The ones digit `1` becomes `เอ็ด` when *any* higher digit in the chunk
//     is non-zero (so 11 -> `สิบเอ็ด`, 21 -> `ยี่สิบเอ็ด`, 101 ->
//     `หนึ่งร้อยเอ็ด`); a standalone 1 stays `หนึ่ง`.

#include "eval/builtins/text_bahttext.h"

#include <cmath>
#include <cstdint>
#include <string>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Thai digit-words (UTF-8). Index = digit value 0..9.
constexpr const char* const kThaiDigit[10] = {
    "\xE0\xB8\xA8\xE0\xB8\xB9\xE0\xB8\x99\xE0\xB8\xA2\xE0\xB9\x8C",  // 0 ศูนย์
    "\xE0\xB8\xAB\xE0\xB8\x99\xE0\xB8\xB6\xE0\xB9\x88\xE0\xB8\x87",  // 1 หนึ่ง
    "\xE0\xB8\xAA\xE0\xB8\xAD\xE0\xB8\x87",                          // 2 สอง
    "\xE0\xB8\xAA\xE0\xB8\xB2\xE0\xB8\xA1",                          // 3 สาม
    "\xE0\xB8\xAA\xE0\xB8\xB5\xE0\xB9\x88",                          // 4 สี่
    "\xE0\xB8\xAB\xE0\xB9\x89\xE0\xB8\xB2",                          // 5 ห้า
    "\xE0\xB8\xAB\xE0\xB8\x81",                                      // 6 หก
    "\xE0\xB9\x80\xE0\xB8\x88\xE0\xB9\x87\xE0\xB8\x94",              // 7 เจ็ด
    "\xE0\xB9\x81\xE0\xB8\x9B\xE0\xB8\x94",                          // 8 แปด
    "\xE0\xB9\x80\xE0\xB8\x81\xE0\xB9\x89\xE0\xB8\xB2",              // 9 เก้า
};

// Place-value suffixes (UTF-8).
constexpr const char* const kSip = "\xE0\xB8\xAA\xE0\xB8\xB4\xE0\xB8\x9A";              // สิบ (10)
constexpr const char* const kRoi = "\xE0\xB8\xA3\xE0\xB9\x89\xE0\xB8\xAD\xE0\xB8\xA2";  // ร้อย (100)
constexpr const char* const kPan = "\xE0\xB8\x9E\xE0\xB8\xB1\xE0\xB8\x99";              // พัน (1,000)
constexpr const char* const kMuen = "\xE0\xB8\xAB\xE0\xB8\xA1\xE0\xB8\xB7\xE0\xB9\x88\xE0\xB8\x99";  // หมื่น (10,000)
constexpr const char* const kSaen = "\xE0\xB9\x81\xE0\xB8\xAA\xE0\xB8\x99";              // แสน (100,000)
constexpr const char* const kLaan = "\xE0\xB8\xA5\xE0\xB9\x89\xE0\xB8\xB2\xE0\xB8\x99";  // ล้าน (1,000,000)

// Special Thai readings for the tens / ones positions.
constexpr const char* const kYi = "\xE0\xB8\xA2\xE0\xB8\xB5\xE0\xB9\x88";  // ยี่ (= 2 in tens)
constexpr const char* const kEt =
    "\xE0\xB9\x80\xE0\xB8\xAD\xE0\xB9\x87\xE0\xB8\x94";  // เอ็ด (= 1 ones, when not standalone)

// Currency / sign words (UTF-8).
constexpr const char* const kBaht = "\xE0\xB8\x9A\xE0\xB8\xB2\xE0\xB8\x97";  // บาท
constexpr const char* const kSatang =
    "\xE0\xB8\xAA\xE0\xB8\x95\xE0\xB8\xB2\xE0\xB8\x87\xE0\xB8\x84\xE0\xB9\x8C";           // สตางค์
constexpr const char* const kThuan = "\xE0\xB8\x96\xE0\xB9\x89\xE0\xB8\xA7\xE0\xB8\x99";  // ถ้วน
constexpr const char* const kLop = "\xE0\xB8\xA5\xE0\xB8\x9A";                            // ลบ

// Spells out a 0..999999 chunk in Thai text. Returns the empty string when
// `value` is zero so the multi-chunk caller can elide an empty group. The
// `is_top_chunk` flag selects between the "standalone" reading rules
// (for an integer that is just `1`, render `หนึ่ง`) and the "subordinate"
// reading rules (the ones digit `1` becomes `เอ็ด` whenever any higher
// digit in the chunk is non-zero, e.g. 11 -> `สิบเอ็ด`).
std::string spell_six_digit_chunk(std::uint32_t value, bool is_top_chunk) {
  if (value == 0) {
    return {};
  }

  // Decompose into 6 digits, most-significant first: hundred-thousands,
  // ten-thousands, thousands, hundreds, tens, ones.
  std::uint32_t digits[6];              // NOLINT(modernize-avoid-c-arrays)
  digits[0] = (value / 100000u) % 10u;  // แสน
  digits[1] = (value / 10000u) % 10u;   // หมื่น
  digits[2] = (value / 1000u) % 10u;    // พัน
  digits[3] = (value / 100u) % 10u;     // ร้อย
  digits[4] = (value / 10u) % 10u;      // สิบ
  digits[5] = value % 10u;              // ones

  static constexpr const char* const kPlace[5] = {kSaen, kMuen, kPan, kRoi, kSip};

  std::string out;
  out.reserve(64);

  // Higher places: แสน, หมื่น, พัน, ร้อย, สิบ.
  for (std::uint32_t i = 0; i < 5; ++i) {
    const std::uint32_t d = digits[i];
    if (d == 0) {
      continue;
    }
    if (i == 4) {
      // Tens place: digit 1 collapses to bare `สิบ`; digit 2 becomes `ยี่`.
      if (d == 1) {
        out.append(kSip);
      } else if (d == 2) {
        out.append(kYi);
        out.append(kSip);
      } else {
        out.append(kThaiDigit[d]);
        out.append(kSip);
      }
    } else {
      out.append(kThaiDigit[d]);
      out.append(kPlace[i]);
    }
  }

  // Ones place. The ones digit `1` is read as `เอ็ด` whenever any higher
  // digit in this chunk is non-zero (Excel matches on the chunk as a whole,
  // not just on the tens place). When the chunk is exactly `1` AND it is
  // the top-most chunk of the integer, render `หนึ่ง`.
  const std::uint32_t ones = digits[5];
  if (ones != 0) {
    const bool any_higher = digits[0] != 0 || digits[1] != 0 || digits[2] != 0 || digits[3] != 0 || digits[4] != 0;
    if (ones == 1 && any_higher) {
      out.append(kEt);
    } else if (ones == 1 && !is_top_chunk) {
      // Subordinate single-`1` chunk inside a multi-chunk number (e.g. the
      // `1` in `1,000,001`) is also `เอ็ด`.
      out.append(kEt);
    } else {
      out.append(kThaiDigit[ones]);
    }
  }
  return out;
}

// Spells out a non-negative integer in Thai text. Chunks the value by
// 1,000,000 from the right and joins with `ล้าน`. Empty chunks (zero groups)
// are skipped — the `ล้าน` separators between non-empty groups still emit
// once per non-empty boundary, so chained `ล้าน` (e.g. 10^12) appear when
// adjacent non-empty groups are separated by one or more zero groups.
std::string spell_integer(std::uint64_t value) {
  if (value == 0) {
    return kThaiDigit[0];  // ศูนย์
  }

  // Split into 6-digit chunks, least-significant first. The caller's `< 1e17`
  // input ceiling means the satang-scaled integer is `< 1e19`, which fits in
  // 3 chunks of up to 999,999 each.
  std::uint32_t chunks[3];  // NOLINT(modernize-avoid-c-arrays)
  std::uint32_t chunk_count = 0;
  std::uint64_t remaining = value;
  while (remaining > 0 && chunk_count < 3) {
    chunks[chunk_count++] = static_cast<std::uint32_t>(remaining % 1000000u);
    remaining /= 1000000u;
  }
  // Anything left over would exceed our ceiling; the caller has already
  // rejected such values.
  (void)remaining;

  // Walk from the highest chunk down, emitting `ล้าน` between every
  // non-empty pair. The most-significant non-empty chunk uses the
  // standalone reading rule; lower chunks use the subordinate rule.
  std::string out;
  out.reserve(64);
  bool emitted_any = false;
  for (std::uint32_t i = chunk_count; i-- > 0;) {
    const bool is_top_chunk = !emitted_any;
    std::string chunk_text = spell_six_digit_chunk(chunks[i], is_top_chunk);
    if (chunk_text.empty()) {
      // Zero group - skip the chunk and the immediately-following ล้าน
      // boundary. (We never want a bare `ล้าน` with nothing in front.)
      continue;
    }
    if (emitted_any) {
      // The previous (higher) chunk was non-empty; emit one `ล้าน` per
      // chunk-boundary that has a non-empty group on either side.
      out.append(kLaan);
    }
    out.append(chunk_text);
    emitted_any = true;
    // For each empty chunk that lies *below* this one but *above* the next
    // non-empty chunk, we still owe one `ล้าน`. Excel chains them.
    // Implementation: a single `ล้าน` per chunk-step; emit the extras here.
    if (i > 0) {
      // We will append a `ล้าน` for the boundary between this chunk and
      // the next non-empty one *via the loop above*. But empty chunks in
      // between also each contribute one `ล้าน` (chained). Count them.
      std::uint32_t empty_below = 0;
      for (std::uint32_t j = i; j-- > 0;) {
        if (chunks[j] == 0) {
          ++empty_below;
        } else {
          break;
        }
      }
      for (std::uint32_t k = 0; k < empty_below; ++k) {
        out.append(kLaan);
      }
    }
  }
  return out;
}

// Top-level formatter. `n` is finite and unrestricted in sign; the absolute
// value must already be < 1e17 (the caller enforces this).
std::string format_bahttext(double n) {
  // Round to 2 decimals away from zero, then split into integer-baht and
  // satang components.
  const double sign = (n < 0.0) ? -1.0 : 1.0;
  const double abs_n = (n < 0.0) ? -n : n;
  const double scaled = std::floor(abs_n * 100.0 + 0.5);  // away-from-zero on the magnitude
  const auto total_satang = static_cast<std::uint64_t>(scaled);
  const std::uint64_t integer_baht = total_satang / 100ull;
  const std::uint32_t satang = static_cast<std::uint32_t>(total_satang % 100ull);
  const bool negative = (sign < 0.0) && total_satang != 0;  // -0.00 -> no `ลบ`.

  std::string out;
  out.reserve(64);
  if (negative) {
    out.append(kLop);
  }
  // Integer-baht clause: always emitted, even for zero (`ศูนย์บาท...`).
  out.append(spell_integer(integer_baht));
  out.append(kBaht);

  // Satang clause: `ถ้วน` when zero, otherwise the spell-out + `สตางค์`.
  if (satang == 0) {
    out.append(kThuan);
  } else {
    out.append(spell_integer(satang));
    out.append(kSatang);
  }
  return out;
}

// Mac Excel 365 (16.108.1) accepts BAHTTEXT inputs that older Microsoft docs
// flagged as out-of-range: `=BAHTTEXT(1E15)` returns the chained-`ล้าน`
// spell-out (`หนึ่งพันล้านล้านบาทถ้วน`), not `#VALUE!`. The `1e17` cutoff is
// chosen to stay safely below the u64 overflow boundary of the
// `floor(|n|*100 + 0.5)` satang scaling step.
constexpr double kBahttextLimit = 1e17;

Value Bahttext(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  auto coerced = coerce_to_number(args[0]);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double n = coerced.value();
  if (!std::isfinite(n)) {
    return Value::error(ErrorCode::Value);
  }
  const double abs_n = (n < 0.0) ? -n : n;
  if (abs_n >= kBahttextLimit) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(format_bahttext(n)));
}

}  // namespace

void register_bahttext_builtin(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"BAHTTEXT", 1u, 1u, &Bahttext});
}

}  // namespace eval
}  // namespace formulon
