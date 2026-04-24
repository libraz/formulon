// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the scalar coercion helpers declared in `coerce.h`.

#include "eval/coerce.h"

#include <cmath>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

#include "eval/date_text_parse.h"
#include "utils/double_format.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

// Parses `s` as a full double using std::strtod. Returns true iff the entire
// input (no leftover bytes) parsed cleanly and the result is finite; the
// finiteness guard lets callers treat true as "usable as a number". The stack
// buffer is sized for any IEEE-754 literal including subnormals; longer
// inputs take the heap path.
bool strtod_full(std::string_view s, double* out) {
  if (s.empty()) {
    return false;
  }
  char stack_buf[64];
  char* heap_buf = nullptr;
  const std::size_t n = s.size();
  char* buf = stack_buf;
  if (n + 1 > sizeof(stack_buf)) {
    heap_buf = static_cast<char*>(std::malloc(n + 1));
    if (heap_buf == nullptr) {
      return false;
    }
    buf = heap_buf;
  }
  std::memcpy(buf, s.data(), n);
  buf[n] = '\0';
  char* end_ptr = nullptr;
  const double parsed = std::strtod(buf, &end_ptr);
  const bool ok = end_ptr == buf + n;
  if (heap_buf != nullptr) {
    std::free(heap_buf);
  }
  if (!ok) {
    return false;
  }
  *out = parsed;
  return true;
}

// Currency symbols accepted by Mac Excel 365 for implicit numeric coercion.
// Allowlist of single-codepoint UTF-8 byte sequences; no locale lookup.
struct CurrencyToken {
  const char* bytes;
  std::size_t len;
};
constexpr CurrencyToken kCurrencyTokens[] = {
    {"\x24", 1},                  // $  U+0024
    {"\xC2\xA2", 2},              // cent  U+00A2
    {"\xC2\xA3", 2},              // pound  U+00A3
    {"\xC2\xA5", 2},              // yen  U+00A5
    {"\xE2\x82\xAC", 3},          // euro  U+20AC
    {"\xE2\x82\xA9", 3},          // won  U+20A9
};

bool try_strip_leading_currency(std::string_view s, std::string_view* out) {
  for (const auto& tok : kCurrencyTokens) {
    if (s.size() > tok.len && std::memcmp(s.data(), tok.bytes, tok.len) == 0) {
      *out = s.substr(tok.len);
      return true;
    }
  }
  return false;
}

bool try_strip_trailing_currency(std::string_view s, std::string_view* out) {
  for (const auto& tok : kCurrencyTokens) {
    if (s.size() > tok.len &&
        std::memcmp(s.data() + s.size() - tok.len, tok.bytes, tok.len) == 0) {
      *out = s.substr(0, s.size() - tok.len);
      return true;
    }
  }
  return false;
}

}  // namespace

Expected<double, ErrorCode> coerce_to_number(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      return d;
    }
    case ValueKind::Bool:
      return v.as_boolean() ? 1.0 : 0.0;
    case ValueKind::Blank:
      return 0.0;
    case ValueKind::Text: {
      const std::string_view trimmed = strings::trim(v.as_text());
      if (trimmed.empty()) {
        // Empty / whitespace-only text is #VALUE! in every numeric-coercion
        // context Mac Excel 365 was tested against (`=""+1`, `=SIN("")`,
        // `=EXP("")`, ... all yield #VALUE!). Blank cells still coerce to
        // 0 via the `ValueKind::Blank` branch above; only the explicit
        // empty string is rejected here.
        return ErrorCode::Value;
      }
      // Layered numeric-coercion fallback, in order:
      //   1. strtod(trimmed)                    - plain numeric fast path
      //   2. trailing '%' stripped, strtod, /100 - percent literals
      //   3. leading currency stripped, strtod   - "$100", "€100", ...
      //   4. trailing currency stripped, strtod  - "100$", "100€", ...
      //   5. date / datetime fallback (raw text) - DATEVALUE-style shapes
      //   6. #VALUE!
      // Percent + currency combinations and currency-only markers remain
      // rejected; the date fallback still runs against the raw, untrimmed
      // text so padded date strings stay #VALUE! (see WhitespacePaddedDate
      // rejection test).
      double parsed = 0.0;
      if (strtod_full(trimmed, &parsed)) {
        if (std::isnan(parsed) || std::isinf(parsed)) {
          return ErrorCode::Num;
        }
        return parsed;
      }
      if (trimmed.back() == '%') {
        const std::string_view body = trimmed.substr(0, trimmed.size() - 1);
        if (strtod_full(body, &parsed)) {
          const double scaled = parsed / 100.0;
          if (std::isnan(scaled) || std::isinf(scaled)) {
            return ErrorCode::Num;
          }
          return scaled;
        }
      }
      std::string_view stripped;
      if (try_strip_leading_currency(trimmed, &stripped) && strtod_full(stripped, &parsed)) {
        if (std::isnan(parsed) || std::isinf(parsed)) {
          return ErrorCode::Num;
        }
        return parsed;
      }
      if (try_strip_trailing_currency(trimmed, &stripped) && strtod_full(stripped, &parsed)) {
        if (std::isnan(parsed) || std::isinf(parsed)) {
          return ErrorCode::Num;
        }
        return parsed;
      }
      // Mac Excel 365 accepts date / datetime text wherever a number is
      // expected: e.g. `=FLOOR(10, "2024-01-10")` coerces the second
      // argument to its serial (45301). Reuse the shared DATEVALUE /
      // TIMEVALUE / VALUE parser; only fires after the numeric fallbacks
      // have rejected the input so plain numerics keep their fast path.
      // The raw, un-trimmed text is passed: implicit numeric coercion is
      // strict about whitespace around date strings (`=FLOOR(10,
      // " 2024-01-10 ")` -> #VALUE!), even though `strtod` and DATEVALUE
      // both tolerate it.
      double serial = 0.0;
      double frac = 0.0;
      bool has_date = false;
      bool has_time = false;
      if (date_parse::parse_date_time_text(v.as_text(), &serial, &frac, &has_date, &has_time)) {
        const double combined = serial + frac;
        if (std::isnan(combined) || std::isinf(combined)) {
          return ErrorCode::Num;
        }
        return combined;
      }
      return ErrorCode::Value;
    }
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<std::string, ErrorCode> coerce_to_text(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Number: {
      std::string out;
      format_double(out, v.as_number());
      return out;
    }
    case ValueKind::Bool:
      return std::string(v.as_boolean() ? "TRUE" : "FALSE");
    case ValueKind::Blank:
      return std::string();
    case ValueKind::Text:
      return std::string(v.as_text());
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<double, ErrorCode> apply_pow(double base, double exp) {
  // Excel treats 0^0 as indeterminate and reports #NUM!, diverging from the
  // IEEE-754 pow convention of 1. Guarded explicitly before the std::pow
  // call so both the POWER() builtin and the `^` binary operator share the
  // same behaviour.
  if (base == 0.0 && exp == 0.0) {
    return ErrorCode::Num;
  }
  // For all other cases Excel matches std::pow: negative base with a
  // non-integer exponent yields NaN -> #NUM!, and overflow / underflow to
  // Inf also yields #NUM!.
  const double r = std::pow(base, exp);
  if (std::isnan(r) || std::isinf(r)) {
    return ErrorCode::Num;
  }
  return r;
}

Expected<bool, ErrorCode> coerce_to_bool(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Bool:
      return v.as_boolean();
    case ValueKind::Number: {
      const double d = v.as_number();
      if (std::isnan(d) || std::isinf(d)) {
        return ErrorCode::Num;
      }
      return d != 0.0;
    }
    case ValueKind::Blank:
      return false;
    case ValueKind::Text: {
      // Excel coerces text to bool by routing through the numeric rule:
      // the text is parsed as a number, and a successful parse becomes
      // `false` iff the value is exactly zero. The literal strings
      // `"TRUE"` and `"FALSE"` are NOT recognised here (they fail the
      // numeric parse and surface as `#VALUE!`); only bool literals
      // (TRUE / FALSE without quotes) and numeric strings round-trip.
      auto coerced = coerce_to_number(v);
      if (!coerced) {
        return coerced.error();
      }
      return coerced.value() != 0.0;
    }
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

}  // namespace eval
}  // namespace formulon
