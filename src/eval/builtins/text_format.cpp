// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's text-conversion builtins: TEXT, VALUE, and
// NUMBERVALUE. All three mediate between numeric and textual
// representations, so they share the format-string engine in
// `eval/text_format/number_format.h` (TEXT) and the date/time parser in
// `eval/date_text_parse.h` (VALUE).

#include "eval/builtins/text_format.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/date_text_parse.h"
#include "eval/function_registry.h"
#include "eval/text_format/number_format.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Numeric parsing used by VALUE and NUMBERVALUE.
// ---------------------------------------------------------------------------

bool is_ascii_ws(unsigned char c) noexcept {
  return c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\v' || c == '\f';
}

std::string_view trim_ascii(std::string_view s) noexcept {
  while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.front()))) {
    s.remove_prefix(1);
  }
  while (!s.empty() && is_ascii_ws(static_cast<unsigned char>(s.back()))) {
    s.remove_suffix(1);
  }
  return s;
}

// Strips a leading currency prefix recognised by Excel's VALUE. Returns
// the remaining view. Supported prefixes: ASCII '$', UTF-8 '¥' (0xC2 0xA5),
// UTF-8 '￥' (0xEF 0xBF 0xA5). The prefix is optional and the caller must
// still handle a missing one.
std::string_view strip_currency(std::string_view s) noexcept {
  if (s.empty()) {
    return s;
  }
  if (s.front() == '$') {
    return s.substr(1);
  }
  if (s.size() >= 2 && static_cast<unsigned char>(s[0]) == 0xC2u && static_cast<unsigned char>(s[1]) == 0xA5u) {
    return s.substr(2);
  }
  if (s.size() >= 3 && static_cast<unsigned char>(s[0]) == 0xEFu && static_cast<unsigned char>(s[1]) == 0xBFu &&
      static_cast<unsigned char>(s[2]) == 0xA5u) {
    return s.substr(3);
  }
  return s;
}

// Parses a numeric string using `decimal_sep` and `group_sep`. `group_sep`
// is ignored during parsing (it may appear any number of times to the left
// of the decimal point). A leading sign (`+` / `-`), an optional currency
// prefix, and a trailing `%` (applied AFTER the numeric parse) are all
// accepted. Returns true on success and writes the parsed value into
// `*out`.
bool parse_numeric(std::string_view s, char decimal_sep, char group_sep, double* out) noexcept {
  // Trim ASCII whitespace first.
  s = trim_ascii(s);
  if (s.empty()) {
    return false;
  }
  // Sign.
  bool negative = false;
  if (s.front() == '+' || s.front() == '-') {
    negative = s.front() == '-';
    s.remove_prefix(1);
    if (s.empty()) {
      return false;
    }
  }
  // Optional currency symbol.
  s = strip_currency(s);
  if (s.empty()) {
    return false;
  }
  // Trailing percent signs. Each `%` multiplies the parsed value by 0.01,
  // so `"50%%"` yields `0.005` (matches Mac Excel ja-JP NUMBERVALUE).
  int percent_count = 0;
  while (!s.empty() && s.back() == '%') {
    ++percent_count;
    s.remove_suffix(1);
  }
  if (s.empty()) {
    return false;
  }
  // Scan and assemble a canonical C-locale numeric string (digits, one
  // optional `.`, optional exponent `e[+/-]digits`). Reject on any
  // unexpected byte.
  std::string canonical;
  canonical.reserve(s.size());
  bool seen_digit = false;
  bool seen_point = false;
  bool seen_exp = false;
  for (std::size_t i = 0; i < s.size(); ++i) {
    const char c = s[i];
    if (c >= '0' && c <= '9') {
      canonical.push_back(c);
      seen_digit = true;
      continue;
    }
    if (c == decimal_sep && !seen_point && !seen_exp) {
      canonical.push_back('.');
      seen_point = true;
      continue;
    }
    if (group_sep != '\0' && c == group_sep && !seen_point && !seen_exp) {
      // Group separators are only valid in the integer part and are
      // discarded by the parser. `group_sep == '\0'` means the caller
      // opted out of group separators entirely (see `NumberValue_` below).
      continue;
    }
    if ((c == 'e' || c == 'E') && seen_digit && !seen_exp) {
      canonical.push_back('e');
      seen_exp = true;
      if (i + 1 < s.size() && (s[i + 1] == '+' || s[i + 1] == '-')) {
        canonical.push_back(s[i + 1]);
        ++i;
      }
      continue;
    }
    return false;
  }
  if (!seen_digit) {
    return false;
  }
  // Parse via std::strtod over a NUL-terminated buffer.
  char stack_buf[64];
  char* heap_buf = nullptr;
  const std::size_t n = canonical.size();
  char* buf = stack_buf;
  if (n + 1 > sizeof(stack_buf)) {
    heap_buf = static_cast<char*>(std::malloc(n + 1));
    if (heap_buf == nullptr) {
      return false;
    }
    buf = heap_buf;
  }
  std::memcpy(buf, canonical.data(), n);
  buf[n] = '\0';
  char* end_ptr = nullptr;
  double parsed = std::strtod(buf, &end_ptr);
  const bool ok = end_ptr == buf + n;
  if (heap_buf != nullptr) {
    std::free(heap_buf);
  }
  if (!ok) {
    return false;
  }
  if (std::isnan(parsed) || std::isinf(parsed)) {
    return false;
  }
  for (int k = 0; k < percent_count; ++k) {
    parsed *= 0.01;
  }
  if (negative) {
    parsed = -parsed;
  }
  *out = parsed;
  return true;
}

// ---------------------------------------------------------------------------
// TEXT(value, format_text)
// ---------------------------------------------------------------------------

Value Text_(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  const Value& v = args[0];

  // Error and non-scalar inputs short-circuit before we even look at the
  // format string: errors propagate, arrays/refs/lambdas are #VALUE!.
  if (v.is_error()) {
    return v;
  }
  if (v.kind() == ValueKind::Array || v.kind() == ValueKind::Ref || v.kind() == ValueKind::Lambda) {
    return Value::error(ErrorCode::Value);
  }

  // Mac Excel ja-JP returns the uppercase boolean text and ignores the
  // format string entirely for a bool value. This matches the observable
  // oracle and the documented Excel contract for TEXT(TRUE, ...) /
  // TEXT(FALSE, ...).
  if (v.is_boolean()) {
    return Value::text(v.as_boolean() ? std::string_view{"TRUE"} : std::string_view{"FALSE"});
  }

  auto fmt = coerce_to_text(args[1]);
  if (!fmt) {
    return Value::error(fmt.error());
  }
  const std::string& format_text = fmt.value();
  if (format_text.empty()) {
    return Value::text({});
  }

  // Non-coercible text values are rejected: Mac Excel ja-JP returns
  // #VALUE! for `TEXT("abc", ...)` regardless of the format. Numeric
  // strings ("42", " $1,234 ") still succeed via `parse_numeric` below.
  double number = 0.0;
  if (v.is_text()) {
    const std::string_view raw = v.as_text();
    double parsed = 0.0;
    if (!parse_numeric(raw, '.', ',', &parsed)) {
      return Value::error(ErrorCode::Value);
    }
    number = parsed;
  } else if (v.is_number()) {
    number = v.as_number();
  } else {
    // Blank -> 0; any other kind would have been caught above.
    number = 0.0;
  }

  if (std::isnan(number) || std::isinf(number)) {
    return Value::error(ErrorCode::Num);
  }

  std::string out;
  out.reserve(32);
  const auto status = text_format::apply_format(number, format_text, out);
  if (status != text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

// ---------------------------------------------------------------------------
// VALUE(text)
// ---------------------------------------------------------------------------

Value Value_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  switch (v.kind()) {
    case ValueKind::Number:
      return v;
    case ValueKind::Bool:
      // Excel's VALUE deliberately rejects boolean inputs, even though
      // they coerce to 1/0 in arithmetic contexts.
      return Value::error(ErrorCode::Value);
    case ValueKind::Error:
      return v;
    case ValueKind::Blank: {
      // VALUE("") returns 0 in Excel; a truly blank cell coerces to ""
      // first and then to 0.
      return Value::number(0.0);
    }
    case ValueKind::Text: {
      const std::string_view raw = v.as_text();
      // Phase 1: numeric parse (the common case).
      double numeric = 0.0;
      if (parse_numeric(raw, '.', ',', &numeric)) {
        return Value::number(numeric);
      }
      // Phase 2: date / time parse. Leading whitespace is trimmed (the
      // date/time parser rejects leading U+3000, so we only strip ASCII).
      const std::string_view trimmed = date_parse::trim_date_text(raw);
      if (!trimmed.empty()) {
        double date_serial = 0.0;
        double time_frac = 0.0;
        bool has_date = false;
        bool has_time = false;
        if (date_parse::parse_date_time_text(trimmed, &date_serial, &time_frac, &has_date, &has_time)) {
          return Value::number(date_serial + time_frac);
        }
      }
      return Value::error(ErrorCode::Value);
    }
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// ---------------------------------------------------------------------------
// NUMBERVALUE(text, [decimal_sep], [group_sep])
// ---------------------------------------------------------------------------

Value NumberValue_(const Value* args, std::uint32_t arity, Arena& /*arena*/) {
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  char decimal_sep = '.';
  char group_sep = ',';
  // Track whether the caller supplied an explicit group separator; when
  // they only passed `decimal_sep`, we silently disable grouping so the
  // 2-arity call `NUMBERVALUE("3,14", ",")` cannot collide with the
  // en-US default group sep of `,`.
  bool group_sep_supplied = false;
  if (arity >= 2) {
    auto dsep = coerce_to_text(args[1]);
    if (!dsep) {
      return Value::error(dsep.error());
    }
    if (dsep.value().empty()) {
      return Value::error(ErrorCode::Value);
    }
    decimal_sep = dsep.value().front();
  }
  if (arity >= 3) {
    auto gsep = coerce_to_text(args[2]);
    if (!gsep) {
      return Value::error(gsep.error());
    }
    if (gsep.value().empty()) {
      return Value::error(ErrorCode::Value);
    }
    group_sep = gsep.value().front();
    group_sep_supplied = true;
  }
  // Only the explicit 3-arg form can produce an identical-separator error.
  // When `group_sep` is the implicit default that happens to collide with
  // the user's `decimal_sep`, disable grouping instead of erroring.
  if (group_sep_supplied && decimal_sep == group_sep) {
    return Value::error(ErrorCode::Value);
  }
  if (!group_sep_supplied && decimal_sep == group_sep) {
    group_sep = '\0';
  }
  double parsed = 0.0;
  if (parse_numeric(text.value(), decimal_sep, group_sep, &parsed)) {
    return Value::number(parsed);
  }
  // Mac Excel ja-JP NUMBERVALUE accepts date / time strings in addition
  // to the numeric grammar documented by Microsoft. Fall through to the
  // shared date-parse helper when the numeric path fails.
  const std::string_view trimmed = date_parse::trim_date_text(text.value());
  if (!trimmed.empty()) {
    double date_serial = 0.0;
    double time_frac = 0.0;
    bool has_date = false;
    bool has_time = false;
    if (date_parse::parse_date_time_text(trimmed, &date_serial, &time_frac, &has_date, &has_time)) {
      return Value::number(date_serial + time_frac);
    }
  }
  return Value::error(ErrorCode::Value);
}

}  // namespace

void register_text_format_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"TEXT", 2u, 2u, &Text_});
  registry.register_function(FunctionDef{"VALUE", 1u, 1u, &Value_});
  registry.register_function(FunctionDef{"NUMBERVALUE", 1u, 3u, &NumberValue_});
}

}  // namespace eval
}  // namespace formulon
