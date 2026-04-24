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
// FIXED(number, [decimals=2], [no_commas=FALSE])
// ---------------------------------------------------------------------------
//
// Rounds `number` to `decimals` places and renders it with thousands group
// separators unless `no_commas` is truthy. `decimals` is truncated toward
// zero; Excel caps the decimals parameter at 127 (values outside [-127, 127]
// surface `#VALUE!`). Negative `decimals` rounds left of the decimal point
// (e.g. `FIXED(1234.56, -2) = "1,200"`). The actual rounding at negative
// decimals is done manually before formatting because `apply_format`'s
// numeric walker does not support left-of-decimal-point rounding.

Expected<int, ErrorCode> fixed_read_int(const Value& v) {
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

Value Fixed_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num = coerce_to_number(args[0]);
  if (!num) {
    return Value::error(num.error());
  }
  if (std::isnan(num.value()) || std::isinf(num.value())) {
    return Value::error(ErrorCode::Num);
  }
  int decimals = 2;
  if (arity >= 2) {
    auto parsed = fixed_read_int(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    decimals = parsed.value();
  }
  if (decimals > 127 || decimals < -127) {
    return Value::error(ErrorCode::Value);
  }
  bool no_commas = false;
  if (arity >= 3) {
    auto parsed = coerce_to_bool(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    no_commas = parsed.value();
  }
  // Apply negative-decimals rounding manually: round to the nearest
  // multiple of 10^|decimals| with the same half-away-from-zero convention
  // as `std::round`.
  double value = num.value();
  if (decimals < 0) {
    const double scale = std::pow(10.0, -decimals);
    value = std::round(value / scale) * scale;
  }
  const int effective_decimals = decimals < 0 ? 0 : decimals;
  std::string fmt;
  fmt.reserve(16 + static_cast<std::size_t>(effective_decimals));
  fmt.append(no_commas ? "0" : "#,##0");
  if (effective_decimals > 0) {
    fmt.push_back('.');
    fmt.append(static_cast<std::size_t>(effective_decimals), '0');
  }
  std::string out;
  out.reserve(32);
  const auto status = text_format::apply_format(value, fmt, out);
  if (status != text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

// ---------------------------------------------------------------------------
// DOLLAR(number, [decimals=2])
// ---------------------------------------------------------------------------
//
// Mac Excel ja-JP formats with the yen sign `¥` (UTF-8 0xC2 0xA5) rather
// than the dollar sign: positive values render as `¥1,234.56`, negative
// values as `(¥1,234.56)` (parentheses, no leading minus). Uses a two-
// section format `¥#,##0.00;(¥#,##0.00)`; the format engine's section
// selector emits `std::fabs(value)` for section 1 so the negative branch
// already strips the minus sign. Negative `decimals` rounds left of the
// decimal point (same rule as FIXED); `|decimals| > 127` -> `#VALUE!`.

Value Dollar_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num = coerce_to_number(args[0]);
  if (!num) {
    return Value::error(num.error());
  }
  if (std::isnan(num.value()) || std::isinf(num.value())) {
    return Value::error(ErrorCode::Num);
  }
  int decimals = 2;
  if (arity >= 2) {
    auto parsed = fixed_read_int(args[1]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    decimals = parsed.value();
  }
  if (decimals > 127 || decimals < -127) {
    return Value::error(ErrorCode::Value);
  }
  double value = num.value();
  if (decimals < 0) {
    const double scale = std::pow(10.0, -decimals);
    value = std::round(value / scale) * scale;
  }
  const int effective_decimals = decimals < 0 ? 0 : decimals;
  // Two-section ¥ format: positive uses `¥#,##0[.00]`, negative uses
  // `(¥#,##0[.00])`. The format engine passes `std::fabs(value)` into
  // section 1, so the trailing `)` lands after the formatted digits.
  std::string fraction;
  if (effective_decimals > 0) {
    fraction.reserve(1u + static_cast<std::size_t>(effective_decimals));
    fraction.push_back('.');
    fraction.append(static_cast<std::size_t>(effective_decimals), '0');
  }
  std::string fmt;
  fmt.reserve(32 + 2u * fraction.size());
  // Positive section: "¥#,##0[.00]"
  fmt.append("\xC2\xA5#,##0");
  fmt.append(fraction);
  fmt.push_back(';');
  // Negative section: "(¥#,##0[.00])"
  fmt.push_back('(');
  fmt.append("\xC2\xA5#,##0");
  fmt.append(fraction);
  fmt.push_back(')');
  std::string out;
  out.reserve(32);
  const auto status = text_format::apply_format(value, fmt, out);
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

// VALUETOTEXT(value, [format])
//
// Converts `value` to text, exactly as Excel 365 does when the user types
// `=VALUETOTEXT(x)` into a cell. The `format` second argument is 0
// ("concise", the default) or 1 ("strict").
//
//   concise:
//     * Numbers → General format (same as `coerce_to_text`).
//     * Bools   → "TRUE" / "FALSE".
//     * Text    → unchanged, no quoting.
//     * Blank   → "".
//   strict:
//     * Text    → wrapped in double-quotes; embedded `"` become `""`.
//     * Booleans, numbers, blanks → same as concise.
//
// Errors are NOT suppressed — they propagate as the function's result
// (matching Excel's behaviour where `VALUETOTEXT(#DIV/0!)` returns
// `#DIV/0!`, not the text "#DIV/0!").
Value ValueToText_(const Value* args, std::uint32_t arity, Arena& arena) {
  const Value& v = args[0];
  if (v.is_error()) {
    return v;
  }
  bool strict = false;
  if (arity >= 2) {
    const Value& fmt = args[1];
    if (fmt.is_error()) {
      return fmt;
    }
    auto n = coerce_to_number(fmt);
    if (!n) {
      return Value::error(n.error());
    }
    const double nv = n.value();
    if (nv == 0.0) {
      strict = false;
    } else if (nv == 1.0) {
      strict = true;
    } else {
      return Value::error(ErrorCode::Value);
    }
  }
  if (strict && v.is_text()) {
    const std::string_view src = v.as_text();
    std::string out;
    out.reserve(src.size() + 2);
    out.push_back('"');
    for (char c : src) {
      if (c == '"') {
        out.push_back('"');
      }
      out.push_back(c);
    }
    out.push_back('"');
    return Value::text(arena.intern(out));
  }
  auto text = coerce_to_text(v);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(text.value()));
}

// ARRAYTOTEXT(array, [format]) — for a scalar input this is a thin
// alias for VALUETOTEXT. Arrays are not fully modelled yet; when an
// array AST reaches here it has already been reduced to its first
// element by the eager dispatcher, so the behavioural difference only
// shows up in the as-yet-unimplemented spill pipeline.
Value ArrayToText_(const Value* args, std::uint32_t arity, Arena& arena) {
  return ValueToText_(args, arity, arena);
}

}  // namespace

void register_text_format_builtins(FunctionRegistry& registry) {
  registry.register_function(FunctionDef{"TEXT", 2u, 2u, &Text_});
  registry.register_function(FunctionDef{"VALUE", 1u, 1u, &Value_});
  registry.register_function(FunctionDef{"VALUETOTEXT", 1u, 2u, &ValueToText_});
  registry.register_function(FunctionDef{"ARRAYTOTEXT", 1u, 2u, &ArrayToText_});
  registry.register_function(FunctionDef{"NUMBERVALUE", 1u, 3u, &NumberValue_});
  registry.register_function(FunctionDef{"FIXED", 1u, 3u, &Fixed_});
  registry.register_function(FunctionDef{"DOLLAR", 1u, 2u, &Dollar_});
}

}  // namespace eval
}  // namespace formulon
