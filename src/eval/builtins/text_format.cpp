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

#include "eval/builtins/registration_helpers.h"
#include "eval/coerce.h"
#include "eval/date_text_parse.h"
#include "eval/date_time.h"
#include "eval/function_registry.h"
#include "eval/number_parse.h"
#include "eval/shape_ops_lazy.h"
#include "eval/text_format/number_format.h"
#include "eval/text_format/rounding.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

// ---------------------------------------------------------------------------
// TEXT(value, format_text). Exposed (not in the anonymous namespace) because
// it is date1904-sensitive and served through the shared calendar lookup
// (`find_date_entry`) + the lazy TEXT wrapper, not the eager registry.
// ---------------------------------------------------------------------------

Value text_builtin_impl(const Value* args, std::uint32_t /*arity*/, Arena& arena, bool date1904) {
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

  // Text goes through the shared numeric-coercion ladder, the same one
  // arithmetic and FLOOR use, so anything `=A1+0` accepts TEXT formats as
  // well: numeric strings ("42", " $1,234 "), percent, currency, and
  // date / time text. Only a value the ladder itself rejects ("abc") is
  // #VALUE!.
  double number = 0.0;
  if (v.is_text()) {
    bool from_date_text = false;
    auto coerced = coerce_text_to_number(v.as_text(), &from_date_text);
    if (!coerced) {
      return Value::error(coerced.error());
    }
    number = coerced.value();
    // The ladder's date fallback always yields a 1900-system serial. The
    // format codes below read the workbook epoch, so shift the serial into
    // it or a 1904 workbook would render the wrong calendar day.
    if (from_date_text && date1904) {
      number -= date_time::kDate1904EpochGap;
    }
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
  const auto status = ::formulon::text_format::apply_format(number, format_text, out, date1904);
  if (status != ::formulon::text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

namespace {

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
  auto d = coerce_to_number(v);
  if (!d) {
    return d.error();
  }
  if (std::isnan(d.value()) || std::isinf(d.value())) {
    return ErrorCode::Num;
  }
  return static_cast<int>(std::trunc(d.value()));
}

Expected<double, ErrorCode> read_finite_number_arg(const Value* args, std::uint32_t index) {
  auto number = coerce_to_number(args[index]);
  if (!number) {
    return number.error();
  }
  const double d = number.value();
  if (std::isnan(d) || std::isinf(d)) {
    return ErrorCode::Num;
  }
  return d;
}

Expected<int, ErrorCode> read_optional_fixed_decimals(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                      int default_value) {
  int decimals = default_value;
  if (arity >= index + 1u) {
    auto parsed = fixed_read_int(args[index]);
    if (!parsed) {
      return parsed.error();
    }
    decimals = parsed.value();
  }
  if (decimals > 127 || decimals < -127) {
    return ErrorCode::Value;
  }
  return decimals;
}

Value apply_text_number_format(double value, std::string_view format, Arena& arena) {
  std::string out;
  out.reserve(32);
  const auto status = ::formulon::text_format::apply_format(value, format, out);
  if (status != ::formulon::text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

Value Fixed_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num = read_finite_number_arg(args, 0);
  if (!num) {
    return Value::error(num.error());
  }
  auto decimals_e = read_optional_fixed_decimals(args, arity, 1, 2);
  if (!decimals_e) {
    return Value::error(decimals_e.error());
  }
  const int decimals = decimals_e.value();
  bool no_commas = false;
  if (arity >= 3) {
    auto parsed = coerce_to_bool(args[2]);
    if (!parsed) {
      return Value::error(parsed.error());
    }
    no_commas = parsed.value();
  }
  const double value = ::formulon::text_format::round_display_decimal(num.value(), decimals);
  const int effective_decimals = decimals < 0 ? 0 : decimals;
  std::string fmt;
  fmt.reserve(16 + static_cast<std::size_t>(effective_decimals));
  fmt.append(no_commas ? "0" : "#,##0");
  if (effective_decimals > 0) {
    fmt.push_back('.');
    fmt.append(static_cast<std::size_t>(effective_decimals), '0');
  }
  return apply_text_number_format(value, fmt, arena);
}

// ---------------------------------------------------------------------------
// DOLLAR(number, [decimals])
// ---------------------------------------------------------------------------
//
// Mac Excel ja-JP formats with the yen sign `¥` (UTF-8 0xC2 0xA5) rather
// than the dollar sign. The default `decimals` is locale-dependent: ja-JP
// uses 0 (no fractional part) while en-US uses 2. The negative-number
// section also diverges from en-US: ja-JP renders negatives as
// `¥-1,235` (leading minus inside the prefix), not `($1,234.56)` parens.
// Uses a two-section format `¥#,##0[.00];¥-#,##0[.00]`; the format
// engine's section selector emits `std::fabs(value)` for section 1, so
// the literal `-` in the negative section is what produces the sign.
// Negative `decimals` rounds left of the decimal point (same rule as
// FIXED); `|decimals| > 127` -> `#VALUE!`.

Value Dollar_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num = read_finite_number_arg(args, 0);
  if (!num) {
    return Value::error(num.error());
  }
  // ja-JP default is 0 decimals (en-US would default to 2).
  auto decimals_e = read_optional_fixed_decimals(args, arity, 1, 0);
  if (!decimals_e) {
    return Value::error(decimals_e.error());
  }
  const int decimals = decimals_e.value();
  const double value = ::formulon::text_format::round_display_decimal(num.value(), decimals);
  const int effective_decimals = decimals < 0 ? 0 : decimals;
  // Two-section ¥ format: positive uses `¥#,##0[.00]`, negative uses
  // `¥-#,##0[.00]`. The format engine passes `std::fabs(value)` into
  // section 1, so the literal `-` in the negative section emits the sign.
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
  // Negative section: "¥-#,##0[.00]"
  fmt.append("\xC2\xA5-#,##0");
  fmt.append(fraction);
  return apply_text_number_format(value, fmt, arena);
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
      // Phase 1: numeric parse with a locale pre-pass that folds
      // full-width digits/punctuation to ASCII and strips accounting-
      // style outer parentheses (`"(1234)"` -> -1234).
      bool paren_negated = false;
      const std::string normalized = normalize_locale_numeric(raw, &paren_negated);
      double numeric = 0.0;
      if (parse_numeric(normalized, '.', ',', &numeric)) {
        return Value::number(paren_negated ? -numeric : numeric);
      }
      // Phase 2: date / time parse. Leading whitespace is trimmed (the
      // date/time parser rejects leading U+3000, so we only strip ASCII).
      // The original (non-normalised) text is used: the date parser has
      // its own ja-JP / kanji handling and must not see a folded form.
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
  // Empty input (including a Blank cell that `coerce_to_text` flattened
  // to `""`) returns 0, matching Mac Excel ja-JP. This mirrors VALUE's
  // existing empty-string-is-zero handling.
  if (text.value().empty()) {
    return Value::number(0.0);
  }
  // Locale pre-pass: fold full-width digits/punctuation to ASCII and
  // detect accounting-style outer parentheses. Mac Excel accepts
  // `NUMBERVALUE("(1234)", ".")` as -1234 contrary to the original
  // assumption documented in `tests/divergence.yaml`.
  bool paren_negated = false;
  const std::string normalized = normalize_locale_numeric(text.value(), &paren_negated);
  double parsed = 0.0;
  if (parse_numeric(normalized, decimal_sep, group_sep, &parsed)) {
    return Value::number(paren_negated ? -parsed : parsed);
  }
  // Mac Excel ja-JP NUMBERVALUE accepts date / time strings in addition
  // to the numeric grammar documented by Microsoft. Fall through to the
  // shared date-parse helper when the numeric path fails. The original
  // (non-normalised) text is used so the date parser's ja-JP path sees
  // the input verbatim.
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

void append_quoted_text(std::string_view src, std::string& out) {
  out.push_back('"');
  for (char c : src) {
    if (c == '"') {
      out.push_back('"');
    }
    out.push_back(c);
  }
  out.push_back('"');
}

bool append_arraytotext_cell(const Value& v, bool strict, std::string& out, ErrorCode* error) {
  if (v.is_error()) {
    out.append(display_name(v.as_error()));
    return true;
  }
  if (strict && v.is_text()) {
    append_quoted_text(v.as_text(), out);
    return true;
  }
  auto text = coerce_to_text(v);
  if (!text) {
    *error = text.error();
    return false;
  }
  out.append(text.value());
  return true;
}

Value parse_arraytotext_format(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                               const EvalContext& ctx, bool* strict) {
  *strict = false;
  if (call.as_call_arity() < 2U) {
    return Value::blank();
  }
  const Value fmt = eval_node(call.as_call_arg(1), arena, registry, ctx);
  if (fmt.is_error()) {
    return fmt;
  }
  auto n = coerce_to_number(fmt);
  if (!n) {
    return Value::error(n.error());
  }
  if (n.value() == 0.0) {
    *strict = false;
    return Value::blank();
  }
  if (n.value() == 1.0) {
    *strict = true;
    return Value::blank();
  }
  return Value::error(ErrorCode::Value);
}

Value arraytotext_from_array(const ArrayValue& arr, bool strict, Arena& arena) {
  std::string out;
  if (strict) {
    out.push_back('{');
  }
  ErrorCode error = ErrorCode::Value;
  for (std::uint32_t r = 0; r < arr.rows; ++r) {
    for (std::uint32_t c = 0; c < arr.cols; ++c) {
      if (r != 0U || c != 0U) {
        if (strict) {
          out.push_back(c == 0U ? ';' : ',');
        } else {
          out.append(", ");
        }
      }
      const Value& cell = arr.cells[static_cast<std::size_t>(r) * arr.cols + c];
      if (!append_arraytotext_cell(cell, strict, out, &error)) {
        return Value::error(error);
      }
    }
  }
  if (strict) {
    out.push_back('}');
  }
  return Value::text(arena.intern(out));
}

Value arraytotext_from_array_literal(const parser::AstNode& literal, bool strict, Arena& arena,
                                     const FunctionRegistry& registry, const EvalContext& ctx) {
  std::string out;
  if (strict) {
    out.push_back('{');
  }
  ErrorCode error = ErrorCode::Value;
  const std::uint32_t rows = literal.as_array_rows();
  const std::uint32_t cols = literal.as_array_cols();
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      if (r != 0U || c != 0U) {
        if (strict) {
          out.push_back(c == 0U ? ';' : ',');
        } else {
          out.append(", ");
        }
      }
      const Value cell = eval_node(literal.as_array_element(r, c), arena, registry, ctx);
      if (cell.is_array()) {
        return Value::error(ErrorCode::Value);
      }
      if (!append_arraytotext_cell(cell, strict, out, &error)) {
        return Value::error(error);
      }
    }
  }
  if (strict) {
    out.push_back('}');
  }
  return Value::text(arena.intern(out));
}

// ARRAYTOTEXT(array, [format]) — scalar inputs mirror VALUETOTEXT except
// that error values are rendered as their display text. Array inputs must
// preserve row/column shape long enough to emit Excel's concise list or
// strict array-literal representation.
Value ArrayToText_(const Value* args, std::uint32_t arity, Arena& arena) {
  bool strict = false;
  if (arity >= 2U) {
    const Value& fmt = args[1];
    if (fmt.is_error()) {
      return fmt;
    }
    auto n = coerce_to_number(fmt);
    if (!n) {
      return Value::error(n.error());
    }
    if (n.value() == 0.0) {
      strict = false;
    } else if (n.value() == 1.0) {
      strict = true;
    } else {
      return Value::error(ErrorCode::Value);
    }
  }
  if (args[0].is_array()) {
    return arraytotext_from_array(*args[0].as_array(), strict, arena);
  }
  if (args[0].is_error()) {
    return args[0];
  }
  std::string out;
  ErrorCode error = ErrorCode::Value;
  if (!append_arraytotext_cell(args[0], strict, out, &error)) {
    return Value::error(error);
  }
  return Value::text(arena.intern(out));
}

}  // namespace

Value eval_arraytotext_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                            const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U || arity > 2U) {
    return Value::error(ErrorCode::Value);
  }
  bool strict = false;
  const Value fmt_status = parse_arraytotext_format(call, arena, registry, ctx, &strict);
  if (fmt_status.is_error()) {
    return fmt_status;
  }
  const parser::AstNode& array_arg = call.as_call_arg(0);
  if (array_arg.kind() == parser::NodeKind::ArrayLiteral) {
    return arraytotext_from_array_literal(array_arg, strict, arena, registry, ctx);
  }
  const Value array_v = eval_node_as_array(array_arg, arena, registry, ctx);
  if (array_v.is_array()) {
    const ArrayValue& arr = *array_v.as_array();
    // A lone 1x1 error argument propagates rather than rendering as the
    // error's display text (eval_node_as_array wraps scalar args).
    if (arr.rows == 1U && arr.cols == 1U && arr.cells[0].is_error()) {
      return arr.cells[0];
    }
    return arraytotext_from_array(arr, strict, arena);
  }
  if (array_v.is_error()) {
    return array_v;
  }
  std::string out;
  ErrorCode error = ErrorCode::Value;
  if (!append_arraytotext_cell(array_v, strict, out, &error)) {
    return Value::error(error);
  }
  return Value::text(arena.intern(out));
}

void register_text_format_builtins(FunctionRegistry& registry) {
  // TEXT is NOT registered here: it is date1904-sensitive (date format codes
  // read the workbook epoch), so it routes through the lazy TEXT wrapper
  // (`eval_text_lazy`) and the shared `find_date_entry` hook (VM), which pass
  // `EvalContext::date1904()` into `text_builtin_impl`.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"VALUE", 1u, 1u, &Value_},
      {"VALUETOTEXT", 1u, 2u, &ValueToText_},
      {"ARRAYTOTEXT", 1u, 2u, &ArrayToText_},
      {"NUMBERVALUE", 1u, 3u, &NumberValue_},
      {"FIXED", 1u, 3u, &Fixed_},
      {"DOLLAR", 1u, 2u, &Dollar_},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
