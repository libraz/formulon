//
// Implementation of Formulon's text-conversion builtins: TEXT, VALUE, and
// NUMBERVALUE. All three mediate between numeric and textual
// representations, so they share the format-string engine in
// `eval/text_format/number_format.h` (TEXT) and the date/time parser in
// `eval/date_text_parse.h` (VALUE).

#include "eval/builtins/text_format.h"

#include <algorithm>
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
#include "eval/eval_profile_scope.h"
#include "eval/function_registry.h"
#include "eval/locale_text.h"
#include "eval/number_parse.h"
#include "eval/shape_ops_lazy.h"
#include "eval/text_format/number_format.h"
#include "eval/text_format/rounding.h"
#include "excel_locale.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace {

Value apply_text_format_text(std::string_view value, std::string_view format_text, Arena& arena) {
  std::string out;
  const auto status = ::formulon::text_format::apply_text_format(value, format_text, out);
  if (status != ::formulon::text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

}  // namespace

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

  auto fmt = coerce_to_text(args[1]);
  if (!fmt) {
    return Value::error(fmt.error());
  }
  const std::string& format_text = fmt.value();

  // Booleans use the text-section path. That preserves their visible TRUE /
  // FALSE spelling while still applying a text placeholder and validating
  // malformed sections in the format.
  if (v.is_boolean()) {
    return apply_text_format_text(locale_bool_text(v.as_boolean()), format_text, arena);
  }

  // Text goes through the shared numeric-coercion ladder, the same one
  // arithmetic and FLOOR use, so anything `=A1+0` accepts TEXT formats as
  // well: numeric strings ("42", " $1,234 "), percent, currency, and
  // date / time text. A value the ladder rejects (for example "abc") is
  // rendered through the format's text section instead.
  double number = 0.0;
  if (v.is_text()) {
    bool from_date_text = false;
    auto coerced = coerce_text_to_number(v.as_text(), &from_date_text);
    if (!coerced) {
      if (coerced.error() != ErrorCode::Value) {
        return Value::error(coerced.error());
      }
      return apply_text_format_text(v.as_text(), format_text, arena);
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
// separators unless `no_commas` is truthy. `decimals` snaps to a near
// integer, then truncates toward zero; Excel caps the decimals parameter at 127 (values outside [-127, 127]
// surface `#VALUE!`). Negative `decimals` rounds left of the decimal point
// (e.g. `FIXED(1234.56, -2) = "1,200"`). The actual rounding at negative
// decimals is done manually before formatting because `apply_format`'s
// numeric walker does not support left-of-decimal-point rounding.

Expected<int, ErrorCode> fixed_read_int(const Value& v) {
  auto d = coerce_to_number(v);
  if (!d) {
    return std::move(d.error());
  }
  const double truncated = std::trunc(snap_near_integer(d.value()));
  if (truncated < -127.0 || truncated > 127.0) {
    return ErrorCode::Value;
  }
  return static_cast<int>(truncated);
}

Expected<int, ErrorCode> read_optional_fixed_decimals(const Value* args, std::uint32_t arity, std::uint32_t index,
                                                      int default_value) {
  int decimals = default_value;
  if (arity >= index + 1u) {
    auto parsed = fixed_read_int(args[index]);
    if (!parsed) {
      return std::move(parsed.error());
    }
    decimals = parsed.value();
  }
  return decimals;
}

Value apply_text_number_format(double value, std::string_view format, Arena& arena) {
  std::string out;
  out.reserve(32);
  // FIXED / DOLLAR build their format in the invariant stored syntax.
  const auto status = ::formulon::text_format::apply_format(value, format, out, /*date1904=*/false,
                                                            ::formulon::text_format::FormatDialect::kStored);
  if (status != ::formulon::text_format::FormatStatus::kOk) {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(out));
}

std::string build_decimal_body(int decimals, bool no_commas) {
  const int effective_decimals = decimals < 0 ? 0 : decimals;
  std::string body;
  body.reserve(8u + static_cast<std::size_t>(effective_decimals));
  body.append(no_commas ? "0" : "#,##0");
  if (effective_decimals > 0) {
    body.push_back('.');
    body.append(static_cast<std::size_t>(effective_decimals), '0');
  }
  return body;
}

Value Fixed_(const Value* args, std::uint32_t arity, Arena& arena) {
  auto num = coerce_to_number(args[0]);
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
  const std::string fmt = build_decimal_body(decimals, no_commas);
  return apply_text_number_format(value, fmt, arena);
}

// ---------------------------------------------------------------------------
// DOLLAR(number, [decimals]) / USDOLLAR(number, [decimals])
// ---------------------------------------------------------------------------
//
// Both render `number` as currency text through one kernel; they differ in
// the symbol, the default `decimals` and the negative section (measured on
// Mac Excel 365 ja-JP):
//   DOLLAR   `¥1,235`,   `¥-1,235`     (locale currency, default 0 decimals)
//   USDOLLAR `$1,234.50`, `($1,234.50)` (US dollars, default 2)
// Symbol placement and the negative form follow the locale `Currency` fact.
// The chosen section formats the magnitude, so its literal `-` or
// parentheses carry the sign. Negative `decimals` rounds left of the decimal
// point (same rule as FIXED); `|decimals| > 127` -> `#VALUE!`.

Currency locale_dollar_style() {
  return locale_facts(current_eval_profile()).currency;
}

// USDOLLAR formats in US dollars where the locale says so; elsewhere it is DOLLAR itself.
Currency locale_usdollar_style() {
  const std::string_view symbol = locale_facts(current_eval_profile()).usdollar_symbol;
  if (symbol.empty()) {
    return locale_dollar_style();
  }
  return Currency{symbol, false, false, true, false, true, 2U};
}

Value format_currency(const Value* args, std::uint32_t arity, Arena& arena, const Currency& currency) {
  auto num = coerce_to_number(args[0]);
  if (!num) {
    return Value::error(num.error());
  }
  auto decimals_e = read_optional_fixed_decimals(args, arity, 1, static_cast<int>(currency.default_decimals));
  if (!decimals_e) {
    return Value::error(decimals_e.error());
  }
  const int decimals = decimals_e.value();
  const double value = ::formulon::text_format::round_display_decimal(num.value(), decimals);
  const std::string body = build_decimal_body(decimals, /*no_commas=*/false);
  // A number that display-rounds to zero keeps its sign, one rounded left of
  // the decimal point first does not: USDOLLAR(-0.4,0) is `($0)` but
  // USDOLLAR(-4,-1) is `$0`. So the section is picked here, and the chosen
  // one formats the magnitude.
  const bool negative = decimals < 0 ? value < 0.0 : num.value() < 0.0;
  std::string fmt;
  fmt.reserve(body.size() + currency.symbol.size() + 6u);
  const bool parens = negative && currency.negative_parens;
  if (parens) {
    fmt.push_back('(');
  }
  if (negative && !parens && !currency.minus_after_symbol) {
    fmt.push_back('-');
  }
  // Quoted, so a lettered symbol (`US$`) stays text rather than date codes.
  const std::string symbol = "\"" + std::string(currency.symbol) + "\"";
  if (!currency.suffix) {
    fmt.append(symbol);
    if (currency.space) {
      fmt.push_back(' ');
    }
    if (negative && !parens && currency.minus_after_symbol) {
      fmt.push_back('-');
    }
  }
  fmt.append(body);
  if (currency.suffix) {
    if (currency.space) {
      fmt.push_back(' ');
    }
    fmt.append(symbol);
  }
  if (parens) {
    fmt.push_back(')');
  }
  return apply_text_number_format(std::fabs(value), fmt, arena);
}

Value Dollar_(const Value* args, std::uint32_t arity, Arena& arena) {
  return format_currency(args, arity, arena, locale_dollar_style());
}

Value UsDollar_(const Value* args, std::uint32_t arity, Arena& arena) {
  return format_currency(args, arity, arena, locale_usdollar_style());
}

// ---------------------------------------------------------------------------
// VALUE(text)
// ---------------------------------------------------------------------------

// `current_year` supplies the fallback year for a year-less date_text
// ("3/15", "3月15日"), matching Microsoft's documented DATEVALUE rule (which
// VALUE's date-parse fallback shares). `0` disables that form, matching the
// pre-existing year-required behaviour.
Value Value_(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/, int current_year) {
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
      const LocaleFacts& facts = locale_facts(current_eval_profile());
      if (parse_numeric(normalized, facts.decimal_separator, facts.group_separator, &numeric)) {
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
        if (date_parse::parse_date_time_text(trimmed, &date_serial, &time_frac, &has_date, &has_time, current_year)) {
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
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  char decimal_sep = facts.decimal_separator;
  char group_sep = facts.group_separator;
  // Track whether the caller supplied an explicit group separator; when
  // they only passed `decimal_sep`, we silently disable grouping so the
  // 2-arity call `NUMBERVALUE("3,14", ",")` cannot collide with the
  // locale's default group sep.
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
  std::string normalized = normalize_locale_numeric(text.value(), &paren_negated);
  // Spaces are dropped wherever they sit (value_numbervalue.numbervalue_inner_spaces).
  normalized.erase(std::remove(normalized.begin(), normalized.end(), ' '), normalized.end());
  double parsed = 0.0;
  if (parse_numeric(normalized, decimal_sep, group_sep, &parsed, /*strict_groups=*/false)) {
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

Expected<bool, ErrorCode> decode_text_format(const Value& evaluated_value) {
  if (evaluated_value.is_error()) {
    return evaluated_value.as_error();
  }
  auto number = coerce_to_number(evaluated_value);
  if (!number) {
    return std::move(number.error());
  }
  if (number.value() == 0.0) {
    return false;
  }
  if (number.value() == 1.0) {
    return true;
  }
  return ErrorCode::Value;
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
    auto decoded = decode_text_format(args[1]);
    if (!decoded) {
      return Value::error(decoded.error());
    }
    strict = decoded.value();
  }
  if (strict && v.is_text()) {
    const std::string_view src = v.as_text();
    std::string out;
    out.reserve(src.size() + 2);
    append_quoted_text(src, out);
    return Value::text(arena.intern(out));
  }
  auto text = coerce_to_text(v);
  if (!text) {
    return Value::error(text.error());
  }
  return Value::text(arena.intern(text.value()));
}

bool append_arraytotext_cell(const Value& v, bool strict, std::string& out, ErrorCode* error) {
  if (v.is_error()) {
    out.append(locale_error_text(v.as_error()));
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
  auto decoded = decode_text_format(fmt);
  if (!decoded) {
    return Value::error(decoded.error());
  }
  *strict = decoded.value();
  return Value::blank();
}

Value arraytotext_from_array(const ArrayValue& arr, bool strict, Arena& arena) {
  const LocaleFacts& facts = locale_facts(current_eval_profile());
  std::string out;
  if (strict) {
    out.push_back('{');
  }
  ErrorCode error = ErrorCode::Value;
  for (std::uint32_t r = 0; r < arr.rows; ++r) {
    for (std::uint32_t c = 0; c < arr.cols; ++c) {
      if (r != 0U || c != 0U) {
        if (strict) {
          out.push_back(c == 0U ? facts.array_row_separator : facts.array_column_separator);
        } else {
          out.push_back(facts.list_separator);
          out.push_back(' ');
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
  const LocaleFacts& facts = locale_facts(current_eval_profile());
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
          out.push_back(c == 0U ? facts.array_row_separator : facts.array_column_separator);
        } else {
          out.push_back(facts.list_separator);
          out.push_back(' ');
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
    auto decoded = decode_text_format(args[1]);
    if (!decoded) {
      return Value::error(decoded.error());
    }
    strict = decoded.value();
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

// ---------------------------------------------------------------------------
// VALUE(text). Exposed (not in the anonymous namespace) because it is
// clock-sensitive (a year-less date_text fallback reads the wall clock) and
// served through the shared calendar lookup (`find_date_entry`) + the lazy
// VALUE wrapper, not the eager registry. Mirrors `text_builtin_impl` above.
// ---------------------------------------------------------------------------

Value value_builtin_impl(const Value* args, std::uint32_t arity, Arena& arena, bool /*date1904*/,
                         const date_time::CivilTime& now) {
  return Value_(args, arity, arena, now.date.y);
}

// Host-clock fallback for VALUE, registered as `DateEntry::impl` so a
// contextless caller still resolves a year-less date_text. Mirrors
// `DatevalueHostClock_` in `builtins/datetime.cpp`.
Value value_builtin_host_clock_impl(const Value* args, std::uint32_t arity, Arena& arena, bool date1904) {
  return value_builtin_impl(args, arity, arena, date1904, date_time::host_civil_time());
}

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
  // TEXT and VALUE are NOT registered here: both are clock-sensitive (TEXT's
  // date format codes read the workbook epoch; VALUE's year-less date_text
  // fallback reads the wall clock), so they route through the lazy TEXT/
  // VALUE wrappers (`eval_text_lazy`/`eval_datetime_lazy`) and the shared
  // `find_date_entry` hook (VM), which pass `EvalContext::date1904()` /
  // `EvalContext::wall_clock()` into `text_builtin_impl` / `value_builtin_impl`.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"VALUETOTEXT", 1u, 2u, &ValueToText_}, {"ARRAYTOTEXT", 1u, 2u, &ArrayToText_},
      {"NUMBERVALUE", 1u, 3u, &NumberValue_}, {"FIXED", 1u, 3u, &Fixed_},
      {"DOLLAR", 1u, 2u, &Dollar_},           {"USDOLLAR", 1u, 2u, &UsDollar_},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
