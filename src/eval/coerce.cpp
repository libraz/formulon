//
// Implementation of the scalar coercion helpers declared in `coerce.h`.

#include "eval/coerce.h"

#include <cmath>
#include <cstdlib>
#include <cstring>
#include <string>
#include <string_view>

#include "eval/date_text_parse.h"
#include "eval/number_parse.h"
#include "utils/double_format.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

// Parses `s` as a full double using std::strtod. Returns true iff the entire
// input (no leftover bytes) parsed cleanly; the numeric value is written
// to `*out` even if it is NaN / Inf, leaving the finiteness check to the
// caller. The stack buffer is sized for any IEEE-754 literal including
// subnormals; longer inputs take the heap path.
bool strtod_full(std::string_view s, double* out) noexcept {
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
  const double parsed = parse_double_c_locale(buf, &end_ptr);
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

Expected<double, ErrorCode> coerce_text_to_number(std::string_view text, bool* from_date_text) {
  if (from_date_text != nullptr) {
    *from_date_text = false;
  }
  // Implementation factored out of the Value-shaped overload's `Text`
  // arm so hot-path callers (criterion parsing) can avoid wrapping a
  // raw string_view in a `Value::text(...)` temporary. The trim / numeric
  // / percent / currency / date fallback ladder is byte-for-byte
  // identical to the original site so external behaviour is unchanged.
  const std::string_view trimmed = strings::trim(text);
  if (trimmed.empty()) {
    // Empty / whitespace-only text is #VALUE! in every numeric-coercion
    // context Mac Excel 365 was tested against (`=""+1`, `=SIN("")`,
    // `=EXP("")`, ... all yield #VALUE!). Blank cells still coerce to
    // 0 via the `ValueKind::Blank` branch in the Value-shaped overload;
    // only the explicit empty string is rejected here.
    return ErrorCode::Value;
  }
  // Layered numeric-coercion fallback, in order:
  //   1. strtod(trimmed)                    - plain numeric fast path
  //   2. trailing '%' stripped, strtod, /100 - percent literals
  //   3. VALUE()-style locale parse          - grouping, parens, full-width,
  //                                            currency ({$, ¥, ￥, €} on one
  //                                            side only), currency + percent
  //   4. date / datetime fallback (raw text) - DATEVALUE-style shapes
  //   5. #VALUE!
  // Currency handling lives entirely in step 3 (`parse_excel_number`) so
  // implicit coercion and the VALUE() builtin share one code path and agree
  // exactly (Mac Excel 365 ja-JP accepts a currency marker on the leading OR
  // trailing side but not both, and only {$, ¥, ￥, €}). The date fallback
  // runs against the raw, untrimmed text so padded date strings stay #VALUE!
  // (see WhitespacePaddedDate rejection test).
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
  // Locale-aware numeric forms that the fast paths above reject: thousands
  // grouping ("1,000"), accounting parentheses ("(100)" -> -100), full-width
  // digits/punctuation, and currency (with the one-side rule above). Excel
  // applies VALUE's normalisation to implicit coercion too, so `="1,000"+1`
  // is 1001 and `COUNTIF(range, ">1,000")` parses the criterion numerically.
  // `parse_excel_number` rejects NaN/Inf internally, so a success is always
  // finite. Runs on the raw (untrimmed) text because the normaliser does its
  // own ASCII/full-width trimming.
  double locale_parsed = 0.0;
  if (parse_excel_number(text, &locale_parsed)) {
    return locale_parsed;
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
  if (date_parse::parse_date_time_text(text, &serial, &frac, &has_date, &has_time)) {
    const double combined = serial + frac;
    if (std::isnan(combined) || std::isinf(combined)) {
      return ErrorCode::Num;
    }
    if (from_date_text != nullptr) {
      *from_date_text = has_date;
    }
    return combined;
  }
  return ErrorCode::Value;
}

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
    case ValueKind::Text:
      // Delegate to the string_view overload so the two entry points
      // share one implementation. Hot-path callers that already have
      // a `string_view` should call `coerce_text_to_number` directly to
      // avoid materialising a `Value::text(...)` temporary.
      return coerce_text_to_number(v.as_text());
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<double, ErrorCode> coerce_to_index_number(const Value& v) {
  auto number = coerce_to_number(v);
  if (!number) {
    return number.error();
  }
  return truncate_index(number.value());
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
  // A zero base with a negative exponent is `1 / 0^|exp|`, and Excel reports
  // it the way it reports any division by zero — `#DIV/0!`, matching the `/`
  // operator and MOD / QUOTIENT rather than the `#NUM!` that std::pow's
  // `+Inf` would otherwise be folded into below. The distinction is
  // load-bearing for `IFERROR` / `ISERR` / `ERROR.TYPE` routing.
  if (base == 0.0 && exp < 0.0) {
    return ErrorCode::Div0;
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

Expected<double, ErrorCode> matrix_strict_number(const Value& v) {
  // Matches `forecast_ets_lazy.cpp::coerce_strict_numeric`: Number passes
  // through raw (NaN / Inf included — matrix-strict callers run their own
  // finite-result guard at the end), Bool coerces to 1/0, Error
  // propagates, everything else (Blank, Text, Array, Ref, Lambda) is
  // `#VALUE!`. The Text rejection is the point of the strict variant.
  switch (v.kind()) {
    case ValueKind::Number:
      return v.as_number();
    case ValueKind::Bool:
      return v.as_boolean() ? 1.0 : 0.0;
    case ValueKind::Error:
      return v.as_error();
    case ValueKind::Blank:
    case ValueKind::Text:
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

Expected<double, ErrorCode> matrix_strict_number_cell(const Value& v, std::uint32_t /*row*/, std::uint32_t /*col*/) {
  // Row / col are accepted today but unused: matrix-strict callers like
  // LINEST / FORECAST.ETS have always reported a single propagating
  // error for the whole matrix. The arguments are reserved so call
  // sites that already track `(row, col)` for iteration can migrate
  // without a churn round when structured-log enrichment lands.
  return matrix_strict_number(v);
}

Expected<std::vector<double>, ErrorCode> collect_numerics(const Value& v, NumericCollectPolicy policy) {
  // Single-pass flatten: an Array iterates its row-major cells; any
  // other kind is treated as a 1-element input. Blank is always
  // dropped (counting blanks as zero is the "A"-family's job, handled
  // by a separate helper). Ref / Lambda are always dropped — callers
  // that care about them resolve refs before calling.
  std::vector<double> out;
  std::uint32_t n = 1;
  const Value* cells = &v;
  if (v.kind() == ValueKind::Array) {
    n = v.as_array_rows() * v.as_array_cols();
    cells = v.as_array_cells();
  }
  out.reserve(n);
  for (std::uint32_t i = 0; i < n; ++i) {
    const Value& cell = cells[i];
    switch (cell.kind()) {
      case ValueKind::Number:
        out.push_back(cell.as_number());
        break;
      case ValueKind::Bool:
        if (policy.include_bool) {
          out.push_back(cell.as_boolean() ? 1.0 : 0.0);
        }
        break;
      case ValueKind::Text:
        if (policy.include_text_numeric_literal) {
          auto coerced = coerce_to_number(cell);
          if (!coerced) {
            if (policy.error_on_text) {
              return coerced.error();
            }
            // Silent skip on unparseable text — matches the lenient
            // `SMALL` / `LARGE` direct-scalar path when paired with
            // `include_text_numeric_literal = true, error_on_text = false`.
            break;
          }
          out.push_back(coerced.value());
        }
        // include_text_numeric_literal = false: silently skip Text
        // cells. This is the default AVERAGE / SUM behaviour.
        break;
      case ValueKind::Error:
        if (policy.error_on_error_cell) {
          return cell.as_error();
        }
        // Silent skip — the caller (e.g. regression families) runs its
        // own error-propagation pass over the unflattened arrays
        // first, so a stray error here is dropped to avoid double-
        // propagation.
        break;
      case ValueKind::Blank:
        if (policy.blank_as_zero) {
          out.push_back(0.0);
        }
        break;
      case ValueKind::Array:
      case ValueKind::Ref:
      case ValueKind::Lambda:
        // Always drop. Nested Array would only appear via lambda /
        // dynamic-array machinery that has its own flattening step.
        break;
    }
  }
  return out;
}

Expected<std::vector<double>, ErrorCode> collect_numerics(const Value* args, std::uint32_t count,
                                                          NumericCollectPolicy policy) {
  // Multi-Value flatten. The dispatcher has already expanded any range
  // arguments into scalar cells, so walking the flat slice cell-by-cell
  // and applying the same per-kind policy as the single-Value overload
  // produces an identical result. Errors short-circuit at the first
  // offending cell (when `policy.error_on_error_cell` is set), matching
  // the contract used by the AVERAGE / SMALL / LARGE / "A" families.
  std::vector<double> out;
  out.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    const Value& cell = args[i];
    switch (cell.kind()) {
      case ValueKind::Number:
        out.push_back(cell.as_number());
        break;
      case ValueKind::Bool:
        if (policy.include_bool) {
          out.push_back(cell.as_boolean() ? 1.0 : 0.0);
        }
        break;
      case ValueKind::Text:
        if (policy.include_text_numeric_literal) {
          auto coerced = coerce_to_number(cell);
          if (!coerced) {
            if (policy.error_on_text) {
              return coerced.error();
            }
            break;
          }
          out.push_back(coerced.value());
        }
        break;
      case ValueKind::Error:
        if (policy.error_on_error_cell) {
          return cell.as_error();
        }
        break;
      case ValueKind::Blank:
        if (policy.blank_as_zero) {
          out.push_back(0.0);
        }
        break;
      case ValueKind::Array:
      case ValueKind::Ref:
      case ValueKind::Lambda:
        // Always drop. Range / Array arguments are expected to have been
        // expanded by the dispatcher before reaching this helper.
        break;
    }
  }
  return out;
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
      // Mac Excel 365 (ja-JP) accepts ONLY the literal strings "TRUE" /
      // "FALSE" (ASCII case-insensitive, no whitespace tolerance) wherever
      // a Bool is expected via this coercion path: e.g.
      // `=IF("TRUE", 1, 0)` -> 1, `=NOT("false")` -> TRUE,
      // `=BETA.DIST(..., "TRUE", ...)`. Everything else surfaces
      // `#VALUE!` — including numeric strings ("0", "1", "0.5"),
      // whitespace-padded forms ("  TRUE  "), localised truth-words
      // ("VRAI", "WAHR", "真"), and decorated forms ("TRUE!!"). This
      // matches the behaviour the text_to_bool_probes oracle suite
      // recorded against Mac Excel 365 (ja-JP, build 16.108.1) on
      // 2026-04-26 — see `tests/oracle/cases/text_to_bool_probes.yaml`.
      //
      // The contract is identical to the stricter `logical_coerce` helper
      // used by AND / OR / XOR / IFS (`eval/logical_coerce.h`) modulo the
      // Skip vs. HasValue distinction; the two helpers stay in sync.
      const std::string_view text = v.as_text();
      if (strings::case_insensitive_eq(text, "TRUE")) {
        return true;
      }
      if (strings::case_insensitive_eq(text, "FALSE")) {
        return false;
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

}  // namespace eval
}  // namespace formulon
