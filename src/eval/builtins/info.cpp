// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of Formulon's info-style built-in functions:
// ISNUMBER, ISTEXT, ISBLANK, ISLOGICAL, ISERROR, ISERR, ISNA, ISEVEN,
// ISODD, ISNONTEXT, ERROR.TYPE, TYPE, N, T.
//
// The IS* family, plus ERROR.TYPE and TYPE, must inspect error-typed
// arguments verbatim, so each is registered with
// `propagate_errors = false` to opt out of the dispatcher's default
// left-most-error short-circuit. `N` and `T` use the default (errors
// propagate before the body runs). ISEVEN / ISODD coerce through
// `coerce_to_number` and therefore *do* propagate errors.

#include "eval/builtins/info.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// --- Type predicates ----------------------------------------------------
//
// These functions are registered with `propagate_errors = false` so the
// dispatcher hands them error-typed inputs verbatim. Each returns a Bool
// based purely on the input's `ValueKind` (and, for ISERR / ISNA, the
// specific `ErrorCode`). None of them ever surface a formula error of
// their own beyond the registry's arity check.

// ISNUMBER(value) - true iff the input is a Number cell.
Value IsNumber(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Number);
}

// ISTEXT(value) - true iff the input is a Text cell.
Value IsText(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Text);
}

// ISBLANK(value) - true iff the input is the Blank scalar. Empty text
// (`""`) is NOT blank in Excel - `ISBLANK("")` returns FALSE.
Value IsBlank(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Blank);
}

// ISLOGICAL(value) - true iff the input is a Bool. Numeric 0/1 do not
// count: only the actual TRUE/FALSE booleans qualify.
Value IsLogical(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Bool);
}

// ISERROR(value) - true iff the input is any formula error, including
// `#N/A`. Combined with the cleared `propagate_errors` flag this lets
// callers branch on errors without first wrapping them in IFERROR.
Value IsError(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() == ValueKind::Error);
}

// ISERR(value) - true iff the input is a formula error OTHER than `#N/A`.
// `ISERR(#N/A)` is FALSE; `ISERR(#DIV/0!)`, `ISERR(#REF!)`, etc. are TRUE.
Value IsErr(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() != ErrorCode::NA);
}

// ISNA(value) - true iff the input is exactly `#N/A`. All other errors
// (and all non-error values) yield FALSE.
Value IsNa(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::boolean(false);
  }
  return Value::boolean(v.as_error() == ErrorCode::NA);
}

// NA() - zero-argument constant that returns the #N/A error. Used as the
// explicit way to inject #N/A into a formula; typically paired with
// IFNA / ISNA. Distinct from lookup-style #N/A which surfaces from
// MATCH / VLOOKUP miss paths.
Value Na(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::NA);
}

// --- Coercion-style info functions --------------------------------------
//
// `N` and `T` propagate errors via the dispatcher's default short-circuit
// (no flag override needed). They differ from `VALUE` / generic
// `coerce_to_*` in that they NEVER fail: any non-matching input maps to
// the function's neutral element (0 for `N`, "" for `T`).

// N(value) - coerce to a Number with Excel's narrow rules:
//   - Number          -> the number unchanged
//   - Bool            -> 1.0 for TRUE, 0.0 for FALSE
//   - Text            -> 0.0 ALWAYS (N intentionally does not parse text;
//                        contrast with VALUE, which does)
//   - Blank           -> 0.0
//   - Error           -> propagated by the dispatcher before this body runs
Value N(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  switch (v.kind()) {
    case ValueKind::Number:
      return v;
    case ValueKind::Bool:
      return Value::number(v.as_boolean() ? 1.0 : 0.0);
    case ValueKind::Text:
    case ValueKind::Blank:
      return Value::number(0.0);
    case ValueKind::Error:
      // Unreachable in practice: dispatcher short-circuits errors. Defensive
      // fall-through keeps the switch exhaustive.
      return v;
    case ValueKind::Array:
    case ValueKind::Ref:
    case ValueKind::Lambda:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

// T(value) - coerce to Text with Excel's narrow rules:
//   - Text            -> the text unchanged
//   - Number / Bool   -> empty string ""
//   - Blank           -> empty string ""
//   - Error           -> propagated by the dispatcher before this body runs
Value T(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() == ValueKind::Text) {
    return v;
  }
  if (v.kind() == ValueKind::Error) {
    // Unreachable in practice: dispatcher short-circuits errors.
    return v;
  }
  return Value::text({});
}

// --- Parity predicates --------------------------------------------------
//
// ISEVEN / ISODD reject Bool inputs with `#VALUE!` (Mac Excel 365 does
// *not* coerce TRUE/FALSE to 1/0 here, unlike most numeric contexts).
// All other inputs go through `coerce_to_number` (blank->0, numeric text
// parsed), then truncate toward zero and check the low bit. Non-numeric
// coercion failure surfaces as `#VALUE!`. Errors propagate through the
// dispatcher's default short-circuit (no `propagate_errors = false`
// opt-out here).
//
// Example: `ISEVEN(-2.9)` truncates to `-2` -> TRUE. `ISEVEN(TRUE)` ->
// `#VALUE!`. `ISEVEN("3")` -> text "3" parses as number 3 -> FALSE.

// Shared helper: coerces, truncates, and returns the low bit as a bool
// under the caller-supplied predicate. Returns Expected so the caller can
// propagate `#VALUE!` from a failed coercion or a rejected Bool argument.
Expected<bool, ErrorCode> truncated_is_even(const Value& v) {
  // Excel rejects Bool inputs to ISEVEN/ISODD outright, without the usual
  // TRUE->1 / FALSE->0 coercion. Guard before `coerce_to_number`.
  if (v.kind() == ValueKind::Bool) {
    return ErrorCode::Value;
  }
  auto coerced = coerce_to_number(v);
  if (!coerced) {
    return coerced.error();
  }
  const double x = coerced.value();
  if (std::isnan(x) || std::isinf(x)) {
    return ErrorCode::Num;
  }
  // `std::trunc` rounds toward zero; casting through int64_t keeps the
  // low-bit check exact for any |x| <= 2^63. Larger magnitudes wrap, but
  // Excel's own behaviour at that range is also undefined — matching
  // open-source xlcalc implementations.
  const auto truncated = static_cast<std::int64_t>(std::trunc(x));
  // C++ modulo on negative operands is implementation-defined-free only
  // for `%2` returning 0/-1/1, so fold back to unsigned parity.
  return (truncated % 2) == 0;
}

// ISEVEN(value) - truncated integer of value is even. Errors propagate.
Value IsEven(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto even = truncated_is_even(args[0]);
  if (!even) {
    return Value::error(even.error());
  }
  return Value::boolean(even.value());
}

// ISODD(value) - truncated integer of value is odd. Errors propagate.
Value IsOdd(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  auto even = truncated_is_even(args[0]);
  if (!even) {
    return Value::error(even.error());
  }
  return Value::boolean(!even.value());
}

// --- ISNONTEXT(value) --------------------------------------------------
//
// TRUE iff the argument is NOT text. Blank / Number / Bool / Error all
// return TRUE. Only a Text-kind value returns FALSE. Registered with
// `propagate_errors = false` so `ISNONTEXT(#DIV/0!) = TRUE` (errors do
// NOT short-circuit through this function).
Value IsNonText(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::boolean(args[0].kind() != ValueKind::Text);
}

// --- ERROR.TYPE(value) --------------------------------------------------
//
// Maps an Excel error value to its documented integer code. Non-error
// inputs surface as `#N/A` (Excel's documented behaviour). Registered
// with `propagate_errors = false` so the error argument is passed
// through verbatim rather than short-circuited.
//
//   #NULL!         -> 1
//   #DIV/0!        -> 2
//   #VALUE!        -> 3
//   #REF!          -> 4
//   #NAME?         -> 5
//   #NUM!          -> 6
//   #N/A           -> 7
//   #GETTING_DATA  -> 8
//
// Formulon's current ErrorCode enum does not have a #GETTING_DATA variant;
// if one is added later the switch will need a new arm.
Value ErrorType(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  const Value& v = args[0];
  if (v.kind() != ValueKind::Error) {
    return Value::error(ErrorCode::NA);
  }
  switch (v.as_error()) {
    case ErrorCode::Null:
      return Value::number(1.0);
    case ErrorCode::Div0:
      return Value::number(2.0);
    case ErrorCode::Value:
      return Value::number(3.0);
    case ErrorCode::Ref:
      return Value::number(4.0);
    case ErrorCode::Name:
      return Value::number(5.0);
    case ErrorCode::Num:
      return Value::number(6.0);
    case ErrorCode::NA:
      return Value::number(7.0);
    case ErrorCode::GettingData:
      return Value::number(8.0);
    // Newer Excel errors (Spill, Calc, Field, Blocked, Connect, External,
    // Busy, Python, Unknown) are not in the classic 1..8 ERROR.TYPE table
    // and Microsoft's ja-JP Excel 365 Mac surfaces varying codes for them
    // depending on build. We fall through to #N/A (matching the
    // "non-standard error -> N/A" safety rail) until the oracle corpus
    // pins a definitive code.
    case ErrorCode::Spill:
    case ErrorCode::Calc:
    case ErrorCode::Field:
    case ErrorCode::Blocked:
    case ErrorCode::Connect:
    case ErrorCode::External:
    case ErrorCode::Busy:
    case ErrorCode::Python:
    case ErrorCode::Unknown:
      return Value::error(ErrorCode::NA);
  }
  return Value::error(ErrorCode::NA);
}

// --- INFO(type_text) ----------------------------------------------------
//
// Returns a string describing aspects of the runtime environment. Excel's
// real implementation probes the host OS, filesystem, and calc settings;
// Formulon is deliberately host-agnostic (it runs embedded in WASM /
// Python / CLI), so we return fixed, build-time-stable strings that are
// correct for "a hypothetical calc-only sandbox".
//
// Divergence note: ja-JP Mac Excel 365 returns "mac" for INFO("system");
// Formulon returns "pcdos" for all hosts. This matches Excel's historical
// fallback on platforms without OS introspection and keeps results
// reproducible across our build targets. The "directory", "numfile",
// "osversion" values are similarly stubbed - the oracle corpus
// intentionally skips INFO for this reason.
//
// `type_text` is compared ASCII-case-insensitively. Any unsupported
// type_text - including a numeric argument or a non-text coercion - surfaces
// as `#VALUE!`.

// ASCII-lowercase helper for the INFO type_text lookup. Non-ASCII bytes
// pass through unchanged; the supported keys are all ASCII so that's
// sufficient.
std::string ascii_tolower(std::string_view s) {
  std::string out;
  out.reserve(s.size());
  for (char c : s) {
    const auto u = static_cast<unsigned char>(c);
    if (u >= 'A' && u <= 'Z') {
      out.push_back(static_cast<char>(u + 32));
    } else {
      out.push_back(c);
    }
  }
  return out;
}

Value Info(const Value* args, std::uint32_t /*arity*/, Arena& arena) {
  // INFO only accepts Text (and Blank -> ""); numeric / bool arguments
  // surface as `#VALUE!` after the case-insensitive lookup fails. We
  // still route through `coerce_to_text` to keep the contract uniform,
  // but a Number coerces to its textual form ("5", "3.14", ...) which is
  // never a valid info_type and therefore falls through to the #VALUE!
  // default.
  auto text = coerce_to_text(args[0]);
  if (!text) {
    return Value::error(text.error());
  }
  const std::string key = ascii_tolower(text.value());
  std::string_view result;
  if (key == "directory") {
    result = "/";
  } else if (key == "numfile") {
    // The single bound workbook counts as one open file.
    return Value::number(1.0);
  } else if (key == "osversion") {
    result = "Formulon";
  } else if (key == "recalc") {
    result = "Automatic";
  } else if (key == "release") {
    result = "Formulon 0.1";
  } else if (key == "system") {
    // Excel's conventional fallback on hosts without OS introspection.
    // Mac Excel ja-JP returns "mac"; we diverge intentionally (see note
    // above).
    result = "pcdos";
  } else {
    return Value::error(ErrorCode::Value);
  }
  return Value::text(arena.intern(result));
}

// --- TYPE(value) --------------------------------------------------------
//
// Maps the argument's runtime kind to Excel's documented type code:
//
//   Number  -> 1
//   Text    -> 2
//   Bool    -> 4
//   Error   -> 16
//   Array   -> 64
//   Blank   -> 1 (Excel treats an empty cell as Number for TYPE purposes)
//
// Registered with `propagate_errors = false` so `TYPE(#DIV/0!) = 16`
// rather than short-circuiting.
Value Type(const Value* args, std::uint32_t /*arity*/, Arena& /*arena*/) {
  switch (args[0].kind()) {
    case ValueKind::Number:
    case ValueKind::Blank:
      return Value::number(1.0);
    case ValueKind::Text:
      return Value::number(2.0);
    case ValueKind::Bool:
      return Value::number(4.0);
    case ValueKind::Error:
      return Value::number(16.0);
    case ValueKind::Array:
      return Value::number(64.0);
    case ValueKind::Ref:
    case ValueKind::Lambda:
      // Ref / Lambda are not user-visible scalar kinds in Excel; fall
      // through to `#VALUE!` so unexpected shapes surface rather than
      // silently returning a misleading type code.
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

}  // namespace

void register_info_builtins(FunctionRegistry& registry) {
  // The IS* family must inspect error-typed arguments verbatim, so each
  // entry clears `propagate_errors` to opt out of the dispatcher's default
  // left-most-error short-circuit. `N` and `T` use the default (errors
  // propagate before the body runs).
  registry.register_function(FunctionDef{"ISNUMBER", 1u, 1u, &IsNumber, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISTEXT", 1u, 1u, &IsText, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISBLANK", 1u, 1u, &IsBlank, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISLOGICAL", 1u, 1u, &IsLogical, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERROR", 1u, 1u, &IsError, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISERR", 1u, 1u, &IsErr, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ISNA", 1u, 1u, &IsNa, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"NA", 0u, 0u, &Na});
  registry.register_function(FunctionDef{"N", 1u, 1u, &N});
  registry.register_function(FunctionDef{"T", 1u, 1u, &T});

  // Parity predicates: coerce through `coerce_to_number`, so errors
  // propagate through the dispatcher's default short-circuit (do NOT
  // clear `propagate_errors`).
  registry.register_function(FunctionDef{"ISEVEN", 1u, 1u, &IsEven});
  registry.register_function(FunctionDef{"ISODD", 1u, 1u, &IsOdd});

  // ISNONTEXT, ERROR.TYPE, TYPE all inspect error values directly rather
  // than propagating them, so they opt out of the dispatcher's default
  // left-most-error short-circuit just like the IS* family above.
  registry.register_function(FunctionDef{"ISNONTEXT", 1u, 1u, &IsNonText, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"ERROR.TYPE", 1u, 1u, &ErrorType, /*propagate_errors=*/false});
  registry.register_function(FunctionDef{"TYPE", 1u, 1u, &Type, /*propagate_errors=*/false});

  // INFO returns build-time-stable environment strings (host-agnostic).
  // Errors propagate; any non-matching type_text surfaces #VALUE!.
  registry.register_function(FunctionDef{"INFO", 1u, 1u, &Info});
  // ROWS / COLUMNS / ROW / COLUMN are routed through the lazy dispatch
  // table in `tree_walker.cpp` (see `eval_rows_lazy` et al. in
  // `shape_ops_lazy.cpp`) because they must introspect each argument's
  // AST shape: a bare single-cell `Ref`, a `RangeOp`, and an inline
  // `{...}` `ArrayLiteral` all produce different answers, and the eager
  // dispatcher's pre-evaluation would erase that distinction. Zero-arg
  // ROW() / COLUMN() read the formula-cell anchor that the recalc driver
  // attaches to `EvalContext` via `with_formula_cell`.
}

}  // namespace eval
}  // namespace formulon
