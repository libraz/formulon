//
// Lazy impls for the short-circuit "special form" family: `IF`,
// `IFERROR`, and `IFNA`. These are routed through the dispatch table in
// `tree_walker.cpp` via the `lazy_impls.h` extern declarations. See that
// header for the dispatch-table contract and the motivation for the
// split.

#include "eval/special_forms_lazy.h"

#include <cstddef>
#include <cstdint>
#include <iterator>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/logical_coerce.h"
#include "eval/name_env.h"
#include "eval/range_args.h"
#include "eval/tree_walker/broadcast.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

LogicalCoerce logical_coerce_for_host(const Value& v, const EvalContext& ctx, bool* out_bool,
                                      ErrorCode* out_err) noexcept {
  if (ctx.excel_profile().host != ExcelHost::kWin365 || !v.is_text()) {
    return logical_coerce(v, out_bool, out_err);
  }
  const std::string_view raw = v.as_text();
  if (raw.empty()) {
    return LogicalCoerce::Skip;
  }
  const std::string_view text = strings::trim(raw);
  if (strings::case_insensitive_eq(text, "TRUE")) {
    *out_bool = true;
    return LogicalCoerce::HasValue;
  }
  if (strings::case_insensitive_eq(text, "FALSE")) {
    *out_bool = false;
    return LogicalCoerce::HasValue;
  }
  *out_err = ErrorCode::Value;
  return LogicalCoerce::Error;
}

// Provenance bucket for a lazy-aggregate argument (COUNT / AND / OR).
// Excel's rules are provenance-sensitive: values arriving from a range,
// spill, array constant, or dynamic-array producer are treated differently
// from a direct scalar argument.
enum class LazyArgShape : std::uint8_t { Range, Scalar };

// Classification of a single COUNT / AND / OR argument. `cells` is populated
// for `Range` shape (row-major); `scalar` for `Scalar` shape (which may hold
// an error Value the caller decides to skip or propagate). `range_failed`
// flags a range resolution that surfaced an Excel error (`#REF!`, ...).
struct LazyAggArg {
  LazyArgShape shape = LazyArgShape::Scalar;
  std::vector<Value> cells;
  Value scalar = Value::blank();
  bool range_failed = false;
  ErrorCode range_error = ErrorCode::Value;
};

// Resolves a COUNT / AND / OR argument to its provenance bucket.
//
// Every shape but one is delegated to `resolve_range_arg`, the engine's
// single range-argument resolver: it looks through LET `NameRef` bindings,
// expands `RangeOp` / single-cell `Ref` / `SpillRef` / `IntersectOp` and the
// range-shaped calls (`OFFSET` / `CHOOSE` / `IF` / `ROW` / `COLUMN`),
// evaluates an array constant `{...}` element-wise, and unwraps the
// `Value::Array` a dynamic-array producer (`SEQUENCE`, `FILTER`, `MUNIT`, a
// lambda helper, ...) returns. Its `from_scalar` flag is exactly the
// provenance distinction this family needs, so an argument shape taught to
// the resolver reaches COUNT / AND / OR without a change here.
//
// The exception is a reference union `(A1:A2, B1:B2)`, which denotes several
// rectangles and therefore has no `RangeResult` the resolver can return; its
// areas are concatenated below.
LazyAggArg resolve_lazy_agg_arg(const parser::AstNode& raw_arg, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx) {
  LazyAggArg out;
  if (raw_arg.kind() == parser::NodeKind::UnionOp) {
    // Areas concatenate in order WITHOUT de-duplication, so an overlapping
    // area is counted twice (`COUNT((A1:A2, A1:A2))` doubles). Mirrors the
    // eager dispatcher's union handling: a failed area contributes a single
    // error cell, which COUNT skips (`propagate_errors = false`) and AND / OR
    // propagate. A LET binding never carries a union AST, so the passthrough
    // the resolver performs cannot reach this shape.
    out.shape = LazyArgShape::Range;
    const std::uint32_t area_count = raw_arg.as_union_arity();
    for (std::uint32_t area = 0; area < area_count; ++area) {
      auto area_result = resolve_range_arg(raw_arg.as_union_child(area), arena, registry, ctx);
      if (!area_result) {
        out.cells.push_back(Value::error(area_result.error()));
        continue;
      }
      std::vector<Value>& area_cells = area_result.value().cells;
      out.cells.insert(out.cells.end(), std::make_move_iterator(area_cells.begin()),
                       std::make_move_iterator(area_cells.end()));
    }
    return out;
  }

  auto resolved = resolve_range_arg(raw_arg, arena, registry, ctx);
  if (!resolved) {
    // A shape the resolver rejects and a propagated cell error arrive the
    // same way; both consumers treat them alike (COUNT skips, AND / OR
    // surface the code), so they share the one failure bucket.
    out.shape = LazyArgShape::Range;
    out.range_failed = true;
    out.range_error = resolved.error();
    return out;
  }
  RangeResult& result = resolved.value();
  // A bare scalar expression keeps direct-argument provenance. The resolver
  // guarantees a single cell alongside `from_scalar`; an empty rectangle
  // falls through to the range bucket, where it contributes nothing.
  if (result.from_scalar && !result.cells.empty()) {
    out.shape = LazyArgShape::Scalar;
    out.scalar = result.cells.front();
    return out;
  }
  out.shape = LazyArgShape::Range;
  out.cells = std::move(result.cells);
  return out;
}

Value eval_and_or_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx, bool is_and) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1U) {
    return Value::error(ErrorCode::Value);
  }
  bool result = is_and;
  bool any_value = false;
  for (std::uint32_t i = 0; i < arity; ++i) {
    LazyAggArg resolved = resolve_lazy_agg_arg(call.as_call_arg(i), arena, registry, ctx);
    if (resolved.shape == LazyArgShape::Range) {
      if (resolved.range_failed) {
        return Value::error(resolved.range_error);
      }
      for (const Value& v : resolved.cells) {
        if (v.is_error()) {
          return v;
        }
        // Text / Blank cells arriving from a range or array are skipped
        // rather than coerced (Excel's range-provenance rule).
        if (v.is_text() || v.is_blank()) {
          continue;
        }
        bool coerced = false;
        ErrorCode err = ErrorCode::Value;
        const LogicalCoerce lc = logical_coerce_for_host(v, ctx, &coerced, &err);
        if (lc == LogicalCoerce::Error) {
          return Value::error(err);
        }
        if (lc == LogicalCoerce::Skip) {
          continue;
        }
        any_value = true;
        result = is_and ? (result && coerced) : (result || coerced);
      }
      continue;
    }
    // Direct scalar argument: coerce with the host-aware strict rule so
    // "TRUE" / "FALSE" text carries a bool and other text surfaces #VALUE!.
    const Value& v = resolved.scalar;
    if (v.is_error()) {
      return v;
    }
    bool coerced = false;
    ErrorCode err = ErrorCode::Value;
    const LogicalCoerce lc = logical_coerce_for_host(v, ctx, &coerced, &err);
    if (lc == LogicalCoerce::Error) {
      return Value::error(err);
    }
    if (lc == LogicalCoerce::Skip) {
      continue;
    }
    any_value = true;
    result = is_and ? (result && coerced) : (result || coerced);
  }
  if (!any_value) {
    return Value::error(ErrorCode::Value);
  }
  return Value::boolean(result);
}

// IF over an Array condition (Excel 365 dynamic-array spill). Unlike the
// scalar path there is NO short-circuit: Excel evaluates BOTH branches,
// broadcasts cond / then / else to the common `max` shape using the same
// rules as the binary operators (shared `ArrayView` helpers from
// `tree_walker/broadcast.h`), and per output cell coerces the condition
// cell to bool and picks the matching branch cell. Errors (a condition cell
// that is or coerces to an error, or a picked branch cell holding an error)
// land in that output cell only; a position an operand cannot supply is
// `#N/A`, exactly as `broadcast_binop` fills it.
Value eval_if_array(const parser::AstNode& call, const Value& cond, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const Value then_val = eval_node(call.as_call_arg(1), arena, registry, ctx);
  const Value else_val =
      call.as_call_arity() == 3 ? eval_node(call.as_call_arg(2), arena, registry, ctx) : Value::boolean(false);
  Value cond_slot = Value::blank();
  Value then_slot = Value::blank();
  Value else_slot = Value::blank();
  const ArrayView cv = as_array_view(cond, &cond_slot);
  const ArrayView tv = as_array_view(then_val, &then_slot);
  const ArrayView ev = as_array_view(else_val, &else_slot);

  std::uint32_t out_rows = cv.rows > tv.rows ? cv.rows : tv.rows;
  out_rows = out_rows > ev.rows ? out_rows : ev.rows;
  std::uint32_t out_cols = cv.cols > tv.cols ? cv.cols : tv.cols;
  out_cols = out_cols > ev.cols ? out_cols : ev.cols;

  Value* buf = nullptr;
  ArrayValue* out = allocate_array_value(out_rows, out_cols, arena, buf, kMaxDerivedArrayCells);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::size_t i = 0;
  for (std::uint32_t r = 0; r < out_rows; ++r) {
    for (std::uint32_t c = 0; c < out_cols; ++c, ++i) {
      const Value* cc = broadcast_cell(cv, r, c);
      if (cc == nullptr) {
        buf[i] = Value::error(ErrorCode::NA);
        continue;
      }
      if (cc->is_error()) {
        buf[i] = *cc;
        continue;
      }
      auto b = coerce_to_bool(*cc);
      if (!b) {
        buf[i] = Value::error(b.error());
        continue;
      }
      const Value* pick = b.value() ? broadcast_cell(tv, r, c) : broadcast_cell(ev, r, c);
      buf[i] = pick == nullptr ? Value::error(ErrorCode::NA) : *pick;
    }
  }
  return Value::array(out);
}

}  // namespace

// IF(cond, then, else?) - then is evaluated iff cond coerces to true; else
// is evaluated iff cond coerces to false. When the third argument is
// omitted Excel returns the boolean `FALSE` for the falsey path. An Array
// condition instead spills element-wise via `eval_if_array` above.
Value eval_if_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                   const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  const Value cond = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (cond.is_error()) {
    return cond;
  }
  if (cond.is_array()) {
    return eval_if_array(call, cond, arena, registry, ctx);
  }
  auto coerced = coerce_to_bool(cond);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  if (coerced.value()) {
    return eval_node(call.as_call_arg(1), arena, registry, ctx);
  }
  if (arity == 3) {
    return eval_node(call.as_call_arg(2), arena, registry, ctx);
  }
  return Value::boolean(false);
}

// IFERROR(value, fallback) - returns `value` unchanged unless it is any
// error, in which case `fallback` is evaluated and returned. The fallback
// subtree is NOT evaluated when `value` is non-error (true short-circuit).
// If `fallback` itself raises an error it is propagated as-is.
Value eval_iferror_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  const Value primary = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!primary.is_error()) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry, ctx);
}

// IFNA(value, fallback) - returns `value` unchanged unless it is exactly
// `#N/A`, in which case `fallback` is evaluated and returned. All other
// errors (including `#DIV/0!`, `#REF!`, `#VALUE!`, `#NAME?`) propagate as
// `value`. The fallback subtree is NOT evaluated unless the trigger fires.
Value eval_ifna_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                     const EvalContext& ctx) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  // Blank handling is kept consistent with IF / IFERROR above: a Blank
  // result is returned as Blank rather than promoted to number 0. The
  // implicit Blank->0 promotion happens later, when a formula cell's
  // ultimate value is returned to the grid, so it must not be applied
  // here or `ISBLANK(IFNA(<blank>, x))` would diverge from the IFERROR
  // analog (which preserves Blank).
  const Value primary = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (!(primary.is_error() && primary.as_error() == ErrorCode::NA)) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry, ctx);
}

Value eval_and_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  return eval_and_or_lazy(call, arena, registry, ctx, /*is_and=*/true);
}

Value eval_or_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                   const EvalContext& ctx) {
  return eval_and_or_lazy(call, arena, registry, ctx, /*is_and=*/false);
}

// COUNT(value, ...) - Excel's rule is provenance-sensitive, so the per-arg
// AST shape decides what counts:
//   * Range argument (A1:B2 etc.): count only Number cells. Range-sourced
//     Bool / Text / Blank / Error are skipped.
//   * Single-cell Ref argument (e.g. `COUNT(B3)`): treat the referenced
//     cell identically to a range cell -- only Number counts. A Bool, Text,
//     Blank, or Error sitting in a referenced cell is skipped.
//   * Any other argument shape (number/bool/text literal, arithmetic
//     subexpression, function call producing a scalar): count if the
//     result is a Number or Bool, or if it is Text that `coerce_to_number`
//     parses to a finite value (so `COUNT("23")` -> 1 but
//     `COUNT("Hola")` -> 0). Blank / Error / Array / Lambda are skipped.
// Implementing this correctly requires per-arg AST inspection, which is
// why COUNT is routed through the lazy dispatch table. A direct-arg error
// is silently skipped (COUNT is registered with `propagate_errors = false`
// in the eager path; this lazy impl mirrors that behaviour).
Value eval_count_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 1) {
    return Value::error(ErrorCode::Value);
  }
  double total = 0.0;
  for (std::uint32_t i = 0; i < arity; ++i) {
    LazyAggArg resolved = resolve_lazy_agg_arg(call.as_call_arg(i), arena, registry, ctx);
    if (resolved.shape == LazyArgShape::Range) {
      // Range / array / spill provenance: only Number cells count. Bool /
      // Text / Blank / Error are skipped, and a resolution failure (#REF!)
      // is silently ignored (COUNT is registered `propagate_errors=false`).
      if (resolved.range_failed) {
        continue;
      }
      for (const Value& v : resolved.cells) {
        if (v.is_number()) {
          total += 1.0;
        }
      }
      continue;
    }
    // Direct (literal / expression) argument: Number and Bool count, and
    // Text counts when it parses as a finite number. Blank / Error / Lambda
    // are skipped.
    const Value& v = resolved.scalar;
    if (v.is_number() || v.is_boolean()) {
      total += 1.0;
      continue;
    }
    if (v.is_text()) {
      if (coerce_to_number(v)) {
        total += 1.0;
      }
    }
  }
  return Value::number(total);
}

// IFS(cond1, val1, ...) - Excel's multi-branch short-circuit: the first
// TRUE condition wins and its paired value is evaluated and returned; all
// remaining branches (both conditions AND values) are skipped. Errors in
// an evaluated condition propagate. When no condition matches (including
// the degenerate odd-arity case) the result is #N/A, matching Excel's
// documented "if none match" behaviour.
//
// Conditions use the *strict* logical coercion shared with AND / OR / XOR
// (`eval/logical_coerce.h`), NOT the generic `coerce_to_bool` used by IF:
//
//   * Bool / finite Number / "TRUE" / "FALSE" text (case-insensitive,
//     trimmed) carry a bool value.
//   * Numeric-text like "0" / "1" surfaces `#VALUE!` — this is the Mac
//     Excel 365 rule that distinguishes IFS from IF.
//   * Blank is treated as FALSE (IFS walks past blank conditions rather
//     than rejecting them).
//   * Empty text is also treated as FALSE, matching the AND-family "Skip"
//     path (IFS collapses Skip to the false branch so `IFS("", 1, TRUE, 2)`
//     hits the catchall instead of erroring).
//
// The coercion uses the host-aware `logical_coerce_for_host`, the same seam
// AND / OR route through, so a text condition is judged consistently with
// the rest of the logical family under the active Excel profile.
Value eval_ifs_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                    const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity < 2) {
    return Value::error(ErrorCode::Value);
  }
  // Iterate in (cond, value) pairs. If the count is odd, the trailing
  // condition has no paired value; we still evaluate it for error
  // propagation, then fall through to #N/A.
  for (std::uint32_t i = 0; i + 1 < arity; i += 2) {
    const Value cond = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (cond.is_error()) {
      return cond;
    }
    bool truth = false;
    ErrorCode err = ErrorCode::Value;
    const LogicalCoerce lc = logical_coerce_for_host(cond, ctx, &truth, &err);
    if (lc == LogicalCoerce::Error) {
      return Value::error(err);
    }
    // Skip (Blank / empty-text) is treated as FALSE: fall through to the
    // next branch.
    if (lc == LogicalCoerce::HasValue && truth) {
      return eval_node(call.as_call_arg(i + 1), arena, registry, ctx);
    }
  }
  if ((arity % 2) == 1) {
    // Trailing unpaired condition: evaluate for error propagation only,
    // then fall through to #N/A regardless of its truth value.
    const Value trailing = eval_node(call.as_call_arg(arity - 1), arena, registry, ctx);
    if (trailing.is_error()) {
      return trailing;
    }
  }
  return Value::error(ErrorCode::NA);
}

// Equality test for SWITCH. This is type-strict and deliberately NOT the
// `=` operator: text comparison is ASCII case-insensitive, cross-type pairs
// never match, and there is no numeric<->text coercion (SWITCH(23, "23")
// and SWITCH("23", 23) both miss). The one Excel special case is that a
// blank subject matches a numeric 0 case -- but NOT "" and NOT FALSE, which
// is where SWITCH diverges from the `=` operator (where blank="" is TRUE).
bool switch_equal(const Value& lhs, const Value& rhs) {
  // Blank subject / case: matches a numeric 0 only. `switch_equal` is called
  // as switch_equal(subject, case), but the rule is symmetric here so both
  // orders are handled for safety.
  if (lhs.kind() == ValueKind::Blank && rhs.kind() == ValueKind::Number) {
    return rhs.as_number() == 0.0;
  }
  if (rhs.kind() == ValueKind::Blank && lhs.kind() == ValueKind::Number) {
    return lhs.as_number() == 0.0;
  }
  if (lhs.kind() != rhs.kind()) {
    return false;
  }
  switch (lhs.kind()) {
    case ValueKind::Number:
      return lhs.as_number() == rhs.as_number();
    case ValueKind::Bool:
      return lhs.as_boolean() == rhs.as_boolean();
    case ValueKind::Text:
      return strings::case_insensitive_eq(lhs.as_text(), rhs.as_text());
    case ValueKind::Blank:
      return true;
    default:
      return false;
  }
}

// SWITCH(expr, case1, val1, ..., [default]) - first case that equals
// `expr` wins; only that branch's value subtree is evaluated. An extra
// trailing argument (odd arity after expr) is the default. No match and
// no default -> #N/A. Errors in `expr` or in any evaluated case expression
// propagate.
Value eval_switch_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                       const EvalContext& ctx) {
  const std::uint32_t arity = call.as_call_arity();
  // Minimum useful form is SWITCH(expr, case, val): 3 args. A bare
  // SWITCH(expr) or SWITCH(expr, default) is rejected as an arity
  // violation (matches Excel's "You've entered too few arguments").
  if (arity < 3) {
    return Value::error(ErrorCode::Value);
  }
  const Value expr = eval_node(call.as_call_arg(0), arena, registry, ctx);
  if (expr.is_error()) {
    return expr;
  }
  // Walk (case, value) pairs starting at index 1. If a trailing single
  // argument remains at the end it is the default.
  std::uint32_t i = 1;
  while (i + 1 < arity) {
    const Value case_val = eval_node(call.as_call_arg(i), arena, registry, ctx);
    if (case_val.is_error()) {
      return case_val;
    }
    if (switch_equal(expr, case_val)) {
      return eval_node(call.as_call_arg(i + 1), arena, registry, ctx);
    }
    i += 2;
  }
  if (i < arity) {
    // Trailing default argument.
    return eval_node(call.as_call_arg(i), arena, registry, ctx);
  }
  return Value::error(ErrorCode::NA);
}

// ISOMITTED returns TRUE only when the argument resolves to a trailing
// optional LAMBDA parameter that the call site did not supply. The check
// is purely structural: the arg must be a bare `NameRef`, and that name
// must resolve in the active `NameEnv` to a binding whose `is_omitted`
// flag is set. Anything else — a literal, arithmetic, an unrelated name,
// or a populated optional slot — yields FALSE. Calls outside any LAMBDA
// invocation have no `NameEnv` (or one without the queried name) and
// therefore also yield FALSE, matching Mac Excel's "outside-lambda"
// behaviour.
//
// This is intentionally lazy: the eager dispatcher would flatten the
// argument to a Value before we could distinguish "omitted optional
// slot" from "regular binding bound to Blank".
Value eval_isomitted_lazy(const parser::AstNode& call, Arena& /*arena*/, const FunctionRegistry& /*registry*/,
                          const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const parser::AstNode& arg = call.as_call_arg(0);
  if (arg.kind() != parser::NodeKind::NameRef) {
    return Value::boolean(false);
  }
  const NameEnv* env = ctx.name_env();
  if (env == nullptr) {
    return Value::boolean(false);
  }
  return Value::boolean(env->lookup_omitted(arg.as_name()));
}

}  // namespace eval
}  // namespace formulon
