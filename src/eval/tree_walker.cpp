// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the tree-walk evaluator. See `tree_walker.h` for the
// public contract and the design references.

#include "eval/tree_walker.h"

#include <cmath>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/coerce.h"
#include "eval/function_registry.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/expected.h"  // FM_CHECK
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// ---------------------------------------------------------------------------
// Cross-type comparison
// ---------------------------------------------------------------------------

// Excel cross-type comparison order: Number < Text < Bool. Blank coerces to
// numeric zero. Text equality and ordering are case-insensitive over ASCII
// letters; locale-aware comparison is deferred. NaN compares as "unordered":
// every relational operator returns FALSE except `<>`.
//
// `out_unordered` is set to true iff one of the operands is NaN; the caller
// uses it to short-circuit relational operators to FALSE while still
// returning TRUE for `<>`. The integer return value (-1/0/+1) is meaningful
// only when `out_unordered` is false.
int compare_values(const Value& lhs, const Value& rhs, bool* out_unordered) {
  *out_unordered = false;

  auto rank = [](ValueKind k) -> int {
    // Blank is treated as numeric (0) for comparison purposes.
    switch (k) {
      case ValueKind::Number:
      case ValueKind::Blank:
        return 0;
      case ValueKind::Text:
        return 1;
      case ValueKind::Bool:
        return 2;
      default:
        return 3;
    }
  };

  const int lr = rank(lhs.kind());
  const int rr = rank(rhs.kind());
  if (lr != rr) {
    return lr < rr ? -1 : 1;
  }

  switch (lr) {
    case 0: {
      const double a = lhs.is_blank() ? 0.0 : lhs.as_number();
      const double b = rhs.is_blank() ? 0.0 : rhs.as_number();
      if (std::isnan(a) || std::isnan(b)) {
        *out_unordered = true;
        return 0;
      }
      if (a < b) {
        return -1;
      }
      if (a > b) {
        return 1;
      }
      return 0;
    }
    case 1: {
      const std::string_view a = lhs.as_text();
      const std::string_view b = rhs.as_text();
      const std::size_t n = a.size() < b.size() ? a.size() : b.size();
      for (std::size_t i = 0; i < n; ++i) {
        const char ca = strings::ascii_to_lower(a[i]);
        const char cb = strings::ascii_to_lower(b[i]);
        if (ca != cb) {
          return ca < cb ? -1 : 1;
        }
      }
      if (a.size() != b.size()) {
        return a.size() < b.size() ? -1 : 1;
      }
      return 0;
    }
    case 2: {
      const bool a = lhs.as_boolean();
      const bool b = rhs.as_boolean();
      if (a == b) {
        return 0;
      }
      // FALSE < TRUE.
      return a ? 1 : -1;
    }
    default:
      return 0;
  }
}

// ---------------------------------------------------------------------------
// Per-operator helpers
// ---------------------------------------------------------------------------

Value finalize_arithmetic(double r) {
  if (std::isnan(r) || std::isinf(r)) {
    return Value::error(ErrorCode::Num);
  }
  return Value::number(r);
}

Value apply_unary(parser::UnaryOp op, const Value& operand) {
  if (operand.is_error()) {
    return operand;
  }
  auto coerced = coerce_to_number(operand);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  const double x = coerced.value();
  switch (op) {
    case parser::UnaryOp::Plus:
      return finalize_arithmetic(x);
    case parser::UnaryOp::Minus:
      return finalize_arithmetic(-x);
    case parser::UnaryOp::Percent:
      return finalize_arithmetic(x / 100.0);
  }
  return Value::error(ErrorCode::Value);
}

Value apply_arithmetic(parser::BinOp op, double lhs, double rhs) {
  switch (op) {
    case parser::BinOp::Add:
      return finalize_arithmetic(lhs + rhs);
    case parser::BinOp::Sub:
      return finalize_arithmetic(lhs - rhs);
    case parser::BinOp::Mul:
      return finalize_arithmetic(lhs * rhs);
    case parser::BinOp::Div:
      // Excel reports #DIV/0! for any division whose divisor is exactly
      // zero, including 0/0 (no #NUM! tie-break).
      if (rhs == 0.0) {
        return Value::error(ErrorCode::Div0);
      }
      return finalize_arithmetic(lhs / rhs);
    case parser::BinOp::Pow: {
      // Delegates to the shared `apply_pow` helper so the `^` operator and
      // the `POWER()` builtin cannot drift apart on edge cases.
      auto r = apply_pow(lhs, rhs);
      if (!r) {
        return Value::error(r.error());
      }
      return Value::number(r.value());
    }
    default:
      // Caller guarantees op is arithmetic.
      FM_CHECK(false, "apply_arithmetic called with non-arithmetic op");
      return Value::error(ErrorCode::Value);
  }
}

Value apply_concat(const Value& lhs, const Value& rhs, Arena& arena) {
  auto lhs_text = coerce_to_text(lhs);
  if (!lhs_text) {
    return Value::error(lhs_text.error());
  }
  auto rhs_text = coerce_to_text(rhs);
  if (!rhs_text) {
    return Value::error(rhs_text.error());
  }
  std::string joined;
  joined.reserve(lhs_text.value().size() + rhs_text.value().size());
  joined.append(lhs_text.value());
  joined.append(rhs_text.value());
  const std::string_view interned = arena.intern(joined);
  // Empty input is fine: Arena::intern returns an empty view that is still
  // a valid Text payload.
  return Value::text(interned);
}

Value apply_comparison(parser::BinOp op, const Value& lhs, const Value& rhs) {
  bool unordered = false;
  const int cmp = compare_values(lhs, rhs, &unordered);
  switch (op) {
    case parser::BinOp::Eq:
      return Value::boolean(!unordered && cmp == 0);
    case parser::BinOp::NotEq:
      // NaN != anything is TRUE, matching IEEE-754 semantics.
      return Value::boolean(unordered || cmp != 0);
    case parser::BinOp::Lt:
      return Value::boolean(!unordered && cmp < 0);
    case parser::BinOp::LtEq:
      return Value::boolean(!unordered && cmp <= 0);
    case parser::BinOp::Gt:
      return Value::boolean(!unordered && cmp > 0);
    case parser::BinOp::GtEq:
      return Value::boolean(!unordered && cmp >= 0);
    default:
      FM_CHECK(false, "apply_comparison called with non-comparison op");
      return Value::error(ErrorCode::Value);
  }
}

// ---------------------------------------------------------------------------
// Recursive evaluator
// ---------------------------------------------------------------------------

Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry);

// ---------------------------------------------------------------------------
// Lazy (short-circuit) function impls
// ---------------------------------------------------------------------------
//
// Each lazy impl receives the full `Call` AST node so it can pull arguments
// out by index and decide which subtrees to evaluate. The eager path in
// `dispatch_call` is bypassed entirely: arity checks and error propagation
// belong inside each impl. On arity mismatch the impls return #VALUE! to
// match the eager dispatcher's behaviour.

// IF(cond, then, else?) - then is evaluated iff cond coerces to true; else
// is evaluated iff cond coerces to false. When the third argument is
// omitted Excel returns the boolean `FALSE` for the falsey path.
Value eval_if_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry) {
  const std::uint32_t arity = call.as_call_arity();
  if (arity != 2 && arity != 3) {
    return Value::error(ErrorCode::Value);
  }
  const Value cond = eval_node(call.as_call_arg(0), arena, registry);
  if (cond.is_error()) {
    return cond;
  }
  auto coerced = coerce_to_bool(cond);
  if (!coerced) {
    return Value::error(coerced.error());
  }
  if (coerced.value()) {
    return eval_node(call.as_call_arg(1), arena, registry);
  }
  if (arity == 3) {
    return eval_node(call.as_call_arg(2), arena, registry);
  }
  return Value::boolean(false);
}

// IFERROR(value, fallback) - returns `value` unchanged unless it is any
// error, in which case `fallback` is evaluated and returned. The fallback
// subtree is NOT evaluated when `value` is non-error (true short-circuit).
// If `fallback` itself raises an error it is propagated as-is.
Value eval_iferror_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  const Value primary = eval_node(call.as_call_arg(0), arena, registry);
  if (!primary.is_error()) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry);
}

// IFNA(value, fallback) - returns `value` unchanged unless it is exactly
// `#N/A`, in which case `fallback` is evaluated and returned. All other
// errors (including `#DIV/0!`, `#REF!`, `#VALUE!`, `#NAME?`) propagate as
// `value`. The fallback subtree is NOT evaluated unless the trigger fires.
Value eval_ifna_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry) {
  if (call.as_call_arity() != 2) {
    return Value::error(ErrorCode::Value);
  }
  const Value primary = eval_node(call.as_call_arg(0), arena, registry);
  if (!(primary.is_error() && primary.as_error() == ErrorCode::NA)) {
    return primary;
  }
  return eval_node(call.as_call_arg(1), arena, registry);
}

using LazyImpl = Value (*)(const parser::AstNode& call, Arena& arena,
                           const FunctionRegistry& registry);

struct LazyEntry {
  const char* name;  // canonical UPPERCASE
  LazyImpl impl;
};

constexpr LazyEntry kLazyDispatch[] = {
    {"IF", &eval_if_lazy},
    {"IFERROR", &eval_iferror_lazy},
    {"IFNA", &eval_ifna_lazy},
};

const LazyEntry* find_lazy(std::string_view name) noexcept {
  for (const auto& e : kLazyDispatch) {
    if (strings::case_insensitive_eq(name, std::string_view(e.name))) {
      return &e;
    }
  }
  return nullptr;
}

// Special-cased function-call dispatch.
//
// Lazy entries (`IF`, `IFERROR`, `IFNA`) are routed through the table above;
// each impl owns its own arity check and chooses which subtrees to evaluate.
//
// All other names are routed through `registry`: unknown name -> #NAME?,
// arity violation -> #VALUE!, otherwise every argument is pre-evaluated in
// order and the left-most error short-circuits before the impl runs.
Value dispatch_call(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  const std::string_view name = node.as_call_name();
  const std::uint32_t arity = node.as_call_arity();

  if (const LazyEntry* lazy = find_lazy(name); lazy != nullptr) {
    return lazy->impl(node, arena, registry);
  }

  const FunctionDef* def = registry.lookup(name);
  if (def == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  if (arity < def->min_arity || arity > def->max_arity) {
    return Value::error(ErrorCode::Value);
  }

  // Pre-evaluate arguments left-to-right; first error wins.
  std::vector<Value> values;
  values.reserve(arity);
  for (std::uint32_t i = 0; i < arity; ++i) {
    Value v = eval_node(node.as_call_arg(i), arena, registry);
    if (v.is_error()) {
      return v;
    }
    values.push_back(v);
  }
  return def->impl(values.data(), arity, arena);
}

Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      return node.as_literal();

    case parser::NodeKind::ErrorLiteral:
      return Value::error(node.as_error_literal());

    case parser::NodeKind::ErrorPlaceholder:
      // Panic-mode skipped this subtree at parse time; we cannot do better
      // than #NAME? since the original tokens are unavailable.
      return Value::error(ErrorCode::Name);

    case parser::NodeKind::ImplicitIntersection:
      // Identity for scalars. Once arrays land this becomes the contraction
      // operator (1x1 selection from a column / row at the call site).
      return eval_node(node.as_implicit_intersection_operand(), arena, registry);

    case parser::NodeKind::UnaryOp:
      return apply_unary(node.as_unary_op(), eval_node(node.as_unary_operand(), arena, registry));

    case parser::NodeKind::BinaryOp: {
      const parser::BinOp op = node.as_binary_op();
      // Evaluate left first so error propagation honours the documented
      // left-most-wins rule from backup/plans/02-calc-engine.md §2.1.1.
      const Value lhs = eval_node(node.as_binary_lhs(), arena, registry);
      if (lhs.is_error()) {
        return lhs;
      }
      const Value rhs = eval_node(node.as_binary_rhs(), arena, registry);
      if (rhs.is_error()) {
        return rhs;
      }

      switch (op) {
        case parser::BinOp::Add:
        case parser::BinOp::Sub:
        case parser::BinOp::Mul:
        case parser::BinOp::Div:
        case parser::BinOp::Pow: {
          auto lhs_n = coerce_to_number(lhs);
          if (!lhs_n) {
            return Value::error(lhs_n.error());
          }
          auto rhs_n = coerce_to_number(rhs);
          if (!rhs_n) {
            return Value::error(rhs_n.error());
          }
          return apply_arithmetic(op, lhs_n.value(), rhs_n.value());
        }
        case parser::BinOp::Concat:
          return apply_concat(lhs, rhs, arena);
        case parser::BinOp::Eq:
        case parser::BinOp::NotEq:
        case parser::BinOp::Lt:
        case parser::BinOp::LtEq:
        case parser::BinOp::Gt:
        case parser::BinOp::GtEq:
          return apply_comparison(op, lhs, rhs);
      }
      return Value::error(ErrorCode::Value);
    }

    case parser::NodeKind::Call:
      return dispatch_call(node, arena, registry);

    // -- Unsupported: name resolution / closures --------------------------
    case parser::NodeKind::Ref:
    case parser::NodeKind::ExternalRef:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::NameRef:
    case parser::NodeKind::LambdaCall:
    case parser::NodeKind::Lambda:
    case parser::NodeKind::LetBinding:
      return Value::error(ErrorCode::Name);

    // -- Unsupported: range-producing operators / array literals ----------
    case parser::NodeKind::RangeOp:
    case parser::NodeKind::UnionOp:
    case parser::NodeKind::IntersectOp:
    case parser::NodeKind::ArrayLiteral:
      return Value::error(ErrorCode::Value);
  }
  return Value::error(ErrorCode::Value);
}

}  // namespace

Value evaluate(const parser::AstNode& node, Arena& arena) {
  return evaluate(node, arena, default_registry());
}

Value evaluate(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry) {
  return eval_node(node, arena, registry);
}

}  // namespace eval
}  // namespace formulon
