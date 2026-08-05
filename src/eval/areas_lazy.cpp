//
// Implementation of the `AREAS` lazy impl. The structural count of
// rectangles in a reference argument is computed by recursing over the
// AST: `UnionOp` children sum, `RangeOp` / `Ref` contribute 1, anything
// else is `#VALUE!`. See `eval/areas_lazy.h` for the dispatch contract.
//
// Reference-returning function calls are handled in two ways. CHOOSE and
// IF return one of their reference branches, so AREAS recurses into the
// branch that the index / condition selects and counts areas there (this
// makes `=AREAS(CHOOSE(2, A1:B2, (C1,D1)))` return 2). Other
// reference-returning calls (INDIRECT, OFFSET, INDEX, XLOOKUP, ...) are
// recognised by static name and counted as 1 area. The latter is an
// approximation: Excel can return a multi-area union via e.g.
// `=INDIRECT("A1,B2")`, which is actually 2 areas. Resolving that would
// require parsing / evaluating the INDIRECT string into a runtime
// reference union, which the engine cannot represent today; see the
// `areas-indirect-union` entry in tests/divergence.yaml.

#include "eval/areas_lazy.h"

#include <cstdint>
#include <string_view>

#include "eval/coerce.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "eval/range_resolvers.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Sentinel return values used by `count_areas` to communicate non-count
// outcomes through a single `int64_t` return. Negative values cannot be
// confused with a real Excel area count (always >= 0).
constexpr std::int64_t kAreasInvalid = -1;  // -> #VALUE!
constexpr std::int64_t kAreasNull = -2;     // -> #NULL! (disjoint intersection)
constexpr std::int64_t kAreasRef = -3;      // -> #REF! (cross-sheet intersection)

// Function names whose return value Excel treats as a reference for the
// purposes of AREAS but whose multi-area count we cannot determine
// statically. Membership counts the call as exactly 1 area. CHOOSE / IF
// are intentionally absent: they are handled by recursing into the
// selected branch (see `count_areas`). Names are uppercase canonical
// forms; AREAS arguments are matched case-insensitively below.
bool returns_single_reference(std::string_view name) noexcept {
  using strings::case_insensitive_eq;
  return case_insensitive_eq(name, "INDIRECT") || case_insensitive_eq(name, "OFFSET") ||
         case_insensitive_eq(name, "IFS") || case_insensitive_eq(name, "SWITCH") ||
         case_insensitive_eq(name, "INDEX") || case_insensitive_eq(name, "XLOOKUP");
}

// Recursively sums the leaf rectangles in a reference-shaped AST subtree.
// Returns a negative sentinel on structural mismatch so callers can
// translate it into the right Excel error without an extra out-parameter.
// `IntersectOp` requires a runtime rectangle test (disjoint -> `#NULL!`),
// so the `arena` / `registry` / `ctx` triple is threaded through to call
// `compute_intersect_rect`. Other kinds answer from shape alone.
std::int64_t count_areas(const parser::AstNode& n, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx) noexcept {
  const parser::NodeKind k = n.kind();
  if (k == parser::NodeKind::Ref || k == parser::NodeKind::RangeOp) {
    return 1;
  }
  if (k == parser::NodeKind::UnionOp) {
    std::int64_t total = 0;
    const std::uint32_t arity = n.as_union_arity();
    for (std::uint32_t i = 0; i < arity; ++i) {
      const std::int64_t c = count_areas(n.as_union_child(i), arena, registry, ctx);
      if (c < 0) {
        return c;  // propagate the most specific sentinel.
      }
      total += c;
    }
    return total;
  }
  if (k == parser::NodeKind::Call) {
    const std::string_view fname = n.as_call_name();
    // IF(cond, then, [else]) returns one of its reference branches; recurse
    // into the branch the condition selects. An omitted else branch with a
    // false condition yields a Bool in Excel (1 "area" by AREAS' contract is
    // not meaningful), so a missing branch falls back to the 1-area count.
    if (strings::case_insensitive_eq(fname, "IF")) {
      const std::uint32_t fn_arity = n.as_call_arity();
      if (fn_arity == 2U || fn_arity == 3U) {
        const Value cond = eval_node(n.as_call_arg(0), arena, registry, ctx);
        if (cond.is_error()) {
          return kAreasInvalid;
        }
        auto b = coerce_to_bool(cond);
        if (!b) {
          return kAreasInvalid;
        }
        if (b.value()) {
          return count_areas(n.as_call_arg(1), arena, registry, ctx);
        }
        if (fn_arity == 3U) {
          return count_areas(n.as_call_arg(2), arena, registry, ctx);
        }
      }
      return 1;
    }
    // CHOOSE(index, v1, v2, ...) returns the selected value branch; recurse
    // into it. A 1-based index outside [1, arity-1] is #VALUE! in Excel.
    if (strings::case_insensitive_eq(fname, "CHOOSE")) {
      const std::uint32_t fn_arity = n.as_call_arity();
      if (fn_arity >= 2U) {
        const Value idx_v = eval_node(n.as_call_arg(0), arena, registry, ctx);
        if (idx_v.is_error()) {
          return kAreasInvalid;
        }
        auto idx_e = coerce_to_number(idx_v);
        if (!idx_e) {
          return kAreasInvalid;
        }
        const std::int64_t idx = static_cast<std::int64_t>(idx_e.value());
        if (idx >= 1 && idx <= static_cast<std::int64_t>(fn_arity - 1U)) {
          return count_areas(n.as_call_arg(static_cast<std::uint32_t>(idx)), arena, registry, ctx);
        }
      }
      return kAreasInvalid;
    }
    if (returns_single_reference(fname)) {
      return 1;
    }
  }
  if (k == parser::NodeKind::IntersectOp) {
    std::string_view sheet;
    std::uint32_t r1 = 0;
    std::uint32_t c1 = 0;
    std::uint32_t r2 = 0;
    std::uint32_t c2 = 0;
    bool disjoint = false;
    ErrorCode err = ErrorCode::Value;
    if (!compute_intersect_rect(n.as_intersect_lhs(), n.as_intersect_rhs(), arena, registry, ctx, &sheet, &r1, &c1, &r2,
                                &c2, &disjoint, &err)) {
      // Cross-sheet -> #REF!; whole-col / whole-row endpoint -> #VALUE!;
      // any other resolution failure -> #VALUE! (the existing fallback).
      return err == ErrorCode::Ref ? kAreasRef : kAreasInvalid;
    }
    if (disjoint) {
      return kAreasNull;
    }
    return 1;
  }
  return kAreasInvalid;
}

}  // namespace

Value eval_areas_lazy(const parser::AstNode& call, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  if (call.as_call_arity() != 1U) {
    return Value::error(ErrorCode::Value);
  }
  const std::int64_t n = count_areas(call.as_call_arg(0), arena, registry, ctx);
  if (n == kAreasNull) {
    return Value::error(ErrorCode::Null);
  }
  if (n == kAreasRef) {
    return Value::error(ErrorCode::Ref);
  }
  if (n < 0) {
    return Value::error(ErrorCode::Value);
  }
  return Value::number(static_cast<double>(n));
}

}  // namespace eval
}  // namespace formulon
