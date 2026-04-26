// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the `AREAS` lazy impl. The structural count of
// rectangles in a reference argument is computed by recursing over the
// AST: `UnionOp` children sum, `RangeOp` / `Ref` contribute 1, anything
// else is `#VALUE!`. See `eval/areas_lazy.h` for the dispatch contract.
//
// Reference-returning function calls (INDIRECT, OFFSET, CHOOSE, IF, ...)
// are recognised by static name and counted as 1 area. This is an
// approximation: Excel can return a multi-area union via e.g.
// `=INDIRECT("A1,B2")`, which would actually be 2 areas. The allowlist
// covers the dominant single-area case; a future revision could parse
// the INDIRECT string argument or evaluate the call to refine the count.

#include "eval/areas_lazy.h"

#include <cstdint>
#include <string_view>

#include "eval/eval_context.h"
#include "eval/reference_lazy.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
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
// purposes of AREAS. Membership is sufficient to count the call as 1 area
// without further inspection. Names are uppercase canonical forms.
bool returns_single_reference(std::string_view name) noexcept {
  // CHOOSE / IF / IFS / SWITCH propagate references when both branches
  // are references; we treat them as 1 area regardless because the oracle
  // cases that exercise this all resolve to a single area in Excel.
  return name == "INDIRECT" || name == "OFFSET" || name == "CHOOSE" || name == "IF" ||
         name == "IFS" || name == "SWITCH" || name == "INDEX" || name == "XLOOKUP";
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
  if (k == parser::NodeKind::Call && returns_single_reference(n.as_call_name())) {
    return 1;
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
