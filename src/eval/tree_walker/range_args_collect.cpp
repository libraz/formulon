//
// Collects the cells of range-shaped call arguments into the flat value
// vector a range-aware (or element-wise lifted) function receives, applying
// the per-cell range filters of the function being called.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/lazy_impls.h"
#include "eval/range_args.h"
#include "eval/range_expanders.h"
#include "eval/range_resolvers.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker/dispatch_internal.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

using RangeCallExpander = bool (*)(const parser::AstNode&, Arena&, const FunctionRegistry&, const EvalContext&,
                                   std::vector<Value>*, ErrorCode*, std::uint32_t*, std::uint32_t*);

struct RangeCallExpansion {
  std::string_view name;
  RangeCallExpander expand;
};

// Calls a range-aware function sees as the cells they produce rather than
// the scalar they return: OFFSET, INDIRECT (a reference, so one cell still
// takes the range filters), CHOOSE (its picked child, recursively), and
// ROW / COLUMN over a range (Excel spills them to index vectors).
const RangeCallExpansion kRangeCallExpansions[] = {
    {"OFFSET", expand_offset_call}, {"INDIRECT", expand_indirect_call}, {"CHOOSE", expand_choose_call},
    {"ROW", expand_row_call},       {"COLUMN", expand_column_call},
};

}  // namespace

// A failed argument either aborts the call or, for a function that inspects
// its error arguments, lands in `values` like any other value. Returns false
// when the call must abort with `err`.
bool absorb_arg_error(const FunctionDef& def, const Value& err, std::vector<Value>* values) {
  if (def.propagate_errors) {
    return false;
  }
  forget_error_origin();
  values->push_back(err);
  return true;
}

namespace {

bool append_expanded_call_argument(const FunctionDef& def, const parser::AstNode& arg_node, std::string_view call_name,
                                   RangeCallExpander expand, Arena& arena, const FunctionRegistry& registry,
                                   const EvalContext& ctx, std::vector<Value>* values, bool* handled,
                                   Value* immediate_return) {
  *handled = false;
  if (!def.accepts_ranges || arg_node.kind() != parser::NodeKind::Call ||
      !strings::case_insensitive_eq(arg_node.as_call_name(), call_name)) {
    return true;
  }
  *handled = true;
  std::vector<Value> cells;
  ErrorCode err_code = ErrorCode::Value;
  if (!expand(arg_node, arena, registry, ctx, &cells, &err_code, nullptr, nullptr)) {
    const Value err = Value::error(err_code);
    if (!absorb_arg_error(def, err, values)) {
      *immediate_return = err;
      return false;
    }
    return true;
  }
  Value range_err = Value::blank();
  if (!append_range_sourced_values(def, cells.data(), cells.size(), values, &range_err)) {
    *immediate_return = range_err;
    return false;
  }
  return true;
}

// Expands the rectangle [`lhs`, `rhs`] into `values` through `def`'s range
// filters. An expansion error is pushed, or returned via `immediate_return`
// (with `false`) when `def` propagates errors.
bool append_expanded_range(const FunctionDef& def, const parser::Reference& lhs, const parser::Reference& rhs,
                           Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                           std::vector<Value>* values, Value* immediate_return) {
  auto expanded = ctx.expand_range(lhs, rhs, arena, registry);
  if (!expanded) {
    const Value err = Value::error(expanded.error());
    if (!absorb_arg_error(def, err, values)) {
      *immediate_return = err;
      return false;
    }
    return true;
  }
  Value range_err = Value::blank();
  const std::vector<Value>& cells = expanded.value();
  values->reserve(values->size() + cells.size());
  if (!append_range_sourced_values(def, cells.data(), cells.size(), values, &range_err)) {
    *immediate_return = range_err;
    return false;
  }
  return true;
}

// Resolves a bounded `RangeOp`'s two endpoints and hands back the union
// rectangle as the corner Refs `expand_range` takes. Endpoints may be
// plain Refs (the simple `A1:B2` form), reference-producing calls
// (`OFFSET(...)` / `INDIRECT(...)`) or names standing for either; anything
// else surfaces per `resolve_range_endpoint`'s error code. Both corners
// carry the sheet `merge_range_endpoint_sheets` settles.
//
// Returns `false` with the endpoint's error code in `*out_err`.
bool union_endpoint_refs(const parser::AstNode& range, Arena& arena, const FunctionRegistry& registry,
                         const EvalContext& ctx, parser::Reference* out_lhs, parser::Reference* out_rhs,
                         ErrorCode* out_err) {
  std::string_view sheet;
  std::uint32_t top = 0;
  std::uint32_t left = 0;
  std::uint32_t bottom = 0;
  std::uint32_t right = 0;
  if (!resolve_range_endpoints(range, arena, registry, ctx, &sheet, &top, &left, &bottom, &right, out_err)) {
    return false;
  }
  set_rect_corner_refs(sheet, top, left, bottom, right, out_lhs, out_rhs);
  return true;
}

// Ends a `collect_range_arg` branch on a failed argument: the error lands in
// `values` per `propagate_errors`, else it is the call's result.
RangeArgOutcome absorb_error_outcome(const FunctionDef& def, const Value& err, std::vector<Value>& values,
                                     Value& result) {
  if (!absorb_arg_error(def, err, &values)) {
    result = err;
    return RangeArgOutcome::kReturn;
  }
  return RangeArgOutcome::kCollected;
}

// Ends a `collect_range_arg` branch by appending `count` range cells through
// `def`'s range filters.
RangeArgOutcome append_sourced_outcome(const FunctionDef& def, const Value* cells, std::size_t count,
                                       std::vector<Value>& values, Value& result) {
  Value range_err = Value::blank();
  if (!append_range_sourced_values(def, cells, count, &values, &range_err)) {
    result = range_err;
    return RangeArgOutcome::kReturn;
  }
  return RangeArgOutcome::kCollected;
}

}  // namespace

// Collects the cells of an argument that denotes a range into `values`: for
// a range-aware function every range-shaped spelling, for a scalar function
// a bounded rectangle (which the caller then lifts element-wise). A failed
// argument aborts the call or lands in `values` per `propagate_errors`.
RangeArgOutcome collect_range_arg(const FunctionDef& def, const parser::AstNode& arg_node, Arena& arena,
                                  const FunctionRegistry& registry, const EvalContext& ctx, std::vector<Value>& values,
                                  bool& had_range_shaped_arg, Value& result) {
  // Minimal array-literal support: when a range-aware function receives
  // a `{a;b;c}` style literal, flatten it in row-major order exactly like
  // a RangeOp argument. This is just enough to let LARGE / SMALL /
  // PERCENTILE.INC / QUARTILE.INC accept brace literals as their "array"
  // input. The same literals are also first-class dynamic arrays when
  // evaluated directly; flattening here preserves this range-aware
  // function dispatch path.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::ArrayLiteral) {
    had_range_shaped_arg = true;
    auto resolved = resolve_range_arg(arg_node, arena, registry, ctx);
    if (!resolved) {
      result = Value::error(resolved.error());
      return RangeArgOutcome::kReturn;
    }
    bool short_circuit = false;
    Value propagated_err = Value::blank();
    for (const Value& value : resolved.value().cells) {
      // A boolean inside an array literal is #VALUE! under the
      // Analysis-ToolPak rule (`GCD({TRUE,2})`).
      if (def.analysis_toolpak_args && value.is_boolean()) {
        result = Value::error(ErrorCode::Value);
        return RangeArgOutcome::kReturn;
      }
      if (!append_array_sourced_value(def, value, &values, &propagated_err)) {
        short_circuit = true;
        break;
      }
    }
    if (short_circuit) {
      result = propagated_err;
      return RangeArgOutcome::kReturn;
    }
    return RangeArgOutcome::kCollected;
  }
  // Spilled-range `A1#` argument: resolve the spill region anchored at
  // the reference and flatten its row-major cells into the values
  // vector. Mirrors the RangeOp branch below; the same provenance-aware
  // filters apply because cells inside a spill region behave the same
  // way as cells inside an ordinary range when consumed by SUM /
  // AVERAGE / MIN / MAX / etc. A missing spill yields `#REF!`; an
  // unbound context yields `#NAME?`. Errors short-circuit per
  // `propagate_errors`.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::SpillRef) {
    had_range_shaped_arg = true;
    std::string_view anchor_sheet;
    std::uint32_t anchor_row = 0;
    std::uint32_t anchor_col = 0;
    ErrorCode spill_err = ErrorCode::Ref;
    if (!resolve_spill_anchor_node(arg_node, arena, registry, ctx, &anchor_sheet, &anchor_row, &anchor_col,
                                   &spill_err)) {
      return absorb_error_outcome(def, Value::error(spill_err), values, result);
    }
    const Sheet* current = ctx.current_sheet();
    const Sheet* target = current;
    if (current == nullptr) {
      spill_err = ErrorCode::Name;
    } else if (!anchor_sheet.empty()) {
      const Workbook* wb = ctx.workbook();
      if (wb == nullptr) {
        target = nullptr;
      } else {
        target = wb->sheet_by_name(anchor_sheet);
      }
    }
    if (target == nullptr || anchor_row >= Sheet::kMaxRows || anchor_col >= Sheet::kMaxCols) {
      return absorb_error_outcome(def, Value::error(spill_err), values, result);
    }
    // Copied out under the sheet lock, with Text payloads re-homed into
    // `arena`: the flattened values feed the whole call and outlive the
    // region they came from.
    std::vector<Value> region_cells;
    if (!ctx.read_spill_region(*target, anchor_row, anchor_col, arena, region_cells, nullptr, nullptr)) {
      return absorb_error_outcome(def, Value::error(ErrorCode::Ref), values, result);
    }
    return append_sourced_outcome(def, region_cells.data(), region_cells.size(), values, result);
  }
  // 3-D reference argument (`SUM(Sheet2:Sheet3!A1)`): resolve the inclusive
  // sheet span by workbook order, read the referenced cell from each sheet,
  // and flatten the resulting cells into the values vector. Mirrors the
  // RangeOp / SpillRef branches; the same provenance filters apply because
  // a 3-D ref is a range shape. A span endpoint that names a missing sheet
  // surfaces the error `expand_ref3d_cells` reports. Errors short-circuit
  // per `propagate_errors`.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::Ref3D && !is_grouped_ref3d(arg_node)) {
    std::vector<Value> ref3d_cells;
    ErrorCode ref3d_err = ErrorCode::Ref;
    if (!expand_ref3d_cells(arg_node, arena, registry, ctx, &ref3d_cells, &ref3d_err)) {
      return absorb_error_outcome(def, Value::error(ref3d_err), values, result);
    }
    had_range_shaped_arg = true;
    return append_sourced_outcome(def, ref3d_cells.data(), ref3d_cells.size(), values, result);
  }
  // Cross-workbook 3-D reference (`SUM([Book.xlsx]S1:S2!A1)`): the same
  // flattening over the link's sheet order, read from its cache.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::ExternalRef &&
      !arg_node.as_external_ref_sheet_end().empty()) {
    std::vector<Value> ref3d_cells;
    if (!collect_external_ref3d_cells(arg_node, arena, ctx, &ref3d_cells)) {
      return absorb_error_outcome(def, Value::error(ErrorCode::Ref), values, result);
    }
    had_range_shaped_arg = true;
    return append_sourced_outcome(def, ref3d_cells.data(), ref3d_cells.size(), values, result);
  }
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::RangeOp) {
    had_range_shaped_arg = true;
    const parser::AstNode& lhs_ast = arg_node.as_range_operand(0);
    const parser::AstNode& rhs_ast = arg_node.as_range_operand(1);
    // Multi-column (`A:C`) / multi-row (`1:3`) whole references parse as a
    // RangeOp over two whole-column / whole-row Refs. `resolve_range_endpoint`
    // rejects whole references (no bounded rectangle on their own), so route
    // the pair straight to `expand_range`, which clamps the unbounded axis to
    // the sheet's used range.
    //
    // A `:` chain of three or more endpoints is no bare pair;
    // `declared_rect_endpoint_pair` reduces the whole chain to the endpoint
    // pair that bounds it. Without it a chain falls into the endpoint union
    // below, which rejects whole references and turns `=SUM(A:C:E:G)` into
    // `#VALUE!` where Excel answers.
    parser::Reference chain_lhs{};
    parser::Reference chain_rhs{};
    const bool is_bare_pair = arg_node.as_range_arity() == 2U && lhs_ast.kind() == parser::NodeKind::Ref &&
                              rhs_ast.kind() == parser::NodeKind::Ref;
    const bool whole_axis_pair = is_bare_pair && has_whole_axis_endpoint(lhs_ast.as_ref(), rhs_ast.as_ref());
    if (whole_axis_pair) {
      chain_lhs = lhs_ast.as_ref();
      chain_rhs = rhs_ast.as_ref();
    } else if (is_bare_pair || !declared_rect_endpoint_pair(arg_node, &chain_lhs, &chain_rhs, ctx.workbook())) {
      // Union the two endpoints into one rectangle and let `expand_range`
      // validate sheet equality (mismatched qualifiers surface as #REF!).
      ErrorCode endpoint_err = ErrorCode::Ref;
      if (!union_endpoint_refs(arg_node, arena, registry, ctx, &chain_lhs, &chain_rhs, &endpoint_err)) {
        return absorb_error_outcome(def, Value::error(endpoint_err), values, result);
      }
    }
    Value range_return = Value::blank();
    if (!append_expanded_range(def, chain_lhs, chain_rhs, arena, registry, ctx, &values, &range_return)) {
      result = range_return;
      return RangeArgOutcome::kReturn;
    }
    return RangeArgOutcome::kCollected;
  }
  // Scalar (non-range-aware) function receiving a bounded `A1:B2`
  // RangeOp: Excel 365 evaluates the function element-wise over the
  // rectangle and spills the result rather than collapsing the range
  // to its implicit-intersection anchor. Materialise the rectangle as a
  // `Value::Array` here; the post-loop broadcast step then lifts the
  // scalar impl over it. Whole-column / whole-row endpoints (`A:A`,
  // `1:1`) keep the legacy anchor projection (they fall through to the
  // scalar `eval_node` path below) to avoid spilling an unbounded
  // rectangle through a scalar function. `@` / SINGLE-wrapped args are
  // Call nodes, not RangeOp, so they never reach this branch and retain
  // implicit-intersection semantics. A rectangle spanning a whole grid axis
  // (`A:C`, `1:3`, `A1:A1048576`) is not expanded here: it falls through to
  // `eval_node_shaped`, which reads it at its declared size without
  // materialising the unused part.
  if (!def.accepts_ranges && arg_node.kind() == parser::NodeKind::RangeOp) {
    const parser::AstNode& lhs_ast = arg_node.as_range_operand(0);
    const parser::AstNode& rhs_ast = arg_node.as_range_operand(1);
    const bool whole =
        parser::any_child_node(arg_node,
                               [](const parser::AstNode& operand) {
                                 return operand.kind() == parser::NodeKind::Ref &&
                                        has_whole_axis_endpoint(operand.as_ref(), operand.as_ref());
                               }) ||
        (arg_node.as_range_arity() == 2U && lhs_ast.kind() == parser::NodeKind::Ref &&
         rhs_ast.kind() == parser::NodeKind::Ref && is_full_axis_range(lhs_ast.as_ref(), rhs_ast.as_ref()));
    if (!whole) {
      parser::Reference union_lhs{};
      parser::Reference union_rhs{};
      ErrorCode endpoint_err = ErrorCode::Ref;
      if (!union_endpoint_refs(arg_node, arena, registry, ctx, &union_lhs, &union_rhs, &endpoint_err)) {
        return absorb_error_outcome(def, Value::error(endpoint_err), values, result);
      }
      auto expanded = ctx.expand_range(union_lhs, union_rhs, arena, registry);
      if (!expanded) {
        return absorb_error_outcome(def, Value::error(expanded.error()), values, result);
      }
      const std::uint32_t rrows = union_rhs.row - union_lhs.row + 1u;
      const std::uint32_t rcols = union_rhs.col - union_lhs.col + 1u;
      const std::vector<Value>& ev = expanded.value();
      // A single-cell range is that cell's value, not a 1x1 array.
      if (rrows == 1U && rcols == 1U && ev.size() == 1U) {
        values.push_back(ev[0]);
        return RangeArgOutcome::kCollected;
      }
      ArrayValue* arr = array_from_values(rrows, rcols, ev.data(), ev.size(), arena);
      if (arr == nullptr) {
        result = Value::error(ErrorCode::Num);
        return RangeArgOutcome::kReturn;
      }
      values.push_back(Value::array(arr));
      return RangeArgOutcome::kCollected;
    }
  }
  // Intersection operator as a range-aware function argument: Excel's
  // space operator (`A1:C3 B1:B5`) yields the overlapping rectangle,
  // and an aggregator must see every cell of that rectangle rather
  // than the single anchor `eval_node` would collapse it to. Compute
  // the clipped intersection rectangle and flatten it row-major,
  // mirroring the RangeOp branch above. Disjoint operands -> #NULL!.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::IntersectOp) {
    std::string_view isect_sheet;
    std::uint32_t isect_r1 = 0;
    std::uint32_t isect_c1 = 0;
    std::uint32_t isect_r2 = 0;
    std::uint32_t isect_c2 = 0;
    bool isect_disjoint = false;
    ErrorCode isect_err = ErrorCode::Value;
    if (!compute_intersect_rect(arg_node.as_intersect_lhs(), arg_node.as_intersect_rhs(), arena, registry, ctx,
                                &isect_sheet, &isect_r1, &isect_c1, &isect_r2, &isect_c2, &isect_disjoint,
                                &isect_err)) {
      return absorb_error_outcome(def, Value::error(isect_err), values, result);
    }
    if (isect_disjoint) {
      return absorb_error_outcome(def, Value::error(ErrorCode::Null), values, result);
    }
    had_range_shaped_arg = true;
    parser::Reference isect_lhs{};
    parser::Reference isect_rhs{};
    set_rect_corner_refs(isect_sheet, isect_r1, isect_c1, isect_r2, isect_c2, &isect_lhs, &isect_rhs);
    Value range_return = Value::blank();
    if (!append_expanded_range(def, isect_lhs, isect_rhs, arena, registry, ctx, &values, &range_return)) {
      result = range_return;
      return RangeArgOutcome::kReturn;
    }
    return RangeArgOutcome::kCollected;
  }
  // Union operator (`(A1:A2,B1:B2)`) as a range-aware function argument:
  // Excel's comma operator concatenates the areas, and an aggregator must
  // see every cell of every area. Resolve each area through
  // `resolve_range_arg` (which handles Ref / RangeOp / range-shaped calls)
  // and flatten them in order. Overlapping areas are intentionally counted
  // more than once — Excel's union does NOT de-duplicate, so
  // `SUM((A1:A2,A1:A2))` doubles the A1:A2 total.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::UnionOp) {
    had_range_shaped_arg = true;
    if (const std::optional<ErrorCode> invalid = union_reference_error(arg_node, ctx)) {
      const Value err = Value::error(*invalid);
      if (!absorb_arg_error(def, err, &values)) {
        result = err;
        return RangeArgOutcome::kReturn;
      }
      return RangeArgOutcome::kCollected;
    }
    const std::uint32_t area_count = arg_node.as_union_arity();
    Value range_err = Value::blank();
    bool short_circuit = false;
    for (std::uint32_t area = 0; area < area_count; ++area) {
      auto resolved = resolve_range_arg(arg_node.as_union_child(area), arena, registry, ctx);
      if (!resolved) {
        const Value err = Value::error(resolved.error());
        if (!absorb_arg_error(def, err, &values)) {
          result = err;
          return RangeArgOutcome::kReturn;
        }
        continue;
      }
      const RangeResult& rr = resolved.value();
      if (!append_range_sourced_values(def, rr.cells.data(), rr.cells.size(), &values, &range_err)) {
        short_circuit = true;
        break;
      }
    }
    if (short_circuit) {
      result = range_err;
      return RangeArgOutcome::kReturn;
    }
    return RangeArgOutcome::kCollected;
  }
  // Calls whose result is a range feed their cells through the range filters.
  {
    bool expanded_call_handled = false;
    Value expanded_call_return = Value::blank();
    for (const RangeCallExpansion& expansion : kRangeCallExpansions) {
      if (!append_expanded_call_argument(def, arg_node, expansion.name, expansion.expand, arena, registry, ctx, &values,
                                         &expanded_call_handled, &expanded_call_return)) {
        result = expanded_call_return;
        return RangeArgOutcome::kReturn;
      }
      if (expanded_call_handled) {
        had_range_shaped_arg = true;
        return RangeArgOutcome::kCollected;
      }
    }
  }
  // IF as a range producer: a picked range branch is flattened through the
  // range filters, so `=LET(r, IF(TRUE, A1:A3, B1:B3), SUM(r))` aggregates
  // three cells. A picked plain value (`IF(TRUE, "", 1)`) is a direct
  // argument, not a range cell, so the provenance filter must not drop it.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::Call &&
      strings::case_insensitive_eq(arg_node.as_call_name(), "IF")) {
    auto picked = resolve_range_arg(arg_node, arena, registry, ctx);
    if (!picked) {
      return absorb_error_outcome(def, Value::error(picked.error()), values, result);
    }
    if (picked.value().from_scalar && picked.value().cells.size() == 1U) {
      const Value& direct = picked.value().cells[0];
      if (direct.is_error()) {
        if (def.propagate_errors) {
          result = direct;
          return RangeArgOutcome::kReturn;
        }
        forget_error_origin();
      }
      values.push_back(direct);
      return RangeArgOutcome::kCollected;
    }
    had_range_shaped_arg = true;
    return append_sourced_outcome(def, picked.value().cells.data(), picked.value().cells.size(), values, result);
  }
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::StructuredRef) {
    // Structured (table) reference argument: evaluate it and unpack the
    // resulting Array row-major into the values vector. Mirrors the
    // SpillRef / RangeOp branches so SUM / AVERAGE / COUNTIF / ... see
    // the rectangle's cells exactly as if `Table[Col]` had been written
    // as a literal `A2:A10`. Errors short-circuit per `propagate_errors`.
    // Scalar (single-cell) results fall through to the generic argument
    // path; the only effect is that the per-cell `range_filter_*` rules
    // are not applied, which matches Mac for a single-cell `Sales[@Col]`
    // (Excel never broadcasts a single-cell structured ref through the
    // range-filter pipeline). The array-shaped path applies the filters
    // exactly like the RangeOp / SpillRef branches.
    Value sr_val = eval_node(arg_node, arena, registry, ctx);
    if (sr_val.is_error()) {
      if (def.propagate_errors) {
        result = sr_val;
        return RangeArgOutcome::kReturn;
      }
      forget_error_origin();
    }
    if (sr_val.is_array()) {
      had_range_shaped_arg = true;
      const ArrayValue* a = sr_val.as_array();
      const std::size_t sr_total = static_cast<std::size_t>(a->rows) * static_cast<std::size_t>(a->cols);
      return append_sourced_outcome(def, a->cells, sr_total, values, result);
    }
    // Scalar single-cell result: push without range-filtering. Falling
    // through to the bottom-of-loop scalar handling would re-eval the
    // node; the value we already have is correct.
    values.push_back(sr_val);
    return RangeArgOutcome::kCollected;
  }
  // Whole-column (`A:A`) / whole-row (`1:1`) single reference: expand
  // against the sheet's used range so range-aware aggregators see the
  // populated cells instead of the `#VALUE!` a scalar `resolve_ref`
  // returns for an unbounded reference. Multi-span (`A:C` / `1:3`) is
  // handled by the RangeOp branch above.
  if (def.accepts_ranges && arg_node.kind() == parser::NodeKind::Ref &&
      (arg_node.as_ref().is_full_col || arg_node.as_ref().is_full_row)) {
    had_range_shaped_arg = true;
    const parser::Reference& ref = arg_node.as_ref();
    Value range_return = Value::blank();
    if (!append_expanded_range(def, ref, ref, arena, registry, ctx, &values, &range_return)) {
      result = range_return;
      return RangeArgOutcome::kReturn;
    }
    return RangeArgOutcome::kCollected;
  }
  return RangeArgOutcome::kNotRange;
}

}  // namespace eval
}  // namespace formulon
