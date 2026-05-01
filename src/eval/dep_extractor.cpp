// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `extract_deps`. See `dep_extractor.h` for the contract.

#include "eval/dep_extractor.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/dep_graph.h"
#include "eval/volatile_tracker.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Walker state: collects results into `out` and tracks already-emitted cells
// in `seen` so the dedup runs in one O(N) pass.
struct WalkState {
  ExtractedDeps* out;
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  std::uint16_t current_sheet_id;
  const Workbook* workbook;
};

// Resolves the sheet id for a `Reference`. Returns true and writes
// `*out_sheet_id` on success; returns false (and writes nothing) when the
// qualifier names an unknown sheet. An empty qualifier resolves to the
// current sheet.
bool resolve_sheet_id(const parser::Reference& ref, const WalkState& state, std::uint16_t* out_sheet_id) {
  if (ref.sheet.empty()) {
    *out_sheet_id = state.current_sheet_id;
    return true;
  }
  // Linear scan over the workbook's sheets to find the index. The workbook
  // typically owns O(1)-O(10) sheets so a hash is premature; mirroring the
  // strategy used by `Workbook::sheet_by_name`.
  const std::size_t idx = state.workbook->sheet_index_by_name(ref.sheet);
  if (idx == static_cast<std::size_t>(-1)) {
    return false;
  }
  // Sheet ids fit in uint16_t per CellNodeId; Excel allows at most a few
  // thousand sheets in a workbook, well within range. Cast is safe.
  *out_sheet_id = static_cast<std::uint16_t>(idx);
  return true;
}

// Adds a single cell to the dependency set, deduplicating against `seen`.
void add_cell_dep(WalkState& state, CellNodeId cell) {
  if (state.seen.insert(cell).second) {
    state.out->cell_deps.push_back(cell);
  }
}

// Flattens the rectangle [lhs, rhs] into per-cell dependencies. Both
// endpoints must be plain `Ref` nodes; complex ranges (OFFSET-based,
// INDIRECT, etc.) are silently ignored — dynamic shapes cannot be statically
// resolved here. Whole-column / whole-row endpoints promote the formula to
// volatile status without adding individual cells, because flattening a
// 1,048,576-row column would blow up the dep graph.
void emit_range_cells(WalkState& state, const parser::Reference& lhs, const parser::Reference& rhs) {
  // Whole-column / whole-row: flag volatile and skip enumeration.
  if (lhs.is_full_col || lhs.is_full_row || rhs.is_full_col || rhs.is_full_row) {
    state.out->is_volatile = true;
    return;
  }

  // Resolve the effective sheet for the rectangle. Mirrors the policy in
  // `EvalContext::expand_range`: the parser keeps the qualifier on the LHS
  // for `Sheet2!A1:B2`, so the RHS often has an empty qualifier and inherits.
  std::string_view effective_sheet;
  if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
    // Mismatched qualifiers are an evaluator-level #REF!; statically we
    // simply skip the range.
    if (lhs.sheet != rhs.sheet) {
      return;
    }
    effective_sheet = lhs.sheet;
  } else if (!lhs.sheet.empty()) {
    effective_sheet = lhs.sheet;
  } else if (!rhs.sheet.empty()) {
    effective_sheet = rhs.sheet;
  }

  parser::Reference probe{};
  probe.sheet = effective_sheet;
  std::uint16_t target_sheet_id = state.current_sheet_id;
  if (!resolve_sheet_id(probe, state, &target_sheet_id)) {
    return;
  }

  // Bounds: clamp out-of-range coordinates by skipping. The evaluator turns
  // these into #REF! at evaluation time; for static dep tracking they are
  // simply absent.
  if (lhs.row >= Sheet::kMaxRows || lhs.col >= Sheet::kMaxCols ||  //
      rhs.row >= Sheet::kMaxRows || rhs.col >= Sheet::kMaxCols) {
    return;
  }

  const std::uint32_t r_min = std::min(lhs.row, rhs.row);
  const std::uint32_t r_max = std::max(lhs.row, rhs.row);
  const std::uint32_t c_min = std::min(lhs.col, rhs.col);
  const std::uint32_t c_max = std::max(lhs.col, rhs.col);

  for (std::uint32_t r = r_min; r <= r_max; ++r) {
    for (std::uint32_t c = c_min; c <= c_max; ++c) {
      add_cell_dep(state, CellNodeId{target_sheet_id, r, c});
    }
  }
}

// Forward decl for the recursive walker.
void walk(const parser::AstNode& node, WalkState& state);

// Handles `NodeKind::RangeOp` specifically so the inner Ref endpoints are
// not double-counted as scalar reads.
void walk_range_op(const parser::AstNode& node, WalkState& state) {
  const parser::AstNode& lhs = node.as_range_lhs();
  const parser::AstNode& rhs = node.as_range_rhs();

  // Static dep extraction only handles `Ref:Ref` rectangles. Anything more
  // exotic (OFFSET-based, INDIRECT, named-range expansion) is dynamic and
  // skipped silently here; the evaluator will resolve it lazily and the
  // recalc engine will pick up dependencies when those refs surface as
  // direct `Ref` nodes inside their argument trees.
  if (lhs.kind() == parser::NodeKind::Ref && rhs.kind() == parser::NodeKind::Ref) {
    emit_range_cells(state, lhs.as_ref(), rhs.as_ref());
    return;
  }

  // Otherwise descend into both sides so any nested `Ref` / `Call` is still
  // visited.
  walk(lhs, state);
  walk(rhs, state);
}

void walk(const parser::AstNode& node, WalkState& state) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::ErrorPlaceholder:
      return;

    case parser::NodeKind::Ref: {
      const parser::Reference& ref = node.as_ref();
      // Whole-column / whole-row used as a bare scalar (rare in practice but
      // legal): treat as volatile, do not enumerate cells.
      if (ref.is_full_col || ref.is_full_row) {
        state.out->is_volatile = true;
        return;
      }
      if (ref.row >= Sheet::kMaxRows || ref.col >= Sheet::kMaxCols) {
        return;  // Out-of-bounds: evaluator surfaces #REF! at eval time.
      }
      std::uint16_t sheet_id = state.current_sheet_id;
      if (!resolve_sheet_id(ref, state, &sheet_id)) {
        return;  // Unknown sheet: skip silently.
      }
      add_cell_dep(state, CellNodeId{sheet_id, ref.row, ref.col});
      return;
    }

    case parser::NodeKind::SpillRef: {
      // `A1#` semantically reads the spill region anchored at A1. The set
      // of cells in the region is dynamic (depends on the spill's current
      // shape), so we conservatively register only the anchor as a direct
      // dep: any change at the anchor invalidates the whole region by
      // construction in `Sheet::clear_spill`.
      const parser::Reference& ref = node.as_spill_ref();
      if (ref.is_full_col || ref.is_full_row) {
        return;  // Parser rejects this shape; defensive.
      }
      if (ref.row >= Sheet::kMaxRows || ref.col >= Sheet::kMaxCols) {
        return;
      }
      std::uint16_t sheet_id = state.current_sheet_id;
      if (!resolve_sheet_id(ref, state, &sheet_id)) {
        return;
      }
      add_cell_dep(state, CellNodeId{sheet_id, ref.row, ref.col});
      return;
    }

    case parser::NodeKind::ExternalRef:
      // TODO: cross-workbook ref tracking once the workbook registry exists.
      return;

    case parser::NodeKind::StructuredRef:
      // TODO: structured (table) ref tracking once tables are wired in.
      return;

    case parser::NodeKind::NameRef:
      // TODO: defined-name ref tracking once the defined-name store exists.
      return;

    case parser::NodeKind::UnaryOp:
      walk(node.as_unary_operand(), state);
      return;

    case parser::NodeKind::BinaryOp:
      walk(node.as_binary_lhs(), state);
      walk(node.as_binary_rhs(), state);
      return;

    case parser::NodeKind::RangeOp:
      walk_range_op(node, state);
      return;

    case parser::NodeKind::UnionOp: {
      const std::uint32_t arity = node.as_union_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        walk(node.as_union_child(i), state);
      }
      return;
    }

    case parser::NodeKind::IntersectOp:
      walk(node.as_intersect_lhs(), state);
      walk(node.as_intersect_rhs(), state);
      return;

    case parser::NodeKind::ImplicitIntersection:
      walk(node.as_implicit_intersection_operand(), state);
      return;

    case parser::NodeKind::Call: {
      // Volatile detection: function names are stored UPPERCASE by the
      // parser-side intern table, matching the registry key convention.
      // `is_volatile_function` is case-sensitive on uppercase input, so we
      // can pass the lexeme through verbatim.
      if (VolatileTracker::is_volatile_function(node.as_call_name())) {
        state.out->is_volatile = true;
      }
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        walk(node.as_call_arg(i), state);
      }
      return;
    }

    case parser::NodeKind::ArrayLiteral: {
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      for (std::uint32_t r = 0; r < rows; ++r) {
        for (std::uint32_t c = 0; c < cols; ++c) {
          walk(node.as_array_element(r, c), state);
        }
      }
      return;
    }

    case parser::NodeKind::Lambda:
      // LAMBDA bodies are evaluated at bind time with a parameter env; the
      // body's references to lambda parameters cannot be statically
      // distinguished from workbook refs without simulating the binding.
      // Skip the body so we do not over-approximate.
      return;

    case parser::NodeKind::LetBinding: {
      // LET introduces local names that shadow workbook references inside
      // its body. Walking the binding initialisers is straightforward (they
      // live in the outer scope). For the body, descending unconditionally
      // is safe today because the `NameRef` case is a no-op pending
      // defined-name support: a LET-bound name reaches `NameRef` and
      // contributes nothing, while real cell `Ref` nodes and `Call` nodes
      // inside the body emit their deps and volatile flags as usual.
      // Skipping the body would under-approximate — `=LET(x, A1, x + B1)`
      // would miss `B1`, and `=LET(x, 1, x + RAND())` would miss the
      // volatile call. When defined-name tracking lands, this case will
      // need a name-env stack so bound names short-circuit before reaching
      // the `NameRef` resolver.
      const std::uint32_t binding_count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < binding_count; ++i) {
        walk(node.as_let_binding_expr(i), state);
      }
      walk(node.as_let_body(), state);
      return;
    }

    case parser::NodeKind::LambdaCall: {
      // Walk the callee and arguments so any embedded refs / volatile calls
      // are recorded.
      walk(node.as_lambda_call_callee(), state);
      const std::uint32_t arity = node.as_lambda_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        walk(node.as_lambda_call_arg(i), state);
      }
      return;
    }
  }
}

}  // namespace

ExtractedDeps extract_deps(const parser::AstNode& node, std::uint16_t current_sheet_id, const Workbook& workbook) {
  ExtractedDeps deps;
  WalkState state{&deps, {}, current_sheet_id, &workbook};
  walk(node, state);
  return deps;
}

}  // namespace formulon::eval
