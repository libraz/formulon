// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `extract_deps`. See `dep_extractor.h` for the contract.

#include "eval/dep_extractor.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "eval/defined_name_resolve.h"
#include "eval/dep_graph.h"
#include "eval/formula_text_utils.h"
#include "eval/structured_ref.h"
#include "eval/volatile_tracker.h"
#include "io/defined_names.h"
#include "io/tables_reader.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// Walker state: collects results into `out` and tracks already-emitted cells
// in `seen` so the dedup runs in one O(N) pass. `name_stack` holds the
// lowercase names currently being expanded so a self-referential or mutually
// recursive defined name (`Loop = Loop + 1`) is broken silently rather than
// driving the walker into unbounded recursion. `arena` owns the parsed ASTs
// produced for defined-name expansion; it is local to a single
// `extract_deps()` invocation and never escapes.
struct WalkState {
  ExtractedDeps* out;
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  std::unordered_set<std::uint32_t> seen_external_books;
  std::uint16_t current_sheet_id;
  const Workbook* workbook;
  Arena* name_arena;
  std::vector<std::string> name_stack;
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

// Expands a defined-name reference: parses its formula text in the
// extractor-local arena and recurses into the resulting AST through the
// shared `walk()`. Cycles (`Loop = Loop + 1`, `A = B; B = A`) are detected
// via `state.name_stack`: the lowercase form of every name currently being
// expanded is pushed before recursion and popped after. A repeated entry
// causes a silent skip — no volatility flag, no diagnostic — matching the
// "graceful skip on unresolvable refs" policy already used for unknown
// sheets and out-of-bounds coordinates. Parser failures on the defined-
// name's formula are also silently skipped: malformed defined-name
// formulas exist in the wild and the dep extractor is not the right layer
// to surface them.
void expand_defined_name(const io::DefinedName& def, WalkState& state) {
  // Cycle detection. Lowercase the name for the stack comparison so a
  // mixed-case re-entry (`Foo` -> `=foo+1`) is still caught.
  std::string lowered = strings::to_ascii_lower(def.name);
  for (const auto& active : state.name_stack) {
    if (active == lowered) {
      return;
    }
  }

  // Strip the leading `=` if present; defined-name formulas are stored as
  // expressions but Excel sometimes keeps the prefix on import.
  const std::string_view src = strip_formula_prefix(def.formula);
  if (src.empty()) {
    return;
  }

  parser::AstNode* root = parser::parse_strict(src, *state.name_arena);
  if (root == nullptr) {
    return;  // Unparseable (or valid-prefix-plus-garbage) formula: skip.
  }

  state.name_stack.push_back(std::move(lowered));
  walk(*root, state);
  state.name_stack.pop_back();
}

// Resolves a `StructuredRef` node into a static rectangle on the table's
// owning sheet and emits one cell dep per cell in the rectangle.
//
// Design (pin-the-rect, mirroring `walk_range_op`):
//   * The bracket payload is captured verbatim by the parser into the
//     node's `column` slot; we re-parse it through
//     `parse_structured_ref_payload` so multi-specifier (`#All`,
//     `#Headers`, ...) and column-range (`[ColA]:[ColB]`) forms flow
//     through a single resolver.
//   * `parse_structured_ref_payload` failures, missing tables, missing
//     columns, and `#Headers`/`#Totals` on tables that lack the
//     corresponding band all surface as `Expected` errors from
//     `resolve_structured_ref`. Every error path is a silent skip here:
//     the recalc engine cares only about *cells the formula reads*; if
//     the structured ref is unresolvable the evaluator will emit
//     `#NAME?` / `#REF!` at eval time and there are no static deps to
//     register.
//   * Implicit intersection (`Table[@Col]`, the `kThisRow` bit) is
//     statically unresolvable because the row depends on the formula
//     cell's evaluation row context, which `extract_deps` does not have.
//     We skip such references silently — the evaluator will surface the
//     actual single-cell dep when the implicit intersection resolves.
//     This mirrors the concession `walk_range_op` already makes for
//     OFFSET / INDIRECT-shaped endpoints: dynamic shapes cannot be
//     statically pinned.
//   * Table-resize events must trigger a dep re-extraction at the recalc
//     layer; this layer pins the rectangle once and never re-evaluates.
//
// Volatility is not affected: structured refs are not themselves
// volatile. A calculated column's own formula may reference a volatile
// function, but that volatility lives on the column's home cell and is
// the dep extractor's concern when *that* cell is registered, not here.
void walk_structured_ref(const parser::AstNode& node, WalkState& state) {
  const std::string_view table_name = node.as_structured_ref_table();
  const std::string_view payload = node.as_structured_ref_column();

  auto sel_or = parse_structured_ref_payload(payload);
  if (!sel_or) {
    return;  // Malformed payload: silent skip.
  }
  StructuredRefSelector sel = std::move(sel_or).value();
  sel.table_name = table_name;

  // Implicit intersection: row context is the formula cell's row, which
  // is not known here. The evaluator owns this dep at eval time.
  if ((sel.specifiers & StructuredRefSpecifiers::kThisRow) != 0u) {
    return;
  }

  // `resolve_structured_ref` only consults `current_sheet_index` for
  // future cross-sheet contracts; the row argument is consumed only when
  // `kThisRow` is set, which we already short-circuited above. Pass the
  // walk's current sheet for `current_sheet_index` and 0 for the row to
  // keep the call shape stable.
  auto rect_or =
      resolve_structured_ref(sel, *state.workbook, /*current_sheet_index=*/state.current_sheet_id, /*current_row=*/0u);
  if (!rect_or) {
    return;  // Unknown table / column / missing band: silent skip.
  }
  const StructuredRefRange rect = std::move(rect_or).value();

  // Sheet ids fit in uint16_t per CellNodeId; Excel allows at most a few
  // thousand sheets per workbook, well within range. Reject defensively
  // if the workbook's sheet index ever overflows.
  if (rect.sheet_index > 0xFFFFu) {
    return;
  }
  const std::uint16_t target_sheet_id = static_cast<std::uint16_t>(rect.sheet_index);

  if (rect.row_first >= Sheet::kMaxRows || rect.row_last >= Sheet::kMaxRows ||  //
      rect.col_first >= Sheet::kMaxCols || rect.col_last >= Sheet::kMaxCols) {
    return;
  }

  for (auto [r, c] : utils::RectRange(rect.row_first, rect.col_first, rect.row_last, rect.col_last)) {
    add_cell_dep(state, CellNodeId{target_sheet_id, r, c});
  }
}

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

    case parser::NodeKind::ExternalRef: {
      // Opaque-sentinel tracking: capture the referenced book id so the
      // recalc engine can invalidate dependents when the external link's
      // stamp changes. The cross-workbook cell is not flattened into a
      // CellNodeId because external sheets live outside the dep graph
      // (CellNodeId.sheet_id indexes the bound workbook). Today the
      // evaluator returns `#NAME?` for ExternalRef so this is forward-
      // compatible plumbing; consumer wiring is a separate task.
      const std::uint32_t book_id = node.as_external_ref_book_id();
      if (state.seen_external_books.insert(book_id).second) {
        state.out->external_book_ids.push_back(book_id);
      }
      return;
    }

    case parser::NodeKind::Ref3D: {
      // A 3-D reference reads the same cell (or cell rectangle) on every
      // sheet in the inclusive workbook-order span. Register a cell dep per
      // (sheet, area cell) so an edit to any of them invalidates this
      // formula.
      const parser::Reference& cell = node.as_ref3d_cell();
      const bool is_range = node.as_ref3d_is_range();
      const parser::Reference& cell_end = node.as_ref3d_cell_end();
      if (cell.is_full_col || cell.is_full_row || cell.row >= Sheet::kMaxRows || cell.col >= Sheet::kMaxCols ||
          (is_range && (cell_end.is_full_col || cell_end.is_full_row || cell_end.row >= Sheet::kMaxRows ||
                        cell_end.col >= Sheet::kMaxCols))) {
        return;
      }
      if (state.workbook == nullptr) {
        return;
      }
      const std::size_t begin_idx = state.workbook->sheet_index_by_name(node.as_ref3d_sheet_begin());
      const std::size_t end_idx = state.workbook->sheet_index_by_name(node.as_ref3d_sheet_end());
      if (begin_idx == static_cast<std::size_t>(-1) || end_idx == static_cast<std::size_t>(-1)) {
        return;  // Missing endpoint: evaluator surfaces #REF! at eval time.
      }
      const std::size_t lo = std::min(begin_idx, end_idx);
      const std::size_t hi = std::max(begin_idx, end_idx);
      const std::uint32_t r_lo = is_range ? std::min(cell.row, cell_end.row) : cell.row;
      const std::uint32_t r_hi = is_range ? std::max(cell.row, cell_end.row) : cell.row;
      const std::uint32_t c_lo = is_range ? std::min(cell.col, cell_end.col) : cell.col;
      const std::uint32_t c_hi = is_range ? std::max(cell.col, cell_end.col) : cell.col;
      for (std::size_t s = lo; s <= hi; ++s) {
        for (std::uint32_t r = r_lo; r <= r_hi; ++r) {
          for (std::uint32_t c = c_lo; c <= c_hi; ++c) {
            add_cell_dep(state, CellNodeId{static_cast<std::uint16_t>(s), r, c});
          }
        }
      }
      return;
    }

    case parser::NodeKind::StructuredRef:
      walk_structured_ref(node, state);
      return;

    case parser::NodeKind::NameRef: {
      // Resolve against the workbook's defined-name list. A miss (the name
      // is undefined, scoped to a different sheet, or hidden behind a cycle
      // already on the expansion stack) is a silent skip — same policy as
      // unknown sheet qualifiers and out-of-bounds coordinates. On a hit we
      // re-enter `walk()` with the parsed body so cells, ranges, volatility,
      // and nested NameRefs all surface naturally.
      const io::DefinedName* def = find_defined_name(*state.workbook, state.current_sheet_id, node.as_name());
      if (def == nullptr) {
        return;
      }
      expand_defined_name(*def, state);
      return;
    }

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
      // Volatile detection: `is_volatile_function` matches names
      // case-insensitively, so the call lexeme is passed through verbatim
      // (a hand-typed `=now()` is recognised as well as `=NOW()`).
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
      // live in the outer scope). The body is descended unconditionally so
      // that `=LET(x, A1, x + B1)` records B1 and `=LET(x, 1, x + RAND())`
      // records the volatile call. Bound names reach the `NameRef` case;
      // when no workbook-scoped defined name shares the identifier the
      // resolver returns null and the binding contributes nothing. A LET
      // binding whose name *does* collide with a workbook-scoped defined
      // name will currently over-approximate (the defined-name body is
      // walked instead of being shadowed). A scoped name-environment stack
      // that short-circuits the `NameRef` resolver inside LET bodies is the
      // proper fix and is deferred — collisions of this shape are rare in
      // practice and over-approximating cell deps is conservative for the
      // recalc engine.
      const std::uint32_t binding_count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < binding_count; ++i) {
        walk(node.as_let_binding_expr(i), state);
      }
      walk(node.as_let_body(), state);
      return;
    }

    case parser::NodeKind::LambdaCall: {
      const parser::AstNode& callee = node.as_lambda_call_callee();
      if (callee.kind() == parser::NodeKind::Lambda) {
        // Directly-invoked lambda (`=LAMBDA(x, x+A1)(5)`): unlike a bare
        // lambda *value*, the body IS evaluated here, so its cell refs and
        // volatile calls are genuine dependencies that must reach the graph.
        // Walk the body with the parameter names shadowed — pushing them onto
        // `name_stack` makes a parameter that collides with a workbook-scoped
        // defined name short-circuit the `NameRef` expander instead of
        // pulling in that name's dependencies.
        const std::uint32_t param_count = callee.as_lambda_param_count();
        for (std::uint32_t i = 0; i < param_count; ++i) {
          state.name_stack.push_back(strings::to_ascii_lower(callee.as_lambda_param(i)));
        }
        walk(callee.as_lambda_body(), state);
        for (std::uint32_t i = 0; i < param_count; ++i) {
          state.name_stack.pop_back();
        }
      } else {
        // Named / computed callee (e.g. a defined-name lambda): walk it
        // generically so embedded refs and volatile calls still surface.
        walk(callee, state);
      }
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
  // The arena owns any ASTs parsed for defined-name expansion. It lives only
  // as long as this call so the parsed nodes never outlive the walk; the
  // caller-supplied `node` is unrelated and stays in its own arena.
  Arena name_arena;
  WalkState state{&deps, {}, {}, current_sheet_id, &workbook, &name_arena, {}};
  walk(node, state);
  return deps;
}

}  // namespace formulon::eval
