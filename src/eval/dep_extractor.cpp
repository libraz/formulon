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

#include "eval/declared_rect.h"
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
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
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
  struct LexicalBinding {
    std::string name;
    const parser::AstNode* lambda = nullptr;
  };

  ExtractedDeps* out;
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  std::uint16_t current_sheet_id;
  const Workbook* workbook;
  Arena* name_arena;
  std::vector<std::string> name_stack;
  std::vector<LexicalBinding> lexical_stack;
  std::vector<const parser::AstNode*> lambda_stack;
};

const WalkState::LexicalBinding* lookup_lexical(std::string_view name, const WalkState& state) {
  const std::string lowered = strings::to_ascii_lower(name);
  for (auto it = state.lexical_stack.rbegin(); it != state.lexical_stack.rend(); ++it) {
    if (it->name == lowered) {
      return &*it;
    }
  }
  return nullptr;
}

bool lambda_is_active(const parser::AstNode* lambda, const WalkState& state) noexcept {
  for (const parser::AstNode* active : state.lambda_stack) {
    if (active == lambda) {
      return true;
    }
  }
  return false;
}

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
  // `idx` is below `sheet_count()`, which `Workbook::kMaxSheets` bounds at
  // the 16-bit ceiling `CellNodeId::sheet_id` can address, so the narrowing
  // keeps the index intact.
  *out_sheet_id = static_cast<std::uint16_t>(idx);
  return true;
}

// Adds a single cell to the dependency set, deduplicating against `seen`.
void add_cell_dep(WalkState& state, CellNodeId cell) {
  if (state.seen.insert(cell).second) {
    state.out->cell_deps.push_back(cell);
  }
}

void add_range_dep(WalkState& state, CellRangeDependency range) {
  for (const CellRangeDependency& existing : state.out->range_deps) {
    if (existing.sheet_id == range.sheet_id && existing.row_first == range.row_first &&
        existing.row_last == range.row_last && existing.col_first == range.col_first &&
        existing.col_last == range.col_last) {
      return;
    }
  }
  state.out->range_deps.push_back(range);
}

void add_three_d_span_dep(WalkState& state, std::size_t sheet_first, std::size_t sheet_last) {
  if (sheet_first > 0xFFFFU || sheet_last > 0xFFFFU) {
    return;
  }
  const ThreeDSheetSpanDependency span{static_cast<std::uint16_t>(sheet_first), static_cast<std::uint16_t>(sheet_last)};
  if (std::find(state.out->three_d_spans.begin(), state.out->three_d_spans.end(), span) ==
      state.out->three_d_spans.end()) {
    state.out->three_d_spans.push_back(span);
  }
}

// Flattens the rectangle [lhs, rhs] into per-cell dependencies. Both
// endpoints must be plain `Ref` nodes; complex ranges (OFFSET-based,
// INDIRECT, etc.) are silently ignored — dynamic shapes cannot be statically
// resolved here. Whole-column / whole-row endpoints are captured as compact
// rectangles without flattening their 1,048,576 potential cells.
void emit_range_cells(WalkState& state, const parser::Reference& lhs, const parser::Reference& rhs) {
  // Resolve the effective sheet for the rectangle. Mirrors the policy in
  // `EvalContext::expand_range`: the parser keeps the qualifier on the LHS
  // for `Sheet2!A1:B2`, so the RHS often has an empty qualifier and inherits.
  std::string_view effective_sheet;
  if (!lhs.sheet.empty() && !rhs.sheet.empty()) {
    // Mismatched qualifiers are an evaluator-level #REF!; statically we
    // simply skip the range.
    if (!sheet_names::equal(lhs.sheet, rhs.sheet)) {
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

  const auto normalized = declared_rect(lhs, rhs);
  if (!normalized) {
    // The evaluator reports malformed or mixed-axis references as an error;
    // static extraction simply omits a shape it cannot model safely.
    return;
  }
  const std::uint32_t r_min = normalized.value().row_first;
  const std::uint32_t r_max = normalized.value().row_last;
  const std::uint32_t c_min = normalized.value().col_first;
  const std::uint32_t c_max = normalized.value().col_last;

  // Every rectangle leaves here as either at most
  // `kMaxMaterializedDependencyCells` per-cell edges or exactly one compact
  // rectangle; no shape registers nothing.
  const std::uint64_t area = static_cast<std::uint64_t>(r_max - r_min + 1U) * (c_max - c_min + 1U);
  if (normalized.value().whole_axis || area > kMaxMaterializedDependencyCells) {
    add_range_dep(state, CellRangeDependency{target_sheet_id, r_min, r_max, c_min, c_max});
    return;
  }

  for (std::uint32_t r = r_min; r <= r_max; ++r) {
    for (std::uint32_t c = c_min; c <= c_max; ++c) {
      add_cell_dep(state, CellNodeId{target_sheet_id, r, c});
    }
  }
}

// Forward decl for the recursive walker.
void walk(const parser::AstNode& node, WalkState& state);

// Walks a LAMBDA body as an *invoked* body. Parameter names are pushed
// onto the shadow stack for the duration so a parameter that collides
// with a workbook-scoped defined name short-circuits the `NameRef`
// expander instead of dragging that name's dependencies in.
//
// Only call this where the body is genuinely evaluated. A bare lambda
// *value* is not (see the `Lambda` case in `walk()`), and walking it
// would invent dependencies the formula never reads.
void walk_invoked_lambda_body(const parser::AstNode& lambda, WalkState& state) {
  if (lambda_is_active(&lambda, state)) {
    return;
  }
  state.lambda_stack.push_back(&lambda);
  const std::uint32_t param_count = lambda.as_lambda_param_count();
  for (std::uint32_t i = 0; i < param_count; ++i) {
    state.lexical_stack.push_back({strings::to_ascii_lower(lambda.as_lambda_param(i)), nullptr});
  }
  walk(lambda.as_lambda_body(), state);
  for (std::uint32_t i = 0; i < param_count; ++i) {
    state.lexical_stack.pop_back();
  }
  state.lambda_stack.pop_back();
}

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
void expand_defined_name(const io::DefinedName& def, WalkState& state, bool invoked) {
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

  // Defined-name evaluation clears the caller's lexical NameEnv. Keep the
  // extractor's lexical shadow stack isolated for the same reason: a LET
  // binding at the call site must not hide a workbook name referenced by the
  // definition body.
  std::vector<WalkState::LexicalBinding> saved_lexical;
  saved_lexical.swap(state.lexical_stack);
  state.name_stack.push_back(std::move(lowered));
  if (invoked && root->kind() == parser::NodeKind::Lambda) {
    // A direct defined-name LAMBDA body is evaluated by a call, so descend
    // into it with parameter names shadowing workbook names. A bare NameRef
    // to the same definition remains a lambda value and must not invent
    // dependencies from its body.
    walk_invoked_lambda_body(*root, state);
  } else if (invoked && root->kind() == parser::NodeKind::NameRef) {
    // Preserve the common alias shape (`Alias = NamedLambda`) without
    // repeatedly walking the lambda body. Any non-defined alias is handled
    // by the ordinary NameRef walker and contributes no static deps.
    const io::DefinedName* aliased = find_defined_name(*state.workbook, state.current_sheet_id, root->as_name());
    if (aliased != nullptr) {
      expand_defined_name(*aliased, state, /*invoked=*/true);
    }
  } else {
    walk(*root, state);
  }
  state.name_stack.pop_back();
  state.lexical_stack.swap(saved_lexical);
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

  // `Workbook::kMaxSheets` keeps every real sheet index inside the 16-bit
  // `CellNodeId::sheet_id`. `rect` arrives from a resolved defined name, so
  // the bound is re-checked rather than assumed: a rectangle that cannot be
  // addressed contributes no edges instead of aliasing another sheet.
  if (rect.sheet_index > 0xFFFFu) {
    return;
  }
  const std::uint16_t target_sheet_id = static_cast<std::uint16_t>(rect.sheet_index);

  if (rect.row_first >= Sheet::kMaxRows || rect.row_last >= Sheet::kMaxRows ||  //
      rect.col_first >= Sheet::kMaxCols || rect.col_last >= Sheet::kMaxCols) {
    return;
  }

  // A table column is unbounded from the extractor's point of view — the
  // table's own row count decides the area — so the rectangle passes the same
  // graph-footprint ceiling every other range shape does.
  const std::uint64_t area = (static_cast<std::uint64_t>(rect.row_last - rect.row_first) + 1U) *
                             (static_cast<std::uint64_t>(rect.col_last - rect.col_first) + 1U);
  if (area > kMaxMaterializedDependencyCells) {
    add_range_dep(state,
                  CellRangeDependency{target_sheet_id, rect.row_first, rect.row_last, rect.col_first, rect.col_last});
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
      const auto normalized = declared_rect(ref, ref);
      if (!normalized) {
        return;
      }
      std::uint16_t sheet_id = state.current_sheet_id;
      if (!resolve_sheet_id(ref, state, &sheet_id)) {
        return;  // Unknown sheet: skip silently.
      }
      if (normalized.value().whole_axis) {
        add_range_dep(state, CellRangeDependency{sheet_id, normalized.value().row_first, normalized.value().row_last,
                                                 normalized.value().col_first, normalized.value().col_last});
        return;
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
      if (const parser::AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        // A computed anchor names no cell until the formula runs, so there
        // is nothing static to register for the region itself. Walking the
        // sub-expression still records the references it reads and any
        // volatile call it makes, which is what keeps
        // `=SUM(OFFSET(A1,1,0)#)` recalculating.
        walk(*anchor, state);
        return;
      }
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

    case parser::NodeKind::Ref3D: {
      // A 3-D reference reads the same cell (or cell rectangle) on every
      // sheet in the inclusive workbook-order span. Register a cell dep per
      // (sheet, area cell) so an edit to any of them invalidates this
      // formula.
      const parser::Reference& cell = node.as_ref3d_cell();
      const bool is_range = node.as_ref3d_is_range();
      const parser::Reference& cell_end = node.as_ref3d_cell_end();
      const auto normalized = declared_rect(cell, is_range ? cell_end : cell);
      if (!normalized) {
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
      add_three_d_span_dep(state, lo, hi);
      // The shared rectangle is read once per sheet in the span, so the
      // graph-footprint ceiling is applied to it before the span multiplies
      // the cost.
      const std::uint64_t area =
          (static_cast<std::uint64_t>(normalized.value().row_last - normalized.value().row_first) + 1U) *
          (static_cast<std::uint64_t>(normalized.value().col_last - normalized.value().col_first) + 1U);
      const bool compact = normalized.value().whole_axis || area > kMaxMaterializedDependencyCells;
      for (std::size_t s = lo; s <= hi; ++s) {
        if (compact) {
          add_range_dep(state, CellRangeDependency{static_cast<std::uint16_t>(s), normalized.value().row_first,
                                                   normalized.value().row_last, normalized.value().col_first,
                                                   normalized.value().col_last});
          continue;
        }
        for (std::uint32_t r = normalized.value().row_first; r <= normalized.value().row_last; ++r) {
          for (std::uint32_t c = normalized.value().col_first; c <= normalized.value().col_last; ++c) {
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
      // Lexical LET / LAMBDA bindings shadow workbook names. The binding
      // itself is already accounted for by its initializer (or by the
      // invoked lambda body), so a bare lexical NameRef contributes no
      // additional workbook dependency.
      if (lookup_lexical(node.as_name(), state) != nullptr) {
        return;
      }
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
      expand_defined_name(*def, state, /*invoked=*/false);
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
      const std::uint32_t arity = node.as_call_arity();
      for (std::uint32_t i = 0; i < arity; ++i) {
        walk(node.as_call_arg(i), state);
      }

      const WalkState::LexicalBinding* lexical = lookup_lexical(node.as_call_name(), state);
      const io::DefinedName* defined =
          lexical == nullptr ? find_defined_name(*state.workbook, state.current_sheet_id, node.as_call_name())
                             : nullptr;
      // A lexical binding or a visible defined name shadows a built-in name;
      // in particular, a LET-bound `NOW` / `RAND` must not be marked
      // volatile merely because its spelling resembles a built-in.
      if (lexical == nullptr && defined == nullptr && VolatileTracker::is_volatile_function(node.as_call_name())) {
        state.out->is_volatile = true;
        if (VolatileTracker::is_dynamic_reference_function(node.as_call_name())) {
          state.out->has_dynamic_reference = true;
        }
      }
      if (lexical != nullptr) {
        if (lexical->lambda != nullptr) {
          walk_invoked_lambda_body(*lexical->lambda, state);
        }
        return;
      }
      if (defined != nullptr) {
        // `Name(args)` is the only call shape for a workbook-defined
        // LAMBDA. Expanding its body once is enough: recursive re-entry is
        // blocked by `name_stack` while its arguments remain ordinary deps.
        expand_defined_name(*defined, state, /*invoked=*/true);
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
      const std::size_t saved_lexical_depth = state.lexical_stack.size();
      for (std::uint32_t i = 0; i < binding_count; ++i) {
        walk(node.as_let_binding_expr(i), state);
        const parser::AstNode& expr = node.as_let_binding_expr(i);
        const parser::AstNode* lambda = nullptr;
        if (expr.kind() == parser::NodeKind::Lambda) {
          lambda = &expr;
        } else if (expr.kind() == parser::NodeKind::NameRef) {
          if (const WalkState::LexicalBinding* prior = lookup_lexical(expr.as_name(), state); prior != nullptr) {
            lambda = prior->lambda;
          }
        }
        state.lexical_stack.push_back({strings::to_ascii_lower(node.as_let_binding_name(i)), lambda});
      }
      walk(node.as_let_body(), state);
      state.lexical_stack.resize(saved_lexical_depth);
      return;
    }

    case parser::NodeKind::LambdaCall: {
      const parser::AstNode& callee = node.as_lambda_call_callee();
      if (callee.kind() == parser::NodeKind::Lambda) {
        // Directly-invoked lambda (`=LAMBDA(x, x+A1)(5)`): unlike a bare
        // lambda *value*, the body IS evaluated here, so its cell refs and
        // volatile calls are genuine dependencies that must reach the graph.
        walk_invoked_lambda_body(callee, state);
      } else {
        // The only callee kind that reaches here is a nested `LambdaCall`
        // (currying, e.g. `LAMBDA(x, LAMBDA(y, x+y))(3)(4)`): the parser
        // gates this postfix `(` to a `Lambda` or `LambdaCall` LHS, so a
        // named callee (`=MyLambda(5)`) always parses as a `Call` and is
        // handled by the `Call` case above instead, including invocation
        // of a workbook-scoped defined name that holds a lambda. Walking
        // generically here recurses back into this same `case` (or into
        // `Lambda`, at the base of the curry chain), which is what still
        // surfaces the chain's embedded refs and volatile calls.
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
  WalkState state{&deps, {}, current_sheet_id, &workbook, &name_arena, {}, {}, {}};
  walk(node, state);
  return deps;
}

}  // namespace formulon::eval
