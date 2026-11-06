//
// Implementation of `extract_deps`. See `dep_extractor.h` for the contract.

#include "eval/dep_extractor.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "defined_name.h"
#include "eval/builtin_names.h"
#include "eval/cell_lazy.h"
#include "eval/declared_rect.h"
#include "eval/defined_name_resolve.h"
#include "eval/dep_graph.h"
#include "eval/range_resolvers.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/volatile_tracker.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/rect_iterator.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon::eval {
namespace {

// The rectangle a reference-valued expression can name, before evaluation
// picks which of its candidate references it returns. `sheet_explicit`
// records whether a qualifier fixed the sheet, since a bare `:` endpoint
// inherits the other endpoint's sheet.
struct Footprint {
  std::uint16_t sheet_id = 0;
  bool sheet_explicit = false;
  DeclaredRect rect;
};

// Walker state: collects results into `out` and tracks already-emitted cells
// in `seen` so the dedup runs in one O(N) pass. `name_stack` holds the
// definitions currently being expanded so a self-referential or mutually
// recursive defined name (`Loop = Loop + 1`) is broken silently rather than
// driving the walker into unbounded recursion. `arena` owns the parsed ASTs
// produced for defined-name expansion; it is local to a single
// `extract_deps()` invocation and never escapes.
struct WalkState {
  struct LexicalBinding {
    // AST names are interned in the arena that owns the formula (or the
    // extractor-local arena for a defined-name body), so the binding can
    // borrow the view for the duration of this walk.
    std::string_view name;
    const parser::AstNode* lambda = nullptr;
    /// The reference a LET initialiser can name; empty for LAMBDA parameters
    /// and for initialisers that are not references.
    std::optional<Footprint> footprint;
    /// A reference written in the formula that this name binds, left unwalked:
    /// a value use of the name walks it, a position-only use (`ROW(r)`) does
    /// not, exactly as for the reference written in place.
    const parser::AstNode* deferred = nullptr;
  };

  ExtractedDeps* out;
  std::unordered_set<CellNodeId, CellNodeIdHash> seen;
  std::uint16_t current_sheet_id;
  /// The formula cell's row, which a `[#This Row]` reference reads.
  std::uint32_t current_row;
  // Scope unqualified defined names resolve in: the formula's own sheet,
  // or the owning sheet while a sheet-local name's body is expanded.
  std::uint16_t name_scope_sheet_id;
  const Workbook* workbook;
  Arena* name_arena;
  std::vector<const DefinedName*> name_stack;
  std::vector<LexicalBinding> lexical_stack;
  std::vector<const parser::AstNode*> lambda_stack;
  /// Set just before walking a call that sits in a position-only slot, and
  /// consumed by that call: the reference it returns is only positioned too.
  bool position_only_call = false;
};

const WalkState::LexicalBinding* lookup_lexical(std::string_view name, const WalkState& state) {
  for (auto it = state.lexical_stack.rbegin(); it != state.lexical_stack.rend(); ++it) {
    if (strings::case_insensitive_eq(it->name, name)) {
      return &*it;
    }
  }
  return nullptr;
}

// The lexical binding an unqualified `NameRef` denotes, or nullptr for any
// other node.
const WalkState::LexicalBinding* lookup_unqualified_lexical(const parser::AstNode& node, const WalkState& state) {
  if (node.kind() == parser::NodeKind::NameRef && node.as_name_sheet().empty()) {
    return lookup_lexical(node.as_name(), state);
  }
  return nullptr;
}

// True for `Name`, `Sheet1!Name` and `[0]!Name`.
bool is_defined_name_ref(const parser::AstNode& node) {
  return node.kind() == parser::NodeKind::NameRef || parser::is_self_book_name_ref(node);
}

// The definition a `NameRef` denotes: `Sheet1!Name` in Sheet1's scope, an
// unqualified name in the formula's own.
const DefinedName* find_name_ref_definition(const parser::AstNode& name_ref, const WalkState& state) {
  if (name_ref.kind() == parser::NodeKind::ExternalRef) {
    // Only `[0]!Name` names a definition of this workbook.
    return parser::is_self_book_name_ref(name_ref)
               ? find_self_book_defined_name(*state.workbook, name_ref.as_external_ref_name())
               : nullptr;
  }
  const std::string_view sheet = name_ref.as_name_sheet();
  if (!sheet.empty()) {
    return find_sheet_defined_name(*state.workbook, sheet, name_ref.as_name());
  }
  return find_defined_name(*state.workbook, state.name_scope_sheet_id, name_ref.as_name());
}

bool lambda_is_active(const parser::AstNode* lambda, const WalkState& state) noexcept {
  for (const parser::AstNode* active : state.lambda_stack) {
    if (active == lambda) {
      return true;
    }
  }
  return false;
}

// Looks up a sheet by display name; false when no sheet matches.
bool find_sheet_index(const Workbook& workbook, std::string_view name, std::size_t* out) {
  const std::size_t idx = workbook.sheet_index_by_name(name);
  if (idx >= workbook.sheet_count()) {
    return false;
  }
  *out = idx;
  return true;
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
  std::size_t idx = 0;
  if (!find_sheet_index(*state.workbook, ref.sheet, &idx)) {
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
  if (std::find(state.out->range_deps.begin(), state.out->range_deps.end(), range) == state.out->range_deps.end()) {
    state.out->range_deps.push_back(range);
  }
}

void add_three_d_span_dep(WalkState& state, std::size_t sheet_first, std::size_t sheet_last) {
  if (sheet_first >= kRangeNodeSheetId || sheet_last >= kRangeNodeSheetId) {
    return;
  }
  const ThreeDSheetSpanDependency span{static_cast<std::uint16_t>(sheet_first), static_cast<std::uint16_t>(sheet_last)};
  if (std::find(state.out->three_d_spans.begin(), state.out->three_d_spans.end(), span) ==
      state.out->three_d_spans.end()) {
    state.out->three_d_spans.push_back(span);
  }
}

// Registers `rect` on `sheet_id`. Every rectangle leaves here as either at
// most `kMaxMaterializedDependencyCells` per-cell edges or exactly one
// compact rectangle; no shape registers nothing.
void emit_rect(WalkState& state, std::uint16_t sheet_id, const DeclaredRect& rect) {
  const std::uint64_t area = static_cast<std::uint64_t>(rect.rows()) * rect.cols();
  if (rect.whole_axis || area > kMaxMaterializedDependencyCells) {
    add_range_dep(state, CellRangeDependency{sheet_id, rect.row_first, rect.row_last, rect.col_first, rect.col_last});
    return;
  }
  for (auto [r, c] : utils::RectRange(rect.row_first, rect.col_first, rect.row_last, rect.col_last)) {
    add_cell_dep(state, CellNodeId{sheet_id, r, c});
  }
}

// The footprint of a static endpoint pair. A bare side takes the other side's
// sheet; two qualified sides must agree, else the pair has no footprint.
std::optional<Footprint> pair_footprint(const parser::Reference& lhs, const parser::Reference& rhs,
                                        const WalkState& state) {
  if (!lhs.sheet.empty() && !rhs.sheet.empty() && !sheet_names::equal(lhs.sheet, rhs.sheet)) {
    return std::nullopt;
  }
  parser::Reference probe{};
  probe.sheet = lhs.sheet.empty() ? rhs.sheet : lhs.sheet;
  Footprint out;
  if (!resolve_sheet_id(probe, state, &out.sheet_id)) {
    return std::nullopt;
  }
  const auto rect = declared_rect(lhs, rhs);
  if (!rect) {
    return std::nullopt;
  }
  out.sheet_explicit = !probe.sheet.empty();
  out.rect = rect.value();
  return out;
}

// Flattens the rectangle [lhs, rhs] into per-cell dependencies. Both
// endpoints must be plain `Ref` nodes; complex ranges (OFFSET-based,
// INDIRECT, etc.) are silently ignored — dynamic shapes cannot be statically
// resolved here. Whole-column / whole-row endpoints are captured as compact
// rectangles without flattening their 1,048,576 potential cells.
void emit_range_cells(WalkState& state, const parser::Reference& lhs, const parser::Reference& rhs) {
  // Mismatched qualifiers are an evaluator-level #REF!, and a malformed or
  // mixed-axis pair is an evaluator error; statically both are omitted.
  if (const std::optional<Footprint> footprint = pair_footprint(lhs, rhs, state)) {
    emit_rect(state, footprint->sheet_id, footprint->rect);
  }
}

// Forward decl for the recursive walker.
void walk(const parser::AstNode& node, WalkState& state);

// Per-parameter reference arguments a lambda invocation binds unwalked.
using DeferredArgs = std::vector<const parser::AstNode*>;

// Walks, as value reads, the deferred arguments from index `first` on.
void walk_deferred_args(const DeferredArgs* args, std::size_t first, WalkState& state) {
  if (args == nullptr) {
    return;
  }
  for (std::size_t i = first; i < args->size(); ++i) {
    if ((*args)[i] != nullptr) {
      walk(*(*args)[i], state);
    }
  }
}

// Walks a LAMBDA body as an *invoked* body. Parameter names are pushed
// onto the shadow stack for the duration so a parameter that collides
// with a workbook-scoped defined name short-circuits the `NameRef`
// expander instead of dragging that name's dependencies in.
//
// Only call this where the body is genuinely evaluated. A bare lambda
// *value* is not (see the `Lambda` case in `walk()`), and walking it
// would invent dependencies the formula never reads.
//
// `args` holds, per parameter, the deferred reference argument it binds
// (nullptr for an argument the caller walked itself). A deferred argument no
// parameter takes -- the body is already being walked, or the call passes too
// many -- is walked here as a value read.
void walk_invoked_lambda_body(const parser::AstNode& lambda, WalkState& state, const DeferredArgs* args) {
  if (lambda_is_active(&lambda, state)) {
    walk_deferred_args(args, 0U, state);
    return;
  }
  state.lambda_stack.push_back(&lambda);
  const std::uint32_t param_count = lambda.as_lambda_param_count();
  for (std::uint32_t i = 0; i < param_count; ++i) {
    const parser::AstNode* deferred = (args != nullptr && i < args->size()) ? (*args)[i] : nullptr;
    state.lexical_stack.push_back({lambda.as_lambda_param(i), nullptr, std::nullopt, deferred});
  }
  walk(lambda.as_lambda_body(), state);
  for (std::uint32_t i = 0; i < param_count; ++i) {
    state.lexical_stack.pop_back();
  }
  state.lambda_stack.pop_back();
  walk_deferred_args(args, param_count, state);
}

// Parses a defined name's formula text in the extractor-local arena and
// hands the resulting AST to `visit`. Cycles (`Loop = Loop + 1`, `A = B; B = A`) are detected
// via `state.name_stack`: every definition currently being expanded is
// pushed before recursion and popped after. A repeated entry
// causes a silent skip — no volatility flag, no diagnostic — matching the
// "graceful skip on unresolvable refs" policy already used for unknown
// sheets and out-of-bounds coordinates. Parser failures on the defined-
// name's formula are also silently skipped: malformed defined-name
// formulas exist in the wild and the dep extractor is not the right layer
// to surface them.
template <typename Visit>
void visit_defined_name_body(const DefinedName& def, WalkState& state, Visit&& visit) {
  // Cycle detection by definition identity: a mixed-case re-entry
  // (`Foo` -> `=foo+1`) resolves to the same entry, while `Sheet1!X` naming
  // `Sheet2!X` reaches a different one.
  for (const DefinedName* active : state.name_stack) {
    if (active == &def) {
      return;
    }
  }

  // Strip the leading `=` if present; defined-name formulas are stored as
  // expressions but Excel sometimes keeps the prefix on import.
  const std::string_view src = strip_formula_prefix(def.formula);
  if (src.empty()) {
    return;
  }

  parser::AstNode* root = parse_formula_entry(src, *state.name_arena);
  if (root == nullptr) {
    return;  // Unparseable (or valid-prefix-plus-garbage) formula: skip.
  }

  // Defined-name evaluation clears the caller's lexical NameEnv. Keep the
  // extractor's lexical shadow stack isolated for the same reason: a LET
  // binding at the call site must not hide a workbook name referenced by the
  // definition body.
  std::vector<WalkState::LexicalBinding> saved_lexical;
  saved_lexical.swap(state.lexical_stack);
  // A sheet-local body binds its names in the owning sheet's scope, as
  // `resolve_defined_name` evaluates it.
  const std::uint16_t saved_scope = state.name_scope_sheet_id;
  if (def.local_sheet_id >= 0 && static_cast<std::size_t>(def.local_sheet_id) < state.workbook->sheet_count()) {
    state.name_scope_sheet_id = static_cast<std::uint16_t>(def.local_sheet_id);
  }
  state.name_stack.push_back(&def);
  visit(*root);
  state.name_stack.pop_back();
  state.name_scope_sheet_id = saved_scope;
  state.lexical_stack.swap(saved_lexical);
}

// Expands a defined-name reference, recursing into its body through the
// shared `walk()`.
void expand_defined_name(const DefinedName& def, WalkState& state, bool invoked, const DeferredArgs* args = nullptr) {
  bool bound = false;
  visit_defined_name_body(def, state, [&](const parser::AstNode& root) {
    if (invoked && root.kind() == parser::NodeKind::Lambda) {
      bound = true;
      // A direct defined-name LAMBDA body is evaluated by a call, so descend
      // into it with parameter names shadowing workbook names. A bare NameRef
      // to the same definition remains a lambda value and must not invent
      // dependencies from its body.
      walk_invoked_lambda_body(root, state, args);
    } else if (invoked && is_defined_name_ref(root)) {
      // Preserve the common alias shape (`Alias = NamedLambda`) without
      // repeatedly walking the lambda body. Any non-defined alias is handled
      // by the ordinary NameRef walker and contributes no static deps.
      const DefinedName* aliased = find_name_ref_definition(root, state);
      if (aliased != nullptr) {
        bound = true;
        expand_defined_name(*aliased, state, /*invoked=*/true, args);
      }
    } else {
      walk(root, state);
    }
  });
  if (!bound) {
    walk_deferred_args(args, 0U, state);
  }
}

// A builtin can use a lambda argument only by calling it (MAP, BYROW, REDUCE,
// GROUPBY, ...), so a lambda written inline, bound by LET or named by a defined
// name is walked as an invoked body. Returns false for any other argument.
bool walk_callable_arg(const parser::AstNode& arg, WalkState& state, const DeferredArgs* args) {
  if (arg.kind() == parser::NodeKind::Lambda) {
    walk_invoked_lambda_body(arg, state, args);
    return true;
  }
  if (const WalkState::LexicalBinding* lexical = lookup_unqualified_lexical(arg, state); lexical != nullptr) {
    if (lexical->lambda == nullptr) {
      return false;
    }
    walk_invoked_lambda_body(*lexical->lambda, state, args);
    return true;
  }
  if (!is_defined_name_ref(arg)) {
    return false;
  }
  const DefinedName* def = find_name_ref_definition(arg, state);
  if (def == nullptr) {
    return false;
  }
  // A definition that is not a LAMBDA walks exactly as a non-invoked one.
  expand_defined_name(*def, state, /*invoked=*/true, args);
  return true;
}

// Resolves a `StructuredRef` node into a static rectangle on the table's
// owning sheet.
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
//   * `[#This Row]` (`Table[@Col]`) pins the formula cell's own row; a
//     formula outside the table's data rows resolves to nothing here and
//     to `#VALUE!` in the evaluator.
//   * Table-resize events must trigger a dep re-extraction at the recalc
//     layer; this layer pins the rectangle once and never re-evaluates.
//
// Volatility is not affected: structured refs are not themselves
// volatile. A calculated column's own formula may reference a volatile
// function, but that volatility lives on the column's home cell and is
// the dep extractor's concern when *that* cell is registered, not here.
bool structured_ref_rect(const parser::AstNode& node, const WalkState& state, std::uint16_t* out_sheet_id,
                         DeclaredRect* out_rect) {
  const std::string_view table_name = node.as_structured_ref_table();
  const std::string_view payload = node.as_structured_ref_column();

  auto sel_or = parse_structured_ref_payload(payload);
  if (!sel_or) {
    return false;  // Malformed payload: silent skip.
  }
  StructuredRefSelector sel = std::move(sel_or).value();
  sel.table_name = table_name;

  auto rect_or = resolve_structured_ref(sel, *state.workbook, state.current_sheet_id, state.current_row);
  if (!rect_or) {
    return false;  // Unknown table / column / missing band: silent skip.
  }
  const StructuredRefRange rect = std::move(rect_or).value();

  // `Workbook::kMaxSheets` keeps every real sheet index inside the 16-bit
  // `CellNodeId::sheet_id`. `rect` arrives from a resolved defined name, so
  // the bound is re-checked rather than assumed: a rectangle that cannot be
  // addressed contributes no edges instead of aliasing another sheet.
  if (rect.sheet_index >= kRangeNodeSheetId) {
    return false;
  }
  if (rect.row_first >= Sheet::kMaxRows || rect.row_last >= Sheet::kMaxRows ||  //
      rect.col_first >= Sheet::kMaxCols || rect.col_last >= Sheet::kMaxCols) {
    return false;
  }
  *out_sheet_id = static_cast<std::uint16_t>(rect.sheet_index);
  *out_rect = DeclaredRect{rect.row_first, rect.row_last, rect.col_first, rect.col_last, /*whole_axis=*/false};
  return true;
}

void walk_structured_ref(const parser::AstNode& node, WalkState& state) {
  std::uint16_t sheet_id = 0;
  DeclaredRect rect;
  if (structured_ref_rect(node, state, &sheet_id, &rect)) {
    // A table column is unbounded from the extractor's point of view — the
    // table's own row count decides the area — so the rectangle passes the
    // same graph-footprint ceiling every other range shape does.
    emit_rect(state, sheet_id, rect);
  }
}

// Unions two footprints the way `:` composes its endpoints: two qualified
// sheets must agree, and a bare side takes the other side's sheet.
std::optional<Footprint> union_footprints(const Footprint& a, const Footprint& b) {
  if (a.sheet_explicit && b.sheet_explicit && a.sheet_id != b.sheet_id) {
    return std::nullopt;
  }
  Footprint out;
  out.sheet_explicit = a.sheet_explicit || b.sheet_explicit;
  out.sheet_id = a.sheet_explicit ? a.sheet_id : b.sheet_id;
  out.rect.row_first = std::min(a.rect.row_first, b.rect.row_first);
  out.rect.row_last = std::max(a.rect.row_last, b.rect.row_last);
  out.rect.col_first = std::min(a.rect.col_first, b.rect.col_first);
  out.rect.col_last = std::max(a.rect.col_last, b.rect.col_last);
  out.rect.whole_axis = a.rect.whole_axis || b.rect.whole_axis;
  return out;
}

// Folds `part` into the running union `acc`; false when the two cannot be joined.
bool fold_footprint(std::optional<Footprint>& acc, const Footprint& part) {
  acc = acc ? union_footprints(*acc, part) : std::optional<Footprint>(part);
  return acc.has_value();
}

std::optional<Footprint> reference_footprint(const parser::AstNode& node, WalkState& state);

// The footprint of a reference-returning call: the union over every argument
// that can be the returned reference. Literal arguments name no reference
// (the `:` fails if one is picked) and are skipped.
std::optional<Footprint> call_footprint(const parser::AstNode& node, WalkState& state) {
  const std::string_view name = node.as_call_name();
  if (lookup_lexical(name, state) != nullptr ||
      find_defined_name(*state.workbook, state.name_scope_sheet_id, name) != nullptr ||
      VolatileTracker::is_dynamic_reference_function(name) || !is_reference_call_name(name)) {
    return std::nullopt;
  }
  const std::uint32_t arity = node.as_call_arity();
  std::optional<Footprint> acc;
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg = node.as_call_arg(i);
    if (!is_reference_carrying_arg(name, i, arity) || arg.kind() == parser::NodeKind::Literal ||
        arg.kind() == parser::NodeKind::ErrorLiteral) {
      continue;
    }
    const std::optional<Footprint> part = reference_footprint(arg, state);
    if (!part) {
      return std::nullopt;
    }
    if (!fold_footprint(acc, *part)) {
      return std::nullopt;
    }
  }
  return acc;
}

// The rectangle `node` can name when used as a reference, or nothing when
// that is not statically knowable (OFFSET / INDIRECT, a LAMBDA parameter, a
// non-reference expression).
std::optional<Footprint> reference_footprint(const parser::AstNode& node, WalkState& state) {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
      return pair_footprint(node.as_ref(), node.as_ref(), state);

    case parser::NodeKind::RangeOp: {
      parser::Reference lhs{};
      parser::Reference rhs{};
      if (declared_rect_endpoint_pair(node, &lhs, &rhs)) {
        return pair_footprint(lhs, rhs, state);
      }
      std::optional<Footprint> acc = reference_footprint(node.as_range_operand(0), state);
      for (std::uint32_t i = 1; acc && i < node.as_range_arity(); ++i) {
        const std::optional<Footprint> part = reference_footprint(node.as_range_operand(i), state);
        acc = part ? union_footprints(*acc, *part) : std::nullopt;
      }
      return acc;
    }

    case parser::NodeKind::UnionOp: {
      std::optional<Footprint> acc;
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        const std::optional<Footprint> part = reference_footprint(node.as_union_child(i), state);
        if (!part) {
          return std::nullopt;
        }
        if (!fold_footprint(acc, *part)) {
          return std::nullopt;
        }
      }
      return acc;
    }

    case parser::NodeKind::StructuredRef: {
      Footprint out;
      if (!structured_ref_rect(node, state, &out.sheet_id, &out.rect)) {
        return std::nullopt;
      }
      out.sheet_explicit = true;
      return out;
    }

    case parser::NodeKind::NameRef:
    case parser::NodeKind::ExternalRef: {
      if (const WalkState::LexicalBinding* lexical = lookup_unqualified_lexical(node, state); lexical != nullptr) {
        return lexical->footprint;
      }
      const DefinedName* def = find_name_ref_definition(node, state);
      if (def == nullptr) {
        return std::nullopt;
      }
      std::optional<Footprint> out;
      visit_defined_name_body(*def, state,
                              [&](const parser::AstNode& root) { out = reference_footprint(root, state); });
      return out;
    }

    case parser::NodeKind::Call:
      return call_footprint(node, state);

    default:
      return std::nullopt;
  }
}

// Whether CELL's info_type is a constant that reads the referenced cell's
// value, which makes the reference a static read.
bool cell_info_is_constant_value_read(const parser::AstNode& call) {
  if (call.as_call_arity() < 1U || call.as_call_arg(0).kind() != parser::NodeKind::Literal) {
    return false;
  }
  const Value& info = call.as_call_arg(0).as_literal();
  return info.is_text() && cell_info_reads_value(info.as_text());
}

// A CELL reference is only positioned unless a constant info_type reads its
// value; a computed info_type's read is reported at run time instead.
bool cell_info_is_positional(const parser::AstNode& call) {
  return !cell_info_is_constant_value_read(call);
}

// True when a builtin call may read a cell no formula text names: OFFSET /
// INDIRECT, and CELL with a computed info_type.
bool reads_cells_at_run_time(const parser::AstNode& call) {
  if (VolatileTracker::is_dynamic_reference_function(call.as_call_name())) {
    return true;
  }
  return strings::case_insensitive_eq(strip_future_prefix(call.as_call_name()), "CELL") && call.as_call_arity() == 2U &&
         call.as_call_arg(0).kind() != parser::NodeKind::Literal;
}

// Builtins that read only the position or shape of a reference argument,
// never its values: a static reference there is no dependency, so `=ROW(A1)`
// in A1 is not circular. OFFSET's base only anchors the rectangle it
// returns, whose cells are recorded when the formula runs. FORMULATEXT /
// ISFORMULA read the cell and stay out.
struct ReferenceOnlyArg {
  std::string_view name;
  std::uint32_t arg_index;
  /// Further condition on the call, or nullptr when the slot always qualifies.
  bool (*applies)(const parser::AstNode& call);
};
constexpr ReferenceOnlyArg kReferenceOnlyArgs[] = {
    {"ROW", 0U, nullptr},     {"COLUMN", 0U, nullptr}, {"ROWS", 0U, nullptr},
    {"COLUMNS", 0U, nullptr}, {"AREAS", 0U, nullptr},  {"ISREF", 0U, nullptr},
    {"SHEET", 0U, nullptr},   {"OFFSET", 0U, nullptr}, {"CELL", 1U, &cell_info_is_positional},
};

bool is_reference_only_arg(const parser::AstNode& call, std::uint32_t arg_index) {
  const std::string_view bare = strip_future_prefix(call.as_call_name());
  return std::any_of(std::begin(kReferenceOnlyArgs), std::end(kReferenceOnlyArgs), [&](const ReferenceOnlyArg& e) {
    return e.arg_index == arg_index && strings::case_insensitive_eq(bare, e.name) &&
           (e.applies == nullptr || e.applies(call));
  });
}

// True for a reference whose rectangle is fixed by the formula text: a `Ref`,
// a static `:` chain, a 3-D or structured reference, a union of those, or a
// defined name standing for one, or a LET / LAMBDA name bound to such a
// reference written in the formula. A spill or a computed endpoint is not.
bool is_static_reference(const parser::AstNode& node, WalkState& state) {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::StructuredRef:
      return true;
    case parser::NodeKind::RangeOp: {
      parser::Reference lhs{};
      parser::Reference rhs{};
      return declared_rect_endpoint_pair(node, &lhs, &rhs);
    }
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (!is_static_reference(node.as_union_child(i), state)) {
          return false;
        }
      }
      return true;
    case parser::NodeKind::NameRef:
    case parser::NodeKind::ExternalRef: {
      if (const WalkState::LexicalBinding* lexical = lookup_unqualified_lexical(node, state); lexical != nullptr) {
        return lexical->deferred != nullptr;
      }
      const DefinedName* def = find_name_ref_definition(node, state);
      if (def == nullptr) {
        return false;
      }
      bool is_static = false;
      visit_defined_name_body(*def, state,
                              [&](const parser::AstNode& root) { is_static = is_static_reference(root, state); });
      return is_static;
    }
    default:
      return false;
  }
}

// The reference written in the formula that binding `node` to a LET name or
// a LAMBDA parameter can leave unwalked: a static reference spelled without
// names, or a name already bound to one. nullptr for anything else, which is
// walked where it is bound.
const parser::AstNode* deferrable_reference(const parser::AstNode& node, WalkState& state) {
  switch (node.kind()) {
    case parser::NodeKind::Ref:
    case parser::NodeKind::Ref3D:
    case parser::NodeKind::StructuredRef:
    case parser::NodeKind::RangeOp:
      return is_static_reference(node, state) ? &node : nullptr;
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (deferrable_reference(node.as_union_child(i), state) != &node.as_union_child(i)) {
          return nullptr;
        }
      }
      return &node;
    case parser::NodeKind::NameRef:
      if (const WalkState::LexicalBinding* lexical = lookup_unqualified_lexical(node, state); lexical != nullptr) {
        return lexical->deferred;
      }
      return nullptr;
    default:
      return nullptr;
  }
}

// The lambda helpers whose last argument is a callback fed element by element
// from source arguments, grouped by how their sources map to callback parameters.
enum class CallbackHelper { None, Map, ByRowOrCol, Fold };

CallbackHelper callback_helper_kind(std::string_view call_name) {
  const std::string_view bare = strip_future_prefix(call_name);
  if (strings::case_insensitive_eq(bare, "MAP")) {
    return CallbackHelper::Map;
  }
  if (strings::case_insensitive_eq(bare, "BYROW") || strings::case_insensitive_eq(bare, "BYCOL")) {
    return CallbackHelper::ByRowOrCol;
  }
  if (strings::case_insensitive_eq(bare, "REDUCE") || strings::case_insensitive_eq(bare, "SCAN")) {
    return CallbackHelper::Fold;
  }
  return CallbackHelper::None;
}

// The source argument whose cells callback parameter `param` of helper
// `kind` receives: MAP's arrays in order, BYROW / BYCOL's array, REDUCE /
// SCAN's array (their accumulator is a result, not a source). -1 for none.
int callback_source_arg(CallbackHelper kind, std::uint32_t arity, std::uint32_t param) {
  if (kind == CallbackHelper::Map) {
    return param + 1U < arity ? static_cast<int>(param) : -1;
  }
  if (kind == CallbackHelper::ByRowOrCol) {
    return (arity == 2U && param == 0U) ? 0 : -1;
  }
  return (arity == 3U && param == 1U) ? 1 : -1;
}

// Handles `NodeKind::RangeOp` specifically so the inner Ref endpoints are
// not double-counted as scalar reads.
void walk_range_op(const parser::AstNode& node, WalkState& state) {
  const parser::AstNode& lhs = node.as_range_operand(0);
  const parser::AstNode& rhs = node.as_range_operand(1);

  if (node.as_range_arity() == 2U && lhs.kind() == parser::NodeKind::Ref && rhs.kind() == parser::NodeKind::Ref) {
    emit_range_cells(state, lhs.as_ref(), rhs.as_ref());
    return;
  }

  // Otherwise descend into every endpoint so any nested `Ref` / `Call` is still
  // visited, then register the bounding box the `:` reads. An endpoint that
  // is a reference-returning call contributes every reference it could
  // return (`A1:INDEX(C1:C10,n)` reads within A1:C10); OFFSET / INDIRECT
  // endpoints have no static footprint and rely on their volatility.
  for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
    walk(node.as_range_operand(i), state);
  }
  if (const std::optional<Footprint> footprint = reference_footprint(node, state)) {
    emit_rect(state, footprint->sheet_id, footprint->rect);
  }
}

void walk_ref(const parser::AstNode& node, WalkState& state) {
  const parser::Reference& ref = node.as_ref();
  emit_range_cells(state, ref, ref);
}

void walk_spill_ref(const parser::AstNode& node, WalkState& state) {
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
}

void walk_ref3d(const parser::AstNode& node, WalkState& state) {
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
  std::size_t begin_idx = 0;
  std::size_t end_idx = 0;
  if (!find_sheet_index(*state.workbook, node.as_ref3d_sheet_begin(), &begin_idx) ||
      !find_sheet_index(*state.workbook, node.as_ref3d_sheet_end(), &end_idx)) {
    return;  // Missing endpoint: evaluator surfaces #REF! at eval time.
  }
  const std::size_t lo = std::min(begin_idx, end_idx);
  const std::size_t hi = std::max(begin_idx, end_idx);
  add_three_d_span_dep(state, lo, hi);
  // The shared rectangle is read once per sheet in the span, so the
  // graph-footprint ceiling is applied to it before the span multiplies
  // the cost.
  for (std::size_t s = lo; s <= hi; ++s) {
    emit_rect(state, static_cast<std::uint16_t>(s), normalized.value());
  }
}

void walk_name_ref(const parser::AstNode& node, WalkState& state) {
  // Lexical LET / LAMBDA bindings shadow workbook names. A binding's
  // initializer or argument was walked where it was bound, unless it is a
  // deferred reference, which this value read walks now.
  if (const WalkState::LexicalBinding* lexical = lookup_unqualified_lexical(node, state); lexical != nullptr) {
    if (lexical->deferred != nullptr) {
      walk(*lexical->deferred, state);
    }
    return;
  }
  // Resolve against the workbook's defined-name list. A miss (the name
  // is undefined, scoped to a different sheet, or hidden behind a cycle
  // already on the expansion stack) is a silent skip — same policy as
  // unknown sheet qualifiers and out-of-bounds coordinates. On a hit we
  // re-enter `walk()` with the parsed body so cells, ranges, volatility,
  // and nested NameRefs all surface naturally.
  const DefinedName* def = find_name_ref_definition(node, state);
  if (def == nullptr) {
    return;
  }
  expand_defined_name(*def, state, /*invoked=*/false);
}

void walk_call(const parser::AstNode& node, WalkState& state) {
  const bool in_position_only = state.position_only_call;
  state.position_only_call = false;
  const WalkState::LexicalBinding* lexical = lookup_lexical(node.as_call_name(), state);
  const DefinedName* defined =
      lexical == nullptr ? find_defined_name(*state.workbook, state.name_scope_sheet_id, node.as_call_name()) : nullptr;
  const bool builtin = lexical == nullptr && defined == nullptr;
  // Read before the arguments are walked: a LET among them grows the stack.
  const parser::AstNode* lexical_lambda = lexical != nullptr ? lexical->lambda : nullptr;
  const std::uint32_t arity = node.as_call_arity();
  // A reference argument a lambda binds -- the call's own, or a helper's
  // source feeding its callback -- is walked only where it is read as a value.
  DeferredArgs deferred(arity, nullptr);
  DeferredArgs callback_args;
  const CallbackHelper helper_kind =
      builtin && arity > 0U ? callback_helper_kind(node.as_call_name()) : CallbackHelper::None;
  const bool helper = helper_kind != CallbackHelper::None;
  if (lexical_lambda != nullptr || defined != nullptr) {
    for (std::uint32_t i = 0; i < arity; ++i) {
      deferred[i] = deferrable_reference(node.as_call_arg(i), state);
    }
  } else if (helper) {
    for (std::uint32_t param = 0; param + 1U < arity; ++param) {
      const int source = callback_source_arg(helper_kind, arity, param);
      const parser::AstNode* ref =
          source >= 0 ? deferrable_reference(node.as_call_arg(static_cast<std::uint32_t>(source)), state) : nullptr;
      callback_args.push_back(ref);
      if (ref != nullptr) {
        deferred[static_cast<std::uint32_t>(source)] = ref;
      }
    }
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg = node.as_call_arg(i);
    // A reference-returning call whose result is only positioned passes
    // that on to the arguments it can return (`ROW(INDEX(A1:A3,2))`).
    const bool position_only =
        builtin && (is_reference_only_arg(node, i) ||
                    (in_position_only && is_reference_carrying_arg(node.as_call_name(), i, arity)));
    if (position_only && is_static_reference(arg, state)) {
      continue;
    }
    if (deferred[i] != nullptr) {
      continue;
    }
    const DeferredArgs* bound = (helper && i + 1U == arity) ? &callback_args : nullptr;
    if (builtin && walk_callable_arg(arg, state, bound)) {
      continue;
    }
    walk_deferred_args(bound, 0U, state);
    state.position_only_call = position_only && arg.kind() == parser::NodeKind::Call;
    walk(arg, state);
    state.position_only_call = false;
  }

  // A lexical binding or a visible defined name shadows a built-in name;
  // in particular, a LET-bound `NOW` / `RAND` must not be marked
  // volatile merely because its spelling resembles a built-in.
  if (builtin && VolatileTracker::is_volatile_function(node.as_call_name())) {
    state.out->is_volatile = true;
    if (reads_cells_at_run_time(node)) {
      state.out->has_dynamic_reference = true;
    }
  }
  if (lexical != nullptr) {
    if (lexical_lambda != nullptr) {
      walk_invoked_lambda_body(*lexical_lambda, state, &deferred);
    }
    return;
  }
  if (defined != nullptr) {
    // `Name(args)` is the only call shape for a workbook-defined
    // LAMBDA. Expanding its body once is enough: recursive re-entry is
    // blocked by `name_stack`, which then walks the deferred arguments.
    expand_defined_name(*defined, state, /*invoked=*/true, &deferred);
  }
}

void walk_let_binding(const parser::AstNode& node, WalkState& state) {
  // LET introduces local names that shadow workbook references inside
  // its body. Walking the binding initialisers is straightforward (they
  // live in the outer scope). The body is descended unconditionally so
  // that `=LET(x, A1, x + B1)` records B1 and `=LET(x, 1, x + RAND())`
  // records the volatile call. Each binding is pushed onto the lexical
  // stack, so a bound name reaching the `NameRef` case shadows a
  // workbook-scoped defined name of the same identifier.
  const std::uint32_t binding_count = node.as_let_binding_count();
  const std::size_t saved_lexical_depth = state.lexical_stack.size();
  for (std::uint32_t i = 0; i < binding_count; ++i) {
    const parser::AstNode& expr = node.as_let_binding_expr(i);
    const parser::AstNode* deferred = deferrable_reference(expr, state);
    if (deferred == nullptr) {
      walk(expr, state);
    }
    std::optional<Footprint> footprint = reference_footprint(expr, state);
    const parser::AstNode* lambda = nullptr;
    if (expr.kind() == parser::NodeKind::Lambda) {
      lambda = &expr;
    } else if (const WalkState::LexicalBinding* prior = lookup_unqualified_lexical(expr, state); prior != nullptr) {
      lambda = prior->lambda;
    }
    state.lexical_stack.push_back({node.as_let_binding_name(i), lambda, footprint, deferred});
  }
  walk(node.as_let_body(), state);
  state.lexical_stack.resize(saved_lexical_depth);
}

void walk_lambda_call(const parser::AstNode& node, WalkState& state) {
  const parser::AstNode& callee = node.as_lambda_call_callee();
  const std::uint32_t arity = node.as_lambda_call_arity();
  DeferredArgs deferred(arity, nullptr);
  const DefinedName* def = nullptr;
  if (is_defined_name_ref(callee)) {
    def = find_name_ref_definition(callee, state);
  }
  if (callee.kind() == parser::NodeKind::Lambda || def != nullptr) {
    for (std::uint32_t i = 0; i < arity; ++i) {
      deferred[i] = deferrable_reference(node.as_lambda_call_arg(i), state);
    }
  }
  if (callee.kind() == parser::NodeKind::Lambda) {
    // Directly-invoked lambda (`=LAMBDA(x, x+A1)(5)`): unlike a bare
    // lambda *value*, the body IS evaluated here, so its cell refs and
    // volatile calls are genuine dependencies that must reach the graph.
    walk_invoked_lambda_body(callee, state, &deferred);
  } else if (is_defined_name_ref(callee)) {
    // `Sheet1!Fn(5)` / `[0]!Fn(5)`: qualified spellings of `Fn(5)`, which
    // invoke the definition just as the `Call` case does.
    if (def != nullptr) {
      expand_defined_name(*def, state, /*invoked=*/true, &deferred);
    }
  } else {
    // The remaining callee kind is a nested `LambdaCall` (currying,
    // e.g. `LAMBDA(x, LAMBDA(y, x+y))(3)(4)`): the parser gates this
    // postfix `(` to a `Lambda`, a `LambdaCall` or a qualified name,
    // and an unqualified named callee (`=MyLambda(5)`) parses as
    // a `Call` handled above. Walking generically here recurses back
    // into this same `case` (or into `Lambda`, at the base of the curry
    // chain), which is what still surfaces the chain's embedded refs
    // and volatile calls.
    walk(callee, state);
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    if (deferred[i] == nullptr) {
      walk(node.as_lambda_call_arg(i), state);
    }
  }
}

void walk(const parser::AstNode& node, WalkState& state) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::ErrorPlaceholder:
      return;

    // A cross-workbook reference reads a cache attached to the workbook,
    // never a cell of it, so it contributes no edge to the dependency
    // graph. The cache changes only when the file is reloaded, which
    // rebuilds the graph anyway. `[0]!Name` is this workbook's own name.
    case parser::NodeKind::ExternalRef:
      if (const DefinedName* def = find_name_ref_definition(node, state); def != nullptr) {
        expand_defined_name(*def, state, /*invoked=*/false);
      }
      return;

    case parser::NodeKind::Ref:
      walk_ref(node, state);
      return;

    case parser::NodeKind::SpillRef:
      walk_spill_ref(node, state);
      return;

    case parser::NodeKind::Ref3D:
      walk_ref3d(node, state);
      return;

    case parser::NodeKind::StructuredRef:
      walk_structured_ref(node, state);
      return;

    case parser::NodeKind::NameRef:
      walk_name_ref(node, state);
      return;

    case parser::NodeKind::UnaryOp:
      walk(node.as_unary_operand(), state);
      return;

    case parser::NodeKind::BinaryOp:
      for (std::uint32_t i = 0; i < node.as_binary_arity(); ++i) {
        walk(node.as_binary_operand(i), state);
      }
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

    case parser::NodeKind::Call:
      walk_call(node, state);
      return;

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

    case parser::NodeKind::LetBinding:
      walk_let_binding(node, state);
      return;

    case parser::NodeKind::LambdaCall:
      walk_lambda_call(node, state);
      return;
  }
}

}  // namespace

ExtractedDeps extract_deps(const parser::AstNode& node, CellNodeId formula_cell, const Workbook& workbook) {
  ExtractedDeps deps;
  // The arena owns any ASTs parsed for defined-name expansion. It lives only
  // as long as this call so the parsed nodes never outlive the walk; the
  // caller-supplied `node` is unrelated and stays in its own arena.
  Arena name_arena;
  WalkState state{&deps, {}, formula_cell.sheet_id, formula_cell.row, formula_cell.sheet_id, &workbook, &name_arena, {},
                  {},    {}};
  walk(node, state);
  return deps;
}

}  // namespace formulon::eval
