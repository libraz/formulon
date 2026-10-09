//
// Keeps the recalc engine's dependency graph in step with the workbook's
// formulas when its structure changes: full and scoped re-registration, and
// the reference scans that decide which formulas a change affects.

#include "workbook_formula_index.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/builtin_names.h"
#include "eval/recalc_engine.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {

parser::AstNode* parse_indexable_formula(std::string_view body, Arena& arena) {
  return eval::parse_formula_entry(body, arena);
}

// Parses stored formula text (leading '=' optional) into `arena`, resetting it
// first. Null when the body is empty (the arena is then left untouched) or
// does not parse.
parser::AstNode* parse_stored_formula(std::string_view text, Arena& arena) {
  if (!text.empty() && text.front() == '=') {
    text.remove_prefix(1);
  }
  if (text.empty()) {
    return nullptr;
  }
  arena.reset();
  return parse_indexable_formula(text, arena);
}

// Rebuilds the dependency graph from scratch after a sheet permutation.
//
// A `CellNodeId.sheet_id` is the workbook-relative sheet index, so
// removing or moving a sheet invalidates the `sheet_id` of every node on
// every shifted sheet at once. Patching individual edges is error-prone
// (the pre-move ids no longer name the right sheet once the vector has
// been reordered), so the safe baseline is to drop the whole graph and
// re-register each formula against its current position. Every formula
// cell is then marked dirty so the next recalc re-evaluates against the
// rearranged workbook — matching Excel's post-rearrange recalculation.
//
// Precondition: the caller holds the engine mutex for the lifetime of the
// `LockedMutator&`, and `sheets` is already in its final post-permutation
// order.
void reindex_all_formulas(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                          const Workbook& workbook) {
  // Defined-name retargets and sheet permutations invalidate every cached
  // formula result. Clear both committed and blocked spill geometry before
  // rebuilding edges; otherwise a NameRef that changed from SEQUENCE to a
  // scalar expression could leave stale phantom values visible to a range
  // watcher. Structural row/column edits snapshot and restore blocked
  // records around this helper, so their pending-spill contract is retained.
  for (Sheet& sheet : sheets) {
    sheet.clear_all_spills();
  }
  mutator.reset_graph();
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    const Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        const eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        if (parser::AstNode* root = parse_stored_formula(cell.formula_text, parser_arena); root != nullptr) {
          mutator.register_formula(node, *root, workbook);
        }
        // Every formula cell recomputes after a rearrangement, regardless
        // of whether its refs changed.
        mutator.mark_dirty(node);
      }
    }
  }
}

namespace {

// Which reference spelling `references_any` compares against its candidates.
enum class RefKind : std::uint8_t { kName, kTable, kSheet };

// True once any entry of `candidates` (case-insensitive) is named by a
// reference of `kind` in `node`'s subtree. A name is a `NameRef`, a `[0]!Name`
// or a `Call` callee (a LAMBDA-valued defined name is called as `Fn(args)`; a
// built-in callee simply matches no defined name); a table is a
// `StructuredRef`'s table specifier. Node-kind coverage mirrors
// `parser::TransformNode` (ast_shift.cpp) exhaustively, so a reference nested
// inside any expression form (calls, LET/LAMBDA bodies, array literals,
// unions...) is found; a reference shadowed by an enclosing LET/LAMBDA
// parameter of the same spelling still counts -- treating it as a real
// reference only costs an unnecessary reindex, never a missed one.
bool references_any(const parser::AstNode& node, const std::vector<std::string>& candidates, RefKind kind) {
  const auto matches = [&](RefKind ref_kind, std::string_view ref) {
    if (ref_kind != kind) {
      return false;
    }
    return std::any_of(candidates.begin(), candidates.end(), [&](const std::string& n) {
      return kind == RefKind::kSheet ? sheet_names::equal(n, ref) : strings::case_insensitive_eq(n, ref);
    });
  };
  switch (node.kind()) {
    case parser::NodeKind::NameRef:
      return matches(RefKind::kSheet, node.as_name_sheet()) || matches(RefKind::kName, node.as_name());
    case parser::NodeKind::StructuredRef:
      return matches(RefKind::kTable, node.as_structured_ref_table());
    case parser::NodeKind::ExternalRef:
      // External workbook sheets and literals are outside this workbook. The
      // self-book form is a defined-name spelling and has no sheet qualifier.
      return kind != RefKind::kSheet && parser::is_self_book_name_ref(node) &&
             matches(RefKind::kName, node.as_external_ref_name());
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::ErrorPlaceholder:
      return false;
    case parser::NodeKind::Ref:
      return matches(RefKind::kSheet, node.as_ref().sheet);
    case parser::NodeKind::SpillRef:
      if (node.as_spill_ref_anchor_expr() == nullptr) {
        return matches(RefKind::kSheet, node.as_spill_ref().sheet);
      }
      break;
    case parser::NodeKind::Ref3D:
      return matches(RefKind::kSheet, node.as_ref3d_sheet_begin()) ||
             matches(RefKind::kSheet, node.as_ref3d_sheet_end());
    case parser::NodeKind::Call:
      if (matches(RefKind::kName, node.as_call_name())) {
        return true;
      }
      break;
    default:
      break;
  }

  return parser::any_child_node(node,
                                [&](const parser::AstNode& child) { return references_any(child, candidates, kind); });
}

}  // namespace

bool references_any_name(const parser::AstNode& root, const std::vector<std::string>& candidates) {
  return references_any(root, candidates, RefKind::kName);
}

namespace {

bool references_any_table(const parser::AstNode& root, const std::vector<std::string>& candidates) {
  return references_any(root, candidates, RefKind::kTable);
}

bool references_any_sheet(const parser::AstNode& root, const std::vector<std::string>& candidates) {
  return references_any(root, candidates, RefKind::kSheet);
}

bool contains_affected_name(const std::vector<std::string>& affected, std::string_view candidate) {
  return std::any_of(affected.begin(), affected.end(),
                     [&](const std::string& name) { return strings::case_insensitive_eq(name, candidate); });
}

}  // namespace

std::vector<std::string> close_affected_names(const std::vector<DefinedName>& defined_names,
                                              std::vector<std::string> affected) {
  Arena arena;
  bool grew = true;
  while (grew) {
    grew = false;
    for (const DefinedName& entry : defined_names) {
      if (contains_affected_name(affected, entry.name)) {
        continue;
      }
      const parser::AstNode* root = parse_stored_formula(entry.formula, arena);
      if (root != nullptr && references_any_name(*root, affected)) {
        affected.emplace_back(entry.name);
        grew = true;
      }
    }
  }
  return affected;
}

namespace {

// Closure of defined names whose resolved value can change when
// `changed_name` is added, retargeted, or removed: `changed_name` itself,
// plus every other defined name whose own formula references it
// (transitively). `defined_names` is a handful of entries in practice, so
// the O(names^2) fixpoint here is cheap next to reparsing every formula
// cell in the workbook, which the caller below does exactly once per
// affected cell rather than once per cell regardless of relevance.
std::vector<std::string> collect_affected_names(const std::vector<DefinedName>& defined_names,
                                                std::string_view changed_name) {
  return close_affected_names(defined_names, {std::string(changed_name)});
}

// Seeds the same reverse-alias closure from every defined name whose body
// mentions `changed_sheet`. Direct sheet references are handled separately;
// these seeds cover formulas that reach the sheet only through a name.
std::vector<std::string> collect_affected_names_referencing_sheet(const std::vector<DefinedName>& defined_names,
                                                                  std::string_view changed_sheet) {
  std::vector<std::string> affected;
  const std::vector<std::string> sheet_candidates{std::string(changed_sheet)};
  Arena arena;
  for (const DefinedName& entry : defined_names) {
    const parser::AstNode* root = parse_stored_formula(entry.formula, arena);
    if (root != nullptr && references_any_sheet(*root, sheet_candidates) &&
        !contains_affected_name(affected, entry.name)) {
      affected.emplace_back(entry.name);
    }
  }
  return close_affected_names(defined_names, std::move(affected));
}

}  // namespace

void reindex_formulas_referencing_name(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                       const Workbook& workbook, std::string_view changed_name) {
  const std::vector<std::string> affected = collect_affected_names(workbook.defined_names(), changed_name);
  reindex_formulas_if(sheets, mutator, workbook,
                      [&](const parser::AstNode& root) { return references_any_name(root, affected); });
}

void reindex_formulas_referencing_sheet(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                        const Workbook& workbook, std::string_view changed_sheet) {
  const std::vector<std::string> sheet_candidates{std::string(changed_sheet)};
  const std::vector<std::string> affected_names =
      collect_affected_names_referencing_sheet(workbook.defined_names(), changed_sheet);
  reindex_formulas_if(sheets, mutator, workbook, [&](const parser::AstNode& root) {
    return references_any_sheet(root, sheet_candidates) || references_any_name(root, affected_names);
  });
}

/// Body of `Workbook::mark_row_visibility_dependents_dirty`; the caller
/// holds the engine mutex.
void mark_row_visibility_dependents_dirty_locked(const std::vector<Sheet>& sheets,
                                                 const eval::RecalcEngine::LockedMutator& mutator) {
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    const Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        // SUBTOTAL/AGGREGATE can only be called by literally spelling one
        // of these names, so this text scan has no false negatives; a
        // string literal that happens to contain one costs a harmless
        // extra dirty mark rather than reparsing every formula in the
        // workbook to test more precisely.
        if (strings::case_insensitive_contains(cell.formula_text, "SUBTOTAL") ||
            strings::case_insensitive_contains(cell.formula_text, "AGGREGATE")) {
          mutator.mark_dirty(
              eval::CellNodeId{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)});
        }
      }
    }
  }
}

// Re-registers and dirties only the formula cells whose structured
// reference resolves through `dep_extractor` into a static rectangle
// derived from a table's current `ref`/columns (see
// `eval/dep_extractor.cpp`'s `StructuredRef` handling), so a table create,
// ref/column edit, or removal needs the same scoped treatment as a defined-
// name edit: only a formula naming one of `table_names` (a table's `name`
// and `display_name` can each appear in a structured reference) can be
// affected.
void reindex_formulas_referencing_table(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                        const Workbook& workbook, const std::vector<std::string>& table_names) {
  reindex_formulas_if(sheets, mutator, workbook,
                      [&](const parser::AstNode& root) { return references_any_table(root, table_names); });
}

// Builds a CellNodeId for a workbook-relative coordinate. `sheet_index` is
// below `sheet_count()`, which every append path bounds by
// `Workbook::kMaxSheets`, so the narrowing to the dep graph's 16-bit sheet
// id is lossless.
eval::CellNodeId make_node(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  return eval::CellNodeId{static_cast<std::uint16_t>(sheet_index), row, col};
}

}  // namespace formulon
