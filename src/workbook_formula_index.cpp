//
// Keeps the recalc engine's dependency graph in step with the workbook's
// formulas when its structure changes: full and scoped re-registration, and
// the reference scans that decide which formulas a change affects.

#include "workbook_formula_index.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/builtin_names.h"
#include "eval/recalc_engine.h"
#include "merge_range.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_metadata.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {

parser::AstNode* parse_indexable_formula(std::string_view body, Arena& arena) {
  return eval::parse_formula_entry(body, arena);
}

// Parses stored formula text (leading '=' optional) into `arena`, resetting it
// first. Null when the body is empty (the arena is then left untouched) or
// does not parse.
parser::AstNode* parse_stored_formula(std::string_view text, Arena& arena) {
  text = strip_formula_prefix(text);
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
      for (const auto& [col, cell] : cells) {
        if (cell.formula_text.empty()) {
          continue;
        }
        const eval::CellNodeId node = make_node(sheet_idx, row, static_cast<std::uint32_t>(col));
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

// Appends to `found` each entry of `locals` that `node` names as a NameRef
// qualified by `sheet`.
void collect_local_name_refs(const parser::AstNode& node, std::string_view sheet,
                             const std::vector<std::string>& locals, std::vector<std::string>& found) {
  if (node.kind() == parser::NodeKind::NameRef) {
    const std::string_view qualifier = node.as_name_sheet();
    if (qualifier.empty() || !sheet_names::equal(qualifier, sheet)) {
      return;
    }
    for (const std::string& local : locals) {
      if (strings::case_insensitive_eq(local, node.as_name()) && !contains_affected_name(found, local)) {
        found.push_back(local);
      }
    }
    return;
  }
  parser::any_child_node(node, [&](const parser::AstNode& child) {
    collect_local_name_refs(child, sheet, locals, found);
    return false;
  });
}

}  // namespace

std::vector<std::string> referenced_local_names_of_removed_sheet(const std::vector<Sheet>& sheets,
                                                                 const std::vector<DefinedName>& defined_names,
                                                                 const std::vector<TableMetadata>& tables,
                                                                 std::size_t removed_index,
                                                                 const std::vector<std::string>& extra_formulas) {
  const std::int32_t removed_id = static_cast<std::int32_t>(removed_index);
  std::vector<std::string> locals;
  for (const DefinedName& entry : defined_names) {
    if (entry.local_sheet_id == removed_id) {
      locals.push_back(entry.name);
    }
  }
  if (locals.empty()) {
    return {};
  }
  const std::string_view sheet_name = sheets[removed_index].name();
  std::vector<std::string> referenced;
  Arena arena;
  const auto scan = [&](std::string_view text) {
    if (const parser::AstNode* root = parse_stored_formula(text, arena); root != nullptr) {
      collect_local_name_refs(*root, sheet_name, locals, referenced);
    }
  };

  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    if (sheet_idx == removed_index) {
      continue;
    }
    const Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (const auto& [col, cell] : cells) {
        scan(cell.formula_text);
      }
    }
    for (const cf::ConditionalFormat& conditional_format : sheet.conditional_formats()) {
      for (const cf::CFRule& rule : conditional_format.rules) {
        if (rule.formula1.has_value()) {
          scan(*rule.formula1);
        }
        if (rule.formula2.has_value()) {
          scan(*rule.formula2);
        }
      }
    }
    for (const DataValidation& validation : sheet.validations()) {
      scan(validation.formula1);
      scan(validation.formula2);
    }
  }
  for (const DefinedName& entry : defined_names) {
    if (entry.local_sheet_id != removed_id) {
      scan(entry.formula);
    }
  }
  for (const TableMetadata& table : tables) {
    if (table.sheet_index == removed_index) {
      continue;
    }
    for (const TableColumn& column : table.columns) {
      scan(column.calculated_column_formula);
    }
  }
  for (const std::string& text : extra_formulas) {
    scan(text);
  }

  // A promoted name keeps its body, so the locals that body reaches survive too.
  const auto has_book_namesake = [&](std::string_view name) { return detail::has_book_namesake(defined_names, name); };
  for (std::size_t i = 0; i < referenced.size(); ++i) {
    const std::string name = referenced[i];
    if (has_book_namesake(name)) {
      continue;
    }
    for (const DefinedName& entry : defined_names) {
      if (entry.local_sheet_id == removed_id && strings::case_insensitive_eq(entry.name, name)) {
        scan(entry.formula);
      }
    }
  }
  return referenced;
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
                                                 const eval::RecalcEngine::LockedMutator& mutator,
                                                 const std::vector<DefinedName>& defined_names) {
  const auto visibility_sensitive = [](std::string_view formula) {
    return strings::case_insensitive_contains(formula, "SUBTOTAL") ||
           strings::case_insensitive_contains(formula, "AGGREGATE");
  };
  std::vector<std::string> affected_names;
  for (const DefinedName& name : defined_names) {
    if (visibility_sensitive(name.formula)) {
      affected_names.push_back(name.name);
    }
  }
  if (!affected_names.empty()) {
    affected_names = close_affected_names(defined_names, std::move(affected_names));
  }
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    const Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (const auto& [col, cell] : cells) {
        if (cell.formula_text.empty()) {
          continue;
        }
        // A call must spell its name, so text scans miss nothing; a string literal costs a spare dirty mark.
        bool affected = visibility_sensitive(cell.formula_text);
        const std::string_view text = cell.formula_text;
        const auto spells_name = [text](const std::string& name) {
          return strings::case_insensitive_contains(text, name);
        };
        if (!affected && std::any_of(affected_names.begin(), affected_names.end(), spells_name)) {
          const parser::AstNode* root = parse_stored_formula(cell.formula_text, parser_arena);
          affected = root != nullptr && references_any_name(*root, affected_names);
        }
        if (affected) {
          mutator.mark_dirty(make_node(sheet_idx, row, static_cast<std::uint32_t>(col)));
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

void mark_all_formulas_dirty_locked(const std::vector<Sheet>& sheets,
                                    const eval::RecalcEngine::LockedMutator& mutator) {
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    for (const auto& [row, cells] : sheets[sheet_idx].rows()) {
      for (const auto& [col, cell] : cells) {
        if (cell.formula_text.empty()) {
          continue;
        }
        mutator.mark_dirty(eval::CellNodeId{static_cast<std::uint16_t>(sheet_idx), row, col});
      }
    }
  }
}

void sync_table_spill_barriers_locked(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                      const std::vector<TableMetadata>& tables) {
  std::vector<std::vector<MergeRange>> per_sheet(sheets.size());
  for (const TableMetadata& table : tables) {
    const std::optional<MergeRange> rect = parse_a1_rectangle(table.ref);
    if (rect && table.sheet_index < sheets.size()) {
      per_sheet[table.sheet_index].push_back(normalized(*rect));
    }
  }
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    std::vector<MergeRange> touched = per_sheet[sheet_idx];
    const std::vector<MergeRange> before = sheet.set_table_spill_barriers(std::move(per_sheet[sheet_idx]));
    if (before == touched) {
      continue;
    }
    touched.insert(touched.end(), before.begin(), before.end());
    for (const MergeRange& rect : touched) {
      const std::uint32_t rows = rect.last_row - rect.first_row + 1U;
      const std::uint32_t cols = rect.last_col - rect.first_col + 1U;
      for (const auto query :
           {&Sheet::blocked_spill_anchors_intersecting, &Sheet::committed_spill_anchors_intersecting}) {
        for (const CellAddress anchor : (sheet.*query)(rect.first_row, rect.first_col, rows, cols)) {
          mutator.mark_dirty(make_node(sheet_idx, anchor.row, anchor.col));
        }
      }
    }
  }
}

}  // namespace formulon
