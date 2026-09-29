//
// Workbook implementation. Wires the OOXML save path and the embedded
// recalc engine. The `RecalcEngine` is held via `unique_ptr` (PIMPL-style)
// so the public header does not need to include the heavyweight recalc /
// dep-graph headers.

#include "workbook.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "defined_name.h"
#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "external_link.h"
#include "io/format_detect.h"
#include "io/ooxml_writer.h"
#include "io/workbook_kind_ooxml.h"
#include "io/xlsb/writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook_ref_rewrite.h"

namespace formulon {
namespace {
// Forward declaration; defined further below alongside the other
// sheet-name helpers. Declared here so the early `add_sheet_validated`
// definition can call it.
Expected<void, Error> validate_sheet_name(std::string_view name);
}  // namespace

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()), kind_(WorkbookKind::kXlsx) {}
Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;
Workbook::~Workbook() = default;

namespace {

/// OOXML pattern ordinal for `gray125`.
constexpr std::uint8_t kFillPatternGray125 = 17U;

/// Fills `styles` with the minimum style table Excel puts in a new
/// workbook.
///
/// Leaving the table empty is representable — the writer synthesises a
/// single default record per section when it sees one — but the synthesis
/// stops the moment a caller appends anything. A caller that adds one fill
/// therefore lands it at `fills[0]`, taking the slot Excel reserves for
/// `none`, and every `fillId` in the file shifts by one relative to what
/// Excel expects. The same trap applies to fonts and borders, and it is
/// only visible after the file reaches Excel.
///
/// Seeding also makes index 0 a usable template: `get_font(0)` returns the
/// default font, so "copy the default and change the size" works without a
/// separate partial-input type.
void seed_default_styles(StylesTable& styles) {
  FontRecord font;
  font.name = "Calibri";
  font.size = 11.0;
  font.has_family = true;
  font.family = 2U;
  styles.fonts.push_back(std::move(font));

  // Excel reserves the first two fills and never renders either: `none` is
  // the no-fill default every unstyled cell points at, and `gray125` is a
  // legacy placeholder it writes unconditionally.
  styles.fills.emplace_back();
  FillRecord gray125;
  gray125.pattern = kFillPatternGray125;
  styles.fills.push_back(gray125);

  styles.borders.emplace_back();
  styles.cell_style_xfs.emplace_back();
  styles.cell_xfs.emplace_back();

  CellStyleRecord normal;
  normal.name = "Normal";
  normal.xf_id = 0U;
  normal.builtin_id = 0U;
  styles.cell_styles.push_back(std::move(normal));
}

}  // namespace

Workbook Workbook::create() {
  Workbook wb;
  wb.sheets_.emplace_back(Sheet{std::string("Sheet1")});
  seed_default_styles(wb.styles_);
  return wb;
}

Workbook Workbook::create_empty() {
  // No default sheet; callers are expected to populate via add_sheet().
  // The style table is seeded all the same: the reserved-slot trap that
  // `seed_default_styles` documents is a property of appending to an empty
  // table, not of having a sheet, so both construction paths must hand the
  // caller a table whose reserved entries are already taken. A reader that
  // builds on top of this replaces the whole table with the one it parsed.
  Workbook wb;
  seed_default_styles(wb.styles_);
  return wb;
}

std::string_view Workbook::intern_text(std::string_view text) {
  text_storage_.emplace_back(text.data(), text.size());
  return std::string_view(text_storage_.back());
}

namespace {
// Rejects an append that would take the workbook past `kMaxSheets`. Every
// append routes through here, so `sheets_.size() <= kMaxSheets` holds for
// the lifetime of the workbook and each narrowing of a sheet index to the
// dep graph's 16-bit `sheet_id` is lossless by construction.
Expected<void, Error> check_sheet_headroom(std::size_t current_count) {
  if (current_count >= Workbook::kMaxSheets) {
    return make_error(FormulonErrorCode::kSheetCountLimitExceeded,
                      "add_sheet: workbook already holds the maximum sheets",
                      "limit=" + std::to_string(Workbook::kMaxSheets));
  }
  return {};
}
}  // namespace

std::size_t Workbook::add_sheet(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  // No error channel here; at the ceiling the workbook is left as-is and
  // the caller gets the out-of-range sentinel (see the header contract).
  if (!check_sheet_headroom(sheets_.size()).has_value()) {
    return kMaxSheets;
  }
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.size() - 1U;
}

Expected<std::size_t, Error> Workbook::add_sheet_checked(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.size() - 1U;
}

Expected<Sheet*, Error> Workbook::add_sheet_validated(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(validate_sheet_name(name));
  for (const Sheet& existing : sheets_) {
    if (sheet_names::equal(existing.name(), name)) {
      return make_error(FormulonErrorCode::kInvalidSheetName, "add_sheet: name collides with an existing sheet",
                        "name=\"" + name + "\"");
    }
  }
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  sheets_.emplace_back(Sheet{std::move(name)});
  return &sheets_.back();
}

namespace {

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
        std::string_view body = cell.formula_text;
        if (!body.empty() && body.front() == '=') {
          body.remove_prefix(1);
        }
        const eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        if (!body.empty()) {
          parser_arena.reset();
          parser::AstNode* root = parser::parse_strict(body, parser_arena);
          if (root != nullptr) {
            mutator.register_formula(node, *root, workbook);
          }
        }
        // Every formula cell recomputes after a rearrangement, regardless
        // of whether its refs changed.
        mutator.mark_dirty(node);
      }
    }
  }
}

// Which reference spelling `references_any` compares against its candidates.
enum class RefKind : std::uint8_t { kName, kTable };

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
    return ref_kind == kind && std::any_of(candidates.begin(), candidates.end(),
                                           [&](const std::string& n) { return strings::case_insensitive_eq(n, ref); });
  };
  const auto any = [&](const parser::AstNode& child) { return references_any(child, candidates, kind); };
  switch (node.kind()) {
    case parser::NodeKind::NameRef:
      return matches(RefKind::kName, node.as_name());
    case parser::NodeKind::StructuredRef:
      return matches(RefKind::kTable, node.as_structured_ref_table());
    case parser::NodeKind::ExternalRef:
      return parser::is_self_book_name_ref(node) && matches(RefKind::kName, node.as_external_ref_name());
    case parser::NodeKind::Literal:
    case parser::NodeKind::ErrorLiteral:
    case parser::NodeKind::ErrorPlaceholder:
    case parser::NodeKind::Ref:
    case parser::NodeKind::SpillRef:
    case parser::NodeKind::Ref3D:
      return false;
    case parser::NodeKind::UnaryOp:
      return any(node.as_unary_operand());
    case parser::NodeKind::BinaryOp:
      return any(node.as_binary_lhs()) || any(node.as_binary_rhs());
    case parser::NodeKind::RangeOp:
      return any(node.as_range_lhs()) || any(node.as_range_rhs());
    case parser::NodeKind::UnionOp:
      for (std::uint32_t i = 0; i < node.as_union_arity(); ++i) {
        if (any(node.as_union_child(i))) {
          return true;
        }
      }
      return false;
    case parser::NodeKind::IntersectOp:
      return any(node.as_intersect_lhs()) || any(node.as_intersect_rhs());
    case parser::NodeKind::ImplicitIntersection:
      return any(node.as_implicit_intersection_operand());
    case parser::NodeKind::Call:
      if (matches(RefKind::kName, node.as_call_name())) {
        return true;
      }
      for (std::uint32_t i = 0; i < node.as_call_arity(); ++i) {
        if (any(node.as_call_arg(i))) {
          return true;
        }
      }
      return false;
    case parser::NodeKind::ArrayLiteral:
      for (std::uint32_t r = 0; r < node.as_array_rows(); ++r) {
        for (std::uint32_t c = 0; c < node.as_array_cols(); ++c) {
          if (any(node.as_array_element(r, c))) {
            return true;
          }
        }
      }
      return false;
    case parser::NodeKind::Lambda:
      return any(node.as_lambda_body());
    case parser::NodeKind::LetBinding:
      for (std::uint32_t i = 0; i < node.as_let_binding_count(); ++i) {
        if (any(node.as_let_binding_expr(i))) {
          return true;
        }
      }
      return any(node.as_let_body());
    case parser::NodeKind::LambdaCall:
      if (any(node.as_lambda_call_callee())) {
        return true;
      }
      for (std::uint32_t i = 0; i < node.as_lambda_call_arity(); ++i) {
        if (any(node.as_lambda_call_arg(i))) {
          return true;
        }
      }
      return false;
  }
  return false;
}

bool references_any_name(const parser::AstNode& root, const std::vector<std::string>& candidates) {
  return references_any(root, candidates, RefKind::kName);
}

bool references_any_table(const parser::AstNode& root, const std::vector<std::string>& candidates) {
  return references_any(root, candidates, RefKind::kTable);
}

// Closure of defined names whose resolved value can change when
// `changed_name` is added, retargeted, or removed: `changed_name` itself,
// plus every other defined name whose own formula references it
// (transitively). `defined_names` is a handful of entries in practice, so
// the O(names^2) fixpoint here is cheap next to reparsing every formula
// cell in the workbook, which the caller below does exactly once per
// affected cell rather than once per cell regardless of relevance.
std::vector<std::string> collect_affected_names(const std::vector<DefinedName>& defined_names,
                                                std::string_view changed_name) {
  std::vector<std::string> affected{std::string(changed_name)};
  Arena arena;
  bool grew = true;
  while (grew) {
    grew = false;
    for (const DefinedName& entry : defined_names) {
      if (std::any_of(affected.begin(), affected.end(),
                      [&](const std::string& n) { return strings::case_insensitive_eq(n, entry.name); })) {
        continue;
      }
      std::string_view body = entry.formula;
      if (!body.empty() && body.front() == '=') {
        body.remove_prefix(1);
      }
      if (body.empty()) {
        continue;
      }
      arena.reset();
      const parser::AstNode* root = parser::parse_strict(body, arena);
      if (root != nullptr && references_any_name(*root, affected)) {
        affected.emplace_back(entry.name);
        grew = true;
      }
    }
  }
  return affected;
}

// Re-registers and dirties only the formula cells whose value can change
// when `changed_name` (and every name that transitively references it, per
// `collect_affected_names`) is added, retargeted, or removed. Unlike
// `reindex_all_formulas`, this leaves every unaffected formula's dep-graph
// edges, cached value, and spill geometry untouched, so setting a print
// area/print titles/unrelated name no longer wipes spills workbook-wide or
// forces every formula to recompute. Shared by every scoped-reindex caller
// below (defined names, table structure) via `affected`, which tests the
// already-parsed AST; the parse itself cannot be skipped for the unaffected
// majority (no per-name/per-table dependent index exists in the recalc
// engine to look this up directly), so cost is proportional to the
// workbook's formula count regardless of how few are actually affected. It
// no longer re-registers dep-graph edges, resets the graph, or clears
// spills for that unaffected majority, unlike `reindex_all_formulas`.
template <typename Affected>
void reindex_formulas_if(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                         const Workbook& workbook, const Affected& affected) {
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        std::string_view body = cell.formula_text;
        if (!body.empty() && body.front() == '=') {
          body.remove_prefix(1);
        }
        if (body.empty()) {
          continue;
        }
        parser_arena.reset();
        parser::AstNode* root = parser::parse_strict(body, parser_arena);
        if (root == nullptr || !affected(*root)) {
          continue;
        }
        const eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        // No-op when `(row, col)` is not a spill anchor.
        sheet.clear_spill(row, static_cast<std::uint32_t>(col));
        mutator.register_formula(node, *root, workbook);
        mutator.mark_dirty(node);
      }
    }
  }
}

void reindex_formulas_referencing_name(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                       const Workbook& workbook, std::string_view changed_name) {
  const std::vector<std::string> affected = collect_affected_names(workbook.defined_names(), changed_name);
  reindex_formulas_if(sheets, mutator, workbook,
                      [&](const parser::AstNode& root) { return references_any_name(root, affected); });
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

// Excel's structural validation for sheet names: strict UTF-8, non-empty,
// ≤ 31 UTF-16 units, and no `: \ / ? * [ ]`. The forbidden-character scan is
// a byte-level check (the forbidden set is ASCII), so it works correctly on
// valid UTF-8 sheet names because none of the disallowed code units appear as
// continuation bytes (all are < 0x80).
bool is_valid_sheet_name_chars(std::string_view name) noexcept {
  for (char byte : name) {
    switch (byte) {
      case ':':
      case '\\':
      case '/':
      case '?':
      case '*':
      case '[':
      case ']':
        return false;
      default:
        break;
    }
  }
  return true;
}

// Excel measures the 31-"character" sheet-name limit in UTF-16 code units
// (its internal string representation), not UTF-8 bytes. Counting bytes
// wrongly rejects a 31-character Japanese name (up to 93 bytes) far short
// of the real limit; counting code units matches Excel and treats a
// supplementary-plane emoji as the two units Excel charges for it.
constexpr std::uint32_t kMaxSheetNameUnits = 31U;

// Shared structural validator for a sheet name across every mutation
// surface (add / rename / any future import-side check). Verifies strict
// UTF-8, a non-empty name within the code-unit length limit, and no forbidden
// characters. Duplicate/case-folding collision is caller-scoped (it needs
// the target index) and handled at each call site.
Expected<void, Error> validate_sheet_name(std::string_view name) {
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name must not be empty", "name=\"\"");
  }
  if (!sheet_names::valid_utf8(name)) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name is not valid UTF-8", "name=invalid-utf8");
  }
  if (utf16_units_in(name) > kMaxSheetNameUnits) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name exceeds 31 characters",
                      "name=\"" + std::string(name) + "\"");
  }
  if (!is_valid_sheet_name_chars(name)) {
    return make_error(FormulonErrorCode::kInvalidSheetName,
                      "sheet name contains a forbidden character (: \\ / ? * [ ])",
                      "name=\"" + std::string(name) + "\"");
  }
  return Expected<void, Error>::Ok();
}

void remove_and_reindex_tables(std::vector<TableMetadata>& tables, std::size_t removed_sheet_index) {
  std::vector<TableMetadata> retained;
  retained.reserve(tables.size());
  for (TableMetadata& table : tables) {
    if (table.sheet_index == removed_sheet_index) {
      continue;
    }
    if (table.sheet_index > removed_sheet_index) {
      --table.sheet_index;
    }
    retained.push_back(std::move(table));
  }
  tables = std::move(retained);
}

void move_table_sheet_indices(std::vector<TableMetadata>& tables, std::size_t from_index, std::size_t to_index) {
  for (TableMetadata& table : tables) {
    if (table.sheet_index == from_index) {
      table.sheet_index = to_index;
    } else if (from_index < to_index && table.sheet_index > from_index && table.sheet_index <= to_index) {
      --table.sheet_index;
    } else if (from_index > to_index && table.sheet_index >= to_index && table.sheet_index < from_index) {
      ++table.sheet_index;
    }
  }
}

}  // namespace

Expected<void, Error> Workbook::rename_sheet(std::uint32_t index, std::string new_name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "rename_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Validate the new name (non-empty, ≤ 31 code units, no forbidden
  // characters) via the shared validator so add / rename agree.
  RETURN_IF_ERROR(validate_sheet_name(new_name));
  // Collision check (Unicode simple-fold). A no-op rename — same case-fold
  // as the current name — bypasses the collision check so callers can
  // change only the casing.
  for (std::size_t idx = 0; idx < sheets_.size(); ++idx) {
    if (idx == static_cast<std::size_t>(index)) {
      continue;
    }
    if (sheet_names::equal(sheets_[idx].name(), new_name)) {
      return make_error(FormulonErrorCode::kInvalidSheetName, "rename_sheet: name collides with another sheet",
                        "name=\"" + new_name + "\" colliding_index=" + std::to_string(idx));
    }
  }

  const std::string old_name = sheets_[index].name();

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  std::vector<std::uint32_t> ignored_cache_ids;
  bool defined_names_changed = false;
  const parser::SheetRenameTransform transform(old_name, new_name);
  // Keep the sheet name in its pre-rename state until every formula and
  // metadata holder has been transformed. In particular, do not re-register
  // a changed formula against `new_name` before the workbook can resolve it.
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, old_name, new_name,
                              {}, ignored_cache_ids, defined_names_changed);
  if (defined_names_changed) {
    // Defined-name expansion can change a formula's value without changing
    // the formula cell's own text. Sheet ids survive a rename, so retaining
    // the existing graph edges is safe; dirty every formula so those cached
    // values are nevertheless refreshed on the next pass.
    for (std::size_t sheet_idx = 0; sheet_idx < sheets_.size(); ++sheet_idx) {
      for (const auto& [row, cells] : sheets_[sheet_idx].rows()) {
        for (std::size_t col = 0; col < cells.size(); ++col) {
          if (!cells[col].formula_text.empty()) {
            mutator.mark_dirty(
                eval::CellNodeId{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)});
          }
        }
      }
    }
  }
  // Rename last: the transform above intentionally reads the old workbook
  // name while formatting, then the final name makes all newly-written text
  // resolvable for the next recalc.
  sheets_[index].set_name(std::move(new_name));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_sheet(std::uint32_t index) {
  // Hold the engine mutex from the first workbook-state read through the
  // final graph rebuild. A parallel recalc therefore observes one complete
  // pre-removal or post-removal workbook, never the metadata halfway state.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "remove_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (sheets_.size() <= 1U) {
    return make_error(FormulonErrorCode::kCannotRemoveLastSheet, "remove_sheet: cannot remove the only sheet",
                      "sheet_count=" + std::to_string(sheets_.size()));
  }

  const std::string removed_name = sheets_[index].name();

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  std::vector<std::string_view> pre_removal_sheet_order;
  pre_removal_sheet_order.reserve(sheets_.size());
  for (const Sheet& sheet : sheets_) {
    pre_removal_sheet_order.push_back(sheet.name());
  }

  const parser::SheetRemovalTransform transform(pre_removal_sheet_order, index);
  std::vector<std::uint32_t> dropped_cache_ids;
  bool defined_names_changed = false;
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, removed_name,
                              dropped_cache_ids, defined_names_changed);
  (void)defined_names_changed;

  // Sheet-scoped names owned by the removed sheet disappear; names scoped to
  // later sheets follow their owner as the sheet vector closes the gap.
  std::vector<DefinedName> retained;
  retained.reserve(defined_names_.size());
  for (DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id == static_cast<std::int32_t>(index)) {
      continue;
    }
    if (entry.local_sheet_id > static_cast<std::int32_t>(index)) {
      entry.local_sheet_id -= 1;
    }
    retained.push_back(std::move(entry));
  }
  defined_names_ = std::move(retained);
  remove_and_reindex_tables(tables_, index);

  // Remove caches whose worksheet source resolved to the deleted sheet and
  // remove every surviving pivot table that would otherwise retain one of
  // those cache ids. This keeps the workbook graph free of dangling cache
  // bindings after the structural mutation.
  std::vector<std::unique_ptr<pivot::PivotCache>> retained_caches;
  retained_caches.reserve(pivot_caches_.size());
  for (std::unique_ptr<pivot::PivotCache>& cache : pivot_caches_) {
    if (cache == nullptr) {
      continue;
    }
    const bool dropped =
        std::find(dropped_cache_ids.begin(), dropped_cache_ids.end(), cache->cache_id()) != dropped_cache_ids.end();
    if (!dropped) {
      retained_caches.push_back(std::move(cache));
    }
  }
  pivot_caches_ = std::move(retained_caches);
  const auto cache_survives = [this](std::uint32_t cache_id) {
    return std::any_of(pivot_caches_.begin(), pivot_caches_.end(),
                       [cache_id](const auto& cache) { return cache != nullptr && cache->cache_id() == cache_id; });
  };
  for (Sheet& sheet : sheets_) {
    std::vector<std::unique_ptr<pivot::PivotTable>> retained_pivots;
    retained_pivots.reserve(sheet.pivot_tables().size());
    for (std::unique_ptr<pivot::PivotTable>& pivot_table : sheet.mutable_pivot_tables()) {
      if (pivot_table == nullptr) {
        continue;
      }
      const std::uint32_t cache_id = pivot_table->pivot_cache_id();
      if (cache_survives(cache_id) &&
          std::find(dropped_cache_ids.begin(), dropped_cache_ids.end(), cache_id) == dropped_cache_ids.end()) {
        retained_pivots.push_back(std::move(pivot_table));
      }
    }
    sheet.mutable_pivot_tables() = std::move(retained_pivots);
  }

  // Erase only after all transforms have resolved against the pre-removal
  // order. One and only one final reindex sees the final sheet/name/table/
  // cache topology and dirties every surviving formula consistently.
  sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(index));
  reindex_all_formulas(sheets_, mutator, *this);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::move_sheet(std::uint32_t from_index, std::uint32_t to_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(from_index) >= sheets_.size() || static_cast<std::size_t>(to_index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "move_sheet: index out of range",
                      "from=" + std::to_string(from_index) + " to=" + std::to_string(to_index) +
                          " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (from_index == to_index) {
    return Expected<void, Error>::Ok();
  }

  // `to_index` is the destination in the *post-removal* sheet list, which
  // matches Excel's UI semantics. Implementation: lift the sheet out,
  // then insert at the destination. The sheet vector mutation and the
  // subsequent per-cell `mark_dirty` loop run under a single hold of
  // the engine mutex so a concurrent `recalc_parallel` either sees the
  // pre-move workbook (sheets in their original order, dirty set
  // unchanged) or the fully patched one, never a half-applied move
  // where `sheets_` has been reordered but the dep-graph still indexes
  // cells by their pre-move sheet_id.
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  Sheet moving = std::move(sheets_[from_index]);
  sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(from_index));
  sheets_.insert(sheets_.begin() + static_cast<std::ptrdiff_t>(to_index), std::move(moving));

  // Patch sheet-scoped defined names so their `local_sheet_id` follows
  // the rearrangement. Workbook-scoped names reference sheets by name,
  // so they need no update.
  const auto from_id = static_cast<std::int32_t>(from_index);
  const auto to_id = static_cast<std::int32_t>(to_index);
  for (DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id < 0) {
      continue;
    }
    if (entry.local_sheet_id == from_id) {
      entry.local_sheet_id = to_id;
      continue;
    }
    if (from_id < to_id && entry.local_sheet_id > from_id && entry.local_sheet_id <= to_id) {
      entry.local_sheet_id -= 1;
    } else if (from_id > to_id && entry.local_sheet_id >= to_id && entry.local_sheet_id < from_id) {
      entry.local_sheet_id += 1;
    }
  }
  move_table_sheet_indices(tables_, from_index, to_index);

  // The recalc engine's `CellNodeId.sheet_id` is the workbook-relative
  // index, so a move renumbers every sheet in the `[min..max]` window and
  // invalidates the graph nodes/edges keyed by their old ids. Rebuild the
  // graph from scratch against the reordered sheet vector; every re-keyed
  // formula is marked dirty so the next `recalc()` re-evaluates it. This
  // mirrors Excel's post-rearrange behaviour where downstream formulas
  // re-evaluate.
  reindex_all_formulas(sheets_, mutator, *this);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_defined_name(std::string name, std::string formula) {
  return set_defined_name_scoped(std::move(name), std::move(formula), -1);
}

Expected<void, Error> Workbook::set_defined_name_scoped(std::string name, std::string formula,
                                                        std::int32_t local_sheet_id) {
  // Validation, mutation, and the dependent graph rebuild must share one
  // critical section. Otherwise a concurrent sheet removal can change the
  // count between validation and insertion, or a recalc can observe the
  // new definition before its graph is rebuilt.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: name is empty");
  }
  if (local_sheet_id < -1 || (local_sheet_id >= 0 && static_cast<std::size_t>(local_sheet_id) >= sheets_.size())) {
    return make_error(
        FormulonErrorCode::kInvalidArgument, "set_defined_name_scoped: local_sheet_id out of range",
        "local_sheet_id=" + std::to_string(local_sheet_id) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Case-insensitive lookup: Excel resolves defined names case-folded.
  // Restrict the search to the requested scope.
  for (auto it = defined_names_.begin(); it != defined_names_.end(); ++it) {
    if (it->local_sheet_id != local_sheet_id) {
      continue;
    }
    if (strings::case_insensitive_eq(it->name, name)) {
      if (formula.empty()) {
        defined_names_.erase(it);
      } else {
        it->formula = std::move(formula);
      }
      // Retargeting or removing an existing name changes what every formula
      // that references it (directly or through another name that expands
      // to it) resolves to, so their dep-graph edges and cached values are
      // now stale; a formula that never mentions this name is unaffected.
      // Re-register and dirty just that subset, and clear spill geometry
      // only where one of them anchors a spill, rather than the workbook-
      // wide rebuild `reindex_all_formulas` would force.
      const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
      reindex_formulas_referencing_name(sheets_, mutator, *this, name);
      return Expected<void, Error>::Ok();
    }
  }
  // No existing entry; appending only makes sense if a formula is
  // supplied. An empty-formula "remove" against a non-existent name is
  // a successful no-op.
  if (formula.empty()) {
    return Expected<void, Error>::Ok();
  }
  DefinedName entry;
  entry.name = std::move(name);
  entry.formula = std::move(formula);
  entry.local_sheet_id = local_sheet_id;
  defined_names_.push_back(std::move(entry));
  // Adding a name can make an already-calculated #NAME? formula resolvable.
  // Only a formula that actually mentions this name (or another name that
  // expands to it) could have produced that stale #NAME?, so scope the
  // re-register/dirty pass to that subset rather than every formula.
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_name(sheets_, mutator, *this, defined_names_.back().name);
  return Expected<void, Error>::Ok();
}

void Workbook::reindex_formulas_for_table_change(const std::vector<std::string>& table_names) {
  if (table_names.empty()) {
    return;
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_table(sheets_, mutator, *this, table_names);
}

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  for (const Sheet& s : sheets_) {
    if (sheet_names::equal(s.name(), name)) {
      return &s;
    }
  }
  return nullptr;
}

std::size_t Workbook::sheet_index_by_name(std::string_view name) const noexcept {
  for (std::size_t i = 0; i < sheets_.size(); ++i) {
    if (sheet_names::equal(sheets_[i].name(), name)) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}

std::size_t Workbook::approximate_memory_bytes() const noexcept {
  // Per-row cost of the cell store's hash map beyond the run itself: one
  // node holding the key and the `RowCells` header, plus the bucket slot
  // pointing at it. A constant is enough — the figure is a pressure
  // signal, and no standard library exposes its real per-node overhead.
  constexpr std::size_t kRowNodeOverheadBytes = sizeof(std::uint32_t) + 2U * sizeof(void*) + 32U;

  std::size_t total = sizeof(Workbook);

  for (const Sheet& sheet_ref : sheets_) {
    total += sizeof(Sheet) + sheet_ref.name().capacity();
    for (const auto& [row, cells] : sheet_ref.rows()) {
      (void)row;
      total += kRowNodeOverheadBytes + (cells.run().capacity() * sizeof(Cell));
      for (const Cell& cell : cells.run()) {
        total += cell.formula_text.capacity();
        total += cell.phonetic_runs.capacity() * sizeof(PhoneticRun);
        for (const PhoneticRun& run : cell.phonetic_runs) {
          total += run.text.capacity();
        }
        if (cell.cached_text_owned != nullptr) {
          total += sizeof(std::string) + cell.cached_text_owned->capacity();
        }
      }
    }
  }

  // Every `Text` value in the workbook is a view into this deque, so the
  // cell walk above deliberately does not add string payloads a second
  // time.
  for (const std::string& text : text_storage_) {
    total += sizeof(std::string) + text.capacity();
  }

  for (const PassthroughPart& part : passthrough_parts_) {
    total += sizeof(PassthroughPart) + part.path.capacity() + part.content_type.capacity() + part.bytes.capacity();
  }

  for (const DefinedName& name : defined_names_) {
    total += sizeof(DefinedName) + name.name.capacity() + name.formula.capacity() + name.comment.capacity();
  }

  total += workbook_pr_xml_.capacity() + book_views_xml_.capacity() + workbook_protection_xml_.capacity();

  return total;
}

void Workbook::add_pivot_cache(std::unique_ptr<pivot::PivotCache> cache) {
  if (cache == nullptr) {
    return;
  }
  pivot_caches_.push_back(std::move(cache));
}

const pivot::PivotCache* Workbook::find_pivot_cache(std::uint32_t cache_id) const noexcept {
  for (const std::unique_ptr<pivot::PivotCache>& c : pivot_caches_) {
    if (c != nullptr && c->cache_id() == cache_id) {
      return c.get();
    }
  }
  return nullptr;
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save() const {
  return save_as(WorkbookFormat::Ooxml);
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save_as(WorkbookFormat format) const {
  switch (format) {
    case WorkbookFormat::Ooxml:
      return io::write_ooxml(*this);
    case WorkbookFormat::Xlsb:
      return io::xlsb::write_xlsb(*this);
    case WorkbookFormat::Unknown:
      break;
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "Workbook::save_as: unsupported format",
                    "context=workbook_save_as");
}

namespace {

// Builds a CellNodeId for a workbook-relative coordinate. `sheet_index` is
// below `sheet_count()`, which every append path bounds by
// `Workbook::kMaxSheets`, so the narrowing to the dep graph's 16-bit sheet
// id is lossless.
eval::CellNodeId make_node(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  return eval::CellNodeId{static_cast<std::uint16_t>(sheet_index), row, col};
}

// Eagerly marks every existing dependent of `cell` dirty in `mutator`'s
// engine. The next `recalc()` pass would discover them via BFS anyway,
// but eager marking keeps the dirty set self-consistent between
// mutations and matches what callers see when they introspect the
// engine via `recalc_engine()`.
//
// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The helper routes through the facade so the
// compound mutation in `set_cell_value` / `set_cell_formula` stays
// under a single critical section, which keeps the `Sheet` write that
// follows from racing against a concurrent `recalc_parallel`.
void mark_dependents_dirty(const eval::RecalcEngine::LockedMutator& mutator, eval::CellNodeId cell) {
  for (eval::CellNodeId dep : mutator.dep_graph().dependents_of_ref(cell)) {
    mutator.mark_dirty(dep);
  }
  mutator.mark_range_dependents_dirty(cell);
}

// When `(row, col)` is a *phantom* of a committed spill region (i.e. a
// spill target that is not the anchor), a write here clears the region via
// `Sheet::set_cell_value` / `set_cell_formula` but leaves the anchor's
// cached value untouched. The anchor is not a dep-graph dependent of the
// phantom, so nothing else dirties it; mark it dirty so the next recalc
// re-evaluates it — re-spilling if the write vacated the cell, or surfacing
// `#SPILL!` if it now blocks the footprint. Must be called BEFORE the sheet
// write, while the region still covers the cell. Precondition: caller holds
// the engine mutex for the lifetime of the `LockedMutator&`.
void mark_spill_anchor_dirty_if_covered(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                        const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  const SpillRegion* covering = sheets[sheet_index].spill_region_covering(row, col);
  if (covering == nullptr) {
    return;
  }
  if (covering->anchor_row == row && covering->anchor_col == col) {
    return;  // Writing the anchor itself is already handled by the caller.
  }
  mutator.mark_dirty(make_node(sheet_index, covering->anchor_row, covering->anchor_col));
}

// A failed dynamic-array commit keeps its attempted rectangle even though no
// phantom cells were materialised.  A mutation that vacates any one cell in
// that rectangle must wake the blocked producer so it can retry on the next
// recalc pass.  The Sheet query returns coordinate copies while holding its
// own lock; marking happens after that lock is released, preserving the
// workbook's engine-mutex -> sheet/spill-mutex order.
void mark_blocked_spill_anchors_intersecting(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                             const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  for (const CellAddress anchor : sheets[sheet_index].blocked_spill_anchors_intersecting(row, col, 1U, 1U)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

void mark_blocked_spill_anchors_intersecting(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                             const std::vector<Sheet>& sheets, const BlockedSpillFootprint& rectangle) {
  for (const CellAddress anchor : sheets[sheet_index].blocked_spill_anchors_intersecting(
           rectangle.anchor_row, rectangle.anchor_col, rectangle.rows, rectangle.cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

void mark_blocked_spill_anchors_released_by_cell(const eval::RecalcEngine::LockedMutator& mutator,
                                                 std::size_t sheet_index, const std::vector<Sheet>& sheets,
                                                 std::uint32_t row, std::uint32_t col) {
  // A write into an existing spill clears the entire committed rectangle,
  // not just the addressed cell. Snapshot that rectangle before the write and
  // wake every pending producer intersecting it.
  if (const auto rectangle = sheets[sheet_index].committed_spill_footprint_covering(row, col); rectangle.has_value()) {
    mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, *rectangle);
  }
  mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, row, col);
}

void mark_blocked_spill_anchors_intersecting_merge(const eval::RecalcEngine::LockedMutator& mutator,
                                                   std::size_t sheet_index, const std::vector<Sheet>& sheets,
                                                   const MergeRange& merge) {
  if (merge.first_row > merge.last_row || merge.first_col > merge.last_col ||
      !Sheet::coord_in_grid(merge.first_row, merge.first_col) ||
      !Sheet::coord_in_grid(merge.last_row, merge.last_col)) {
    return;
  }
  const std::uint32_t rows = merge.last_row - merge.first_row + 1U;
  const std::uint32_t cols = merge.last_col - merge.first_col + 1U;
  for (const CellAddress anchor :
       sheets[sheet_index].blocked_spill_anchors_intersecting(merge.first_row, merge.first_col, rows, cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

void mark_committed_spill_anchors_intersecting_merge(const eval::RecalcEngine::LockedMutator& mutator,
                                                     std::size_t sheet_index, const std::vector<Sheet>& sheets,
                                                     const MergeRange& merge) {
  if (merge.first_row > merge.last_row || merge.first_col > merge.last_col ||
      !Sheet::coord_in_grid(merge.first_row, merge.first_col) ||
      !Sheet::coord_in_grid(merge.last_row, merge.last_col)) {
    return;
  }
  const std::uint32_t rows = merge.last_row - merge.first_row + 1U;
  const std::uint32_t cols = merge.last_col - merge.first_col + 1U;
  for (const CellAddress anchor :
       sheets[sheet_index].committed_spill_anchors_intersecting(merge.first_row, merge.first_col, rows, cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_value: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_value: coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // The cell is becoming a literal: drop outgoing edges (the old formula's
  // reads), but preserve incoming edges so other formulas that *read* this
  // cell continue to re-evaluate when the literal changes.
  // `clear_cell_dependencies` does exactly that — `unregister_formula`
  // would also remove the incoming edges, which is wrong here.
  //
  // Mark dependents *before* clearing the cell's outgoing edges so the
  // reverse-edge snapshot still describes the pre-clear graph. (Clearing
  // outgoing edges does not actually drop incoming edges, but doing the
  // mark first keeps the ordering robust against future API changes.)
  //
  // The entire compound mutation — three engine operations plus the
  // `Sheet` write — is performed under a single hold of the engine
  // mutex so a concurrent `recalc_parallel` (which holds the same
  // mutex for the duration of its pass) cannot observe a half-applied
  // edit. Going through the public engine API would release and
  // re-acquire the lock between every step and let the `Sheet` write
  // race against the recalc worker reading the same cell.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    mutator.mark_dirty(node);
    mark_dependents_dirty(mutator, node);
    mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets_, row, col);
    mutator.clear_cell_dependencies(node);
    mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets_, row, col);
    sheets_[sheet_index].set_cell_value(row, col, value);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              std::string_view text) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_text: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_text: coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  mutator.mark_dirty(node);
  mark_dependents_dirty(mutator, node);
  mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets_, row, col);
  mutator.clear_cell_dependencies(node);
  mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets_, row, col);
  sheets_[sheet_index].set_cell_text(row, col, text);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_formula: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_formula: coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }

  // Normalize Excel's `_xlfn.` / `_xlfn._xlws.` / `_xlws.` / `_xlpm.`
  // storage prefixes to the canonical formula-bar form at the single
  // ingestion point every reader (OOXML DOM / SAX, XLSB) and binding
  // funnels through. This keeps the stored `formula_text` (and the
  // dependency-extraction parse below) matching what Excel's formula bar
  // shows, and lets LET / LAMBDA resolve their `_xlpm.`-prefixed
  // parameter names. The transform is a no-op on an already-canonical
  // formula, so hand-authored / test formulas are unaffected. The writer
  // re-applies the prefixes on save for Excel readability.
  {
    std::string normalized = parser::strip_storage_prefixes(formula);
    if (normalized != formula) {
      formula = std::move(normalized);
    }
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // Parse the formula in a throwaway arena to extract its dependency list.
  // Strip a leading '=' so the parser sees a bare expression (the tokenizer
  // accepts both shapes, but the dep extractor walks the AST regardless).
  std::string_view src = formula;
  if (!src.empty() && src.front() == '=') {
    src.remove_prefix(1);
  }

  Arena tmp_arena;
  parser::AstNode* root = parser::parse_strict(src, tmp_arena);

  // The compound mutation runs under a single hold of the engine mutex
  // so a concurrent `recalc_parallel` does not see a half-applied
  // edit — see the comment in `set_cell_value` for the full rationale.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    // Query the covering spill BEFORE the sheet write clears it, so a write
    // into a live spill's phantom re-dirties the anchor.
    mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets_, row, col);
    mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets_, row, col);
    // Persist the formula text on the sheet first so a later `recalc()`
    // reads what the user actually typed. This also resets `cached_value`
    // to blank.
    sheets_[sheet_index].set_cell_formula(row, col, std::move(formula));

    if (root != nullptr) {
      mutator.register_formula(node, *root, *this);
    } else {
      // Hard parse failure, or a valid prefix trailed by unparseable
      // tokens. Drop any stale edges rather than register dependencies for
      // a recovered prefix that is not the whole formula; the cell surfaces
      // `#NAME?` at the next recalc (via the strict gate in cell_evaluator).
      mutator.unregister_formula(node);
    }

    // Mark the cell dirty and propagate to direct dependents.
    mutator.mark_dirty(node);
    mark_dependents_dirty(mutator, node);
  }
  return Expected<void, Error>::Ok();
}

void Workbook::mark_cell_dependents_dirty(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  if (sheet_index >= sheets_.size() || !Sheet::coord_in_grid(row, col)) {
    return;
  }
  const eval::CellNodeId node = make_node(sheet_index, row, col);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  mark_dependents_dirty(mutator, node);
}

Expected<void, Error> Workbook::add_merge(std::size_t sheet_index, MergeRange merge) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "add_merge: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  sheets_[sheet_index].add_merge(merge);
  // A merge added over a committed spill becomes a blocker for the next
  // evaluation. Keep the anchor dirty so recalc clears the old rectangle and
  // records the resulting #SPILL! state under the normal commit contract.
  mark_committed_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, merge);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_merges_intersecting(std::size_t sheet_index, MergeRange merge) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_merges_intersecting: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::vector<MergeRange> removed = sheets_[sheet_index].remove_merges_intersecting(merge);
  for (const MergeRange& erased : removed) {
    mark_blocked_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, erased);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_merge_at(std::size_t sheet_index, std::size_t index) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_merge_at: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  MergeRange removed;
  if (!sheets_[sheet_index].remove_merge_at(index, &removed)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_merge_at: index out of range",
                      "index=" + std::to_string(index));
  }
  mark_blocked_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, removed);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_merges(std::size_t sheet_index) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "clear_merges: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  for (const CellAddress anchor : sheets_[sheet_index].blocked_spill_anchors()) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
  sheets_[sheet_index].clear_merges();
  return Expected<void, Error>::Ok();
}

Expected<eval::RecalcStats, Error> Workbook::recalc(const eval::FunctionRegistry& registry) {
  return engine_->recalc(*this, registry);
}

Expected<void, Error> Workbook::recalc_parallel(const eval::FunctionRegistry& registry,
                                                const eval::SchedulerConfig& cfg, eval::SchedulerStats* stats) {
  return eval::recalc_parallel(*this, registry, cfg, stats);
}

void Workbook::set_iterative_options(IterativeOptions opts) {
  // Bound the iteration budget where it enters the model rather than at
  // each producer. The solver has no wall-clock limit, and cancellation
  // exists only when the host registered a progress callback, so an
  // unbounded count is an unrecoverable hang rather than a slow
  // calculation. Clamping here means every consumer of
  // `iterative_options()` reads a bounded value by construction — including
  // the single-cell self-reference driver in the tree walker, which runs
  // its own loop without any cyclic component ever forming and so would
  // not be covered by a bound enforced inside the solver.
  opts.max_iterations = std::min(opts.max_iterations, kMaxIterationsCap);
  engine_->set_iterative_options(opts);
}

const IterativeOptions& Workbook::iterative_options() const noexcept {
  return engine_->iterative_options();
}

Expected<eval::RecalcStats, Error> Workbook::partial_recalc(const eval::FunctionRegistry& registry,
                                                            const eval::SheetCellRange& viewport) {
  return engine_->partial_recalc(*this, registry, viewport);
}

void Workbook::set_iterative_progress(eval::IterativeProgressCb cb, void* user_data) noexcept {
  engine_->set_iterative_progress(cb, user_data);
}

void Workbook::mark_row_visibility_dependents_dirty() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  for (std::size_t sheet_idx = 0; sheet_idx < sheets_.size(); ++sheet_idx) {
    const Sheet& sheet = sheets_[sheet_idx];
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

void Workbook::mark_all_formulas_dirty() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  for (std::size_t sheet_idx = 0; sheet_idx < sheets_.size(); ++sheet_idx) {
    for (const auto& [row, cells] : sheets_[sheet_idx].rows()) {
      for (std::size_t col = 0; col < cells.size(); ++col) {
        if (cells[col].formula_text.empty()) {
          continue;
        }
        mutator.mark_dirty(
            eval::CellNodeId{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)});
      }
    }
  }
}

namespace {

// Computes the post-edit coordinates of a cell on the edited sheet.
// `kept == false` means the cell is dropped by the upcoming
// `Sheet::delete_rows/cols` (deletion fully covers the cell) or pushed
// past the sheet bound by an insert.
struct CellShift {
  bool kept = true;
  std::uint32_t new_row = 0;
  std::uint32_t new_col = 0;
};

CellShift shift_cell_coords_for_row_col_edit(parser::RowColAxis axis, parser::RowColEdit edit, std::uint32_t index,
                                             std::uint32_t count, std::uint32_t row, std::uint32_t col) {
  CellShift r;
  r.new_row = row;
  r.new_col = col;
  std::uint32_t& target = (axis == parser::RowColAxis::kRow) ? r.new_row : r.new_col;
  const std::uint32_t coord = target;
  const std::uint32_t bound = (axis == parser::RowColAxis::kRow) ? Sheet::kMaxRows : Sheet::kMaxCols;

  if (edit == parser::RowColEdit::kInsert) {
    if (coord >= index) {
      const std::uint64_t shifted = static_cast<std::uint64_t>(coord) + count;
      if (shifted >= bound) {
        r.kept = false;
      } else {
        target = static_cast<std::uint32_t>(shifted);
      }
    }
    return r;
  }

  if (coord >= index + count) {
    target = coord - count;
  } else if (coord >= index) {
    r.kept = false;
  }
  return r;
}

// Walks every formula cell in `sheets`, applying a row/col shift
// transform. The transform is rebuilt per sheet so its
// `local_means_target` flag tracks whether the current sheet's
// unqualified references should be rewritten — a local reference on
// the target sheet is in scope, but a local reference on any other
// sheet refers to that other sheet and must be left alone.
//
// The transformed AST drives both the rewritten formula text and the
// dep-graph re-registration; we keep the arena alive across both
// operations so `register_formula` can read the same nodes the
// formatter just emitted, instead of re-parsing the formatted output.
//
// Coordinate handling: cells on the target sheet have their keys shifted
// by the same transform as the AST refs (the physical move happens
// later in `Sheet::insert_rows` / `delete_rows`; here we re-key the
// dep-graph entries and dirty marks so they match the post-shift cell
// position). Without this re-key the dirty marks land at stale
// coordinates and recalc evaluates the wrong cell — visible as blank
// cached values on the band edge and `#REF!` on aggregators that span
// the band.
//
// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The helper drives a long per-cell loop of dep-graph
// mutations and must share the critical section that the surrounding
// `apply_row_col_edit_operation` takes, so the eventual
// `Sheet::insert_rows` / `delete_rows` does not race against a
// concurrent `recalc_parallel`.
void rewrite_formulas_for_row_col_edit(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                                       const Workbook& workbook, std::string_view target_sheet, parser::RowColAxis axis,
                                       parser::RowColEdit edit, std::uint32_t index, std::uint32_t count) {
  struct FormulaUpdate {
    Sheet* sheet = nullptr;
    eval::CellNodeId old_node;
    eval::CellNodeId new_node;
    bool dropped = false;
    std::string formula;
  };

  // Collect the complete mutation set before touching the graph. A single
  // structural edit can map an old key onto a key that is still occupied by
  // another formula (for example, delete row 1 moves B3 -> B2 while B2 is
  // not yet unregistered). DepGraph::remove_node is coordinate based, so
  // interleaving unregister/register can then delete the newly registered
  // node's edges. The two phases below make every old key absent before any
  // new key is registered, independent of unordered_map iteration order.
  std::vector<FormulaUpdate> updates;
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
    const bool local_means_target = sheet_names::equal(sheet.name(), target_sheet);
    const parser::RowColShiftTransform transform(target_sheet, axis, edit, index, count, local_means_target);

    std::vector<std::uint32_t> row_keys;
    row_keys.reserve(sheet.rows().size());
    for (const auto& kv : sheet.rows()) {
      row_keys.push_back(kv.first);
    }
    for (std::uint32_t row : row_keys) {
      const auto it = sheet.rows().find(row);
      if (it == sheet.rows().end()) {
        continue;
      }
      const RowCells& cells = it->second;
      for (std::size_t col = 0; col < cells.size(); ++col) {
        const Cell& cell = cells[col];
        if (cell.formula_text.empty()) {
          continue;
        }
        std::string_view body = cell.formula_text;
        bool had_equals = false;
        if (!body.empty() && body.front() == '=') {
          body = body.substr(1);
          had_equals = true;
        }
        if (body.empty()) {
          continue;
        }
        parser_arena.reset();
        parser::Parser parser(body, parser_arena);
        parser::AstNode* root = parser.parse();
        if (root == nullptr || !parser.errors().empty()) {
          continue;  // Unparseable formula; leave alone (matches Excel "carry through unchanged").
        }
        const parser::AstNode* shifted = parser::shift_refs(*root, parser_arena, transform);
        const bool ast_changed = (shifted != nullptr && shifted != root);

        // Resolve the cell's post-shift coordinates. Only cells on the
        // target sheet move; off-target cells keep their (row, col).
        std::uint32_t new_row = row;
        std::uint32_t new_col = static_cast<std::uint32_t>(col);
        bool dropped = false;
        if (local_means_target) {
          const CellShift sr =
              shift_cell_coords_for_row_col_edit(axis, edit, index, count, row, static_cast<std::uint32_t>(col));
          if (!sr.kept) {
            dropped = true;
          } else {
            new_row = sr.new_row;
            new_col = sr.new_col;
          }
        }
        const bool cell_moves = (new_row != row) || (new_col != static_cast<std::uint32_t>(col));

        if (!ast_changed && !cell_moves && !dropped) {
          continue;  // Cell unaffected by the edit.
        }

        // Keep the rewritten text alongside the key transition. It is
        // applied only after every disappearing / moving old key has been
        // unregistered.
        std::string effective_formula = cell.formula_text;
        if (ast_changed) {
          effective_formula.clear();
          if (had_equals) {
            effective_formula.push_back('=');
          }
          effective_formula.append(parser::format_formula(*shifted));
        }

        const eval::CellNodeId old_node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        const eval::CellNodeId new_node{static_cast<std::uint16_t>(sheet_idx), new_row, new_col};
        updates.push_back(FormulaUpdate{&sheet, old_node, new_node, dropped, std::move(effective_formula)});
      }
    }
  }

  // Phase 1: no old graph key may survive a coordinate transition.
  for (const FormulaUpdate& update : updates) {
    if (update.dropped || update.old_node != update.new_node) {
      mutator.unregister_formula(update.old_node);
    }
  }

  // Phase 2: apply formula text and create the post-edit graph. Re-parse the
  // stored text because the first-pass AST arenas intentionally expired
  // before phase 1; this keeps memory bounded for large workbooks.
  for (FormulaUpdate& update : updates) {
    if (update.dropped) {
      continue;
    }
    const Cell* current = update.sheet->cell_at(update.old_node.row, update.old_node.col);
    if (current == nullptr || current->formula_text != update.formula) {
      update.sheet->set_cell_formula(update.old_node.row, update.old_node.col, std::move(update.formula));
    }
    const Cell* formula_cell = update.sheet->cell_at(update.old_node.row, update.old_node.col);
    if (formula_cell == nullptr || formula_cell->formula_text.empty()) {
      continue;
    }
    std::string_view body = formula_cell->formula_text;
    if (!body.empty() && body.front() == '=') {
      body.remove_prefix(1);
    }
    parser_arena.reset();
    parser::Parser parser(body, parser_arena);
    parser::AstNode* root = parser.parse();
    if (root == nullptr || !parser.errors().empty()) {
      continue;
    }
    mutator.register_formula(update.new_node, *root, workbook);
    mutator.mark_dirty(update.new_node);
  }
}

// Re-register the owners of Ref3D spans after the target sheet's cell store
// has completed its physical move. Their shared inner coordinates intentionally
// remain unchanged, so the ordinary AST rewrite cannot wake them. The owner
// snapshot is taken before rewrite/unregister; map only owners that survive
// the edit, then parse the final formula text against the final workbook
// coordinates so defined-name and nested-Lambda expansion follows the same
// extractor path as ordinary registration.
void reregister_three_d_span_owners_after_row_col_edit(const std::vector<eval::CellNodeId>& owners,
                                                       std::vector<Sheet>& sheets,
                                                       const eval::RecalcEngine::LockedMutator& mutator,
                                                       const Workbook& workbook, std::size_t edited_sheet_index,
                                                       parser::RowColAxis axis, parser::RowColEdit edit,
                                                       std::uint32_t index, std::uint32_t count) {
  Arena parser_arena;
  for (const eval::CellNodeId owner : owners) {
    eval::CellNodeId mapped = owner;
    if (owner.sheet_id == edited_sheet_index) {
      const CellShift shift = shift_cell_coords_for_row_col_edit(axis, edit, index, count, owner.row, owner.col);
      if (!shift.kept) {
        continue;
      }
      mapped.row = shift.new_row;
      mapped.col = shift.new_col;
    }
    if (mapped.sheet_id >= sheets.size()) {
      continue;
    }
    const Cell* formula_cell = sheets[mapped.sheet_id].cell_at(mapped.row, mapped.col);
    if (formula_cell == nullptr || formula_cell->formula_text.empty()) {
      continue;
    }
    std::string_view body = formula_cell->formula_text;
    if (!body.empty() && body.front() == '=') {
      body.remove_prefix(1);
    }
    if (body.empty()) {
      continue;
    }
    parser_arena.reset();
    parser::AstNode* root = parser::parse_strict(body, parser_arena);
    if (root == nullptr) {
      continue;
    }
    mutator.register_formula(mapped, *root, workbook);
    mutator.mark_dirty(mapped);
  }
}

std::vector<BlockedSpillFootprint> remap_blocked_spill_footprints(const std::vector<BlockedSpillFootprint>& footprints,
                                                                  parser::RowColAxis axis, parser::RowColEdit edit,
                                                                  std::uint32_t index, std::uint32_t count) {
  std::vector<BlockedSpillFootprint> mapped;
  mapped.reserve(footprints.size());
  const std::uint32_t bound = axis == parser::RowColAxis::kRow ? Sheet::kMaxRows : Sheet::kMaxCols;
  const std::uint64_t delete_end = static_cast<std::uint64_t>(index) + count;
  for (const BlockedSpillFootprint& original : footprints) {
    std::uint32_t coordinate = axis == parser::RowColAxis::kRow ? original.anchor_row : original.anchor_col;
    if (edit == parser::RowColEdit::kDelete) {
      if (static_cast<std::uint64_t>(coordinate) >= index && static_cast<std::uint64_t>(coordinate) < delete_end) {
        continue;
      }
      if (static_cast<std::uint64_t>(coordinate) >= delete_end) {
        coordinate -= count;
      }
    } else if (coordinate >= index) {
      const std::uint64_t shifted = static_cast<std::uint64_t>(coordinate) + count;
      if (shifted >= bound) {
        continue;
      }
      coordinate = static_cast<std::uint32_t>(shifted);
    }
    BlockedSpillFootprint next = original;
    next.anchor_row = axis == parser::RowColAxis::kRow ? coordinate : original.anchor_row;
    next.anchor_col = axis == parser::RowColAxis::kRow ? original.anchor_col : coordinate;
    mapped.push_back(next);
  }
  return mapped;
}

Expected<void, Error> apply_row_col_edit(Workbook& wb, std::size_t sheet_index, parser::RowColAxis axis,
                                         parser::RowColEdit edit, std::uint32_t origin, std::uint32_t count,
                                         const char* op_name) {
  if (sheet_index >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  if (count == 0U) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": count must be >= 1");
  }
  const std::uint32_t bound = (axis == parser::RowColAxis::kRow) ? Sheet::kMaxRows : Sheet::kMaxCols;
  if (origin >= bound) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": origin out of bounds",
                      "origin=" + std::to_string(origin));
  }
  // Deletion helpers use `origin + count` as their exclusive endpoint.
  // Validate against the remaining grid with subtraction so an unsigned
  // count received through the C/WASM ABI cannot wrap and turn a delete
  // into a corrupting row/column move.
  if (edit == parser::RowColEdit::kDelete && count > bound - origin) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op_name) + ": count exceeds sheet bounds",
                      "origin=" + std::to_string(origin) + " count=" + std::to_string(count));
  }
  return Expected<void, Error>::Ok();
}

// Precondition: caller holds the engine mutex for the lifetime of the
// `LockedMutator&`. The compound edit (per-cell dep-graph re-keys plus
// the eventual `Sheet::insert_rows` / `delete_rows` / `insert_cols` /
// `delete_cols`) runs entirely under that single critical section so
// a concurrent `recalc_parallel` either sees the pre-edit state or
// the fully patched one, never a half-applied edit.
// `rewrite_defined_names` does not touch the engine but is included
// here to preserve the "defined-name table matches dep-graph state"
// invariant for any other reader that consults both under the same
// lock, and because the dep-graph re-index it can trigger has to run
// against the rewritten table.
Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                   const eval::RecalcEngine::LockedMutator& mutator,
                                                   std::vector<DefinedName>& defined_names, std::size_t sheet_index,
                                                   parser::RowColAxis axis, parser::RowColEdit edit,
                                                   std::uint32_t origin, std::uint32_t count, const char* op_name) {
  RETURN_IF_ERROR(apply_row_col_edit(wb, sheet_index, axis, edit, origin, count, op_name));
  const std::string target_sheet_name = sheets[sheet_index].name();
  // `apply_row_col_edit` above rejected an out-of-range `sheet_index`, and
  // `Workbook::kMaxSheets` bounds `sheet_count()`, so the narrowing to the
  // dep graph's 16-bit sheet id keeps the index intact.
  const std::vector<eval::CellNodeId> three_d_owners =
      mutator.three_d_span_owners_covering_sheet(static_cast<std::uint16_t>(sheet_index));
  // Defined names are rewritten before cell formulas, and a change to any of
  // them forces a full re-index afterwards. A formula that reaches a shifted
  // range only through a name — `=MyRef*1` — is textually unchanged by the
  // shift, because a `NameRef` node is the identity case for the transform.
  // The per-formula rewriter only re-registers formulas whose text changed,
  // so without the re-index that formula keeps dep-graph edges pointing at
  // the range `MyRef` used to cover. `set_defined_name_scoped` and
  // `remove_sheet` already take this fallback for the same reason.
  const parser::RowColShiftTransform name_transform(target_sheet_name, axis, edit, origin, count);
  const bool names_changed = rewrite_defined_names(defined_names, name_transform);
  // Snapshot the pending-footprint records before formula text rewrites.
  // `Sheet::set_cell_formula` quite correctly clears a user-overwritten
  // blocked anchor; a structural rewrite of that same formula is not a user
  // overwrite, so restore the mapped record after the physical move below.
  // Keep the other sheets too: a name rewrite can force the full graph reset,
  // whose spill invalidation must not silently discard unrelated pending
  // producers.
  std::vector<std::vector<BlockedSpillFootprint>> blocked_before_all;
  blocked_before_all.reserve(sheets.size());
  for (const Sheet& sheet : sheets) {
    blocked_before_all.push_back(sheet.blocked_spill_footprints());
  }
  Sheet& target = sheets[sheet_index];
  const std::vector<BlockedSpillFootprint> blocked_before = blocked_before_all[sheet_index];
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, axis, edit, origin, count);
  // Metadata formulas follow the same per-sheet policy as cell formulas: a
  // qualified reference to the edited sheet shifts no matter which sheet owns
  // the rule, while an unqualified one is in scope only on the edited sheet.
  // The text is rewritten against pre-edit coordinates, like the cells above
  // and unlike the sqref/anchor rectangles, which `Sheet::insert_rows` and
  // friends move afterwards.
  const parser::RowColShiftTransform local_transform(target_sheet_name, axis, edit, origin, count,
                                                     /*local_means_target=*/true);
  std::vector<const parser::RefTransform*> per_sheet_transforms(sheets.size(), &name_transform);
  per_sheet_transforms[sheet_index] = &local_transform;
  std::vector<std::uint32_t> ignored_cache_ids;
  rewrite_sheet_metadata_formulas(sheets, per_sheet_transforms, wb.mutable_tables(), wb.mutable_pivot_caches(), {}, {},
                                  {}, ignored_cache_ids);
  if (axis == parser::RowColAxis::kRow) {
    if (edit == parser::RowColEdit::kInsert) {
      target.insert_rows(origin, count);
    } else {
      target.delete_rows(origin, count);
    }
  } else if (edit == parser::RowColEdit::kInsert) {
    target.insert_cols(origin, count);
  } else {
    target.delete_cols(origin, count);
  }
  // Re-index only after the physical move. Formula text is rewritten while
  // cells still occupy their pre-edit coordinates, so rebuilding the graph
  // before the move would register moved owners (including Ref3D owners) at
  // stale coordinates and leave duplicate registry entries behind.
  if (names_changed) {
    reindex_all_formulas(sheets, mutator, wb);
    for (std::size_t preserved_sheet = 0; preserved_sheet < sheets.size(); ++preserved_sheet) {
      if (preserved_sheet == sheet_index) {
        continue;
      }
      sheets[preserved_sheet].restore_blocked_spill_footprints(std::move(blocked_before_all[preserved_sheet]));
    }
  }
  reregister_three_d_span_owners_after_row_col_edit(three_d_owners, sheets, mutator, wb, sheet_index, axis, edit,
                                                    origin, count);
  std::vector<BlockedSpillFootprint> blocked_mapped =
      remap_blocked_spill_footprints(blocked_before, axis, edit, origin, count);
  // Do not resurrect a record whose formula anchor was deleted or moved past
  // the grid. The cell store has already completed the physical move, so a
  // formula check is stable while the engine mutex is held.
  blocked_mapped.erase(std::remove_if(blocked_mapped.begin(), blocked_mapped.end(),
                                      [&](const auto& footprint) {
                                        const Cell* cell = target.cell_at(footprint.anchor_row, footprint.anchor_col);
                                        return cell == nullptr || cell->formula_text.empty();
                                      }),
                       blocked_mapped.end());
  target.restore_blocked_spill_footprints(std::move(blocked_mapped));
  // Sheet-local row/column edits remap the pending spill-footprint reverse
  // index alongside formula cells.  Dirty the surviving anchors at their new
  // coordinates; the next recalc retries each formula against the moved
  // blocker/footprint. Anchors deleted or shifted past the grid were dropped
  // by the Sheet remap and therefore do not appear here.
  for (const CellAddress anchor : target.blocked_spill_anchors()) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<void, Error> Workbook::insert_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count) {
  // The whole edit — including the per-cell dep-graph re-keys done by
  // `rewrite_formulas_for_row_col_edit` and the eventual
  // `Sheet::insert_rows` — must run under a single hold of the engine
  // mutex. A concurrent `recalc_parallel` holds the same mutex for its
  // full pass, so the compound mutation either runs entirely before
  // or entirely after a recalc, never half-applied alongside it. The
  // `LockedMutator` facade routes through the engine's `*_locked` API
  // and assumes this lock is already held.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kRow,
                                      parser::RowColEdit::kInsert, row, count, "insert_rows");
}

Expected<void, Error> Workbook::delete_rows(std::size_t sheet_index, std::uint32_t row, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kRow,
                                      parser::RowColEdit::kDelete, row, count, "delete_rows");
}

Expected<void, Error> Workbook::insert_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kCol,
                                      parser::RowColEdit::kInsert, col, count, "insert_cols");
}

Expected<void, Error> Workbook::delete_cols(std::size_t sheet_index, std::uint32_t col, std::uint32_t count) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  return apply_row_col_edit_operation(*this, sheets_, mutator, defined_names_, sheet_index, parser::RowColAxis::kCol,
                                      parser::RowColEdit::kDelete, col, count, "delete_cols");
}

Expected<void, Error> Workbook::set_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                  std::uint32_t xf_index) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_xf_index: sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_cell_xf_index: coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  // A style write can grow the sheet's sparse row store, so serialize it
  // with recalc just like all other workbook-level cell mutations. The
  // sheet-level setter deliberately bypasses literal-write spill invalidation
  // so formatting a live spill phantom preserves the dynamic array.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].set_cell_xf_index(row, col, xf_index);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
