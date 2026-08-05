//
// Workbook implementation. Wires the OOXML save path and the embedded
// recalc engine. The `RecalcEngine` is held via `unique_ptr` (PIMPL-style)
// so the public header does not need to include the heavyweight recalc /
// dep-graph headers.

#include "workbook.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "eval/utf8_length.h"
#include "io/defined_names.h"
#include "io/external_links.h"
#include "io/format_detect.h"
#include "io/formula_prefix.h"
#include "io/ooxml_writer.h"
#include "io/passthrough_part.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/xlsb/writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace {
// Forward declaration; defined further below alongside the other
// sheet-name helpers. Declared here so the early `add_sheet_validated`
// definition can call it.
Expected<void, Error> validate_sheet_name(std::string_view name);
}  // namespace

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()), kind_(io::WorkbookKind::kXlsx) {}
Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;
Workbook::~Workbook() = default;

Workbook Workbook::create() {
  Workbook wb;
  wb.sheets_.emplace_back(Sheet{std::string("Sheet1")});
  return wb;
}

Workbook Workbook::create_empty() {
  // No default sheet; callers are expected to populate via add_sheet().
  return Workbook{};
}

std::string_view Workbook::intern_text(std::string_view text) {
  text_storage_.emplace_back(text.data(), text.size());
  return std::string_view(text_storage_.back());
}

Sheet& Workbook::add_sheet(std::string name) {
  sheets_.emplace_back(Sheet{std::move(name)});
  return sheets_.back();
}

Expected<Sheet*, Error> Workbook::add_sheet_validated(std::string name) {
  RETURN_IF_ERROR(validate_sheet_name(name));
  for (const Sheet& existing : sheets_) {
    if (strings::case_insensitive_eq(existing.name(), name)) {
      return make_error(FormulonErrorCode::kInvalidSheetName, "add_sheet: name collides with an existing sheet",
                        "name=\"" + name + "\"");
    }
  }
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
void reindex_all_formulas(const std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                          const Workbook& workbook) {
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

// Forward declaration: the text-level sheet-rename rewriter lives further
// down in this anonymous namespace but is needed by the cell rewriter.
std::string replace_sheet_in_formula(std::string_view formula, std::string_view old_name, std::string_view new_name);
bool formula_references_sheet(std::string_view formula, std::string_view sheet_name) noexcept;

// Rewrites every cell formula that references `old_name` to `new_name`
// after a sheet rename. Cell formulas store sheet qualifiers by name and
// resolve them at evaluation time, so without this rewrite a renamed
// sheet's dependents keep the stale name and surface `#REF!` / `#NAME?` on
// the next recalc. Each rewritten cell is re-registered (its
// `CellNodeId.sheet_id` is unchanged by a rename, but re-registration
// keeps the graph in lockstep with the new text) and marked dirty so the
// blanked cached value is restored by the next recalc.
//
// Precondition: the caller holds the engine mutex for the lifetime of the
// `LockedMutator&`.
void rewrite_cell_formulas_for_sheet_rename(std::vector<Sheet>& sheets,
                                            const eval::RecalcEngine::LockedMutator& mutator, const Workbook& workbook,
                                            std::string_view old_name, std::string_view new_name) {
  Arena parser_arena;
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    Sheet& sheet = sheets[sheet_idx];
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
      const std::size_t col_count = it->second.size();
      for (std::size_t col = 0; col < col_count; ++col) {
        const Cell& cell = sheet.rows().at(row)[col];
        if (cell.formula_text.empty() || !formula_references_sheet(cell.formula_text, old_name)) {
          continue;
        }
        std::string rewritten = replace_sheet_in_formula(cell.formula_text, old_name, new_name);
        if (rewritten == cell.formula_text) {
          continue;  // Name matched a substring only; nothing to rewrite.
        }
        sheet.set_cell_formula(row, static_cast<std::uint32_t>(col), std::move(rewritten));

        const eval::CellNodeId node{static_cast<std::uint16_t>(sheet_idx), row, static_cast<std::uint32_t>(col)};
        std::string_view body = sheet.rows().at(row)[col].formula_text;
        if (!body.empty() && body.front() == '=') {
          body.remove_prefix(1);
        }
        if (!body.empty()) {
          parser_arena.reset();
          parser::AstNode* root = parser::parse_strict(body, parser_arena);
          if (root != nullptr) {
            mutator.register_formula(node, *root, workbook);
          }
        }
        mutator.mark_dirty(node);
      }
    }
  }
}

// Excel's structural validation for sheet names: non-empty, ≤ 31
// characters, no `: \ / ? * [ ]`. The forbidden-character scan is a
// byte-level check (the forbidden set is ASCII), so it works correctly on
// UTF-8 sheet names because none of the disallowed code units appear as
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
// surface (add / rename / any future import-side check). Verifies the
// name is non-empty, within the code-unit length limit, and free of
// forbidden characters. Duplicate/case-folding collision is caller-scoped
// (it needs the target index) and handled at each call site.
Expected<void, Error> validate_sheet_name(std::string_view name) {
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name must not be empty", "name=\"\"");
  }
  if (eval::utf16_units_in(name) > kMaxSheetNameUnits) {
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

// Quotes `sheet` for use in a formula reference. Excel quotes a sheet
// name with single quotes when the name contains anything other than
// alphanumerics, underscores or periods. We err on the side of always
// considering both shapes (raw and quoted) when matching defined-name
// targets so a rename catches both.
std::string single_quote_sheet(std::string_view sheet) {
  std::string out;
  out.reserve(sheet.size() + 4U);
  out.push_back('\'');
  for (char byte : sheet) {
    if (byte == '\'') {
      // Embedded single quotes double up in OOXML formula syntax.
      out.append("''");
    } else {
      out.push_back(byte);
    }
  }
  out.push_back('\'');
  return out;
}

// Returns true if `byte` is part of an unquoted Excel identifier
// (letter, digit, underscore). UTF-8 continuation bytes (high bit set)
// are deliberately excluded so a sheet name whose preceding byte is a
// multi-byte CJK character still treats the sheet token as starting
// fresh.
constexpr unsigned char kAsciiHighBit = 0x80U;
constexpr bool is_id_byte(char byte) noexcept {
  const auto value = static_cast<unsigned char>(byte);
  if (value >= kAsciiHighBit) {
    return false;
  }
  return (byte >= 'A' && byte <= 'Z') || (byte >= 'a' && byte <= 'z') || (byte >= '0' && byte <= '9') || byte == '_';
}

// Returns true iff `formula` references the sheet `sheet_name` by name
// (either bare `Sheet1!` or quoted `'Sheet One'!`). Strict prefix match
// of the bang-suffixed sheet token followed by a `!`. The check is
// case-insensitive, mirroring Excel's sheet-name resolution semantics.
//
// This is a textual scan deliberately kept out of the parser: the
// AST-level reference shifter (a separate follow-up bundle) handles the
// general case. For this bundle we only need to decide whether a
// defined-name target string mentions the renamed sheet at all so the
// bulk-replace helper below knows whether the entry is affected.
bool formula_references_sheet(std::string_view formula, std::string_view sheet_name) noexcept {
  // Exhaustive bare-token scan: walk every position and check whether a
  // sheet token followed by `!` matches `sheet_name` (case-insensitive).
  // Quoted-sheet matches are handled by also scanning for the
  // single-quoted form.
  const std::string quoted = single_quote_sheet(sheet_name);
  for (std::size_t pos = 0; pos + sheet_name.size() < formula.size(); ++pos) {
    if (formula[pos + sheet_name.size()] != '!') {
      continue;
    }
    if (strings::case_insensitive_eq(formula.substr(pos, sheet_name.size()), sheet_name)) {
      // Reject matches that are part of a longer identifier (e.g. the
      // suffix of `OtherSheet1` should not match `Sheet1`). A sheet
      // token starts at pos==0 or after a non-identifier byte.
      const bool token_start = pos == 0 || !is_id_byte(formula[pos - 1]);
      if (token_start) {
        return true;
      }
    }
  }
  if (quoted.size() + 1U <= formula.size()) {
    for (std::size_t pos = 0; pos + quoted.size() < formula.size(); ++pos) {
      if (formula[pos + quoted.size()] != '!') {
        continue;
      }
      if (strings::case_insensitive_eq(formula.substr(pos, quoted.size()), quoted)) {
        return true;
      }
    }
  }
  return false;
}

// Rewrites every reference to sheet `old_name` in `formula` to
// `new_name` by parsing through the AST and re-formatting. Both bare
// and quoted forms are recognised by the parser; the formatter emits
// the canonical shape (quoted only when the new name's bytes demand
// it). On parse failure or arena exhaustion the input is returned
// unchanged so the caller's behaviour stays conservative — a formula
// we cannot understand is left alone rather than corrupted.
std::string replace_sheet_in_formula(std::string_view formula, std::string_view old_name, std::string_view new_name) {
  // Defined-name targets are typically authored without a leading `=`;
  // tolerate either by parsing the body and prepending the marker back
  // when the source had one.
  std::string_view body = formula;
  bool had_equals = false;
  if (!body.empty() && body.front() == '=') {
    body = body.substr(1);
    had_equals = true;
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return std::string(formula);
  }
  const parser::SheetRenameTransform transform(old_name, new_name);
  const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    return std::string(formula);
  }
  std::string out;
  if (had_equals) {
    out.push_back('=');
  }
  out.append(parser::format_formula(*shifted));
  return out;
}

std::string replace_sheet_in_hyperlink_location(std::string_view location, std::string_view old_name,
                                                std::string_view new_name) {
  if (location.empty()) {
    return std::string(location);
  }
  const bool has_fragment_prefix = location.front() == '#';
  const std::string_view formula = has_fragment_prefix ? location.substr(1) : location;
  if (!formula_references_sheet(formula, old_name)) {
    return std::string(location);
  }
  std::string rewritten = replace_sheet_in_formula(formula, old_name, new_name);
  if (has_fragment_prefix) {
    rewritten.insert(rewritten.begin(), '#');
  }
  return rewritten;
}

void rewrite_sheet_named_metadata(std::vector<Sheet>& sheets, std::vector<io::DefinedName>& defined_names,
                                  Workbook& workbook, std::string_view old_name, std::string_view new_name) {
  for (io::DefinedName& entry : defined_names) {
    if (formula_references_sheet(entry.formula, old_name)) {
      entry.formula = replace_sheet_in_formula(entry.formula, old_name, new_name);
    }
  }
  for (Sheet& sheet : sheets) {
    for (Hyperlink& hyperlink : sheet.mutable_hyperlinks()) {
      hyperlink.location = replace_sheet_in_hyperlink_location(hyperlink.location, old_name, new_name);
    }
    for (DataValidation& validation : sheet.mutable_validations()) {
      if (formula_references_sheet(validation.formula1, old_name)) {
        validation.formula1 = replace_sheet_in_formula(validation.formula1, old_name, new_name);
      }
      if (formula_references_sheet(validation.formula2, old_name)) {
        validation.formula2 = replace_sheet_in_formula(validation.formula2, old_name, new_name);
      }
    }
  }
  for (std::unique_ptr<pivot::PivotCache>& cache : workbook.mutable_pivot_caches()) {
    if (cache == nullptr) {
      continue;
    }
    pivot::WorksheetSource& source = cache->mutable_worksheet_source();
    if (strings::case_insensitive_eq(source.sheet, old_name)) {
      source.sheet.assign(new_name);
    }
    if (formula_references_sheet(source.ref, old_name)) {
      source.ref = replace_sheet_in_formula(source.ref, old_name, new_name);
    }
  }
}

void freeze_cell_formulas_referencing_sheet(std::vector<Sheet>& sheets, std::uint32_t removed_index,
                                            std::string_view removed_name) {
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    if (sheet_idx == removed_index) {
      continue;
    }
    Sheet& sheet = sheets[sheet_idx];
    for (const auto& row_entry : sheet.rows()) {
      const std::uint32_t row = row_entry.first;
      for (std::size_t col = 0; col < row_entry.second.size(); ++col) {
        const Cell& cell = row_entry.second[col];
        if (!cell.formula_text.empty() && formula_references_sheet(cell.formula_text, removed_name)) {
          sheet.set_cell_formula(row, static_cast<std::uint32_t>(col), "=#REF!");
        }
      }
    }
  }
}

void remove_and_reindex_tables(std::vector<io::TableMetadata>& tables, std::size_t removed_sheet_index) {
  std::vector<io::TableMetadata> retained;
  retained.reserve(tables.size());
  for (io::TableMetadata& table : tables) {
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

void remove_pivot_caches_referencing_sheet(std::vector<std::unique_ptr<pivot::PivotCache>>& caches,
                                           std::string_view removed_sheet_name) {
  std::vector<std::unique_ptr<pivot::PivotCache>> retained;
  retained.reserve(caches.size());
  for (std::unique_ptr<pivot::PivotCache>& cache : caches) {
    if (cache == nullptr) {
      continue;
    }
    const pivot::WorksheetSource& source = cache->worksheet_source();
    // A cache whose worksheet source was removed can no longer refresh and
    // is not valid for any surviving pivot table. Source-less caches remain
    // supported as deliberate static snapshots.
    if (source.present && (strings::case_insensitive_eq(source.sheet, removed_sheet_name) ||
                           formula_references_sheet(source.ref, removed_sheet_name))) {
      continue;
    }
    retained.push_back(std::move(cache));
  }
  caches = std::move(retained);
}

void move_table_sheet_indices(std::vector<io::TableMetadata>& tables, std::size_t from_index, std::size_t to_index) {
  for (io::TableMetadata& table : tables) {
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
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "rename_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Validate the new name (non-empty, ≤ 31 code units, no forbidden
  // characters) via the shared validator so add / rename agree.
  RETURN_IF_ERROR(validate_sheet_name(new_name));
  // Collision check (case-insensitive). A no-op rename — same case-fold
  // as the current name — bypasses the collision check so callers can
  // change only the casing.
  for (std::size_t idx = 0; idx < sheets_.size(); ++idx) {
    if (idx == static_cast<std::size_t>(index)) {
      continue;
    }
    if (strings::case_insensitive_eq(sheets_[idx].name(), new_name)) {
      return make_error(FormulonErrorCode::kInvalidSheetName, "rename_sheet: name collides with another sheet",
                        "name=\"" + new_name + "\" colliding_index=" + std::to_string(idx));
    }
  }

  const std::string old_name = sheets_[index].name();

  // Rewrite cell formulas that reference the renamed sheet, and apply the
  // sheet rename, under a single hold of the engine mutex so a concurrent
  // `recalc_parallel` never observes a half-renamed state (some formulas
  // still naming the old sheet while the sheet already carries the new
  // name). See `set_cell_value` for the full rationale.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    rewrite_cell_formulas_for_sheet_rename(sheets_, mutator, *this, old_name, new_name);
    rewrite_sheet_named_metadata(sheets_, defined_names_, *this, old_name, new_name);
    // Move the rename into the sheet last so the rewriter above still has
    // the pre-move name state to read from.
    sheets_[index].set_name(std::move(new_name));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_sheet(std::uint32_t index) {
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "remove_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (sheets_.size() <= 1U) {
    return make_error(FormulonErrorCode::kCannotRemoveLastSheet, "remove_sheet: cannot remove the only sheet",
                      "sheet_count=" + std::to_string(sheets_.size()));
  }

  const std::string removed_name = sheets_[index].name();

  // Erase the sheet, then rebuild the dependency graph from scratch. A
  // remove shifts the workbook-relative index of every sheet after the
  // removed one, so their `CellNodeId.sheet_id`s — and the graph edges
  // keyed by them — are all invalidated at once; a full re-registration
  // is the safe baseline. Both halves run under a single hold of the
  // engine mutex so a concurrent `recalc_parallel` either sees the sheet
  // (and its graph nodes) fully present or fully gone, never a
  // half-erased intermediate.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    freeze_cell_formulas_referencing_sheet(sheets_, index, removed_name);
    sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(index));
    reindex_all_formulas(sheets_, mutator, *this);
  }

  // Drop defined names that target the removed sheet. We compare against
  // the removed sheet's name verbatim; sheet-scoped names whose
  // local_sheet_id matches the removed index are also dropped (the
  // index they referenced no longer exists). Note: indices for
  // sheet-scoped names that point at sheets *after* the removed one are
  // shifted down by 1 to preserve the binding. This matches Excel's UI
  // observation that sheet-scoped names stay attached to the same sheet
  // as the workbook is rearranged.
  std::vector<io::DefinedName> retained;
  retained.reserve(defined_names_.size());
  for (io::DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id == static_cast<std::int32_t>(index)) {
      continue;
    }
    if (formula_references_sheet(entry.formula, removed_name)) {
      // Keep surviving names but freeze their target. This prevents a later
      // newly-created sheet with the same name from silently rebinding it.
      entry.formula = "#REF!";
    }
    if (entry.local_sheet_id > static_cast<std::int32_t>(index)) {
      entry.local_sheet_id -= 1;
    }
    retained.push_back(std::move(entry));
  }
  defined_names_ = std::move(retained);
  remove_and_reindex_tables(tables_, index);
  remove_pivot_caches_referencing_sheet(pivot_caches_, removed_name);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::move_sheet(std::uint32_t from_index, std::uint32_t to_index) {
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
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  Sheet moving = std::move(sheets_[from_index]);
  sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(from_index));
  sheets_.insert(sheets_.begin() + static_cast<std::ptrdiff_t>(to_index), std::move(moving));

  // Patch sheet-scoped defined names so their `local_sheet_id` follows
  // the rearrangement. Workbook-scoped names reference sheets by name,
  // so they need no update.
  const auto from_id = static_cast<std::int32_t>(from_index);
  const auto to_id = static_cast<std::int32_t>(to_index);
  for (io::DefinedName& entry : defined_names_) {
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
      // that references it resolves to, so their dep-graph edges (extracted
      // by expanding the old definition) and cached values are now stale.
      // Rebuild the graph from the current definitions and mark all formulas
      // dirty so the next recalc re-resolves the name. Redefinition is a rare
      // user edit; workbook load appends fresh unique names and never reaches
      // this branch, so the load path keeps its per-name cost.
      {
        std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
        const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
        reindex_all_formulas(sheets_, mutator, *this);
      }
      return Expected<void, Error>::Ok();
    }
  }
  // No existing entry; appending only makes sense if a formula is
  // supplied. An empty-formula "remove" against a non-existent name is
  // a successful no-op.
  if (formula.empty()) {
    return Expected<void, Error>::Ok();
  }
  io::DefinedName entry;
  entry.name = std::move(name);
  entry.formula = std::move(formula);
  entry.local_sheet_id = local_sheet_id;
  defined_names_.push_back(std::move(entry));
  // Adding a name can make an already-calculated #NAME? formula resolvable.
  // Its old dependency entry was extracted without this definition, so use
  // the same full rebuild as name updates and removals before the next
  // recalc.
  {
    std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
    const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
    reindex_all_formulas(sheets_, mutator, *this);
  }
  return Expected<void, Error>::Ok();
}

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  for (const Sheet& s : sheets_) {
    if (strings::case_insensitive_eq(s.name(), name)) {
      return &s;
    }
  }
  return nullptr;
}

std::size_t Workbook::sheet_index_by_name(std::string_view name) const noexcept {
  for (std::size_t i = 0; i < sheets_.size(); ++i) {
    if (strings::case_insensitive_eq(sheets_[i].name(), name)) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
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
  return save_ex(io::WorkbookFormat::Ooxml);
}

Expected<std::vector<std::uint8_t>, Error> Workbook::save_ex(io::WorkbookFormat format) const {
  switch (format) {
    case io::WorkbookFormat::Ooxml:
      return io::write_ooxml(*this);
    case io::WorkbookFormat::Xlsb:
      return io::xlsb::write_xlsb(*this);
    case io::WorkbookFormat::Unknown:
      break;
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "Workbook::save_ex: unsupported format",
                    "context=workbook_save_ex");
}

namespace {

// Builds a CellNodeId for a workbook-relative coordinate. Sheet ids fit in
// uint16_t per the dep graph contract; Excel allows at most a few thousand
// sheets per workbook, well within range.
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
  for (eval::CellNodeId dep : mutator.dep_graph().dependents_of(cell)) {
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
    std::string normalized = io::strip_storage_prefixes(formula);
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

Expected<eval::RecalcStats, Error> Workbook::recalc(const eval::FunctionRegistry& registry) {
  return engine_->recalc(*this, registry);
}

Expected<void, Error> Workbook::recalc_parallel(const eval::FunctionRegistry& registry,
                                                const eval::SchedulerConfig& cfg, eval::SchedulerStats* stats) {
  return eval::recalc_parallel(*this, registry, cfg, stats);
}

void Workbook::set_iterative_options(eval::IterativeOptions opts) {
  engine_->set_iterative_options(opts);
}

const eval::IterativeOptions& Workbook::iterative_options() const noexcept {
  return engine_->iterative_options();
}

Expected<eval::RecalcStats, Error> Workbook::partial_recalc(const eval::FunctionRegistry& registry,
                                                            const eval::SheetCellRange& viewport) {
  return engine_->partial_recalc(*this, registry, viewport);
}

void Workbook::set_iterative_progress(eval::IterativeProgressCb cb, void* user_data) noexcept {
  engine_->set_iterative_progress(cb, user_data);
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

// Rewrites a single formula text through `transform`. Returns the
// rewritten body alongside a flag indicating whether any reference was
// actually changed; an unchanged body lets the caller skip the
// dep-graph re-register and avoid touching the cell.
struct FormulaRewriteResult {
  std::string text;
  bool changed = false;
  bool error = false;  // Parse failure or arena exhaustion.
};

FormulaRewriteResult rewrite_formula(std::string_view formula, const parser::RefTransform& transform) {
  FormulaRewriteResult out;
  out.text.assign(formula);
  std::string_view body = formula;
  bool had_equals = false;
  if (!body.empty() && body.front() == '=') {
    body = body.substr(1);
    had_equals = true;
  }
  if (body.empty()) {
    return out;
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    out.error = true;
    return out;
  }
  const parser::AstNode* shifted = parser::shift_refs(*root, arena, transform);
  if (shifted == nullptr) {
    out.error = true;
    return out;
  }
  if (shifted == root) {
    return out;  // Identity walk; no change required.
  }
  std::string rewritten;
  if (had_equals) {
    rewritten.push_back('=');
  }
  rewritten.append(parser::format_formula(*shifted));
  out.text = std::move(rewritten);
  out.changed = true;
  return out;
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
    const bool local_means_target = strings::case_insensitive_eq(sheet.name(), target_sheet);
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
      const std::vector<Cell>& cells = it->second;
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

void rewrite_defined_names(std::vector<io::DefinedName>& names, const parser::RefTransform& transform) {
  for (io::DefinedName& entry : names) {
    FormulaRewriteResult result = rewrite_formula(entry.formula, transform);
    if (result.changed) {
      entry.formula = std::move(result.text);
    }
  }
}

void rewrite_conditional_format_formulas(std::vector<cf::ConditionalFormat>& formats,
                                         const parser::RefTransform& transform) {
  for (cf::ConditionalFormat& format : formats) {
    for (cf::CFRule& rule : format.rules) {
      if (rule.formula1.has_value()) {
        FormulaRewriteResult result = rewrite_formula(*rule.formula1, transform);
        if (result.changed) {
          rule.formula1 = std::move(result.text);
        }
      }
      if (rule.formula2.has_value()) {
        FormulaRewriteResult result = rewrite_formula(*rule.formula2, transform);
        if (result.changed) {
          rule.formula2 = std::move(result.text);
        }
      }
    }
  }
}

void rewrite_tables_for_row_col_edit(std::vector<io::TableMetadata>& tables, std::size_t target_sheet_index,
                                     const parser::RefTransform& transform) {
  for (io::TableMetadata& table : tables) {
    if (table.sheet_index != target_sheet_index) {
      continue;
    }
    FormulaRewriteResult range = rewrite_formula(table.ref, transform);
    if (range.changed) {
      table.ref = std::move(range.text);
    }
    for (io::TableColumn& column : table.columns) {
      if (!column.calculated_column_formula.empty()) {
        FormulaRewriteResult formula = rewrite_formula(column.calculated_column_formula, transform);
        if (formula.changed) {
          column.calculated_column_formula = std::move(formula.text);
        }
      }
    }
  }
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
// lock.
Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                   const eval::RecalcEngine::LockedMutator& mutator,
                                                   std::vector<io::DefinedName>& defined_names, std::size_t sheet_index,
                                                   parser::RowColAxis axis, parser::RowColEdit edit,
                                                   std::uint32_t origin, std::uint32_t count, const char* op_name) {
  RETURN_IF_ERROR(apply_row_col_edit(wb, sheet_index, axis, edit, origin, count, op_name));
  const std::string target_sheet_name = sheets[sheet_index].name();
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, axis, edit, origin, count);
  const parser::RowColShiftTransform name_transform(target_sheet_name, axis, edit, origin, count);
  rewrite_defined_names(defined_names, name_transform);
  Sheet& target = sheets[sheet_index];
  const parser::RowColShiftTransform cf_transform(target_sheet_name, axis, edit, origin, count,
                                                  /*local_means_target=*/true);
  rewrite_conditional_format_formulas(target.mutable_conditional_formats(), cf_transform);
  rewrite_tables_for_row_col_edit(wb.mutable_tables(), sheet_index, cf_transform);
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
