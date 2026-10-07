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
#include <unordered_map>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "defined_name.h"
#include "drawing/drawing_edit.h"
#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/scheduler.h"
#include "external_link.h"
#include "io/auto_filter_xml.h"
#include "io/dynamic_array_formula.h"
#include "io/format_detect.h"
#include "io/future_functions.h"
#include "io/ooxml_writer.h"
#include "io/theme_part.h"
#include "io/workbook_kind_ooxml.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/writer.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/ast_shift.h"
#include "parser/external_book_spelling.h"
#include "parser/formula_prefix.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_ref.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook_formula_index.h"
#include "workbook_ref_rewrite.h"

namespace formulon {
namespace {
// Forward declaration; defined further below alongside the other
// sheet-name helpers. Declared here so the early `add_sheet_validated`
// definition can call it.
Expected<void, Error> validate_sheet_name(std::string_view name);

// `op`-prefixed range check for a mutator addressing one sheet.
Expected<void, Error> check_sheet_index(const char* op, std::size_t sheet_index, std::size_t sheet_count) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  return Expected<void, Error>::Ok();
}

// `op`-prefixed target check for a cell setter: the sheet index, then the grid coordinate.
Expected<void, Error> check_cell_target(const char* op, std::size_t sheet_index, std::size_t sheet_count,
                                        std::uint32_t row, std::uint32_t col) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index) + " sheet_count=" + std::to_string(sheet_count));
  }
  if (!Sheet::coord_in_grid(row, col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": coordinate out of grid",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}
}  // namespace

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()), kind_(WorkbookKind::kXlsx) {}
Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;
Workbook::~Workbook() = default;

Expected<void, Error> Workbook::add_passthrough_part(PassthroughPart part) {
  if (part.path.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "add_passthrough_part: empty part path");
  }
  for (const PassthroughPart& existing : passthrough_parts_) {
    if (existing.path == part.path) {
      return make_error(FormulonErrorCode::kInvalidArgument, "add_passthrough_part: part already present",
                        "path=" + part.path);
    }
  }
  passthrough_parts_.push_back(std::move(part));
  sort_passthrough_parts(passthrough_parts_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::replace_passthrough_part(std::string_view path, std::vector<std::uint8_t> bytes) {
  for (PassthroughPart& existing : passthrough_parts_) {
    if (existing.path == path) {
      existing.bytes = std::move(bytes);
      return Expected<void, Error>::Ok();
    }
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "replace_passthrough_part: no such part",
                    "path=" + std::string(path));
}

void Workbook::add_workbook_relationship(std::string type, std::string target) {
  for (const UnknownRelationship& rel : unknown_workbook_rels_) {
    if (!rel.target_external && rel.type == type && rel.target == target) {
      return;
    }
  }
  UnknownRelationship rel;
  rel.type = std::move(type);
  rel.target = std::move(target);
  unknown_workbook_rels_.push_back(std::move(rel));
}

LoadedTheme Workbook::load_theme() const {
  return io::load_theme(*this);
}

Expected<void, Error> Workbook::set_theme_colors(const ThemeColors& colors) {
  return io::set_theme_colors(*this, colors);
}

Expected<void, Error> Workbook::set_theme_fonts(const ThemeFonts& fonts) {
  return io::set_theme_fonts(*this, fonts);
}

Expected<void, Error> Workbook::reset_theme() {
  return io::reset_theme(*this);
}

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

// Appends a sheet and re-points formulas that already name it. The caller
// holds the compound-mutation mutex.
void append_sheet(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                  const Workbook& workbook, std::string name) {
  sheets.emplace_back(Sheet{std::move(name)});
  if (mutator.has_ever_registered_formula()) {
    reindex_formulas_referencing_sheet(sheets, mutator, workbook, sheets.back().name());
  }
}
}  // namespace

std::size_t Workbook::add_sheet(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  // No error channel here; at the ceiling the workbook is left as-is and
  // the caller gets the out-of-range sentinel (see the header contract).
  if (!check_sheet_headroom(sheets_.size()).has_value()) {
    return kMaxSheets;
  }
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name));
  return sheets_.size() - 1U;
}

Expected<std::size_t, Error> Workbook::add_sheet_checked(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name));
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
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name));
  return &sheets_.back();
}

namespace {

// Hidden sheet-scoped name Excel keeps over a sheet AutoFilter's range.
constexpr std::string_view kFilterDatabaseName = "_xlnm._FilterDatabase";

/// `Sheet1!$B$2:$F$12`, quoting the sheet name where Excel would.
std::string filter_database_formula(std::string_view sheet_name, const MergeRange& range) {
  std::string out;
  parser::append_sheet_name(sheet_name, false, out);
  out.push_back('!');
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    const std::string a1 = a1::encode_a1(row, col);
    const std::size_t digits = a1.find_first_of("0123456789");
    out.push_back('$');
    out.append(a1, 0, digits);
    out.push_back('$');
    out.append(a1, digits, std::string::npos);
  };
  append_cell(range.first_row, range.first_col);
  out.push_back(':');
  append_cell(range.last_row, range.last_col);
  return out;
}

/// Removes the `_FilterDatabase` name scoped to `sheet_index`. Returns true
/// when one existed.
bool erase_filter_database_name(std::vector<DefinedName>& names, std::size_t sheet_index) {
  const auto it = std::find_if(names.begin(), names.end(), [sheet_index](const DefinedName& entry) {
    return entry.local_sheet_id == static_cast<std::int32_t>(sheet_index) &&
           strings::case_insensitive_eq(entry.name, kFilterDatabaseName);
  });
  if (it == names.end()) {
    return false;
  }
  names.erase(it);
  return true;
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
  if (mutator.has_ever_registered_formula()) {
    reindex_formulas_referencing_sheet(sheets_, mutator, *this, sheets_[index].name());
  }
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

  std::vector<std::uint32_t> old_to_new(sheets_.size());
  const std::uint32_t new_last_index = static_cast<std::uint32_t>(sheets_.size() - 2U);
  for (std::uint32_t old_index = 0; old_index < old_to_new.size(); ++old_index) {
    if (old_index < index) {
      old_to_new[old_index] = old_index;
    } else if (old_index > index) {
      old_to_new[old_index] = old_index - 1U;
    } else {
      old_to_new[old_index] = std::min(index, new_last_index);
    }
  }
  remap_book_views_xml(book_views_xml_, old_to_new);

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

  std::vector<std::uint32_t> old_to_new(sheets_.size());
  for (std::uint32_t old_index = 0; old_index < old_to_new.size(); ++old_index) {
    if (old_index == from_index) {
      old_to_new[old_index] = to_index;
    } else if (from_index < to_index && old_index > from_index && old_index <= to_index) {
      old_to_new[old_index] = old_index - 1U;
    } else if (from_index > to_index && old_index >= to_index && old_index < from_index) {
      old_to_new[old_index] = old_index + 1U;
    } else {
      old_to_new[old_index] = old_index;
    }
  }
  remap_book_views_xml(book_views_xml_, old_to_new);

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

void Workbook::set_defined_names(std::vector<DefinedName> names) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  std::vector<std::string> changed;
  {
    // Declaration order matters for duplicate definitions. Metadata and
    // case-only spelling changes do not alter name resolution.
    using Definitions = std::vector<std::pair<std::int32_t, std::string_view>>;
    const auto group = [](const std::vector<DefinedName>& entries) {
      std::unordered_map<std::string, Definitions> grouped;
      for (const DefinedName& entry : entries) {
        grouped[strings::to_ascii_lower(entry.name)].emplace_back(entry.local_sheet_id, entry.formula);
      }
      return grouped;
    };
    const auto before = group(defined_names_);
    const auto after = group(names);
    for (const auto& [name, definitions] : before) {
      const auto found = after.find(name);
      if (found == after.end() || definitions != found->second) {
        changed.push_back(name);
      }
    }
    for (const auto& [name, definitions] : after) {
      if (before.find(name) == before.end()) {
        changed.push_back(name);
      }
    }
  }
  for (const DefinedName& entry : names) {
    bind_external_books_unlocked(std::string_view(entry.formula));
  }
  defined_names_ = std::move(names);
  if (changed.empty()) {
    return;
  }
  const auto affected = close_affected_names(defined_names_, std::move(changed));
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_if(sheets_, mutator, *this,
                      [&](const parser::AstNode& root) { return references_any_name(root, affected); });
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
  bind_external_books_unlocked(std::string_view(formula));
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

Expected<void, Error> Workbook::set_defined_name_hidden(std::string_view name, std::int32_t local_sheet_id,
                                                        bool hidden) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  for (DefinedName& entry : defined_names_) {
    if (entry.local_sheet_id == local_sheet_id && strings::case_insensitive_eq(entry.name, name)) {
      entry.hidden = hidden;
      return Expected<void, Error>::Ok();
    }
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "set_defined_name_hidden: no such defined name",
                    "name=" + std::string(name) + " local_sheet_id=" + std::to_string(local_sheet_id));
}

Expected<void, Error> Workbook::set_sheet_auto_filter(std::size_t sheet_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("set_sheet_auto_filter", sheet_index, sheets_.size()));
  RETURN_IF_ERROR(validate_auto_filter(filter));
  std::string formula = filter_database_formula(sheets_[sheet_index].name(), filter.range);
  sheets_[sheet_index].set_auto_filter(std::move(filter));
  const auto scope = static_cast<std::int32_t>(sheet_index);
  const auto it = std::find_if(defined_names_.begin(), defined_names_.end(), [scope](const DefinedName& entry) {
    return entry.local_sheet_id == scope && strings::case_insensitive_eq(entry.name, kFilterDatabaseName);
  });
  if (it != defined_names_.end()) {
    it->formula = std::move(formula);
    it->hidden = true;
  } else {
    DefinedName entry;
    entry.name = std::string(kFilterDatabaseName);
    entry.formula = std::move(formula);
    entry.local_sheet_id = scope;
    entry.hidden = true;
    defined_names_.push_back(std::move(entry));
  }
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  reindex_formulas_referencing_name(sheets_, mutator, *this, kFilterDatabaseName);
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_sheet_auto_filter(std::size_t sheet_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("remove_sheet_auto_filter", sheet_index, sheets_.size()));
  sheets_[sheet_index].clear_auto_filter_model();
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  if (erase_filter_database_name(defined_names_, sheet_index)) {
    reindex_formulas_referencing_name(sheets_, mutator, *this, kFilterDatabaseName);
  }
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_sheet_auto_filter_xml(std::size_t sheet_index, std::string_view xml) {
  if (xml.empty()) {
    return remove_sheet_auto_filter(sheet_index);
  }
  Expected<AutoFilter, Error> parsed = io::parse_auto_filter_xml(xml);
  if (parsed && validate_auto_filter(parsed.value())) {
    return set_sheet_auto_filter(sheet_index, std::move(parsed.value()));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("set_sheet_auto_filter_xml", sheet_index, sheets_.size()));
  sheets_[sheet_index].set_auto_filter(io::auto_filter_from_xml(xml));
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_table_auto_filter(std::size_t table_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (table_index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "set_table_auto_filter: table_index out of range",
                      "table_index=" + std::to_string(table_index));
  }
  RETURN_IF_ERROR(validate_auto_filter(filter));
  tables_[table_index].auto_filter_xml.set(std::move(filter));
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_table_auto_filter(std::size_t table_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (table_index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_table_auto_filter: table_index out of range",
                      "table_index=" + std::to_string(table_index));
  }
  tables_[table_index].auto_filter_xml.reset();
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator());
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

/// `Sheet::blocked_spill_anchors_intersecting` or
/// `Sheet::committed_spill_anchors_intersecting`.
using SpillAnchorQuery = std::vector<CellAddress> (Sheet::*)(std::uint32_t, std::uint32_t, std::uint32_t,
                                                             std::uint32_t) const;

MergeRange normalize_merge_range(MergeRange merge) noexcept {
  const std::uint32_t first_row = std::min(merge.first_row, merge.last_row);
  const std::uint32_t first_col = std::min(merge.first_col, merge.last_col);
  const std::uint32_t last_row = std::max(merge.first_row, merge.last_row);
  const std::uint32_t last_col = std::max(merge.first_col, merge.last_col);
  return MergeRange{first_row, first_col, last_row, last_col};
}

void mark_spill_anchors_intersecting_merge(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                           const std::vector<Sheet>& sheets, const MergeRange& merge,
                                           SpillAnchorQuery query) {
  if (merge.first_row > merge.last_row || merge.first_col > merge.last_col ||
      !Sheet::coord_in_grid(merge.first_row, merge.first_col) ||
      !Sheet::coord_in_grid(merge.last_row, merge.last_col)) {
    return;
  }
  const std::uint32_t rows = merge.last_row - merge.first_row + 1U;
  const std::uint32_t cols = merge.last_col - merge.first_col + 1U;
  for (const CellAddress anchor : (sheets[sheet_index].*query)(merge.first_row, merge.first_col, rows, cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  RETURN_IF_ERROR(check_cell_target("set_cell_value", sheet_index, sheets_.size(), row, col));

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
  RETURN_IF_ERROR(check_cell_target("set_cell_text", sheet_index, sheets_.size(), row, col));

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

void Workbook::apply_legacy_implicit_intersections() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  for (std::size_t sheet_index = 0; sheet_index < sheets_.size(); ++sheet_index) {
    const io::xlsb::NameShapes name_shapes = io::name_shapes(*this, sheet_index);
    Sheet& sheet = sheets_[sheet_index];
    for (const CellAddress address : sheet.formula_cells_in(0U, 0U, Sheet::kMaxRows - 1U, Sheet::kMaxCols - 1U)) {
      const Cell* cell = sheet.cell_at(address.row, address.col);
      // A dynamic-array formula already says where it intersects, and a CSE
      // block evaluates as an array.
      if (cell == nullptr || cell->dynamic_array || sheet.spill_region_at_anchor(address.row, address.col) != nullptr) {
        continue;
      }
      std::string text = cell->formula_text;
      const std::size_t body_at = !text.empty() && text.front() == '=' ? 1U : 0U;
      Arena arena;
      const parser::AstNode* root = parser::parse_strict(std::string_view(text).substr(body_at), arena);
      if (root == nullptr) {
        continue;
      }
      std::vector<std::uint32_t> at;
      for (const parser::AstNode* node : io::xlsb::legacy_intersections(*root, name_shapes)) {
        if (node->kind() != parser::NodeKind::ImplicitIntersection) {
          at.push_back(node->range().start);
        }
      }
      if (at.empty()) {
        continue;
      }
      // Right to left, so each insertion leaves the earlier offsets valid.
      sort_ascending(at);
      for (auto it = at.rbegin(); it != at.rend(); ++it) {
        text.insert(body_at + *it, 1U, '@');
      }
      sheet.set_cell_formula_text(address.row, address.col, std::move(text));
    }
  }
}

std::string Workbook::normalize_formula_text(std::string formula) {
  return parser::spell_storage_operators(parser::strip_storage_prefixes(formula, &io::has_storage_prefix));
}

namespace {

// The file name a link target ends in; a `file:` URL is percent-decoded.
std::string LinkFileName(std::string_view target) {
  const std::size_t sep = target.find_last_of("/\\");
  const std::string_view name = sep == std::string_view::npos ? target : target.substr(sep + 1);
  if (target.compare(0, 5, "file:") != 0) {
    return std::string(name);
  }
  return parser::percent_decode(name);
}

// Directory the link displays: from the alternate absolute URL when Excel
// recorded one, otherwise from the target itself.
std::string LinkDisplayPath(const ExternalLinkRecord& rec) {
  return parser::display_path_for_link_target(rec.absolute_target.empty() ? rec.target : rec.absolute_target);
}

std::string LinkBookName(const ExternalLinkRecord& rec) {
  return LinkFileName(rec.target.empty() ? rec.absolute_target : rec.target);
}

bool IsPathSeparator(char c) noexcept {
  return c == '/' || c == '\\';
}

// Path equality under ASCII case folding, with `/` and `\` interchangeable.
bool SamePath(std::string_view a, std::string_view b) noexcept {
  if (a.size() != b.size()) {
    return false;
  }
  for (std::size_t i = 0; i < a.size(); ++i) {
    if (IsPathSeparator(a[i]) && IsPathSeparator(b[i])) {
      continue;
    }
    if (strings::ascii_to_lower(a[i]) != strings::ascii_to_lower(b[i])) {
      return false;
    }
  }
  return true;
}

// The target a new link records for `path` + `book`: the file name alone,
// a POSIX path as written, and a Windows drive or UNC path as a `file:` URL.
std::string TargetForNewLink(std::string_view path, std::string_view book) {
  std::string out;
  const bool drive =
      path.size() >= 2 && ((path[0] >= 'A' && path[0] <= 'Z') || (path[0] >= 'a' && path[0] <= 'z')) && path[1] == ':';
  const bool unc = path.size() >= 2 && path[0] == '\\' && path[1] == '\\';
  if (drive) {
    out = "file:///";
    out.append(path);
  } else if (unc) {
    out = "file:";
    for (const char c : path) {
      out.push_back(c == '\\' ? '/' : c);
    }
  } else {
    out.append(path);
  }
  out.append(book);
  return out;
}

// Whether `formula` can contain a cross-workbook reference at all: every
// spelling has a `[`, or a `!` after a workbook extension or a quoted path.
bool MayNameExternalBook(std::string_view formula) noexcept {
  if (formula.find('[') != std::string_view::npos) {
    return true;
  }
  if (formula.find('!') == std::string_view::npos) {
    return false;
  }
  return strings::case_insensitive_contains(formula, ".xl") ||
         (formula.find('\'') != std::string_view::npos && formula.find_first_of("/\\") != std::string_view::npos);
}

// Position in `links` of the link `path` + `book` names (see
// `Workbook::find_external_link`), or `links.size()` when none does.
std::size_t FindLinkPosition(const std::vector<ExternalLinkRecord>& links, std::string_view path,
                             std::string_view book) noexcept {
  std::size_t found = links.size();
  const auto take_lowest = [&links, &found](std::size_t i) {
    if (found == links.size() || links[i].index < links[found].index) {
      found = i;
    }
  };
  if (!path.empty()) {
    std::string key(path);
    key.append(book);
    for (std::size_t i = 0; i < links.size(); ++i) {
      const std::string dir = LinkDisplayPath(links[i]);
      if (!dir.empty() && SamePath(dir + LinkBookName(links[i]), key)) {
        return i;
      }
    }
    for (std::size_t i = 0; i < links.size(); ++i) {
      if (LinkDisplayPath(links[i]).empty() && strings::case_insensitive_eq(LinkBookName(links[i]), book)) {
        take_lowest(i);
      }
    }
    return found;
  }
  for (std::size_t i = 0; i < links.size(); ++i) {
    if (strings::case_insensitive_eq(LinkFileName(links[i].target), book) ||
        strings::case_insensitive_eq(LinkFileName(links[i].absolute_target), book)) {
      take_lowest(i);
    }
  }
  return found;
}

std::uint32_t IndexForBook(const void* ctx, std::string_view path, std::string_view book) {
  const ExternalLinkRecord* rec = static_cast<const Workbook*>(ctx)->find_external_link(path, book);
  return rec == nullptr ? 0U : rec->index;
}

std::uint32_t IndexForSheetQualifier(const void* ctx, std::string_view qualifier) {
  const ExternalLinkRecord* rec = static_cast<const Workbook*>(ctx)->link_for_sheet_qualifier(qualifier);
  return rec == nullptr ? 0U : rec->index;
}

bool ResolveLinkIndex(const void* ctx, std::uint32_t index, parser::ExternalBookDisplay* out) {
  for (const ExternalLinkRecord& rec : *static_cast<const std::vector<ExternalLinkRecord>*>(ctx)) {
    if (rec.index != index) {
      continue;
    }
    out->book = LinkBookName(rec);
    if (out->book.empty()) {
      return false;
    }
    out->path = LinkDisplayPath(rec);
    return true;
  }
  return false;
}

}  // namespace

std::vector<const ExternalLinkRecord*> Workbook::external_links_by_index(bool skip_ole_dde) const {
  std::vector<const ExternalLinkRecord*> out;
  out.reserve(external_links_.size());
  for (const ExternalLinkRecord& rec : external_links_) {
    if (skip_ole_dde &&
        (rec.kind == ExternalLinkRecord::Kind::kOleLink || rec.kind == ExternalLinkRecord::Kind::kDdeLink)) {
      continue;
    }
    // Insertion sort: stable, and the list is almost always in order already.
    std::size_t at = out.size();
    out.push_back(&rec);
    for (; at > 0 && out[at - 1]->index > rec.index; --at) {
      out[at] = out[at - 1];
    }
    out[at] = &rec;
  }
  return out;
}

const ExternalLinkRecord* Workbook::find_external_link(std::string_view path, std::string_view book) const noexcept {
  const std::size_t at = FindLinkPosition(external_links_, path, book);
  return at == external_links_.size() ? nullptr : &external_links_[at];
}

const ExternalLinkRecord* Workbook::link_for_sheet_qualifier(std::string_view qualifier) const noexcept {
  if (sheet_by_name(qualifier) != nullptr) {
    return nullptr;
  }
  const ExternalLinkRecord* rec = find_external_link({}, qualifier);
  if (rec == nullptr || rec->kind == ExternalLinkRecord::Kind::kOleLink ||
      rec->kind == ExternalLinkRecord::Kind::kDdeLink) {
    return nullptr;
  }
  return rec;
}

parser::ExternalBookIndexer Workbook::external_book_indexer() const noexcept {
  return parser::ExternalBookIndexer{&IndexForBook, this, &IndexForSheetQualifier};
}

void Workbook::bind_external_books(const parser::AstNode& root) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  bind_external_books_unlocked(root);
}

void Workbook::bind_external_books(std::string_view formula) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  bind_external_books_unlocked(formula);
}

void Workbook::bind_external_books_unlocked(std::string_view formula) {
  if (!MayNameExternalBook(formula)) {
    return;
  }
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  if (const parser::AstNode* root = parse_indexable_formula(formula, arena); root != nullptr) {
    bind_external_books_unlocked(*root);
  }
}

void Workbook::bind_external_books_unlocked(const parser::AstNode& root) {
  std::vector<const parser::AstNode*> pending{&root};
  while (!pending.empty()) {
    const parser::AstNode* node = pending.back();
    pending.pop_back();
    if (node->kind() != parser::NodeKind::ExternalRef) {
      // Reversed, so references bind in source order.
      const std::vector<const parser::AstNode*> children = parser::child_nodes(*node);
      pending.insert(pending.end(), children.rbegin(), children.rend());
      continue;
    }
    if (parser::is_self_book_name_ref(*node)) {
      continue;
    }
    const std::string_view path = node->as_external_ref_path();
    const std::string_view book = node->as_external_ref_book();
    const std::size_t at = FindLinkPosition(external_links_, path, book);
    ExternalLinkRecord* rec = at == external_links_.size() ? nullptr : &external_links_[at];
    if (rec == nullptr) {
      std::uint32_t next = 1;
      for (const ExternalLinkRecord& existing : external_links_) {
        next = existing.index >= next ? existing.index + 1U : next;
      }
      ExternalLinkRecord created;
      created.index = next;
      created.kind = ExternalLinkRecord::Kind::kExternalBook;
      created.target_external = true;
      created.target = TargetForNewLink(path, book);
      external_links_.push_back(std::move(created));
      rec = &external_links_.back();
    } else if (!path.empty() && LinkDisplayPath(*rec).empty()) {
      // Excel keeps one link for the book and retargets it at the absolute path.
      rec->target = TargetForNewLink(path, book);
      rec->absolute_target = rec->target;
      rec->body_stale = rec->body_stale || !rec->part_path.empty();
    }
    for (const std::string_view sheet : {node->as_external_ref_sheet(), node->as_external_ref_sheet_end()}) {
      if (sheet.empty() || rec->book.sheet_index(sheet) != ExternalBook::kNoSheet) {
        continue;
      }
      rec->book.sheet_names.emplace_back(sheet);
      // A loaded body no longer lists every sheet the model does.
      rec->body_stale = rec->body_stale || !rec->part_path.empty();
    }
  }
}

std::string Workbook::ingest_stored_formula(std::string_view stored) {
  std::string text = normalize_formula_text(std::string(stored));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (text.find('[') != std::string::npos) {
    // A link whose rels named no target is the book its index spells.
    for (ExternalLinkRecord& rec : external_links_) {
      const bool book_kind =
          rec.kind == ExternalLinkRecord::Kind::kExternalBook || rec.kind == ExternalLinkRecord::Kind::kUnknown;
      if (book_kind && rec.target.empty() && rec.absolute_target.empty()) {
        rec.target = std::to_string(rec.index);
      }
    }
    text = parser::spell_external_books(text, parser::ExternalBookResolver{&ResolveLinkIndex, &external_links_});
  }
  bind_external_books_unlocked(std::string_view(text));
  return text;
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula) {
  RETURN_IF_ERROR(check_cell_target("set_cell_formula", sheet_index, sheets_.size(), row, col));

  // Normalize Excel's `_xlfn.` / `_xlfn._xlws.` / `_xlws.` / `_xlpm.`
  // storage prefixes to the canonical formula-bar form at the single
  // ingestion point every reader (OOXML DOM / SAX, XLSB) and binding
  // funnels through. This keeps the stored `formula_text` (and the
  // dependency-extraction parse below) matching what Excel's formula bar
  // shows, and lets LET / LAMBDA resolve their `_xlpm.`-prefixed
  // parameter names. The transform is a no-op on an already-canonical
  // formula, so hand-authored / test formulas are unaffected. The writer
  // re-applies the prefixes on save for Excel readability. The stored
  // `SINGLE(x)` / `ANCHORARRAY(x)` calls read back as the `@x` / `x#` the
  // formula bar shows.
  formula = normalize_formula_text(std::move(formula));

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // Parse the formula in a throwaway arena to extract its dependency list.
  // Strip a leading '=' so the parser sees a bare expression (the tokenizer
  // accepts both shapes, but the dep extractor walks the AST regardless).
  std::string_view src = formula;
  if (!src.empty() && src.front() == '=') {
    src.remove_prefix(1);
  }

  Arena tmp_arena;
  parser::AstNode* root = parse_indexable_formula(src, tmp_arena);
  const bool may_name_external_book = MayNameExternalBook(src);

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
    // Marked as Excel 365 marks a typed formula; a reader replaces this with
    // the file's own mark.
    sheets_[sheet_index].set_cell_dynamic_array(
        row, col, root != nullptr && io::entered_as_dynamic_array(*this, sheet_index, *root));

    if (root != nullptr) {
      if (may_name_external_book) {
        bind_external_books_unlocked(*root);
      }
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
  RETURN_IF_ERROR(check_sheet_index("add_merge", sheet_index, sheets_.size()));
  merge = normalize_merge_range(merge);
  if (!Sheet::rect_in_grid(merge.first_row, merge.first_col, merge.last_row, merge.last_col)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "add_merge: range out of grid",
                      "first_row=" + std::to_string(merge.first_row) + " first_col=" + std::to_string(merge.first_col) +
                          " last_row=" + std::to_string(merge.last_row) +
                          " last_col=" + std::to_string(merge.last_col));
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  sheets_[sheet_index].add_merge(merge);
  // A merge added over a committed spill becomes a blocker for the next
  // evaluation. Keep the anchor dirty so recalc clears the old rectangle and
  // records the resulting #SPILL! state under the normal commit contract.
  mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, merge,
                                        &Sheet::committed_spill_anchors_intersecting);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_merges_intersecting(std::size_t sheet_index, MergeRange merge) {
  RETURN_IF_ERROR(check_sheet_index("remove_merges_intersecting", sheet_index, sheets_.size()));
  merge = normalize_merge_range(merge);
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::vector<MergeRange> removed = sheets_[sheet_index].remove_merges_intersecting(merge);
  for (const MergeRange& erased : removed) {
    mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, erased,
                                          &Sheet::blocked_spill_anchors_intersecting);
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_merge_at(std::size_t sheet_index, std::size_t index) {
  RETURN_IF_ERROR(check_sheet_index("remove_merge_at", sheet_index, sheets_.size()));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  MergeRange removed;
  if (!sheets_[sheet_index].remove_merge_at(index, &removed)) {
    return make_error(FormulonErrorCode::kInvalidArgument, "remove_merge_at: index out of range",
                      "index=" + std::to_string(index));
  }
  mark_spill_anchors_intersecting_merge(mutator, sheet_index, sheets_, removed,
                                        &Sheet::blocked_spill_anchors_intersecting);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_merges(std::size_t sheet_index) {
  RETURN_IF_ERROR(check_sheet_index("clear_merges", sheet_index, sheets_.size()));
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

namespace {

void mark_all_formulas_dirty_locked(const std::vector<Sheet>& sheets,
                                    const eval::RecalcEngine::LockedMutator& mutator) {
  for (std::size_t sheet_idx = 0; sheet_idx < sheets.size(); ++sheet_idx) {
    for (const auto& [row, cells] : sheets[sheet_idx].rows()) {
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

void invalidate_context_dependent_state_locked(std::vector<Sheet>& sheets,
                                               const eval::RecalcEngine::LockedMutator& mutator) {
  for (Sheet& sheet : sheets) {
    for (const std::unique_ptr<pivot::PivotTable>& table : sheet.mutable_pivot_tables()) {
      if (table == nullptr) {
        continue;
      }
      table->clear_last_result();
      table->clear_span_authored();
    }
  }
  mark_all_formulas_dirty_locked(sheets, mutator);
}

bool same_civil_time(const date_time::CivilTime& lhs, const date_time::CivilTime& rhs) noexcept {
  return lhs.date.y == rhs.date.y && lhs.date.m == rhs.date.m && lhs.date.d == rhs.date.d && lhs.time.h == rhs.time.h &&
         lhs.time.m == rhs.time.m && lhs.time.s == rhs.time.s;
}

}  // namespace

void Workbook::set_excel_profile(ExcelProfile profile) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (same_profile(excel_profile_, profile)) {
    return;
  }
  excel_profile_ = profile;
  invalidate_context_dependent_state_locked(sheets_, engine_->locked_mutator());
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
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const IterativeOptions previous = engine_->iterative_options();
  const bool changed = previous.enabled != opts.enabled || previous.max_iterations != opts.max_iterations ||
                       previous.max_change != opts.max_change;
  engine_->set_iterative_options(opts);
  if (changed) {
    mark_all_formulas_dirty_locked(sheets_, engine_->locked_mutator());
  }
}

void Workbook::set_date1904(bool value) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (date1904_ == value) {
    return;
  }
  date1904_ = value;
  invalidate_context_dependent_state_locked(sheets_, engine_->locked_mutator());
}

void Workbook::set_pinned_now(date_time::CivilTime value) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (pinned_now_.has_value() && same_civil_time(*pinned_now_, value)) {
    return;
  }
  pinned_now_ = value;
  invalidate_context_dependent_state_locked(sheets_, engine_->locked_mutator());
}

void Workbook::clear_pinned_now() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (!pinned_now_.has_value()) {
    return;
  }
  pinned_now_.reset();
  invalidate_context_dependent_state_locked(sheets_, engine_->locked_mutator());
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
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator());
}

void Workbook::mark_all_formulas_dirty() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  mark_all_formulas_dirty_locked(sheets_, engine_->locked_mutator());
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

        // Resolve the cell's post-shift coordinates before parsing the formula.
        // An opaque formula still moves with its cell; skipping it on parse
        // failure would leave the dependency graph keyed at the old address.
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

        std::string_view body = cell.formula_text;
        bool had_equals = false;
        if (!body.empty() && body.front() == '=') {
          body = body.substr(1);
          had_equals = true;
        }

        bool ast_changed = false;
        std::string effective_formula = cell.formula_text;
        parser_arena.reset();
        if (!body.empty()) {
          parser::AstNode* root = parse_indexable_formula(body, parser_arena);
          if (root != nullptr) {
            const parser::AstNode* shifted = parser::shift_refs(*root, parser_arena, transform);
            ast_changed = (shifted != nullptr && shifted != root);
            if (ast_changed) {
              effective_formula.clear();
              if (had_equals) {
                effective_formula.push_back('=');
              }
              effective_formula.append(parser::format_formula(*shifted));
            }
          }
        }

        if (!ast_changed && !cell_moves && !dropped) {
          continue;  // Cell unaffected by the edit.
        }

        // Keep the rewritten text alongside the key transition. It is
        // applied only after every disappearing / moving old key has been
        // unregistered. Parse failures and an empty body remain opaque and
        // therefore retain their original text exactly.

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
    parser::AstNode* root = parse_indexable_formula(body, parser_arena);
    if (root == nullptr) {
      // The cell has moved, but its formula is opaque. Keep the text as-is
      // and wake the new coordinate without inventing dependency edges.
      mutator.mark_dirty(update.new_node);
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
    parser::AstNode* root = parse_stored_formula(formula_cell->formula_text, parser_arena);
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
  for (const BlockedSpillFootprint& original : footprints) {
    const CellShift shift =
        shift_cell_coords_for_row_col_edit(axis, edit, index, count, original.anchor_row, original.anchor_col);
    if (!shift.kept) {
      continue;
    }
    BlockedSpillFootprint next = original;
    next.anchor_row = shift.new_row;
    next.anchor_col = shift.new_col;
    mapped.push_back(next);
  }
  return mapped;
}

Expected<void, Error> apply_row_col_edit(Workbook& wb, std::size_t sheet_index, parser::RowColAxis axis,
                                         parser::RowColEdit edit, std::uint32_t origin, std::uint32_t count,
                                         const char* op_name) {
  RETURN_IF_ERROR(check_sheet_index(op_name, sheet_index, wb.sheet_count()));
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
[[gnu::noinline]] Expected<void, Error> apply_row_col_edit_operation(Workbook& wb, std::vector<Sheet>& sheets,
                                                                     const eval::RecalcEngine::LockedMutator& mutator,
                                                                     std::vector<DefinedName>& defined_names,
                                                                     std::size_t sheet_index, parser::RowColAxis axis,
                                                                     parser::RowColEdit edit, std::uint32_t origin,
                                                                     std::uint32_t count, const char* op_name) {
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
  const parser::RowColShiftTransform local_transform(target_sheet_name, axis, edit, origin, count,
                                                     /*local_means_target=*/true);
  std::vector<const parser::RefTransform*> per_sheet_transforms(sheets.size(), &name_transform);
  per_sheet_transforms[sheet_index] = &local_transform;
  bool names_changed = rewrite_defined_names(defined_names, per_sheet_transforms, name_transform);
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
  // The edit clears committed spills; snapshot their anchors so those formulas are re-dirtied.
  const std::vector<SpillFootprint> committed_before = target.committed_spill_footprints();
  // AutoFilters on the edited sheet, serialized before the move, to detect a
  // change that alters which hidden rows count as filtered.
  const bool had_sheet_filter = target.has_auto_filter();
  std::vector<std::string> filters_before;
  filters_before.push_back(io::auto_filter_xml(target.auto_filter()));
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index == sheet_index) {
      filters_before.push_back(io::auto_filter_xml(table.auto_filter_xml.get()));
    }
  }
  rewrite_formulas_for_row_col_edit(sheets, mutator, wb, target_sheet_name, axis, edit, origin, count);
  // Metadata formulas follow the same per-sheet policy as cell formulas: a
  // qualified reference to the edited sheet shifts no matter which sheet owns
  // the rule, while an unqualified one is in scope only on the edited sheet.
  // The text is rewritten against pre-edit coordinates, like the cells above
  // and unlike the sqref/anchor rectangles, which `Sheet::insert_rows` and
  // friends move afterwards.
  std::vector<std::uint32_t> ignored_cache_ids;
  rewrite_sheet_metadata_formulas(sheets, per_sheet_transforms, wb.mutable_tables(), wb.mutable_pivot_caches(), {}, {},
                                  {}, ignored_cache_ids, name_transform);
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
  shift_retained_extension_ranges(target, origin, count, edit == parser::RowColEdit::kDelete,
                                  axis == parser::RowColAxis::kRow);
  shift_drawing_anchors(wb, sheet_index, origin, count, edit == parser::RowColEdit::kDelete,
                        axis == parser::RowColAxis::kRow);
  std::vector<std::string> filters_after;
  filters_after.push_back(io::auto_filter_xml(target.auto_filter()));
  for (TableMetadata& table : wb.mutable_tables()) {
    if (table.sheet_index != sheet_index) {
      continue;
    }
    if (AutoFilter* filter = table.auto_filter_xml.get();
        filter != nullptr && !shift_auto_filter(*filter, origin, count, edit == parser::RowColEdit::kDelete,
                                                axis == parser::RowColAxis::kRow, /*header_delete_removes=*/false)) {
      table.auto_filter_xml.reset();
    }
    filters_after.push_back(io::auto_filter_xml(table.auto_filter_xml.get()));
  }
  // A sheet AutoFilter the edit removed takes its `_FilterDatabase` name
  // with it; the full re-index below covers formulas naming it.
  if (had_sheet_filter && !target.has_auto_filter() && erase_filter_database_name(defined_names, sheet_index)) {
    names_changed = true;
  }
  if (filters_after != filters_before) {
    mark_row_visibility_dependents_dirty_locked(sheets, mutator);
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
  // Map each snapshotted anchor through the edit; deleted anchors and non-formula cells are skipped.
  for (const SpillFootprint& footprint : committed_before) {
    const CellShift shift =
        shift_cell_coords_for_row_col_edit(axis, edit, origin, count, footprint.anchor_row, footprint.anchor_col);
    if (!shift.kept) {
      continue;
    }
    const Cell* cell = target.cell_at(shift.new_row, shift.new_col);
    if (cell == nullptr || cell->formula_text.empty()) {
      continue;
    }
    mutator.mark_dirty(make_node(sheet_index, shift.new_row, shift.new_col));
  }
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
  RETURN_IF_ERROR(check_cell_target("set_cell_xf_index", sheet_index, sheets_.size(), row, col));
  // A style write can grow the sheet's sparse row store, so serialize it
  // with recalc just like all other workbook-level cell mutations. The
  // sheet-level setter deliberately bypasses literal-write spill invalidation
  // so formatting a live spill phantom preserves the dynamic array.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].set_cell_xf_index(row, col, xf_index);
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
