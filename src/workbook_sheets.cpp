//
// Workbook sheet list mutation: add, rename, remove, move and the active-sheet
// bookkeeping, with the sheet-name validators and the sheet-index remapping
// they share.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "chart/chart_ref.h"
#include "defined_name.h"
#include "drawing/drawing_xml.h"
#include "drawing/slicer_edit.h"
#include "eval/recalc_engine.h"
#include "io/ooxml/part_graph.h"
#include "new_workbook_defaults.h"
#include "parser/ref_transforms.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/utf8_length.h"
#include "workbook.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"
#include "workbook_sheet_mutation.h"
#include "workbook_table_loss.h"

namespace formulon {
namespace {
// Forward declarations; defined further below alongside the other
// sheet-name helpers. Declared here so the early `add_sheet_validated` and
// `add_sheet_entered` definitions can call them.
Expected<void, Error> validate_sheet_name(std::string_view name);
Expected<void, Error> validate_entered_sheet_name(std::string_view name, ExcelHost host);

}  // namespace

namespace {
// Rejects an append that would take the workbook past `kMaxSheets`. Every
// append routes through here, so `sheets_.size() <= kMaxSheets` holds for
// the lifetime of the workbook and each narrowing of a sheet index to the
// dep graph's 16-bit `sheet_id` is lossless by construction.
Expected<void, Error> check_sheet_headroom(std::size_t current_count) {
  if (current_count >= Workbook::kMaxSheets) {
    return make_error(FormulonErrorCode::kSheetCountLimitExceeded, "workbook already holds the maximum sheets",
                      "limit=" + std::to_string(Workbook::kMaxSheets));
  }
  return {};
}

// Appends a sheet and re-points formulas that already name it. The caller
// holds the compound-mutation mutex.
void append_sheet(std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator,
                  const Workbook& workbook, std::string name, const ExcelProfile* new_sheet_profile) {
  sheets.emplace_back(Sheet{std::move(name)});
  if (new_sheet_profile != nullptr) {
    apply_new_sheet_defaults(*new_sheet_profile, sheets.back());
  }
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
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name), new_sheet_profile());
  return sheets_.size() - 1U;
}

Expected<Sheet*, Error> Workbook::add_sheet_validated(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(validate_sheet_name(name));
  // The file readers' duplicate rule: the case fold only, so wide-fold twins Excel writes still load.
  if (std::any_of(sheets_.begin(), sheets_.end(), [&name](const Sheet& sheet) {
        return sheet_names::same(sheet.name(), name, sheet_names::Fold::kNarrow);
      })) {
    return make_error(FormulonErrorCode::kAlreadyExists, "name collides with an existing sheet",
                      "name=\"" + name + "\"");
  }
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name), new_sheet_profile());
  return &sheets_.back();
}

Expected<void, Error> Workbook::check_entered_new_sheet_name(std::string_view name) const {
  RETURN_IF_ERROR(validate_sheet_name(name));
  RETURN_IF_ERROR(validate_entered_sheet_name(name, excel_profile_.host));
  if (detail::sheet_name_collides(sheets_, name, kNoSheet)) {
    return make_error(FormulonErrorCode::kAlreadyExists, "name collides with an existing sheet",
                      "name=\"" + std::string(name) + "\"");
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::check_sheet_capacity() const {
  return check_sheet_headroom(sheets_.size());
}

std::string Workbook::next_default_sheet_name() const {
  std::string_view word = new_workbook_defaults(excel_profile_.locale).sheet_name;
  while (!word.empty() && strings::is_ascii_digit(word.back())) {
    word.remove_suffix(1);
  }
  for (std::size_t n = 1;; ++n) {
    std::string name = std::string(word) + std::to_string(n);
    if (!detail::sheet_name_collides(sheets_, name, kNoSheet)) {
      return name;
    }
  }
}

Expected<Sheet*, Error> Workbook::add_sheet_entered(std::string name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_entered_new_sheet_name(name));
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name), new_sheet_profile());
  return &sheets_.back();
}

namespace {

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

// Shared structural validator for a sheet name across every mutation
// surface (add / rename / any future import-side check). Verifies strict
// UTF-8, a non-empty name within the code-unit length limit, no forbidden
// characters and no edge apostrophe. Duplicate/case-folding collision is caller-scoped (it needs
// the target index) and handled at each call site.
Expected<void, Error> validate_sheet_name(std::string_view name) {
  if (name.empty()) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name must not be empty", "name=\"\"");
  }
  if (!sheet_names::valid_utf8(name)) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name is not valid UTF-8", "name=invalid-utf8");
  }
  if (utf16_units_in(name) > sheet_names::kMaxNameUnits) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name exceeds 31 characters",
                      "name=\"" + std::string(name) + "\"");
  }
  if (!is_valid_sheet_name_chars(name)) {
    return make_error(FormulonErrorCode::kInvalidSheetName,
                      "sheet name contains a forbidden character (: \\ / ? * [ ])",
                      "name=\"" + std::string(name) + "\"");
  }
  // Excel refuses an edge apostrophe on entry and a file carrying one fails to open, like a forbidden character.
  if (name.front() == '\'' || name.back() == '\'') {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name must not begin or end with an apostrophe",
                      "name=\"" + std::string(name) + "\"");
  }
  return Expected<void, Error>::Ok();
}

// The name Windows Excel's naming UI reserves beyond the structural rules:
// `History`, in any case. Mac Excel accepts it on entry and on load.
Expected<void, Error> validate_entered_sheet_name(std::string_view name, ExcelHost host) {
  if (host == ExcelHost::kWin365 && strings::case_insensitive_eq(name, "History")) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "sheet name \"History\" is reserved",
                      "name=\"" + std::string(name) + "\"");
  }
  return Expected<void, Error>::Ok();
}

/// Position of the sheet at `i` after the sheet at `from` is lifted out and
/// reinserted at `to`.
template <typename T>
T moved_index(T i, T from, T to) {
  if (i == from) {
    return to;
  }
  if (from < to && i > from && i <= to) {
    return i - 1;
  }
  if (from > to && i >= to && i < from) {
    return i + 1;
  }
  return i;
}

/// Position of the sheet at `i` after the sheet at `removed` is erased.
template <typename T>
T removed_index(T i, T removed) {
  return i > removed ? i - 1 : i;
}

void remove_and_reindex_tables(std::vector<TableMetadata>& tables, std::size_t removed_sheet_index) {
  std::vector<TableMetadata> retained;
  retained.reserve(tables.size());
  for (TableMetadata& table : tables) {
    if (table.sheet_index == removed_sheet_index) {
      continue;
    }
    table.sheet_index = removed_index(table.sheet_index, removed_sheet_index);
    retained.push_back(std::move(table));
  }
  tables = std::move(retained);
}

void move_table_sheet_indices(std::vector<TableMetadata>& tables, std::size_t from_index, std::size_t to_index) {
  for (TableMetadata& table : tables) {
    table.sheet_index = moved_index(table.sheet_index, from_index, to_index);
  }
}

}  // namespace

Expected<void, Error> Workbook::rename_sheet(std::uint32_t index, std::string new_name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Validate the new name (non-empty, ≤ 31 code units, no forbidden
  // characters, none of the names Excel's naming UI refuses) via the shared
  // validators, so a rename and `add_sheet_entered` agree.
  RETURN_IF_ERROR(validate_sheet_name(new_name));
  RETURN_IF_ERROR(validate_entered_sheet_name(new_name, excel_profile_.host));
  // Collision check against every other sheet; the sheet's own name is
  // excluded so a rename can change only case, width or kana type.
  if (detail::sheet_name_collides(sheets_, new_name, index)) {
    return make_error(FormulonErrorCode::kAlreadyExists, "name collides with another sheet",
                      "name=\"" + new_name + "\"");
  }

  const std::string old_name = sheets_[index].name();

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  std::vector<std::uint32_t> ignored_cache_ids;
  bool defined_names_changed = false;
  const sheet_names::NameList sheet_list = detail::sheet_name_list(sheets_);
  const parser::SheetRenameTransform transform(old_name, new_name, &sheet_list);
  // Keep the sheet name in its pre-rename state until every formula and
  // metadata holder has been transformed. In particular, do not re-register
  // a changed formula against `new_name` before the workbook can resolve it.
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, old_name, new_name,
                              {}, ignored_cache_ids, defined_names_changed);
  RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));
  if (defined_names_changed) {
    // Defined-name expansion can change a formula's value without changing
    // the formula cell's own text. Sheet ids survive a rename, so retaining
    // the existing graph edges is safe; dirty every formula so those cached
    // values are nevertheless refreshed on the next pass.
    mark_all_formulas_dirty_locked(sheets_, mutator);
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

std::uint32_t Workbook::active_sheet() const {
  if (book_views_xml_.empty()) {
    for (std::size_t i = 0; i < sheets_.size(); ++i) {
      if (sheets_[i].view().tab_selected) {
        return static_cast<std::uint32_t>(i);
      }
    }
    return 0U;
  }
  const std::uint32_t index = book_views_active_tab(book_views_xml_);
  return index < sheets_.size() ? index : 0U;
}

std::uint32_t Workbook::nearest_visible_sheet(std::uint32_t index) const {
  for (std::size_t i = static_cast<std::size_t>(index) + 1U; i < sheets_.size(); ++i) {
    if (sheets_[i].view().visibility() == SheetVisibility::kVisible) {
      return static_cast<std::uint32_t>(i);
    }
  }
  for (std::size_t i = index; i > 0U; --i) {
    if (sheets_[i - 1U].view().visibility() == SheetVisibility::kVisible) {
      return static_cast<std::uint32_t>(i - 1U);
    }
  }
  return index;
}

void Workbook::set_active_sheet(std::uint32_t index) {
  set_book_views_active_tab(book_views_xml_, index);
  for (std::size_t i = 0; i < sheets_.size(); ++i) {
    sheets_[i].mutable_view().tab_selected = i == index;
  }
}

Expected<void, Error> Workbook::remove_sheet(std::uint32_t index) {
  // Hold the engine mutex from the first workbook-state read through the
  // final graph rebuild. A parallel recalc therefore observes one complete
  // pre-removal or post-removal workbook, never the metadata halfway state.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (sheets_.size() <= 1U) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "cannot remove the only sheet",
                      "sheet_count=" + std::to_string(sheets_.size()));
  }
  return remove_sheet_locked(index);
}

Expected<void, Error> Workbook::remove_sheet_locked(std::uint32_t index) {
  // Read before anything changes: a surviving chart series counts toward the promoted names.
  auto chart_formulas = chart_ref::collect_formula_texts(*this, index);
  if (!chart_formulas) {
    return chart_formulas.error();
  }
  const std::string removed_name = sheets_[index].name();
  erase_slicer_cache_names(defined_names_, remove_sheet_slicers(*this, index));

  std::vector<std::uint32_t> old_to_new(sheets_.size());
  for (std::uint32_t old_index = 0; old_index < old_to_new.size(); ++old_index) {
    old_to_new[old_index] = removed_index(old_index, index);
  }
  // A view naming the removed sheet moves to the sheet Excel activates in its place.
  const std::uint32_t successor = nearest_visible_sheet(index);
  const std::uint32_t new_last_index = static_cast<std::uint32_t>(sheets_.size() - 2U);
  old_to_new[index] = successor != index ? old_to_new[successor] : std::min(index, new_last_index);
  const bool removes_active = active_sheet() == index;
  remap_book_views_xml(book_views_xml_, old_to_new);

  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  std::vector<std::string_view> pre_removal_sheet_order;
  pre_removal_sheet_order.reserve(sheets_.size());
  for (const Sheet& sheet : sheets_) {
    pre_removal_sheet_order.push_back(sheet.name());
  }

  // Resolved before the sheet leaves: its chart parts are cleaned up after the erase.
  std::vector<std::string> drawing_candidates;
  if (auto drawing = sheet_drawing_path(*this, index); drawing && !drawing.value().empty()) {
    drawing_candidates.push_back(drawing.value());
  }

  // Decided before any formula is rewritten: the removed sheet's own cells would otherwise count.
  std::vector<std::string> rebound_names =
      referenced_local_names_of_removed_sheet(sheets_, defined_names_, tables_, index, chart_formulas.value());
  const auto has_book_namesake = [this](std::string_view name) {
    return detail::has_book_namesake(defined_names_, name);
  };
  std::vector<bool> promotes(defined_names_.size(), false);
  for (std::size_t i = 0; i < defined_names_.size(); ++i) {
    const DefinedName& entry = defined_names_[i];
    promotes[i] = entry.local_sheet_id == static_cast<std::int32_t>(index) && !has_book_namesake(entry.name) &&
                  std::any_of(rebound_names.begin(), rebound_names.end(),
                              [&entry](const std::string& name) { return collection_names::equal(name, entry.name); });
  }
  const parser::SheetRemovalTransform removal(pre_removal_sheet_order, index, std::move(rebound_names));
  std::vector<TableLoss> table_losses;
  for (const TableMetadata& table : tables_) {
    if (table.sheet_index == index) {
      table_losses.push_back(TableLoss{table.name, table.display_name, true, {}, {}});
    }
  }
  const sheet_names::NameList table_list = sheet_names::table_name_list(tables_);
  const TableLossTransform transform(removal, table_losses, &table_list);
  std::vector<std::uint32_t> dropped_cache_ids;
  bool defined_names_changed = false;
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, removed_name,
                              dropped_cache_ids, defined_names_changed);
  (void)defined_names_changed;
  RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));

  // The removed sheet's referenced names without a workbook-scoped namesake
  // become workbook-scoped in place; its other names disappear. Names scoped
  // to later sheets follow their owner as the sheet vector closes the gap.
  std::vector<DefinedName> retained;
  retained.reserve(defined_names_.size());
  for (std::size_t i = 0; i < defined_names_.size(); ++i) {
    DefinedName& entry = defined_names_[i];
    if (promotes[i]) {
      entry.local_sheet_id = -1;
    } else if (entry.local_sheet_id == static_cast<std::int32_t>(index)) {
      continue;
    }
    entry.local_sheet_id = removed_index(entry.local_sheet_id, static_cast<std::int32_t>(index));
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
  for (std::size_t sheet_pos = 0; sheet_pos < sheets_.size(); ++sheet_pos) {
    Sheet& sheet = sheets_[sheet_pos];
    std::vector<std::unique_ptr<pivot::PivotTable>> retained_pivots;
    retained_pivots.reserve(sheet.pivot_tables().size());
    // Position of each pre-removal pivot among the survivors; SIZE_MAX marks a dropped one.
    std::vector<std::size_t> pivot_new_pos(sheet.pivot_tables().size(), static_cast<std::size_t>(-1));
    std::size_t pivot_pos = 0;
    for (std::unique_ptr<pivot::PivotTable>& pivot_table : sheet.mutable_pivot_tables()) {
      const std::size_t old_pos = pivot_pos++;
      if (pivot_table == nullptr) {
        continue;
      }
      const std::uint32_t cache_id = pivot_table->pivot_cache_id();
      if (cache_survives(cache_id) &&
          std::find(dropped_cache_ids.begin(), dropped_cache_ids.end(), cache_id) == dropped_cache_ids.end()) {
        pivot_new_pos[old_pos] = retained_pivots.size();
        retained_pivots.push_back(std::move(pivot_table));
      }
    }
    sheet.mutable_pivot_tables() = std::move(retained_pivots);
    for (PassthroughPart& part : passthrough_parts_) {
      if (part.retained_origin == PassthroughPart::RetainedOrigin::kPivotTable &&
          part.origin_sheet_index == sheet_pos && part.origin_pivot_index < pivot_new_pos.size() &&
          pivot_new_pos[part.origin_pivot_index] != static_cast<std::size_t>(-1)) {
        part.origin_pivot_index = pivot_new_pos[part.origin_pivot_index];
      }
    }
  }
  // Parts owned by the removed sheet point past the end afterwards, so they read as stale.
  for (PassthroughPart& part : passthrough_parts_) {
    if (part.retained_origin != PassthroughPart::RetainedOrigin::kPivotTable) {
      continue;
    }
    if (part.origin_sheet_index == index) {
      part.origin_sheet_index = sheets_.size();
    } else if (part.origin_sheet_index > index) {
      --part.origin_sheet_index;
    }
  }

  // Erase only after all transforms have resolved against the pre-removal
  // order. One and only one final reindex sees the final sheet/name/table/
  // cache topology and dirties every surviving formula consistently.
  sheets_.erase(sheets_.begin() + static_cast<std::ptrdiff_t>(index));
  if (removes_active) {
    set_active_sheet(old_to_new[index]);
  }
  io::ooxml::remove_orphans(*this, std::move(drawing_candidates));
  reindex_all_formulas(sheets_, mutator, *this);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::move_sheet(std::uint32_t from_index, std::uint32_t to_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (static_cast<std::size_t>(from_index) >= sheets_.size() || static_cast<std::size_t>(to_index) >= sheets_.size()) {
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "index out of range",
                      "from=" + std::to_string(from_index) + " to=" + std::to_string(to_index) +
                          " sheet_count=" + std::to_string(sheets_.size()));
  }
  move_sheet_locked(from_index, to_index);
  return Expected<void, Error>::Ok();
}

void Workbook::move_sheet_locked(std::uint32_t from_index, std::uint32_t to_index) {
  if (from_index == to_index) {
    return;
  }

  std::vector<std::uint32_t> old_to_new(sheets_.size());
  for (std::uint32_t old_index = 0; old_index < old_to_new.size(); ++old_index) {
    old_to_new[old_index] = moved_index(old_index, from_index, to_index);
  }
  remap_book_views_xml(book_views_xml_, old_to_new);
  for (PassthroughPart& part : passthrough_parts_) {
    if (part.retained_origin == PassthroughPart::RetainedOrigin::kPivotTable &&
        part.origin_sheet_index < old_to_new.size()) {
      part.origin_sheet_index = old_to_new[part.origin_sheet_index];
    }
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
    if (entry.local_sheet_id >= 0) {
      entry.local_sheet_id = moved_index(entry.local_sheet_id, from_id, to_id);
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
}

}  // namespace formulon
