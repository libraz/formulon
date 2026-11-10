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
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "chart/chart_ref.h"
#include "defined_name.h"
#include "drawing/drawing_edit.h"
#include "drawing/slicer_edit.h"
#include "eval/dep_graph.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/recalc_reentry.h"
#include "eval/scheduler.h"
#include "external_link.h"
#include "io/auto_filter_xml.h"
#include "io/dynamic_array_formula.h"
#include "io/format_detect.h"
#include "io/future_functions.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml_writer.h"
#include "io/sparkline_groups.h"
#include "io/theme_part.h"
#include "io/workbook_kind_ooxml.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/writer.h"
#include "new_workbook_defaults.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/external_book_spelling.h"
#include "parser/formula_prefix.h"
#include "parser/literal_spelling.h"
#include "parser/parser.h"
#include "parser/parser_detail.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "parser/structured_ref_payload.h"
#include "parser/structured_ref_spelling.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "sheet_name.h"
#include "styles.h"
#include "table.h"
#include "utils/a1_column.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/index_sort.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/text_ops.h"
#include "utils/utf8_length.h"
#include "value.h"
#include "workbook_formula_index.h"
#include "workbook_internal.h"
#include "workbook_ref_rewrite.h"
#include "workbook_sheet_mutation.h"
#include "workbook_table_loss.h"

namespace formulon {
namespace {
using detail::check_cell_target;
}  // namespace

// Range check for a mutator addressing one sheet.
Expected<void, Error> check_sheet_index(std::size_t sheet_index, std::size_t sheet_count) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index is out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  return Expected<void, Error>::Ok();
}

Workbook::Workbook() : engine_(std::make_unique<eval::RecalcEngine>()), kind_(WorkbookKind::kXlsx) {}
Workbook::Workbook(Workbook&&) noexcept = default;
Workbook& Workbook::operator=(Workbook&&) noexcept = default;
Workbook::~Workbook() = default;

void Workbook::set_passthrough_parts(std::vector<PassthroughPart> parts) {
  passthrough_parts_ = std::move(parts);
  passthrough_parts_sorted_ = true;
  for (std::size_t i = 1; i < passthrough_parts_.size(); ++i) {
    if (passthrough_parts_[i].path < passthrough_parts_[i - 1U].path) {
      passthrough_parts_sorted_ = false;
      break;
    }
  }
}

Expected<void, Error> Workbook::add_passthrough_part(PassthroughPart part) {
  if (part.path.empty()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Part path is empty");
  }

  if (passthrough_parts_sorted_) {
    // Readers and most callers hand us path-sorted parts. Locate the insertion
    // point directly; equal paths are rejected before the vector is mutated.
    std::size_t first = 0;
    std::size_t last = passthrough_parts_.size();
    while (first < last) {
      const std::size_t middle = first + (last - first) / 2U;
      if (passthrough_parts_[middle].path < part.path) {
        first = middle + 1U;
      } else {
        last = middle;
      }
    }
    if (first < passthrough_parts_.size() && passthrough_parts_[first].path == part.path) {
      return make_error(FormulonErrorCode::kInvalidArgument, "Part is already present", "path=" + part.path);
    }
    passthrough_parts_.insert(passthrough_parts_.begin() + static_cast<std::ptrdiff_t>(first), std::move(part));
    return Expected<void, Error>::Ok();
  }

  // A caller may intentionally hand over an unsorted list through the setter.
  // Preserve the old append-then-stable-sort behaviour for that one fallback
  // path, and only mark the invariant restored after the successful sort.
  for (const PassthroughPart& existing : passthrough_parts_) {
    if (existing.path == part.path) {
      return make_error(FormulonErrorCode::kInvalidArgument, "Part is already present", "path=" + part.path);
    }
  }
  passthrough_parts_.push_back(std::move(part));
  sort_passthrough_parts(passthrough_parts_);
  passthrough_parts_sorted_ = true;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::erase_passthrough_parts(std::vector<std::uint8_t> tombstones) {
  if (tombstones.size() != passthrough_parts_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Tombstone mask size does not match the part count");
  }
  auto write = passthrough_parts_.begin();
  for (std::size_t i = 0; i < passthrough_parts_.size(); ++i) {
    if (tombstones[i] != 0U) {
      continue;
    }
    auto read = passthrough_parts_.begin() + static_cast<std::ptrdiff_t>(i);
    if (write != read) {
      *write = std::move(*read);
    }
    ++write;
  }
  passthrough_parts_.erase(write, passthrough_parts_.end());
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::replace_passthrough_part(std::string_view path, std::vector<std::uint8_t> bytes) {
  for (PassthroughPart& existing : passthrough_parts_) {
    if (existing.path == path) {
      existing.bytes = std::move(bytes);
      return Expected<void, Error>::Ok();
    }
  }
  return make_error(FormulonErrorCode::kInvalidArgument, "No part has that path", "path=" + std::string(path));
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

Expected<void, Error> Workbook::set_theme_script_font(bool major, std::string_view script, std::string_view typeface) {
  return io::set_theme_script_font(*this, major, script, typeface);
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

Workbook Workbook::create_for_profile(const ExcelProfile& profile) {
  Workbook wb;
  wb.excel_profile_ = profile;
  wb.new_sheet_profile_ = profile;
  wb.sheets_.emplace_back(Sheet{std::string(new_workbook_defaults(profile.locale).sheet_name)});
  seed_default_styles(wb.styles_);
  apply_new_workbook_defaults(profile, wb.styles_, wb.sheets_.front());
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

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  const std::size_t index = sheet_index_by_name(name);
  return index == kNoSheet ? nullptr : &sheets_[index];
}

std::size_t Workbook::sheet_index_by_name(std::string_view name) const noexcept {
  return detail::find_sheet_index(sheets_, name);
}

std::size_t Workbook::approximate_memory_bytes() const noexcept {
  // Per-row cost of the cell store's hash map beyond the row's arrays: one
  // node holding the key and the `RowCells` header (two vectors, 48 B on
  // 64-bit), plus the bucket slot pointing at it. A constant is enough — the
  // figure is a pressure signal, and no standard library exposes its real
  // per-node overhead.
  constexpr std::size_t kRowNodeOverheadBytes = sizeof(std::uint32_t) + 2U * sizeof(void*) + sizeof(RowCells);

  std::size_t total = sizeof(Workbook);

  for (const Sheet& sheet_ref : sheets_) {
    total += sizeof(Sheet) + sheet_ref.name().capacity();
    for (const auto& [row, cells] : sheet_ref.rows()) {
      (void)row;
      total += kRowNodeOverheadBytes + cells.capacity_bytes();
      for (const auto& [col, cell] : cells) {
        (void)col;
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
  return make_error(FormulonErrorCode::kInvalidArgument, "format is not a supported save format",
                    "context=workbook_save_as");
}

bool Workbook::has_dirty_cells() const {
  // A nested call already holds the mutex; the engine rejects it as re-entrant.
  if (eval::detail::g_in_recalc) {
    return false;
  }
  // Another thread's mutation writes the dirty set under the engine mutex.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return engine_->dirty().size() != 0U;
}

Expected<eval::RecalcStats, Error> Workbook::recalc(const eval::FunctionRegistry& registry) {
  // A write that dirtied only a data table's header or input cell evaluates nothing, yet changes the table.
  const bool had_dirty = has_dirty_cells();
  auto result = engine_->recalc(*this, registry);
  if (!result) {
    return result;
  }
  RETURN_IF_ERROR(recalc_data_tables_after_pass(had_dirty || result.value().cells_evaluated != 0U, registry));
  return result;
}

Expected<void, Error> Workbook::ensure_calculated(const eval::FunctionRegistry& registry) {
  if (calc_mode_ == CalcMode::kManual) {
    return {};
  }
  // Inside a recalc the dirty set cannot be read; recalc() rejects the nested pass as re-entrant.
  if (!eval::detail::g_in_recalc && !has_dirty_cells()) {
    return {};
  }
  auto result = recalc(registry);
  if (!result) {
    return result.error();
  }
  return {};
}

Expected<void, Error> Workbook::recalc_parallel(const eval::FunctionRegistry& registry,
                                                const eval::SchedulerConfig& cfg, eval::SchedulerStats* stats) {
  eval::SchedulerStats local_stats;
  eval::SchedulerStats& pass_stats = stats != nullptr ? *stats : local_stats;
  const bool had_dirty = has_dirty_cells();
  RETURN_IF_ERROR(eval::recalc_parallel(*this, registry, cfg, &pass_stats));
  return recalc_data_tables_after_pass(had_dirty || pass_stats.cells_evaluated != 0U, registry);
}

namespace {

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
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
}

void Workbook::mark_all_formulas_dirty() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  mark_all_formulas_dirty_locked(sheets_, engine_->locked_mutator());
}

Expected<void, Error> Workbook::set_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                  std::uint32_t xf_index) {
  RETURN_IF_ERROR(check_cell_target(sheet_index, sheets_.size(), row, col));
  // A style write can grow the sheet's sparse row store, so serialize it
  // with recalc just like all other workbook-level cell mutations. The
  // sheet-level setter deliberately bypasses literal-write spill invalidation
  // so formatting a live spill phantom preserves the dynamic array.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].set_cell_xf_index(row, col, xf_index);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  RETURN_IF_ERROR(check_cell_target(sheet_index, sheets_.size(), row, col));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].clear_cell_xf_index(row, col);
  return Expected<void, Error>::Ok();
}

namespace {

Expected<void, Error> check_style_rect(std::size_t sheet_index, std::size_t sheet_count, const StyleRect& rect) {
  RETURN_IF_ERROR(check_cell_target(sheet_index, sheet_count, rect.last_row, rect.last_col));
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range is inverted");
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::vector<std::uint32_t>, Error> Workbook::range_style_bases(std::size_t sheet_index, const StyleRect& rect,
                                                                        std::vector<StyleCell> excluded) const {
  RETURN_IF_ERROR(check_style_rect(sheet_index, sheets_.size(), rect));
  return formulon::range_style_bases(sheets_[sheet_index], rect, std::move(excluded));
}

Expected<void, Error> Workbook::remap_range_style(std::size_t sheet_index, const StyleRect& rect,
                                                  std::vector<XfRemap> map) {
  RETURN_IF_ERROR(check_style_rect(sheet_index, sheets_.size(), rect));
  sort_by_index(map, [](const XfRemap& lhs, const XfRemap& rhs) { return lhs.from < rhs.from; });
  for (std::size_t i = 0; i < map.size(); ++i) {
    if (map[i].to >= styles_.cell_xfs.size() || (i > 0U && map[i - 1U].from == map[i].from)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "pairs holds a repeated or out-of-range xf");
    }
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return formulon::remap_range_style(sheets_[sheet_index], rect, map, kMaxStyledRangeCells);
}

Expected<std::vector<std::uint8_t>, Error> Workbook::snapshot_range_style(std::size_t sheet_index,
                                                                          const StyleRect& rect) const {
  RETURN_IF_ERROR(check_style_rect(sheet_index, sheets_.size(), rect));
  return formulon::snapshot_range_style(sheets_[sheet_index], rect);
}

Expected<void, Error> Workbook::restore_range_style(std::size_t sheet_index, const std::uint8_t* bytes,
                                                    std::size_t len) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "sheet_index out of range");
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return formulon::restore_range_style(sheets_[sheet_index], bytes, len, styles_.cell_xfs.size());
}

namespace {

Error sheet_out_of_range(std::size_t sheet, std::size_t count) {
  return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "sheet index out of range",
                    "sheet=" + std::to_string(sheet) + " sheet_count=" + std::to_string(count));
}

// `id`'s value, assigning the next session id first when it has none.
std::uint32_t assigned(const Workbook& wb, RuntimeId& id) {
  if (id.value() == 0U) {
    id.assign(wb.allocate_runtime_id());
  }
  return id.value();
}

}  // namespace

std::uint32_t Workbook::validation_runtime_id(std::size_t sheet, std::size_t index) const {
  return assigned(*this, sheets_[sheet].validations()[index].runtime_id);
}

std::optional<std::size_t> Workbook::find_validation(std::size_t sheet, std::uint32_t id) const noexcept {
  if (sheet >= sheets_.size() || id == 0U) {
    return std::nullopt;
  }
  const std::vector<DataValidation>& validations = sheets_[sheet].validations();
  for (std::size_t i = 0; i < validations.size(); ++i) {
    if (validations[i].runtime_id.value() == id) {
      return i;
    }
  }
  return std::nullopt;
}

Expected<std::uint32_t, Error> Workbook::add_validation(std::size_t sheet, DataValidation validation) {
  if (sheet >= sheets_.size()) {
    return sheet_out_of_range(sheet, sheets_.size());
  }
  validation.runtime_id.assign(allocate_runtime_id());
  std::vector<DataValidation>& validations = sheets_[sheet].mutable_validations();
  validations.push_back(std::move(validation));
  return validations.back().runtime_id.value();
}

Expected<void, Error> Workbook::remove_validation(std::size_t sheet, std::uint32_t id) {
  if (sheet >= sheets_.size()) {
    return sheet_out_of_range(sheet, sheets_.size());
  }
  const std::optional<std::size_t> index = find_validation(sheet, id);
  if (!index.has_value()) {
    return make_error(FormulonErrorCode::kNotFound, "id does not name a validation", "id=" + std::to_string(id));
  }
  std::vector<DataValidation>& validations = sheets_[sheet].mutable_validations();
  validations.erase(validations.begin() + static_cast<std::ptrdiff_t>(*index));
  return Expected<void, Error>::Ok();
}

std::uint32_t Workbook::sparkline_group_runtime_id(std::size_t sheet, std::size_t index) const {
  std::vector<RuntimeId>& ids = sheets_[sheet].sparkline_group_ids();
  if (index >= ids.size()) {
    ids.resize(index + 1U);
  }
  return assigned(*this, ids[index]);
}

std::optional<std::size_t> Workbook::find_sparkline_group(std::size_t sheet, std::uint32_t id) const {
  if (sheet >= sheets_.size() || id == 0U) {
    return std::nullopt;
  }
  const std::vector<RuntimeId>& ids = sheets_[sheet].sparkline_group_ids();
  const std::size_t count = io::read_sparkline_groups(sheets_[sheet].ext_lst_xml()).size();
  for (std::size_t i = 0; i < ids.size() && i < count; ++i) {
    if (ids[i].value() == id) {
      return i;
    }
  }
  return std::nullopt;
}

}  // namespace formulon
