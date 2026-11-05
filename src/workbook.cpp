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
using detail::find_defined_name;
using detail::kFilterDatabaseName;
using detail::MayNameExternalBook;

// Forward declarations; defined further below alongside the other
// sheet-name helpers. Declared here so the early `add_sheet_validated` and
// `add_sheet_entered` definitions can call them.
Expected<void, Error> validate_sheet_name(std::string_view name);
Expected<void, Error> validate_entered_sheet_name(std::string_view name, ExcelHost host);

void mark_all_formulas_dirty_locked(const std::vector<Sheet>& sheets, const eval::RecalcEngine::LockedMutator& mutator);

}  // namespace

// `op`-prefixed range check for a mutator addressing one sheet.
Expected<void, Error> check_sheet_index(const char* op, std::size_t sheet_index, std::size_t sheet_count) {
  if (sheet_index >= sheet_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": sheet_index out of range",
                      "sheet_index=" + std::to_string(sheet_index));
  }
  return Expected<void, Error>::Ok();
}

namespace {

// `op`-prefixed range check for a mutator addressing one table.
Expected<void, Error> check_table_index(const char* op, std::size_t table_index, std::size_t table_count) {
  if (table_index >= table_count) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": table_index out of range",
                      "table_index=" + std::to_string(table_index));
  }
  return Expected<void, Error>::Ok();
}

// `op`-prefixed refusal of a write into part of a legacy array formula
// block: a block is only ever replaced as a whole, so a write into one of
// its non-anchor cells leaves it unchanged. A write to the anchor replaces
// the whole block.
Expected<void, Error> check_not_part_of_array(const char* op, const Sheet& sheet, std::uint32_t row,
                                              std::uint32_t col) {
  const std::optional<SpillFootprint> block = sheet.array_block_covering(row, col);
  if (block.has_value() && (block->anchor_row != row || block->anchor_col != col)) {
    return make_error(FormulonErrorCode::kPreconditionFailed, std::string(op) + ": cannot change part of an array",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}

// `op`-prefixed refusal of a write into a What-If data table's result range:
// unlike an array, the whole table is only ever replaced or removed as one.
Expected<void, Error> check_not_part_of_data_table(const char* op, const Sheet& sheet, std::uint32_t row,
                                                   std::uint32_t col) {
  if (sheet.data_table_covering(row, col) != nullptr) {
    return make_error(FormulonErrorCode::kPreconditionFailed, std::string(op) + ": cannot change part of a data table",
                      "row=" + std::to_string(row) + " col=" + std::to_string(col));
  }
  return Expected<void, Error>::Ok();
}

// Target and block checks shared by the cell value, text and formula setters.
Expected<void, Error> check_cell_writable(const char* op, const std::vector<Sheet>& sheets, std::size_t sheet_index,
                                          std::uint32_t row, std::uint32_t col) {
  RETURN_IF_ERROR(check_cell_target(op, sheet_index, sheets.size(), row, col));
  RETURN_IF_ERROR(check_not_part_of_array(op, sheets[sheet_index], row, col));
  return check_not_part_of_data_table(op, sheets[sheet_index], row, col);
}
}  // namespace

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
    return make_error(FormulonErrorCode::kInvalidArgument, "add_passthrough_part: empty part path");
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
      return make_error(FormulonErrorCode::kInvalidArgument, "add_passthrough_part: part already present",
                        "path=" + part.path);
    }
    passthrough_parts_.insert(passthrough_parts_.begin() + static_cast<std::ptrdiff_t>(first), std::move(part));
    return Expected<void, Error>::Ok();
  }

  // A caller may intentionally hand over an unsorted list through the setter.
  // Preserve the old append-then-stable-sort behaviour for that one fallback
  // path, and only mark the invariant restored after the successful sort.
  for (const PassthroughPart& existing : passthrough_parts_) {
    if (existing.path == part.path) {
      return make_error(FormulonErrorCode::kInvalidArgument, "add_passthrough_part: part already present",
                        "path=" + part.path);
    }
  }
  passthrough_parts_.push_back(std::move(part));
  sort_passthrough_parts(passthrough_parts_);
  passthrough_parts_sorted_ = true;
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::erase_passthrough_parts(std::vector<std::uint8_t> tombstones) {
  if (tombstones.size() != passthrough_parts_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument,
                      "erase_passthrough_parts: tombstone mask size does not match part count");
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
  if (sheet_by_name(name) != nullptr) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "add_sheet: name collides with an existing sheet",
                      "name=\"" + name + "\"");
  }
  RETURN_IF_ERROR(check_sheet_headroom(sheets_.size()));
  append_sheet(sheets_, engine_->locked_mutator(), *this, std::move(name), new_sheet_profile());
  return &sheets_.back();
}

Expected<void, Error> Workbook::check_entered_new_sheet_name(std::string_view name) const {
  RETURN_IF_ERROR(validate_sheet_name(name));
  RETURN_IF_ERROR(validate_entered_sheet_name(name, excel_profile_.host));
  if (sheet_by_name(name) != nullptr) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "copy_sheet: name collides with an existing sheet",
                      "name=\"" + std::string(name) + "\"");
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::check_sheet_capacity() const {
  return check_sheet_headroom(sheets_.size());
}

Expected<Sheet*, Error> Workbook::add_sheet_entered(std::string name) {
  RETURN_IF_ERROR(validate_entered_sheet_name(name, excel_profile_.host));
  return add_sheet_validated(std::move(name));
}

namespace {

/// `Sheet1!$B$2:$F$12`, quoting the sheet name where Excel would.
std::string filter_database_formula(std::string_view sheet_name, const MergeRange& range) {
  std::string out;
  parser::append_sheet_name(sheet_name, false, out);
  out.push_back('!');
  const auto append_cell = [&out](std::uint32_t row, std::uint32_t col) {
    out.push_back('$');
    a1::append_column_letters(out, col);
    out.push_back('$');
    out.append(std::to_string(row + 1U));
  };
  append_cell(range.first_row, range.first_col);
  out.push_back(':');
  append_cell(range.last_row, range.last_col);
  return out;
}

}  // namespace

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
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "rename_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  // Validate the new name (non-empty, ≤ 31 code units, no forbidden
  // characters, none of the names Excel's naming UI refuses) via the shared
  // validators, so a rename and `add_sheet_entered` agree.
  RETURN_IF_ERROR(validate_sheet_name(new_name));
  RETURN_IF_ERROR(validate_entered_sheet_name(new_name, excel_profile_.host));
  // Collision check (Unicode simple-fold). A no-op rename — same case-fold
  // as the current name — bypasses the collision check so callers can
  // change only the casing.
  // Sheet names are unique under folding, so the one match is either the
  // sheet being renamed or a genuine collision.
  const std::size_t idx = sheet_index_by_name(new_name);
  if (idx != static_cast<std::size_t>(-1) && idx != static_cast<std::size_t>(index)) {
    return make_error(FormulonErrorCode::kInvalidSheetName, "rename_sheet: name collides with another sheet",
                      "name=\"" + new_name + "\" colliding_index=" + std::to_string(idx));
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
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "remove_sheet: index out of range",
                      "index=" + std::to_string(index) + " sheet_count=" + std::to_string(sheets_.size()));
  }
  if (sheets_.size() <= 1U) {
    return make_error(FormulonErrorCode::kCannotRemoveLastSheet, "remove_sheet: cannot remove the only sheet",
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
                  std::any_of(rebound_names.begin(), rebound_names.end(), [&entry](const std::string& name) {
                    return strings::case_insensitive_eq(name, entry.name);
                  });
  }
  const parser::SheetRemovalTransform removal(pre_removal_sheet_order, index, std::move(rebound_names));
  std::vector<TableLoss> table_losses;
  for (const TableMetadata& table : tables_) {
    if (table.sheet_index == index) {
      table_losses.push_back(TableLoss{table.name, table.display_name, true, {}});
    }
  }
  const TableLossTransform transform(removal, table_losses);
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
    return make_error(FormulonErrorCode::kSheetIndexOutOfRange, "move_sheet: index out of range",
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

Expected<void, Error> Workbook::set_sheet_auto_filter(std::size_t sheet_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_sheet_index("set_sheet_auto_filter", sheet_index, sheets_.size()));
  RETURN_IF_ERROR(validate_auto_filter(filter));
  std::string formula = filter_database_formula(sheets_[sheet_index].name(), filter.range);
  sheets_[sheet_index].set_auto_filter(std::move(filter));
  const auto scope = static_cast<std::int32_t>(sheet_index);
  const auto it = find_defined_name(defined_names_, kFilterDatabaseName, scope);
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
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator, defined_names_);
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
  mark_row_visibility_dependents_dirty_locked(sheets_, mutator, defined_names_);
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
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::set_table_auto_filter(std::size_t table_index, AutoFilter filter) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_table_index("set_table_auto_filter", table_index, tables_.size()));
  RETURN_IF_ERROR(validate_auto_filter(filter));
  tables_[table_index].auto_filter_xml.set(std::move(filter));
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::remove_table_auto_filter(std::size_t table_index) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  RETURN_IF_ERROR(check_table_index("remove_table_auto_filter", table_index, tables_.size()));
  tables_[table_index].auto_filter_xml.reset();
  mark_row_visibility_dependents_dirty_locked(sheets_, engine_->locked_mutator(), defined_names_);
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

namespace {

/// Maps a structured reference's table name from either of a table's old
/// names to its new display name.
class TableRenameTransform final : public parser::RefTransform {
 public:
  TableRenameTransform(std::string_view old_name, std::string_view old_display_name, std::string_view new_name)
      : old_name_(old_name), old_display_name_(old_display_name), new_name_(new_name) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view /*column*/) const override {
    if (strings::case_insensitive_eq(table, old_name_) || strings::case_insensitive_eq(table, old_display_name_)) {
      return parser::StructuredRefTarget{new_name_, {}};
    }
    return parser::StructuredRefTarget{table, {}};
  }

 private:
  std::string_view old_name_;
  std::string_view old_display_name_;
  std::string_view new_name_;
};

}  // namespace

Expected<void, Error> Workbook::rename_table(std::size_t index, std::string name, std::string display_name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (index >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "rename_table: index out of range",
                      "index=" + std::to_string(index) + " table_count=" + std::to_string(tables_.size()));
  }
  const std::string old_name = tables_[index].name;
  const std::string old_display_name = tables_[index].display_name;
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const TableRenameTransform transform(old_name, old_display_name, display_name);
  std::vector<std::uint32_t> ignored_cache_ids;
  bool defined_names_changed = false;
  rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, {},
                              ignored_cache_ids, defined_names_changed);
  RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));
  for (std::unique_ptr<pivot::PivotCache>& cache : pivot_caches_) {
    if (cache == nullptr) {
      continue;
    }
    pivot::WorksheetSource& source = cache->mutable_worksheet_source();
    if (!source.name.empty() && (strings::case_insensitive_eq(source.name, old_name) ||
                                 strings::case_insensitive_eq(source.name, old_display_name))) {
      source.name = display_name;
    }
  }
  tables_[index].name = std::move(name);
  tables_[index].display_name = std::move(display_name);
  if (defined_names_changed) {
    mark_all_formulas_dirty_locked(sheets_, mutator);
  }
  reindex_formulas_referencing_table(sheets_, mutator, *this,
                                     {old_name, old_display_name, tables_[index].name, tables_[index].display_name});
  return Expected<void, Error>::Ok();
}

namespace {

/// Maps a structured reference through one table naming `old_column` to the
/// same reference naming `new_column`, re-spelled in the stored form.
class TableColumnRenameTransform final : public parser::RefTransform {
 public:
  TableColumnRenameTransform(std::string_view table_name, std::string_view table_display_name,
                             std::string_view old_column, std::string_view new_column)
      : table_name_(table_name),
        table_display_name_(table_display_name),
        old_column_(old_column),
        new_column_(new_column) {}

  std::optional<parser::Reference> apply(const parser::Reference& ref) const override { return ref; }

  std::optional<parser::StructuredRefTarget> apply_structured_ref(std::string_view table,
                                                                  std::string_view column) const override {
    if (!strings::case_insensitive_eq(table, table_name_) &&
        !strings::case_insensitive_eq(table, table_display_name_)) {
      return parser::StructuredRefTarget{table, {}};
    }
    std::optional<parser::StructuredRefPayload> payload = parser::parse_structured_ref_payload_text(column);
    if (!payload) {
      return parser::StructuredRefTarget{table, {}};
    }
    bool renamed = false;
    for (std::string& name : payload->columns) {
      if (strings::case_insensitive_eq(name, old_column_)) {
        name.assign(new_column_);
        renamed = true;
      }
    }
    if (!renamed) {
      return parser::StructuredRefTarget{table, {}};
    }
    return parser::StructuredRefTarget{table, parser::stored_payload(*payload)};
  }

 private:
  std::string_view table_name_;
  std::string_view table_display_name_;
  std::string_view old_column_;
  std::string_view new_column_;
};

}  // namespace

Expected<std::string, Error> Workbook::rename_table_column(std::size_t table, std::size_t column,
                                                           std::string_view name) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  if (table >= tables_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "rename_table_column: table index out of range",
                      "table=" + std::to_string(table) + " table_count=" + std::to_string(tables_.size()));
  }
  std::vector<TableColumn>& columns = tables_[table].columns;
  if (column >= columns.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "rename_table_column: column index out of range",
                      "column=" + std::to_string(column) + " column_count=" + std::to_string(columns.size()));
  }
  const auto taken = [&columns, column](std::string_view candidate) {
    for (std::size_t i = 0; i < columns.size(); ++i) {
      if (i != column && strings::case_insensitive_eq(columns[i].name, candidate)) {
        return true;
      }
    }
    return false;
  };
  // An empty header takes `<word>N` from N = 1; a duplicate takes `<name>N` from N = 2.
  std::string final_name(name);
  if (name.empty() || taken(name)) {
    const std::string stem =
        name.empty() ? std::string(new_workbook_defaults(excel_profile_.locale).column_word) : std::string(name);
    for (std::size_t n = name.empty() ? 1U : 2U;; ++n) {
      final_name = stem + std::to_string(n);
      if (!taken(final_name)) {
        break;
      }
    }
  }

  const TableMetadata& meta = tables_[table];
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  const std::string old_name = columns[column].name;
  const bool renamed = final_name != old_name;
  if (renamed) {
    const TableColumnRenameTransform transform(meta.name, meta.display_name, old_name, final_name);
    std::vector<std::uint32_t> ignored_cache_ids;
    bool defined_names_changed = false;
    rewrite_workbook_references(sheets_, defined_names_, tables_, pivot_caches_, transform, mutator, {}, {}, {},
                                ignored_cache_ids, defined_names_changed);
    RETURN_IF_ERROR(chart_ref::rewrite_all(*this, transform));
    tables_[table].columns[column].name = final_name;
    if (defined_names_changed) {
      mark_all_formulas_dirty_locked(sheets_, mutator);
    }
  }
  const std::optional<MergeRange> rect = parse_a1_rectangle(tables_[table].ref);
  if (tables_[table].header_row && rect.has_value() && tables_[table].sheet_index < sheets_.size()) {
    const std::uint32_t header_col = rect->first_col + static_cast<std::uint32_t>(column);
    sheets_[tables_[table].sheet_index].set_cell_text(rect->first_row, header_col, final_name);
    const eval::CellNodeId node = make_node(tables_[table].sheet_index, rect->first_row, header_col);
    for (const eval::CellNodeId dependent : mutator.dep_graph().dependents_of_ref(node)) {
      mutator.mark_dirty(dependent);
    }
    mutator.mark_range_dependents_dirty(node);
  }
  if (renamed) {
    reindex_formulas_referencing_table(sheets_, mutator, *this, {tables_[table].name, tables_[table].display_name});
  }
  return final_name;
}

const Sheet* Workbook::sheet_by_name(std::string_view name) const noexcept {
  const std::size_t index = sheet_index_by_name(name);
  return index == static_cast<std::size_t>(-1) ? nullptr : &sheets_[index];
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
                                             const std::vector<Sheet>& sheets, const BlockedSpillFootprint& rectangle) {
  for (const CellAddress anchor : sheets[sheet_index].blocked_spill_anchors_intersecting(
           rectangle.anchor_row, rectangle.anchor_col, rectangle.rows, rectangle.cols)) {
    mutator.mark_dirty(make_node(sheet_index, anchor.row, anchor.col));
  }
}

void mark_blocked_spill_anchors_intersecting(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                                             const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  mark_blocked_spill_anchors_intersecting(mutator, sheet_index, sheets, BlockedSpillFootprint{row, col, 1U, 1U});
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

// Dirties the cell, its dependents and every spill anchor the write wakes,
// and drops the cell's outgoing edges. Precondition: caller holds the engine
// mutex; call before the sheet write.
void prepare_cell_overwrite(const eval::RecalcEngine::LockedMutator& mutator, std::size_t sheet_index,
                            const std::vector<Sheet>& sheets, std::uint32_t row, std::uint32_t col) {
  const eval::CellNodeId node = make_node(sheet_index, row, col);
  mutator.mark_dirty(node);
  mark_dependents_dirty(mutator, node);
  mark_blocked_spill_anchors_released_by_cell(mutator, sheet_index, sheets, row, col);
  mutator.clear_cell_dependencies(node);
  mark_spill_anchor_dirty_if_covered(mutator, sheet_index, sheets, row, col);
}

}  // namespace

Expected<void, Error> Workbook::set_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                               Value value) {
  RETURN_IF_ERROR(check_cell_writable("set_cell_value", sheets_, sheet_index, row, col));
  put_cell_value(sheet_index, row, col, value);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_value(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, Value value) {
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
    prepare_cell_overwrite(mutator, sheet_index, sheets_, row, col);
    sheets_[sheet_index].set_cell_value(row, col, value);
  }
}

Expected<void, Error> Workbook::set_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                              std::string_view text) {
  RETURN_IF_ERROR(check_cell_writable("set_cell_text", sheets_, sheet_index, row, col));
  put_cell_text(sheet_index, row, col, text);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_text(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string_view text) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  prepare_cell_overwrite(mutator, sheet_index, sheets_, row, col);
  sheets_[sheet_index].set_cell_text(row, col, text);
}

void Workbook::erase_cells_in(std::size_t sheet_index, const MergeRange& rect) {
  Sheet& sheet = sheets_[sheet_index];
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  const eval::RecalcEngine::LockedMutator mutator = engine_->locked_mutator();
  std::vector<CellAddress> stored;
  for (std::uint64_t row = rect.first_row; row <= rect.last_row; ++row) {
    const auto it = sheet.rows().find(static_cast<std::uint32_t>(row));
    if (it == sheet.rows().end()) {
      continue;
    }
    for (const RowCells::Entry entry :
         it->second.entries_in(rect.first_col, static_cast<std::uint64_t>(rect.last_col) + 1U)) {
      stored.push_back(CellAddress{static_cast<std::uint32_t>(row), entry.col});
    }
  }
  // Like a literal write, the erase re-evaluates whatever read the cell.
  for (const CellAddress address : stored) {
    prepare_cell_overwrite(mutator, sheet_index, sheets_, address.row, address.col);
    sheet.erase_cell(address.row, address.col);
  }
}

void Workbook::apply_legacy_implicit_intersections() {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  for (std::size_t sheet_index = 0; sheet_index < sheets_.size(); ++sheet_index) {
    const io::xlsb::NameShapes name_shapes = io::name_shapes(*this, sheet_index);
    Sheet& sheet = sheets_[sheet_index];
    for (const CellAddress address : sheet.all_formula_cells()) {
      const Cell* cell = sheet.cell_at(address.row, address.col);
      // A dynamic-array formula already says where it intersects, and a CSE
      // block evaluates as an array.
      if (cell == nullptr || cell->dynamic_array || cell->array_block_rows != 0U) {
        continue;
      }
      std::string text = cell->formula_text;
      const std::size_t body_at = !text.empty() && text.front() == '=' ? 1U : 0U;
      Arena arena;
      const parser::AstNode* root = parser::parse_strict(std::string_view(text).substr(body_at), arena);
      if (root == nullptr) {
        continue;
      }
      // Byte offsets into the body; a node's range counts UTF-16 units.
      const std::string_view body = std::string_view(text).substr(body_at);
      std::vector<std::uint32_t> at;
      for (const parser::AstNode* node : io::xlsb::legacy_intersections(*root, name_shapes)) {
        if (node->kind() != parser::NodeKind::ImplicitIntersection) {
          // A parenthesised operand takes the `@` inside its parentheses: `(@A1:A3)`.
          auto start = static_cast<std::uint32_t>(utf16_to_byte_offset(body, node->range().start));
          for (std::uint16_t pairs = node->paren_depth(); pairs > 0U && body_at + start < text.size(); ++start) {
            if (text[body_at + start] == '(') {
              --pairs;
            }
          }
          at.push_back(start);
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

std::string Workbook::normalize_formula_text(std::string formula, FormulaSource source) {
  std::string text = parser::spell_storage_operators(parser::strip_storage_prefixes(formula, &io::has_storage_prefix));
  if (source == FormulaSource::kEntered) {
    text = parser::spell_entered_formula(text);
  }
  return text;
}

Expected<void, Error> Workbook::set_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                                 std::string formula, FormulaSource source) {
  RETURN_IF_ERROR(check_cell_writable("set_cell_formula", sheets_, sheet_index, row, col));
  put_cell_formula(sheet_index, row, col, std::move(formula), source, 0U, 0U);
  return Expected<void, Error>::Ok();
}

void Workbook::put_cell_formula(std::size_t sheet_index, std::uint32_t row, std::uint32_t col, std::string formula,
                                FormulaSource source, std::uint32_t block_rows, std::uint32_t block_cols) {
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
  // formula bar shows. Entered text also takes the whitespace and number
  // literal spelling Excel stores.
  formula = normalize_formula_text(std::move(formula), source);
  if (source == FormulaSource::kEntered && !file_name_.empty()) {
    formula = parser::bind_self_book_names(formula, file_name_);
  }

  const eval::CellNodeId node = make_node(sheet_index, row, col);

  // Parse the formula in a throwaway arena to extract its dependency list.
  // Strip a leading '=' so the parser sees a bare expression (the tokenizer
  // accepts both shapes, but the dep extractor walks the AST regardless).
  const std::string_view src = strip_formula_prefix(formula);

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
        row, col, block_rows == 0U && root != nullptr && io::entered_as_dynamic_array(*this, sheet_index, *root));
    sheets_[sheet_index].set_cell_array_block(row, col, block_rows, block_cols);

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

Expected<void, Error> Workbook::recalc_parallel(const eval::FunctionRegistry& registry,
                                                const eval::SchedulerConfig& cfg, eval::SchedulerStats* stats) {
  eval::SchedulerStats local_stats;
  eval::SchedulerStats& pass_stats = stats != nullptr ? *stats : local_stats;
  const bool had_dirty = has_dirty_cells();
  RETURN_IF_ERROR(eval::recalc_parallel(*this, registry, cfg, &pass_stats));
  return recalc_data_tables_after_pass(had_dirty || pass_stats.cells_evaluated != 0U, registry);
}

namespace {

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
  RETURN_IF_ERROR(check_cell_target("set_cell_xf_index", sheet_index, sheets_.size(), row, col));
  // A style write can grow the sheet's sparse row store, so serialize it
  // with recalc just like all other workbook-level cell mutations. The
  // sheet-level setter deliberately bypasses literal-write spill invalidation
  // so formatting a live spill phantom preserves the dynamic array.
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].set_cell_xf_index(row, col, xf_index);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::clear_cell_xf_index(std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  RETURN_IF_ERROR(check_cell_target("clear_cell_xf_index", sheet_index, sheets_.size(), row, col));
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  sheets_[sheet_index].clear_cell_xf_index(row, col);
  return Expected<void, Error>::Ok();
}

namespace {

Expected<void, Error> check_style_rect(const char* op, std::size_t sheet_index, std::size_t sheet_count,
                                       const StyleRect& rect) {
  RETURN_IF_ERROR(check_cell_target(op, sheet_index, sheet_count, rect.last_row, rect.last_col));
  if (rect.first_row > rect.last_row || rect.first_col > rect.last_col) {
    return make_error(FormulonErrorCode::kInvalidArgument, std::string(op) + ": inverted range");
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<std::vector<std::uint32_t>, Error> Workbook::range_style_bases(std::size_t sheet_index, const StyleRect& rect,
                                                                        std::vector<StyleCell> excluded) const {
  RETURN_IF_ERROR(check_style_rect("range_style_bases", sheet_index, sheets_.size(), rect));
  return formulon::range_style_bases(sheets_[sheet_index], rect, std::move(excluded));
}

Expected<void, Error> Workbook::remap_range_style(std::size_t sheet_index, const StyleRect& rect,
                                                  std::vector<XfRemap> map) {
  RETURN_IF_ERROR(check_style_rect("remap_range_style", sheet_index, sheets_.size(), rect));
  sort_by_index(map, [](const XfRemap& lhs, const XfRemap& rhs) { return lhs.from < rhs.from; });
  for (std::size_t i = 0; i < map.size(); ++i) {
    if (map[i].to >= styles_.cell_xfs.size() || (i > 0U && map[i - 1U].from == map[i].from)) {
      return make_error(FormulonErrorCode::kInvalidArgument, "remap_range_style: repeated or out-of-range xf");
    }
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return formulon::remap_range_style(sheets_[sheet_index], rect, map, kMaxStyledRangeCells);
}

Expected<std::vector<std::uint8_t>, Error> Workbook::snapshot_range_style(std::size_t sheet_index,
                                                                          const StyleRect& rect) const {
  RETURN_IF_ERROR(check_style_rect("snapshot_range_style", sheet_index, sheets_.size(), rect));
  return formulon::snapshot_range_style(sheets_[sheet_index], rect);
}

Expected<void, Error> Workbook::restore_range_style(std::size_t sheet_index, const std::uint8_t* bytes,
                                                    std::size_t len) {
  if (sheet_index >= sheets_.size()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "restore_range_style: sheet_index out of range");
  }
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  return formulon::restore_range_style(sheets_[sheet_index], bytes, len, styles_.cell_xfs.size());
}

}  // namespace formulon
