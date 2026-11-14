// Save-time plan for the named cell styles: which `<cellStyles>` /
// `<cellStyleXfs>` entries a save writes, under which names, in which order,
// and the `xfId` each `<cellXfs>` entry takes. The model is never touched.
//
// The plan reproduces what Excel writes when it saves the same book under the
// profile's UI locale: built-ins take the locale name of their builtinId,
// styles whose saved names are equal under the profile's compare collation
// merge (name and flags of the first, style xf of the last; Normal wins a
// group it is in), unused built-ins drop, `<cellStyles>` sorts by saved name
// under the profile's sort collation and `<cellStyleXfs>` puts Normal first,
// then the outline styles, then the rest in model order.
//
// `collect_used_cell_xfs` is the single enumeration of the structures that
// hold a cell xf. A new structure holding a cell xf must be added there, or
// the styles it uses are dropped on save.
//
// The plan also compacts `<cellXfs>` as Excel does on save: cell xfs that
// write identically collapse into the first one's position, groups no
// structure references drop (xf 0 always stays), and the rest keep source
// order. Identity compares the written record, apply flags, parent and xf
// extLst included. A cell inside a merged range takes a separate entry of
// the same content: those entries follow the others, one per group, in the
// order the cells first appear (sheet, row, column).

#ifndef FORMULON_CELL_STYLE_SAVE_PLAN_H_
#define FORMULON_CELL_STYLE_SAVE_PLAN_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "excel_profile.h"
#include "merge_range.h"
#include "phonetic.h"
#include "styles.h"

namespace formulon {

class Workbook;

/// The `<cellXfs>` entries a worksheet references.
struct UsedCellXfs {
  /// `used[c]` is true when cell xf `c` is referenced; indices past the end are unused.
  std::vector<bool> used;
  /// Every cell xf counts as referenced (a sheet the engine cannot scan is present).
  bool all_used = false;
  /// Cell xfs referenced by cells inside a merged range, each once, in the
  /// order such a cell first appears. `used` does not count these references.
  std::vector<std::uint32_t> merged;
};

/// Answers whether a cell lies inside one of a sheet's merged ranges. Queries
/// for one row reuse that row's column spans.
class MergedCellIndex {
 public:
  explicit MergedCellIndex(const std::vector<MergeRange>& merges);
  bool empty() const noexcept { return by_first_row_.empty(); }
  bool covers(std::uint32_t row, std::uint32_t col);
  /// Sorted, possibly overlapping `{first, last}` column spans merged on `row`.
  const std::vector<std::pair<std::uint32_t, std::uint32_t>>& spans(std::uint32_t row);

 private:
  std::vector<MergeRange> by_first_row_;
  std::uint32_t row_ = 0xFFFFFFFFU;
  std::vector<std::pair<std::uint32_t, std::uint32_t>> spans_;
};

/// Collects the cell xfs referenced by every sheet: stored cells, style runs,
/// row styles (`has_style`) and column styles (`has_style`); a stored cell or
/// run column inside a merged range goes to `merged` instead. `all_used` is set
/// by a part the save writes back unread that names cell xfs: an opaque sheet
/// other than a chartsheet, or a retained revision log, macro sheet or dialog
/// sheet part. Indices outside `styles().cell_xfs` are ignored.
UsedCellXfs collect_used_cell_xfs(const Workbook& wb);

/// What a save writes for the named cell styles.
class CellStyleSavePlan {
 public:
  /// `<cellStyles>` in output order. `name` is the saved name and `xf_id` the
  /// output `<cellStyleXfs>` index; the other fields are the written flags.
  const std::vector<CellStyleRecord>& cell_styles() const noexcept { return cell_styles_; }

  /// `<cellStyleXfs>` in output order, as model `cell_style_xfs` indices.
  /// When the model table is empty this is `{0}`, the default xf a writer
  /// synthesizes.
  const std::vector<std::uint32_t>& style_xf_order() const noexcept { return style_xf_order_; }

  /// Output `xfId` of a `<cellXfs>` entry whose model `xf_id` is `old_xf_id`.
  /// 0 when that style xf is not written (only an unused cell xf reaches one).
  /// An index outside the model table is read as 0, as the writers clamp it.
  std::uint32_t cell_xf_parent(std::uint32_t old_xf_id) const noexcept {
    return old_xf_id < parent_of_.size() ? parent_of_[old_xf_id] : parent_of_[0];
  }

  /// `<cellXfs>` in output order, as model `cell_xfs` indices (a group's first
  /// member; a merged-cell entry's first merged member). Empty when the model
  /// table is empty, as the writers then synthesize the default xf.
  const std::vector<std::uint32_t>& cell_xf_order() const noexcept { return cell_xf_order_; }
  /// Output index of model cell xf `model_xf` referenced from outside a merged
  /// range. An index outside the model table is returned unchanged, as the
  /// writers have always written it.
  std::uint32_t cell_xf_index(std::uint32_t model_xf) const noexcept {
    return model_xf < cell_xf_index_.size() ? cell_xf_index_[model_xf] : model_xf;
  }
  /// Output index of model cell xf `model_xf` referenced by a cell inside a
  /// merged range.
  std::uint32_t merged_cell_xf_index(std::uint32_t model_xf) const noexcept {
    return model_xf < merged_cell_xf_index_.size() ? merged_cell_xf_index_[model_xf] : cell_xf_index(model_xf);
  }
  /// True when every cell xf is written at its model index and every
  /// reference keeps its model index.
  bool cell_xfs_identity() const noexcept { return cell_xfs_identity_; }
  /// True when non-default model cell xf `model_xf` is written as xf 0 because it
  /// equals it; a reference to it outside a merged range then reads as xf 0.
  bool collapsed_to_default(std::uint32_t model_xf) const noexcept {
    return model_xf != 0U && cell_xf_index(model_xf) == 0U;
  }
  /// The profile the names, merges and order were decided under.
  ExcelProfile profile() const noexcept { return profile_; }

  /// True when the plan writes the model unchanged: nothing merged, dropped or
  /// synthesized, every saved name equal to its stored name, and both tables
  /// already in the planned order. A table with no named styles is identity.
  bool is_identity() const noexcept { return identity_; }

 private:
  friend CellStyleSavePlan build_cell_style_save_plan(const StylesTable& styles, const UsedCellXfs& used,
                                                      ExcelProfile profile);
  friend CellStyleSavePlan stored_cell_style_save_plan(const StylesTable& styles, ExcelProfile profile);

  std::vector<CellStyleRecord> cell_styles_;
  std::vector<std::uint32_t> style_xf_order_;
  std::vector<std::uint32_t> parent_of_;
  std::vector<std::string> planned_names_;
  std::vector<std::uint32_t> cell_xf_order_;
  std::vector<std::uint32_t> cell_xf_index_;
  std::vector<std::uint32_t> merged_cell_xf_index_;
  bool cell_xfs_identity_ = true;
  ExcelProfile profile_;
  bool identity_ = false;

  friend bool saved_file_names_cell_style(const CellStyleSavePlan& plan, std::string_view name);
};

/// Plans the named cell styles of `styles` for a save under `profile`. A book
/// without a builtin 0 record gains Normal (locale name, style xf 0) as its
/// first style; an empty `cell_style_xfs` is planned as one default xf.
CellStyleSavePlan build_cell_style_save_plan(const StylesTable& styles, const UsedCellXfs& used, ExcelProfile profile);

/// The plan of a save that keeps the stored styles part: every record, name
/// and index as loaded. A retained xlsb styles.bin that passes through is the
/// only style table in the file, so the parts written beside it follow this.
CellStyleSavePlan stored_cell_style_save_plan(const StylesTable& styles, ExcelProfile profile);

/// Whether a table cell-style attribute naming `name` is kept on save: `name`
/// equals, under the profile's compare collation, the saved name of a planned
/// style, written or dropped as unused. A built-in matches only under its
/// locale name. This describes the saved file and is not a name lookup.
bool saved_file_names_cell_style(const CellStyleSavePlan& plan, std::string_view name);

/// The `<fonts>` a save writes, and the `<phoneticPr>` of the sheets that state
/// none. A font linked to the theme (`scheme` minor or major) is written with
/// the theme's face for the script of the profile's UI locale, which is what
/// Excel stores for it on save; the model keeps the stored name. A font without
/// a scheme, or one the theme names no face for, is written as stored.
///
/// Under a locale whose new workbooks carry a phonetic font (ja, zh_CN, zh_TW,
/// ko), a book with a sheet lacking `<phoneticPr>` gets that locale's phonetic
/// font: font 0's written face and scheme at the locale's phonetic size,
/// family 3 and charset. An equal font already in the table is reused, else the
/// font is appended last; every such sheet then names it. A font 0 without a
/// scheme gets none, as Excel's choice of face there follows the host's
/// installed fonts.
class FontSavePlan {
 public:
  /// The fonts in output order: the model's, then any appended phonetic font.
  const std::vector<FontRecord>& fonts() const noexcept { return fonts_; }

  /// The `<phoneticPr>` a sheet without its own is written with; empty when
  /// such sheets are written without one.
  const std::optional<PhoneticProperties>& absent_sheet_phonetic() const noexcept { return absent_sheet_phonetic_; }

  /// True when every font is written as the model holds it and none is added.
  bool is_identity() const noexcept { return identity_; }

 private:
  friend FontSavePlan build_font_save_plan(const Workbook& wb);

  std::vector<FontRecord> fonts_;
  std::optional<PhoneticProperties> absent_sheet_phonetic_;
  bool identity_ = true;
};

/// The saved index of model cell xf `xf_index`, for a cell inside a merged
/// range when `merged`; unchanged without a plan.
inline std::uint32_t saved_cell_xf(const CellStyleSavePlan* style_plan, std::uint32_t xf_index, bool merged = false) {
  if (style_plan == nullptr) {
    return xf_index;
  }
  return merged ? style_plan->merged_cell_xf_index(xf_index) : style_plan->cell_xf_index(xf_index);
}

/// Plans the fonts of `wb` for a save under its profile, with its theme (the
/// default theme when the book has no theme part).
FontSavePlan build_font_save_plan(const Workbook& wb);

}  // namespace formulon

#endif  // FORMULON_CELL_STYLE_SAVE_PLAN_H_
