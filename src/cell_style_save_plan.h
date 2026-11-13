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

#ifndef FORMULON_CELL_STYLE_SAVE_PLAN_H_
#define FORMULON_CELL_STYLE_SAVE_PLAN_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "excel_profile.h"
#include "styles.h"

namespace formulon {

class Workbook;

/// The `<cellXfs>` entries a worksheet references.
struct UsedCellXfs {
  /// `used[c]` is true when cell xf `c` is referenced; indices past the end are unused.
  std::vector<bool> used;
  /// Every cell xf counts as referenced (a sheet the engine cannot scan is present).
  bool all_used = false;
};

/// Collects the cell xfs referenced by every sheet: stored cells, style runs,
/// row styles (`has_style`) and column styles (`has_style`). An opaque sheet
/// sets `all_used`. Indices outside `styles().cell_xfs` are ignored.
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

  /// The profile the names, merges and order were decided under.
  ExcelProfile profile() const noexcept { return profile_; }

  /// True when the plan writes the model unchanged: nothing merged, dropped or
  /// synthesized, every saved name equal to its stored name, and both tables
  /// already in the planned order. A table with no named styles is identity.
  bool is_identity() const noexcept { return identity_; }

 private:
  friend CellStyleSavePlan build_cell_style_save_plan(const StylesTable& styles, const UsedCellXfs& used,
                                                      ExcelProfile profile);

  std::vector<CellStyleRecord> cell_styles_;
  std::vector<std::uint32_t> style_xf_order_;
  std::vector<std::uint32_t> parent_of_;
  std::vector<std::string> planned_names_;
  ExcelProfile profile_;
  bool identity_ = false;

  friend bool saved_file_names_cell_style(const CellStyleSavePlan& plan, std::string_view name);
};

/// Plans the named cell styles of `styles` for a save under `profile`. A book
/// without a builtin 0 record gains Normal (locale name, style xf 0) as its
/// first style; an empty `cell_style_xfs` is planned as one default xf.
CellStyleSavePlan build_cell_style_save_plan(const StylesTable& styles, const UsedCellXfs& used, ExcelProfile profile);

/// Whether a table cell-style attribute naming `name` is kept on save: `name`
/// equals, under the profile's compare collation, the saved name of a planned
/// style, written or dropped as unused. A built-in matches only under its
/// locale name. This describes the saved file and is not a name lookup.
bool saved_file_names_cell_style(const CellStyleSavePlan& plan, std::string_view name);

}  // namespace formulon

#endif  // FORMULON_CELL_STYLE_SAVE_PLAN_H_
