//
// Sheet-attached metadata records: structural-edit mapping, hyperlinks, comments, data validations and spill regions.

#ifndef FORMULON_SHEET_METADATA_H_
#define FORMULON_SHEET_METADATA_H_

#include <array>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "data_table.h"
#include "merge_range.h"
#include "utils/band.h"
#include "utils/runtime_id.h"
#include "value.h"

namespace formulon {

namespace pivot {
class PivotTable;
}  // namespace pivot

/// Direction of an Insert Cells / Delete Cells edit: insert takes `kDown` or
/// `kRight`, delete takes `kUp` or `kLeft`.
enum class CellShift : std::uint8_t { kDown, kRight, kUp, kLeft };

/// One row/column structural edit, as the coordinate mapping every
/// sheet-attached structure derives from.
///
/// `index` is the 0-based band the edit starts at and `count` its width. On
/// an insert, a coordinate `>= index` moves forward by `count`. On a delete,
/// `[index, index + count)` is the deleted band — a coordinate inside it is
/// dropped or clamped, and a coordinate past it moves back by `count`.
///
/// The public `Sheet` edit methods each build or take one of these and hand it
/// to `Sheet::finalize_structural_edit_locked`, which invokes the single
/// place that enumerates the modelled structures. A structure that is not
/// listed there does not move, and that failure is silent in a user's file —
/// `tests/integration/structural_edit_matrix_test.cpp` asserts each one's
/// post-edit coordinate.
struct StructuralEdit {
  std::uint32_t index = 0;
  std::uint32_t count = 0;
  bool is_delete = false;
  bool row_axis = false;
  /// Perpendicular extent the edit is confined to (Insert / Delete Cells);
  /// absent for a whole row/column edit.
  std::optional<Band> band{};
};

/// How a row/column edit meets a structure that moves whole or goes whole (a
/// data table's block, a legacy array block, a PivotTable), given its extent
/// `[first, last]` along the edit's axis.
enum class StructureEditCut : std::uint8_t {
  kMisses,     ///< The edit lies outside the extent; the structure shifts whole or stays.
  kRemoves,    ///< A delete spans the whole extent; the structure goes.
  kSplits,     ///< The edit cuts into the extent.
  kPushesOff,  ///< An insert would push the far edge of the extent off the grid.
};

/// Classifies `edit` against the extent `[first, last]`; see `StructureEditCut`.
StructureEditCut structure_edit_cut(std::uint32_t first, std::uint32_t last, const StructuralEdit& edit) noexcept;

/// `structure_edit_cut` for a data table's block: its result range plus the
/// header row above it and the header column left of it.
StructureEditCut data_table_structure_cut(const DataTable& table, const StructuralEdit& edit) noexcept;

/// A data table's block as cells: its result range plus the header row above
/// it and the header column left of it.
MergeRange data_table_block(const DataTable& table) noexcept;

/// A PivotTable's location range as cells; a zero span counts as the anchor
/// line alone.
MergeRange pivot_location(const pivot::PivotTable& pivot) noexcept;

/// `structure_edit_cut` for a PivotTable's location range; a zero span counts
/// as the anchor line alone.
StructureEditCut pivot_structure_cut(const pivot::PivotTable& pivot, const StructuralEdit& edit) noexcept;

/// Moves the ranges of one sqref (a conditional format's, or a retained x14
/// copy of one) for a row/column edit: `count` rows or columns inserted or
/// deleted at 0-based `index`. Ranges the edit deletes are removed, and a
/// delete joins the ranges it brings edge to edge.
void shift_sqref_ranges(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                        bool row_axis);

/// Moves the ranges of an sqref that an edit cuts rather than stretches (same
/// arguments as `shift_sqref_ranges`): a range the edit falls inside splits
/// at `index`, so an insert leaves its new band uncovered and a delete
/// leaves the two sides as separate ranges. Excel moves `<ignoredError>`
/// ranges this way.
void cut_sqref_ranges(std::vector<MergeRange>& ranges, std::uint32_t index, std::uint32_t count, bool is_delete,
                      bool row_axis);

/// Where a rectangle sits against a band edit's moving region: the band x
/// from the edit's edge to the grid's end (the deleted cells included).
enum class BandFit : std::uint8_t {
  kOutside,    ///< Misses the moving region; it stays.
  kInside,     ///< Within the band and reaching the moving region; it follows the edit.
  kStraddles,  ///< Reaches the moving region across a band edge.
};

/// Classifies `rect` against `edit`, which must carry a band.
BandFit band_fit(const MergeRange& rect, const StructuralEdit& edit) noexcept;

/// `shift_sqref_ranges` for `edit`, band or not. Under a band, a range inside
/// it follows the span rules, one straddling the moving region splits into
/// its moved and unmoved parts, and nothing is joined.
void shift_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit);

/// `shift_sqref_ranges` for a conditional format's sqref, legacy or x14: a
/// band delete also joins the ranges it brings together into the one
/// rectangle they now tile, unless they tiled it before the edit.
void shift_cf_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit);

/// Moves retained ranges that follow references (protected ranges, sort and
/// web-publish ranges, watched cells): `shift_sqref_ranges` for a whole
/// row/column edit; under a band only a range inside it moves.
void shift_reference_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit);

/// `cut_sqref_ranges` for `edit`; under a band only a range inside it is cut.
void cut_sqref_ranges(std::vector<MergeRange>& ranges, const StructuralEdit& edit);

/// `shift_auto_filter` for `edit`; under a band only a filter inside it moves.
bool shift_auto_filter(AutoFilter& filter, const StructuralEdit& edit, bool header_delete_removes);

/// Moves an AutoFilter for a row/column edit (same arguments as
/// `shift_sqref_ranges`): its range and sort ranges follow the span rules, a
/// sort state goes with the last of its conditions,
/// a column delete removes the deleted columns' criteria and renumbers the
/// later `colId`s down, and an insert strictly inside the range renumbers the
/// `colId`s at or after it up. Returns false when the filter must be removed:
/// its range was consumed, or `header_delete_removes` is set and the header
/// row was deleted (a sheet AutoFilter; Excel cannot delete a table's).
/// An opaque filter is left untouched.
bool shift_auto_filter(AutoFilter& filter, std::uint32_t index, std::uint32_t count, bool is_delete, bool row_axis,
                       bool header_delete_removes);

/// One `<hyperlink>` entry attached to a sheet. The numeric rectangle is the
/// sole source of truth; the OOXML writer regenerates its `ref` attribute from
/// these four 0-based inclusive endpoints. The OOXML wire form is
///
///     <hyperlink ref="A1" r:id="rId3" tooltip="..." display="..."/>
///
/// where `r:id` resolves to a `Target` in the sheet's
/// `_rels/<sheet>.xml.rels`. Hyperlinks may be external (URLs / mailto)
/// or internal (e.g. `#Sheet2!A1`); for internal links Excel emits the
/// `location` attribute on the `<hyperlink>` itself rather than going
/// through rels. We capture both shapes in a single struct so the
/// writer can reproduce either.
///
/// Lifetime: `target` / `display` / `tooltip` / `location` are owned
/// `std::string`s. `rid` is the relationship id observed at read time;
/// the writer reuses it verbatim for round-trip stability and assigns a
/// fresh `rId<N>` when the field is empty (newly added entries).
struct Hyperlink {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
  std::string target;    ///< External URL / mailto / file path. Empty for purely internal links.
  std::string location;  ///< `Sheet2!A1` style internal target, or empty.
  std::string display;   ///< Optional `display="..."` attribute.
  std::string tooltip;   ///< Optional `tooltip="..."` attribute.
  std::string rid;       ///< Reader populates from sheet rels; empty for fresh entries.

  /// Whether the link's cells include `(r, c)`.
  bool covers(std::uint32_t r, std::uint32_t c) const noexcept {
    return row <= r && r <= last_row && col <= c && c <= last_col;
  }
  /// Whether `other` covers exactly the same cells, the identity of a link.
  bool same_range(const Hyperlink& other) const noexcept {
    return row == other.row && col == other.col && last_row == other.last_row && last_col == other.last_col;
  }
  /// Whether the link shares a cell with the inclusive rectangle given.
  bool intersects(std::uint32_t first_r, std::uint32_t first_c, std::uint32_t last_r,
                  std::uint32_t last_c) const noexcept {
    return row <= last_r && first_r <= last_row && col <= last_c && first_c <= last_col;
  }
};

/// One per-cell text comment (`<comment ref="A1" authorId="N"><text>...</text></comment>`).
///
/// The struct is intentionally narrow: a single plain-text payload, no
/// rich-text runs. Reads of multi-run rich text concatenate the runs
/// into `text`; writes always emit the plain form. Authors are
/// preserved verbatim so the workbook-wide author list round-trips.
/// `uid` is the comment's `xr:uid` revision id, empty when absent.
struct CellComment {
  CellComment() = default;
  CellComment(std::uint32_t row_in, std::uint32_t col_in, std::string author_in, std::string text_in,
              std::string uid_in = {})
      : row(row_in), col(col_in), author(std::move(author_in)), text(std::move(text_in)), uid(std::move(uid_in)) {}

  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string author;
  std::string text;
  std::string uid;
  /// The note box is shown without hovering (`visibility:visible` in the VML).
  bool visible = false;
  /// The VML `x:Anchor` box when the source carried one: left column and
  /// offset, top row and offset, right column and offset, bottom row and
  /// offset. Columns and rows are relative to the note's cell so the box
  /// follows the note through row and column edits.
  bool has_box = false;
  std::array<std::int32_t, 8> box{};
  /// The source `<v:shape>` element text, empty for a note authored here. A
  /// regenerated VML drawing re-emits it with only the anchor, visibility and
  /// cell patched.
  std::string vml_shape;
};

/// What a comment VML drawing says about one note shape, compared to decide
/// whether the source VML can be re-emitted as it was.
struct CommentShape {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  bool visible = false;
  bool has_box = false;
  std::array<std::int32_t, 8> box{};
  std::string xml;  ///< `CellComment::vml_shape` of the note, empty for none.

  friend bool operator==(const CommentShape& lhs, const CommentShape& rhs) noexcept {
    return lhs.row == rhs.row && lhs.col == rhs.col && lhs.visible == rhs.visible && lhs.has_box == rhs.has_box &&
           lhs.box == rhs.box && lhs.xml == rhs.xml;
  }
};

/// One `<dataValidation>` block: a list of cell ranges that share a
/// validation rule, plus the rule itself. Storage is plain POD; the
/// engine does not yet evaluate the rules — the data is captured for
/// round-trip and surfacing through the binding APIs.
///
/// Field semantics (matches OOXML `dataValidations.xsd`):
///   * `type`         — 0 none, 1 whole, 2 decimal, 3 list, 4 date,
///                      5 time, 6 textLength, 7 custom.
///   * `op`           — 0 between, 1 notBetween, 2 equal, 3 notEqual,
///                      4 greaterThan, 5 lessThan, 6 greaterThanOrEqual,
///                      7 lessThanOrEqual.
///   * `error_style`  — 0 stop, 1 warning, 2 information.
///   * `formula1`/`formula2` — raw OOXML formula text without leading
///                      `=`; the evaluator does not consume these.
struct DataValidation {
  std::vector<MergeRange> ranges;
  std::uint8_t type = 0;
  std::uint8_t op = 0;
  std::uint8_t error_style = 0;
  bool allow_blank = true;
  bool show_input_message = false;
  bool show_error_message = false;
  // Whether the in-cell dropdown arrow is shown for a `list` validation.
  // Default matches Excel's default (arrow shown). Note the OOXML
  // `showDropDown` attribute has inverted semantics per ECMA-376: its
  // presence with value "1" SUPPRESSES the arrow. This field holds the
  // user-facing meaning (arrow shown/hidden); the reader/writer invert it
  // when translating to/from the raw XML attribute.
  bool show_dropdown = true;
  std::string formula1;
  std::string formula2;
  std::string error_title;
  std::string error_message;
  std::string prompt_title;
  std::string prompt_message;
  /// Attributes not represented by the model (`imeMode`, ...) retained as
  /// escaped ` name="value"` pairs.
  std::string extra_attrs;
  /// Session identity of the rule, never saved; assigned by
  /// `Workbook::validation_runtime_id` (see `utils/runtime_id.h`).
  mutable RuntimeId runtime_id;
};

/// A registered spill region produced by a dynamic-array formula.
///
/// `cells` is row-major (size = `rows * cols`) and holds heap-owned `Value`
/// copies. Any `Text` cell's payload is interned in `owned_strings` so the
/// `SpillRegion`'s lifetime is independent of the per-evaluation arena that
/// produced the original `ArrayValue`. The order of strings in
/// `owned_strings` is insertion order (left-to-right, top-to-bottom across
/// `cells`), which is purely an implementation detail; consumers must not
/// rely on it.
struct SpillRegion {
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  std::vector<Value> cells;
  // Backing store for Text values: a `Text` entry in `cells` is a string_view
  // into one of these strings. Kept as a separate vector so the `cells`
  // vector itself remains a dense array of `Value` (which is trivially
  // copyable).
  std::vector<std::string> owned_strings;
};

/// The geometry of a dynamic-array spill footprint. It stores no values and
/// is safe to copy while the sheet spill mutex is held. The same value type
/// is used for committed and blocked spill snapshots.
struct SpillFootprint {
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
};

/// What refused a spill: the first obstructed cell in row-major order is in a
/// merge (`kMerged`) or holds anything else (`kOccupied`), or the rectangle
/// leaves the grid (`kBeyondEdge`), which wins over any in-grid obstruction.
enum class SpillCause : std::uint8_t { kOccupied, kMerged, kBeyondEdge };

/// A refused spill: the rectangle the anchor needs, which may run past the
/// grid when `cause` is `kBeyondEdge`, and why it was refused.
struct BlockedSpillFootprint {
  std::uint32_t anchor_row = 0;
  std::uint32_t anchor_col = 0;
  std::uint32_t rows = 0;
  std::uint32_t cols = 0;
  SpillCause cause = SpillCause::kOccupied;
};

/// The rich detail of an error cell that a spill footprint or a blocked pivot
/// does not already carry: a measured `#CALC!` / `#VALUE!` origin, or an error
/// that arrived from another cell, shaped for `#VALUE!` by the formula's root.
struct CellErrorDetail {
  enum class Shape : std::uint8_t { kSubType, kIftab, kPropagated, kPropagatedPtg, kPropagatedIftab };
  ErrorCode code = ErrorCode::Calc;  // Calc or Value; Spill only with kPropagated
  Shape shape = Shape::kSubType;
  std::uint32_t sub_type = 0;  // kSubType, kIftab
  std::uint32_t iftab = 0;     // kIftab, kPropagatedIftab
  std::uint32_t ptg = 0;       // kPropagatedPtg read from a file: the root operator's Ptg id
  // kPropagatedPtg from the engine: 1 + the root operator's parser::BinOp; the writer maps it to the Ptg id.
  std::uint8_t root_binop = 0;
  // kPropagatedIftab from the engine: the root function's canonical name, in static storage.
  std::string_view root_function{};
  std::uint8_t root_argpos = 0;  // 1-based root argument that carried the error; 0 when not determined
  bool elements = false;         // also applies to the error elements of the spill anchored here
};

}  // namespace formulon

#endif  // FORMULON_SHEET_METADATA_H_
