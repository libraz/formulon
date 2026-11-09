//
// `PivotTable` is the workbook-level pivot definition: identity, cache
// binding, field configuration, layout, anchor location, transient slicer
// state, and the most recent evaluation result. One pivot table is owned
// by exactly one sheet (the sheet supplies the sheet identity); it points
// at a workbook-owned `PivotCache` by id.

#ifndef FORMULON_PIVOT_PIVOT_TABLE_H_
#define FORMULON_PIVOT_PIVOT_TABLE_H_

#include <cstddef>
#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"
#include "utils/runtime_id.h"

namespace formulon::pivot {

/// `<pivotTableStyleInfo>`: the table style and its banding / header toggles.
/// An empty `name` means the element carries no style name.
struct PivotStyleInfo {
  std::string name;
  bool show_row_headers = false;
  bool show_col_headers = false;
  bool show_row_stripes = false;
  bool show_col_stripes = false;
  bool show_last_column = false;
};

/// A rectangle of sheet cells, inclusive on all four sides, zero-based.
struct PivotRect {
  std::uint32_t row0 = 0;
  std::uint32_t col0 = 0;
  std::uint32_t row1 = 0;
  std::uint32_t col1 = 0;

  bool operator==(const PivotRect& other) const {
    return row0 == other.row0 && col0 == other.col0 && row1 == other.row1 && col1 == other.col1;
  }
};

/// One pivot table definition + most recent evaluation result.
///
/// Move-only; `clone()` is the one copy. Designed to be held by the owning
/// `Sheet` via `std::unique_ptr<PivotTable>` so the address is stable for as
/// long as the sheet retains the pivot.
class PivotTable {
 public:
  PivotTable() = default;

  PivotTable& operator=(const PivotTable&) = delete;
  PivotTable(PivotTable&&) = default;
  PivotTable& operator=(PivotTable&&) = default;

  /// An independent copy of the definition with its own result mutex and no
  /// cached result; the cache id is kept, so the copy shares the cache.
  std::unique_ptr<PivotTable> clone() const {
    std::unique_ptr<PivotTable> copy(new PivotTable(*this));
    copy->last_result_mutex_ = std::make_shared<std::mutex>();
    copy->last_result_.reset();
    return copy;
  }

  // Identity / cache binding -------------------------------------------------

  const std::string& name() const { return name_; }
  void set_name(std::string n) { name_ = std::move(n); }

  std::uint32_t pivot_cache_id() const { return pivot_cache_id_; }
  void set_pivot_cache_id(std::uint32_t id) { pivot_cache_id_ = id; }

  /// Session identity the C API assigns on first request; never saved. A
  /// `clone()` starts at zero (see `utils/runtime_id.h`).
  RuntimeId& runtime_id() const { return runtime_id_; }

  /// The `dataCaption` attribute on `<pivotTableDefinition>` — the header
  /// shown over the Values area (e.g. "Values" / "値"). ECMA-376 marks it
  /// required, so the writer always emits it; the reader captures the
  /// authored value and defaults to "Values" when absent so a file built
  /// from scratch still round-trips a schema-valid definition.
  const std::string& data_caption() const { return data_caption_; }
  void set_data_caption(std::string caption) { data_caption_ = std::move(caption); }

  // Field configuration ------------------------------------------------------

  const std::vector<PivotField>& fields() const { return fields_; }
  std::vector<PivotField>& mutable_fields() { return fields_; }

  /// Data-field entries from `<dataFields>`. Keyed by display name for
  /// GETPIVOTDATA; one source `PivotField` may have multiple data-field
  /// entries (e.g. Sum + Average of the same column).
  const std::vector<PivotDataField>& data_fields() const { return data_fields_; }
  std::vector<PivotDataField>& mutable_data_fields() { return data_fields_; }

  /// Document-order indices into `fields()` of the row-axis fields,
  /// captured from `<rowFields>`. Used by the pivot evaluator to walk
  /// the row hierarchy.
  const std::vector<std::uint32_t>& row_field_order() const { return row_field_order_; }
  std::vector<std::uint32_t>& mutable_row_field_order() { return row_field_order_; }

  /// Document-order indices into `fields()` of the column-axis fields,
  /// captured from `<colFields>`.
  const std::vector<std::uint32_t>& col_field_order() const { return col_field_order_; }
  std::vector<std::uint32_t>& mutable_col_field_order() { return col_field_order_; }

  /// Document position of the `<field x="-2"/>` Values pseudo-field
  /// marker (`pivot::kValuesFieldPosition`) within `<rowFields>`, or
  /// `nullopt` when the row axis carries none. Counted over the
  /// original `<rowFields>` children including the marker's own slot,
  /// so the writer can re-interleave it among `row_field_order()` at
  /// the same position. Kept apart from `row_field_order()`, which
  /// holds only real indices into `fields()`.
  std::optional<std::size_t> row_values_position() const { return row_values_position_; }
  void set_row_values_position(std::optional<std::size_t> position) { row_values_position_ = position; }

  /// Same as `row_values_position()`, for `<colFields>`.
  std::optional<std::size_t> col_values_position() const { return col_values_position_; }
  void set_col_values_position(std::optional<std::size_t> position) { col_values_position_ = position; }

  // Page (report filter) axis -----------------------------------------------
  //
  // The model owns the page axis: the writer emits `<pageFields>` from
  // `page_field_order()` and each field's entry here, and nothing of the
  // element is passed through. An entry adds what `<pivotFields>` cannot
  // record: the report order and the single selected item (`item_index`).
  //
  // The list is kept in step with the axis: `set_field_axis` appends a field
  // joining the page axis and drops one leaving it, and `page_field_order()`
  // skips an entry whose field is no longer on the axis and appends a page
  // field without an entry in `fields()` order, so a table whose `axis` was
  // set directly still reads consistently.
  //
  // A page field's hidden items prune records through
  // `PreparedRecordFilter`, which keys off `PivotField::items` on every axis
  // alike; a single selected item filters on that item too.
  const std::vector<PivotPageField>& page_fields() const { return page_fields_; }
  std::vector<PivotPageField>& mutable_page_fields() { return page_fields_; }

  /// The first page-field entry for `field_index`, or null when the field is
  /// not on the page axis or has no entry.
  const PivotPageField* find_page_field(std::uint32_t field_index) const {
    if (field_index >= fields_.size() || fields_[field_index].axis != PivotAxis::Page) {
      return nullptr;
    }
    for (const PivotPageField& page : page_fields_) {
      if (page.field_index == field_index) {
        return &page;
      }
    }
    return nullptr;
  }

  /// Indices into `fields()` of the page-axis fields, in report order: the
  /// entries of `page_fields()` whose field is on the page axis, then every
  /// other page-axis field in `fields()` order (Excel's own fallback for a
  /// page field without a `<pageField>`).
  std::vector<std::uint32_t> page_field_order() const {
    std::vector<std::uint32_t> order;
    std::vector<bool> listed(fields_.size(), false);
    for (const PivotPageField& page : page_fields_) {
      if (page.field_index < fields_.size() && fields_[page.field_index].axis == PivotAxis::Page &&
          !listed[page.field_index]) {
        listed[page.field_index] = true;
        order.push_back(page.field_index);
      }
    }
    for (std::size_t i = 0; i < fields_.size(); ++i) {
      if (fields_[i].axis == PivotAxis::Page && !listed[i]) {
        order.push_back(static_cast<std::uint32_t>(i));
      }
    }
    return order;
  }

  /// Moves field `field_index` to `axis`, keeping the page list in step: a
  /// field joining the page axis is appended to it, one leaving it is
  /// removed. Row / column orders and data fields are the caller's. An
  /// out-of-range index is ignored.
  void set_field_axis(std::uint32_t field_index, PivotAxis axis) {
    if (field_index >= fields_.size()) {
      return;
    }
    const bool was_page = fields_[field_index].axis == PivotAxis::Page;
    fields_[field_index].axis = axis;
    if (was_page == (axis == PivotAxis::Page)) {
      return;
    }
    std::vector<PivotPageField> kept;
    kept.reserve(page_fields_.size() + 1U);
    for (PivotPageField& page : page_fields_) {
      if (page.field_index != field_index) {
        kept.push_back(std::move(page));
      }
    }
    if (axis == PivotAxis::Page) {
      PivotPageField added;
      added.field_index = field_index;
      kept.push_back(std::move(added));
    }
    page_fields_ = std::move(kept);
  }

  // Layout / location --------------------------------------------------------

  PivotLayout layout() const { return layout_; }

  /// Sets the report form of the table and of every field, as Excel's report
  /// layout command does: a per-field or table-level form read from a file
  /// (kept in `passthrough_attrs`) is dropped in favour of the derived one.
  void set_layout(PivotLayout l) {
    layout_ = l;
    erase_form_attrs(passthrough_attrs_);
    for (PivotField& field : fields_) {
      erase_form_attrs(field.passthrough_attrs);
    }
  }

  std::uint32_t anchor_row() const { return anchor_row_; }
  std::uint32_t anchor_col() const { return anchor_col_; }
  std::uint32_t span_rows() const { return span_rows_; }
  std::uint32_t span_cols() const { return span_cols_; }

  void set_anchor(std::uint32_t row, std::uint32_t col, std::uint32_t rows, std::uint32_t cols) {
    anchor_row_ = row;
    anchor_col_ = col;
    span_rows_ = rows;
    span_cols_ = cols;
    span_read_ = false;
  }

  /// True when the span came from a `<location ref>` the reader decoded,
  /// rather than from a model assembled in memory.
  ///
  /// The saved `ref` has to describe the grid Excel will actually draw.
  /// Nothing in this model keeps the span in step with the fields, so a
  /// table assembled through the C API still carries the placeholder span
  /// installed at creation however many fields are added afterwards, and a
  /// `ref` smaller than the rendered report makes Excel terminate when it
  /// refreshes. The writer therefore projects the real extent for a table
  /// whose span is not authored (see `write_pivot_table_definition`).
  ///
  /// An authored span is re-emitted exactly as it was read. Re-deriving it
  /// would rewrite a `ref` Excel itself wrote, on the strength of our own
  /// layout projection — the round-trip fidelity gate exists to catch
  /// precisely that.
  /// Cleared again as soon as anything about the table changes: the span
  /// Excel wrote described the report as it stood, and a field, an order,
  /// a filter or a move all resize it. Position stays the caller's to set;
  /// the extent goes back to being projected.
  bool has_authored_span() const { return has_authored_span_; }
  void mark_span_authored() {
    has_authored_span_ = true;
    span_read_ = true;
  }
  void clear_span_authored() { has_authored_span_ = false; }

  /// True while the span is still the one read from a file, authored or gone
  /// stale through an edit; `set_anchor` makes it the caller's. A stale read
  /// span is replaced by the projection, whereas a caller's span is a floor.
  bool span_read() const { return span_read_; }

  /// True once the definition was edited after load, so the `<rowItems>` /
  /// `<colItems>` read with it no longer describe the report. The writer then
  /// drops them and asks Excel to rebuild the report on open. Unlike
  /// `has_authored_span`, a change of evaluation context leaves it alone.
  bool layout_dirty() const { return layout_dirty_; }
  void mark_layout_dirty() { layout_dirty_ = true; }

  // `<location>` required / commonly-present attributes ----------------------
  //
  // ECMA-376 marks `firstHeaderRow`, `firstDataRow`, and `firstDataCol` as
  // required on `<location>`; `rowPageCount` / `colPageCount` are optional
  // but commonly present. The `ref` attribute is modelled separately as the
  // anchor/spans above. These offsets are relative to the top-left of the
  // pivot's `ref` range. They are captured verbatim from the source so a
  // read -> write round trip re-emits a schema-valid `<location>` (dropping
  // the required attributes makes Excel flag the file for repair). Stored as
  // `optional` so the writer can preserve authored values, including zero,
  // and independently default an absent required value to 1.

  std::optional<std::uint32_t> location_first_header_row() const { return location_first_header_row_; }
  std::optional<std::uint32_t> location_first_data_row() const { return location_first_data_row_; }
  std::optional<std::uint32_t> location_first_data_col() const { return location_first_data_col_; }
  std::optional<std::uint32_t> location_row_page_count() const { return location_row_page_count_; }
  std::optional<std::uint32_t> location_col_page_count() const { return location_col_page_count_; }

  void set_location_attributes(std::optional<std::uint32_t> first_header_row,
                               std::optional<std::uint32_t> first_data_row, std::optional<std::uint32_t> first_data_col,
                               std::optional<std::uint32_t> row_page_count,
                               std::optional<std::uint32_t> col_page_count) {
    location_first_header_row_ = first_header_row;
    location_first_data_row_ = first_data_row;
    location_first_data_col_ = first_data_col;
    location_row_page_count_ = row_page_count;
    location_col_page_count_ = col_page_count;
  }

  /// True iff `(row, col)` is inside the pivot's bounds.
  bool contains(std::uint32_t row, std::uint32_t col) const noexcept {
    return row >= anchor_row_ && row < anchor_row_ + span_rows_ && col >= anchor_col_ && col < anchor_col_ + span_cols_;
  }

  // Label / value / date filters (`<filters>`) ------------------------------
  //
  // Kept separate from `PivotField::items` (which is the manual item
  // filter state).
  //
  // `<filters>` is owned by two lists. `active_filters()` holds every entry
  // the `PivotFilter` model expresses: the ones added through the API and
  // the ones the reader decoded (each carrying its `PivotFilterSource`).
  // `retained_filters()` holds the loaded entries it cannot express,
  // verbatim. The writer serialises both, interleaved back into document
  // order, so an API filter survives a save + reload and a loaded block is
  // re-emitted byte-for-byte while nothing in it changed.
  const std::vector<PivotFilter>& active_filters() const { return active_filters_; }
  std::vector<PivotFilter>& mutable_active_filters() { return active_filters_; }

  const std::vector<RetainedPivotFilter>& retained_filters() const { return retained_filters_; }
  std::vector<RetainedPivotFilter>& mutable_retained_filters() { return retained_filters_; }

  /// Removes every active filter `remove` returns true for (it may also edit
  /// the ones it keeps), re-anchoring each retained entry after the same
  /// surviving active entries so document order holds.
  /// Appending needs no such care and goes through `mutable_active_filters`.
  template <typename Remove>
  void erase_active_filters_if(Remove remove) {
    const std::size_t n = active_filters_.size();
    std::vector<std::size_t> kept_before(n + 1U, 0U);
    std::vector<PivotFilter> kept;
    for (std::size_t i = 0; i < n; ++i) {
      kept_before[i] = kept.size();
      if (!remove(active_filters_[i])) {
        kept.push_back(std::move(active_filters_[i]));
      }
    }
    kept_before[n] = kept.size();
    active_filters_ = std::move(kept);
    for (RetainedPivotFilter& entry : retained_filters_) {
      entry.active_before = kept_before[entry.active_before < n ? entry.active_before : n];
    }
  }

  // Most-recent evaluation result -------------------------------------------
  //
  // `last_result_` is logical-const memoisation: GETPIVOTDATA observes a
  // `PivotTable` through a `const Workbook&` and may publish a result on
  // demand. Return a shared snapshot, never a reference into the cache.
  // Parallel formula cells can therefore keep reading an older result while
  // another cell publishes a replacement without a dangling reference.
  std::shared_ptr<const PivotResult> last_result() const {
    std::lock_guard<std::mutex> lock(*last_result_mutex_);
    return last_result_;
  }

  void set_last_result(PivotResult result) const {
    auto snapshot = std::make_shared<const PivotResult>(std::move(result));
    std::lock_guard<std::mutex> lock(*last_result_mutex_);
    last_result_ = std::move(snapshot);
  }

  void clear_last_result() const {
    std::lock_guard<std::mutex> lock(*last_result_mutex_);
    last_result_.reset();
  }

  // OOXML round-trip passthrough --------------------------------------------
  //
  // OOXML pivot definitions can carry several elements this version does
  // not model structurally — `<rowItems>`, `<colItems>`,
  // `<calculatedFields>`, `<calculatedItems>`, `<chartFormats>`,
  // `<formats>`, etc. The reader captures them as
  // concatenated raw XML so the writer can emit them back verbatim,
  // keeping a round trip stable for files that depend on those features.
  //
  // `CT_pivotTableDefinition` mandates a strict child order, so a single
  // trailing buffer would re-emit `<rowItems>`/`<colItems>` after
  // `<dataFields>` and make Excel flag the file for repair. Those two are
  // the only unmodelled elements the schema places *before* `<dataFields>`,
  // so the reader bins them by name into the two pre-dataFields slots
  // (rowItems after `<rowFields>`, colItems after `<colFields>`). `<pivotTableStyleInfo>` is modelled
  // (`style_info`) and splits the post-dataFields elements: formats,
  // conditionalFormats, chartFormats and pivotHierarchies go to
  // `raw_passthrough_before_style_info_`; everything else — hierarchy
  // usage, extLst, ... — goes to the tail bin `raw_passthrough_xml_`, which
  // the writer flushes after the modelled `<filters>`. It flushes each bin at the matching
  // slot so schema order is preserved. Mutating the structured state does not invalidate these
  // buffers; consumers that need exact bit parity should regenerate or
  // strip them.
  const std::string& raw_passthrough_xml() const { return raw_passthrough_xml_; }
  std::string& mutable_raw_passthrough_xml() { return raw_passthrough_xml_; }

  const std::string& raw_passthrough_before_style_info() const { return raw_passthrough_before_style_info_; }
  std::string& mutable_raw_passthrough_before_style_info() { return raw_passthrough_before_style_info_; }

  const std::string& raw_passthrough_after_row_fields() const { return raw_passthrough_after_row_fields_; }
  std::string& mutable_raw_passthrough_after_row_fields() { return raw_passthrough_after_row_fields_; }

  /// `<colItems>`, which the writer drops from a `layout_dirty()` table.
  const std::string& raw_col_items() const { return raw_col_items_; }
  std::string& mutable_raw_col_items() { return raw_col_items_; }

  /// `<pivotTableStyleInfo>`; nullopt when the definition has none.
  const std::optional<PivotStyleInfo>& style_info() const { return style_info_; }
  void set_style_info(PivotStyleInfo info) { style_info_ = std::move(info); }

  // Unmodelled `<pivotTableDefinition>` root attributes (`updatedVersion`,
  // `createdVersion`, `itemPrintTitles`, `indent`, ...), captured as
  // `(name, value)` pairs so the writer re-emits them verbatim. Modelled
  // attributes (name / cacheId / dataCaption / grand totals) are excluded
  // and written from the structured state. The form attributes (`compact`,
  // `compactData`, `outline`, `outlineData`) are captured only when they
  // depart from what `layout()` derives, and then replace the derived set.
  const std::vector<std::pair<std::string, std::string>>& passthrough_attrs() const { return passthrough_attrs_; }
  std::vector<std::pair<std::string, std::string>>& mutable_passthrough_attrs() { return passthrough_attrs_; }

  // Grand totals layout flags -----------------------------------------------

  bool grand_totals_rows() const { return grand_totals_rows_; }
  bool grand_totals_cols() const { return grand_totals_cols_; }
  void set_grand_totals(bool rows, bool cols) {
    grand_totals_rows_ = rows;
    grand_totals_cols_ = cols;
  }

  /// `subtotalHiddenItems`: hidden page items still count in every cell, as
  /// if the page filter were absent.
  bool subtotal_hidden_items() const { return subtotal_hidden_items_; }
  void set_subtotal_hidden_items(bool on) { subtotal_hidden_items_ = on; }

  /// `grandTotalCaption`: the label of the grand-total row and column, in
  /// place of the locale's word. Absent when not renamed.
  const std::optional<std::string>& grand_total_caption() const { return grand_total_caption_; }
  void set_grand_total_caption(std::optional<std::string> caption) { grand_total_caption_ = std::move(caption); }

  /// `useAutoFormatting`: Excel autofits the report's columns on update.
  /// A table read from a file takes the attribute (absent = false); one
  /// created in memory starts on, as Excel creates every pivot.
  bool auto_formatting() const { return auto_formatting_; }
  void set_auto_formatting(bool on) { auto_formatting_ = on; }

  /// `enableDrill`: a data cell may be expanded into its source records.
  bool enable_drill() const { return enable_drill_; }
  void set_enable_drill(bool on) { enable_drill_ = on; }

  /// `showError`: error cells show `error_caption`.
  bool show_error() const { return show_error_; }
  void set_show_error(bool v) { show_error_ = v; }

  /// `errorCaption`: the text error cells show while `show_error` is on.
  const std::optional<std::string>& error_caption() const { return error_caption_; }
  void set_error_caption(std::optional<std::string> v) { error_caption_ = std::move(v); }

  /// `showMissing`: empty cells show `missing_caption`.
  bool show_missing() const { return show_missing_; }
  void set_show_missing(bool v) { show_missing_ = v; }

  /// `missingCaption`: the text empty cells show while `show_missing` is on.
  const std::optional<std::string>& missing_caption() const { return missing_caption_; }
  void set_missing_caption(std::optional<std::string> v) { missing_caption_ = std::move(v); }

  /// `showHeaders`: the field headers are shown.
  bool show_headers() const { return show_headers_; }
  void set_show_headers(bool v) { show_headers_ = v; }

  /// `showDrill`: expand / collapse buttons are shown.
  bool show_drill() const { return show_drill_; }
  void set_show_drill(bool v) { show_drill_ = v; }

  /// `printDrill`: expand / collapse buttons are printed.
  bool print_drill() const { return print_drill_; }
  void set_print_drill(bool v) { print_drill_ = v; }

  /// `showDataTips`: cell tooltips are shown.
  bool show_data_tips() const { return show_data_tips_; }
  void set_show_data_tips(bool v) { show_data_tips_ = v; }

  /// `gridDropZones`: the classic layout, with drop zones drawn on the grid.
  bool grid_drop_zones() const { return grid_drop_zones_; }
  void set_grid_drop_zones(bool v) { grid_drop_zones_ = v; }

  /// `showEmptyRow`: items without data are shown on rows.
  bool show_empty_row() const { return show_empty_row_; }
  void set_show_empty_row(bool v) { show_empty_row_ = v; }

  /// `showEmptyCol`: items without data are shown on columns.
  bool show_empty_col() const { return show_empty_col_; }
  void set_show_empty_col(bool v) { show_empty_col_ = v; }

  /// `fieldPrintTitles`: field captions are repeated as print titles.
  bool field_print_titles() const { return field_print_titles_; }
  void set_field_print_titles(bool v) { field_print_titles_ = v; }

  /// `itemPrintTitles`: item labels are repeated on every printed page.
  bool item_print_titles() const { return item_print_titles_; }
  void set_item_print_titles(bool v) { item_print_titles_ = v; }

  /// `showDataDropDown`: the field drop-down buttons are shown.
  bool show_data_drop_down() const { return show_data_drop_down_; }
  void set_show_data_drop_down(bool v) { show_data_drop_down_ = v; }

  /// `multipleFieldFilters`: a field may carry more than one filter type.
  bool multiple_field_filters() const { return multiple_field_filters_; }
  void set_multiple_field_filters(bool v) { multiple_field_filters_ = v; }

  /// `preserveFormatting`: cell formatting survives a refresh.
  bool preserve_formatting() const { return preserve_formatting_; }
  void set_preserve_formatting(bool v) { preserve_formatting_ = v; }

  /// `mergeItem`: item labels are merged and centred.
  bool merge_item() const { return merge_item_; }
  void set_merge_item(bool v) { merge_item_ = v; }

  /// `indent`: the per-level indent of compact-form labels.
  std::uint32_t indent() const { return indent_; }
  void set_indent(std::uint32_t v) { indent_ = v; }

  /// `pageOverThenDown`: page fields fill rows before columns.
  bool page_over_then_down() const { return page_over_then_down_; }
  void set_page_over_then_down(bool v) { page_over_then_down_ = v; }

  /// `pageWrap`: page fields per row or column before wrapping; 0 is no wrap.
  std::uint32_t page_wrap() const { return page_wrap_; }
  void set_page_wrap(std::uint32_t v) { page_wrap_ = v; }

  /// `customListSort`: custom lists take part in sorting.
  bool custom_list_sort() const { return custom_list_sort_; }
  void set_custom_list_sort(bool v) { custom_list_sort_ = v; }

  /// x14 `altText`: the alternative-text title.
  const std::string& alt_text() const { return alt_text_; }
  void set_alt_text(std::string v) { alt_text_ = std::move(v); }

  /// x14 `altTextSummary`: the alternative-text description.
  const std::string& alt_text_summary() const { return alt_text_summary_; }
  void set_alt_text_summary(std::string v) { alt_text_summary_ = std::move(v); }

  /// Sets the options Excel gives a pivot it creates, which differ from the
  /// schema defaults a loaded definition falls back to.
  void apply_new_pivot_defaults() {
    item_print_titles_ = true;
    indent_ = 0;
    multiple_field_filters_ = false;
  }

  // Projection state ---------------------------------------------------------
  //
  // The sheet cells the last projection of the report covered and the merges
  // it added, so the next projection can clear exactly those.

  const std::optional<PivotRect>& projected_extent() const { return projected_extent_; }
  void set_projected_extent(std::optional<PivotRect> extent) { projected_extent_ = extent; }

  const std::vector<PivotRect>& projected_merges() const { return projected_merges_; }
  std::vector<PivotRect>& mutable_projected_merges() { return projected_merges_; }

  /// True when a change to the shared cache left the sheet cells stale.
  bool projection_pending() const { return projection_pending_; }
  void set_projection_pending(bool pending) { projection_pending_ = pending; }

 private:
  // Private so `clone()` is the only copy: the defaulted member-wise copy
  // shares the result mutex and cache until `clone()` replaces them.
  PivotTable(const PivotTable&) = default;

  static void erase_form_attrs(std::vector<std::pair<std::string, std::string>>& attrs) {
    std::vector<std::pair<std::string, std::string>> kept;
    for (auto& attr : attrs) {
      if (attr.first != "compact" && attr.first != "compactData" && attr.first != "outline" &&
          attr.first != "outlineData") {
        kept.push_back(std::move(attr));
      }
    }
    attrs = std::move(kept);
  }

  std::string name_;
  std::string data_caption_ = "Values";
  std::uint32_t pivot_cache_id_ = 0;
  std::vector<PivotField> fields_;
  std::vector<PivotDataField> data_fields_;
  std::vector<std::uint32_t> row_field_order_;
  std::vector<std::uint32_t> col_field_order_;
  std::optional<std::size_t> row_values_position_;
  std::optional<std::size_t> col_values_position_;
  std::vector<PivotPageField> page_fields_;
  PivotLayout layout_ = PivotLayout::Compact;
  std::uint32_t anchor_row_ = 0;
  std::uint32_t anchor_col_ = 0;
  std::uint32_t span_rows_ = 0;
  std::uint32_t span_cols_ = 0;
  bool has_authored_span_ = false;
  bool span_read_ = false;
  bool layout_dirty_ = false;
  std::optional<std::uint32_t> location_first_header_row_;
  std::optional<std::uint32_t> location_first_data_row_;
  std::optional<std::uint32_t> location_first_data_col_;
  std::optional<std::uint32_t> location_row_page_count_;
  std::optional<std::uint32_t> location_col_page_count_;
  bool grand_totals_rows_ = true;
  bool grand_totals_cols_ = true;
  bool subtotal_hidden_items_ = false;
  std::optional<std::string> grand_total_caption_;
  bool auto_formatting_ = true;
  bool enable_drill_ = true;
  bool show_error_ = false;
  std::optional<std::string> error_caption_;
  bool show_missing_ = true;
  std::optional<std::string> missing_caption_;
  bool show_headers_ = true;
  bool show_drill_ = true;
  bool print_drill_ = false;
  bool show_data_tips_ = true;
  bool grid_drop_zones_ = false;
  bool show_empty_row_ = false;
  bool show_empty_col_ = false;
  bool field_print_titles_ = false;
  bool item_print_titles_ = false;
  bool show_data_drop_down_ = true;
  bool multiple_field_filters_ = true;
  bool preserve_formatting_ = true;
  bool merge_item_ = false;
  std::uint32_t indent_ = 1;
  bool page_over_then_down_ = false;
  std::uint32_t page_wrap_ = 0;
  bool custom_list_sort_ = true;
  std::string alt_text_;
  std::string alt_text_summary_;
  std::optional<PivotRect> projected_extent_;
  std::vector<PivotRect> projected_merges_;
  bool projection_pending_ = false;
  std::vector<PivotFilter> active_filters_;
  std::vector<RetainedPivotFilter> retained_filters_;
  std::optional<PivotStyleInfo> style_info_;
  std::string raw_passthrough_xml_;
  std::string raw_passthrough_before_style_info_;
  std::string raw_passthrough_after_row_fields_;
  std::string raw_col_items_;
  std::vector<std::pair<std::string, std::string>> passthrough_attrs_;
  // Keep the mutex separately allocated so PivotTable remains movable.
  mutable std::shared_ptr<std::mutex> last_result_mutex_ = std::make_shared<std::mutex>();
  mutable std::shared_ptr<const PivotResult> last_result_;
  mutable RuntimeId runtime_id_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TABLE_H_
