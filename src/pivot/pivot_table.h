//
// `PivotTable` is the workbook-level pivot definition: identity, cache
// binding, field configuration, layout, anchor location, transient slicer
// state, and the most recent evaluation result. One pivot table is owned
// by exactly one sheet (the sheet supplies the sheet identity); it points
// at a workbook-owned `PivotCache` by id.

#ifndef FORMULON_PIVOT_PIVOT_TABLE_H_
#define FORMULON_PIVOT_PIVOT_TABLE_H_

#include <cstdint>
#include <memory>
#include <mutex>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_result.h"
#include "pivot/pivot_types.h"

namespace formulon::pivot {

/// One pivot table definition + most recent evaluation result.
///
/// Move-only. Designed to be held by the owning `Sheet` via
/// `std::unique_ptr<PivotTable>` so the address is stable for as long as
/// the sheet retains the pivot.
class PivotTable {
 public:
  PivotTable() = default;

  PivotTable(const PivotTable&) = delete;
  PivotTable& operator=(const PivotTable&) = delete;
  PivotTable(PivotTable&&) = default;
  PivotTable& operator=(PivotTable&&) = default;

  // Identity / cache binding -------------------------------------------------

  const std::string& name() const { return name_; }
  void set_name(std::string n) { name_ = std::move(n); }

  std::uint32_t pivot_cache_id() const { return pivot_cache_id_; }
  void set_pivot_cache_id(std::uint32_t id) { pivot_cache_id_ = id; }

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

  // Page (report filter) axis -----------------------------------------------
  //
  // Decoded from `<pageFields>` for rendering only. The block itself stays
  // in the pre-dataFields passthrough bin and the writer re-emits it from
  // there, so this list is read-side state exactly like
  // `authored_caption_filters()`: populating it changes not one byte of a
  // round trip, and clearing it would not remove the page field from a
  // saved file. The one thing the writer does own is the element's absence
  // — see `write_pivot_table_definition`, which synthesises `<pageFields>`
  // only for a table that never carried one.
  //
  // Selection is not a filter: a page field's hidden items already prune
  // records through `PreparedRecordFilter`, which keys off
  // `PivotField::items` on every axis alike. What this list adds is the
  // report order and the explicit single-item selection, neither of which
  // is recoverable from `<pivotFields>`.
  const std::vector<PivotPageField>& page_fields() const { return page_fields_; }
  std::vector<PivotPageField>& mutable_page_fields() { return page_fields_; }

  /// Indices into `fields()` of the page-axis fields, in report order.
  ///
  /// `<pageFields>` states that order explicitly and wins when it was
  /// decoded. A table assembled in memory (C API, the workbook-oracle
  /// builder) sets only `PivotField::axis`, so the fallback is `fields()`
  /// document order — which is what Excel itself falls back to when a page
  /// field carries no `<pageField>` entry of its own.
  std::vector<std::uint32_t> page_field_order() const {
    std::vector<std::uint32_t> order;
    if (!page_fields_.empty()) {
      order.reserve(page_fields_.size());
      for (const PivotPageField& page : page_fields_) {
        if (page.field_index < fields_.size()) {
          order.push_back(page.field_index);
        }
      }
      return order;
    }
    for (std::size_t i = 0; i < fields_.size(); ++i) {
      if (fields_[i].axis == PivotAxis::Page) {
        order.push_back(static_cast<std::uint32_t>(i));
      }
    }
    return order;
  }

  // Layout / location --------------------------------------------------------

  PivotLayout layout() const { return layout_; }
  void set_layout(PivotLayout l) { layout_ = l; }

  std::uint32_t anchor_row() const { return anchor_row_; }
  std::uint32_t anchor_col() const { return anchor_col_; }
  std::uint32_t span_rows() const { return span_rows_; }
  std::uint32_t span_cols() const { return span_cols_; }

  void set_anchor(std::uint32_t row, std::uint32_t col, std::uint32_t rows, std::uint32_t cols) {
    anchor_row_ = row;
    anchor_col_ = col;
    span_rows_ = rows;
    span_cols_ = cols;
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
  void mark_span_authored() { has_authored_span_ = true; }
  void clear_span_authored() { has_authored_span_ = false; }

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

  // Transient slicer-applied filters ----------------------------------------
  //
  // Kept separate from `PivotField::items` (which is the XML-defined manual
  // filter state). Folding slicer state into `items` would make the
  // `evaluator -> XML round-trip` path overwrite the authored definition.
  //
  // This list is session state, deliberately, in both directions:
  //
  //   * it is never serialised — `write_pivot_table_definition` emits no
  //     `<filters>` block, so entries added here shape every evaluation for
  //     as long as this object lives and are gone after a save + reload;
  //   * it is never populated by the reader — an authored `<filters>` block
  //     decodes into `authored_caption_filters()` instead, so a slicer
  //     selection and a persisted filter never overwrite one another.
  const std::vector<PivotFilter>& active_filters() const { return active_filters_; }
  std::vector<PivotFilter>& mutable_active_filters() { return active_filters_; }

  // Authored `<filters>` caption filters ------------------------------------
  //
  // Decoded from the OOXML `<filters>` block for evaluation only. The
  // block itself stays in the verbatim tail bin (`raw_passthrough_xml_`)
  // and the writer re-emits it from there, so this list is read-side
  // state: populating it does not change a single byte of the round trip,
  // and clearing it would not remove the filter from a saved file.
  //
  // Splitting the decode from the serialisation is what makes modelling
  // `<filters>` affordable. Owning the block outright would mean
  // authoring filter XML — including the nested `<autoFilter>` criteria —
  // which this version has no Excel-produced reference file to check
  // against. Reading it needs no such reference: the criteria reduce to a
  // predicate the bound cache already spells out.
  //
  // The two lists split by when the rule can be decided. A caption or
  // date criterion is a property of one source record, so it prunes
  // before aggregation; a value criterion ranks an axis leaf by its
  // aggregate, so it cannot be decided until the aggregates exist. That
  // is the same split `active_filters()` already drives, and both lists
  // feed the same two passes.
  //
  // The relative-period families (`thisMonth`, `yearToDate`, ...) form a
  // third list. They carry no criteria in the file — the window is implied
  // by the type name — so they resolve against a clock reading supplied at
  // evaluation time. Pinning that reading (`Workbook::pinned_now`) is what
  // keeps a pivot that uses them reproducible rather than dependent on the
  // day it happened to be computed.
  const std::vector<AuthoredCaptionFilter>& authored_caption_filters() const { return authored_caption_filters_; }
  std::vector<AuthoredCaptionFilter>& mutable_authored_caption_filters() { return authored_caption_filters_; }

  const std::vector<AuthoredValueFilter>& authored_value_filters() const { return authored_value_filters_; }
  std::vector<AuthoredValueFilter>& mutable_authored_value_filters() { return authored_value_filters_; }

  const std::vector<AuthoredPeriodFilter>& authored_period_filters() const { return authored_period_filters_; }
  std::vector<AuthoredPeriodFilter>& mutable_authored_period_filters() { return authored_period_filters_; }

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
  // OOXML pivot definitions can carry several elements that v1.0 does
  // not model structurally — `<rowItems>`, `<colItems>`, `<pageFields>`,
  // `<calculatedFields>`, `<calculatedItems>`, `<pivotTableStyleInfo>`,
  // `<chartFormats>`, `<formats>`, etc. The reader captures them as
  // concatenated raw XML so the writer can emit them back verbatim,
  // keeping a round trip stable for files that depend on those features.
  //
  // `CT_pivotTableDefinition` mandates a strict child order, so a single
  // trailing buffer would re-emit `<rowItems>`/`<colItems>`/`<pageFields>`
  // after `<dataFields>` and make Excel flag the file for repair. Those
  // three are the only unmodelled elements the schema places *before*
  // `<dataFields>`, so the reader bins them by name into the two
  // pre-dataFields slots (rowItems after `<rowFields>`; colItems and
  // pageFields after `<colFields>`) and routes everything else — formats,
  // style info, extLst, ... — into the tail bin `raw_passthrough_xml_`.
  // The writer flushes each bin at the matching slot so schema order is
  // preserved. Mutating the structured state does not invalidate these
  // buffers; consumers that need exact bit parity should regenerate or
  // strip them.
  const std::string& raw_passthrough_xml() const { return raw_passthrough_xml_; }
  std::string& mutable_raw_passthrough_xml() { return raw_passthrough_xml_; }

  const std::string& raw_passthrough_after_row_fields() const { return raw_passthrough_after_row_fields_; }
  std::string& mutable_raw_passthrough_after_row_fields() { return raw_passthrough_after_row_fields_; }

  const std::string& raw_passthrough_after_col_fields() const { return raw_passthrough_after_col_fields_; }
  std::string& mutable_raw_passthrough_after_col_fields() { return raw_passthrough_after_col_fields_; }

  // Unmodelled `<pivotTableDefinition>` root attributes (`updatedVersion`,
  // `createdVersion`, `itemPrintTitles`, `indent`, ...), captured as
  // `(name, value)` pairs so the writer re-emits them verbatim. Modelled
  // attributes (name / cacheId / dataCaption / grand totals / compact /
  // outline) are excluded and written from the structured state.
  const std::vector<std::pair<std::string, std::string>>& passthrough_attrs() const { return passthrough_attrs_; }
  std::vector<std::pair<std::string, std::string>>& mutable_passthrough_attrs() { return passthrough_attrs_; }

  // Grand totals layout flags -----------------------------------------------

  bool grand_totals_rows() const { return grand_totals_rows_; }
  bool grand_totals_cols() const { return grand_totals_cols_; }
  void set_grand_totals(bool rows, bool cols) {
    grand_totals_rows_ = rows;
    grand_totals_cols_ = cols;
  }

 private:
  std::string name_;
  std::string data_caption_ = "Values";
  std::uint32_t pivot_cache_id_ = 0;
  std::vector<PivotField> fields_;
  std::vector<PivotDataField> data_fields_;
  std::vector<std::uint32_t> row_field_order_;
  std::vector<std::uint32_t> col_field_order_;
  std::vector<PivotPageField> page_fields_;
  PivotLayout layout_ = PivotLayout::Compact;
  std::uint32_t anchor_row_ = 0;
  std::uint32_t anchor_col_ = 0;
  std::uint32_t span_rows_ = 0;
  std::uint32_t span_cols_ = 0;
  bool has_authored_span_ = false;
  std::optional<std::uint32_t> location_first_header_row_;
  std::optional<std::uint32_t> location_first_data_row_;
  std::optional<std::uint32_t> location_first_data_col_;
  std::optional<std::uint32_t> location_row_page_count_;
  std::optional<std::uint32_t> location_col_page_count_;
  bool grand_totals_rows_ = true;
  bool grand_totals_cols_ = true;
  std::vector<PivotFilter> active_filters_;
  std::vector<AuthoredCaptionFilter> authored_caption_filters_;
  std::vector<AuthoredValueFilter> authored_value_filters_;
  std::vector<AuthoredPeriodFilter> authored_period_filters_;
  std::string raw_passthrough_xml_;
  std::string raw_passthrough_after_row_fields_;
  std::string raw_passthrough_after_col_fields_;
  std::vector<std::pair<std::string, std::string>> passthrough_attrs_;
  // Keep the mutex separately allocated so PivotTable remains movable.
  mutable std::shared_ptr<std::mutex> last_result_mutex_ = std::make_shared<std::mutex>();
  mutable std::shared_ptr<const PivotResult> last_result_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_TABLE_H_
