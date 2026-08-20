//
// In-memory representation of an OOXML pivot cache (pair of
// `xl/pivotCache/cacheDefinition*.xml` and `cacheRecords*.xml`). The cache
// is a snapshot of the source data taken at refresh time; the pivot
// evaluator consumes it without reaching back into the source workbook.

#ifndef FORMULON_PIVOT_PIVOT_CACHE_H_
#define FORMULON_PIVOT_PIVOT_CACHE_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "value.h"

namespace formulon::pivot {

/// Numeric / date range and grouping hints carried on a `<sharedItems>`
/// element. Excel writes these so it can group and refresh without
/// re-scanning every record. Each attribute is stored as the raw OOXML
/// string body when present (and `has_*` records presence) so the writer
/// re-emits exactly what was read; absent attributes stay absent. The
/// `min_value` / `max_value` are kept as strings rather than doubles so a
/// date-typed range (`minDate`/`maxDate` ISO bodies) and a numeric range
/// share one carrier without lossy reformatting.
struct SharedItemsHints {
  bool contains_number = false;
  bool contains_integer = false;
  bool contains_date = false;
  bool contains_string = false;
  bool contains_blank = false;
  bool contains_semi_mixed = false;
  bool contains_non_date = false;
  bool contains_mixed_types = false;
  // Presence flags for the boolean content hints. Several of these default
  // to true in OOXML (containsString / containsSemiMixedTypes /
  // containsNonDate), so a plain value flag cannot round-trip an absent
  // attribute (absent would be re-emitted as its false value, flipping the
  // meaning). Tracking presence lets the writer re-emit each attribute
  // verbatim — present stays present with its body, absent stays absent —
  // sidestepping default semantics entirely.
  bool has_contains_number = false;
  bool has_contains_integer = false;
  bool has_contains_date = false;
  bool has_contains_string = false;
  bool has_contains_blank = false;
  bool has_contains_semi_mixed = false;
  bool has_contains_non_date = false;
  bool has_contains_mixed_types = false;
  // Presence flags + raw bodies for the bound attributes.
  bool has_min_value = false;
  bool has_max_value = false;
  bool has_min_date = false;
  bool has_max_date = false;
  std::string min_value;
  std::string max_value;
  std::string min_date;
  std::string max_date;
  /// True when the source `<sharedItems>` carried any of the above hint
  /// attributes. When false the writer falls back to its legacy minimal
  /// placeholder for an empty (range-typed) field.
  bool present = false;
};

/// One column of the pivot cache.
///
/// `shared_items` enumerates the distinct values for a discrete (string /
/// boolean) column. Range-typed (numeric / date) columns leave it empty
/// and store inline values on each `PivotCacheRecord`.
struct PivotCacheField {
  PivotCacheField() = default;
  /// Convenience constructor used by hand-built caches (tests, the cache
  /// reader): the `<sharedItems>` hints default to absent. Keeping this
  /// lets call sites write `PivotCacheField{"Region", {...}}` without
  /// spelling the hint set every time.
  PivotCacheField(std::string field_name, std::vector<Value> items)
      : name(std::move(field_name)), shared_items(std::move(items)) {}

  std::string name;
  std::vector<Value> shared_items;
  /// Numeric / date range + grouping hints from `<sharedItems>`. Preserved
  /// verbatim so Excel's Refresh keeps its grouping boundaries.
  SharedItemsHints shared_items_hints;

  /// Whether this cache field is a *database* field (backed by a source
  /// column) rather than one derived by grouping. OOXML marks a derived
  /// field with `databaseField="0"`; only database fields contribute a
  /// cell to each `<r>` record, so the record arity is the number of
  /// database fields, not the total field count. Defaults to true so
  /// hand-built caches keep the "every field is a source column" shape.
  bool is_database_field = true;

  /// Raw `<fieldGroup>` element captured verbatim (date / numeric / discrete
  /// grouping definition) so a grouped field round-trips even though v1.0
  /// does not model grouping structurally. Empty when the field carries no
  /// `<fieldGroup>`.
  std::string field_group_xml;
};

/// The `<cacheSource>/<worksheetSource>` reference that tells Excel where
/// to re-read the cache from on Refresh. Any of the attributes may be
/// absent; presence is tracked so the writer re-emits only what was read.
/// Dropping these makes Excel's Refresh fail or repoint incorrectly.
struct WorksheetSource {
  bool present = false;
  std::string ref;    ///< A1 range, e.g. "Sheet1!$A$1:$C$9" or "$A$1:$C$9".
  std::string sheet;  ///< Sheet name when `ref` is unqualified.
  std::string name;   ///< Defined-name source (alternative to ref).
};

/// One row of the pivot cache.
///
/// Each `cells` entry is either a `Number` index into the corresponding
/// field's `shared_items` (source `<x v="N"/>`) or an inline `Value`
/// (source `<n>`/`<s>`/`<b>`/`<d>`/`<m>`/`<e>`). Encoding the index as a
/// `Number` keeps the `Value` shape uniform.
///
/// `cell_is_index` records, per cell, which of the two encodings the
/// source used, so a numeric inline value in a shared field is not
/// mistaken for an index (or vice versa) on write. It is sized in
/// lock-step with `cells` by the reader. When it is empty — a hand-built
/// cache (C API, tests) that does not populate it — consumers fall back
/// to inferring the encoding from the field's `shared_items` being
/// non-empty, preserving the legacy behaviour.
/// `cell_text_slot` entry for a cell whose Text payload this cache does not
/// own — either the cell is not Text, or its bytes come from somewhere else
/// (the reader's own append, the workbook shared-string table, ...).
inline constexpr std::size_t kNoCacheTextSlot = static_cast<std::size_t>(-1);

struct PivotCacheRecord {
  std::vector<Value> cells;
  std::vector<bool> cell_is_index;
  /// Per-cell index into `PivotCache::text_storage()` for the entry that
  /// backs this cell's Text payload, or `kNoCacheTextSlot`.
  ///
  /// Only `PivotCache::set_record_text` populates this; the reader leaves
  /// it empty and consumers that never overwrite a cell never look at it.
  /// Recording the slot is what lets an overwrite reuse the storage the
  /// previous write allocated instead of stranding it.
  std::vector<std::size_t> cell_text_slot;
};

/// Owning container for the cache definition + records. Held by the
/// workbook (one cache may back multiple pivot tables) and accessed by
/// the pivot evaluator and OOXML round-trip path.
class PivotCache {
 public:
  PivotCache() = default;

  PivotCache(const PivotCache&) = delete;
  PivotCache& operator=(const PivotCache&) = delete;
  PivotCache(PivotCache&&) = default;
  PivotCache& operator=(PivotCache&&) = default;

  std::uint32_t cache_id() const { return cache_id_; }
  void set_cache_id(std::uint32_t id) { cache_id_ = id; }

  const std::vector<PivotCacheField>& fields() const { return fields_; }
  std::vector<PivotCacheField>& mutable_fields() { return fields_; }

  const std::vector<PivotCacheRecord>& records() const { return records_; }
  std::vector<PivotCacheRecord>& mutable_records() { return records_; }

  /// Unmodelled `<pivotCacheDefinition>` root attributes (`refreshedBy`,
  /// `refreshedDate`, `refreshOnLoad`, `createdVersion`, ...), captured as
  /// `(name, value)` pairs so the writer re-emits them verbatim. The
  /// modelled attributes (`r:id` / `recordCount`, plus the namespace
  /// declarations) are excluded and written from the structured state.
  const std::vector<std::pair<std::string, std::string>>& passthrough_attrs() const { return passthrough_attrs_; }
  std::vector<std::pair<std::string, std::string>>& mutable_passthrough_attrs() { return passthrough_attrs_; }

  /// Lifetime-stable backing store for any `Value::text` payload that the
  /// reader produces (whether on a `shared_items` entry or inline on a
  /// record cell). The reader appends one entry per decoded string and
  /// builds the `Value` from `string_view(back())`. A `std::deque` is
  /// chosen for pointer / iterator stability across appends; using a
  /// `std::vector` would invalidate the views on growth. Callers that
  /// hand-build a `PivotCache` (e.g. tests) may also append here when
  /// they want the cache to own a literal string.
  std::deque<std::string>& mutable_text_storage() { return text_storage_; }
  const std::deque<std::string>& text_storage() const { return text_storage_; }

  /// Writes `utf8` into record `record_idx`'s cell `field_idx` as a Text
  /// value backed by this cache, reusing the storage a previous write to
  /// the same coordinate allocated. Returns false — leaving the cache
  /// untouched — when the coordinate is out of range.
  ///
  /// Appending a fresh `text_storage_` entry per write is right for
  /// `shared_items`, which only ever grows, and wrong for a record cell:
  /// overwriting one that way strands the replaced string in the deque
  /// forever, so a live editor's memory tracks how many times the source
  /// was edited rather than how much data it holds. One storage slot per
  /// written coordinate bounds the store by the cells that exist.
  bool set_record_text(std::size_t record_idx, std::size_t field_idx, std::string_view utf8) {
    if (record_idx >= records_.size()) {
      return false;
    }
    PivotCacheRecord& record = records_[record_idx];
    if (field_idx >= record.cells.size()) {
      return false;
    }
    if (record.cell_text_slot.size() < record.cells.size()) {
      record.cell_text_slot.resize(record.cells.size(), kNoCacheTextSlot);
    }
    std::size_t& slot = record.cell_text_slot[field_idx];
    if (slot == kNoCacheTextSlot) {
      slot = text_storage_.size();
      text_storage_.emplace_back(utf8);
    } else {
      // `std::deque` keeps the element at a stable address, so replacing the
      // string's contents in place does not disturb any other slot; the
      // `Value` rebuilt below is the only view of these bytes.
      text_storage_[slot].assign(utf8);
    }
    record.cells[field_idx] = Value::text(text_storage_[slot]);
    return true;
  }

  /// The `<worksheetSource>` reference under `<cacheSource>`. Preserved so
  /// Excel's Refresh can locate the source range / defined name.
  const WorksheetSource& worksheet_source() const { return worksheet_source_; }
  WorksheetSource& mutable_worksheet_source() { return worksheet_source_; }

 private:
  std::uint32_t cache_id_ = 0;
  std::vector<PivotCacheField> fields_;
  std::vector<PivotCacheRecord> records_;
  std::deque<std::string> text_storage_;
  WorksheetSource worksheet_source_;
  std::vector<std::pair<std::string, std::string>> passthrough_attrs_;
};

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_CACHE_H_
