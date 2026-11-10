//
// Rebuilding a pivot cache from the current values of its source, as Excel's
// Refresh does. The function here is pure: the workbook reads the source and
// applies the result (`Workbook::refresh_pivot_cache`).
//
// Field names come from the header row's displayed text. A blank header takes
// the locale's word plus its 1-based column; names equal under the header
// fold (case, width and kana type, spaces significant) are told apart left to
// right by the smallest free suffix from 2, literal headers before generated
// ones. Old fields match new ones by name under that fold, keep their state
// and take the new spelling; an unmatched old field is dropped, an unmatched
// new one is unplaced. Cache fields follow the source's column order, then the
// old group and calculated fields in their order.
//
// Only fields some table puts on an axis (and bases of discrete groups, and
// fields with calculated items) keep shared items: the source's distinct
// values in first-appearance order under the item fold (case only; numbers,
// dates and text never fold together), the first spelling kept; then the old
// items the source lost, kept as missing unless `missingItemsLimit` is 0; then
// the calculated items. Every other field holds its values in the records,
// each text folded to the first spelling of its item.
//
// A field that joins an axis later takes shared items by the same rule
// (`share_field_items`), and keeps them, and its item list, when it leaves.

#ifndef FORMULON_PIVOT_PIVOT_REFRESH_H_
#define FORMULON_PIVOT_PIVOT_REFRESH_H_

#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "excel_profile.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot {

/// The source as read: the header row's displayed text and the data rows
/// below it.
struct RefreshSource {
  std::uint32_t columns = 0;
  std::uint32_t rows = 0;
  /// Per column, the header cell's displayed text; empty for a blank header.
  std::vector<std::string> headers;
  /// `rows` x `columns` values, row-major. Text views point into `text`.
  std::vector<Value> cells;
  /// Per cell, true for a number whose cell is formatted as a date or time.
  std::vector<bool> cell_is_date;
  /// Per column, the number format id every data cell shares, else 0.
  std::vector<std::uint32_t> num_fmt_ids;
  std::deque<std::string> text;
};

/// What the rebuild needs from the workbook besides the source.
struct RefreshContext {
  /// Per old cache field, true when a table on the cache places it on the
  /// row, column or page axis.
  std::vector<bool> axis_fields;
  /// The word a blank header is named by.
  std::string_view blank_header_word;
  ExcelProfile profile;
  bool date1904 = false;
};

/// The rebuilt cache and how its fields and items moved from the old one.
struct RefreshPlan {
  PivotCache cache;
  FieldRemap remap;
};

/// Names for source columns with displayed headers `headers`, as the header
/// rules give them; a blank header is `blank_word` plus its 1-based column.
std::vector<std::string> refreshed_field_names(const std::vector<std::string>& headers, std::string_view blank_word);

/// True when `a` and `b` name the same field: equal under case, width and
/// kana-type folding, spaces significant.
bool same_header(std::string_view a, std::string_view b);

/// Rebuilds `old` from `source`.
///
/// A numeric range group with an automatic bound takes the base's current
/// minimum or maximum and regenerates its items as `set_numeric_group` does;
/// one with both bounds fixed keeps its items, values outside falling in the
/// first or last.
///
/// Errors, none of which builds anything: `kNotImplemented` for a numeric
/// range group with a non-integer start, end or interval, for a years or
/// multi-day date group whose bounds move, and for a blank header in a locale
/// whose word is unmeasured.
Expected<RefreshPlan, Error> build_refreshed_cache(const PivotCache& old, const RefreshSource& source,
                                                   const RefreshContext& context);

/// Gives database field `field` of `cache`, when it holds its values inline,
/// the shared items a refresh would: its record values in record order under
/// the item fold, a blank included, each item keeping its first spelling and
/// date mark. Every record cell of the field then names its item by index, and
/// later record writes go through the items (`items_follow_values`). The
/// values, and so the field's hints, are unchanged.
///
/// Returns false, changing nothing, for a field that already has shared
/// items, a calculated or group field, a field index out of range, and a
/// cache without records.
bool share_field_items(PivotCache& cache, std::size_t field);

/// Shares (`share_field_items`) every field `table` places on the row, column
/// or page axis.
void share_axis_fields(PivotCache& cache, const PivotTable& table);

/// The shared item of `field` that `value` (a date when `is_date`) is under the
/// item fold, if any.
std::optional<std::uint32_t> find_shared_item(const PivotCacheField& field, const Value& value, bool is_date);

/// Writes `value` (a date when `is_date`) to cell `(record, field)` of a field
/// whose items follow its values: the cell names the item `value` is under the
/// item fold, appended with its own spelling when there is none. `record` and
/// `field` must be in range.
void write_shared_record_value(PivotCache& cache, std::size_t record, std::size_t field, const Value& value,
                               bool is_date);

/// Lists in `field` every item its shared items give it, as a field joining or
/// leaving the axes lists them and a save writes them: with no items, one per shared item in the field's
/// sort order (ascending by default, blank last; shared-item order when sorted
/// manually or by a data field); otherwise the items it has, then the shared
/// items none of them names. Missing and calculated shared items are never
/// added. The caller holds an `eval::ScopedEvalProfile`.
void keep_field_items(PivotField& field, const PivotCacheField& cache_field);

/// Lists (`keep_field_items`) the items of every field `table` places on the
/// row, column or page axis that has none yet. The caller holds an
/// `eval::ScopedEvalProfile`.
void list_axis_field_items(PivotTable& table, const PivotCache& cache);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_REFRESH_H_
