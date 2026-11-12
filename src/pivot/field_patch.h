//
// Partial updates of one PivotTable field or field item.
//
// Every key of a patch is checked against the table before any is written, so
// a refused patch leaves the field as it was; the caller re-projects the
// report once afterwards. Unset keys leave the field alone and an all-unset
// patch changes nothing. Name and caption collisions use the pivot collation
// of the workbook profile, as the evaluator merges items.

#ifndef FORMULON_PIVOT_FIELD_PATCH_H_
#define FORMULON_PIVOT_FIELD_PATCH_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>

#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Workbook;
}  // namespace formulon

namespace formulon::pivot {

/// Longest field custom name or item caption Excel accepts, in UTF-16 code units.
// @limit: excel
inline constexpr std::uint32_t kMaxPivotNameUnits = 255U;

/// Item order a field patch sets; kManual keeps `PivotField::items` order and takes no by-field.
enum class FieldSortOrder : std::uint8_t { kManual = 0, kAscending = 1, kDescending = 2 };

struct PivotFieldPatch {
  /// Empty removes it; otherwise it may repeat the field's own source name, but no other field's names or any
  /// data field's name.
  std::optional<std::string> custom_name;
  std::optional<std::string> number_format;  ///< `numFmtId` as decimal text; empty removes it.
  std::optional<bool> default_subtotal;
  std::optional<bool> insert_blank_row;
  std::optional<bool> fill_down_labels;
  std::optional<bool> show_all;
  std::optional<bool> subtotal_top;
  std::optional<FieldSortOrder> sort_order;
  std::optional<std::string> sort_by_field;  ///< Empty sorts by label.
  std::optional<bool> show_detail;           ///< Expands or collapses every item and the default marker.
};

struct PivotItemPatch {
  std::optional<bool> visible;
  std::optional<bool> show_detail;
  /// Empty, or the item's own label, removes it; otherwise no other item of the field may show it.
  std::optional<std::string> caption;
};

/// True when every item and the default marker of `field` are expanded (a field without items is).
bool field_show_detail(const PivotField& field) noexcept;

/// Checks `patch` against field `field_index` of `table` without changing it. `kInvalidArgument` for a field
/// index out of range, a name too long, or a manual order with a by-field; `kAlreadyExists` for a name taken.
Expected<void, Error> check_field_patch(const Workbook& wb, const PivotTable& table, std::size_t field_index,
                                        const PivotFieldPatch& patch);

/// `check_field_patch`, then writes every set key; on failure the table is unchanged.
Expected<void, Error> apply_field_patch(const Workbook& wb, PivotTable& table, std::size_t field_index,
                                        const PivotFieldPatch& patch);

/// Checks `patch` against item `item_index` of field `field_index` without changing it. `kInvalidArgument` for
/// an index out of range or a caption too long; `kAlreadyExists` for a caption shown by another item.
Expected<void, Error> check_item_patch(const Workbook& wb, const PivotTable& table, std::size_t field_index,
                                       std::size_t item_index, const PivotItemPatch& patch);

/// `check_item_patch`, then writes every set key; on failure the table is unchanged.
Expected<void, Error> apply_item_patch(const Workbook& wb, PivotTable& table, std::size_t field_index,
                                       std::size_t item_index, const PivotItemPatch& patch);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_FIELD_PATCH_H_
