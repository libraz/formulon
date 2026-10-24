//
// Typed access to the table slicers of a worksheet: listing, adding, updating
// and removing them.
//
// A table slicer spans several parts, all of which stay raw passthrough
// parts: the sheet's `xl/slicers/slicerN.xml`, the workbook's
// `xl/slicerCaches/slicerCacheN.xml`, the `slicerList` / `slicerCaches`
// entries of the sheet and workbook extension lists, the `Slicer_*` defined
// name, and the slicer's frame in the sheet's drawing. Every edit patches the
// DOM of the parts it touches, so attributes and extensions the model does
// not carry survive. The table AutoFilter a slicer drives is not part of this
// model; use the table AutoFilter API for the selection.
//
// Slicers over a pivot cache are neither listed nor edited. A slicer is
// addressed by its position in `list_slicers`, which follows the order of the
// sheet's slicer parts.

#ifndef FORMULON_DRAWING_SLICER_EDIT_H_
#define FORMULON_DRAWING_SLICER_EDIT_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_xml.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// `tableSlicerCache/@sortOrder`.
enum class SlicerSortOrder : std::int32_t { kAscending = 0, kDescending = 1 };

/// `tableSlicerCache/@crossFilter`.
enum class SlicerCrossFilter : std::int32_t { kShowItemsWithDataAtTop = 0, kShowItemsWithNoData = 1, kNone = 2 };

/// A listed slicer. `object` is the drawing frame (id, anchor, size); its
/// `object_id` is 0 when the sheet's drawing holds no frame for the slicer.
struct SlicerInfo {
  DrawingObject object;
  std::string name;
  std::string caption;
  std::string cache_name;
  std::string table_name;
  std::string column_name;
  std::string style;               ///< Empty when the slicer names no style.
  std::uint32_t column_count = 1;  ///< Item columns.
  std::int64_t row_height = 0;     ///< EMU.
  SlicerSortOrder sort_order = SlicerSortOrder::kAscending;
  SlicerCrossFilter cross_filter = SlicerCrossFilter::kShowItemsWithDataAtTop;
};

/// `SlicerSpec::fields` bits: the fields an add or update applies.
namespace slicer_field {
inline constexpr std::uint32_t kName = 1U << 0;  ///< Add only.
inline constexpr std::uint32_t kCaption = 1U << 1;
inline constexpr std::uint32_t kStyle = 1U << 2;  ///< Empty removes the style.
inline constexpr std::uint32_t kColumnCount = 1U << 3;
inline constexpr std::uint32_t kRowHeight = 1U << 4;
inline constexpr std::uint32_t kAnchor = 1U << 5;  ///< Required by add.
inline constexpr std::uint32_t kSortOrder = 1U << 6;
inline constexpr std::uint32_t kCrossFilter = 1U << 7;
inline constexpr std::uint32_t kAll = (1U << 8) - 1U;
}  // namespace slicer_field

/// Values for an add or update; fields outside `fields` are ignored. The
/// anchor is the two-cell `from` / `to` marker pair of the slicer's frame.
struct SlicerSpec {
  std::uint32_t fields = 0;
  std::string name;     ///< Made unique among the workbook's slicers; defaults to the column name.
  std::string caption;  ///< Defaults to the column name.
  std::string style;
  std::uint32_t column_count = 1;
  std::int64_t row_height = 0;  ///< EMU; defaults to 241300.
  AnchorPoint from;
  AnchorPoint to;
  SlicerSortOrder sort_order = SlicerSortOrder::kAscending;
  SlicerCrossFilter cross_filter = SlicerCrossFilter::kShowItemsWithDataAtTop;
};

/// Table slicers hosted by sheet `sheet`, in list order. Fails with
/// `kInvalidArgument` for a bad sheet and `kIoDrawingUnparseable` when the
/// sheet's drawing does not parse.
Expected<std::vector<SlicerInfo>, Error> list_slicers(const Workbook& wb, std::size_t sheet);

/// Adds a slicer on `sheet` over column `column` (0-based) of table
/// `table_index`, with its cache, relationships, content types, `Slicer_*`
/// defined name and drawing frame, and returns the frame's `cNvPr` id.
/// Fails with `kInvalidArgument` for a bad sheet, table, column, anchor or
/// text, and `kIoDrawingUnparseable` for a drawing that does not parse or was
/// retained from an `.xlsb` package.
Expected<std::uint32_t, Error> add_slicer(Workbook& wb, std::size_t sheet, std::size_t table_index, std::size_t column,
                                          const SlicerSpec& spec);

/// Applies the fields `spec.fields` names (all but `kName`) to slicer `index`.
Expected<void, Error> update_slicer(Workbook& wb, std::size_t sheet, std::size_t index, const SlicerSpec& spec);

/// Removes slicer `index` with its drawing frame; its cache goes too when no
/// other slicer uses it.
Expected<void, Error> remove_slicer(Workbook& wb, std::size_t sheet, std::size_t index);

/// Removes every slicer and cache bound to the table with `table_id`. Call
/// before the table leaves the workbook.
void remove_table_slicers(Workbook& wb, std::uint32_t table_id);

/// `remove_table_slicers` for a caller that already holds the workbook lock:
/// returns the names of the caches that went and leaves their `Slicer_*`
/// defined names to the caller.
std::vector<std::string> remove_table_slicers_locked(Workbook& wb, std::uint32_t table_id);

/// Removes the slicers hosted on `sheet` and those bound to its tables, and
/// returns the names of the caches that went with them. It leaves their
/// `Slicer_*` defined names to the caller, which already holds the workbook
/// lock. Call before the sheet leaves the workbook.
std::vector<std::string> remove_sheet_slicers(Workbook& wb, std::size_t sheet);

/// Names of the slicers of `sheet` that have a drawing frame; pass them to `remove_unframed_slicers` after an
/// edit that may delete frames.
std::vector<std::string> framed_slicers(const Workbook& wb, std::size_t sheet);

/// Removes the slicers named in `framed` that no longer have a frame, with their caches when unused, and returns
/// the names of the caches that went. The caller removes their `Slicer_*` defined names (it holds the workbook
/// lock already).
std::vector<std::string> remove_unframed_slicers(Workbook& wb, std::size_t sheet,
                                                 const std::vector<std::string>& framed);

}  // namespace formulon

#endif  // FORMULON_DRAWING_SLICER_EDIT_H_
