//
// Typed drawing edits: insert and remove images in a sheet's drawing part.
//
// Every edit parses the drawing part, patches the DOM and re-serialises it,
// so charts, shapes and anything else the part holds stay intact; a part no
// edit touched keeps its bytes. Inserting into a sheet without a drawing
// creates `xl/drawings/drawingN.xml`, its rels part and the sheet's drawing
// relationship. Media parts are `Default`-typed by extension.
//
// Refusals: a drawing part (or its rels part) that does not parse fails with
// `kIoDrawingUnparseable`, as does a sheet whose drawing was retained from an
// `.xlsb` package (reachable only through `Sheet::unknown_relationships()`);
// bytes that are not a recognised image fail with `kIoImageUnsupported`.
//
// Row/column edits follow Excel's anchor rules: a move-and-size (two-cell)
// object moves each marker, a marker inside a deleted span collapses to the
// deletion boundary with offset 0, and an object whose span is deleted
// entirely is removed; a move-only (one-cell) object moves its `from` the same
// way and keeps its size; an absolute object never moves.

#ifndef FORMULON_DRAWING_DRAWING_EDIT_H_
#define FORMULON_DRAWING_DRAWING_EDIT_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_xml.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Placement of an inserted image.
struct ImageInsertOptions {
  std::string name;                               ///< `cNvPr/@name`; empty yields `Picture <id>`.
  std::string descr;                              ///< Alternative text; omitted when empty.
  AnchorKind anchor_kind = AnchorKind::kOneCell;  ///< `kOneCell` or `kTwoCell`.
  EditAs edit_as = EditAs::kTwoCell;              ///< Honoured for `kTwoCell` only.
  std::uint32_t row = 0;                          ///< 0-based top-left cell.
  std::uint32_t col = 0;
  std::int64_t row_off = 0;  ///< EMU offset inside the cell.
  std::int64_t col_off = 0;
  std::int64_t width_emu = 0;   ///< 0 means pixel width x 9525.
  std::int64_t height_emu = 0;  ///< 0 means pixel height x 9525.
};

/// Largest image `insert_image` accepts, in bytes.
inline constexpr std::size_t kMaxImageBytes = 32U * 1024U * 1024U;

/// Lists the anchored objects of sheet `sheet`'s drawing in document order,
/// with `image_format` filled for pictures whose media part is present. A
/// sheet without a drawing yields an empty list.
Expected<std::vector<DrawingObject>, Error> list_drawing_objects(const Workbook& wb, std::size_t sheet);

/// Inserts `bytes` as a picture and returns its `cNvPr` id, which is unique
/// within the sheet's drawing. Fails with `kInvalidArgument` for a bad sheet,
/// an absolute anchor kind, a negative or 2^31-EMU-or-larger offset or size,
/// an anchor outside the sheet, an image over `kMaxImageBytes`, or a package
/// that would exceed the reader's part or size limits.
Expected<std::uint32_t, Error> insert_image(Workbook& wb, std::size_t sheet, const std::uint8_t* bytes, std::size_t len,
                                            const ImageInsertOptions& options);

/// Removes the picture with `object_id`, plus its image relationship and
/// media part once nothing else references them. Fails with
/// `kInvalidArgument` when the sheet has no picture with that id.
Expected<void, Error> remove_image(Workbook& wb, std::size_t sheet, std::uint32_t object_id);

/// Moves sheet `sheet`'s anchors for a row (`row_axis`) or column edit of
/// `count` lines at `index`, after the sheet's own cells and layout moved. A
/// drawing that does not parse is left untouched.
void shift_drawing_anchors(Workbook& wb, std::size_t sheet, std::uint32_t index, std::uint32_t count, bool is_delete,
                           bool row_axis);

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_EDIT_H_
