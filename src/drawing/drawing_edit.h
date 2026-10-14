//
// Typed drawing edits: insert images and graphic frames, and place, reorder,
// snapshot, restore and remove pictures and charts in a sheet's drawing part.
//
// Every edit parses the drawing part, patches the DOM and re-serialises it,
// so charts, shapes and anything else the part holds stay intact; a part no
// edit touched keeps its bytes. Inserting into a sheet without a drawing
// creates `xl/drawings/drawingN.xml`, its rels part and the sheet's drawing
// relationship. Media parts are `Default`-typed by extension. A chart's part
// and every part reachable from its rels travel with it through snapshot,
// restore and remove.
//
// Refusals: a drawing part (or its rels part) that does not parse fails with
// `kIoDrawingUnparseable`, as does a sheet whose drawing was retained from an
// `.xlsb` package (reachable only through `Sheet::unknown_relationships()`);
// bytes that are not a recognised image fail with `kIoImageUnsupported`.
// Edits target worksheets; a chartsheet's drawing is listed but not edited.
// Shapes, groups, connectors and other frames are refused with
// `kInvalidArgument`.
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
#include <string_view>
#include <vector>

#include "drawing/drawing_xml.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Placement of an inserted image or graphic frame.
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

/// Placement `set_drawing_anchor` applies, in the quantities
/// `list_drawing_objects` reports: the anchor box's top-left marker and the
/// object's unrotated size.
struct DrawingAnchor {
  AnchorKind anchor_kind = AnchorKind::kOneCell;  ///< `kOneCell` or `kTwoCell`.
  EditAs edit_as = EditAs::kTwoCell;              ///< Honoured for `kTwoCell` only.
  std::uint32_t row = 0;                          ///< 0-based top-left cell of the anchor box.
  std::uint32_t col = 0;
  std::int64_t row_off = 0;  ///< EMU offset inside the cell.
  std::int64_t col_off = 0;
  std::int64_t width_emu = 0;   ///< 0 keeps the current width.
  std::int64_t height_emu = 0;  ///< 0 keeps the current height.
};

/// `restore_drawing` flag: restore as a copy with a fresh id, appended on top.
inline constexpr std::uint32_t kDrawingRestoreNewId = 1U;

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

/// Appends an `xdr:graphicFrame` holding `graphic_data_xml` (an
/// `a:graphicData` element declaring every prefix it uses but `a`) on top of
/// sheet `sheet`'s drawing, with a new relationship of `rel_type` to the part
/// at `target_part_path`; each `{rId}` in the fragment becomes its id. The
/// caller adds the target part. Returns the frame's `cNvPr` id; an empty
/// `name` yields `Chart <id>`. Fails like `insert_image`, and with
/// `kInvalidArgument` for a zero size or a fragment that does not parse.
Expected<std::uint32_t, Error> insert_graphic_frame(Workbook& wb, std::size_t sheet,
                                                    const ImageInsertOptions& placement,
                                                    std::string_view graphic_data_xml,
                                                    std::string_view target_part_path, std::string_view rel_type);

/// Distance in EMU from the sheet's top (`row_axis`) or left edge to `line`
/// under the display geometry anchors are laid out with. `sheet` must exist.
std::int64_t sheet_offset_emu(const Workbook& wb, std::size_t sheet, bool row_axis, std::uint32_t line);

/// Moves and resizes the first picture or chart with `object_id`, rewriting
/// only its anchor markers, `editAs` when its value changes, and the
/// `a:off`/`a:ext` of its `a:xfrm` (a chart's `xdr:xfrm`); every other
/// attribute and child is kept, and both branches of an `mc:AlternateContent`
/// move together. A two-cell chart's current size is its anchor box. A
/// picture rotated into [45, 135) or [225, 315) degrees gets an anchor box
/// with width and height swapped about the same centre. A placement equal to
/// the current one writes nothing. Fails like `insert_image` for bad input,
/// and with `kInvalidArgument` when the sheet has no picture or chart with
/// that id.
Expected<void, Error> set_drawing_anchor(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                         const DrawingAnchor& anchor);

/// Moves the first picture or chart with `object_id` so it becomes item
/// `index` of `list_drawing_objects`, whose order is the z-order (0 is the
/// back). The unit moved is the object's top-level element, a whole
/// `mc:AlternateContent` included. Fails with `kInvalidArgument` for an
/// unknown id or an `index` past the last item; an unchanged order writes
/// nothing.
Expected<void, Error> set_drawing_z_order(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                          std::uint32_t index);

/// Captures the first picture or chart with `object_id` as opaque versioned
/// bytes: its list position, its top-level element, the relationships it
/// names with their parts, and every part reachable from those parts' rels
/// (a chart's style, colors, user shapes and embeddings) with the rels parts.
/// Fails with `kInvalidArgument` for an unknown id.
Expected<std::vector<std::uint8_t>, Error> snapshot_drawing(const Workbook& wb, std::size_t sheet,
                                                            std::uint32_t object_id);

/// Puts a `snapshot_drawing` capture (current or earlier version) back and
/// returns the object's id. An object with the same id is replaced, the
/// parts it reaches are rewritten in place, and the capture returns to its
/// recorded list position; with `kDrawingRestoreNewId` it is added as a copy
/// with a fresh id (and no `a16:creationId`) on top, its chart parts copied
/// to fresh paths. A part whose path is free is written there; a media part
/// with identical bytes is shared; any other occupied path takes a fresh one
/// and the rels naming it are re-pointed. Fails with `kInvalidArgument` for
/// malformed bytes, unknown flags or an id held by another kind of object,
/// and with the package limits of `insert_image`.
Expected<std::uint32_t, Error> restore_drawing(Workbook& wb, std::size_t sheet, const std::uint8_t* bytes,
                                               std::size_t len, std::uint32_t flags);

/// Removes the picture or chart with `object_id`, plus the relationships it
/// named and the parts (a chart's part and its dependents) nothing else
/// references. Fails with `kInvalidArgument` when the sheet has no picture or
/// chart with that id.
Expected<void, Error> remove_drawing(Workbook& wb, std::size_t sheet, std::uint32_t object_id);

/// Moves sheet `sheet`'s anchors for a row (`row_axis`) or column edit of
/// `count` lines at `index`, after the sheet's own cells and layout moved. A
/// drawing that does not parse is left untouched.
void shift_drawing_anchors(Workbook& wb, std::size_t sheet, std::uint32_t index, std::uint32_t count, bool is_delete,
                           bool row_axis);

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_EDIT_H_
