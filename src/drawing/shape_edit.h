//
// Typed access to the drawing shapes of a sheet: `xdr:sp` (shapes and text
// boxes), `xdr:cxnSp` (lines and connectors) and `xdr:grpSp` (groups).
//
// Like the picture and chart edits, every edit parses the sheet's drawing
// part, patches the DOM and re-serialises it. An update rewrites only the
// fields `ShapeSpec::fields` names; the rest of the shape's XML (geometry
// guides, effects, style references, hyperlinks, text formatting, children of
// a group) is kept. A shape the model cannot express is listed with its
// unrepresentable fields reported as `ShapePaint::kOther` or empty and passes
// through verbatim until it is updated or removed.
//
// Only top-level anchors are listed; the children of a group are not. A shape
// inside an `mc:AlternateContent` is read and updated through its first
// `mc:Choice` only. Edits target worksheets (see `drawing_edit.h`).

#ifndef FORMULON_DRAWING_SHAPE_EDIT_H_
#define FORMULON_DRAWING_SHAPE_EDIT_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "drawing/drawing_xml.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// How a shape's fill or outline is specified.
enum class ShapePaint : std::int32_t {
  kUnset = 0,  ///< No fill or outline element.
  kSolid = 1,  ///< `a:solidFill` of an `a:srgbClr`.
  kNone = 2,   ///< `a:noFill`.
  kOther = 3,  ///< Gradient, pattern, picture or theme colour.
};

/// A listed shape. `object` carries id, kind (shape, connector or group),
/// anchor, size and name as `list_drawing_objects` reports them.
struct ShapeInfo {
  DrawingObject object;
  std::string prst;       ///< `a:prstGeom/@prst`; empty for custom geometry or a group.
  std::string text;       ///< Paragraphs joined with `\n`; empty for connectors and groups.
  bool text_box = false;  ///< `cNvSpPr/@txBox`.
  ShapePaint fill = ShapePaint::kUnset;
  std::uint32_t fill_rgb = 0;  ///< 0xRRGGBB when `fill` is `kSolid`.
  ShapePaint line = ShapePaint::kUnset;
  std::uint32_t line_rgb = 0;       ///< 0xRRGGBB when `line` is `kSolid`.
  std::int64_t line_width_emu = 0;  ///< `a:ln/@w`; 0 when absent.
  std::int32_t rotation = 0;        ///< `a:xfrm/@rot` in 60000ths of a degree.
  bool hidden = false;
};

/// `ShapeSpec::fields` bits: the fields an insert or update applies.
namespace shape_field {
inline constexpr std::uint32_t kName = 1U << 0;
inline constexpr std::uint32_t kPrst = 1U << 1;
inline constexpr std::uint32_t kAnchor = 1U << 2;
inline constexpr std::uint32_t kText = 1U << 3;
inline constexpr std::uint32_t kTextBox = 1U << 4;
inline constexpr std::uint32_t kFill = 1U << 5;       ///< `fill`: `kSolid` with `fill_rgb`, or `kNone`.
inline constexpr std::uint32_t kLine = 1U << 6;       ///< `line`: `kSolid` with `line_rgb`, or `kNone`.
inline constexpr std::uint32_t kLineWidth = 1U << 7;  ///< `line_width_emu`.
inline constexpr std::uint32_t kRotation = 1U << 8;
inline constexpr std::uint32_t kHidden = 1U << 9;
inline constexpr std::uint32_t kAll = (1U << 10) - 1U;
}  // namespace shape_field

/// Values for an insert or update; fields outside `fields` are ignored.
/// The anchor reads as `DrawingObject` does: `from` and `to` markers (an
/// absolute anchor's position is `from.col_off` = x, `from.row_off` = y) and
/// `cx`/`cy` for one-cell and absolute anchors (a two-cell anchor's size is
/// its box).
struct ShapeSpec {
  std::uint32_t fields = 0;
  std::string name;
  std::string prst;
  std::string text;
  bool text_box = false;
  AnchorKind anchor_kind = AnchorKind::kTwoCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;
  std::int64_t cx = 0;
  std::int64_t cy = 0;
  ShapePaint fill = ShapePaint::kNone;
  std::uint32_t fill_rgb = 0;
  ShapePaint line = ShapePaint::kNone;
  std::uint32_t line_rgb = 0;
  std::int64_t line_width_emu = 0;
  std::int32_t rotation = 0;
  bool hidden = false;
};

/// The shapes, connectors and groups of sheet `sheet`'s drawing in z-order
/// (back first). A sheet without a drawing yields an empty list. Fails like
/// `list_drawing_objects`.
Expected<std::vector<ShapeInfo>, Error> list_shapes(const Workbook& wb, std::size_t sheet);

/// Appends a shape on top and returns its `cNvPr` id. `kAnchor` is required;
/// `prst` defaults to `rect`, the name to `Shape <id>`. A `prst` naming a line
/// or connector geometry makes an `xdr:cxnSp`, which carries no text. A sheet
/// without a drawing gets one. Fails with `kInvalidArgument` for a bad sheet,
/// a missing anchor, a `prst` that is not an alphanumeric token, an
/// offset or size outside [0, 2^31) EMU, a cell outside the sheet, a
/// two-cell `to` before `from`, a fill or outline that is neither solid nor none, or text with
/// control characters.
Expected<std::uint32_t, Error> insert_shape(Workbook& wb, std::size_t sheet, const ShapeSpec& spec);

/// Applies the fields `spec.fields` names to the shape, connector or group
/// with `object_id`. The anchor kind cannot change. Text applies to a shape
/// only and keeps the first paragraph's properties and first run's
/// formatting. Fails like `insert_shape`, and with `kInvalidArgument` for an
/// unknown id.
Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id, const ShapeSpec& spec);

/// Removes the shape, connector or group (with its children) with `object_id`.
Expected<void, Error> remove_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id);

}  // namespace formulon

#endif  // FORMULON_DRAWING_SHAPE_EDIT_H_
