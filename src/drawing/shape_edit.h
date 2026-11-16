//
// Typed access to the drawing shapes of a sheet: `xdr:sp` (shapes and text
// boxes), `xdr:cxnSp` (lines and connectors) and `xdr:grpSp` (groups).
//
// Like the picture and chart edits, every edit parses the sheet's drawing
// part, patches the DOM and re-serialises it. An update rewrites only the
// fields `ShapeSpec::fields` names; the rest of the shape's XML (geometry
// guides, effects, style references, hyperlinks, text formatting, children of
// a group) is kept. A shape the model cannot express is listed with its
// unrepresentable fields reported as empty and passes through verbatim until
// it is updated or removed.
//
// Only top-level anchors are listed; the children of a group are not. A shape
// inside an `mc:AlternateContent` is read and updated through its first
// `mc:Choice` only. Edits target worksheets (see `drawing_edit.h`).

#ifndef FORMULON_DRAWING_SHAPE_EDIT_H_
#define FORMULON_DRAWING_SHAPE_EDIT_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "chart/chart_model.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// A listed shape. `object` carries id, kind (shape, connector or group),
/// anchor, size, name, description and hidden flag as `list_drawing_objects`
/// reports them.
/// One end of a connector glued to a shape (`a:stCxn` / `a:endCxn`).
struct ShapeConnection {
  std::uint32_t object_id = 0;   ///< `@id`: the `cNvPr/@id` of the shape glued to.
  std::uint32_t site_index = 0;  ///< `@idx`: the connection site on that shape, from 0.
  bool operator==(const ShapeConnection& o) const { return object_id == o.object_id && site_index == o.site_index; }
};

struct ShapeInfo {
  DrawingObject object;
  std::string prst;           ///< `a:prstGeom/@prst`; empty for custom geometry or a group.
  std::string text;           ///< Paragraphs joined with `\n`; empty for connectors and groups.
  bool text_box = false;      ///< `cNvSpPr/@txBox`.
  std::int32_t rotation = 0;  ///< `a:xfrm/@rot` in 60000ths of a degree, as stored.
  bool flip_h = false;        ///< `a:xfrm/@flipH`.
  bool flip_v = false;        ///< `a:xfrm/@flipV`.
  /// The `spPr` fill and outline as the chart paint model reads them; `other`
  /// marks a gradient, pattern or picture paint the record cannot state.
  ChartFill paint_fill;
  ChartLine paint_line;
  /// A connector's glued start and end; unset when that end is free or the object is not a connector.
  std::optional<ShapeConnection> start_connection;
  std::optional<ShapeConnection> end_connection;
};

/// `ShapeSpec::fields` bits: the fields an insert or update applies.
namespace shape_field {
inline constexpr std::uint32_t kName = 1U << 0;
inline constexpr std::uint32_t kPrst = 1U << 1;
inline constexpr std::uint32_t kAnchor = 1U << 2;
inline constexpr std::uint32_t kText = 1U << 3;
inline constexpr std::uint32_t kTextBox = 1U << 4;
inline constexpr std::uint32_t kRotation = 1U << 5;
inline constexpr std::uint32_t kHidden = 1U << 6;
inline constexpr std::uint32_t kFlipH = 1U << 7;
inline constexpr std::uint32_t kFlipV = 1U << 8;
inline constexpr std::uint32_t kAll = (1U << 9) - 1U;
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
  std::int32_t rotation = 0;  ///< 60000ths of a degree, stored as given; 0 removes `rot`.
  bool flip_h = false;        ///< `kFlipH`; false removes the attribute.
  bool flip_v = false;        ///< `kFlipV`.
  bool hidden = false;
  /// Fill and outline edits in the chart paint model, applied whatever
  /// `fields` says; a Clear field returns that part to the style's paint.
  ChartFillPatch paint_fill;
  ChartLinePatch paint_line;
};

/// The shapes, connectors and groups of sheet `sheet`'s drawing in z-order
/// (back first). A sheet without a drawing yields an empty list. Fails like
/// `list_drawing_objects`.
Expected<std::vector<ShapeInfo>, Error> list_shapes(const Workbook& wb, std::size_t sheet);

/// Appends a shape on top and returns its `cNvPr` id. `kAnchor` is required;
/// an unset or empty `prst` is `rect`, the name defaults to `Shape <id>`. A
/// `prst` naming a line or connector geometry makes an `xdr:cxnSp`, which
/// carries no text. A sheet without a drawing gets one. Fails with
/// `kInvalidArgument` for a bad sheet, a missing anchor, a `prst` that is not
/// an ECMA-376 `ST_ShapeType` name, a paint patch `write_paint_properties`
/// refuses, an offset or size outside [0, 2^31) EMU, a cell outside the
/// sheet, a two-cell `to` before `from`, a fill or outline that is not solid,
/// none or unset, or text with control characters.
Expected<std::uint32_t, Error> insert_shape(Workbook& wb, std::size_t sheet, const ShapeSpec& spec);

/// Applies the fields `spec.fields` names to the shape, connector or group
/// with `object_id`. An anchor of another kind converts the anchor. Text applies to a shape
/// only and keeps the first paragraph's properties and first run's
/// formatting. Fails like `insert_shape`, with `kNotFound` for an unknown id,
/// and with `kInvalidArgument` for an object of another kind.
Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id, const ShapeSpec& spec);

/// Applies `common` (the `cNvPr` fields and anchor, as `update_drawing`) and
/// `shape` (as the spec overload) to the object with `object_id` as one edit:
/// both halves are checked before anything is written, so nothing changes on
/// failure. Either pointer may be null; both null is a no-op.
Expected<void, Error> update_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                   const DrawingCommonPatch* common, const ShapeSpec* shape);

/// Removes the shape, connector or group (with its children) with `object_id`,
/// the relationships only it named, and a drawing part left empty with the
/// sheet's link to it. Fails as `update_shape` for an unknown id or an object
/// of another kind.
Expected<void, Error> remove_shape(Workbook& wb, std::size_t sheet, std::uint32_t object_id);

}  // namespace formulon

#endif  // FORMULON_DRAWING_SHAPE_EDIT_H_
