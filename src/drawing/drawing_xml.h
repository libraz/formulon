//
// DrawingML part model: anchors and objects of a sheet's drawing part.
//
// The drawing part's bytes stay the source of truth: this module reads a
// parsed `xdr:wsDr` document into a typed list of objects but never rebuilds
// the part from that list, so everything it does not model (shape geometry,
// effects, extension lists, unknown elements) survives an edit made through
// the DOM. Element names are matched by local name, so a part that binds the
// SpreadsheetDrawing or DrawingML namespaces to other prefixes reads the same.
//
// A top-level `mc:AlternateContent` contributes the anchors of its first
// `mc:Choice`; its `mc:Fallback` is a rendering of the same object.

#ifndef FORMULON_DRAWING_DRAWING_XML_H_
#define FORMULON_DRAWING_DRAWING_XML_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "drawing/image_header.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// What an anchor holds.
enum class DrawingObjectKind : std::int32_t {
  kPicture = 0,       ///< `xdr:pic`.
  kShape = 1,         ///< `xdr:sp`.
  kChart = 2,         ///< `xdr:graphicFrame` holding a chart or chartex.
  kGroup = 3,         ///< `xdr:grpSp`.
  kConnector = 4,     ///< `xdr:cxnSp`.
  kGraphicFrame = 5,  ///< Any other `xdr:graphicFrame` (SmartArt, slicer, ...).
  kOther = 6,         ///< Content part or an unrecognised element.
};

/// The anchor element type.
enum class AnchorKind : std::int32_t {
  kOneCell = 0,   ///< `xdr:oneCellAnchor` (from + ext).
  kTwoCell = 1,   ///< `xdr:twoCellAnchor` (from + to).
  kAbsolute = 2,  ///< `xdr:absoluteAnchor` (pos + ext).
};

/// How the object follows row/column edits. A two-cell anchor carries it in
/// `editAs` (absent means `kTwoCell`); a one-cell anchor behaves as
/// `kOneCell` and an absolute anchor as `kAbsolute`.
enum class EditAs : std::int32_t {
  kTwoCell = 0,
  kOneCell = 1,
  kAbsolute = 2,
};

/// One `xdr:from` / `xdr:to` marker: a 0-based cell plus EMU offsets.
struct AnchorPoint {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::int64_t row_off = 0;
  std::int64_t col_off = 0;
};

/// One anchored object. An absolute anchor reports its `xdr:pos` as offsets
/// from cell A1 in `from` (`col_off` = x, `row_off` = y).
struct DrawingObject {
  std::uint32_t object_id = 0;  ///< `cNvPr/@id`.
  DrawingObjectKind kind = DrawingObjectKind::kOther;
  AnchorKind anchor_kind = AnchorKind::kTwoCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;       ///< Two-cell anchors only.
  std::int64_t cx = 0;  ///< Width in EMU (`xdr:ext`, else the object's `a:xfrm/a:ext`).
  std::int64_t cy = 0;
  std::string name;
  std::string descr;
  std::string image_rel_id;                          ///< Picture `a:blip/@r:embed`.
  std::string media_path;                            ///< Package path the blip resolves to, or empty.
  ImageFormat image_format = ImageFormat::kUnknown;  ///< Filled by callers that hold the media bytes.
};

/// One relationship of a drawing part's rels part. `target` is the resolved
/// package path for an internal target and the raw value for an external one.
struct DrawingRel {
  std::string id;
  std::string type;
  std::string target;
  bool external = false;
};

/// `Target=` value reaching package path `target` from a part in directory
/// `from_dir` (no trailing slash; `xl/drawings` + `xl/media/image1.png` ->
/// `../media/image1.png`).
std::string relative_target(std::string_view from_dir, std::string_view target);

/// Parses drawing part bytes. Fails with `kIoDrawingUnparseable` when they
/// are not well-formed XML or the root is not `wsDr`.
Expected<void, Error> parse_drawing_part(const std::vector<std::uint8_t>& bytes, pugi::xml_document& doc);

/// Parses the rels part of a part in directory `owner_dir` (no trailing slash). Fails with
/// `kIoDrawingUnparseable` on malformed XML or a target that cannot resolve.
Expected<std::vector<DrawingRel>, Error> parse_part_rels(const std::vector<std::uint8_t>& bytes,
                                                         std::string_view owner_dir);

/// Local name of `node` (the part after any `prefix:`).
std::string_view local_name(const pugi::xml_node& node);

/// First element child of `parent` whose local name is `name`.
pugi::xml_node child_local(const pugi::xml_node& parent, std::string_view name);

/// The anchor elements of `root` in document order. `include_fallback`
/// adds the `mc:Fallback` anchors of `mc:AlternateContent` wrappers, which a
/// row/column edit must move along with the `mc:Choice` ones.
std::vector<pugi::xml_node> drawing_anchors(const pugi::xml_node& root, bool include_fallback);

/// The `cNvPr` element naming the object `anchor` holds, or a null node.
pugi::xml_node anchor_cnvpr(const pugi::xml_node& anchor);

/// Reads one anchor element.
DrawingObject read_drawing_object(const pugi::xml_node& anchor, const std::vector<DrawingRel>& rels);

/// Reads every anchored object of a parsed drawing part.
std::vector<DrawingObject> read_drawing_objects(const pugi::xml_document& doc, const std::vector<DrawingRel>& rels);

/// Reads an `xdr:from` / `xdr:to` marker.
AnchorPoint read_anchor_point(const pugi::xml_node& marker);

/// Overwrites the four children of an `xdr:from` / `xdr:to` marker.
void write_anchor_point(const pugi::xml_node& marker, const AnchorPoint& point);

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_XML_H_
