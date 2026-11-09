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
//
// A sheet's drawing is found through the worksheet's drawing relationship,
// one retained from an `.xlsb` package, or, for a chartsheet, the drawing
// relationship in the chartsheet part's rels part.

#ifndef FORMULON_DRAWING_DRAWING_XML_H_
#define FORMULON_DRAWING_DRAWING_XML_H_

#include <cstddef>
#include <cstdint>
#include <set>
#include <string>
#include <string_view>
#include <vector>

#include "drawing/image_header.h"
#include "io/ooxml/part_graph.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

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
  AnchorPoint to;  ///< Two-cell anchors only.
  /// Width in EMU: a picture's unrotated `a:xfrm/a:ext`, else `xdr:ext`; other
  /// objects read `xdr:ext` first.
  std::int64_t cx = 0;
  std::int64_t cy = 0;
  std::string name;
  std::string descr;
  bool hidden = false;                               ///< `cNvPr/@hidden`.
  std::string image_rel_id;                          ///< Picture `a:blip/@r:embed`.
  std::string media_path;                            ///< Package path the blip resolves to, or empty.
  ImageFormat image_format = ImageFormat::kUnknown;  ///< Filled by callers that hold the media bytes.
};

/// A chart frame and the chart part it shows.
struct ChartPartRef {
  std::uint32_t object_id = 0;  ///< `cNvPr/@id` of the `graphicFrame`.
  std::string part_path;        ///< Chart part; empty when the frame's relationship does not resolve.
  bool is_chartex = false;      ///< A `cx:chart` (chartex) part rather than a `c:chart` one.
};

/// A `ChartPartRef` with the index of the sheet (worksheet or chartsheet) holding it.
struct SheetChartPartRef {
  std::size_t sheet = 0;
  ChartPartRef ref;
};

/// `Target=` value reaching package path `target` from a part in directory
/// `from_dir` (no trailing slash; `xl/drawings` + `xl/media/image1.png` ->
/// `../media/image1.png`).
std::string relative_target(std::string_view from_dir, std::string_view target);

/// Parses drawing part bytes. Fails with `kIoDrawingUnparseable` when they
/// are not well-formed XML or the root is not `wsDr`.
Expected<void, Error> parse_drawing_part(const std::vector<std::uint8_t>& bytes, pugi::xml_document& doc);

// Local-name helpers shared with the io readers.
using io::child_local;
using io::children_local;
using io::local_name;

/// The anchor elements of `root` in document order. `include_fallback`
/// adds the `mc:Fallback` anchors of `mc:AlternateContent` wrappers, which a
/// row/column edit must move along with the `mc:Choice` ones.
std::vector<pugi::xml_node> drawing_anchors(const pugi::xml_node& root, bool include_fallback);

/// The element `anchor` positions (its first child that is not a marker),
/// looking through a nested `mc:AlternateContent` to its first `mc:Choice`.
pugi::xml_node anchor_content(const pugi::xml_node& anchor);

/// The `cNvPr` element naming the object `anchor` holds, or a null node.
pugi::xml_node anchor_cnvpr(const pugi::xml_node& anchor);

/// Reads one anchor element.
DrawingObject read_drawing_object(const pugi::xml_node& anchor, const std::vector<io::ooxml::DrawingRel>& rels);

/// The first listed anchor holding an object of id `object_id` whose kind
/// `kind_ok` accepts, or a null node.
pugi::xml_node find_drawing_object(const pugi::xml_node& root, std::uint32_t object_id,
                                   bool (*kind_ok)(DrawingObjectKind));

/// As above, with `anchor_ok` also seeing the anchor.
pugi::xml_node find_drawing_object(const pugi::xml_node& root, std::uint32_t object_id,
                                   bool (*anchor_ok)(const pugi::xml_node& anchor, DrawingObjectKind kind));

/// Whether `anchor` holds a slicer: a graphic frame whose `a:graphicData`
/// uri is the slicer one.
bool is_slicer_frame(const pugi::xml_node& anchor);

/// Reads every anchored object of a parsed drawing part.
std::vector<DrawingObject> read_drawing_objects(const pugi::xml_document& doc,
                                                const std::vector<io::ooxml::DrawingRel>& rels);

/// Package path of sheet `sheet`'s drawing, or empty when it has none. Fails
/// with `kInvalidArgument` for a bad sheet and with `kIoDrawingUnparseable`
/// when a chartsheet's rels part does not parse.
Expected<std::string, Error> sheet_drawing_path(const Workbook& wb, std::size_t sheet);

/// Parses the drawing part at `path` into `doc`. Fails with
/// `kIoDrawingUnparseable` when the part is missing or does not parse.
Expected<void, Error> load_drawing_part(const Workbook& wb, std::string_view path, pugi::xml_document& doc);

/// Path of sheet `sheet`'s typed-editable drawing, or empty when it has none.
/// Fails with `kInvalidArgument` for a bad sheet or a non-worksheet, and with
/// `kIoDrawingUnparseable` for an `.xlsb` drawing the reader could not lift
/// (still a drawing relationship in `Sheet::unknown_relationships()`).
Expected<std::string, Error> editable_drawing_path(const Workbook& wb, std::size_t sheet);

/// The relationships of the part at `part_path`; none when it has no rels
/// part. Fails with `kIoDrawingUnparseable` when the rels part does not parse.
Expected<std::vector<io::ooxml::DrawingRel>, Error> load_part_rels(const Workbook& wb, std::string_view part_path);

/// The chart frames of sheet `sheet`'s drawing in list order, resolved
/// `graphicFrame` -> `c:chart` (or `cx:chart`) `r:id` -> drawing rels -> part.
/// Fails like `list_drawing_objects`.
Expected<std::vector<ChartPartRef>, Error> list_chart_parts(const Workbook& wb, std::size_t sheet);

/// `list_chart_parts` over every sheet, chartsheets included. A sheet whose
/// drawing does not parse contributes nothing.
Expected<std::vector<SheetChartPartRef>, Error> all_chart_parts(const Workbook& wb);

/// The entries of `parts` with a non-empty part path, in order, keeping the
/// first entry per path. Entries for which `skip(entry)` is true are dropped
/// before the path is recorded, so a later entry of the same part still counts.
template <class Skip>
std::vector<const SheetChartPartRef*> unique_chart_parts(const std::vector<SheetChartPartRef>& parts, Skip&& skip) {
  std::vector<const SheetChartPartRef*> out;
  std::set<std::string> seen;
  for (const SheetChartPartRef& entry : parts) {
    if (entry.ref.part_path.empty() || skip(entry) || !seen.insert(entry.ref.part_path).second) {
      continue;
    }
    out.push_back(&entry);
  }
  return out;
}

/// Reads an `xdr:from` / `xdr:to` marker.
AnchorPoint read_anchor_point(const pugi::xml_node& marker);

/// Overwrites the four children of an `xdr:from` / `xdr:to` marker.
void write_anchor_point(const pugi::xml_node& marker, const AnchorPoint& point);

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_XML_H_
