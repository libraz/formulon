#ifndef FORMULON_DRAWING_DRAWING_EDIT_INTERNAL_H_
#define FORMULON_DRAWING_DRAWING_EDIT_INTERNAL_H_

// Drawing-part editing primitives shared by `drawing_edit.cpp` and `drawing_snapshot.cpp`.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "io/ooxml/part_graph.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;
class Workbook;

namespace drawing_edit {

inline Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

inline Error not_found(const char* message) {
  return make_error(FormulonErrorCode::kNotFound, message);
}

inline Error unparseable(const char* message, std::string_view detail) {
  return make_error(FormulonErrorCode::kIoDrawingUnparseable, message, std::string(detail));
}

template <typename Fn>
void for_each_element(const pugi::xml_node& node, const Fn& fn) {
  node.find_node([&fn](const pugi::xml_node& child) {
    if (child.type() == pugi::node_element) {
      fn(child);
    }
    return false;
  });
}

inline constexpr const char* kNoObject = "object_id does not name a picture or chart";
inline constexpr const char* kNoShapeOrObject = "object_id does not name a picture, chart or shape";
inline constexpr const char* kNoOrderable = "object_id does not name a picture, chart, shape or slicer";

// The refusal of `object_id` by an edit whose kinds exclude it:
// `kInvalidArgument` when an object of another kind holds the id under `root`,
// `kNotFound` when none does.
Error refuse_object(const pugi::xml_node& root, std::uint32_t object_id, const char* message);

// Loads sheet `sheet`'s editable drawing into `doc` and returns its path; a
// sheet with no drawing holds no object (`kNotFound`).
Expected<std::string, Error> load_object_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc);

// The id for a new object on `sheet`: the first id not used by a `cNvPr` in
// `doc` at or above the sheet's drawing counter, which then moves past it.
// `kPreconditionFailed` once the id space is exhausted.
Expected<std::uint32_t, Error> allocate_drawing_id(Sheet& sheet, const pugi::xml_document& doc);

// "<word> <id - 1>" with the workbook locale's word for `kind` (chart or
// picture); empty for any other kind.
std::string default_drawing_name(const Workbook& wb, DrawingObjectKind kind, std::uint32_t id);

// Whether `name` is "<word> <digits>" with `kind`'s word in the workbook's
// locale or in English.
bool is_default_drawing_name(const Workbook& wb, DrawingObjectKind kind, std::string_view name);

// Values of the prefixed, non-`xmlns:` attributes of `node` and its subtree:
// the relationship ids it may name.
std::vector<std::string> relationship_refs(const pugi::xml_node& node);

// First `<stem><N>.<suffix>` with no part, and no path in `reserved`, named
// `<stem><N>.*`.
std::string unused_part_path(const Workbook& wb, const std::string& stem, const char* suffix,
                             const std::vector<std::string>& reserved = {});

Expected<void, Error> prepare_drawing(const Workbook& wb, std::string& path, pugi::xml_document& doc);

// Removes the top-level drawing elements `tops`, then the relationships no
// remaining element names and the parts they leave unreferenced.
Expected<void, Error> drop_objects(Workbook& wb, const std::string& path, pugi::xml_document& doc,
                                   const std::vector<pugi::xml_node>& tops);

// Creates an empty part at `path` so `store_part_dom` can fill it.
Expected<void, Error> add_empty_part(Workbook& wb, const std::string& path, const char* content_type);

void ensure_default_content_type(Workbook& wb, const char* extension, const char* content_type);

// The `xdr:wsDr` child holding `node`.
pugi::xml_node top_of(const pugi::xml_node& root, pugi::xml_node node);

// The kinds the drawing edits move and remove.
inline bool is_editable(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kPicture || kind == DrawingObjectKind::kChart;
}

// Writes `cNvPr/@hidden` as Excel does: `1` when hidden, absent otherwise.
inline void set_cnvpr_hidden(pugi::xml_node cnvpr, bool hidden) {
  if (hidden) {
    (cnvpr.attribute("hidden") ? cnvpr.attribute("hidden") : cnvpr.append_attribute("hidden")).set_value("1");
  } else {
    cnvpr.remove_attribute("hidden");
  }
}

// The kinds the shape edits address: a shape, a connector or a group.
inline bool is_shape_kind(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kShape || kind == DrawingObjectKind::kConnector ||
         kind == DrawingObjectKind::kGroup;
}

// The kinds a rename accepts.
inline bool is_nameable(DrawingObjectKind kind) {
  return is_editable(kind) || kind == DrawingObjectKind::kShape;
}

// The kinds a snapshot captures and a restore places.
inline bool is_snapshot_kind(DrawingObjectKind kind) {
  return is_editable(kind) || kind == DrawingObjectKind::kShape;
}

// The objects a z-order move accepts: the nameable kinds and slicer frames.
inline bool is_orderable(const pugi::xml_node& anchor, DrawingObjectKind kind) {
  return is_nameable(kind) || is_slicer_frame(anchor);
}

// Where a top-level element moved to list position `index` goes: before the
// element holding that item once `moving` is out of the list, or at the end
// (a null node).
pugi::xml_node insertion_point(const pugi::xml_node& root, const pugi::xml_node& moving, std::size_t index);

std::string next_rel_id(const std::vector<io::ooxml::DrawingRel>& rels);

Expected<void, Error> load_rels_doc(const Workbook& wb, const std::string& rels_path, pugi::xml_document& rels_doc);

void append_rel(pugi::xml_document& rels_doc, const std::string& id, std::string_view type, const std::string& target,
                bool external);

}  // namespace drawing_edit
}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_EDIT_INTERNAL_H_
