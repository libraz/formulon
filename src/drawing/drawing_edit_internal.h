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

class Workbook;

namespace drawing_edit {

inline Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
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

inline constexpr const char* kNoObject = "drawing: no picture or chart with that id";

// Loads sheet `sheet`'s editable drawing into `doc` and returns its path; a
// sheet with no drawing holds no object.
Expected<std::string, Error> load_object_drawing(const Workbook& wb, std::size_t sheet, pugi::xml_document& doc);

// The id for a new object: one past the largest `cNvPr/@id` in `doc`, and at
// least 2 (Excel numbers drawing objects from 2). `kPreconditionFailed` once
// the largest id is UINT32_MAX.
Expected<std::uint32_t, Error> next_object_id(const pugi::xml_document& doc);

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

// The kinds the drawing edits move, snapshot and remove.
inline bool is_editable(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kPicture || kind == DrawingObjectKind::kChart;
}

// The first listed picture or chart anchor with `object_id`, or a null node.
pugi::xml_node find_object(const pugi::xml_node& root, std::uint32_t object_id);

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
