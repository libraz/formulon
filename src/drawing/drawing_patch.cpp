//
// Field patches on existing drawing objects: anchor merges, common cNvPr
// fields routed per kind, and picture transforms.

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "drawing/connector_route.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_edit_internal.h"
#include "drawing/drawing_geometry.h"
#include "drawing/shape_edit.h"
#include "drawing/slicer_edit.h"
#include "io/ooxml/part_dom.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/status_macros.h"
#include "workbook.h"

namespace formulon {

using namespace drawing_edit;

namespace {

constexpr const char* kNoDrawingObject = "object_id does not name a picture, chart, shape or slicer";

enum class DrawingRoute { kFrame, kShape, kSlicer };

bool any_anchor_field(const DrawingAnchorPatch& p) {
  return p.kind || p.edit_as || p.from_row || p.from_col || p.from_row_off || p.from_col_off || p.to_row || p.to_col ||
         p.to_row_off || p.to_col_off || p.width_emu || p.height_emu;
}

// `patch` merged into the anchor of `obj` (whose box is `box_cx` x `box_cy`).
// `to` is meaningful for a two-cell result, `cx`/`cy` for any other.
struct MergedAnchor {
  AnchorKind kind = AnchorKind::kTwoCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;
  std::int64_t cx = 0;
  std::int64_t cy = 0;
};

Expected<MergedAnchor, Error> merge_anchor(const Geometry& geo, const DrawingObject& obj, std::int64_t box_cx,
                                           std::int64_t box_cy, const DrawingAnchorPatch& p) {
  MergedAnchor m;
  m.kind = p.kind.value_or(obj.anchor_kind);
  m.edit_as = p.edit_as.value_or(obj.edit_as);
  if (m.kind == AnchorKind::kAbsolute && obj.anchor_kind != AnchorKind::kAbsolute) {
    return invalid(kNotAbsolute);
  }
  const bool two = m.kind == AnchorKind::kTwoCell;
  const bool has_to = p.to_row || p.to_col || p.to_row_off || p.to_col_off;
  const bool has_size = p.width_emu || p.height_emu;
  // A read anchor carries both the end marker and the size: the field the resulting kind holds wins.
  if (two && has_size && !has_to) {
    return invalid("anchor.width_emu and anchor.height_emu do not apply to a two-cell anchor");
  }
  if (!two && has_to && !has_size) {
    return invalid("anchor.to_row and anchor.to_col apply only to a two-cell anchor");
  }
  m.from = obj.from;
  m.from.row = p.from_row.value_or(m.from.row);
  m.from.col = p.from_col.value_or(m.from.col);
  m.from.row_off = p.from_row_off.value_or(m.from.row_off);
  m.from.col_off = p.from_col_off.value_or(m.from.col_off);
  if (m.kind == AnchorKind::kAbsolute && (m.from.row != 0 || m.from.col != 0)) {
    return invalid(kAbsoluteCells);
  }
  if (!two) {
    m.cx = p.width_emu.value_or(box_cx);
    m.cy = p.height_emu.value_or(box_cy);
    return m;
  }
  if (obj.anchor_kind == AnchorKind::kTwoCell) {
    m.to = obj.to;
  } else {
    // A converted anchor ends where the current size reaches from the new start.
    m.to = m.from;
    if (m.from.row >= Sheet::kMaxRows || m.from.col >= Sheet::kMaxCols ||
        !geo.advance(true, m.to.row, m.to.row_off, box_cy) || !geo.advance(false, m.to.col, m.to.col_off, box_cx)) {
      return invalid("anchor must lie within the sheet");
    }
  }
  m.to.row = p.to_row.value_or(m.to.row);
  m.to.col = p.to_col.value_or(m.to.col);
  m.to.row_off = p.to_row_off.value_or(m.to.row_off);
  m.to.col_off = p.to_col_off.value_or(m.to.col_off);
  if (!cell_in_range(m.to.row, m.to.col)) {
    return invalid("anchor offsets, size and cells must be in range");
  }
  m.cx = geo.span(false, m.from, m.to);
  m.cy = geo.span(true, m.from, m.to);
  if (m.cx < 0 || m.cy < 0) {
    return invalid("anchor.to_row and anchor.to_col must not precede anchor.from_row and anchor.from_col");
  }
  return m;
}

// Routes a merged anchor through the edit of the object's kind.
Expected<void, Error> apply_merged_anchor(Workbook& wb, std::size_t sheet, std::uint32_t object_id, DrawingRoute route,
                                          const MergedAnchor& m, bool swapped) {
  switch (route) {
    case DrawingRoute::kFrame: {
      DrawingAnchor a;
      a.anchor_kind = m.kind;
      a.edit_as = m.edit_as;
      a.row = m.from.row;
      a.col = m.from.col;
      a.row_off = m.from.row_off;
      a.col_off = m.from.col_off;
      // `set_drawing_anchor` takes the unrotated size of a picture laid out in a swapped box.
      a.width_emu = swapped ? m.cy : m.cx;
      a.height_emu = swapped ? m.cx : m.cy;
      if (a.width_emu == 0 || a.height_emu == 0) {
        return invalid("anchor.width_emu and anchor.height_emu must be non-zero");
      }
      return set_drawing_anchor(wb, sheet, object_id, a);
    }
    case DrawingRoute::kShape: {
      ShapeSpec spec;
      spec.fields = shape_field::kAnchor;
      spec.anchor_kind = m.kind;
      spec.edit_as = m.edit_as;
      spec.from = m.from;
      spec.to = m.to;
      spec.cx = m.cx;
      spec.cy = m.cy;
      return update_shape(wb, sheet, object_id, spec);
    }
    case DrawingRoute::kSlicer: {
      if (m.kind != AnchorKind::kTwoCell) {
        return invalid("anchor.kind must be a two-cell anchor for a slicer");
      }
      ASSIGN_OR_RETURN(std::vector<SlicerInfo> slicers, list_slicers(wb, sheet));
      if (const std::optional<std::size_t> index = slicer_index_of(slicers, object_id)) {
        SlicerSpec spec;
        spec.fields = slicer_field::kAnchor;
        spec.from = m.from;
        spec.to = m.to;
        return update_slicer(wb, sheet, *index, spec);
      }
      return not_found(kNoDrawingObject);
    }
  }
  return invalid(kNoDrawingObject);
}

// Every `cNvPr` of `top` carrying `object_id`: both branches of an `mc:AlternateContent`.
std::vector<pugi::xml_node> object_cnvprs(const pugi::xml_node& top, std::uint32_t object_id) {
  std::vector<pugi::xml_node> out;
  for_each_element(top, [&](const pugi::xml_node& node) {
    if (local_name(node) == "cNvPr" && node.attribute("id").as_uint() == object_id) {
      out.push_back(node);
    }
  });
  return out;
}

// Applies the `cNvPr` fields of `patch`; true when anything changed.
bool apply_cnvpr_fields(pugi::xml_node cnvpr, const DrawingCommonPatch& patch) {
  bool changed = false;
  if (patch.name && std::string_view(cnvpr.attribute("name").value()) != *patch.name) {
    io::set_attr(cnvpr, "name", *patch.name);
    changed = true;
  }
  if (patch.description && std::string_view(cnvpr.attribute("descr").value()) != *patch.description) {
    if (patch.description->empty()) {
      cnvpr.remove_attribute("descr");
    } else {
      io::set_attr(cnvpr, "descr", *patch.description);
    }
    changed = true;
  }
  if (patch.hidden && cnvpr.attribute("hidden").as_bool() != *patch.hidden) {
    set_cnvpr_hidden(cnvpr, *patch.hidden);
    changed = true;
  }
  return changed;
}

}  // namespace

Expected<void, Error> update_drawing(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                     const DrawingCommonPatch& patch) {
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  pugi::xml_node root = doc.document_element();
  DrawingRoute route = DrawingRoute::kFrame;
  // A slicer frame is checked before the shapes: its `mc:Fallback` holds a shape with the same id.
  pugi::xml_node target = find_drawing_object(root, object_id, is_slicer_object);
  if (target) {
    route = DrawingRoute::kSlicer;
  } else if ((target = find_drawing_object(root, object_id, is_editable))) {
    route = DrawingRoute::kFrame;
  } else if ((target = find_drawing_object(root, object_id, is_shape_kind))) {
    route = DrawingRoute::kShape;
  } else {
    return refuse_object(root, object_id, kNoDrawingObject);
  }
  if (patch.name && route == DrawingRoute::kSlicer) {
    return make_error(FormulonErrorCode::kPreconditionFailed, "name cannot be changed on a slicer");
  }
  if (patch.name && patch.name->empty()) {
    return invalid("name must not be empty");
  }

  std::optional<MergedAnchor> merged;
  bool swapped = false;
  if (any_anchor_field(patch.anchor)) {
    const Geometry geo(wb, wb.sheet(sheet));
    const DrawingObject obj = read_drawing_object(target, {});
    swapped =
        route == DrawingRoute::kFrame && obj.kind == DrawingObjectKind::kPicture && swaps_box(object_xfrm(target));
    std::int64_t box_cx = swapped ? obj.cy : obj.cx;
    std::int64_t box_cy = swapped ? obj.cx : obj.cy;
    if (obj.anchor_kind == AnchorKind::kTwoCell) {
      box_cx = geo.span(false, obj.from, obj.to);
      box_cy = geo.span(true, obj.from, obj.to);
    }
    ASSIGN_OR_RETURN(MergedAnchor m, merge_anchor(geo, obj, box_cx, box_cy, patch.anchor));
    merged = m;
  }

  // The anchor edit stores the part itself; keep the bytes so a later failure leaves the drawing as it was.
  std::vector<std::uint8_t> before;
  if (const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path)) {
    before = part->bytes;
  }
  if (merged) {
    RETURN_IF_ERROR(apply_merged_anchor(wb, sheet, object_id, route, *merged, swapped));
    RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
    root = doc.document_element();
    const auto any_kind = [](DrawingObjectKind) { return true; };
    target = find_drawing_object(root, object_id, +any_kind);
    if (!target) {
      (void)wb.replace_passthrough_part(path, std::move(before));
      return not_found(kNoDrawingObject);
    }
  }
  bool changed = false;
  for (pugi::xml_node cnvpr : object_cnvprs(top_of(root, target), object_id)) {
    changed = apply_cnvpr_fields(cnvpr, patch) || changed;
  }
  if (!changed) {
    return Expected<void, Error>::Ok();
  }
  if (auto stored = io::ooxml::store_part_dom(wb, path, doc); !stored) {
    if (merged) {
      (void)wb.replace_passthrough_part(path, std::move(before));
    }
    return stored;
  }
  return Expected<void, Error>::Ok();
}

namespace {

bool is_picture_kind(DrawingObjectKind kind) {
  return kind == DrawingObjectKind::kPicture;
}

constexpr const char* kNoPicture = "object_id does not name a picture";

}  // namespace

Expected<std::optional<PictureTransform>, Error> get_picture_transform(const Workbook& wb, std::size_t sheet,
                                                                       std::uint32_t object_id) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  if (path.empty()) {
    return std::optional<PictureTransform>{};
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_drawing_part(wb, path, doc));
  const pugi::xml_node target = find_drawing_object(doc.document_element(), object_id, is_picture_kind);
  if (!target) {
    return std::optional<PictureTransform>{};
  }
  const XfrmTransform t = read_xfrm_transform(object_xfrm(target));
  return std::optional<PictureTransform>{PictureTransform{t.rot, t.flip_h, t.flip_v}};
}

Expected<void, Error> update_picture(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                     const DrawingCommonPatch* common, const PictureTransformPatch* picture) {
  const bool has_picture = picture != nullptr && (picture->rotation || picture->flip_h || picture->flip_v);
  if (!has_picture) {
    return common != nullptr ? update_drawing(wb, sheet, object_id, *common) : Expected<void, Error>::Ok();
  }
  pugi::xml_document doc;
  ASSIGN_OR_RETURN(std::string path, load_object_drawing(wb, sheet, doc));
  const pugi::xml_node root = doc.document_element();
  const pugi::xml_node target = find_drawing_object(root, object_id, is_picture_kind);
  if (!target) {
    return refuse_object(root, object_id, kNoPicture);
  }
  pugi::xml_node props = child_local(anchor_content(target), "spPr");
  if (!props) {
    return invalid("The picture has no shape properties");
  }
  pugi::xml_node xfrm = child_local(props, "xfrm");
  const XfrmTransform current = read_xfrm_transform(xfrm);
  XfrmTransform wanted = current;
  wanted.rot = picture->rotation.value_or(current.rot);
  wanted.flip_h = picture->flip_h.value_or(current.flip_h);
  wanted.flip_v = picture->flip_v.value_or(current.flip_v);
  std::vector<std::uint8_t> before;
  const bool changed = wanted != current;
  if (changed) {
    if (const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path)) {
      before = part->bytes;
    }
    if (!xfrm) {
      xfrm = props.prepend_child(
          (prefix_of(child_local(child_local(anchor_content(target), "blipFill"), "blip")) + "xfrm").c_str());
    }
    const std::vector<GluedEnd> glued = glued_ends(root);
    write_xfrm_transform(xfrm, wanted);
    follow_glued_shapes(wb, sheet, root, glued);
    RETURN_IF_ERROR(io::ooxml::store_part_dom(wb, path, doc));
  }
  if (common == nullptr) {
    return Expected<void, Error>::Ok();
  }
  auto applied = update_drawing(wb, sheet, object_id, *common);
  if (!applied && changed) {
    (void)wb.replace_passthrough_part(path, std::move(before));
  }
  return applied;
}

}  // namespace formulon
