#include "drawing/drawing_xml.h"

#include <cstring>
#include <string_view>
#include <unordered_map>
#include <utility>

#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "workbook.h"

namespace formulon {
namespace {

Error unparseable(std::string_view what, std::string_view detail) {
  return make_error(FormulonErrorCode::kIoDrawingUnparseable, std::string("drawing: ") + std::string(what),
                    std::string(detail));
}

bool is_anchor(const pugi::xml_node& node) {
  const std::string_view name = local_name(node);
  return name == "twoCellAnchor" || name == "oneCellAnchor" || name == "absoluteAnchor";
}

DrawingObjectKind classify(const pugi::xml_node& content) {
  const std::string_view name = local_name(content);
  if (name == "pic") {
    return DrawingObjectKind::kPicture;
  }
  if (name == "sp") {
    return DrawingObjectKind::kShape;
  }
  if (name == "grpSp") {
    return DrawingObjectKind::kGroup;
  }
  if (name == "cxnSp") {
    return DrawingObjectKind::kConnector;
  }
  if (name == "graphicFrame") {
    const std::string_view uri = child_local(child_local(content, "graphic"), "graphicData").attribute("uri").value();
    return strings::ends_with(uri, "/chart") || strings::ends_with(uri, "/chartex") ? DrawingObjectKind::kChart
                                                                                    : DrawingObjectKind::kGraphicFrame;
  }
  return DrawingObjectKind::kOther;
}

// The object's own `a:xfrm`: a child of the content (graphic frames) or of
// its shape-properties element.
pugi::xml_node content_xfrm(const pugi::xml_node& content) {
  for (pugi::xml_node child = content.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = local_name(child);
    if (name == "xfrm") {
      return child;
    }
    if (name == "spPr" || name == "grpSpPr") {
      return child_local(child, "xfrm");
    }
  }
  return {};
}

pugi::xml_attribute attr_local(const pugi::xml_node& node, std::string_view name) {
  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    const char* colon = std::strchr(attr.name(), ':');
    if (std::string_view(colon != nullptr ? colon + 1 : attr.name()) == name) {
      return attr;
    }
  }
  return {};
}

pugi::xml_node find_descendant(const pugi::xml_node& node, std::string_view name) {
  return node.find_node(
      [name](const pugi::xml_node& child) { return child.type() == pugi::node_element && local_name(child) == name; });
}

// The `c:chart` / `cx:chart` element of a chart frame, looking through an
// `mc:AlternateContent` inside `a:graphicData`.
pugi::xml_node frame_chart(const pugi::xml_node& content) {
  const pugi::xml_node data = child_local(child_local(content, "graphic"), "graphicData");
  if (const pugi::xml_node chart = child_local(data, "chart")) {
    return chart;
  }
  return child_local(child_local(child_local(data, "AlternateContent"), "Choice"), "chart");
}

}  // namespace

std::string relative_target(std::string_view from_dir, std::string_view target) {
  const std::string dir = std::string(from_dir) + "/";
  std::size_t common = 0;
  for (std::size_t i = 0; i < dir.size() && i < target.size() && dir[i] == target[i]; ++i) {
    if (dir[i] == '/') {
      common = i + 1;
    }
  }
  std::string out;
  for (std::size_t i = common; i < dir.size(); ++i) {
    if (dir[i] == '/') {
      out.append("../");
    }
  }
  out.append(target.substr(common));
  return out;
}

Expected<void, Error> parse_drawing_part(const std::vector<std::uint8_t>& bytes, pugi::xml_document& doc) {
  if (!io::load_xml_buffer(doc, bytes, "drawing", "drawing")) {
    return unparseable("part is not well-formed XML", {});
  }
  if (local_name(doc.document_element()) != "wsDr") {
    return unparseable("part root is not wsDr", doc.document_element().name());
  }
  return Expected<void, Error>::Ok();
}

std::string_view local_name(const char* qname) {
  const char* colon = std::strchr(qname, ':');
  return colon != nullptr ? colon + 1 : qname;
}

std::string_view local_name(const pugi::xml_node& node) {
  return local_name(node.name());
}

pugi::xml_node child_local(const pugi::xml_node& parent, std::string_view name) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && (name.empty() || local_name(child) == name)) {
      return child;
    }
  }
  return {};
}

std::vector<pugi::xml_node> children_local(const pugi::xml_node& parent, std::string_view name) {
  std::vector<pugi::xml_node> out;
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && local_name(child) == name) {
      out.push_back(child);
    }
  }
  return out;
}

std::vector<pugi::xml_node> drawing_anchors(const pugi::xml_node& root, bool include_fallback) {
  std::vector<pugi::xml_node> anchors;
  for (pugi::xml_node child = root.first_child(); child; child = child.next_sibling()) {
    if (is_anchor(child)) {
      anchors.push_back(child);
      continue;
    }
    if (local_name(child) != "AlternateContent") {
      continue;
    }
    bool first_choice = true;
    for (pugi::xml_node branch = child.first_child(); branch; branch = branch.next_sibling()) {
      const std::string_view name = local_name(branch);
      const bool take = include_fallback || (name == "Choice" && first_choice);
      first_choice = first_choice && name != "Choice";
      if (!take) {
        continue;
      }
      for (pugi::xml_node inner = branch.first_child(); inner; inner = inner.next_sibling()) {
        if (is_anchor(inner)) {
          anchors.push_back(inner);
        }
      }
    }
  }
  return anchors;
}

pugi::xml_node anchor_content(const pugi::xml_node& anchor) {
  for (pugi::xml_node child = anchor.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    const std::string_view name = local_name(child);
    if (name == "from" || name == "to" || name == "ext" || name == "pos" || name == "clientData") {
      continue;
    }
    if (name == "AlternateContent") {
      return child_local(child_local(child, "Choice"), "");
    }
    return child;
  }
  return {};
}

pugi::xml_node anchor_cnvpr(const pugi::xml_node& anchor) {
  const pugi::xml_node content = anchor_content(anchor);
  for (pugi::xml_node child = content.first_child(); child; child = child.next_sibling()) {
    const std::string_view name = local_name(child);
    if (name.size() > 4 && name.substr(0, 2) == "nv" && name.substr(name.size() - 2) == "Pr") {
      return child_local(child, "cNvPr");
    }
  }
  return {};
}

namespace {

using RelationshipTargetIndex = std::unordered_map<std::string_view, const io::ooxml::DrawingRel*>;

RelationshipTargetIndex index_drawing_relationships(const std::vector<io::ooxml::DrawingRel>& rels) {
  RelationshipTargetIndex index;
  index.reserve(rels.size());
  for (const io::ooxml::DrawingRel& rel : rels) {
    if (!rel.external && !rel.id.empty()) {
      index.insert_or_assign(std::string_view(rel.id), &rel);
    }
  }
  return index;
}

DrawingObject read_drawing_object_fields(const pugi::xml_node& anchor) {
  DrawingObject obj;
  const std::string_view anchor_name = local_name(anchor);
  if (anchor_name == "oneCellAnchor") {
    obj.anchor_kind = AnchorKind::kOneCell;
    obj.edit_as = EditAs::kOneCell;
  } else if (anchor_name == "absoluteAnchor") {
    obj.anchor_kind = AnchorKind::kAbsolute;
    obj.edit_as = EditAs::kAbsolute;
    const pugi::xml_node pos = child_local(anchor, "pos");
    obj.from.col_off = pos.attribute("x").as_llong();
    obj.from.row_off = pos.attribute("y").as_llong();
  } else {
    const std::string_view edit_as = anchor.attribute("editAs").value();
    obj.edit_as = edit_as == "oneCell"    ? EditAs::kOneCell
                  : edit_as == "absolute" ? EditAs::kAbsolute
                                          : EditAs::kTwoCell;
    obj.to = read_anchor_point(child_local(anchor, "to"));
  }
  if (obj.anchor_kind != AnchorKind::kAbsolute) {
    obj.from = read_anchor_point(child_local(anchor, "from"));
  }
  const pugi::xml_node content = anchor_content(anchor);
  obj.kind = classify(content);
  // A picture's own size is its unrotated `a:xfrm/a:ext`; the anchor box may be rotated.
  const pugi::xml_node xfrm_ext = child_local(content_xfrm(content), "ext");
  pugi::xml_node ext = obj.kind == DrawingObjectKind::kPicture && xfrm_ext ? xfrm_ext : child_local(anchor, "ext");
  if (!ext) {
    ext = xfrm_ext;
  }
  obj.cx = ext.attribute("cx").as_llong();
  obj.cy = ext.attribute("cy").as_llong();
  const pugi::xml_node cnvpr = anchor_cnvpr(anchor);
  obj.object_id = cnvpr.attribute("id").as_uint();
  obj.name = cnvpr.attribute("name").value();
  obj.descr = cnvpr.attribute("descr").value();
  if (obj.kind == DrawingObjectKind::kPicture) {
    obj.image_rel_id = attr_local(find_descendant(content, "blip"), "embed").value();
  }
  return obj;
}

}  // namespace

DrawingObject read_drawing_object(const pugi::xml_node& anchor, const std::vector<io::ooxml::DrawingRel>& rels) {
  DrawingObject obj = read_drawing_object_fields(anchor);
  if (obj.kind == DrawingObjectKind::kPicture) {
    for (const io::ooxml::DrawingRel& rel : rels) {
      if (!rel.external && !obj.image_rel_id.empty() && rel.id == obj.image_rel_id) {
        obj.media_path = rel.target;
      }
    }
  }
  return obj;
}

std::vector<DrawingObject> read_drawing_objects(const pugi::xml_document& doc,
                                                const std::vector<io::ooxml::DrawingRel>& rels) {
  std::vector<DrawingObject> objects;
  const RelationshipTargetIndex relation_index = index_drawing_relationships(rels);
  for (const pugi::xml_node& anchor : drawing_anchors(doc.document_element(), false)) {
    DrawingObject obj = read_drawing_object_fields(anchor);
    if (obj.kind == DrawingObjectKind::kPicture && !obj.image_rel_id.empty()) {
      const auto relation = relation_index.find(obj.image_rel_id);
      if (relation != relation_index.end()) {
        obj.media_path = relation->second->target;
      }
    }
    objects.push_back(std::move(obj));
  }
  return objects;
}

Expected<std::string, Error> sheet_drawing_path(const Workbook& wb, std::size_t sheet) {
  if (sheet >= wb.sheet_count()) {
    return make_error(FormulonErrorCode::kInvalidArgument, "drawing: sheet index out of range");
  }
  const Sheet& ws = wb.sheet(sheet);
  if (!ws.drawing_rel_target().empty()) {
    return ws.drawing_rel_target();
  }
  for (const UnknownRelationship& rel : ws.unknown_relationships()) {
    if (rel.type == io::kRelDrawing && !rel.target_external) {
      return rel.target;
    }
  }
  if (ws.is_opaque_ooxml_sheet()) {
    ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, ws.opaque_ooxml_part_path()));
    for (const io::ooxml::DrawingRel& rel : rels) {
      if (rel.type == io::kRelDrawing && !rel.external) {
        return rel.target;
      }
    }
  }
  return std::string();
}

Expected<std::vector<io::ooxml::DrawingRel>, Error> load_part_rels(const Workbook& wb, std::string_view part_path) {
  const PassthroughPart* rels = io::ooxml::find_passthrough_part(wb, io::ooxml::rels_path_for_part(part_path));
  if (rels == nullptr) {
    return std::vector<io::ooxml::DrawingRel>{};
  }
  return io::ooxml::parse_part_rels(rels->bytes, io::ooxml::dir_of(part_path));
}

Expected<std::vector<ChartPartRef>, Error> list_chart_parts(const Workbook& wb, std::size_t sheet) {
  ASSIGN_OR_RETURN(std::string path, sheet_drawing_path(wb, sheet));
  std::vector<ChartPartRef> charts;
  if (path.empty()) {
    return charts;
  }
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  if (part == nullptr) {
    return unparseable("part is missing", path);
  }
  pugi::xml_document doc;
  RETURN_IF_ERROR(parse_drawing_part(part->bytes, doc));
  ASSIGN_OR_RETURN(std::vector<io::ooxml::DrawingRel> rels, load_part_rels(wb, path));
  const RelationshipTargetIndex relation_index = index_drawing_relationships(rels);
  for (const pugi::xml_node& anchor : drawing_anchors(doc.document_element(), false)) {
    const pugi::xml_node content = anchor_content(anchor);
    if (classify(content) != DrawingObjectKind::kChart) {
      continue;
    }
    ChartPartRef ref;
    ref.object_id = anchor_cnvpr(anchor).attribute("id").as_uint();
    ref.is_chartex = strings::ends_with(
        child_local(child_local(content, "graphic"), "graphicData").attribute("uri").value(), "/chartex");
    const std::string_view rid = attr_local(frame_chart(content), "id").value();
    if (!rid.empty()) {
      const auto relation = relation_index.find(rid);
      if (relation != relation_index.end()) {
        ref.part_path = relation->second->target;
      }
    }
    charts.push_back(std::move(ref));
  }
  return charts;
}

Expected<std::vector<SheetChartPartRef>, Error> all_chart_parts(const Workbook& wb) {
  std::vector<SheetChartPartRef> charts;
  for (std::size_t sheet = 0; sheet < wb.sheet_count(); ++sheet) {
    auto refs = list_chart_parts(wb, sheet);
    if (!refs) {
      continue;
    }
    for (ChartPartRef& ref : refs.value()) {
      charts.push_back(SheetChartPartRef{sheet, std::move(ref)});
    }
  }
  return charts;
}

AnchorPoint read_anchor_point(const pugi::xml_node& marker) {
  AnchorPoint point;
  point.col = child_local(marker, "col").text().as_uint();
  point.col_off = child_local(marker, "colOff").text().as_llong();
  point.row = child_local(marker, "row").text().as_uint();
  point.row_off = child_local(marker, "rowOff").text().as_llong();
  return point;
}

void write_anchor_point(const pugi::xml_node& marker, const AnchorPoint& point) {
  child_local(marker, "col").text().set(point.col);
  child_local(marker, "colOff").text().set(static_cast<long long>(point.col_off));
  child_local(marker, "row").text().set(point.row);
  child_local(marker, "rowOff").text().set(static_cast<long long>(point.row_off));
}

}  // namespace formulon
