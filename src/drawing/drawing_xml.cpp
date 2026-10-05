#include "drawing/drawing_xml.h"

#include <cstring>
#include <utility>

#include "io/ooxml/package_validator.h"
#include "io/xml_utils.h"

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

// The element an anchor positions: its first child that is not a marker.
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
    const auto ends_with = [uri](std::string_view suffix) {
      return uri.size() >= suffix.size() && uri.substr(uri.size() - suffix.size()) == suffix;
    };
    return ends_with("/chart") || ends_with("/chartex") ? DrawingObjectKind::kChart : DrawingObjectKind::kGraphicFrame;
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
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (local_name(child) == name) {
      return child;
    }
    if (pugi::xml_node found = find_descendant(child, name)) {
      return found;
    }
  }
  return {};
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

Expected<std::vector<DrawingRel>, Error> parse_part_rels(const std::vector<std::uint8_t>& bytes,
                                                         std::string_view owner_dir) {
  pugi::xml_document doc;
  if (!io::load_xml_buffer(doc, bytes, "drawing", "rels")) {
    return unparseable("rels part is not well-formed XML", owner_dir);
  }
  std::vector<DrawingRel> rels;
  for (pugi::xml_node rel = doc.child("Relationships").child("Relationship"); rel;
       rel = rel.next_sibling("Relationship")) {
    DrawingRel entry;
    entry.id = rel.attribute("Id").value();
    entry.type = rel.attribute("Type").value();
    entry.external = std::strcmp(rel.attribute("TargetMode").value(), "External") == 0;
    entry.target = rel.attribute("Target").value();
    if (!entry.external) {
      auto resolved = io::ooxml::resolve_relative_path(owner_dir, entry.target);
      if (!resolved) {
        return unparseable("rels target does not resolve", entry.target);
      }
      entry.target = std::move(resolved.value());
    }
    rels.push_back(std::move(entry));
  }
  return rels;
}

std::string_view local_name(const pugi::xml_node& node) {
  const char* name = node.name();
  const char* colon = std::strchr(name, ':');
  return colon != nullptr ? colon + 1 : name;
}

pugi::xml_node child_local(const pugi::xml_node& parent, std::string_view name) {
  for (pugi::xml_node child = parent.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element && (name.empty() || local_name(child) == name)) {
      return child;
    }
  }
  return {};
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

DrawingObject read_drawing_object(const pugi::xml_node& anchor, const std::vector<DrawingRel>& rels) {
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
  pugi::xml_node ext = child_local(anchor, "ext");
  if (!ext) {
    ext = child_local(content_xfrm(content), "ext");
  }
  obj.cx = ext.attribute("cx").as_llong();
  obj.cy = ext.attribute("cy").as_llong();
  const pugi::xml_node cnvpr = anchor_cnvpr(anchor);
  obj.object_id = cnvpr.attribute("id").as_uint();
  obj.name = cnvpr.attribute("name").value();
  obj.descr = cnvpr.attribute("descr").value();
  if (obj.kind == DrawingObjectKind::kPicture) {
    obj.image_rel_id = attr_local(find_descendant(content, "blip"), "embed").value();
    for (const DrawingRel& rel : rels) {
      if (!rel.external && !obj.image_rel_id.empty() && rel.id == obj.image_rel_id) {
        obj.media_path = rel.target;
      }
    }
  }
  return obj;
}

std::vector<DrawingObject> read_drawing_objects(const pugi::xml_document& doc, const std::vector<DrawingRel>& rels) {
  std::vector<DrawingObject> objects;
  for (const pugi::xml_node& anchor : drawing_anchors(doc.document_element(), false)) {
    objects.push_back(read_drawing_object(anchor, rels));
  }
  return objects;
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
