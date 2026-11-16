#include "io/worksheet_child_refs.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/ext_lst_refs.h"
#include "io/sqref_text.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/a1_ref.h"

namespace formulon::io {
namespace {

enum class Outcome { kUnchanged, kChanged, kEmptied };

/// The two ways an edit moves a retained range, and the edit itself for the
/// children that hold cell anchors rather than ranges.
struct ChildRefRemap {
  /// Ranges that stretch over an insert inside them and shrink under a delete.
  SqrefRemap move;
  /// Ranges an edit inside them cuts in two.
  SqrefRemap cut;
  const StructuralEdit& edit;
};

// Maps the A1 text of `attr` through `remap`; an attribute that does not parse
// is left alone.
Outcome RemapAttr(pugi::xml_attribute attr, const SqrefRemap& remap) {
  std::vector<MergeRange> ranges;
  if (!attr || !parse_sqref_text(attr.value(), ranges)) {
    return Outcome::kUnchanged;
  }
  const std::vector<MergeRange> before = ranges;
  remap(ranges);
  if (ranges.empty()) {
    return Outcome::kEmptied;
  }
  if (ranges == before) {
    return Outcome::kUnchanged;
  }
  attr.set_value(format_sqref_text(ranges).c_str());
  return Outcome::kChanged;
}

void SyncCount(pugi::xml_node node, const char* item) {
  if (pugi::xml_attribute count = node.attribute("count")) {
    unsigned n = 0;
    for (pugi::xml_node child = node.child(item); child; child = child.next_sibling(item)) {
      ++n;
    }
    count.set_value(n);
  }
}

// Remaps `attr` on each `item` child of `root` that `selected` accepts,
// removing the ones the edit empties. Returns how many it removed.
template <typename Selected>
std::size_t RemapEach(pugi::xml_node root, const char* item, const char* attr, const SqrefRemap& remap,
                      Selected selected, bool& changed) {
  std::size_t removed = 0;
  for (pugi::xml_node node = root.child(item); node;) {
    const pugi::xml_node next = node.next_sibling(item);
    if (selected(node)) {
      const Outcome outcome = RemapAttr(node.attribute(attr), remap);
      if (outcome == Outcome::kEmptied) {
        root.remove_child(node);
        ++removed;
      }
      changed = changed || outcome != Outcome::kUnchanged;
    }
    node = next;
  }
  return removed;
}

constexpr auto kAll = [](const pugi::xml_node&) { return true; };

// Each returns false when the edit leaves the child without content.

bool RemapProtectedRanges(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  return RemapEach(root, "protectedRange", "sqref", remap.move, kAll, changed) == 0U || root.child("protectedRange");
}

bool RemapWebPublishItems(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  const auto range_source = [](const pugi::xml_node& node) {
    return std::string_view(node.attribute("sourceType").value()) == "range";
  };
  if (RemapEach(root, "webPublishItem", "sourceRef", remap.move, range_source, changed) == 0U) {
    return true;
  }
  SyncCount(root, "webPublishItem");
  return static_cast<bool>(root.child("webPublishItem"));
}

bool RemapSortState(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  const Outcome outcome = RemapAttr(root.attribute("ref"), remap.move);
  if (outcome == Outcome::kEmptied) {
    return false;
  }
  changed = changed || outcome == Outcome::kChanged;
  return RemapEach(root, "sortCondition", "ref", remap.move, kAll, changed) == 0U || root.child("sortCondition");
}

// A watch on a deleted cell keeps its address.
bool RemapCellWatches(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  for (pugi::xml_node node = root.child("cellWatch"); node; node = node.next_sibling("cellWatch")) {
    pugi::xml_attribute r = node.attribute("r");
    const std::string before = r.value();
    if (RemapAttr(r, remap.move) == Outcome::kEmptied) {
      r.set_value(before.c_str());
    }
    changed = changed || before != r.value();
  }
  return true;
}

// The first entry an edit empties takes every later entry with it.
bool RemapIgnoredErrors(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  for (pugi::xml_node node = root.child("ignoredError"); node; node = node.next_sibling("ignoredError")) {
    const Outcome outcome = RemapAttr(node.attribute("sqref"), remap.cut);
    if (outcome == Outcome::kEmptied) {
      while (pugi::xml_node doomed = node.next_sibling("ignoredError")) {
        root.remove_child(doomed);
      }
      root.remove_child(node);
      changed = true;
      return static_cast<bool>(root.child("ignoredError"));
    }
    changed = changed || outcome == Outcome::kChanged;
  }
  return true;
}

// A scenario goes with its last input cell; `current` and `show` keep their
// index unless it now points past the last scenario left, which they clamp to.
bool RemapScenarios(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  unsigned remaining = 0;
  bool removed = false;
  for (pugi::xml_node scenario = root.child("scenario"); scenario;) {
    const pugi::xml_node next = scenario.next_sibling("scenario");
    if (RemapEach(scenario, "inputCells", "r", remap.move, kAll, changed) != 0U) {
      if (scenario.child("inputCells")) {
        SyncCount(scenario, "inputCells");
      } else {
        root.remove_child(scenario);
        removed = true;
        scenario = next;
        continue;
      }
    }
    ++remaining;
    scenario = next;
  }
  if (!removed) {
    return true;
  }
  if (remaining == 0U) {
    return false;
  }
  for (const char* name : {"current", "show"}) {
    if (pugi::xml_attribute attr = root.attribute(name); attr && attr.as_uint() >= remaining) {
      attr.set_value(remaining - 1U);
    }
  }
  return true;
}

// Reads the `edit`-axis line and offset of an anchor marker (`xdr:col` /
// `xdr:colOff` or `xdr:row` / `xdr:rowOff`).
AnchorLine ReadMarkerLine(const pugi::xml_node& marker, bool row_axis) {
  return {static_cast<std::uint32_t>(child_local(marker, row_axis ? "row" : "col").text().as_ullong()),
          child_local(marker, row_axis ? "rowOff" : "colOff").text().as_llong()};
}

void WriteMarkerLine(const pugi::xml_node& marker, bool row_axis, const AnchorLine& at) {
  child_local(marker, row_axis ? "row" : "col").text().set(static_cast<unsigned long long>(at.line));
  child_local(marker, row_axis ? "rowOff" : "colOff").text().set(static_cast<long long>(at.offset));
}

// The `anchor` of each form control (`controlPr`) and OLE object (`objectPr`),
// in either branch of an `mc:AlternateContent`.
std::vector<pugi::xml_node> ObjectAnchors(pugi::xml_node root) {
  struct Walker : pugi::xml_tree_walker {
    std::vector<pugi::xml_node> anchors;
    bool for_each(pugi::xml_node& node) override {
      const std::string_view parent = local_name(node.parent());
      if (local_name(node) == "anchor" && (parent == "controlPr" || parent == "objectPr")) {
        anchors.push_back(node);
      }
      return true;
    }
  } walker;
  root.traverse(walker);
  return walker.anchors;
}

// Moves each object anchor; an object that does not move with cells keeps its anchor.
bool RemapObjectAnchors(pugi::xml_node root, const ChildRefRemap& remap, bool& changed) {
  const StructuralEdit& edit = remap.edit;
  if (edit.band.has_value()) {
    return true;
  }
  for (const pugi::xml_node& anchor : ObjectAnchors(root)) {
    const pugi::xml_node from_node = child_local(anchor, "from");
    const pugi::xml_node to_node = child_local(anchor, "to");
    if (!from_node || !to_node || !parse_xml_bool_attr(anchor.attribute("moveWithCells"))) {
      continue;
    }
    const AnchorLine from = ReadMarkerLine(from_node, edit.row_axis);
    const AnchorLine to = ReadMarkerLine(to_node, edit.row_axis);
    AnchorLine moved_from = from;
    AnchorLine moved_to = to;
    move_cell_anchor(moved_from, moved_to, parse_xml_bool_attr(anchor.attribute("sizeWithCells")), edit);
    if (moved_from.line != from.line || moved_from.offset != from.offset) {
      WriteMarkerLine(from_node, edit.row_axis, moved_from);
      changed = true;
    }
    if (moved_to.line != to.line || moved_to.offset != to.offset) {
      WriteMarkerLine(to_node, edit.row_axis, moved_to);
      changed = true;
    }
  }
  return true;
}

using ChildRemap = bool (*)(pugi::xml_node, const ChildRefRemap&, bool&);

struct ChildRule {
  std::string_view name;
  ChildRemap apply;
};

constexpr ChildRule kRules[] = {
    {"protectedRanges", RemapProtectedRanges},
    {"scenarios", RemapScenarios},
    {"sortState", RemapSortState},
    {"cellWatches", RemapCellWatches},
    {"ignoredErrors", RemapIgnoredErrors},
    {"webPublishItems", RemapWebPublishItems},
    {"controls", RemapObjectAnchors},
    {"oleObjects", RemapObjectAnchors},
    // Excel wraps `<controls>` in an `mc:AlternateContent`, which takes the
    // slot of the child before it.
    {"AlternateContent", RemapObjectAnchors},
};

ChildRemap RuleFor(std::string_view name) {
  for (const ChildRule& rule : kRules) {
    if (rule.name == name) {
      return rule.apply;
    }
  }
  return nullptr;
}

// Whether a retained child's root element is an `mc:AlternateContent`.
bool IsAlternateContent(std::string_view xml) {
  constexpr std::string_view kName = "AlternateContent";
  const std::string_view tag = xml.substr(0, xml.find_first_of(" />"));
  return tag.size() >= kName.size() && tag.substr(tag.size() - kName.size()) == kName;
}

bool IsObjectContainer(const pugi::xml_node& node) {
  const std::string_view name = local_name(node);
  return name == "controls" || name == "oleObjects";
}

// Whether a retained child is a `<controls>` / `<oleObjects>` container or
// the `mc:AlternateContent` Excel wraps one in.
bool MayHoldObjects(const WorksheetRawChild& child) {
  const bool container_slot =
      child.slot < worksheet_child::kCount && (worksheet_child::kOrder[child.slot] == std::string_view("controls") ||
                                               worksheet_child::kOrder[child.slot] == std::string_view("oleObjects"));
  return container_slot || IsAlternateContent(child.xml);
}

// Appends the value of every prefixed `id` attribute (`r:id`) under `node`.
void CollectRelIds(const pugi::xml_node& node, std::vector<std::string>& ids) {
  for (const pugi::xml_attribute& attr : node.attributes()) {
    const std::string_view name = attr.name();
    if (name.size() > 3U && name.substr(name.size() - 3U) == ":id" &&
        std::find(ids.begin(), ids.end(), attr.value()) == ids.end()) {
      ids.emplace_back(attr.value());
    }
  }
  for (const pugi::xml_node& child : node.children()) {
    CollectRelIds(child, ids);
  }
}

// Removes the listed objects under `root`; returns false when it empties an
// object container, which drops `root` as a whole.
bool RemoveObjects(pugi::xml_node root, const std::vector<std::uint32_t>& shape_ids, std::vector<std::string>& ids,
                   bool& changed) {
  struct Walker : pugi::xml_tree_walker {
    std::vector<pugi::xml_node> objects;
    bool for_each(pugi::xml_node& node) override {
      const std::string_view name = local_name(node);
      if (name == "control" || name == "oleObject") {
        objects.push_back(node);
      }
      return true;
    }
  } walker;
  root.traverse(walker);
  std::vector<pugi::xml_node> containers;
  for (const pugi::xml_node& object : walker.objects) {
    const std::uint32_t shape_id = object.attribute("shapeId").as_uint();
    if (std::find(shape_ids.begin(), shape_ids.end(), shape_id) == shape_ids.end()) {
      continue;
    }
    // The object, or the `mc:AlternateContent` wrapping it, is the container's child.
    pugi::xml_node entry = object;
    while (entry.parent() && !IsObjectContainer(entry.parent())) {
      entry = entry.parent();
    }
    pugi::xml_node container = entry.parent();
    if (!container || entry.type() != pugi::node_element) {
      continue;
    }
    CollectRelIds(entry, ids);
    container.remove_child(entry);
    changed = true;
    if (std::find(containers.begin(), containers.end(), container) == containers.end()) {
      containers.push_back(container);
    }
  }
  // An emptied container goes, and with it the top-level `mc:AlternateContent` wrapping it.
  return std::all_of(containers.begin(), containers.end(),
                     [](const pugi::xml_node& container) { return has_element_child(container); });
}

}  // namespace

std::vector<std::string> remove_worksheet_objects(WorksheetRawExtensions& children,
                                                  const std::vector<std::uint32_t>& shape_ids) {
  std::vector<std::string> ids;
  if (shape_ids.empty()) {
    return ids;
  }
  for (auto it = children.begin(); it != children.end();) {
    if (!MayHoldObjects(*it)) {
      ++it;
      continue;
    }
    pugi::xml_document doc;
    const pugi::xml_node root = doc.load_string(it->xml.c_str()) ? doc.document_element() : pugi::xml_node();
    bool changed = false;
    if (!root || !RemoveObjects(root, shape_ids, ids, changed)) {
      it = root ? children.erase(it) : it + 1;
      continue;
    }
    if (changed) {
      it->xml = raw_xml(root);
    }
    ++it;
  }
  return ids;
}

AnchorLine shift_anchor_line(AnchorLine marker, const StructuralEdit& edit) {
  if (marker.line < edit.index) {
    return marker;
  }
  if (!edit.is_delete) {
    marker.line = static_cast<std::uint32_t>(std::min<std::uint64_t>(
        static_cast<std::uint64_t>(marker.line) + edit.count, a1::axis_extent(edit.row_axis) - 1U));
  } else if (marker.line >= edit.end()) {
    marker.line -= edit.count;
  } else {
    marker.line = edit.index;
    marker.offset = 0;
  }
  return marker;
}

void move_cell_anchor(AnchorLine& from, AnchorLine& to, bool size_with_cells, const StructuralEdit& edit) {
  const AnchorLine moved_from = shift_anchor_line(from, edit);
  if (size_with_cells) {
    to = shift_anchor_line(to, edit);
  } else {
    // `from` moved by `moved_from.line - from.line` lines; `to` keeps the span.
    const std::int64_t delta = static_cast<std::int64_t>(moved_from.line) - static_cast<std::int64_t>(from.line);
    const std::int64_t last = static_cast<std::int64_t>(a1::axis_extent(edit.row_axis)) - 1;
    to.line = static_cast<std::uint32_t>(std::clamp<std::int64_t>(static_cast<std::int64_t>(to.line) + delta, 0, last));
  }
  from = moved_from;
}

bool cell_anchor_deleted(const AnchorLine& from, const AnchorLine& to, const StructuralEdit& edit) {
  if (!edit.is_delete || edit.band.has_value() || edit.count == 0U) {
    return false;
  }
  const auto deleted = [&edit](std::uint32_t line) { return line >= edit.index && line - edit.index < edit.count; };
  // The `to` marker is exclusive at offset 0: an object ending on the line
  // after the span lies wholly inside it.
  return deleted(from.line) && (deleted(to.line) || (to.line == edit.index + edit.count && to.offset == 0));
}

std::vector<std::uint32_t> deleted_cell_objects(const WorksheetRawExtensions& children, const StructuralEdit& edit) {
  std::vector<std::uint32_t> shape_ids;
  if (!edit.is_delete) {
    return shape_ids;
  }
  for (const WorksheetRawChild& child : children) {
    pugi::xml_document doc;
    if (!MayHoldObjects(child) || !doc.load_string(child.xml.c_str())) {
      continue;
    }
    for (const pugi::xml_node& anchor : ObjectAnchors(doc.document_element())) {
      const pugi::xml_node from_node = child_local(anchor, "from");
      const pugi::xml_node to_node = child_local(anchor, "to");
      if (!from_node || !to_node || !parse_xml_bool_attr(anchor.attribute("moveWithCells")) ||
          !parse_xml_bool_attr(anchor.attribute("sizeWithCells")) ||
          !cell_anchor_deleted(ReadMarkerLine(from_node, edit.row_axis), ReadMarkerLine(to_node, edit.row_axis),
                               edit)) {
        continue;
      }
      // anchor -> controlPr / objectPr -> control / oleObject
      const std::uint32_t id = anchor.parent().parent().attribute("shapeId").as_uint();
      if (id != 0U && std::find(shape_ids.begin(), shape_ids.end(), id) == shape_ids.end()) {
        shape_ids.push_back(id);
      }
    }
  }
  return shape_ids;
}

void remap_worksheet_children(WorksheetRawExtensions& children, const StructuralEdit& edit) {
  const ChildRefRemap remap{[&edit](std::vector<MergeRange>& ranges) { shift_reference_ranges(ranges, edit); },
                            [&edit](std::vector<MergeRange>& ranges) { cut_sqref_ranges(ranges, edit); }, edit};
  for (auto it = children.begin(); it != children.end();) {
    // The slot names the element unless it was an unknown one parked there.
    const bool named_rule = it->slot < worksheet_child::kCount && RuleFor(worksheet_child::kOrder[it->slot]) != nullptr;
    if (!named_rule && !IsAlternateContent(it->xml)) {
      ++it;
      continue;
    }
    pugi::xml_document doc;
    const pugi::xml_node root = doc.load_string(it->xml.c_str()) ? doc.document_element() : pugi::xml_node();
    const ChildRemap apply = root ? RuleFor(local_name(root)) : nullptr;
    if (apply == nullptr) {
      ++it;
      continue;
    }
    bool changed = false;
    if (!apply(root, remap, changed)) {
      it = children.erase(it);
      continue;
    }
    if (changed) {
      it->xml = raw_xml(root);
    }
    ++it;
  }
}

}  // namespace formulon::io
