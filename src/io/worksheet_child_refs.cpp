#include "io/worksheet_child_refs.h"

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "io/ext_lst_refs.h"
#include "io/sqref_text.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

enum class Outcome { kUnchanged, kChanged, kEmptied };

/// The two ways an edit moves a retained range.
struct ChildRefRemap {
  /// Ranges that stretch over an insert inside them and shrink under a delete.
  SqrefRemap move;
  /// Ranges an edit inside them cuts in two.
  SqrefRemap cut;
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
};

ChildRemap RuleFor(std::string_view name) {
  for (const ChildRule& rule : kRules) {
    if (rule.name == name) {
      return rule.apply;
    }
  }
  return nullptr;
}

}  // namespace

void remap_worksheet_children(WorksheetRawExtensions& children, const StructuralEdit& edit) {
  const ChildRefRemap remap{[&edit](std::vector<MergeRange>& ranges) {
                              shift_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
                            },
                            [&edit](std::vector<MergeRange>& ranges) {
                              cut_sqref_ranges(ranges, edit.index, edit.count, edit.is_delete, edit.row_axis);
                            }};
  for (auto it = children.begin(); it != children.end();) {
    // The slot names the element unless it was an unknown one parked there.
    if (it->slot >= worksheet_child::kCount || RuleFor(worksheet_child::kOrder[it->slot]) == nullptr) {
      ++it;
      continue;
    }
    pugi::xml_document doc;
    const pugi::xml_node root = doc.load_string(it->xml.c_str()) ? doc.document_element() : pugi::xml_node();
    const ChildRemap apply = root ? RuleFor(root.name()) : nullptr;
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
