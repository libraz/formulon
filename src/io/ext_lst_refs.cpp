#include "io/ext_lst_refs.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/a1_column.h"
#include "utils/a1_ref.h"

namespace formulon::io {
namespace {

void Collect(const pugi::xml_node& node, std::string_view name, std::vector<pugi::xml_node>& out) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      continue;
    }
    if (name == child.name()) {
      out.push_back(child);
    } else {
      Collect(child, name, out);
    }
  }
}

bool ParseSqref(std::string_view text, std::vector<MergeRange>& out) {
  std::size_t i = 0;
  while (i < text.size()) {
    if (text[i] == ' ') {
      ++i;
      continue;
    }
    const std::size_t end = std::min(text.find(' ', i), text.size());
    const std::string_view token = text.substr(i, end - i);
    const std::size_t colon = token.find(':');
    MergeRange range;
    if (!a1::parse_a1_ref(token.substr(0, colon), &range.first_row, &range.first_col)) {
      return false;
    }
    range.last_row = range.first_row;
    range.last_col = range.first_col;
    if (colon != std::string_view::npos &&
        !a1::parse_a1_ref(token.substr(colon + 1U), &range.last_row, &range.last_col)) {
      return false;
    }
    out.push_back(range);
    i = end;
  }
  return !out.empty();
}

std::string EncodeSqref(const std::vector<MergeRange>& ranges) {
  std::string out;
  const auto cell = [&out](std::uint32_t row, std::uint32_t col) {
    a1::append_column_letters(out, col);
    out += std::to_string(static_cast<std::uint64_t>(row) + 1U);
  };
  for (const MergeRange& r : ranges) {
    if (!out.empty()) {
      out += ' ';
    }
    cell(r.first_row, r.first_col);
    if (r.first_row != r.last_row || r.first_col != r.last_col) {
      out += ':';
      cell(r.last_row, r.last_col);
    }
  }
  return out;
}

void SyncCount(pugi::xml_node container) {
  pugi::xml_attribute count = container.attribute("count");
  if (!count) {
    return;
  }
  unsigned n = 0;
  for (pugi::xml_node child = container.first_child(); child; child = child.next_sibling()) {
    n += child.type() == pugi::node_element ? 1U : 0U;
  }
  count.set_value(n);
}

// A sparkline group is only its sparklines plus styling, so losing the last
// sparkline takes the group with it.
void RemoveOwner(const pugi::xml_node& sqref) {
  pugi::xml_node doomed = sqref.parent();
  if (std::string_view(doomed.name()) == "x14:sparkline") {
    pugi::xml_node list = doomed.parent();
    list.remove_child(doomed);
    if (has_element_child(list) || std::string_view(list.name()) != "x14:sparklines") {
      return;
    }
    doomed = list.parent();
  }
  for (;;) {
    pugi::xml_node parent = doomed.parent();
    parent.remove_child(doomed);
    if (parent.parent().type() == pugi::node_document || has_element_child(parent)) {
      SyncCount(parent);
      return;
    }
    doomed = parent;
  }
}

bool Load(pugi::xml_document& doc, const std::string& xml) {
  return !xml.empty() && doc.load_string(xml.c_str()) && doc.document_element();
}

void Store(std::string& xml, const pugi::xml_document& doc) {
  const pugi::xml_node root = doc.document_element();
  xml = has_element_child(root) ? raw_xml(root) : std::string();
}

}  // namespace

bool remap_ext_lst_sqrefs(std::string& xml, const SqrefRemap& remap) {
  if (xml.find("xm:sqref") == std::string::npos) {
    return false;
  }
  pugi::xml_document doc;
  if (!Load(doc, xml)) {
    return false;
  }
  std::vector<pugi::xml_node> sqrefs;
  Collect(doc.document_element(), "xm:sqref", sqrefs);
  bool changed = false;
  for (const pugi::xml_node& node : sqrefs) {
    std::vector<MergeRange> ranges;
    if (!ParseSqref(node.text().get(), ranges)) {
      continue;
    }
    const std::vector<MergeRange> before = ranges;
    remap(ranges);
    if (ranges.empty()) {
      RemoveOwner(node);
      changed = true;
    } else if (before != ranges) {
      node.text().set(EncodeSqref(ranges).c_str());
      changed = true;
    }
  }
  if (changed) {
    Store(xml, doc);
  }
  return changed;
}

bool remap_ext_lst_formulas(std::string& xml, const FormulaRemap& remap) {
  if (xml.find("xm:f") == std::string::npos) {
    return false;
  }
  pugi::xml_document doc;
  if (!Load(doc, xml)) {
    return false;
  }
  std::vector<pugi::xml_node> formulas;
  Collect(doc.document_element(), "xm:f", formulas);
  bool changed = false;
  for (const pugi::xml_node& node : formulas) {
    if (const std::optional<std::string> text = remap(node.text().get())) {
      node.text().set(text->c_str());
      changed = true;
    }
  }
  if (changed) {
    Store(xml, doc);
  }
  return changed;
}

}  // namespace formulon::io
