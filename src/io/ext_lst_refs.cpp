#include "io/ext_lst_refs.h"

#include <cstddef>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_rule_traits.h"
#include "io/cf_attr_names.h"
#include "io/sqref_text.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"

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

// Position of `group` among the document's sparkline groups.
std::size_t GroupPosition(const pugi::xml_node& group) {
  std::vector<pugi::xml_node> groups;
  Collect(group.root(), "x14:sparklineGroup", groups);
  std::size_t position = 0;
  while (position < groups.size() && groups[position] != group) {
    ++position;
  }
  return position;
}

// A sparkline group is only its sparklines plus styling, so losing the last
// sparkline takes the group with it; its position then goes to `removed_groups`.
void RemoveOwner(const pugi::xml_node& sqref, std::vector<std::size_t>* removed_groups) {
  pugi::xml_node doomed = sqref.parent();
  if (std::string_view(doomed.name()) == "x14:sparkline") {
    pugi::xml_node list = doomed.parent();
    list.remove_child(doomed);
    if (has_element_child(list) || std::string_view(list.name()) != "x14:sparklines") {
      return;
    }
    doomed = list.parent();
    if (removed_groups != nullptr) {
      removed_groups->push_back(GroupPosition(doomed));
    }
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

// Whether an x14 icon set's thresholds hold a formula.
bool HasFormulaCfvo(const pugi::xml_node& rule) {
  for (const pugi::xml_node cfvo : rule.child("x14:iconSet").children("x14:cfvo")) {
    if (std::string_view(cfvo.attribute("type").value()) == "formula") {
      return true;
    }
  }
  return false;
}

// Whether a far-line cell delete splits the range for an x14 rule of `type`, read from its thresholds.
bool FarEdgeSplits(const pugi::xml_node& rule, cf::RuleType type) {
  if (cf::far_edge_edit_splits_type(type)) {
    return true;
  }
  for (const pugi::xml_node visual : rule.children()) {
    for (const pugi::xml_node cfvo : visual.children("x14:cfvo")) {
      if (cf::far_edge_edit_splits_threshold(
              enum_from_name(kCfvoTypeNames, cfvo.attribute("type").value(), cf::CfvoType::Number))) {
        return true;
      }
    }
  }
  return false;
}

bool RemapSqrefs(const pugi::xml_document& doc, const SqrefRemap& remap, const CfSqrefRemap* cf_remap,
                 const SqrefRemap* dv_remap, std::vector<std::size_t>* removed_groups) {
  std::vector<pugi::xml_node> sqrefs;
  Collect(doc.document_element(), "xm:sqref", sqrefs);
  bool changed = false;
  for (const pugi::xml_node& node : sqrefs) {
    std::vector<MergeRange> ranges;
    if (!parse_sqref_text(node.text().get(), ranges)) {
      continue;
    }
    const std::vector<MergeRange> before = ranges;
    const std::string_view owner = node.parent().name();
    if (cf_remap != nullptr && owner == "x14:conditionalFormatting") {
      std::vector<X14CfRuleRef> rules;
      for (const pugi::xml_node rule : node.parent().children("x14:cfRule")) {
        const cf::RuleType type =
            enum_from_name(kCfRuleTypeNames, rule.attribute("type").value(), cf::RuleType::DataBar);
        rules.push_back(X14CfRuleRef{rule.attribute("id").value(),
                                     x14_cf_rule_class(type, type == cf::RuleType::IconSet && HasFormulaCfvo(rule)),
                                     FarEdgeSplits(rule, type)});
      }
      (*cf_remap)(ranges, rules);
    } else {
      (dv_remap != nullptr && owner == "x14:dataValidation" ? *dv_remap : remap)(ranges);
    }
    if (ranges.empty()) {
      RemoveOwner(node, removed_groups);
      changed = true;
    } else if (before != ranges) {
      node.text().set(format_sqref_text(ranges).c_str());
      changed = true;
    }
  }
  return changed;
}

bool CloneSparklines(const pugi::xml_document& doc, const SparklineClone& clone, FormulaTranslate translate) {
  std::vector<pugi::xml_node> sparklines;
  Collect(doc.document_element(), "x14:sparkline", sparklines);
  bool changed = false;
  for (const pugi::xml_node& node : sparklines) {
    std::vector<MergeRange> host;
    if (!parse_sqref_text(node.child("xm:sqref").text().get(), host) || host.size() != 1U ||
        !clone.hosted(host.front())) {
      continue;
    }
    pugi::xml_node after = node;
    for (std::uint32_t step = 1; step <= clone.count; ++step) {
      const std::optional<MergeRange> moved = clone.host_at(host.front(), step);
      if (!moved) {
        break;
      }
      pugi::xml_node copy = node.parent().insert_copy_after(node, after);
      copy.child("xm:sqref").text().set(format_sqref_text({*moved}).c_str());
      if (pugi::xml_node f = copy.child("xm:f")) {
        if (const std::optional<std::string> text =
                translate(f.text().get(), TranslateTransform(clone.row_axis, step))) {
          f.text().set(text->c_str());
        }
      }
      after = copy;
      changed = true;
    }
    SyncCount(node.parent());
  }
  return changed;
}

}  // namespace

X14CfRuleClass x14_cf_rule_class(cf::RuleType type, bool formula_cfvo) noexcept {
  return type == cf::RuleType::IconSet && formula_cfvo ? X14CfRuleClass::kFormulaIconSet : X14CfRuleClass::kOther;
}

bool remap_ext_lst_sqrefs(std::string& xml, const SqrefRemap& remap, const SparklineClone* clone,
                          FormulaTranslate translate, std::vector<std::size_t>* removed_groups,
                          const CfSqrefRemap* cf_remap, const SqrefRemap* dv_remap) {
  if (xml.find("xm:sqref") == std::string::npos) {
    return false;
  }
  pugi::xml_document doc;
  if (!Load(doc, xml)) {
    return false;
  }
  bool changed = RemapSqrefs(doc, remap, cf_remap, dv_remap, removed_groups);
  if (clone != nullptr) {
    changed = CloneSparklines(doc, *clone, translate) || changed;
  }
  if (changed) {
    store_root_xml(xml, doc);
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
    store_root_xml(xml, doc);
  }
  return changed;
}

}  // namespace formulon::io
