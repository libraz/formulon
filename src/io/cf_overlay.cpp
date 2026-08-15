//
// Implementation of the x14 conditional-formatting overlay
// reconciliation declared in `cf_overlay.h`. Operates purely on the raw
// `<extLst>` string: parse, prune, re-serialise. pugixml is
// non-validating and namespace-unaware, so the undeclared `x14:` / `xm:`
// prefixes inside the captured fragment parse as plain element names —
// the same convention `cf_reader.cpp` relies on.

#include "io/cf_overlay.h"

#include <cstddef>
#include <string>
#include <string_view>
#include <unordered_set>
#include <vector>

#include "io/xml_utils.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

/// True when `node` has at least one element child (text / comment /
/// PI children do not count as extension payload).
bool HasElementChild(const pugi::xml_node& node) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() == pugi::node_element) {
      return true;
    }
  }
  return false;
}

/// Collects the children of `parent` named `name` for which `doomed`
/// returns true, then removes them. Two-phase so removal never races the
/// sibling iteration.
template <typename Pred>
bool RemoveMatchingChildren(pugi::xml_node parent, const char* name, Pred doomed) {
  std::vector<pugi::xml_node> victims;
  for (pugi::xml_node child = parent.child(name); child; child = child.next_sibling(name)) {
    if (doomed(child)) {
      victims.push_back(child);
    }
  }
  for (const pugi::xml_node& victim : victims) {
    parent.remove_child(victim);
  }
  return !victims.empty();
}

/// URI of the worksheet-level `<ext>` that carries the x14
/// conditional-formatting overlay, and the x14 namespace, both fixed
/// constants Excel matches literally.
constexpr const char* kWorksheetCfExtUri = "{78C0D931-6437-407d-A8EE-F0AAD7539E65}";
constexpr const char* kX14Ns = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

/// Collects every id already claimed by an `<x14:cfRule>` anywhere under
/// `ext_lst`, at any nesting depth, so a rebuilt entry can be recognised
/// as a duplicate no matter which `<ext>` block holds the original.
void CollectClaimedRuleIds(const pugi::xml_node& node, std::unordered_set<std::string>* out) {
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (std::string_view(child.name()) == "x14:cfRule") {
      if (const char* id = child.attribute("id").value(); id[0] != '\0') {
        out->insert(id);
      }
    }
    CollectClaimedRuleIds(child, out);
  }
}

}  // namespace

std::string merge_x14_cf_entries(const std::string& ext_lst_xml, const std::string& entries) {
  if (entries.empty()) {
    return ext_lst_xml;
  }

  // The entries arrive as a bare sibling sequence; give them a root so
  // pugixml will parse more than the first one.
  pugi::xml_document entries_doc;
  const std::string wrapped = "<entries>" + entries + "</entries>";
  if (!entries_doc.load_string(wrapped.c_str())) {
    return ext_lst_xml;
  }

  pugi::xml_document doc;
  pugi::xml_node ext_lst;
  if (ext_lst_xml.empty()) {
    ext_lst = doc.append_child("extLst");
  } else {
    if (!doc.load_string(ext_lst_xml.c_str())) {
      return ext_lst_xml;
    }
    ext_lst = doc.child("extLst");
    if (!ext_lst) {
      return ext_lst_xml;
    }
  }

  std::unordered_set<std::string> claimed;
  CollectClaimedRuleIds(ext_lst, &claimed);

  pugi::xml_node formattings;
  bool added = false;
  for (pugi::xml_node entry = entries_doc.child("entries").first_child(); entry; entry = entry.next_sibling()) {
    const char* id = entry.child("x14:cfRule").attribute("id").value();
    if (id[0] == '\0' || claimed.count(id) != 0U) {
      continue;
    }
    if (!formattings) {
      for (pugi::xml_node ext = ext_lst.child("ext"); ext && !formattings; ext = ext.next_sibling("ext")) {
        formattings = ext.child("x14:conditionalFormattings");
      }
    }
    if (!formattings) {
      pugi::xml_node ext = ext_lst.append_child("ext");
      ext.append_attribute("uri") = kWorksheetCfExtUri;
      ext.append_attribute("xmlns:x14") = kX14Ns;
      formattings = ext.append_child("x14:conditionalFormattings");
    }
    formattings.append_copy(entry);
    claimed.insert(id);
    added = true;
  }

  if (!added) {
    // Nothing new: hand the original bytes back untouched rather than
    // re-serialising a document that only round-tripped through pugixml.
    return ext_lst_xml;
  }
  return raw_xml(ext_lst);
}

std::string reconcile_x14_cf_overlay(const std::string& ext_lst_xml,
                                     const std::vector<cf::ConditionalFormat>& formats) {
  if (ext_lst_xml.empty()) {
    return std::string();
  }

  std::unordered_set<std::string> surviving_ids;
  for (const auto& block : formats) {
    for (const auto& rule : block.rules) {
      if (!rule.id.empty()) {
        surviving_ids.insert(rule.id);
      }
    }
  }

  pugi::xml_document doc;
  const pugi::xml_parse_result rc = doc.load_string(ext_lst_xml.c_str());
  const pugi::xml_node ext_lst = doc.child("extLst");
  if (!rc || !ext_lst) {
    // Unparseable or unexpected shape: the ids the overlay references
    // cannot be enumerated, so surgical pruning is impossible. Drop the
    // whole overlay rather than re-emit bytes that may reference a rule
    // the mutation just removed.
    return std::string();
  }

  // Pass 1: drop every id-bearing <x14:cfRule> whose id no longer exists
  // in the model. Rules without an id have no legacy counterpart to
  // dangle against and are kept.
  bool changed = false;
  for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
    for (pugi::xml_node formattings = ext.child("x14:conditionalFormattings"); formattings;
         formattings = formattings.next_sibling("x14:conditionalFormattings")) {
      for (pugi::xml_node block = formattings.child("x14:conditionalFormatting"); block;
           block = block.next_sibling("x14:conditionalFormatting")) {
        changed |= RemoveMatchingChildren(block, "x14:cfRule", [&surviving_ids](const pugi::xml_node& rule) {
          const char* id = rule.attribute("id").value();
          return id[0] != '\0' && surviving_ids.count(id) == 0U;
        });
      }
    }
  }

  if (!changed) {
    // Nothing was pruned; hand the original bytes back untouched so an
    // unrelated mutation cannot perturb the overlay's serialisation.
    return ext_lst_xml;
  }

  // Pass 2: prune containers hollowed out by the rule removal, innermost
  // first: <x14:conditionalFormatting> with no <x14:cfRule> left, then
  // <x14:conditionalFormattings> with no block left, then <ext> elements
  // with no element payload at all.
  for (pugi::xml_node ext = ext_lst.child("ext"); ext; ext = ext.next_sibling("ext")) {
    for (pugi::xml_node formattings = ext.child("x14:conditionalFormattings"); formattings;
         formattings = formattings.next_sibling("x14:conditionalFormattings")) {
      RemoveMatchingChildren(formattings, "x14:conditionalFormatting",
                             [](const pugi::xml_node& block) { return !block.child("x14:cfRule"); });
    }
    RemoveMatchingChildren(ext, "x14:conditionalFormattings",
                           [](const pugi::xml_node& formattings) { return !HasElementChild(formattings); });
  }
  RemoveMatchingChildren(ext_lst, "ext", [](const pugi::xml_node& ext) { return !HasElementChild(ext); });

  if (!HasElementChild(ext_lst)) {
    return std::string();
  }
  return raw_xml(ext_lst);
}

}  // namespace formulon::io
