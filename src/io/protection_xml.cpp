#include "io/protection_xml.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pugixml.hpp"

namespace formulon::io {
namespace {

constexpr std::string_view kWorkbookProtection = "workbookProtection";
constexpr std::string_view kProtectedRanges = "protectedRanges";

void SetBoolAttr(pugi::xml_node node, const char* name, bool value) {
  node.remove_attribute(name);
  if (value) {
    node.append_attribute(name).set_value("1");
  }
}

void SetTextAttr(pugi::xml_node node, const char* name, const std::string& value) {
  node.remove_attribute(name);
  if (!value.empty()) {
    node.append_attribute(name).set_value(value.c_str());
  }
}

void SetUintAttr(pugi::xml_node node, const char* name, std::uint32_t value) {
  node.remove_attribute(name);
  if (value != 0U) {
    node.append_attribute(name).set_value(value);
  }
}

bool IsProtectedRanges(const WorksheetRawChild& child) {
  return child.xml.compare(0, kProtectedRanges.size() + 1U, "<protectedRanges") == 0;
}

}  // namespace

WorkbookProtection parse_workbook_protection(std::string_view xml) {
  WorkbookProtection out;
  pugi::xml_document doc;
  if (xml.empty() || !doc.load_buffer(xml.data(), xml.size())) {
    return out;
  }
  const pugi::xml_node node = doc.child(std::string(kWorkbookProtection).c_str());
  out.lock_structure = read_xsd_bool(node, "lockStructure", false);
  out.lock_windows = read_xsd_bool(node, "lockWindows", false);
  out.lock_revision = read_xsd_bool(node, "lockRevision", false);
  out.workbook_password = attr_str(node, "workbookPassword");
  out.workbook_algorithm_name = attr_str(node, "workbookAlgorithmName");
  out.workbook_hash_value = attr_str(node, "workbookHashValue");
  out.workbook_salt_value = attr_str(node, "workbookSaltValue");
  out.workbook_spin_count = attr_u32(node, "workbookSpinCount", 0U);
  out.revisions_password = attr_str(node, "revisionsPassword");
  out.revisions_algorithm_name = attr_str(node, "revisionsAlgorithmName");
  out.revisions_hash_value = attr_str(node, "revisionsHashValue");
  out.revisions_salt_value = attr_str(node, "revisionsSaltValue");
  out.revisions_spin_count = attr_u32(node, "revisionsSpinCount", 0U);
  return out;
}

std::string apply_workbook_protection(std::string_view existing, const WorkbookProtection& protection) {
  pugi::xml_document doc;
  const std::string name(kWorkbookProtection);
  if (existing.empty() || !doc.load_buffer(existing.data(), existing.size()) || !doc.child(name.c_str())) {
    doc.reset();
    doc.append_child(name.c_str());
  }
  pugi::xml_node node = doc.child(name.c_str());
  SetTextAttr(node, "workbookPassword", protection.workbook_password);
  SetTextAttr(node, "workbookAlgorithmName", protection.workbook_algorithm_name);
  SetTextAttr(node, "workbookHashValue", protection.workbook_hash_value);
  SetTextAttr(node, "workbookSaltValue", protection.workbook_salt_value);
  SetUintAttr(node, "workbookSpinCount", protection.workbook_spin_count);
  SetTextAttr(node, "revisionsPassword", protection.revisions_password);
  SetTextAttr(node, "revisionsAlgorithmName", protection.revisions_algorithm_name);
  SetTextAttr(node, "revisionsHashValue", protection.revisions_hash_value);
  SetTextAttr(node, "revisionsSaltValue", protection.revisions_salt_value);
  SetUintAttr(node, "revisionsSpinCount", protection.revisions_spin_count);
  SetBoolAttr(node, "lockStructure", protection.lock_structure);
  SetBoolAttr(node, "lockWindows", protection.lock_windows);
  SetBoolAttr(node, "lockRevision", protection.lock_revision);
  if (!node.first_attribute() && !node.first_child()) {
    return {};
  }
  return raw_xml(node);
}

std::vector<ProtectedRange> read_protected_ranges(const WorksheetRawExtensions& children) {
  std::vector<ProtectedRange> out;
  for (const WorksheetRawChild& child : children) {
    if (!IsProtectedRanges(child)) {
      continue;
    }
    pugi::xml_document doc;
    if (!doc.load_buffer(child.xml.data(), child.xml.size())) {
      continue;
    }
    for (pugi::xml_node node : doc.child(std::string(kProtectedRanges).c_str()).children("protectedRange")) {
      ProtectedRange range;
      range.name = attr_str(node, "name");
      range.sqref = attr_str(node, "sqref");
      range.password = attr_str(node, "password");
      range.algorithm_name = attr_str(node, "algorithmName");
      range.hash_value = attr_str(node, "hashValue");
      range.salt_value = attr_str(node, "saltValue");
      range.spin_count = attr_u32(node, "spinCount", 0U);
      range.security_descriptor = attr_str(node, "securityDescriptor");
      out.push_back(std::move(range));
    }
  }
  return out;
}

void write_protected_ranges(WorksheetRawExtensions& children, const std::vector<ProtectedRange>& ranges) {
  // Child elements of the existing ranges, by range name, re-emitted with the range.
  std::unordered_map<std::string, std::string> kept_children;
  for (const WorksheetRawChild& child : children) {
    pugi::xml_document doc;
    if (!IsProtectedRanges(child) || !doc.load_buffer(child.xml.data(), child.xml.size())) {
      continue;
    }
    for (pugi::xml_node node : doc.child(std::string(kProtectedRanges).c_str()).children("protectedRange")) {
      std::string inner;
      for (pugi::xml_node element : node.children()) {
        append_raw_xml(inner, element);
      }
      if (!inner.empty()) {
        kept_children.emplace(std::string(attr_str(node, "name")), std::move(inner));
      }
    }
  }
  for (std::size_t i = children.size(); i-- > 0;) {
    if (IsProtectedRanges(children[i])) {
      children.erase(children.begin() + static_cast<std::ptrdiff_t>(i));
    }
  }
  if (ranges.empty()) {
    return;
  }
  std::string xml = "<protectedRanges>";
  for (const ProtectedRange& range : ranges) {
    xml.append("<protectedRange");
    if (!range.password.empty()) {
      append_xml_attr(xml, "password", range.password);
    }
    append_xml_attr(xml, "sqref", range.sqref);
    if (!range.security_descriptor.empty()) {
      append_xml_attr(xml, "securityDescriptor", range.security_descriptor);
    }
    append_xml_attr(xml, "name", range.name);
    if (!range.algorithm_name.empty()) {
      append_xml_attr(xml, "algorithmName", range.algorithm_name);
    }
    if (!range.hash_value.empty()) {
      append_xml_attr(xml, "hashValue", range.hash_value);
    }
    if (!range.salt_value.empty()) {
      append_xml_attr(xml, "saltValue", range.salt_value);
    }
    if (range.spin_count != 0U) {
      append_xml_attr_uint(xml, "spinCount", range.spin_count);
    }
    if (const auto it = kept_children.find(range.name); it != kept_children.end()) {
      xml.append(">").append(it->second).append("</protectedRange>");
    } else {
      xml.append("/>");
    }
  }
  xml.append("</protectedRanges>");
  const auto slot = static_cast<std::uint32_t>(worksheet_child::slot_of(kProtectedRanges));
  std::size_t at = 0;
  while (at < children.size() && children[at].slot <= slot) {
    ++at;
  }
  children.insert(children.begin() + static_cast<std::ptrdiff_t>(at), WorksheetRawChild{slot, std::move(xml)});
}

}  // namespace formulon::io
