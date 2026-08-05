//
// Implementation of the OOXML writer's relationship / override helpers.

#include "io/ooxml/relationship_writer.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "io/xml_escape.h"

namespace formulon {
namespace io {

std::string RelsPathForPart(std::string_view part_path) {
  const std::size_t slash = part_path.find_last_of('/');
  if (slash == std::string_view::npos) {
    std::string rels_path("_rels/");
    rels_path.append(part_path);
    rels_path.append(".rels");
    return rels_path;
  }
  std::string rels_path;
  rels_path.append(part_path.substr(0, slash));
  rels_path.append("/_rels/");
  rels_path.append(part_path.substr(slash + 1));
  rels_path.append(".rels");
  return rels_path;
}

std::string_view WithoutXlPrefix(std::string_view path) {
  constexpr std::string_view kXlPrefix = "xl/";
  if (path.size() >= kXlPrefix.size() && path.substr(0, kXlPrefix.size()) == kXlPrefix) {
    path.remove_prefix(kXlPrefix.size());
  }
  return path;
}

std::string TargetRelativeToWorksheet(std::string_view package_path) {
  constexpr std::string_view kXlPrefix = "xl/";
  std::string out;
  if (package_path.size() >= kXlPrefix.size() && package_path.substr(0, kXlPrefix.size()) == kXlPrefix) {
    out.assign("../");
    out.append(package_path.substr(kXlPrefix.size()));
    return out;
  }
  out.assign(package_path);
  return out;
}

void AppendOverride(std::string& out, std::string_view path, std::string_view ct, bool escape_path) {
  out.append("  <Override PartName=\"/");
  if (escape_path) {
    AppendXmlEscaped(out, path);
  } else {
    out.append(path.data(), path.size());
  }
  out.append("\" ContentType=\"");
  out.append(ct.data(), ct.size());
  out.append("\"/>\n");
}

void AppendRelationship(std::string& out, std::string_view id, std::string_view type, std::string_view target,
                        bool target_external, bool escape_target) {
  out.append("  <Relationship Id=\"");
  AppendXmlEscaped(out, id);
  out.append("\" Type=\"");
  out.append(type);
  out.append("\" Target=\"");
  if (escape_target) {
    AppendXmlEscaped(out, target);
  } else {
    out.append(target);
  }
  if (target_external) {
    out.append("\" TargetMode=\"External\"/>\n");
  } else {
    out.append("\"/>\n");
  }
}

void AppendRelationship(std::string& out, std::uint32_t rid, std::string_view type, std::string_view target,
                        bool target_external, bool escape_target) {
  AppendRelationship(out, "rId" + std::to_string(rid), type, target, target_external, escape_target);
}

}  // namespace io
}  // namespace formulon
