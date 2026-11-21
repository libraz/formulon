//
// Implementation of the OOXML writer's relationship / override helpers.

#include "io/ooxml/relationship_writer.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_set>

#include "io/ooxml_defs.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "passthrough_part.h"
#include "unknown_relationship.h"
#include "utils/strings.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kXlPrefix = "xl/";

}  // namespace

std::string TargetRelativeToWorkbook(std::string_view path) {
  if (strings::starts_with(path, kXlPrefix)) {
    return std::string(path.substr(kXlPrefix.size()));
  }
  std::string out("../");
  out.append(path);
  return out;
}

std::string TargetRelativeToWorksheet(std::string_view package_path) {
  std::string out;
  if (strings::starts_with(package_path, kXlPrefix)) {
    out.assign("../");
    out.append(package_path.substr(kXlPrefix.size()));
    return out;
  }
  out.assign("../../");
  out.append(package_path);
  return out;
}

void BeginRelationships(std::string& out) {
  out.append("<Relationships xmlns=\"");
  out.append(kNsPackageRelationships);
  out.append("\">\n");
}

void EndRelationships(std::string& out) {
  out.append("</Relationships>\n");
}

void AppendOverride(std::string& out, std::string_view path, std::string_view ct, bool escape_path,
                    bool escape_content_type) {
  out.append("  <Override PartName=\"/");
  if (escape_path) {
    AppendXmlAttrEscaped(out, path);
  } else {
    out.append(path.data(), path.size());
  }
  out.append("\" ContentType=\"");
  if (escape_content_type) {
    AppendXmlAttrEscaped(out, ct);
  } else {
    out.append(ct.data(), ct.size());
  }
  out.append("\"/>\n");
}

void AppendPassthroughOverrides(std::string& out, const std::vector<const PassthroughPart*>& parts) {
  for (const PassthroughPart* part : parts) {
    if (part->content_type.empty()) {
      continue;
    }
    AppendOverride(out, part->path, part->content_type, /*escape_path=*/true,
                   /*escape_content_type=*/true);
  }
}

void AppendDefault(std::string& out, std::string_view extension, std::string_view ct) {
  out.append("  <Default");
  append_xml_attr(out, "Extension", extension);
  append_xml_attr(out, "ContentType", ct);
  out.append("/>\n");
}

void AppendRelationship(std::string& out, std::string_view id, std::string_view type, std::string_view target,
                        bool target_external, bool escape_target) {
  out.append("  <Relationship");
  append_xml_attr(out, "Id", id);
  append_xml_attr(out, "Type", type);
  out.append(" Target=\"");
  if (escape_target) {
    // `Target` is an `xsd:anyURI` the rels reader takes verbatim from the
    // parser, so it gets the attribute rule and no OOXML escaping.
    AppendXmlAttrEscaped(out, target);
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
  AppendRelationship(out, rel_id_text(rid), type, target, target_external, escape_target);
}

std::string EncodeTargetSpaces(std::string_view target) {
  std::string out;
  out.reserve(target.size());
  for (const char c : target) {
    if (c == ' ') {
      out.append("%20");
    } else {
      out.push_back(c);
    }
  }
  return out;
}

void RecordDroppedRelationship(std::string_view event, std::string_view type, std::string_view target,
                               WriteDiagnostics* diagnostics) {
  StructuredLog(event)
      .field("reason", std::string_view("target_part_absent"))
      .field("type", type)
      .field("target", target)
      .warn();
  if (diagnostics != nullptr) {
    diagnostics->unwritten_relationship_targets.emplace_back(target);
  }
}

void RecordDroppedPart(std::string_view event, std::string_view path, std::string_view reason,
                       WriteDiagnostics* diagnostics) {
  StructuredLog(event).field("path", path).field("reason", reason).warn();
  CountDroppedPart(path, diagnostics);
}

void CountDroppedPart(std::string_view path, WriteDiagnostics* diagnostics) {
  if (diagnostics != nullptr) {
    ++diagnostics->dropped_part_count;
    diagnostics->dropped_part_paths.emplace_back(path);
  }
}

void SettleDroppedRelationships(WriteDiagnostics& diagnostics) {
  const std::unordered_set<std::string_view> dropped(diagnostics.dropped_part_paths.begin(),
                                                     diagnostics.dropped_part_paths.end());
  diagnostics.dropped_relationship_count = 0;
  for (const std::string& target : diagnostics.unwritten_relationship_targets) {
    diagnostics.dropped_relationship_count += dropped.count(target) != 0U ? 1U : 0U;
  }
}

}  // namespace io
}  // namespace formulon
