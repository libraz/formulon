// @size-budget: 12 KB
//
// Opaque rich values: the strict reader, the cell restore, the fallback
// diagnosis and the richValueRel save plan. See rich_values.h.

#include "io/rich_values.h"

#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/part_paths.h"
#include "io/xml_utils.h"
#include "io/xsd_int.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kXmlDeclCrLf = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n";
constexpr std::string_view kRichValueRelKeyPrefix = "_rvRel:";
constexpr std::string_view kRichDataDir = "xl/richData";

// True when every attribute of `node` is in `allowed` or declares a namespace.
bool OnlyAttributes(const pugi::xml_node& node, std::initializer_list<std::string_view> allowed) {
  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    const std::string_view name = attr.name();
    bool known = strings::starts_with(name, "xmlns");
    for (const std::string_view a : allowed) {
      known = known || name == a;
    }
    if (!known) {
      return false;
    }
  }
  return true;
}

// Key types whose values are self-contained text: decimal (absent or `d`),
// integer, boolean, error, string, and a supporting-property bag index.
bool CarriedKeyType(const OpaqueRichValueKey& key) {
  const std::string_view t = key.type;
  if (is_rich_value_rel_key(key.name)) {
    return t == "i";
  }
  return t.empty() || t == "d" || t == "i" || t == "b" || t == "e" || t == "s" || t == "spb";
}

struct StructureDecl {
  std::string type;
  std::vector<OpaqueRichValueKey> keys;
};

std::string_view ReadStructures(const pugi::xml_node& root, std::vector<StructureDecl>* out) {
  for (pugi::xml_node s : root.children()) {
    if (s.type() != pugi::node_element) {
      continue;
    }
    if (local_name(s) != "s" || !OnlyAttributes(s, {"t"})) {
      return kRichFallbackValueShape;
    }
    StructureDecl decl;
    decl.type = s.attribute("t").value();
    for (pugi::xml_node k : s.children()) {
      if (k.type() != pugi::node_element) {
        continue;
      }
      if (local_name(k) != "k" || !k.attribute("n") || !OnlyAttributes(k, {"n", "t", "s"})) {
        return kRichFallbackValueShape;
      }
      OpaqueRichValueKey key;
      key.name = k.attribute("n").value();
      key.type = k.attribute("t").value();
      key.style = k.attribute("s").value();
      if (decl.type != "_error" && !CarriedKeyType(key)) {
        return kRichFallbackKeyType;
      }
      decl.keys.push_back(std::move(key));
    }
    out->push_back(std::move(decl));
  }
  return {};
}

// One `<rel>` of `richValueRel.xml`, resolved through its rels part.
struct RelEntry {
  std::string target;
  std::string type;
  bool external = false;
};

std::string_view LoadRels(const RichValuePackage& p, std::vector<RelEntry>* out) {
  if (p.rel == nullptr || p.rel_rels == nullptr || !p.rel_registered) {
    return kRichFallbackRelPart;
  }
  pugi::xml_document rel_doc;
  pugi::xml_document rels_doc;
  if (!rel_doc.load_buffer(p.rel->data(), p.rel->size(), pugi::parse_default, pugi::encoding_utf8) ||
      !rels_doc.load_buffer(p.rel_rels->data(), p.rel_rels->size(), pugi::parse_default, pugi::encoding_utf8) ||
      !ooxml::relationship_root(rels_doc)) {
    return kRichFallbackRelPart;
  }
  std::unordered_map<std::string, RelEntry> by_id;
  bool resolved_all = true;
  const auto visited = ooxml::visit_relationship_nodes(rels_doc, [&](const pugi::xml_node& rel) {
    RelEntry entry;
    entry.type = rel.attribute("Type").value();
    entry.external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    if (!entry.external) {
      auto target = ooxml::resolve_relative_path(kRichDataDir, rel.attribute("Target").value());
      if (!target) {
        resolved_all = false;
        return Expected<void, Error>::Ok();
      }
      entry.target = std::move(target.value());
    }
    by_id[rel.attribute("Id").value()] = std::move(entry);
    return Expected<void, Error>::Ok();
  });
  const pugi::xml_node root = child_local(rel_doc, "richValueRels");
  if (!visited || !resolved_all || !root) {
    return kRichFallbackRelPart;
  }
  for (pugi::xml_node rel : root.children()) {
    if (rel.type() != pugi::node_element) {
      continue;
    }
    const auto it = by_id.find(ooxml::relationship_ref_id(rel));
    if (local_name(rel) != "rel" || it == by_id.end()) {
      return kRichFallbackRelPart;
    }
    out->push_back(it->second);
  }
  return {};
}

// Reads the values of `p` into `set`; returns the fallback reason, or empty.
std::string_view ReadValues(const RichValuePackage& p, RichValueSet* set) {
  pugi::xml_document values_doc;
  pugi::xml_document structures_doc;
  // A whitespace-only string value is the sole child of its `<v>`, so it is kept.
  if (!values_doc.load_buffer(p.values->data(), p.values->size(), pugi::parse_default | pugi::parse_ws_pcdata_single,
                              pugi::encoding_utf8) ||
      !structures_doc.load_buffer(p.structures->data(), p.structures->size(), pugi::parse_default,
                                  pugi::encoding_utf8)) {
    return kRichFallbackValueParts;
  }
  const pugi::xml_node values_root = child_local(values_doc, "rvData");
  const pugi::xml_node structures_root = child_local(structures_doc, "rvStructures");
  if (!values_root || !structures_root) {
    return kRichFallbackValueParts;
  }
  std::vector<StructureDecl> structures;
  if (const std::string_view reason = ReadStructures(structures_root, &structures); !reason.empty()) {
    return reason;
  }
  const std::vector<std::optional<RichError>> errors = parse_rich_errors(*p.values, *p.structures);
  std::optional<std::vector<RelEntry>> rels;
  std::size_t index = 0;
  for (pugi::xml_node rv : values_root.children()) {
    if (rv.type() != pugi::node_element) {
      continue;
    }
    if (local_name(rv) != "rv") {
      return kRichFallbackValueShape;
    }
    if (!OnlyAttributes(rv, {"s"})) {
      return kRichFallbackRvAttribute;
    }
    std::uint32_t s = 0;
    if (!parse_xsd_nonneg_int(rv.attribute("s").value(), &s) || s >= structures.size()) {
      return kRichFallbackValueShape;
    }
    if (structures[s].type == "_error") {
      if (index >= errors.size() || !errors[index].has_value()) {
        return kRichFallbackErrorValue;
      }
      set->errors.values.push_back(*errors[index]);
      set->opaque.emplace_back();
      ++index;
      continue;
    }
    OpaqueRichValue value;
    value.structure_type = structures[s].type;
    value.keys = structures[s].keys;
    std::size_t k = 0;
    for (pugi::xml_node v : rv.children()) {
      if (v.type() != pugi::node_element || local_name(v) != "v" || v.first_attribute() || has_element_child(v) ||
          k >= value.keys.size()) {
        return kRichFallbackValueShape;
      }
      OpaqueRichValueKey& key = value.keys[k++];
      if (!is_rich_value_rel_key(key.name)) {
        key.value = v.text().get();
        continue;
      }
      if (!rels.has_value()) {
        rels.emplace();
        if (const std::string_view reason = LoadRels(p, &*rels); !reason.empty()) {
          return reason;
        }
      }
      std::uint32_t rel = 0;
      if (!parse_xsd_nonneg_int(v.text().get(), &rel) || rel >= rels->size()) {
        return kRichFallbackRelPart;
      }
      if ((*rels)[rel].external) {
        return kRichFallbackRelExternal;
      }
      key.rel_target = (*rels)[rel].target;
      key.rel_type = (*rels)[rel].type;
    }
    if (k != value.keys.size()) {
      return kRichFallbackValueShape;
    }
    set->errors.values.push_back(RichError{});
    set->opaque.push_back(std::move(value));
    ++index;
  }
  return {};
}

std::string_view ReadInto(const RichValuePackage& p, RichValueMetadata info, RichValueSet* set) {
  if (!p.unknown_part.empty()) {
    return kRichFallbackRichDataPart;
  }
  if (!info.consumable) {
    return info.fallback_reason.empty() ? kRichFallbackMetadataBlock : info.fallback_reason;
  }
  if (p.values != nullptr || !info.vm_rv.empty()) {
    if (p.values == nullptr || p.structures == nullptr) {
      return kRichFallbackValueParts;
    }
    if (const std::string_view reason = ReadValues(p, set); !reason.empty()) {
      return reason;
    }
  }
  for (const std::uint32_t rv : info.vm_rv) {
    if (rv >= set->errors.values.size()) {
      return kRichFallbackVmIndex;
    }
  }
  set->errors.vm_rv = std::move(info.vm_rv);
  return {};
}

// The internal targets of the relationship part `rels_path`, resolved against its source part's folder.
std::vector<std::string> RelationshipTargets(const std::vector<std::uint8_t>& bytes, std::string_view rels_path) {
  std::vector<std::string> out;
  pugi::xml_document doc;
  if (!doc.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8) ||
      !ooxml::relationship_root(doc)) {
    return out;
  }
  const std::string base = ooxml::dir_of(ooxml::dir_of(rels_path));
  (void)ooxml::visit_relationship_nodes(doc, [&](const pugi::xml_node& rel) {
    if (std::string_view(rel.attribute("TargetMode").value()) != "External") {
      if (auto target = ooxml::resolve_relative_path(base, rel.attribute("Target").value())) {
        out.push_back(std::move(target.value()));
      }
    }
    return Expected<void, Error>::Ok();
  });
  return out;
}

// `path` as a target relative to `xl/richData/`.
std::string RelativeToRichData(std::string_view path) {
  constexpr std::string_view kXl = "xl/";
  if (strings::starts_with(path, kRichDataDirPrefix)) {
    return std::string(path.substr(kRichDataDirPrefix.size()));
  }
  if (strings::starts_with(path, kXl)) {
    return "../" + std::string(path.substr(kXl.size()));
  }
  return "../../" + std::string(path);
}

std::vector<std::uint8_t> Bytes(const std::string& text) {
  return std::vector<std::uint8_t>(text.begin(), text.end());
}

}  // namespace

bool is_rich_value_rel_key(std::string_view name) {
  return strings::starts_with(name, kRichValueRelKeyPrefix);
}

bool rich_value_parts_kept_verbatim(const std::vector<PassthroughPart>& parts) {
  for (const PassthroughPart& part : parts) {
    if (part.path == "xl/metadata.xml" || part.path == kRdRichValuePartPath ||
        part.path == kRdRichValueStructurePartPath) {
      return true;
    }
  }
  return false;
}

bool is_retained_rich_data_part(std::string_view path) {
  return path == kRichValueRelPartPath || path == kRichValueRelRelsPartPath ||
         path == "xl/richData/rdsupportingpropertybag.xml" ||
         path == "xl/richData/rdsupportingpropertybagstructure.xml" || path == "xl/richData/richStyles.xml";
}

bool rich_value_rel_registered(const std::vector<UnknownRelationship>& rels) {
  for (const UnknownRelationship& r : rels) {
    if (!r.target_external && r.type == kRelRichValueRel && r.target == kRichValueRelPartPath) {
      return true;
    }
  }
  return false;
}

RichValueRead read_rich_value_parts(const RichValuePackage& package) {
  RichValueMetadata info;
  if (package.metadata == nullptr) {
    info.fallback_reason = kRichFallbackMetadataMissing;
  } else {
    info = parse_metadata_xml(*package.metadata);
  }
  return read_rich_value_parts(package, std::move(info));
}

RichValueRead read_rich_value_parts(const RichValuePackage& package, RichValueMetadata metadata) {
  RichValueRead out;
  out.fallback_reason = ReadInto(package, std::move(metadata), &out.set);
  if (!out.fallback_reason.empty()) {
    out.set = RichValueSet{};
  }
  return out;
}

std::vector<RichErrorCell> restore_opaque_rich_values(Sheet& sheet, const std::vector<RichErrorCell>& cells,
                                                      const RichValueSet& set) {
  std::vector<RichErrorCell> rest;
  for (const RichErrorCell& c : cells) {
    const std::vector<std::uint32_t>& vm_rv = set.errors.vm_rv;
    const std::size_t rv = c.vm != 0U && c.vm <= vm_rv.size() ? vm_rv[c.vm - 1U] : set.opaque.size();
    if (rv >= set.opaque.size() || !set.opaque[rv].has_value()) {
      rest.push_back(c);
      continue;
    }
    const Cell* cell = sheet.cell_at(c.row, c.col);
    if (cell != nullptr && cell->cached_value.is_error() && cell->cached_value.as_error() == ErrorCode::Value) {
      sheet.set_rich_value(c.row, c.col, *set.opaque[rv]);
    }
  }
  return rest;
}

std::string_view rich_value_fallback_reason(const Workbook& wb) {
  RichValuePackage p;
  const std::vector<std::uint8_t>* metadata_bin = nullptr;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == "xl/metadata.xml") {
      p.metadata = &part.bytes;
    } else if (part.path == xlsb::kMetadataBinPath) {
      metadata_bin = &part.bytes;
    } else if (part.path == kRdRichValuePartPath) {
      p.values = &part.bytes;
    } else if (part.path == kRdRichValueStructurePartPath) {
      p.structures = &part.bytes;
    } else if (part.path == kRichValueRelPartPath) {
      p.rel = &part.bytes;
    } else if (part.path == kRichValueRelRelsPartPath) {
      p.rel_rels = &part.bytes;
    } else if (strings::starts_with(part.path, kRichDataDirPrefix) && part.path != kRdRichValueTypesPartPath &&
               !is_retained_rich_data_part(part.path) && p.unknown_part.empty()) {
      p.unknown_part = part.path;
    }
  }
  if (p.values == nullptr && p.structures == nullptr) {
    return {};
  }
  p.rel_registered = rich_value_rel_registered(wb.unknown_workbook_rels());
  const RichValueRead read =
      metadata_bin != nullptr
          ? read_rich_value_parts(p, xlsb::parse_metadata_bin(ByteSpan{metadata_bin->data(), metadata_bin->size()}))
          : read_rich_value_parts(p);
  return read.fallback_reason.empty() ? kRichFallbackPackage : read.fallback_reason;
}

RichValueRelPlan plan_rich_value_rels(const Workbook& wb, const RichErrorTable& table) {
  RichValueRelPlan out;
  const std::vector<PassthroughPart>& parts = wb.passthrough_parts();
  if (rich_value_parts_kept_verbatim(parts)) {
    return out;
  }
  const PassthroughPart* source_rels = nullptr;
  for (const PassthroughPart& part : parts) {
    if (part.path == kRichValueRelPartPath || part.path == kRichValueRelRelsPartPath) {
      out.skipped.insert(part.path);
      source_rels = part.path == kRichValueRelRelsPartPath ? &part : source_rels;
    }
  }
  const std::vector<std::string> listed = source_rels != nullptr
                                              ? RelationshipTargets(source_rels->bytes, kRichValueRelRelsPartPath)
                                              : std::vector<std::string>();
  if (!listed.empty()) {
    // Everything else that still names a part: the values, every other
    // relationship part, and the relationships the model keeps.
    std::unordered_set<std::string> referenced;
    for (const auto& rel : table.rels) {
      referenced.insert(rel.first);
    }
    for (const PassthroughPart& part : parts) {
      if (part.path != kRichValueRelRelsPartPath && strings::ends_with(part.path, ".rels")) {
        for (std::string& target : RelationshipTargets(part.bytes, part.path)) {
          referenced.insert(std::move(target));
        }
      }
    }
    const auto add_model_rels = [&referenced](const std::vector<UnknownRelationship>& rels) {
      for (const UnknownRelationship& r : rels) {
        if (!r.target_external) {
          referenced.insert(r.target);
        }
      }
    };
    add_model_rels(wb.unknown_workbook_rels());
    for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
      add_model_rels(wb.sheet(s).unknown_relationships());
    }
    for (const std::string& target : listed) {
      if (referenced.count(target) == 0U) {
        out.skipped.insert(target);
      }
    }
  }
  if (table.rels.empty()) {
    return out;
  }
  std::string rel_xml(kXmlDeclCrLf);
  rel_xml.append("<richValueRels");
  append_xml_ns_decl(rel_xml, "", kNsRichValueRel);
  append_xml_ns_decl(rel_xml, "r", kNsOfficeRelationships);
  rel_xml.push_back('>');
  std::string rels_xml(kXmlDecl);
  BeginRelationships(rels_xml);
  for (std::size_t i = 0; i < table.rels.size(); ++i) {
    const std::uint32_t rid = static_cast<std::uint32_t>(i + 1U);
    rel_xml.append("<rel r:id=\"rId").append(std::to_string(rid)).append("\"/>");
    AppendRelationship(rels_xml, rid, table.rels[i].second, RelativeToRichData(table.rels[i].first),
                       /*target_external=*/false, /*escape_target=*/true);
  }
  rel_xml.append("</richValueRels>");
  EndRelationships(rels_xml);
  out.parts.emplace_back(std::string(kRichValueRelPartPath), std::string(kCtRichValueRel), Bytes(rel_xml));
  out.parts.emplace_back(std::string(kRichValueRelRelsPartPath), std::string(), Bytes(rels_xml));
  return out;
}

}  // namespace io
}  // namespace formulon
