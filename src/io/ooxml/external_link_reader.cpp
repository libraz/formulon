
#include "io/ooxml/external_link_reader.h"

#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/external_links.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/ooxml_defs.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace ooxml {

ExternalLinkLoadResult load_external_links(const ZipReader& zip, const pugi::xml_node& wb_root,
                                           const WorkbookRels& wb_rels) {
  ExternalLinkLoadResult out;
  pugi::xml_node refs_node = wb_root.child("externalReferences");
  if (!refs_node) {
    return out;
  }
  std::uint32_t index = 1;
  for (pugi::xml_node ref = refs_node.child("externalReference"); ref;
       ref = ref.next_sibling("externalReference"), ++index) {
    std::string rid = relationship_ref_id(ref);
    if (rid.empty()) {
      continue;
    }
    auto it = wb_rels.external_link_paths_by_rid.find(rid);
    if (it == wb_rels.external_link_paths_by_rid.end()) {
      continue;
    }
    ExternalLinkRecord rec;
    rec.index = index;
    rec.rel_id = std::move(rid);
    rec.part_path = it->second;
    rec.kind = ExternalLinkRecord::Kind::kUnknown;

    // Body part — detect kind and capture the inner r:id reference.
    if (zip.has_entry(rec.part_path)) {
      auto body_or = zip.read_entry(rec.part_path);
      if (body_or) {
        const std::vector<std::uint8_t>& body_bytes = body_or.value();
        pugi::xml_document body_doc;
        pugi::xml_parse_result body_parse =
            body_doc.load_buffer(body_bytes.data(), body_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
        if (body_parse) {
          pugi::xml_node link_root = body_doc.child("externalLink");
          if (link_root) {
            if (pugi::xml_node book = link_root.child("externalBook"); book) {
              rec.kind = ExternalLinkRecord::Kind::kExternalBook;
              rec.body_rel_id = relationship_ref_id(book);
            } else if (pugi::xml_node ole = link_root.child("oleLink"); ole) {
              rec.kind = ExternalLinkRecord::Kind::kOleLink;
              rec.body_rel_id = relationship_ref_id(ole);
            } else if (link_root.child("ddeLink")) {
              rec.kind = ExternalLinkRecord::Kind::kDdeLink;
              // ddeLink carries its connection metadata inline; no inner r:id.
            }
          }
        }
      }
    }

    // Per-link rels — capture target URL + target_mode for round-trip.
    const std::string body_rels_path = rels_path_for_part(rec.part_path);
    if (zip.has_entry(body_rels_path)) {
      auto rels_or = zip.read_entry(body_rels_path);
      if (rels_or) {
        const std::vector<std::uint8_t>& rels_bytes = rels_or.value();
        pugi::xml_document rels_doc;
        pugi::xml_parse_result rels_parse =
            rels_doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
        if (rels_parse) {
          pugi::xml_node rels_root = rels_doc.child("Relationships");
          if (rels_root) {
            // Pick the relationship whose Id matches the body's inner
            // r:id when available; otherwise take the first link-typed
            // relationship as a best-effort fallback.
            for (pugi::xml_node rel = rels_root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
              const std::string_view type = rel.attribute("Type").value();
              if (type != kRelExternalLinkPath && type != kRelOleLink && type != kRelDdeLink) {
                continue;
              }
              const std::string_view rel_id = rel.attribute("Id").value();
              const bool id_match = !rec.body_rel_id.empty() && rel_id == rec.body_rel_id;
              if (rec.target.empty() || id_match) {
                rec.target = rel.attribute("Target").value();
                const std::string_view target_mode = rel.attribute("TargetMode").value();
                rec.target_external = (target_mode == "External") || target_mode.empty();
                if (rec.body_rel_id.empty()) {
                  rec.body_rel_id = rel_id;
                }
                if (id_match) {
                  break;
                }
              }
            }
          }
        }
      }
      out.consumed_rels_paths.push_back(body_rels_path);
    }
    out.records.push_back(std::move(rec));
  }
  return out;
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
