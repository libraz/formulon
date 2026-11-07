//
// Implementation of the XLSB package envelope walkers. See
// `io/xlsb/package_loader.h`.

#include "io/xlsb/package_loader.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "io/ooxml/package_rels_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace detail {
namespace {

// ---------------------------------------------------------------------------
// XML envelope: same shape as the OOXML reader, but the relationship and
// content types target binary parts (`*.bin`) instead of XML.
// ---------------------------------------------------------------------------

/// Drops the leading `/` from a part path (relationship targets can be
/// absolute or relative, but the ZIP catalogue uses relative names).
std::string NormalisePartName(std::string_view raw) {
  if (!raw.empty() && raw.front() == '/') {
    return std::string(raw.substr(1));
  }
  return std::string(raw);
}

/// Resolves an OOXML rels target relative to a base directory, normalising
/// `.` / `..` segments. Thin wrapper over the OOXML package validator's
/// shared helper so both readers share a single path-traversal defence.
///
/// Path-traversal hardening: a `Target` with excess `..` segments or a
/// leading `/` (package-absolute) is refused with `kIoZipSlip`. The
/// surfaced error context is the validator's standard `context=ooxml_reader`
/// flavour; xlsb's xml envelope is structurally identical to xlsx, so the
/// shared diagnostic stays accurate for both.
Expected<std::string, Error> ResolveRelativePath(std::string_view base_dir, std::string_view target) {
  return ooxml::resolve_relative_path(base_dir, target);
}

}  // namespace

// ---------------------------------------------------------------------------
// XML envelope walkers (return `Expected` on every failure path).
// ---------------------------------------------------------------------------

Expected<ContentTypesView, Error> LoadContentTypes(const std::vector<std::uint8_t>& ct_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, ct_bytes, "xlsb_reader", "[Content_Types].xml"));
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=xlsb_reader part=[Content_Types].xml");
  }
  ContentTypesView view;
  std::unordered_map<std::string, std::size_t> default_indices;
  std::unordered_map<std::string, std::size_t> override_indices;
  bool saw_workbook = false;
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    const std::string_view node_name(node.name());
    if (node_name == "Default") {
      // Real Excel-365 output does not always carry a per-part
      // `<Override>` for `xl/workbook.bin`: the workbook part can rely
      // on the blanket `<Default Extension="bin">` registration alone
      // (verified against a real Excel-365-produced `.xlsb`). Accept
      // that form too, rather than requiring an Override that a
      // genuinely valid package may not emit.
      const std::string_view ext(node.attribute("Extension").value());
      const std::string_view ct(node.attribute("ContentType").value());
      if (strings::case_insensitive_eq(ext, "bin") && is_xlsb_workbook_content_type(ct)) {
        saw_workbook = true;
      }
      if (ext.empty()) {
        continue;
      }
      const std::string normalized_ext = strings::to_ascii_lower(ext);
      const auto [index_it, inserted] = default_indices.try_emplace(normalized_ext, view.defaults.size());
      if (!inserted) {
        if (view.defaults[index_it->second].content_type != ct) {
          return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                            "[Content_Types].xml: conflicting Default content types for extension " + normalized_ext,
                            "context=xlsb_reader part=[Content_Types].xml extension=" + normalized_ext);
        }
        continue;
      }
      view.defaults.push_back(DefaultContentType{normalized_ext, std::string(ct)});
      continue;
    }
    if (node_name != "Override") {
      continue;
    }
    std::string part_name = NormalisePartName(node.attribute("PartName").value());
    if (part_name.empty()) {
      continue;
    }
    std::string ct = node.attribute("ContentType").value();
    if (is_xlsb_workbook_content_type(ct)) {
      saw_workbook = true;
    }
    const auto [index_it, inserted] = override_indices.try_emplace(part_name, view.overrides.size());
    if (!inserted) {
      if (view.overrides[index_it->second].second != ct) {
        return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                          "[Content_Types].xml: conflicting Override content types for part " + part_name,
                          "context=xlsb_reader part=[Content_Types].xml override=" + part_name);
      }
      continue;
    }
    view.overrides.emplace_back(std::move(part_name), std::move(ct));
  }
  if (!saw_workbook) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "[Content_Types].xml: no xlsb-workbook content-type override",
                      "context=xlsb_reader part=[Content_Types].xml");
  }
  return view;
}

Expected<std::string, Error> ResolveOfficeDocumentPath(const std::vector<std::uint8_t>& rels_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=xlsb_reader part=_rels/.rels desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "package-level rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = ooxml::relationship_root(doc);
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=xlsb_reader part=_rels/.rels");
  }
  for (const pugi::xml_node rel : root.children()) {
    if (!ooxml::is_relationship_element(rel, root)) {
      continue;
    }
    if (std::string_view(rel.attribute("Type").value()) == kRelOfficeDocument) {
      std::string target = NormalisePartName(rel.attribute("Target").value());
      if (target.empty()) {
        return make_error(FormulonErrorCode::kIoRelationshipBroken,
                          "package-level rels: empty Target for OfficeDocument",
                          "context=xlsb_reader part=_rels/.rels");
      }
      return target;
    }
  }
  return make_error(FormulonErrorCode::kIoRelationshipBroken,
                    "package-level rels: no OfficeDocument relationship found", "context=xlsb_reader part=_rels/.rels");
}

Expected<WorkbookRels, Error> LoadWorkbookRels(const ZipReader& zip, std::string_view workbook_path) {
  WorkbookRels rels;
  const std::string rels_path = ooxml::rels_path_for_part(workbook_path);
  if (!zip.has_entry(rels_path)) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=xlsb_reader rels_path=" + rels_path);
  }
  const std::string base_dir = ooxml::dir_of(workbook_path);
  RETURN_IF_ERROR(ooxml::visit_relationship_nodes(
      zip, rels_path, "workbook rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        const std::string_view type = rel.attribute("Type").value();
        const std::string_view target = rel.attribute("Target").value();
        if (target.empty()) {
          return Expected<void, Error>::Ok();
        }
        if (type == kRelWorksheet) {
          const std::string id = rel.attribute("Id").value();
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          ASSIGN_OR_RETURN(auto resolved, ResolveRelativePath(base_dir, target));
          rels.sheet_targets.emplace(id, std::move(resolved));
        } else if (type == kRelSharedStrings) {
          ASSIGN_OR_RETURN(rels.sst_path, ResolveRelativePath(base_dir, target));
        } else if (type == kRelStyles) {
          ASSIGN_OR_RETURN(rels.styles_path, ResolveRelativePath(base_dir, target));
        } else {
          ASSIGN_OR_RETURN(UnknownRelationship unknown, ooxml::make_unknown_relationship(rel, base_dir));
          if (!unknown.type.empty()) {
            rels.unknown_rels.push_back(std::move(unknown));
          }
        }
        return Expected<void, Error>::Ok();
      }));
  return rels;
}

/// Reads every entry of a worksheet's `.rels` part, keeping the original
/// relationship ids so the retained tail records still resolve. Internal
/// targets are normalised to package-relative paths (matching what the OOXML
/// reader stores); external targets stay verbatim.
Expected<std::vector<UnknownRelationship>, Error> LoadSheetRelationships(const ZipReader& zip,
                                                                         std::string_view rels_path,
                                                                         std::string_view sheet_dir) {
  std::vector<UnknownRelationship> out;
  RETURN_IF_ERROR(ooxml::visit_relationship_nodes(
      zip, rels_path, "sheet rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (std::string_view(rel.attribute("Id").value()).empty()) {
          return Expected<void, Error>::Ok();
        }
        ASSIGN_OR_RETURN(UnknownRelationship entry, ooxml::make_unknown_relationship(rel, sheet_dir));
        out.push_back(std::move(entry));
        return Expected<void, Error>::Ok();
      }));
  return out;
}

/// Returns the package-relative target of the last internal relationship
/// declared by `part_path` whose `attr` attribute equals `value`, or an
/// empty string when the part has no rels file or no such relationship.
/// `label` names the rels file in error messages.
Expected<std::string, Error> FindRelationship(const ZipReader& zip, std::string_view part_path, const char* attr,
                                              std::string_view value, std::string_view label) {
  const std::string rels_path = ooxml::rels_path_for_part(part_path);
  if (value.empty() || !zip.has_entry(rels_path)) {
    return std::string();
  }
  const std::string dir = ooxml::dir_of(part_path);
  std::string found;
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, label, "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (std::string_view(rel.attribute(attr).value()) != value) {
          return Expected<void, Error>::Ok();
        }
        if (std::string_view(rel.attribute("TargetMode").value()) == "External") {
          return Expected<void, Error>::Ok();
        }
        auto resolved = ResolveRelativePath(dir, rel.attribute("Target").value());
        if (!resolved) {
          return Expected<void, Error>(std::move(resolved.error()));
        }
        found = std::move(resolved).value();
        return Expected<void, Error>::Ok();
      });
  if (!status) {
    return std::move(status.error());
  }
  return found;
}

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
