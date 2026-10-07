//
// Implementation of the XLSB read pipeline. See `io/xlsb/reader.h` for
// the design context.
//
// The reader reuses the OOXML zip envelope plus three XML parts
// (`[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`).
// The binary parts (`xl/workbook.bin`, `xl/worksheets/sheet*.bin`,
// `xl/sharedStrings.bin`) are decoded via the record framing in
// `io/xlsb/record.h`; worksheets decode in `io/xlsb/sheet_reader.cpp`,
// the shared-string table in `io/xlsb/sst_reader.cpp`. Formulas are
// decoded through a full `Ptg → AST → Excel-formula-text` pipeline
// (`DecodeFormulaText` in `io/xlsb/sheet_reader.cpp`); a Ptg stream
// outside the supported token set logs a structured warning and leaves
// the formula text empty instead.

#include "io/xlsb/reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "io/cf_reader.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sst_reader.h"
#include "io/xlsb/styles_reader.h"
#include "io/xlsb/workbook_bin_reader.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// ---------------------------------------------------------------------------
// XML envelope: same shape as the OOXML reader, but the relationship and
// content types target binary parts (`*.bin`) instead of XML.
// ---------------------------------------------------------------------------

// `[Content_Types].xml` registers the workbook part under the binary
// content type for `.xlsb`. The reader gates on this content type to
// avoid acting on a `.xlsx` archive that happened to be passed in.
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
// Some non-macro xlsb writers emit the alternative content type below;
// we accept either flavor as "this is an xlsb workbook".
constexpr std::string_view kCtWorkbookXlsbAlt = "application/vnd.ms-excel.sheet.macroEnabled.main";

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

// ---------------------------------------------------------------------------
// XML envelope walkers (return `Expected` on every failure path).
// ---------------------------------------------------------------------------

/// Walks `[Content_Types].xml` and asserts at least one Override is
/// for the xlsb workbook content type. Also returns the full Override
/// list so callers can compute passthrough parts.
struct ContentTypesView {
  std::vector<std::pair<std::string, std::string>> overrides;  // (part_name, content_type)
  std::vector<DefaultContentType> defaults;
};

Expected<ContentTypesView, Error> LoadContentTypes(const std::vector<std::uint8_t>& ct_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(ct_bytes.data(), ct_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=xlsb_reader part=[Content_Types].xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "[Content_Types].xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=xlsb_reader part=[Content_Types].xml");
  }
  ContentTypesView view;
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
      if (strings::case_insensitive_eq(ext, "bin") && (ct == kCtWorkbookXlsb || ct == kCtWorkbookXlsbAlt)) {
        saw_workbook = true;
      }
      if (ext.empty()) {
        continue;
      }
      const std::string normalized_ext = ooxml::lowercase_extension(ext);
      auto existing = std::find_if(
          view.defaults.begin(), view.defaults.end(),
          [&normalized_ext](const DefaultContentType& value) { return value.extension == normalized_ext; });
      if (existing != view.defaults.end()) {
        if (existing->content_type != ct) {
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
    if (ct == kCtWorkbookXlsb || ct == kCtWorkbookXlsbAlt) {
      saw_workbook = true;
    }
    auto existing = std::find_if(
        view.overrides.begin(), view.overrides.end(),
        [&part_name](const std::pair<std::string, std::string>& value) { return value.first == part_name; });
    if (existing != view.overrides.end()) {
      if (existing->second != ct) {
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
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=xlsb_reader part=_rels/.rels");
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
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

Expected<std::vector<UnknownRelationship>, Error> ReadUnknownPackageRels(const std::vector<std::uint8_t>& bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    return make_error(FormulonErrorCode::kIoXmlParse, "package-level rels: pugixml parse failed",
                      "context=xlsb_reader part=_rels/.rels desc=" + std::string(parse.description()));
  }
  const pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=xlsb_reader part=_rels/.rels");
  }
  std::vector<UnknownRelationship> result;
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type(rel.attribute("Type").value());
    if (type == kRelOfficeDocument || type == kRelCoreProperties || type == kRelExtendedProperties ||
        type == kRelCustomProperties) {
      continue;
    }
    const bool external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    std::string target = NormalisePartName(rel.attribute("Target").value());
    if (type.empty() || target.empty())
      continue;
    if (!external && !ooxml::is_safe_part_name(target)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "package relationship target escapes package root",
                        "context=xlsb_reader part=_rels/.rels target=" + target);
    }
    result.push_back(
        UnknownRelationship{std::string(rel.attribute("Id").value()), std::string(type), std::move(target), external});
  }
  return result;
}

struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;  // rId -> resolved part path
  std::string sst_path;
  std::string styles_path;
  std::vector<UnknownRelationship> unknown_rels;
};

Expected<WorkbookRels, Error> LoadWorkbookRels(const ZipReader& zip, std::string_view workbook_path) {
  WorkbookRels rels;
  const std::string rels_path = ooxml::rels_path_for_part(workbook_path);
  if (!zip.has_entry(rels_path)) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=xlsb_reader rels_path=" + rels_path);
  }
  const std::string base_dir = ooxml::dir_of(workbook_path);
  auto visit_status = ooxml::visit_relationship_nodes(
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
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return std::move(resolved.error());
          }
          rels.sheet_targets.emplace(id, std::move(resolved).value());
        } else if (type == kRelSharedStrings) {
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return std::move(resolved.error());
          }
          rels.sst_path = std::move(resolved).value();
        } else if (type == kRelStyles) {
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return std::move(resolved.error());
          }
          rels.styles_path = std::move(resolved).value();
        } else {
          UnknownRelationship unknown;
          unknown.id.assign(rel.attribute("Id").value());
          unknown.type.assign(type);
          unknown.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
          if (unknown.target_external) {
            unknown.target.assign(target);
          } else {
            auto resolved = ResolveRelativePath(base_dir, target);
            if (!resolved) {
              return std::move(resolved.error());
            }
            unknown.target = std::move(resolved).value();
          }
          if (!unknown.type.empty()) {
            rels.unknown_rels.push_back(std::move(unknown));
          }
        }
        return Expected<void, Error>::Ok();
      });
  if (!visit_status) {
    return std::move(visit_status.error());
  }
  return rels;
}

// ---------------------------------------------------------------------------
// Binary part decoders.
// ---------------------------------------------------------------------------

/// Reads every entry of a worksheet's `.rels` part, keeping the original
/// relationship ids so the retained tail records still resolve. Internal
/// targets are normalised to package-relative paths (matching what the OOXML
/// reader stores); external targets stay verbatim.
Expected<std::vector<UnknownRelationship>, Error> LoadSheetRelationships(const ZipReader& zip,
                                                                         std::string_view rels_path,
                                                                         std::string_view sheet_dir) {
  std::vector<UnknownRelationship> out;
  auto status =
      ooxml::visit_relationship_nodes(zip, rels_path, "sheet rels", "xlsb_reader", [&](const pugi::xml_node& rel) {
        UnknownRelationship entry;
        entry.id = rel.attribute("Id").value();
        if (entry.id.empty()) {
          return Expected<void, Error>::Ok();
        }
        entry.type = rel.attribute("Type").value();
        const std::string_view target = rel.attribute("Target").value();
        entry.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
        if (entry.target_external) {
          entry.target = std::string(target);
        } else {
          auto resolved = ResolveRelativePath(sheet_dir, target);
          if (!resolved) {
            return Expected<void, Error>(std::move(resolved.error()));
          }
          entry.target = std::move(resolved).value();
        }
        out.push_back(std::move(entry));
        return Expected<void, Error>::Ok();
      });
  if (!status) {
    return std::move(status.error());
  }
  return out;
}

/// Returns the package-relative target of the first internal relationship
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
        if (!found.empty() || std::string_view(rel.attribute(attr).value()) != value) {
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

/// Links a retained passthrough part's package path to the specific model
/// object `PassthroughPart::model_fingerprint` must track for it. Built
/// while pivot parts and styles load below; applied once, by the
/// passthrough capture sweep, to tag each `PassthroughPart` as it is
/// created (see `io/xlsb/retained_part_fingerprint.h`).
struct RetainedPartOriginTag {
  PassthroughPart::RetainedOrigin origin = PassthroughPart::RetainedOrigin::kNone;
  std::size_t sheet_index = 0;
  std::size_t pivot_index = 0;
  std::uint32_t cache_id = 0;
};

/// Package path -> `RetainedPartOriginTag`. The path index rides on the
/// `string -> uint32_t` hash map the pivot-cache loader already uses, and
/// the tags live beside it; a later `set` for the same path replaces the tag.
struct RetainedPartOrigins {
  std::unordered_map<std::string, std::uint32_t> index;
  std::vector<RetainedPartOriginTag> tags;

  void set(const std::string& path, const RetainedPartOriginTag& tag) {
    const auto [it, inserted] = index.emplace(path, static_cast<std::uint32_t>(tags.size()));
    if (inserted) {
      tags.push_back(tag);
    } else {
      tags[it->second] = tag;
    }
  }

  const RetainedPartOriginTag* find(const std::string& path) const {
    const auto it = index.find(path);
    return it == index.end() ? nullptr : &tags[it->second];
  }
};

/// Decodes the cache a pivot-table part binds to, returning the model id
/// it was registered under. Caches already loaded are reused so two
/// tables over one cache share it, as they do in the file.
///
/// The id is ours to assign: the workbook's own cache-id table is not
/// decoded, and nothing outside the model consults these ids -- the parts
/// round-trip through passthrough rather than being rewritten from the
/// model.
std::optional<std::uint32_t> LoadPivotCacheFor(const ZipReader& zip, Workbook& wb, std::string_view pivot_table_path,
                                               std::unordered_map<std::string, std::uint32_t>& loaded,
                                               RetainedPartOrigins& origins) {
  auto def_path_or = FindRelationship(zip, pivot_table_path, "Type", kRelPivotCacheDefinition, "pivot rels");
  if (!def_path_or || def_path_or.value().empty()) {
    return std::nullopt;
  }
  const std::string def_path = std::move(def_path_or).value();
  const auto seen = loaded.find(def_path);
  if (seen != loaded.end()) {
    return seen->second;
  }
  auto rec_path_or = FindRelationship(zip, def_path, "Type", kRelPivotCacheRecords, "pivot rels");
  if (!rec_path_or || rec_path_or.value().empty()) {
    return std::nullopt;
  }
  const std::string rec_path = std::move(rec_path_or).value();
  if (!zip.has_entry(def_path) || !zip.has_entry(rec_path)) {
    return std::nullopt;
  }
  auto def_bytes_or = zip.read_entry(def_path);
  if (!def_bytes_or) {
    return std::nullopt;
  }
  auto rec_bytes_or = zip.read_entry(rec_path);
  if (!rec_bytes_or) {
    return std::nullopt;
  }
  const std::vector<std::uint8_t>& def_bytes = def_bytes_or.value();
  const std::vector<std::uint8_t>& rec_bytes = rec_bytes_or.value();
  auto cache_or =
      read_pivot_cache_bin(ByteSpan{def_bytes.data(), def_bytes.size()}, ByteSpan{rec_bytes.data(), rec_bytes.size()});
  if (!cache_or) {
    return std::nullopt;
  }
  const auto cache_id = static_cast<std::uint32_t>(wb.pivot_caches().size());
  pivot::PivotCache cache = std::move(cache_or.value());
  cache.set_cache_id(cache_id);
  origins.set(def_path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheDefinition, 0, 0, cache_id});
  origins.set(rec_path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheRecords, 0, 0, cache_id});
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(std::move(cache)));
  loaded.emplace(def_path, cache_id);
  return cache_id;
}

/// Reads the remote workbook URLs recorded in an external link part's own
/// rels file into `record`: the relationship `record.body_rel_id` names is
/// the target, and another external-path relationship beside it is the
/// absolute URL Excel records for a relative target. Without a body rel id
/// the first external-path relationship is the target and no absolute URL
/// is read. Leaves both empty when the part has no rels, which is what an
/// unresolvable link already looks like.
void ReadExternalLinkTargets(const ZipReader& zip, ExternalLinkRecord& record) {
  const std::string rels_path = ooxml::rels_path_for_part(record.part_path);
  if (!zip.has_entry(rels_path)) {
    return;
  }
  const bool body_known = !record.body_rel_id.empty();
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, "external link rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (std::string_view(rel.attribute("Type").value()) != kRelExternalLinkPath) {
          return Expected<void, Error>::Ok();
        }
        const std::string_view id = rel.attribute("Id").value();
        const bool is_body = body_known ? id == record.body_rel_id : record.target.empty();
        if (is_body) {
          record.target = rel.attribute("Target").value();
        } else if (body_known && record.absolute_target.empty()) {
          record.absolute_target = rel.attribute("Target").value();
          record.absolute_rel_id = id;
        }
        return Expected<void, Error>::Ok();
      });
  (void)status;
}

/// `BrtBeginExternalBook` `sbt` values other than a workbook (0).
constexpr std::uint16_t kSbtDde = 1;
constexpr std::uint16_t kSbtOle = 2;

/// The `sbt` of the part's opening `BrtBeginExternalBook`; 0 when the part
/// does not open with one.
std::uint16_t ExternalLinkSbt(ByteSpan part) {
  auto rec_or = read_record(part);
  if (!rec_or || rec_or.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtBeginExternalBook)) {
    return 0;
  }
  ByteSpan payload = rec_or.value().payload;
  auto sbt_or = read_u16(payload);
  return sbt_or ? sbt_or.value() : std::uint16_t{0};
}

/// Loads every external link part the supporting-book list names, in
/// `[N]` order, so `external_links()[N - 1]` is the book a formula's
/// `[N]` prefix selects.
///
/// A part that is missing, unreadable or uses an encoding the decoder
/// has not measured contributes an entry with an empty cache rather than
/// none: the position has to be held or every later `[N]` would shift by
/// one and resolve against the wrong workbook. A reference into an empty
/// cache reads `#REF!`, which is the behaviour before this existed.
///
/// A link whose `BrtBeginExternalBook` carries `sbt` 1 (DDE) or 2 (OLE) is
/// recorded as that kind and its body is not read as a supporting book. The
/// XLSB writer regenerates the parts of external-book links only, so a DDE
/// or OLE part keeps round-tripping verbatim through passthrough.
void LoadExternalLinkParts(const ZipReader& zip, Workbook& wb, const std::vector<XlsbSupBook>& books) {
  std::vector<ExternalLinkRecord> links;
  for (const XlsbSupBook& sup : books) {
    if (sup.external_book == 0) {
      continue;  // This workbook.
    }
    ExternalLinkRecord record;
    record.index = sup.external_book;
    record.rel_id = sup.rel_id;
    auto path_or = FindRelationship(zip, "xl/workbook.bin", "Id", sup.rel_id, "external link rels");
    if (path_or && !path_or.value().empty() && zip.has_entry(path_or.value())) {
      record.part_path = path_or.value();
      auto bytes_or = zip.read_entry(record.part_path);
      if (bytes_or) {
        const std::vector<std::uint8_t>& bytes = bytes_or.value();
        const std::uint16_t sbt = ExternalLinkSbt(ByteSpan{bytes.data(), bytes.size()});
        if (sbt == kSbtDde || sbt == kSbtOle) {
          record.kind = sbt == kSbtDde ? ExternalLinkRecord::Kind::kDdeLink : ExternalLinkRecord::Kind::kOleLink;
        } else if (auto book_or = read_external_link_bin(ByteSpan{bytes.data(), bytes.size()}, &record.body_rel_id)) {
          record.kind = ExternalLinkRecord::Kind::kExternalBook;
          record.book = std::move(book_or.value());
        } else {
          record.body_rel_id.clear();
        }
      }
      ReadExternalLinkTargets(zip, record);
    }
    links.push_back(std::move(record));
  }
  if (!links.empty()) {
    wb.set_external_links(std::move(links));
  }
}

/// Projects the loaded external links onto the two name tables
/// `decode_ptgs` consults. The projection exists so the Ptg decoder does
/// not have to know about `Workbook` or the link records at all.
XlsbExternalBooks CollectExternalBookTables(const Workbook& wb) {
  XlsbExternalBooks books;
  books.reserve(wb.external_links().size());
  for (const ExternalLinkRecord& link : wb.external_links()) {
    XlsbExternalBook entry;
    entry.sheet_names = link.book.sheet_names;
    entry.names.reserve(link.book.names.size());
    for (const ExternalBookName& name : link.book.names) {
      entry.names.push_back(name.name);
    }
    books.push_back(std::move(entry));
  }
  return books;
}

/// Decodes every pivot table reachable from a sheet's relationships.
///
/// A pivot this reader cannot account for is skipped, not failed. Before
/// this existed no XLSB pivot was decoded at all, so refusing one leaves
/// that workbook loading exactly as it used to, whereas propagating the
/// error would turn a working open into a failure. Skipping is also why
/// the parts are left out of `consumed_parts`: the XLSB writer has no
/// pivot output, so a decoded-but-not-consumed part still round-trips
/// verbatim, and a skipped one is indistinguishable from before.
void LoadPivotParts(const ZipReader& zip, Workbook& wb, RetainedPartOrigins& origins) {
  std::unordered_map<std::string, std::uint32_t> loaded_caches;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    std::vector<std::string> pivot_table_paths;
    for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
      if (rel.type == kRelPivotTable && !rel.target_external && !rel.target.empty()) {
        pivot_table_paths.push_back(rel.target);
      }
    }
    for (const std::string& path : pivot_table_paths) {
      if (!zip.has_entry(path)) {
        continue;
      }
      auto bytes_or = zip.read_entry(path);
      if (!bytes_or) {
        continue;
      }
      const std::vector<std::uint8_t>& bytes = bytes_or.value();
      auto table_or = read_pivot_table_bin(ByteSpan{bytes.data(), bytes.size()});
      if (!table_or) {
        continue;
      }
      const std::optional<std::uint32_t> cache_id = LoadPivotCacheFor(zip, wb, path, loaded_caches, origins);
      if (!cache_id) {
        continue;
      }
      pivot::PivotTable table = std::move(table_or.value());
      table.set_pivot_cache_id(*cache_id);
      origins.set(path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotTable, i,
                                              wb.sheet(i).pivot_tables().size(), 0});
      wb.sheet(i).add_pivot_table(std::make_unique<pivot::PivotTable>(std::move(table)));
    }
  }
  // Backfills each field's source-column name and each item's label from
  // the bound cache; without it GETPIVOTDATA cannot match a field by name.
  pivot::resolve_all_pivot_names(wb);
}

/// Resolves the relationship ids carried by model-owned BrtHLink records.
/// Only relationships actually consumed by a hyperlink are removed from the
/// sheet's unknown relationship list; unrelated drawing/table/custom rels,
/// including unused hyperlink rels, remain available to the writer.
Expected<void, Error> ResolveSheetHyperlinks(Sheet& sheet, std::vector<UnknownRelationship>& relationships) {
  std::unordered_set<std::string> consumed;
  for (Hyperlink& hyperlink : sheet.mutable_hyperlinks()) {
    if (hyperlink.rid.empty()) {
      // A present-but-empty RelID is the internal target form. Its location
      // string carries the destination and no relationship is consumed.
      continue;
    }
    const UnknownRelationship* matched = nullptr;
    for (const UnknownRelationship& relationship : relationships) {
      if (relationship.id != hyperlink.rid) {
        continue;
      }
      if (matched != nullptr) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id is duplicated",
                          "context=xlsb_reader rid=" + hyperlink.rid);
      }
      matched = &relationship;
    }
    if (matched == nullptr || matched->type != kRelHyperlink || !matched->target_external || matched->target.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                        matched == nullptr ? "xlsb BrtHLink relationship id is dangling"
                                           : "xlsb BrtHLink relationship has wrong type or target mode",
                        "context=xlsb_reader rid=" + hyperlink.rid);
    }
    hyperlink.target = matched->target;
    consumed.insert(hyperlink.rid);
  }
  if (!consumed.empty()) {
    relationships.erase(std::remove_if(relationships.begin(), relationships.end(),
                                       [&consumed](const UnknownRelationship& relationship) {
                                         return consumed.count(relationship.id) != 0U;
                                       }),
                        relationships.end());
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<XlsbReadResult, Error> read_xlsb(ByteSpan bytes) {
  ZipReader zip;
  if (auto open_result = zip.open(bytes); !open_result) {
    return std::move(open_result.error());
  }

  // 1. [Content_Types].xml — gate + Override list.
  if (!zip.has_entry("[Content_Types].xml")) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing from package",
                      "context=xlsb_reader");
  }
  ASSIGN_OR_RETURN(auto ct_bytes, zip.read_entry("[Content_Types].xml"));
  auto ct_view_or = LoadContentTypes(ct_bytes);
  if (!ct_view_or) {
    return std::move(ct_view_or.error());
  }
  ContentTypesView ct_view = ct_view_or.take();

  // 2. _rels/.rels — locate the workbook part path.
  if (!zip.has_entry("_rels/.rels")) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "_rels/.rels: missing from package",
                      "context=xlsb_reader");
  }
  ASSIGN_OR_RETURN(auto root_rels, zip.read_entry("_rels/.rels"));
  ASSIGN_OR_RETURN(auto wb_path, ResolveOfficeDocumentPath(root_rels));
  const std::string workbook_path = wb_path;
  ASSIGN_OR_RETURN(auto package_rels, ReadUnknownPackageRels(root_rels));

  // 3. xl/_rels/workbook.xml.rels (still XML in xlsb).
  auto wb_rels_or = LoadWorkbookRels(zip, workbook_path);
  if (!wb_rels_or) {
    return std::move(wb_rels_or.error());
  }
  const WorkbookRels& wb_rels = wb_rels_or.value();

  // 4. xl/workbook.bin — sheet bundle list.
  if (!zip.has_entry(workbook_path)) {
    return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: not found at relationship target",
                      "context=xlsb_reader workbook_path=" + workbook_path);
  }
  ASSIGN_OR_RETURN(auto wb_bytes, zip.read_entry(workbook_path));
  auto bundle_or = DecodeWorkbookBin(wb_bytes);
  if (!bundle_or) {
    return std::move(bundle_or.error());
  }
  const WorkbookBinInfo& workbook_info = bundle_or.value();
  const std::vector<SheetBundleEntry>& bundle = workbook_info.sheets;

  // `BrtName` (defined names + hidden future-function / LET-parameter
  // placeholders) and `BrtExternSheet` (qualified-reference sheet
  // ranges) both live in `xl/workbook.bin` globals; decode them once so
  // every sheet's `PtgName` / `PtgRef3d` / `PtgArea3d` tokens can
  // resolve against the same tables.
  auto name_table_or = DecodeWorkbookNames(wb_bytes);
  if (!name_table_or) {
    return std::move(name_table_or.error());
  }
  const std::vector<XlsbName>& name_table = name_table_or.value();
  const std::vector<XlsbSupBook> sup_books = DecodeSupBooks(wb_bytes);
  auto sheet_ranges_or = DecodeExternSheet(wb_bytes, sup_books);
  if (!sheet_ranges_or) {
    return std::move(sheet_ranges_or.error());
  }
  const std::vector<XlsbSheetRange>& sheet_ranges = sheet_ranges_or.value();

  // 5. Build the workbook bottom-up.
  Workbook wb = Workbook::create_empty();
  // The package's style table is authoritative, including when the package
  // has none: back-filling the factory's seeded defaults would invent style
  // records the file never carried and add an unwanted styles part on write.
  wb.set_styles(StylesTable{});
  wb.set_date1904(workbook_info.date1904);
  wb.set_workbook_protection_xml(workbook_info.protection_xml);
  std::vector<std::string> sheet_part_paths;
  sheet_part_paths.reserve(bundle.size());
  for (const SheetBundleEntry& b : bundle) {
    if (b.name.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: BrtBundleSh with empty name",
                        "context=xlsb_reader");
    }
    if (b.rid.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: BrtBundleSh with empty rId",
                        "context=xlsb_reader sheet=" + b.name);
    }
    auto it = wb_rels.sheet_targets.find(b.rid);
    if (it == wb_rels.sheet_targets.end()) {
      std::string ctx("context=xlsb_reader rid=");
      ctx.append(b.rid);
      ctx.append(" sheet=").append(b.name);
      return make_error(FormulonErrorCode::kIoRelationshipBroken,
                        "workbook.bin: BrtBundleSh rId has no matching workbook relationship", std::move(ctx));
    }
    // Same boundary validation the OOXML reader applies, and the same
    // error code: a duplicate sheet name makes every lookup resolve to the
    // first match, so the workbook would compute from the wrong sheet with
    // no ambiguity signal. Callers should not have to switch on the source
    // format to recognise that condition.
    auto added = wb.add_sheet_validated(b.name);
    if (!added) {
      if (added.error().code != FormulonErrorCode::kInvalidSheetName) {
        return std::move(added.error());
      }
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "workbook.bin: BrtBundleSh name is invalid or collides with an earlier sheet",
                        "context=xlsb_reader sheet=\"" + b.name + "\"");
    }
    if (b.visibility != SheetVisibility::kVisible) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().set_visibility(b.visibility);
    }
    sheet_part_paths.push_back(it->second);
  }

  // Ordered sheet display names — the Ptg decoder maps a 3-D reference's
  // `ixti` (0-based sheet index) to the qualifying sheet name through
  // this list.
  std::vector<std::string> sheet_names;
  sheet_names.reserve(bundle.size());
  for (const SheetBundleEntry& b : bundle) {
    sheet_names.push_back(b.name);
  }

  // Runs before both the defined names and the sheets, because a formula
  // in either can carry a `PtgNameX` that only the supporting workbook's
  // own name table can bind.
  LoadExternalLinkParts(zip, wb, sup_books);
  const XlsbExternalBooks external_books = CollectExternalBookTables(wb);

  // 5b. Register every non-hidden `BrtName` entry as a defined name.
  // Needs `sheet_names` (for qualified references inside a name's own
  // formula) and the complete `name_table` / `sheet_ranges` (for
  // self-consistent `PtgName` / `PtgRef3d` resolution), so this can
  // only run after both are fully built above.
  std::uint32_t undecoded_formula_count = 0;
  std::uint32_t undecoded_defined_name_count = 0;
  if (auto r = RegisterDefinedNames(wb_bytes, name_table, sheet_names, sheet_ranges, external_books, wb,
                                    &undecoded_defined_name_count);
      !r) {
    return std::move(r.error());
  }

  // 6. xl/sharedStrings.bin — load before the per-sheet decode loop so
  // BrtCellIsst can resolve indices in-pipeline. The text deque is
  // owned by the workbook itself so `Value::text` views remain valid
  // after the caller moves the workbook out of the read result.
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  std::vector<std::string_view> sst_entries;
  // Parallel to `sst_entries`, one (possibly empty) run list per index.
  std::vector<std::vector<PhoneticRun>> sst_phonetic;
  std::vector<PhoneticProperties> sst_phonetic_props;
  if (!wb_rels.sst_path.empty() && zip.has_entry(wb_rels.sst_path)) {
    ASSIGN_OR_RETURN(auto sst_bytes, zip.read_entry(wb_rels.sst_path));
    ASSIGN_OR_RETURN(auto sst, DecodeSharedStringsBin(sst_bytes, text_storage, sst_phonetic, sst_phonetic_props));
    sst_entries = std::move(sst);
  }

  // Links each retained pivot / styles part's package path to the model
  // object its `PassthroughPart::model_fingerprint` must track, so step 8
  // below can tag the `PassthroughPart` it captures for that path.
  RetainedPartOrigins retained_part_origins;

  // 6b. xl/styles.bin — numFmt + cellXfs/cellStyleXfs (see
  // `io/xlsb/styles_reader.h`). Deliberately NOT added to
  // `consumed_parts` below: leaving it out lets step 8's passthrough
  // loop capture the raw bytes (with the correct `[Content_Types].xml`
  // content-type) alongside the parsed `StylesTable`, so a
  // read-modify-write cycle keeps the original font/fill/border detail
  // this reader does not model in-memory.
  if (!wb_rels.styles_path.empty() && zip.has_entry(wb_rels.styles_path)) {
    auto styles_bytes_or = zip.read_entry(wb_rels.styles_path);
    if (!styles_bytes_or) {
      return std::move(styles_bytes_or.error());
    }
    const std::vector<std::uint8_t>& styles_bytes = styles_bytes_or.value();
    ByteSpan styles_span{styles_bytes.data(), styles_bytes.size()};
    ASSIGN_OR_RETURN(auto styles, read_styles_bin(styles_span));
    retained_part_origins.set(wb_rels.styles_path,
                              RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kStyles, 0, 0, 0});
    wb.set_styles(std::move(styles));
  }

  // 7. Each sheet binary.
  std::uint32_t cells_read = 0;
  std::uint32_t dropped_record_count = 0;
  std::unordered_set<std::string> consumed_parts;
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  consumed_parts.insert(ooxml::rels_path_for_part(workbook_path));
  if (!wb_rels.sst_path.empty()) {
    consumed_parts.insert(wb_rels.sst_path);
  }

  // A formula is a dynamic-array formula exactly when a `BrtCellMeta`
  // naming this entry of `xl/metadata.bin` precedes it.
  std::uint32_t dynamic_array_ifmd = 0U;
  if (zip.has_entry("xl/metadata.bin")) {
    auto metadata_or = zip.read_entry("xl/metadata.bin");
    if (metadata_or) {
      dynamic_array_ifmd =
          find_dynamic_array_cell_meta_index(ByteSpan{metadata_or.value().data(), metadata_or.value().size()});
    }
  }

  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string& sheet_path = sheet_part_paths[i];
    if (!zip.has_entry(sheet_path)) {
      std::string ctx("context=xlsb_reader sheet_path=");
      ctx.append(sheet_path);
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "sheet binary part missing from package", std::move(ctx));
    }
    ASSIGN_OR_RETURN(auto sheet_bytes, zip.read_entry(sheet_path));
    auto state_or = DecodeSheetBin(sheet_bytes, i, wb, sst_entries, sst_phonetic, sst_phonetic_props, text_storage,
                                   sheet_names, name_table, sheet_ranges, external_books, &undecoded_formula_count);
    if (!state_or) {
      return std::move(state_or.error());
    }
    cells_read += state_or.value().cells_decoded;
    dropped_record_count += state_or.value().dropped_records;
    std::unordered_set<std::uint64_t> dynamic_cells;
    for (const auto& [row, col, ifmd] : state_or.value().cell_metadata) {
      if (dynamic_array_ifmd != 0U && ifmd == dynamic_array_ifmd) {
        dynamic_cells.insert(dynamic_array_cell_key(row, col));
      }
    }
    apply_loaded_dynamic_array_marks(wb.sheet(i), dynamic_cells);
    consumed_parts.insert(sheet_path);

    // The sheet's own rels file resolves the relationship ids carried by the
    // retained tail records (hyperlink targets, drawing and table parts). It
    // is `rels`-Default-typed, so the Override-driven passthrough loop below
    // never sees it; keeping every entry — none of them are modelled by the
    // binary reader — preserves both the ids and their targets.
    const std::string sheet_rels_path = ooxml::rels_path_for_part(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      auto rels_or = LoadSheetRelationships(zip, sheet_rels_path, ooxml::dir_of(sheet_path));
      if (!rels_or) {
        return std::move(rels_or.error());
      }
      auto relationships = std::move(rels_or.value());
      auto hyperlinks = ResolveSheetHyperlinks(wb.sheet(i), relationships);
      if (!hyperlinks) {
        return std::move(hyperlinks.error());
      }
      wb.sheet(i).set_unknown_relationships(std::move(relationships));
      consumed_parts.insert(sheet_rels_path);
    } else {
      std::vector<UnknownRelationship> no_relationships;
      auto hyperlinks = ResolveSheetHyperlinks(wb.sheet(i), no_relationships);
      if (!hyperlinks) {
        return std::move(hyperlinks.error());
      }
    }
  }

  // Conditional-format and data-validation formulas name other books
  // through the same `[N]` cell formulas do.
  ingest_feature_formulas(wb);

  // Pivot caches and tables, reached through the sheet relationships the
  // loop above stored. Runs after every sheet is decoded because a table
  // is attached to its sheet and its cache is workbook-level.
  LoadPivotParts(zip, wb, retained_part_origins);

  // 8. Passthrough parts. First capture every Override-listed part the
  // reader did not consume, retaining its explicit content type. A second
  // residual sweep resolves the remaining entries through the source
  // Default registry. Default-typed parts intentionally carry an empty
  // content_type; the registry itself travels on the Workbook and the
  // writer re-emits the matching Default entry.
  //
  // A part named in `retained_part_origins` (pivot table/cache, styles) is
  // tagged with the model object it depends on and fingerprinted against
  // the now-fully-loaded workbook, so `write_xlsb` can tell an unmodified
  // retained part from one whose model twin has since been mutated.
  auto apply_retained_origin = [&retained_part_origins, &wb](PassthroughPart& part) {
    const RetainedPartOriginTag* found = retained_part_origins.find(part.path);
    if (found == nullptr) {
      return;
    }
    const RetainedPartOriginTag& tag = *found;
    part.retained_origin = tag.origin;
    part.origin_sheet_index = tag.sheet_index;
    part.origin_pivot_index = tag.pivot_index;
    part.origin_cache_id = tag.cache_id;
    part.model_fingerprint = current_retained_part_fingerprint(wb, part);
  };

  std::vector<PassthroughPart> unknown_parts;
  unknown_parts.reserve(ct_view.overrides.size());
  std::unordered_set<std::string> captured_parts;
  captured_parts.reserve(ct_view.overrides.size());
  for (const auto& [part_name, content_type] : ct_view.overrides) {
    if (consumed_parts.find(part_name) != consumed_parts.end()) {
      continue;
    }
    if (captured_parts.find(part_name) != captured_parts.end()) {
      continue;
    }
    // Refuse a traversal-shaped passthrough name so a round-tripped .xlsb
    // never hands a downstream extractor a zip-slip primitive.
    if (!ooxml::is_safe_part_name(part_name)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "Override part name escapes package root; refusing to load",
                        "context=xlsb_reader part=" + part_name);
    }
    if (!zip.has_entry(part_name)) {
      continue;
    }
    ASSIGN_OR_RETURN(auto part_bytes, zip.read_entry(part_name));
    PassthroughPart part;
    part.path = part_name;
    part.content_type = content_type;
    part.bytes = std::move(part_bytes);
    apply_retained_origin(part);
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(part_name);
  }

  auto default_content_type_for = [&ct_view](std::string_view path) -> const DefaultContentType* {
    const std::string extension = ooxml::extension_of_part(path);
    if (extension.empty()) {
      return nullptr;
    }
    const auto it =
        std::find_if(ct_view.defaults.begin(), ct_view.defaults.end(),
                     [&extension](const DefaultContentType& value) { return value.extension == extension; });
    return it == ct_view.defaults.end() ? nullptr : &*it;
  };

  std::uint32_t dropped_part_count = 0;
  std::string first_dropped;
  for (const std::string& entry : zip.list_entries()) {
    // Directory markers are not OPC parts. Skip them before path validation:
    // ZIP producers commonly record a trailing-slash directory entry, and
    // the trailing slash is intentionally not a canonical OPC part name.
    // Payload entries are validated below, including consumed paths, so a
    // hostile ZIP catalogue cannot hide a traversal-shaped name behind the
    // modelled path set.
    if (entry.empty() || entry.back() == '/') {
      continue;
    }
    if (!ooxml::is_safe_part_name(entry)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "archive entry name escapes package root; refusing to load",
                        "context=xlsb_reader part=" + entry);
    }
    if (consumed_parts.find(entry) != consumed_parts.end() || captured_parts.find(entry) != captured_parts.end()) {
      continue;
    }
    const DefaultContentType* default_type = default_content_type_for(entry);
    if (default_type == nullptr || default_type->content_type.empty()) {
      if (dropped_part_count == 0U) {
        first_dropped = entry;
      }
      ++dropped_part_count;
      continue;
    }
    ASSIGN_OR_RETURN(auto part_bytes, zip.read_entry(entry));
    PassthroughPart part;
    part.path = entry;
    // Default-typed parts deliberately do not copy the effective content
    // type into the per-part record. This keeps Override and Default
    // semantics distinguishable on the write side.
    part.bytes = std::move(part_bytes);
    apply_retained_origin(part);
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(entry);
  }

  sort_passthrough_parts(unknown_parts);

  if (dropped_part_count != 0U) {
    StructuredLog("xlsb.package.parts_dropped")
        .field("count", static_cast<std::int64_t>(dropped_part_count))
        .field("first_part", first_dropped)
        .field("reason", std::string_view("part content type could not be resolved from Override or Default"))
        .warn();
  }

  // The workbook is the sole owner; the read result does not mirror the
  // payload. See `XlsbReadResult`.
  wb.set_passthrough_parts(std::move(unknown_parts));
  wb.set_default_content_types(std::move(ct_view.defaults));
  wb.set_unknown_package_rels(std::move(package_rels));
  wb.set_unknown_workbook_rels(std::move(wb_rels_or.value().unknown_rels));

  wb.apply_legacy_implicit_intersections();

  XlsbReadResult result{std::move(wb),      cells_read,          undecoded_formula_count, undecoded_defined_name_count,
                        dropped_part_count, dropped_record_count};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
