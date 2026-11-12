//
// Implementation of the XLSB read pipeline. See `io/xlsb/reader.h` for
// the design context.
//
// The reader reuses the OOXML zip envelope plus three XML parts
// (`[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`).
// The binary parts (`xl/workbook.bin`, `xl/worksheets/sheet*.bin`,
// `xl/sharedStrings.bin`) are decoded via the record framing in
// `io/xlsb/record.h`; worksheets decode in `io/xlsb/sheet_reader.cpp` and
// `io/xlsb/sheet_cell_decode.cpp`, the shared-string table in `io/xlsb/sst_reader.cpp`. Formulas are
// decoded through a full `Ptg → AST → Excel-formula-text` pipeline
// (`DecodeFormulaText` in `io/xlsb/sheet_cell_decode.cpp`); a Ptg stream
// outside the supported token set logs a structured warning and leaves
// the formula text empty instead.

#include "io/xlsb/reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "hyperlink_address.h"
#include "io/cf_reader.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/package_rels_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/pivot_table_reader.h"
#include "io/rich_value_errors.h"
#include "io/xlsb/drawing_record.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/package_loader.h"
#include "io/xlsb/part_paths.h"
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/slicer_records.h"
#include "io/xlsb/sst_reader.h"
#include "io/xlsb/styles_reader.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/workbook_bin_reader.h"
#include "io/xml_utils.h"
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

using detail::ContentTypesView;
using detail::FindRelationship;
using detail::LoadContentTypes;
using detail::LoadSheetRelationships;
using detail::LoadWorkbookRels;
using detail::WorkbookRels;

// ---------------------------------------------------------------------------
// Binary part decoders.
// ---------------------------------------------------------------------------

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
    StructuredLog("xlsb.pivot.skipped").field("part", def_path).field("reason", cache_or.error().message).warn();
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
/// does not support contributes an entry with an empty cache rather than
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
    auto path_or = FindRelationship(zip, kWorkbookBinPath, "Id", sup.rel_id, "external link rels");
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

/// Consumes `xl/metadata.bin` and the `xl/richData/` parts into `*set` when
/// they hold nothing but standard dynamic-array metadata and measured rich
/// errors; their workbook relationships go with them, and the writer
/// regenerates the parts. Otherwise every part stays passthrough and false is
/// returned. No sheet of a binary workbook is kept opaque, so that condition
/// of the shared consumption rule always holds.
bool ConsumeRichValueParts(const ZipReader& zip, RichErrorSet* set, std::unordered_set<std::string>& consumed_parts,
                           std::vector<UnknownRelationship>& unknown_rels) {
  if (!zip.has_entry(kMetadataBinPath)) {
    return false;
  }
  constexpr std::string_view kExternalLinks = "xl/externalLinks/";
  std::vector<std::string> parts{std::string(kMetadataBinPath)};
  for (const std::string& name : zip.list_entries()) {
    // An external link's cached cells naming a value-metadata entry belong to a set this reader does not own.
    if (name.compare(0, kExternalLinks.size(), kExternalLinks) == 0 && name.size() > 4U &&
        name.compare(name.size() - 4U, 4U, ".bin") == 0) {
      auto link = zip.read_entry(name);
      if (!link) {
        return false;
      }
      for (ByteSpan cursor{link.value().data(), link.value().size()}; cursor.size > 0U;) {
        auto record = read_record(cursor);
        if (!record || record.value().type == static_cast<std::uint16_t>(XlsbRecordType::BrtValueMeta)) {
          return false;
        }
      }
    }
    if (name.compare(0, kRichDataDirPrefix.size(), kRichDataDirPrefix) == 0 && name.back() != '/') {
      if (name != kRdRichValuePartPath && name != kRdRichValueStructurePartPath && name != kRdRichValueTypesPartPath) {
        return false;
      }
      parts.push_back(name);
    }
  }
  auto metadata = zip.read_entry(kMetadataBinPath);
  if (!metadata) {
    return false;
  }
  RichValueMetadata info = parse_metadata_bin(ByteSpan{metadata.value().data(), metadata.value().size()});
  if (!info.consumable) {
    return false;
  }
  if (zip.has_entry(kRdRichValuePartPath) || !info.vm_rv.empty()) {
    auto values = zip.read_entry(kRdRichValuePartPath);
    auto structures = zip.read_entry(kRdRichValueStructurePartPath);
    if (!values || !structures) {
      return false;
    }
    for (const std::optional<RichError>& v : parse_rich_errors(values.value(), structures.value())) {
      if (!v.has_value()) {
        return false;
      }
      set->values.push_back(*v);
    }
  }
  for (const std::uint32_t rv : info.vm_rv) {
    if (rv >= set->values.size()) {
      set->values.clear();
      return false;
    }
  }
  set->vm_rv = std::move(info.vm_rv);
  for (const std::string& part : parts) {
    consumed_parts.insert(part);
  }
  unknown_rels.erase(std::remove_if(unknown_rels.begin(), unknown_rels.end(),
                                    [&parts](const UnknownRelationship& r) {
                                      return !r.target_external &&
                                             std::find(parts.begin(), parts.end(), r.target) != parts.end();
                                    }),
                     unknown_rels.end());
  return true;
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
void LoadPivotParts(const ZipReader& zip, Workbook& wb, RetainedPartOrigins& origins,
                    const std::vector<DeferredSpillAnchor>& deferred_spill_anchors) {
  std::unordered_map<std::string, std::uint32_t> loaded_caches;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const std::size_t first_new_pivot = wb.sheet(i).pivot_tables().size();
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
        // Not modelled: the part stays passthrough and the report is not a model pivot.
        StructuredLog("xlsb.pivot.skipped").field("part", path).field("reason", table_or.error().message).warn();
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
    assign_read_creation_orders(wb, i, first_new_pivot);
  }
  // Backfills each field's source-column name and each item's label from
  // the bound cache; without it GETPIVOTDATA cannot match a field by name.
  pivot::resolve_all_pivot_names(wb);
  // A pivot saved blocked is recognised by its anchor's rich value; one whose
  // anchor has none (an older file, or rich parts left passthrough) by its shape.
  restore_deferred_spill_anchors(wb, deferred_spill_anchors);
  wb.adopt_saved_blocked_pivots();
  // Cells and tables are both loaded, so the totals cells settle the columns' totals metadata.
  wb.normalize_table_totals_cells();
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
    hyperlink.target = loaded_hyperlink_target(matched->target);
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

/// Why `LiftSheetDrawing` left a sheet's drawing retained, or empty when it
/// lifted it (or the sheet has no drawing at all).
std::string_view DrawingLiftRefusal(const Sheet& sheet, const DrawingRecordScan& scan, const ContentTypesView& ct_view,
                                    const ZipReader& zip, const UnknownRelationship** lifted) {
  if (scan.count != 1U || !scan.parsed) {
    return "record_count";
  }
  if (!scan.rel_id) {
    return "payload";
  }
  std::size_t drawing_rels = 0;
  std::size_t same_id = 0;
  const UnknownRelationship* matched = nullptr;
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    if (rel.id == *scan.rel_id) {
      ++same_id;
    }
    if (rel.type == kRelDrawing) {
      ++drawing_rels;
      if (rel.id == *scan.rel_id) {
        matched = &rel;
      }
    }
  }
  if (same_id > 1U || drawing_rels > 1U) {
    return "rel_count";
  }
  if (matched == nullptr) {
    return "rel_missing";
  }
  if (matched->target_external) {
    return "rel_external";
  }
  if (!ooxml::is_safe_part_name(matched->target) || !zip.has_entry(matched->target)) {
    return "part_absent";
  }
  // The first Override naming the part is the one passthrough capture keeps.
  const auto it = std::find_if(ct_view.overrides.begin(), ct_view.overrides.end(),
                               [matched](const auto& entry) { return entry.first == matched->target; });
  if (it == ct_view.overrides.end() || it->second != kCtDrawing) {
    return "content_type";
  }
  *lifted = matched;
  return {};
}

/// Moves a sheet's drawing from the retained tail record and its unknown
/// relationship onto `Sheet::drawing_rel_target`, which makes it editable as
/// an OOXML drawing is. A drawing that does not meet every condition stays
/// retained as before and is logged; the drawing part and everything it
/// reaches stay passthrough either way.
void LiftSheetDrawing(Sheet& sheet, std::size_t sheet_index, const ContentTypesView& ct_view, const ZipReader& zip) {
  const DrawingRecordScan scan = scan_drawing_records(sheet.xlsb_tail().after_hyperlinks);
  const bool has_drawing_rel = std::any_of(sheet.unknown_relationships().begin(), sheet.unknown_relationships().end(),
                                           [](const UnknownRelationship& rel) { return rel.type == kRelDrawing; });
  if (scan.count == 0U && !has_drawing_rel) {
    return;
  }
  const UnknownRelationship* lifted = nullptr;
  const std::string_view refusal = DrawingLiftRefusal(sheet, scan, ct_view, zip, &lifted);
  if (lifted == nullptr) {
    StructuredLog("xlsb.reader.drawing_retained")
        .field("sheet_index", static_cast<std::int64_t>(sheet_index))
        .field("reason", refusal)
        .warn();
    return;
  }
  XlsbSheetTail tail = sheet.xlsb_tail();
  strip_drawing_records(tail.after_hyperlinks);
  sheet.set_xlsb_tail(std::move(tail));
  sheet.set_drawing_rel_target(lifted->target);
  const std::string rel_id = lifted->id;
  std::vector<UnknownRelationship> rels = sheet.unknown_relationships();
  rels.erase(
      std::remove_if(rels.begin(), rels.end(), [&rel_id](const UnknownRelationship& rel) { return rel.id == rel_id; }),
      rels.end());
  sheet.set_unknown_relationships(std::move(rels));
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
                      kReadXlsbContext);
  }
  ASSIGN_OR_RETURN(auto ct_bytes, zip.read_entry("[Content_Types].xml"));
  auto ct_view_or = LoadContentTypes(ct_bytes);
  if (!ct_view_or) {
    return std::move(ct_view_or.error());
  }
  ContentTypesView ct_view = ct_view_or.take();

  // 2. _rels/.rels — locate the workbook part path.
  if (!zip.has_entry("_rels/.rels")) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "_rels/.rels: missing from package", kReadXlsbContext);
  }
  ASSIGN_OR_RETURN(auto root_rels, zip.read_entry("_rels/.rels"));
  ASSIGN_OR_RETURN(const std::string workbook_path, ooxml::resolve_office_document_path(root_rels));
  ASSIGN_OR_RETURN(auto package_rels, ooxml::read_unknown_package_rels(root_rels, "xlsb_reader"));

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
  ASSIGN_OR_RETURN(const std::vector<XlsbSupBook> sup_books, DecodeSupBooks(wb_bytes));
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
  ApplyWorkbookBinInfo(workbook_info, wb);
  std::vector<std::string> sheet_part_paths;
  sheet_part_paths.reserve(bundle.size());
  for (const SheetBundleEntry& b : bundle) {
    if (b.name.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: BrtBundleSh with empty name",
                        kReadXlsbContext);
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
      if (added.error().code != FormulonErrorCode::kInvalidSheetName &&
          added.error().code != FormulonErrorCode::kAlreadyExists) {
        return std::move(added.error());
      }
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "workbook.bin: BrtBundleSh name is invalid or collides with an earlier sheet",
                        "context=xlsb_reader sheet=\"" + b.name + "\"");
    }
    wb.sheet(wb.sheet_count() - 1U).set_source_sheet_id(b.tab_id);
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

  // Table parts decode before any formula so a structured reference
  // (PtgList) resolves; the sheet loop loads these relationships again.
  XlsbTableParts table_parts;
  ReadDiagnostics table_diagnostics;
  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string rels_path = ooxml::rels_path_for_part(sheet_part_paths[i]);
    if (!zip.has_entry(rels_path)) {
      continue;
    }
    auto rels_or = LoadSheetRelationships(zip, rels_path, ooxml::dir_of(sheet_part_paths[i]));
    if (!rels_or) {
      return std::move(rels_or.error());
    }
    read_xlsb_table_parts(zip, i, rels_or.value(), table_parts);
  }
  wb.set_tables(xlsb_table_directory(table_parts));

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
  // Slicer cache lists stay future records keyed by workbook relationship ids.
  wb.set_xlsb_book_frt(read_book_slicer_cache_ids(ByteSpan{wb_bytes.data(), wb_bytes.size()}));

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
  if (zip.has_entry(kMetadataBinPath)) {
    auto metadata_or = zip.read_entry(kMetadataBinPath);
    if (metadata_or) {
      dynamic_array_ifmd =
          find_dynamic_array_cell_meta_index(ByteSpan{metadata_or.value().data(), metadata_or.value().size()});
    }
  }

  // Rich errors are restored while each sheet is decoded, before its spill
  // regions are registered; a constant saved with a spill value waits for the pivots.
  RichErrorSet rich_errors;
  std::vector<DeferredSpillAnchor> deferred_spill_anchors;
  const bool rich_consumed = ConsumeRichValueParts(zip, &rich_errors, consumed_parts, wb_rels_or.value().unknown_rels);

  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string& sheet_path = sheet_part_paths[i];
    if (!zip.has_entry(sheet_path)) {
      std::string ctx("context=xlsb_reader sheet_path=");
      ctx.append(sheet_path);
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "sheet binary part missing from package", std::move(ctx));
    }
    ASSIGN_OR_RETURN(auto sheet_bytes, zip.read_entry(sheet_path));
    auto state_or = DecodeSheetBin(sheet_bytes, i, wb, sst_entries, sst_phonetic, sst_phonetic_props, text_storage,
                                   sheet_names, name_table, sheet_ranges, external_books, &undecoded_formula_count,
                                   rich_consumed ? &rich_errors : nullptr, &deferred_spill_anchors);
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
  ingest_feature_formulas(wb, Workbook::FormulaSource::kRendered);

  // Every sheet view and visibility is loaded: settle the active sheet as Excel does on open.
  wb.apply_open_active_sheet_rule();

  // Table parts the sheets relate to, through the relationships stored above.
  std::unordered_map<std::string, std::vector<std::uint8_t>> retained_table_bytes;
  const std::vector<std::string> dropped_tables =
      load_xlsb_tables(wb, std::move(table_parts), sheet_names, name_table, sheet_ranges, external_books,
                       consumed_parts, retained_table_bytes, table_diagnostics);

  // Drawings the sheets link through one BrtDrawing record and its relationship.
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    LiftSheetDrawing(wb.sheet(i), i, ct_view, zip);
  }

  // Pivot caches and tables, reached through the sheet relationships the
  // loop above stored. Runs after every sheet is decoded because a table
  // is attached to its sheet and its cache is workbook-level.
  LoadPivotParts(zip, wb, retained_part_origins, deferred_spill_anchors);

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

  // A retained table part carries the bytes `load_xlsb_tables` cleared, not the archive's.
  auto read_retained = [&](const std::string& name) -> Expected<std::vector<std::uint8_t>, Error> {
    if (auto it = retained_table_bytes.find(name); it != retained_table_bytes.end()) {
      return std::move(it->second);
    }
    return zip.read_entry(name);
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
    ASSIGN_OR_RETURN(auto part_bytes, read_retained(part_name));
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
    ASSIGN_OR_RETURN(auto part_bytes, read_retained(entry));
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
  // The parts only a dropped table part's rels referenced go with it, once every model root is in place.
  std::vector<std::string> orphaned;
  for (const std::string& path : dropped_tables) {
    auto rels_bytes = zip.read_entry(ooxml::rels_path_for_part(path));
    if (!rels_bytes) {
      continue;
    }
    if (auto rels = ooxml::parse_part_rels(rels_bytes.value(), path.substr(0, path.rfind('/')))) {
      for (const ooxml::DrawingRel& rel : rels.value()) {
        if (!rel.external) {
          orphaned.push_back(rel.target);
        }
      }
    }
  }
  ooxml::remove_orphans(wb, std::move(orphaned));

  wb.apply_legacy_implicit_intersections();

  XlsbReadResult result{std::move(wb),
                        cells_read,
                        undecoded_formula_count,
                        undecoded_defined_name_count,
                        dropped_part_count,
                        dropped_record_count,
                        table_diagnostics.skipped_feature_count};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
