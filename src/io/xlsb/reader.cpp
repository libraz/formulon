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
#include <limits>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "defined_name.h"
#include "hyperlink_address.h"
#include "io/cf_merge.h"
#include "io/cf_reader.h"
#include "io/comments_reader.h"
#include "io/ooxml/package_rels_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_graph.h"
#include "io/ooxml/relationship_xml.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/package_reader_common.h"
#include "io/pivot_table_reader.h"
#include "io/rich_value_errors.h"
#include "io/rich_values.h"
#include "io/sheet_phonetic.h"
#include "io/threaded_comments_io.h"
#include "io/xlsb/chartsheet_records.h"
#include "io/xlsb/comments_records.h"
#include "io/xlsb/drawing_record.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/package_loader.h"
#include "io/xlsb/part_paths.h"
#include "io/xlsb/pivot_cache_bridge.h"
#include "io/xlsb/pivot_cache_definition_records.h"
#include "io/xlsb/pivot_cache_ids.h"
#include "io/xlsb/pivot_part_plan.h"
#include "io/xlsb/pivot_table_bridge.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/slicer_records.h"
#include "io/xlsb/sst_reader.h"
#include "io/xlsb/styles_reader.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/user_book_views.h"
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
#include "utils/index_sort.h"
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

/// One cache the workbook's PIVOTCACHEIDS block lists: its id and definition part.
struct BookPivotCache {
  std::uint32_t id = 0;
  std::string definition_path;
};

/// The caches `block` lists, in order, with each relationship id resolved through `book_rels`; empty when the
/// block does not frame.
std::vector<BookPivotCache> BookPivotCaches(const std::vector<std::uint8_t>& block,
                                            const std::vector<UnknownRelationship>& book_rels) {
  const BookPivotCacheIds listed = list_book_pivot_cache_ids(block);
  std::vector<BookPivotCache> out;
  if (!listed.complete) {
    return out;
  }
  for (const BookPivotCacheId& entry : listed.entries) {
    for (const UnknownRelationship& rel : book_rels) {
      if (!rel.target_external && rel.type == kRelPivotCacheDefinition && rel.id == entry.rel_id) {
        out.push_back(BookPivotCache{entry.id, rel.target});
      }
    }
  }
  return out;
}

/// The bytes of `path`, or nullopt when the package lacks it.
std::optional<std::vector<std::uint8_t>> ReadPart(const ZipReader& zip, const std::string& path) {
  if (path.empty() || !zip.has_entry(path)) {
    return std::nullopt;
  }
  auto bytes = zip.read_entry(path);
  return bytes ? std::optional<std::vector<std::uint8_t>>(std::move(bytes.value())) : std::nullopt;
}

/// True when `path` is a `.bin` part directly inside `dir` (which ends in '/').
bool DirectBinPart(std::string_view path, std::string_view dir) {
  return strings::starts_with(path, dir) && strings::ends_with(path, ".bin") &&
         path.find('/', dir.size()) == std::string_view::npos;
}

/// The internal target of `part_path`'s first relationship of `type`; empty when it has none.
std::string RelatedPart(const ZipReader& zip, const std::string& part_path, std::string_view type) {
  auto target = FindRelationship(zip, part_path, "Type", type, "pivot rels");
  return target ? std::move(target.value()) : std::string();
}

/// Reads the remote workbook URLs recorded in an external link part's own
/// rels file into `record`: the relationship `record.body_rel_id` names is
/// the target, and the ones `record.absolute_rel_id` / `record.relative_rel_id`
/// name (from the part's alternate-URL record) are the absolute and relative
/// URLs Excel records beside it. Without a body rel id
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
        const std::string_view type = rel.attribute("Type").value();
        const bool missing = type == kRelExternalLinkPathMissing;
        if (type != kRelExternalLinkPath && !missing) {
          return Expected<void, Error>::Ok();
        }
        const std::string_view id = rel.attribute("Id").value();
        const bool is_body = body_known ? id == record.body_rel_id : record.target.empty();
        if (is_body) {
          record.target = rel.attribute("Target").value();
          record.target_path_missing = missing;
        } else if (!missing && body_known && !record.absolute_rel_id.empty() && id == record.absolute_rel_id) {
          record.absolute_target = rel.attribute("Target").value();
        } else if (!missing && body_known && !record.relative_rel_id.empty() && id == record.relative_rel_id) {
          record.relative_target = rel.attribute("Target").value();
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

/// Renumbers the external books by first appearance in the `BrtExternSheet` list; books no entry
/// names follow in stored order. `books` and `ranges` are rewritten to the same new numbers.
void NumberExternalBooksByFirstUse(std::vector<XlsbSupBook>& books, std::vector<XlsbSheetRange>& ranges) {
  std::map<std::uint32_t, std::uint32_t> renumber;
  std::uint32_t next = 1;
  const auto assign = [&](std::uint32_t old_number) {
    if (old_number != 0 && renumber.emplace(old_number, next).second) {
      ++next;
    }
  };
  for (const XlsbSheetRange& range : ranges) {
    if (range.external_book != std::numeric_limits<std::uint32_t>::max()) {
      assign(range.external_book);
    }
  }
  for (const XlsbSupBook& sup : books) {
    assign(sup.external_book);
  }
  for (XlsbSupBook& sup : books) {
    if (sup.external_book != 0) {
      sup.external_book = renumber[sup.external_book];
    }
  }
  for (XlsbSheetRange& range : ranges) {
    const auto it = renumber.find(range.external_book);
    if (it != renumber.end()) {
      range.external_book = it->second;
    }
  }
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
  std::vector<const XlsbSupBook*> ordered;
  for (const XlsbSupBook& sup : books) {
    if (sup.external_book != 0) {  // 0 is this workbook.
      ordered.push_back(&sup);
    }
  }
  sort_by_index(ordered,
                [](const XlsbSupBook* l, const XlsbSupBook* r) { return l->external_book < r->external_book; });
  std::vector<ExternalLinkRecord> links;
  for (const XlsbSupBook* sup_ptr : ordered) {
    const XlsbSupBook& sup = *sup_ptr;
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
        } else if (auto book_or = read_external_link_bin(ByteSpan{bytes.data(), bytes.size()}, &record.body_rel_id,
                                                         &record.absolute_rel_id, &record.relative_rel_id)) {
          record.kind = ExternalLinkRecord::Kind::kExternalBook;
          record.book = std::move(book_or.value());
        } else {
          record.body_rel_id.clear();
          record.absolute_rel_id.clear();
          record.relative_rel_id.clear();
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

/// Consumes `xl/metadata.bin` and the rich values into `*set` when every part
/// reads (`read_rich_value_parts`) and no external link carries a
/// `BrtValueMeta`; their workbook relationships go with them, and the writer
/// regenerates the parts. `richValueRel` with its rels, the supporting-property
/// bags and `richStyles` stay passthrough. Otherwise every part stays
/// passthrough and the metadata part's value entries are counted. No sheet of
/// a binary workbook is kept opaque, so that condition of the shared
/// consumption rule always holds.
RichValueConsumption ConsumeRichValueParts(const ZipReader& zip, RichValueSet* set,
                                           std::unordered_set<std::string>& consumed_parts,
                                           std::vector<UnknownRelationship>& unknown_rels) {
  RichValueConsumption out;
  if (!zip.has_entry(kMetadataBinPath)) {
    return out;
  }
  constexpr std::string_view kExternalLinks = "xl/externalLinks/";
  bool consumable = true;
  for (const std::string& name : zip.list_entries()) {
    // An external link's cached cells naming a value-metadata entry belong to a set this reader does not own.
    if (consumable && name.compare(0, kExternalLinks.size(), kExternalLinks) == 0 && name.size() > 4U &&
        name.compare(name.size() - 4U, 4U, ".bin") == 0) {
      auto link = zip.read_entry(name);
      consumable = static_cast<bool>(link);
      for (ByteSpan cursor = link ? ByteSpan{link.value().data(), link.value().size()} : ByteSpan{};
           consumable && cursor.size > 0U;) {
        auto record = read_record(cursor);
        consumable = record && record.value().type != static_cast<std::uint16_t>(XlsbRecordType::BrtValueMeta);
      }
    }
  }
  RichValueSource source;
  load_rich_value_source(zip, kMetadataBinPath, unknown_rels, &source);
  if (source.package.metadata == nullptr) {
    return out;
  }
  const ByteSpan metadata{source.package.metadata->data(), source.package.metadata->size()};
  if (consumable) {
    RichValueRead read = read_rich_value_parts(source.package, parse_metadata_bin(metadata));
    out.consumed = commit_rich_value_read(read, source.parts, set, consumed_parts, unknown_rels);
  }
  if (!out.consumed) {
    out.verbatim_value_entries = metadata_bin_value_entry_count(metadata);
  }
  return out;
}

/// Decodes every pivot cache and table into the model, a cache with every
/// table drawing from it or not at all.
///
/// A group the bridges cannot read is skipped, not failed: its parts stay
/// passthrough with their relationships, as every pivot part did before the
/// model could write them. A modelled group's part bytes are retained, tagged
/// with the model object each fingerprints against, for the writer to keep
/// while the group is unchanged; its relationships are consumed, since the
/// writer relates whichever parts it writes.
void LoadPivotParts(const ZipReader& zip, Workbook& wb, std::vector<UnknownRelationship>& book_rels,
                    std::unordered_set<std::string>& consumed_parts, RetainedPartOrigins& origins,
                    const std::vector<DeferredSpillAnchor>& deferred_spill_anchors) {
  struct TableRead {
    std::size_t sheet = 0;
    std::string rel_id;
    std::string path;
    std::string definition_path;
    pivot::PivotTable table;
  };
  std::vector<TableRead> tables;
  // Definition parts a table drawing from them, or the cache itself, could not be read for.
  std::unordered_set<std::string> refused;
  // A table part related more than once, from one sheet or several, is not modelled.
  std::unordered_map<std::string, std::size_t> table_rels;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
      if (rel.type == kRelPivotTable && !rel.target_external) {
        ++table_rels[rel.target];
      }
    }
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
      if (rel.type != kRelPivotTable || rel.target_external || rel.target.empty()) {
        continue;
      }
      const std::string definition = RelatedPart(zip, rel.target, kRelPivotCacheDefinition);
      // The writer relates a modelled group's parts from where Excel keeps them.
      if (table_rels[rel.target] != 1U || !DirectBinPart(rel.target, "xl/pivotTables/") ||
          !DirectBinPart(definition, "xl/pivotCache/")) {
        StructuredLog("xlsb.pivot.skipped")
            .field("part", rel.target)
            .field("reason", std::string_view("table part related twice or outside the pivot folders"))
            .warn();
        refused.insert(definition);
        continue;
      }
      const std::optional<std::vector<std::uint8_t>> bytes = ReadPart(zip, rel.target);
      auto table =
          bytes ? pivot_table_from_xlsb(ByteSpan{bytes->data(), bytes->size()})
                : Expected<pivot::PivotTable, Error>(make_error(FormulonErrorCode::kIoRelationshipBroken,
                                                                "pivot table part missing", "context=xlsb_reader"));
      if (!table || definition.empty()) {
        StructuredLog("xlsb.pivot.skipped")
            .field("part", rel.target)
            .field("reason", table ? std::string("no cache definition") : table.error().message)
            .warn();
        refused.insert(definition);
        continue;
      }
      tables.push_back(TableRead{i, rel.id, rel.target, definition, std::move(table.value())});
    }
  }
  std::unordered_map<std::string, std::uint32_t> modelled;
  for (const BookPivotCache& entry : BookPivotCaches(wb.xlsb_pivot_cache_ids(), book_rels)) {
    const std::string& definition = entry.definition_path;
    if (refused.count(definition) != 0U || modelled.count(definition) != 0U ||
        wb.find_pivot_cache(entry.id) != nullptr) {
      continue;
    }
    if (!DirectBinPart(definition, "xl/pivotCache/")) {
      continue;
    }
    const std::optional<std::vector<std::uint8_t>> definition_bytes = ReadPart(zip, definition);
    if (!definition_bytes) {
      continue;
    }
    // A cache saved without its data (`saveData` off) has no records part and names none.
    const auto header = decode_pivot_cache_definition(ByteSpan{definition_bytes->data(), definition_bytes->size()});
    const bool save_data = !header || (header.value().flags & pivot_cache_definition::kCacheSaveData) != 0U;
    const std::string records_path = save_data ? RelatedPart(zip, definition, kRelPivotCacheRecords) : std::string();
    if (save_data && !DirectBinPart(records_path, "xl/pivotCache/")) {
      continue;
    }
    const std::optional<std::vector<std::uint8_t>> records_bytes =
        save_data ? ReadPart(zip, records_path) : std::optional<std::vector<std::uint8_t>>(std::vector<std::uint8_t>());
    if (!records_bytes) {
      continue;
    }
    // A kept definition names its records by the id the writer relates them under; a rich cache names them in richInfo.
    if (save_data && header &&
        (header.value().rich ? header.value().rich->records_rel_id : header.value().records_rel_id) !=
            kPivotCacheRecordsRelId) {
      StructuredLog("xlsb.pivot.skipped")
          .field("part", definition)
          .field("reason", std::string_view("records relationship id differs from rId1"))
          .warn();
      continue;
    }
    auto cache = pivot_cache_from_xlsb(ByteSpan{definition_bytes->data(), definition_bytes->size()},
                                       ByteSpan{records_bytes->data(), records_bytes->size()}, wb.date1904());
    if (!cache) {
      StructuredLog("xlsb.pivot.skipped").field("part", definition).field("reason", cache.error().message).warn();
      continue;
    }
    cache.value().set_cache_id(entry.id);
    wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(std::move(cache.value())));
    modelled.emplace(definition, entry.id);
    consumed_parts.insert(ooxml::rels_path_for_part(definition));
    origins.set(definition,
                RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheDefinition, 0, 0, entry.id});
    if (save_data) {
      origins.set(records_path,
                  RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheRecords, 0, 0, entry.id});
    }
    book_rels.erase(std::remove_if(book_rels.begin(), book_rels.end(),
                                   [&definition](const UnknownRelationship& rel) {
                                     return !rel.target_external && rel.target == definition;
                                   }),
                    book_rels.end());
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const std::size_t first_new_pivot = wb.sheet(i).pivot_tables().size();
    std::unordered_set<std::string> consumed_rel_ids;
    for (TableRead& read : tables) {
      if (read.sheet != i) {
        continue;
      }
      const auto cache = modelled.find(read.definition_path);
      if (cache == modelled.end()) {
        StructuredLog("xlsb.pivot.skipped")
            .field("part", read.path)
            .field("reason", std::string_view("cache not modelled"))
            .warn();
        continue;
      }
      read.table.set_pivot_cache_id(cache->second);
      consumed_parts.insert(ooxml::rels_path_for_part(read.path));
      origins.set(read.path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotTable, i,
                                                   wb.sheet(i).pivot_tables().size(), 0});
      consumed_rel_ids.insert(read.rel_id);
      wb.sheet(i).add_pivot_table(std::make_unique<pivot::PivotTable>(std::move(read.table)));
    }
    if (!consumed_rel_ids.empty()) {
      std::vector<UnknownRelationship> rels = wb.sheet(i).unknown_relationships();
      rels.erase(std::remove_if(rels.begin(), rels.end(),
                                [&consumed_rel_ids](const UnknownRelationship& rel) {
                                  return consumed_rel_ids.count(rel.id) != 0U;
                                }),
                 rels.end());
      wb.sheet(i).set_unknown_relationships(std::move(rels));
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

/// Moves a sheet's notes, their VML drawing and its threads from the retained
/// BrtLegacyDrawing record and relationships into the model, as the xlsx
/// reader holds them. The VML part stays passthrough, as it does for xlsx. A
/// sheet whose parts do not resolve one to one stays retained and is logged.
Expected<void, Error> LiftSheetNotes(Sheet& sheet, std::size_t sheet_index, const ZipReader& zip,
                                     const std::vector<FontRecord>& fonts,
                                     std::unordered_set<std::string>& consumed_parts) {
  const UnknownRelationship* comments = nullptr;
  const UnknownRelationship* vml = nullptr;
  const UnknownRelationship* threads = nullptr;
  std::size_t counts[3] = {0, 0, 0};
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    const UnknownRelationship** slot = rel.type == kRelComments          ? &comments
                                       : rel.type == kRelVmlDrawing      ? &vml
                                       : rel.type == kRelThreadedComment ? &threads
                                                                         : nullptr;
    if (slot != nullptr && !rel.target_external) {
      *slot = &rel;
      ++counts[slot == &comments ? 0 : slot == &vml ? 1 : 2];
    }
  }
  if (comments == nullptr && threads == nullptr) {
    return {};
  }
  const DrawingRecordScan scan = scan_drawing_records(sheet.xlsb_tail().after_hyperlinks, kBrtLegacyDrawing);
  if (counts[0] != 1U || counts[1] != 1U || counts[2] > 1U || scan.count != 1U || !scan.rel_id ||
      *scan.rel_id != vml->id || !zip.has_entry(comments->target) || !zip.has_entry(vml->target) ||
      (threads != nullptr && !zip.has_entry(threads->target))) {
    StructuredLog("xlsb.reader.notes_retained").field("sheet_index", static_cast<std::int64_t>(sheet_index)).warn();
    return {};
  }
  if (threads != nullptr) {
    ASSIGN_OR_RETURN(auto thread_bytes, zip.read_entry(threads->target));
    ASSIGN_OR_RETURN(sheet.mutable_threaded_comments(), read_threaded_comments(thread_bytes));
    consumed_parts.insert(threads->target);
  }
  ASSIGN_OR_RETURN(auto vml_bytes, zip.read_entry(vml->target));
  const std::string_view vml_text(reinterpret_cast<const char*>(vml_bytes.data()), vml_bytes.size());
  sheet.set_comment_vml_extra_xml(vml_non_note_shapes(vml_text));
  ASSIGN_OR_RETURN(auto comment_bytes, zip.read_entry(comments->target));
  std::vector<std::string> authors;
  ASSIGN_OR_RETURN(auto notes,
                   read_comments_bin(ByteSpan{comment_bytes.data(), comment_bytes.size()}, fonts, &authors));
  drop_thread_stubs(notes, sheet.threaded_comments());
  apply_vml_note_shapes(vml_text, notes);
  sheet.mutable_comments() = std::move(notes);
  authors.erase(
      std::remove_if(authors.begin(), authors.end(), [](const std::string& a) { return is_thread_stub_author(a); }),
      authors.end());
  sheet.set_comment_authors(std::move(authors));
  sheet.set_comment_vml_path(vml->target);
  sheet.set_comment_vml_shapes(sheet.comment_shapes());
  consumed_parts.insert(comments->target);

  XlsbSheetTail tail = sheet.xlsb_tail();
  strip_drawing_records(tail.after_hyperlinks, kBrtLegacyDrawing);
  sheet.set_xlsb_tail(std::move(tail));
  const std::string ids[3] = {comments->id, vml->id, threads != nullptr ? threads->id : std::string()};
  std::vector<UnknownRelationship> rels = sheet.unknown_relationships();
  rels.erase(std::remove_if(rels.begin(), rels.end(),
                            [&ids](const UnknownRelationship& rel) {
                              return rel.id == ids[0] || rel.id == ids[1] || (!ids[2].empty() && rel.id == ids[2]);
                            }),
             rels.end());
  sheet.set_unknown_relationships(std::move(rels));
  return {};
}

/// Reads the person list the threads name; its relationship leaves the unknown list with it.
Expected<void, Error> ConsumePersons(const ZipReader& zip, Workbook& wb,
                                     std::unordered_set<std::string>& consumed_parts,
                                     std::vector<UnknownRelationship>& unknown_rels) {
  for (auto it = unknown_rels.begin(); it != unknown_rels.end(); ++it) {
    if (it->type != kRelPerson || it->target_external || !zip.has_entry(it->target)) {
      continue;
    }
    ASSIGN_OR_RETURN(auto bytes, zip.read_entry(it->target));
    ASSIGN_OR_RETURN(wb.mutable_persons(), read_persons(bytes));
    consumed_parts.insert(it->target);
    unknown_rels.erase(it);
    break;
  }
  return {};
}

/// Renames each defined name spelled like a built-in in the load locale to that built-in, under every record naming it.
void FoldBuiltinNames(std::vector<XlsbName>& name_table, const Workbook& wb, const ReadOptions& options) {
  std::vector<LoadedNameScope> scopes;
  scopes.reserve(name_table.size());
  for (const XlsbName& entry : name_table) {
    scopes.push_back(LoadedNameScope{entry.name, entry.itab});
  }
  std::vector<std::string> folds = loaded_builtin_name_folds(scopes, load_locale(wb, options));
  for (std::size_t i = 0; i < folds.size(); ++i) {
    if (!folds[i].empty()) {
      name_table[i].name = std::move(folds[i]);
    }
  }
}

/// Adds one sheet per workbook bundle entry, validating its name and relationship, and collects each sheet part path.
Expected<void, Error> AddBundleSheets(const std::vector<SheetBundleEntry>& bundle, const WorkbookRels& wb_rels,
                                      Workbook& wb, std::vector<std::string>& sheet_part_paths) {
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
      return sheet_add_error(std::move(added.error()),
                             "workbook.bin: BrtBundleSh name is invalid or collides with an earlier sheet",
                             "context=xlsb_reader sheet=\"" + b.name + "\"");
    }
    wb.sheet(wb.sheet_count() - 1U).set_source_sheet_id(b.tab_id);
    if (b.visibility != SheetVisibility::kVisible) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().set_visibility(b.visibility);
    }
    if (it->second.relationship_type != kRelWorksheet) {
      // A chartsheet or other non-worksheet keeps its tab; its parts stay passthrough.
      wb.sheet(wb.sheet_count() - 1U).set_opaque_ooxml_sheet(it->second.path, it->second.relationship_type);
    }
    sheet_part_paths.push_back(it->second.path);
  }
  return {};
}

/// Decodes the table parts every non-opaque sheet relates to, before any formula needs a structured reference.
Expected<void, Error> ReadSheetTableParts(const ZipReader& zip, const Workbook& wb,
                                          const std::vector<std::string>& sheet_part_paths,
                                          XlsbTableParts& table_parts) {
  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string rels_path = ooxml::rels_path_for_part(sheet_part_paths[i]);
    if (wb.sheet(i).is_opaque_ooxml_sheet() || !zip.has_entry(rels_path)) {
      continue;
    }
    auto rels_or = LoadSheetRelationships(zip, rels_path, ooxml::dir_of(sheet_part_paths[i]));
    if (!rels_or) {
      return std::move(rels_or.error());
    }
    read_xlsb_table_parts(zip, i, rels_or.value(), table_parts);
  }
  return {};
}

/// Restores the Compatibility Version stamp from the workbook part, absent unless the part carries one.
void ReadCompatibilityStamp(const std::vector<std::uint8_t>& wb_bytes, Workbook& wb) {
  std::optional<std::uint32_t> set_version;
  std::optional<std::uint32_t> warn_below;
  for (const FramedRecord& rec : split_records(wb_bytes)) {
    if (rec.type == kWorkbookCompatibilityVersion && rec.payload.size == 12U) {
      warn_below = load_u32(rec.payload.data + 4) != 0U ? std::optional<std::uint32_t>(load_u32(rec.payload.data + 4))
                                                        : std::nullopt;
      set_version = load_u32(rec.payload.data + 8);
    }
  }
  wb.restore_compatibility_stamp(set_version, warn_below);
}

/// Marks the package parts the reader models or deliberately rebuilds as consumed.
void SeedConsumedParts(const ZipReader& zip, const std::string& workbook_path, const WorkbookRels& wb_rels,
                       std::unordered_set<std::string>& consumed_parts) {
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  consumed_parts.insert(ooxml::rels_path_for_part(workbook_path));
  if (!wb_rels.sst_path.empty()) {
    consumed_parts.insert(wb_rels.sst_path);
  }
  // The calculation chain and the volatile dependencies are stale after any
  // edit, Excel rebuilds them when absent, and neither writer writes them.
  for (const std::string& path : wb_rels.rebuilt_part_paths) {
    if (zip.has_entry(path)) {
      consumed_parts.insert(path);
    }
  }
  if (zip.has_entry("xl/calcChain.bin")) {
    consumed_parts.insert("xl/calcChain.bin");
  }
}

/// Applies a chartsheet part's drawing, tab colour and protection to its sheet.
Expected<void, Error> LoadChartsheetPart(const ZipReader& zip, const std::string& sheet_path, Sheet& sheet) {
  ASSIGN_OR_RETURN(const auto chartsheet_bytes, zip.read_entry(sheet_path));
  if (const Expected<ChartsheetXml, Error> chartsheet =
          chartsheet_xml_from_records(ByteSpan{chartsheet_bytes.data(), chartsheet_bytes.size()})) {
    apply_chartsheet_part(chartsheet.value().xml, sheet);
  }
  // The colour stands even when the record's other state keeps it out of the converted `<sheetPr>`.
  if (const std::optional<ColorSpec> tab_color =
          chartsheet_tab_color_from_records(ByteSpan{chartsheet_bytes.data(), chartsheet_bytes.size()})) {
    sheet.mutable_print_settings().tab_color = *tab_color;
  }
  sheet.set_chartsheet_protection(
      chartsheet_protection_from_records(ByteSpan{chartsheet_bytes.data(), chartsheet_bytes.size()}));
  return {};
}

/// Loads a sheet's relationships, resolves its hyperlinks and keeps the relationships on the sheet.
Expected<void, Error> LoadSheetRels(const ZipReader& zip, const std::string& sheet_path, Sheet& sheet,
                                    std::unordered_set<std::string>& consumed_parts) {
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
    auto hyperlinks = ResolveSheetHyperlinks(sheet, relationships);
    if (!hyperlinks) {
      return std::move(hyperlinks.error());
    }
    sheet.set_unknown_relationships(std::move(relationships));
    consumed_parts.insert(sheet_rels_path);
  } else {
    std::vector<UnknownRelationship> no_relationships;
    auto hyperlinks = ResolveSheetHyperlinks(sheet, no_relationships);
    if (!hyperlinks) {
      return std::move(hyperlinks.error());
    }
  }
  return {};
}

/// Captures every package part the reader did not consume: Override-listed parts keep their content type, the rest
/// resolve through Default.
Expected<void, Error> CapturePassthroughParts(
    const ZipReader& zip, const Workbook& wb, const ContentTypesView& ct_view,
    const std::unordered_set<std::string>& consumed_parts, const RetainedPartOrigins& retained_part_origins,
    std::unordered_map<std::string, std::vector<std::uint8_t>>& retained_table_bytes, RichValueSet& rich_values,
    std::vector<PassthroughPart>& unknown_parts, std::uint32_t& dropped_part_count, std::string& first_dropped) {
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
    if (tag.origin == PassthroughPart::RetainedOrigin::kStyles) {
      part.dxf_reference_fingerprint = current_dxf_reference_fingerprint(wb);
    }
  };

  // A retained table part carries the bytes `load_xlsb_tables` cleared, not the archive's.
  auto read_retained = [&](const std::string& name) -> Expected<std::vector<std::uint8_t>, Error> {
    if (auto it = retained_table_bytes.find(name); it != retained_table_bytes.end()) {
      return std::move(it->second);
    }
    return take_package_part(zip, rich_values, name);
  };
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
  return {};
}

/// Removes the parts only a dropped table part's relationships referenced.
void RemoveDroppedTableOrphans(const ZipReader& zip, Workbook& wb, const std::vector<std::string>& dropped_tables) {
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
}

}  // namespace

Expected<XlsbReadResult, Error> read_xlsb(ByteSpan bytes, const ReadOptions& options) {
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
  std::vector<XlsbName>& name_table = name_table_or.value();
  ASSIGN_OR_RETURN(std::vector<XlsbSupBook> sup_books, DecodeSupBooks(wb_bytes));
  auto sheet_ranges_or = DecodeExternSheet(wb_bytes, sup_books);
  if (!sheet_ranges_or) {
    return std::move(sheet_ranges_or.error());
  }
  std::vector<XlsbSheetRange>& sheet_ranges = sheet_ranges_or.value();
  NumberExternalBooksByFirstUse(sup_books, sheet_ranges);

  // 5. Build the workbook bottom-up.
  Workbook wb = Workbook::create_empty();
  // The package's style table is authoritative, including when the package
  // has none: back-filling the factory's seeded defaults would invent style
  // records the file never carried and add an unwanted styles part on write.
  wb.set_styles(StylesTable{});
  ApplyWorkbookBinInfo(workbook_info, wb);
  // A name spelled like a built-in in the load locale opens as that built-in, under every record naming it.
  FoldBuiltinNames(name_table, wb, options);
  std::vector<std::string> sheet_part_paths;
  sheet_part_paths.reserve(bundle.size());
  RETURN_IF_ERROR(AddBundleSheets(bundle, wb_rels, wb, sheet_part_paths));

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
  RETURN_IF_ERROR(ReadSheetTableParts(zip, wb, sheet_part_paths, table_parts));
  wb.set_tables(xlsb_table_directory(table_parts));

  // 5b. Register every non-hidden `BrtName` entry as a defined name.
  // Needs `sheet_names` (for qualified references inside a name's own
  // formula) and the complete `name_table` / `sheet_ranges` (for
  // self-consistent `PtgName` / `PtgRef3d` resolution), so this can
  // only run after both are fully built above.
  std::uint32_t undecoded_formula_count = 0;
  std::uint32_t undecoded_defined_name_count = 0;
  std::uint32_t skipped_feature_count = 0;
  if (auto r = RegisterDefinedNames(wb_bytes, name_table, sheet_names, sheet_ranges, external_books, wb,
                                    &undecoded_defined_name_count, &skipped_feature_count);
      !r) {
    return std::move(r.error());
  }
  // Slicer cache lists stay future records keyed by workbook relationship ids.
  wb.set_xlsb_book_frt(read_book_slicer_cache_ids(ByteSpan{wb_bytes.data(), wb_bytes.size()}));
  wb.set_xlsb_pivot_cache_ids(read_book_pivot_cache_ids(ByteSpan{wb_bytes.data(), wb_bytes.size()}));
  wb.set_xlsb_user_book_views(read_book_user_views(ByteSpan{wb_bytes.data(), wb_bytes.size()}));
  // The Compatibility Version stamp: absent unless the part carries one.
  ReadCompatibilityStamp(wb_bytes, wb);

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
    complete_loaded_fonts(styles, wb, options);
    retained_part_origins.set(wb_rels.styles_path,
                              RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kStyles, 0, 0, 0});
    wb.set_styles(std::move(styles));
  }

  // 7. Each sheet binary.
  std::uint32_t cells_read = 0;
  std::uint32_t dropped_record_count = 0;
  std::unordered_set<std::string> consumed_parts;
  SeedConsumedParts(zip, workbook_path, wb_rels, consumed_parts);

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

  // Rich values are restored while each sheet is decoded, before its spill
  // regions are registered; a constant saved with a spill value waits for the pivots.
  RichValueSet rich_values;
  std::vector<DeferredSpillAnchor> deferred_spill_anchors;
  std::vector<std::vector<CellAddress>> recalc_marked(sheet_part_paths.size());
  const RichValueConsumption rich_consumption =
      ConsumeRichValueParts(zip, &rich_values, consumed_parts, wb_rels_or.value().unknown_rels);
  rich_values.zip = &zip;

  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string& sheet_path = sheet_part_paths[i];
    if (wb.sheet(i).kind() == SheetKind::kChartsheet && zip.has_entry(sheet_path)) {
      RETURN_IF_ERROR(LoadChartsheetPart(zip, sheet_path, wb.sheet(i)));
    }
    if (wb.sheet(i).is_opaque_ooxml_sheet()) {
      continue;
    }
    if (!zip.has_entry(sheet_path)) {
      std::string ctx("context=xlsb_reader sheet_path=");
      ctx.append(sheet_path);
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "sheet binary part missing from package", std::move(ctx));
    }
    ASSIGN_OR_RETURN(auto sheet_bytes, zip.read_entry(sheet_path));
    auto state_or = DecodeSheetBin(sheet_bytes, i, wb, sst_entries, sst_phonetic, sst_phonetic_props, text_storage,
                                   sheet_names, name_table, sheet_ranges, external_books, &undecoded_formula_count,
                                   rich_consumption.consumed ? &rich_values : nullptr, &deferred_spill_anchors,
                                   rich_consumption.verbatim_value_entries);
    if (!state_or) {
      return std::move(state_or.error());
    }
    cells_read += state_or.value().cells_decoded;
    dropped_record_count += state_or.value().dropped_records;
    skipped_feature_count += state_or.value().skipped_features;
    recalc_marked[i] = std::move(state_or.value().recalc_marked);
    mark_sheet_dynamic_arrays(wb.sheet(i), dynamic_array_ifmd, state_or.value().cell_metadata);
    consumed_parts.insert(sheet_path);

    RETURN_IF_ERROR(LoadSheetRels(zip, sheet_path, wb.sheet(i), consumed_parts));
  }

  // Conditional-format and data-validation formulas name other books
  // through the same `[N]` cell formulas do.
  ingest_feature_formulas(wb, Workbook::FormulaSource::kRendered);
  merge_cf_rules_on_load(wb);

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

  // Notes and threads, through one BrtLegacyDrawing and the sheet's relationships, and the persons the threads name.
  RETURN_IF_ERROR(ConsumePersons(zip, wb, consumed_parts, wb_rels_or.value().unknown_rels));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    RETURN_IF_ERROR(LiftSheetNotes(wb.sheet(i), i, zip, wb.styles().fonts, consumed_parts));
  }

  // Pivot caches and tables, reached through the sheet relationships the
  // loop above stored. Runs after every sheet is decoded because a table
  // is attached to its sheet and its cache is workbook-level.
  LoadPivotParts(zip, wb, wb_rels_or.value().unknown_rels, consumed_parts, retained_part_origins,
                 deferred_spill_anchors);

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
  std::vector<PassthroughPart> unknown_parts;
  std::uint32_t dropped_part_count = 0;
  std::string first_dropped;
  RETURN_IF_ERROR(CapturePassthroughParts(zip, wb, ct_view, consumed_parts, retained_part_origins, retained_table_bytes,
                                          rich_values, unknown_parts, dropped_part_count, first_dropped));

  sort_passthrough_parts(unknown_parts);
  mark_model_derived_parts(wb, unknown_parts);

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
  RemoveDroppedTableOrphans(zip, wb, dropped_tables);

  // Last, once nothing else dirties formulas: the stale set Excel opens with follows fAlwaysCalc.
  finish_loaded_workbook(wb, options, recalc_marked);
  if (!sheet_ranges.empty()) {
    wb.set_xti_order(ToExternSheetEntries(sheet_ranges, sheet_names, external_books));
  }
  // The links keep their stored order: Excel renumbers an .xlsx on load, not an .xlsb.

  XlsbReadResult result{std::move(wb),
                        cells_read,
                        undecoded_formula_count,
                        undecoded_defined_name_count,
                        dropped_part_count,
                        dropped_record_count,
                        table_diagnostics.skipped_feature_count + skipped_feature_count};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
