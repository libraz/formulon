//
// Implementation of the OOXML writer's emission plan. Determines part
// paths, numeric ids, and resolves collisions between writer-generated
// and passthrough parts. No miniz state is touched here; this is pure
// metadata.

#include "io/ooxml/emission_plan.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "external_link.h"
#include "io/ooxml/external_link_writer.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/sheet_xml_builder.h"
#include "io/ooxml_defs.h"
#include "passthrough_part.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {

std::string NumberedPartPath(std::string_view prefix, std::uint32_t id, std::string_view suffix) {
  std::string path;
  path.reserve(prefix.size() + suffix.size() + 10U);
  path.append(prefix);
  path.append(std::to_string(id));
  path.append(suffix);
  return path;
}

bool HasPassthroughPart(const EmissionPlan& plan, std::string_view path) {
  return HasPassthroughPart(plan.passthrough_kept, path);
}

bool HasPassthroughPart(const std::vector<const PassthroughPart*>& passthrough_kept, std::string_view path) {
  for (const PassthroughPart* part : passthrough_kept) {
    if (part != nullptr && part->path == path) {
      return true;
    }
  }
  return false;
}

bool IsXlsbBinaryContentType(std::string_view content_type) {
  constexpr std::string_view kVendorPrefix = "application/vnd.ms-excel.";
  constexpr std::string_view kXmlSuffix = "+xml";
  if (content_type.size() <= kVendorPrefix.size() ||
      content_type.compare(0, kVendorPrefix.size(), kVendorPrefix) != 0) {
    return false;
  }
  return content_type.size() < kXmlSuffix.size() ||
         content_type.compare(content_type.size() - kXmlSuffix.size(), kXmlSuffix.size(), kXmlSuffix) != 0;
}

namespace {

/// Returns writer-owned paths reserved before passthrough collision handling.
/// Some reserved paths (notably empty sheet `.rels` parts) are ultimately not
/// emitted, but reserving them keeps a source passthrough copy from winning
/// the collision before its relationships can be planned.
std::unordered_set<std::string> BuildGeneratedPathSet(
    const Workbook& wb, const std::vector<EmissionPlan::PerSheetTable>& flat_tables,
    const std::vector<EmissionPlan::PivotCachePlan>& pivot_caches,
    const std::vector<std::vector<EmissionPlan::PivotTablePlan>>& pivot_tables_by_sheet,
    const std::vector<EmissionPlan::CommentsPlan>& comments_by_sheet,
    const std::vector<EmissionPlan::ExternalLinkPlan>& external_links, bool generated_shared_strings) {
  std::unordered_set<std::string> paths;
  paths.insert("[Content_Types].xml");
  paths.insert("_rels/.rels");
  paths.insert("xl/workbook.xml");
  paths.insert("xl/_rels/workbook.xml.rels");
  paths.insert("xl/styles.xml");
  if (generated_shared_strings) {
    paths.insert("xl/sharedStrings.xml");
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (!wb.sheet(i).is_opaque_ooxml_sheet()) {
      paths.insert("xl/worksheets/sheet" + std::to_string(i + 1) + ".xml");
    }
  }
  for (const EmissionPlan::PerSheetTable& t : flat_tables) {
    paths.insert(t.path);
    // Sheet rels for sheets that own tables are also generated.
  }
  // Pivot-cache parts: definition, records, and definition rels.
  for (const EmissionPlan::PivotCachePlan& c : pivot_caches) {
    paths.insert(c.definition_path);
    paths.insert(c.records_path);
    paths.insert(c.definition_rels_path);
  }
  // Pivot-table parts (one per pivot table, package-wide) plus the rels
  // file naming the cache definition each one draws from.
  for (const auto& per_sheet : pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      paths.insert(t.path);
      if (!t.cache_definition_target.empty()) {
        paths.insert(t.rels_path);
      }
    }
  }
  // Comment / VML parts (one per sheet that has at least one comment).
  for (const EmissionPlan::CommentsPlan& c : comments_by_sheet) {
    if (c.numeric_id == 0) {
      continue;
    }
    paths.insert(c.comments_path);
    paths.insert(c.vml_path);
    if (!c.threaded_path.empty()) {
      paths.insert(c.threaded_path);
    }
  }
  if (!wb.persons().empty()) {
    paths.insert(std::string(kPersonsPartPath));
  }
  // Per-link rels files for written external links — the writer generates
  // these from the captured `ExternalLinkRecord`s.
  for (const EmissionPlan::ExternalLinkPlan& link : external_links) {
    if (link.written) {
      paths.insert(ooxml::rels_path_for_part(link.part_path));
    }
  }
  // Sheet rels are model-owned paths even when the eventual rels document
  // turns out to have no relationships. Reserve every canonical path before
  // resolving passthrough collisions so a source copy can never win over the
  // writer's own sheet rels output.
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (!wb.sheet(i).is_opaque_ooxml_sheet()) {
      paths.insert("xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels");
    }
  }
  return paths;
}

/// Resolves the content type a passthrough part will actually be served
/// under: its own `<Override>` type when it has one, otherwise the
/// source `<Default>` registration for its extension. Default-typed
/// parts deliberately carry an empty `content_type` (the registry
/// travels separately on the Workbook), so the two must be recombined
/// before any decision keyed on the type.
std::string_view EffectiveContentType(const Workbook& wb, const PassthroughPart& part) {
  if (!part.content_type.empty()) {
    return part.content_type;
  }
  const std::string extension = ooxml::extension_of_part(part.path);
  if (extension.empty()) {
    return {};
  }
  for (const DefaultContentType& def : wb.default_content_types()) {
    if (ooxml::lowercase_extension(def.extension) == extension) {
      return def.content_type;
    }
  }
  return {};
}

/// True when `path` names a loaded part this package can carry verbatim.
bool HasRetainedPart(const Workbook& wb, std::string_view path) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == path) {
      return !IsXlsbBinaryContentType(EffectiveContentType(wb, part));
    }
  }
  return false;
}

bool EndsWith(std::string_view text, std::string_view suffix) {
  return text.size() >= suffix.size() && text.compare(text.size() - suffix.size(), suffix.size(), suffix) == 0;
}

/// Plans every external link in index order. A loaded body the model still
/// matches is kept; an OLE / DDE link without one cannot be rebuilt and is
/// not written; every other link gets a generated body, at its own `.xml`
/// path or else at the smallest `externalLink<k>.xml` no record or loaded
/// part uses.
void PlanExternalLinks(const Workbook& wb, std::size_t first_rid, EmissionPlan& plan) {
  const std::vector<const ExternalLinkRecord*> records = wb.external_links_by_index(/*skip_ole_dde=*/false);
  std::unordered_set<std::string> used_paths;
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    used_paths.insert(rec.part_path);
  }
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    used_paths.insert(part.path);
  }
  std::vector<std::uint32_t> written_indices;
  std::uint32_t next_part_id = 1;
  for (std::size_t i = 0; i < records.size(); ++i) {
    const ExternalLinkRecord& rec = *records[i];
    EmissionPlan::ExternalLinkPlan entry;
    entry.record = &rec;
    entry.workbook_rid = static_cast<std::uint32_t>(first_rid + i);
    entry.part_path = rec.part_path;
    if (!rec.body_stale && HasRetainedPart(wb, rec.part_path)) {
      entry.written = true;
    } else if (rec.kind != ExternalLinkRecord::Kind::kOleLink && rec.kind != ExternalLinkRecord::Kind::kDdeLink) {
      if (!EndsWith(rec.part_path, ".xml")) {
        while (used_paths.count(NumberedPartPath("xl/externalLinks/externalLink", next_part_id, ".xml")) != 0U) {
          ++next_part_id;
        }
        entry.part_path = NumberedPartPath("xl/externalLinks/externalLink", next_part_id++, ".xml");
        used_paths.insert(entry.part_path);
      }
      entry.generated_body = BuildExternalLinkXml(rec);
      entry.written = true;
    }
    if (entry.written) {
      written_indices.push_back(rec.index);
    }
    plan.external_links.push_back(std::move(entry));
  }
  plan.external_link_ordinals = ExternalLinkOrdinals(wb, std::move(written_indices));
}

}  // namespace

EmissionPlan BuildEmissionPlan(const Workbook& wb, bool generated_shared_strings, WriteDiagnostics* diagnostics) {
  EmissionPlan plan;
  plan.generated_shared_strings = generated_shared_strings;
  plan.generated_persons = !wb.persons().empty();
  plan.tables_by_sheet.assign(wb.sheet_count(), {});

  // Distribute tables to their owning sheets, assigning fallback
  // numeric ids when the source `id` is 0 (which would collide with
  // every other id-less table).
  std::vector<EmissionPlan::PerSheetTable> flat_tables;
  std::unordered_set<std::uint32_t> used_ids;
  for (const TableMetadata& t : wb.tables()) {
    used_ids.insert(t.id);
  }
  std::uint32_t next_fallback_id = 1;
  for (const TableMetadata& t : wb.tables()) {
    if (t.sheet_index >= wb.sheet_count()) {
      // Defensive: stale metadata referencing a removed sheet. Skip
      // rather than crash; round-trip preserves what is consistent.
      // `write_ooxml_with_result` rejects this workbook with
      // `kIoWriteFailed` before it builds a plan, so a caller loses the
      // save rather than the table; the branch stays for any future
      // caller that builds a plan directly.
      StructuredLog("ooxml_writer.table_skipped")
          .field("reason", std::string_view("sheet_index_out_of_range"))
          .field("sheet_index", static_cast<std::int64_t>(t.sheet_index))
          .field("sheet_count", static_cast<std::int64_t>(wb.sheet_count()))
          .field("table_name", t.name)
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->deferred_feature_count;
      }
      continue;
    }
    EmissionPlan::PerSheetTable entry;
    entry.table = &t;
    entry.numeric_id = t.id;
    if (entry.numeric_id == 0) {
      // Find the first unused fallback id so generated filenames stay
      // unique across all tables in the package.
      while (used_ids.count(next_fallback_id) != 0U) {
        ++next_fallback_id;
      }
      entry.numeric_id = next_fallback_id;
      used_ids.insert(entry.numeric_id);
      ++next_fallback_id;
      StructuredLog("ooxml_writer.table_id_fallback")
          .field("table_name", t.name)
          .field("assigned_id", static_cast<std::int64_t>(entry.numeric_id))
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->renumbered_part_count;
      }
    }
    entry.path = NumberedPartPath("xl/tables/table", entry.numeric_id, ".xml");
    plan.tables_by_sheet[t.sheet_index].push_back(entry);
    flat_tables.push_back(entry);
  }

  // Pivot caches in document order. The workbook-rels rId integer for
  // each cache definition starts after the styles relationship: sheets
  // occupy rId1..rId(N), styles uses rId(N+1), pivot caches use
  // rId(N+2)+. The numeric_id drives the package-relative path.
  plan.pivot_caches.reserve(wb.pivot_caches().size());
  for (std::size_t i = 0; i < wb.pivot_caches().size(); ++i) {
    const pivot::PivotCache* cache = wb.pivot_caches()[i].get();
    if (cache == nullptr) {
      continue;
    }
    EmissionPlan::PivotCachePlan entry;
    entry.cache = cache;
    entry.numeric_id = static_cast<std::uint32_t>(i + 1);
    entry.cache_id = cache->cache_id();
    entry.definition_path = NumberedPartPath("xl/pivotCache/pivotCacheDefinition", entry.numeric_id, ".xml");
    entry.records_path = NumberedPartPath("xl/pivotCache/pivotCacheRecords", entry.numeric_id, ".xml");
    entry.definition_rels_path =
        NumberedPartPath("xl/pivotCache/_rels/pivotCacheDefinition", entry.numeric_id, ".xml.rels");
    // sheets rId1..rId(sheet_count), styles rId(sheet_count+1),
    // first cache rId(sheet_count+2). Cast safe: workbook size is
    // bounded well within uint32 range.
    entry.workbook_rid = static_cast<std::uint32_t>(wb.sheet_count() + 2 + (generated_shared_strings ? 1U : 0U) + i);
    plan.pivot_caches.push_back(std::move(entry));
  }

  // Pivot tables grouped by sheet, with a package-wide numeric counter.
  plan.pivot_tables_by_sheet.assign(wb.sheet_count(), {});
  std::uint32_t next_pivot_table_id = 1;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const auto& sheet_pivots = wb.sheet(s).pivot_tables();
    for (const std::unique_ptr<pivot::PivotTable>& uptr : sheet_pivots) {
      const pivot::PivotTable* tbl = uptr.get();
      if (tbl == nullptr) {
        continue;
      }
      EmissionPlan::PivotTablePlan entry;
      entry.table = tbl;
      entry.numeric_id = next_pivot_table_id++;
      entry.path = NumberedPartPath("xl/pivotTables/pivotTable", entry.numeric_id, ".xml");
      entry.rels_path = NumberedPartPath("xl/pivotTables/_rels/pivotTable", entry.numeric_id, ".xml.rels");
      // Resolve the cache this table draws from so its rels file can name
      // the definition part. Both live under `xl/`, so the target steps
      // out of `pivotTables/` and into `pivotCache/`. A table whose
      // cache id matches nothing leaves the target empty.
      for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
        if (c.cache_id == tbl->pivot_cache_id()) {
          entry.cache_definition_target = NumberedPartPath("../pivotCache/pivotCacheDefinition", c.numeric_id, ".xml");
          break;
        }
      }
      plan.pivot_tables_by_sheet[s].push_back(std::move(entry));
    }
  }

  // Comments / VML parts. One package-wide numeric counter; each sheet
  // with at least one note or thread gets a `comments<N>.xml` and a
  // matching `vmlDrawing<N>.vml` (a thread is also written as a legacy
  // stub note), plus `threadedComment<N>.xml` when it has threads. The
  // numeric id matches between them so the sheet rels file pairs them by
  // ordinal.
  //
  // A vmlDrawing target still named by some sheet's `unknown_relationships()`
  // (for example a `<legacyDrawingHF>` header/footer VML preserved
  // verbatim, see `sheet_aux_rels_reader.h`) already occupies a
  // `xl/drawings/vmlDrawing<N>.vml` path the counter would otherwise
  // assign. Skip any id that collides so the auto-numbered comment VML
  // never overwrites a part a live relationship still points at.
  //
  // This is deliberately narrower than "any passthrough part at that
  // path": an orphaned passthrough part left over from a since-removed
  // sheet's comment VML is not referenced by anything anymore, and
  // reusing its path for the next sheet's own renumbered comment VML is
  // the desired outcome, not a collision.
  std::unordered_set<std::string> retained_paths;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    for (const UnknownRelationship& rel : wb.sheet(s).unknown_relationships()) {
      if (rel.type == kRelVmlDrawing && !rel.target_external) {
        retained_paths.insert(rel.target);
      }
    }
  }
  plan.comments_by_sheet.assign(wb.sheet_count(), EmissionPlan::CommentsPlan{});
  std::uint32_t next_comments_id = 1;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    if (sheet.comments().empty() && sheet.threaded_comments().empty()) {
      continue;
    }
    while (retained_paths.count(NumberedPartPath("xl/drawings/vmlDrawing", next_comments_id, ".vml")) != 0U) {
      ++next_comments_id;
    }
    EmissionPlan::CommentsPlan entry;
    entry.numeric_id = next_comments_id++;
    entry.comments_path = NumberedPartPath("xl/comments", entry.numeric_id, ".xml");
    entry.vml_path = NumberedPartPath("xl/drawings/vmlDrawing", entry.numeric_id, ".vml");
    if (!sheet.threaded_comments().empty()) {
      entry.threaded_path = NumberedPartPath("xl/threadedComments/threadedComment", entry.numeric_id, ".xml");
    }
    // Detect whether this sheet still carries its original VML bytes. Do
    // not infer the source from a newly assigned output number: removing a
    // preceding commented sheet renumbers output parts but must not discard
    // the surviving sheet's shape geometry. Bytes whose shapes no longer
    // match the comment anchors are regenerated instead.
    if (!sheet.comment_vml_path().empty() && sheet.comment_vml_anchors() == sheet.comment_anchor_set()) {
      for (const PassthroughPart& part : wb.passthrough_parts()) {
        if (part.path == sheet.comment_vml_path()) {
          entry.vml_source = &part;
          break;
        }
      }
    }
    plan.comments_by_sheet[s] = std::move(entry);
  }

  // External link relationships. Assigned rIds follow the pivot caches
  // in the workbook-rels numbering scheme, mirroring how Excel emits
  // them when multiple optional sections coexist.
  PlanExternalLinks(
      wb,
      static_cast<std::size_t>(wb.sheet_count()) + 2U + (generated_shared_strings ? 1U : 0U) + plan.pivot_caches.size(),
      plan);
  std::unordered_set<std::string> regenerated_link_bodies;
  for (const EmissionPlan::ExternalLinkPlan& link : plan.external_links) {
    if (!link.generated_body.empty()) {
      regenerated_link_bodies.insert(link.part_path);
    }
  }

  // Collision detection between generated paths and passthrough paths.
  // Generated paths win; passthrough copy is dropped with a warning.
  std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, flat_tables, plan.pivot_caches, plan.pivot_tables_by_sheet, plan.comments_by_sheet,
                            plan.external_links, generated_shared_strings);

  for (const PassthroughPart& part : wb.passthrough_parts()) {
    // A stale external-link body is superseded by the one generated from
    // the model, which carries everything it did; nothing is lost.
    if (regenerated_link_bodies.count(part.path) != 0U) {
      continue;
    }
    if (generated.count(part.path) != 0U) {
      StructuredLog("ooxml_writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("generated_path_wins"))
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_part_count;
      }
      continue;
    }
    // A workbook read from `.xlsb` still carries that package's binary
    // parts (`xl/styles.bin`, `xl/calcChain.bin`, `xl/metadata.bin`,
    // `xl/worksheets/binaryIndex*.bin`, ...). None of them belongs in an
    // OOXML package: the modelled content is re-emitted as XML by the
    // generated parts, and every downstream consumer of
    // `passthrough_kept` — the `<Override>` list, workbook rels and
    // sheet rels — drops with them, so the output matches what saving
    // the same model from an `.xlsx` source produces.
    if (IsXlsbBinaryContentType(EffectiveContentType(wb, part))) {
      StructuredLog("ooxml_writer.passthrough_dropped")
          .field("path", part.path)
          .field("content_type", part.content_type)
          .field("reason", std::string_view("xlsb_binary_part"))
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_part_count;
      }
      continue;
    }
    plan.passthrough_kept.push_back(&part);
  }

  // Build each sheet's rels file once, after passthrough collision handling,
  // so unknown internal relationships can be retained only when their
  // payload is present in the finalized plan. `relationship_count` remains
  // the single source of truth for whether the writer emits the rels part;
  // opaque sheets keep a default-constructed (unused) entry.
  plan.sheet_rels.assign(wb.sheet_count(), SheetRelsResult{});
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).is_opaque_ooxml_sheet()) {
      continue;
    }
    plan.sheet_rels[i] = BuildSheetRels(wb.sheet(i), plan.tables_by_sheet[i], plan.pivot_tables_by_sheet[i],
                                        plan.comments_by_sheet[i], plan, diagnostics);
  }

  return plan;
}

}  // namespace io
}  // namespace formulon
