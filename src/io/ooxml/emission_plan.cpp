//
// Implementation of the OOXML writer's emission plan. Determines part
// paths, numeric ids, and resolves collisions between writer-generated
// and passthrough parts. No miniz state is touched here; this is pure
// metadata.

#include "io/ooxml/emission_plan.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/external_links.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/sheet_xml_builder.h"
#include "io/ooxml_defs.h"
#include "io/passthrough_part.h"
#include "io/tables_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
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
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part != nullptr && part->path == path) {
      return true;
    }
  }
  return false;
}

namespace {

/// Returns the set of paths the writer always generates, regardless of
/// metadata. Used to detect passthrough collisions.
std::unordered_set<std::string> BuildGeneratedPathSet(
    const Workbook& wb, const std::vector<EmissionPlan::PerSheetTable>& flat_tables,
    const std::vector<EmissionPlan::PivotCachePlan>& pivot_caches,
    const std::vector<std::vector<EmissionPlan::PivotTablePlan>>& pivot_tables_by_sheet,
    const std::vector<EmissionPlan::CommentsPlan>& comments_by_sheet, bool generated_shared_strings) {
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
  }
  // Per-link rels files for external links — the writer generates these
  // from the captured `ExternalLinkRecord`s; the body parts themselves
  // are passthrough.
  for (const ExternalLinkRecord& rec : wb.external_links()) {
    paths.insert(ooxml::rels_path_for_part(rec.part_path));
  }
  // Sheet rels: any sheet that owns at least one table or pivot table.
  // Computed by callers; we enumerate them here for completeness.
  return paths;
}

}  // namespace

EmissionPlan BuildEmissionPlan(const Workbook& wb, bool generated_shared_strings) {
  EmissionPlan plan;
  plan.generated_shared_strings = generated_shared_strings;
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
      StructuredLog("ooxml_writer.table_skipped")
          .field("reason", std::string_view("sheet_index_out_of_range"))
          .field("sheet_index", static_cast<std::int64_t>(t.sheet_index))
          .field("sheet_count", static_cast<std::int64_t>(wb.sheet_count()))
          .field("table_name", t.name)
          .warn();
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
  // with at least one comment gets a `comments<N>.xml` and a matching
  // `vmlDrawing<N>.vml`. The numeric id matches between the two so the
  // sheet rels file pairs them by ordinal.
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
    if (wb.sheet(s).comments().empty()) {
      continue;
    }
    while (retained_paths.count(NumberedPartPath("xl/drawings/vmlDrawing", next_comments_id, ".vml")) != 0U) {
      ++next_comments_id;
    }
    EmissionPlan::CommentsPlan entry;
    entry.numeric_id = next_comments_id++;
    entry.comments_path = NumberedPartPath("xl/comments", entry.numeric_id, ".xml");
    entry.vml_path = NumberedPartPath("xl/drawings/vmlDrawing", entry.numeric_id, ".vml");
    // Detect whether this sheet still carries its original VML bytes. Do
    // not infer the source from a newly assigned output number: removing a
    // preceding commented sheet renumbers output parts but must not discard
    // the surviving sheet's shape geometry.
    for (const PassthroughPart& part : wb.passthrough_parts()) {
      if (!wb.sheet(s).comment_vml_path().empty() && part.path == wb.sheet(s).comment_vml_path()) {
        entry.vml_source = &part;
        break;
      }
    }
    plan.comments_by_sheet[s] = std::move(entry);
  }

  // External link relationships. Assigned rIds follow the pivot caches
  // in the workbook-rels numbering scheme, mirroring how Excel emits
  // them when multiple optional sections coexist. The body parts ride
  // through `passthrough_parts()`; only the per-link rels files are
  // generated below.
  {
    const std::size_t base = static_cast<std::size_t>(wb.sheet_count()) + 2U + (generated_shared_strings ? 1U : 0U) +
                             plan.pivot_caches.size();
    for (std::size_t i = 0; i < wb.external_links().size(); ++i) {
      EmissionPlan::ExternalLinkPlan entry;
      entry.record = &wb.external_links()[i];
      entry.workbook_rid = static_cast<std::uint32_t>(base + i);
      plan.external_links.push_back(entry);
    }
  }

  // Build each sheet's rels file once, here, so every downstream
  // consumer — the collision-detection pass immediately below, and the
  // writer's `AddPart` call — reads the exact same result instead of
  // separately re-deriving whether a sheet has relationships worth
  // writing. `relationship_count` is the single source of truth for
  // that decision; opaque sheets keep a default-constructed (unused)
  // entry.
  plan.sheet_rels.assign(wb.sheet_count(), SheetRelsResult{});
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    if (wb.sheet(i).is_opaque_ooxml_sheet()) {
      continue;
    }
    plan.sheet_rels[i] =
        BuildSheetRels(wb.sheet(i), plan.tables_by_sheet[i], plan.pivot_tables_by_sheet[i], plan.comments_by_sheet[i]);
  }

  // Collision detection between generated paths and passthrough paths.
  // Generated paths win; passthrough copy is dropped with a warning.
  std::unordered_set<std::string> generated = BuildGeneratedPathSet(
      wb, flat_tables, plan.pivot_caches, plan.pivot_tables_by_sheet, plan.comments_by_sheet, generated_shared_strings);
  // A sheet's rels file is generated exactly when its built result
  // declares at least one relationship — an empty rels file is invalid
  // OOXML, so the writer never emits one (see `write_ooxml`'s use of
  // this same `plan.sheet_rels` entry).
  for (std::size_t i = 0; i < plan.sheet_rels.size(); ++i) {
    if (plan.sheet_rels[i].relationship_count > 0) {
      generated.insert("xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".xml.rels");
    }
  }

  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (generated.count(part.path) != 0U) {
      StructuredLog("ooxml_writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("generated_path_wins"))
          .warn();
      continue;
    }
    plan.passthrough_kept.push_back(&part);
  }

  return plan;
}

}  // namespace io
}  // namespace formulon
