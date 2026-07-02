// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include "io/ooxml/relationship_writer.h"
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
    const std::vector<EmissionPlan::CommentsPlan>& comments_by_sheet) {
  std::unordered_set<std::string> paths;
  paths.insert("[Content_Types].xml");
  paths.insert("_rels/.rels");
  paths.insert("xl/workbook.xml");
  paths.insert("xl/_rels/workbook.xml.rels");
  paths.insert("xl/styles.xml");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    paths.insert("xl/worksheets/sheet" + std::to_string(i + 1) + ".xml");
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
  // Pivot-table parts (one per pivot table, package-wide).
  for (const auto& per_sheet : pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      paths.insert(t.path);
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
    paths.insert(RelsPathForPart(rec.part_path));
  }
  // Sheet rels: any sheet that owns at least one table or pivot table.
  // Computed by callers; we enumerate them here for completeness.
  return paths;
}

}  // namespace

EmissionPlan BuildEmissionPlan(const Workbook& wb) {
  EmissionPlan plan;
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
    entry.workbook_rid = static_cast<std::uint32_t>(wb.sheet_count() + 2 + i);
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
      plan.pivot_tables_by_sheet[s].push_back(std::move(entry));
    }
  }

  // Comments / VML parts. One package-wide numeric counter; each sheet
  // with at least one comment gets a `comments<N>.xml` and a matching
  // `vmlDrawing<N>.vml`. The numeric id matches between the two so the
  // sheet rels file pairs them by ordinal.
  plan.comments_by_sheet.assign(wb.sheet_count(), EmissionPlan::CommentsPlan{});
  std::uint32_t next_comments_id = 1;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (wb.sheet(s).comments().empty()) {
      continue;
    }
    EmissionPlan::CommentsPlan entry;
    entry.numeric_id = next_comments_id++;
    entry.comments_path = NumberedPartPath("xl/comments", entry.numeric_id, ".xml");
    entry.vml_path = NumberedPartPath("xl/drawings/vmlDrawing", entry.numeric_id, ".vml");
    // Detect whether the workbook still carries the original VML bytes
    // via passthrough. If so, prefer those bytes over the writer's
    // stub so the round-trip stays byte-identical for unmodified
    // sheets.
    for (const PassthroughPart& part : wb.passthrough_parts()) {
      if (part.path == entry.vml_path) {
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
    const std::size_t base = static_cast<std::size_t>(wb.sheet_count()) + 2U + plan.pivot_caches.size();
    for (std::size_t i = 0; i < wb.external_links().size(); ++i) {
      EmissionPlan::ExternalLinkPlan entry;
      entry.record = &wb.external_links()[i];
      entry.workbook_rid = static_cast<std::uint32_t>(base + i);
      plan.external_links.push_back(entry);
    }
  }

  // Collision detection between generated paths and passthrough paths.
  // Generated paths win; passthrough copy is dropped with a warning.
  std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, flat_tables, plan.pivot_caches, plan.pivot_tables_by_sheet, plan.comments_by_sheet);
  // Sheet rels for sheets that own tables, pivot tables, hyperlinks,
  // comments/VML, or printer settings are also generated.
  for (std::size_t i = 0; i < plan.tables_by_sheet.size(); ++i) {
    const bool has_tables = !plan.tables_by_sheet[i].empty();
    const bool has_pivots = i < plan.pivot_tables_by_sheet.size() && !plan.pivot_tables_by_sheet[i].empty();
    const bool has_hyperlinks = i < wb.sheet_count() && !wb.sheet(i).hyperlinks().empty();
    const bool has_comments = i < plan.comments_by_sheet.size() && plan.comments_by_sheet[i].numeric_id != 0;
    const bool has_print_settings = i < wb.sheet_count() && !wb.sheet(i).print_settings().printer_settings_path.empty();
    const bool has_drawing = i < wb.sheet_count() && !wb.sheet(i).drawing_rel_target().empty();
    if (has_tables || has_pivots || has_hyperlinks || has_comments || has_print_settings || has_drawing) {
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
