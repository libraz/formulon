// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// OOXML (.xlsx) package writer. The writer emits the minimum spreadsheet
// surface that Excel 365 will open without complaint, and additionally
// round-trips the metadata Bundles 2.3 and 2.4 wired into the reader:
// defined names, table parts, and Override-listed parts the reader did
// not consume (unknown-part passthrough). Cells are still emitted with
// inline strings (`t="inlineStr"`); SST emission would force every text
// cell to walk a side table for no observable gain — the inline form
// round-trips cleanly already.

#include "io/ooxml_writer.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "eval/iterative_solver.h"
#include "io/cf_writer.h"
#include "io/comments_writer.h"
#include "io/defined_names.h"
#include "io/external_links.h"
#include "io/ooxml_writer_cell.h"
#include "io/passthrough_part.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/styles_writer.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "miniz.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

// `kXmlDecl` lives in `io/xml_utils.h` and is shared with comments / cf
// writers (single source of truth for the XML 1.0 prologue).

constexpr std::string_view kCtPackageRels = "application/vnd.openxmlformats-package.relationships+xml";
constexpr std::string_view kCtXml = "application/xml";
// The workbook part's content type depends on `Workbook::kind()`; we
// fetch it via `io::workbook_kind_content_type` at emission time rather
// than wiring four constants here.
constexpr std::string_view kCtWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
constexpr std::string_view kCtStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
constexpr std::string_view kCtTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
constexpr std::string_view kCtPivotCacheDefinition =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
constexpr std::string_view kCtPivotCacheRecords =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
constexpr std::string_view kCtPivotTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";

constexpr std::string_view kRelTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
constexpr std::string_view kRelPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
constexpr std::string_view kRelPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
constexpr std::string_view kRelPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
constexpr std::string_view kRelHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
constexpr std::string_view kRelComments =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
constexpr std::string_view kRelVmlDrawing =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";

constexpr std::string_view kCtComments = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
constexpr std::string_view kCtVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing";

// ---------------------------------------------------------------------------
// Per-package emission plan
// ---------------------------------------------------------------------------

/// Plan we build up before any miniz call: which sheets own which
/// table parts, what filename each table uses, what passthrough parts
/// we'll emit. Centralising this here keeps `BuildContentTypes` /
/// `AddPart` calls trivial and avoids re-deriving table numbering from
/// two places.
struct EmissionPlan {
  // For each sheet (by 0-based index), the in-source TableMetadata
  // entries that target it, paired with the package-relative path the
  // writer assigned (`xl/tables/tableN.xml`). `(table_ref, path)` is
  // append-only and 1:1 with `<tablePart>` rels.
  struct PerSheetTable {
    const TableMetadata* table = nullptr;
    std::string path;
    std::uint32_t numeric_id = 0;  // matches the path's `tableN.xml` suffix
  };
  std::vector<std::vector<PerSheetTable>> tables_by_sheet;
  // Workbook-level pivot caches. One entry per `wb.pivot_caches()` in
  // document order. `numeric_id` drives the package paths; `cache_id`
  // is the workbook-level identifier consumers see (PivotTable refers
  // to it via `pivot_cache_id()`); `workbook_rid` is the rId integer
  // the workbook-rels file assigns to this cache definition.
  struct PivotCachePlan {
    const pivot::PivotCache* cache = nullptr;
    std::uint32_t numeric_id = 0;      // 1-based, package-wide
    std::string definition_path;       // "xl/pivotCache/pivotCacheDefinition1.xml"
    std::string records_path;          // "xl/pivotCache/pivotCacheRecords1.xml"
    std::string definition_rels_path;  // "xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"
    std::uint32_t workbook_rid = 0;    // workbook-rels rId number for this cache definition
    std::uint32_t cache_id = 0;        // PivotCache::cache_id(); used by <pivotCaches> entry
  };
  std::vector<PivotCachePlan> pivot_caches;
  // Pivot tables grouped by owning sheet. Each `numeric_id` is a
  // package-wide 1-based counter (independent of cache numbering); the
  // path is `xl/pivotTables/pivotTable<N>.xml`.
  struct PivotTablePlan {
    const pivot::PivotTable* table = nullptr;
    std::uint32_t numeric_id = 0;  // 1-based, package-wide
    std::string path;              // "xl/pivotTables/pivotTable1.xml"
  };
  std::vector<std::vector<PivotTablePlan>> pivot_tables_by_sheet;
  // Per-sheet comments / VML payload. `numeric_id` matches the
  // `comments<N>.xml` filename (1-based, package-wide). `vml_path` is
  // the per-sheet VML drawing path, emitted as a stub when the source
  // had none and as passthrough bytes otherwise.
  struct CommentsPlan {
    std::uint32_t numeric_id = 0;
    std::string comments_path;                    // "xl/comments<N>.xml"
    std::string vml_path;                         // "xl/drawings/vmlDrawing<N>.vml"
    const PassthroughPart* vml_source = nullptr;  // non-null => use bytes verbatim
  };
  // One entry per sheet; engaged only for sheets with at least one
  // comment. Disengaged entries have `numeric_id == 0`.
  std::vector<CommentsPlan> comments_by_sheet;
  // External link relationships. One entry per `wb.external_links()` in
  // document order. `workbook_rid` is the writer-assigned rId integer
  // for the workbook.xml.rels Relationship and the matching
  // `<externalReference r:id>` attribute in workbook.xml. The body part
  // itself rides through `passthrough_parts()`; the per-link rels file
  // (under `xl/externalLinks/_rels/`) is generated by the writer from
  // the captured target URL.
  struct ExternalLinkPlan {
    const ExternalLinkRecord* record = nullptr;
    std::uint32_t workbook_rid = 0;
  };
  std::vector<ExternalLinkPlan> external_links;
  // Passthrough parts we will keep. Entries that collide with a
  // generated path are dropped here (with a warning) so downstream
  // emission can blindly write everything in the list.
  std::vector<const PassthroughPart*> passthrough_kept;
};

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
    // SheetRelsPath equivalent inlined here to avoid a writer-side
    // dependency on the reader's helpers: insert
    // `xl/<dir>/_rels/<file>.rels` for `xl/<dir>/<file>`.
    const std::size_t slash = rec.part_path.find_last_of('/');
    std::string rels_path;
    if (slash == std::string::npos) {
      rels_path = "_rels/" + rec.part_path + ".rels";
    } else {
      rels_path.append(rec.part_path.substr(0, slash));
      rels_path.append("/_rels/");
      rels_path.append(rec.part_path.substr(slash + 1));
      rels_path.append(".rels");
    }
    paths.insert(std::move(rels_path));
  }
  // Sheet rels: any sheet that owns at least one table or pivot table.
  // Computed by callers; we enumerate them here for completeness.
  return paths;
}

/// Builds the emission plan. Performs collision detection between
/// generated and passthrough paths; collisions are logged via
/// `StructuredLog` (warn) and the passthrough copy is dropped.
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
    entry.path = "xl/tables/table" + std::to_string(entry.numeric_id) + ".xml";
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
    const std::string n_str = std::to_string(entry.numeric_id);
    entry.definition_path = "xl/pivotCache/pivotCacheDefinition" + n_str + ".xml";
    entry.records_path = "xl/pivotCache/pivotCacheRecords" + n_str + ".xml";
    entry.definition_rels_path = "xl/pivotCache/_rels/pivotCacheDefinition" + n_str + ".xml.rels";
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
      entry.path = "xl/pivotTables/pivotTable" + std::to_string(entry.numeric_id) + ".xml";
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
    entry.comments_path = "xl/comments" + std::to_string(entry.numeric_id) + ".xml";
    entry.vml_path = "xl/drawings/vmlDrawing" + std::to_string(entry.numeric_id) + ".vml";
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
  // comments, or that need a VML drawing rel are also generated.
  for (std::size_t i = 0; i < plan.tables_by_sheet.size(); ++i) {
    const bool has_tables = !plan.tables_by_sheet[i].empty();
    const bool has_pivots = i < plan.pivot_tables_by_sheet.size() && !plan.pivot_tables_by_sheet[i].empty();
    const bool has_hyperlinks = i < wb.sheet_count() && !wb.sheet(i).hyperlinks().empty();
    const bool has_comments = i < plan.comments_by_sheet.size() && plan.comments_by_sheet[i].numeric_id != 0;
    if (has_tables || has_pivots || has_hyperlinks || has_comments) {
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

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------

/// Escapes `text` and appends it as the body of an XML element. Callers
/// that need attribute escaping should use `AppendXmlEscaped` directly.
inline void AppendEscaped(std::string& out, std::string_view text) {
  AppendXmlEscaped(out, text);
}

/// Appends a single `<Override PartName="/<path>" ContentType="<ct>"/>`
/// entry plus its trailing newline. Used by `BuildContentTypes` for the
/// per-table / per-pivot / per-comments / passthrough Override blocks
/// — each emitting bytes-identical fragments before this helper landed.
/// `path` is escaped to defend against passthrough paths carrying
/// XML-critical characters; `ct` is a writer-controlled string view from
/// our content-type table and is emitted verbatim.
inline void AppendOverride(std::string& out, std::string_view path, std::string_view ct, bool escape_path = false) {
  out.append("  <Override PartName=\"/");
  if (escape_path) {
    AppendXmlEscaped(out, path);
  } else {
    out.append(path.data(), path.size());
  }
  out.append("\" ContentType=\"");
  out.append(ct.data(), ct.size());
  out.append("\"/>\n");
}

// ---------------------------------------------------------------------------
// Part builders
// ---------------------------------------------------------------------------

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 128 + plan.passthrough_kept.size() * 128);
  out.append(kXmlDecl);
  out.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n");
  out.append("  <Default Extension=\"rels\" ContentType=\"");
  out.append(kCtPackageRels);
  out.append("\"/>\n");
  out.append("  <Default Extension=\"xml\" ContentType=\"");
  out.append(kCtXml);
  out.append("\"/>\n");
  // VML drawings are referenced by extension via a Default; this lets
  // the per-sheet VML stub avoid an Override entry. Emitted only when
  // at least one sheet has comments (i.e. a VML companion is needed).
  bool any_comments = false;
  for (const EmissionPlan::CommentsPlan& cplan : plan.comments_by_sheet) {
    if (cplan.numeric_id != 0) {
      any_comments = true;
      break;
    }
  }
  if (any_comments) {
    out.append("  <Default Extension=\"vml\" ContentType=\"");
    out.append(kCtVmlDrawing);
    out.append("\"/>\n");
  }
  AppendOverride(out, "xl/workbook.xml", workbook_kind_content_type(wb.kind()));
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    const std::string sheet_path = "xl/worksheets/sheet" + std::to_string(i + 1) + ".xml";
    AppendOverride(out, sheet_path, kCtWorksheet);
  }
  AppendOverride(out, "xl/styles.xml", kCtStyles);
  // Per-table overrides (one per emitted table part, regardless of
  // owning sheet).
  for (const auto& per_sheet : plan.tables_by_sheet) {
    for (const EmissionPlan::PerSheetTable& t : per_sheet) {
      AppendOverride(out, t.path, kCtTable);
    }
  }
  // Pivot caches: one Override per definition + one per records part.
  for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
    AppendOverride(out, c.definition_path, kCtPivotCacheDefinition);
    AppendOverride(out, c.records_path, kCtPivotCacheRecords);
  }
  // Pivot tables: one Override per part.
  for (const auto& per_sheet : plan.pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      AppendOverride(out, t.path, kCtPivotTable);
    }
  }
  // Comments parts: one Override per part. VML drawings are covered
  // by the `Default Extension="vml"` entry above and need no Override.
  for (const EmissionPlan::CommentsPlan& cplan : plan.comments_by_sheet) {
    if (cplan.numeric_id == 0) {
      continue;
    }
    AppendOverride(out, cplan.comments_path, kCtComments);
  }
  // Passthrough overrides: only for entries that carried an explicit
  // ContentType in the source archive. Default-typed parts (empty
  // content_type) must NOT appear as Overrides — the package's
  // `<Default Extension=...>` entries already cover them.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part->content_type.empty()) {
      continue;
    }
    // Passthrough payloads may carry XML-critical bytes in either path
    // or content type (rare but legal); escape both.
    out.append("  <Override PartName=\"/");
    AppendXmlEscaped(out, part->path);
    out.append("\" ContentType=\"");
    AppendXmlEscaped(out, part->content_type);
    out.append("\"/>\n");
  }
  out.append("</Types>\n");
  return out;
}

std::string BuildPackageRels() {
  std::string out;
  out.reserve(256);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append(
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n");
  out.append("</Relationships>\n");
  return out;
}

void AppendDefinedNamesBlock(std::string& out, const std::vector<DefinedName>& names) {
  if (names.empty()) {
    return;
  }
  out.append("  <definedNames>\n");
  for (const DefinedName& n : names) {
    out.append("    <definedName name=\"");
    AppendXmlEscaped(out, n.name);
    out.push_back('"');
    if (n.local_sheet_id >= 0) {
      out.append(" localSheetId=\"");
      out.append(std::to_string(n.local_sheet_id));
      out.push_back('"');
    }
    if (n.hidden) {
      out.append(" hidden=\"1\"");
    }
    if (!n.comment.empty()) {
      out.append(" comment=\"");
      AppendXmlEscaped(out, n.comment);
      out.push_back('"');
    }
    out.push_back('>');
    AppendEscaped(out, n.formula);
    out.append("</definedName>\n");
  }
  out.append("  </definedNames>\n");
}

std::string BuildWorkbookXml(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 96 + wb.defined_names().size() * 96 + plan.pivot_caches.size() * 64);
  out.append(kXmlDecl);
  out.append(
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  out.append("  <sheets>\n");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out.append("    <sheet name=\"");
    AppendXmlEscaped(out, wb.sheet(i).name());
    out.append("\" sheetId=\"");
    out.append(std::to_string(i + 1));
    out.append("\" r:id=\"rId");
    out.append(std::to_string(i + 1));
    out.push_back('"');
    // Excel records tab visibility on the workbook part via the
    // `state` attribute. We OR-merge that with `<sheetPr><tabHidden/>`
    // on the worksheet part inside `BuildWorksheetXml`, so either
    // path round-trips correctly. Keeping the workbook-side emission
    // here mirrors what Excel writes by default.
    if (wb.sheet(i).view().tab_hidden) {
      out.append(" state=\"hidden\"");
    }
    out.append("/>\n");
  }
  out.append("  </sheets>\n");
  // <externalReferences> precedes <definedNames> per ECMA-376 element
  // order (sheets, functionGroups, externalReferences, definedNames,
  // ...). Emit only when the workbook actually carries cross-workbook
  // references so freshly-created files stay diff-friendly.
  if (!plan.external_links.empty()) {
    out.append("  <externalReferences>\n");
    for (const EmissionPlan::ExternalLinkPlan& e : plan.external_links) {
      out.append("    <externalReference r:id=\"rId");
      out.append(std::to_string(e.workbook_rid));
      out.append("\"/>\n");
    }
    out.append("  </externalReferences>\n");
  }
  // <definedNames> sits between <sheets> and <calcPr>/end-of-workbook
  // per OOXML schema (cf. ECMA-376 sheet ordering).
  AppendDefinedNamesBlock(out, wb.defined_names());
  // <calcPr> persists Excel's workbook-level calculation policy
  // (`calcMode` + iterative trio). Emit only when at least one
  // attribute differs from the spec defaults so a fresh workbook keeps
  // emitting a minimal, diff-friendly part. Element order per ECMA-376:
  // sheets, functionGroups, externalReferences, definedNames, calcPr,
  // oleSize, customWorkbookViews, pivotCaches, ...
  {
    const Workbook::CalcMode calc_mode = wb.calc_mode();
    const eval::IterativeOptions& iter = wb.iterative_options();
    const bool calc_mode_default = calc_mode == Workbook::CalcMode::kAuto;
    const bool iterate_default = !iter.enabled && iter.max_iterations == eval::kDefaultMaxIterations &&
                                 iter.max_change == eval::kDefaultMaxChange;
    if (!calc_mode_default || !iterate_default) {
      out.append("  <calcPr");
      if (calc_mode == Workbook::CalcMode::kManual) {
        out.append(" calcMode=\"manual\"");
      } else if (calc_mode == Workbook::CalcMode::kAutoNoTable) {
        out.append(" calcMode=\"autoNoTable\"");
      }
      if (iter.enabled) {
        out.append(" iterate=\"1\"");
      }
      if (iter.max_iterations != eval::kDefaultMaxIterations) {
        out.append(" iterateCount=\"");
        out.append(std::to_string(iter.max_iterations));
        out.push_back('"');
      }
      if (iter.max_change != eval::kDefaultMaxChange) {
        out.append(" iterateDelta=\"");
        format_double(out, iter.max_change);
        out.push_back('"');
      }
      out.append("/>\n");
    }
  }
  // <pivotCaches> follows <calcPr> in the ECMA-376 schema element order.
  if (!plan.pivot_caches.empty()) {
    out.append("  <pivotCaches>\n");
    for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
      out.append("    <pivotCache cacheId=\"");
      out.append(std::to_string(c.cache_id));
      out.append("\" r:id=\"rId");
      out.append(std::to_string(c.workbook_rid));
      out.append("\"/>\n");
    }
    out.append("  </pivotCaches>\n");
  }
  out.append("</workbook>\n");
  return out;
}

std::string BuildWorkbookRels(std::size_t sheet_count, const EmissionPlan& plan) {
  std::string out;
  out.reserve(256 + sheet_count * 192 + plan.pivot_caches.size() * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  for (std::size_t i = 0; i < sheet_count; ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(i + 1));
    out.append(
        "\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
        "Target=\"worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".xml\"/>\n");
  }
  // Styles relationship follows the worksheet relationships.
  out.append("  <Relationship Id=\"rId");
  out.append(std::to_string(sheet_count + 1));
  out.append(
      "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" "
      "Target=\"styles.xml\"/>\n");
  // Pivot-cache definition relationships, one per planned cache. Targets
  // are relative to the workbook directory (`xl/`); we strip the `xl/`
  // prefix from `definition_path` so the form matches what Excel emits
  // (e.g. `Target="pivotCache/pivotCacheDefinition1.xml"`).
  constexpr std::string_view kXlPrefix = "xl/";
  for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
    std::string_view target = c.definition_path;
    if (target.size() >= kXlPrefix.size() && target.substr(0, kXlPrefix.size()) == kXlPrefix) {
      target.remove_prefix(kXlPrefix.size());
    }
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(c.workbook_rid));
    out.append("\" Type=\"");
    out.append(kRelPivotCacheDefinition);
    out.append("\" Target=\"");
    out.append(target);
    out.append("\"/>\n");
  }
  // External link relationships. Same `xl/` prefix stripping as pivot
  // caches above; targets land as `Target="externalLinks/externalLink1.xml"`.
  for (const EmissionPlan::ExternalLinkPlan& e : plan.external_links) {
    std::string_view target = e.record->part_path;
    if (target.size() >= kXlPrefix.size() && target.substr(0, kXlPrefix.size()) == kXlPrefix) {
      target.remove_prefix(kXlPrefix.size());
    }
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(e.workbook_rid));
    out.append(
        "\" "
        "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\" "
        "Target=\"");
    AppendXmlEscaped(out, target);
    out.append("\"/>\n");
  }
  out.append("</Relationships>\n");
  return out;
}

/// Builds the per-link rels file content for one external link.
/// Returns an empty string when the record has no captured target —
/// callers should skip the AddPart call in that case so the package
/// does not carry an empty rels file Excel would treat as malformed.
std::string BuildExternalLinkRels(const ExternalLinkRecord& rec) {
  if (rec.target.empty()) {
    return {};
  }
  std::string out;
  out.reserve(256);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append("  <Relationship Id=\"");
  AppendXmlEscaped(out, rec.body_rel_id.empty() ? std::string("rId1") : rec.body_rel_id);
  out.append("\" Type=\"");
  switch (rec.kind) {
    case ExternalLinkRecord::Kind::kOleLink:
      out.append("http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleLink");
      break;
    case ExternalLinkRecord::Kind::kDdeLink:
      out.append("http://schemas.openxmlformats.org/officeDocument/2006/relationships/ddeLink");
      break;
    case ExternalLinkRecord::Kind::kExternalBook:
    case ExternalLinkRecord::Kind::kUnknown:
    default:
      out.append("http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath");
      break;
  }
  out.append("\" Target=\"");
  AppendXmlEscaped(out, rec.target);
  if (rec.target_external) {
    out.append("\" TargetMode=\"External\"/>\n");
  } else {
    out.append("\"/>\n");
  }
  out.append("</Relationships>\n");
  return out;
}

// Forward declarations for sheet-view / column-layout builders. The
// definitions live at the bottom of the file alongside the other XML
// helpers; their signatures are needed up here so `BuildWorksheetXml`
// can call them.
std::string BuildSheetViewXml(const SheetView& view);
std::string BuildColsXml(const SheetLayout& layout);
std::string BuildSheetProtectionXml(const SheetProtection& p);

/// Bijective base-26 column letters (`0 -> A`, `25 -> Z`, `26 -> AA`,
/// ...). Mirrors `cli/render.cpp`'s helper; kept inline here so the
/// writer side has no cross-package dependency.
void AppendColumnLettersForRef(std::string& out, std::uint32_t col) {
  char buf[4];
  std::uint32_t i = 0;
  std::uint32_t v = col + 1;
  while (v > 0 && i < 4) {
    const std::uint32_t rem = (v - 1) % 26U;
    buf[i++] = static_cast<char>('A' + rem);
    v = (v - 1) / 26U;
  }
  while (i > 0) {
    out.push_back(buf[--i]);
  }
}

void AppendCellRefForRef(std::string& out, std::uint32_t row, std::uint32_t col) {
  AppendColumnLettersForRef(out, col);
  out.append(std::to_string(row + 1));
}

void AppendRangeRef(std::string& out, const MergeRange& r) {
  AppendCellRefForRef(out, r.first_row, r.first_col);
  if (r.first_row != r.last_row || r.first_col != r.last_col) {
    out.push_back(':');
    AppendCellRefForRef(out, r.last_row, r.last_col);
  }
}

std::string BuildMergeCellsBlock(const Sheet& sheet) {
  if (sheet.merges().empty()) {
    return {};
  }
  std::string out;
  out.reserve(64 + sheet.merges().size() * 32);
  out.append("<mergeCells count=\"");
  out.append(std::to_string(sheet.merges().size()));
  out.append("\">");
  for (const MergeRange& m : sheet.merges()) {
    out.append("<mergeCell ref=\"");
    AppendRangeRef(out, m);
    out.append("\"/>");
  }
  out.append("</mergeCells>");
  return out;
}

std::string_view DataValidationTypeToString(std::uint8_t type) {
  switch (type) {
    case 1:
      return "whole";
    case 2:
      return "decimal";
    case 3:
      return "list";
    case 4:
      return "date";
    case 5:
      return "time";
    case 6:
      return "textLength";
    case 7:
      return "custom";
    default:
      return "";
  }
}

std::string_view DataValidationOperatorToString(std::uint8_t op) {
  switch (op) {
    case 1:
      return "notBetween";
    case 2:
      return "equal";
    case 3:
      return "notEqual";
    case 4:
      return "greaterThan";
    case 5:
      return "lessThan";
    case 6:
      return "greaterThanOrEqual";
    case 7:
      return "lessThanOrEqual";
    default:
      return "";  // 0 == between, omitted
  }
}

std::string_view DataValidationErrorStyleToString(std::uint8_t style) {
  switch (style) {
    case 1:
      return "warning";
    case 2:
      return "information";
    default:
      return "";  // 0 == stop (default)
  }
}

std::string BuildDataValidationsBlock(const Sheet& sheet) {
  if (sheet.validations().empty()) {
    return {};
  }
  std::string out;
  out.reserve(96 + sheet.validations().size() * 96);
  out.append("<dataValidations count=\"");
  out.append(std::to_string(sheet.validations().size()));
  out.append("\">");
  for (const DataValidation& v : sheet.validations()) {
    out.append("<dataValidation");
    if (const std::string_view t = DataValidationTypeToString(v.type); !t.empty()) {
      out.append(" type=\"");
      out.append(t);
      out.append("\"");
    }
    if (const std::string_view op = DataValidationOperatorToString(v.op); !op.empty()) {
      out.append(" operator=\"");
      out.append(op);
      out.append("\"");
    }
    if (const std::string_view es = DataValidationErrorStyleToString(v.error_style); !es.empty()) {
      out.append(" errorStyle=\"");
      out.append(es);
      out.append("\"");
    }
    if (!v.allow_blank) {
      out.append(" allowBlank=\"0\"");
    } else {
      out.append(" allowBlank=\"1\"");
    }
    if (v.show_input_message) {
      out.append(" showInputMessage=\"1\"");
    }
    if (v.show_error_message) {
      out.append(" showErrorMessage=\"1\"");
    }
    if (!v.error_title.empty()) {
      out.append(" errorTitle=\"");
      AppendXmlEscaped(out, v.error_title);
      out.append("\"");
    }
    if (!v.error_message.empty()) {
      out.append(" error=\"");
      AppendXmlEscaped(out, v.error_message);
      out.append("\"");
    }
    if (!v.prompt_title.empty()) {
      out.append(" promptTitle=\"");
      AppendXmlEscaped(out, v.prompt_title);
      out.append("\"");
    }
    if (!v.prompt_message.empty()) {
      out.append(" prompt=\"");
      AppendXmlEscaped(out, v.prompt_message);
      out.append("\"");
    }
    out.append(" sqref=\"");
    for (std::size_t i = 0; i < v.ranges.size(); ++i) {
      if (i > 0) {
        out.push_back(' ');
      }
      AppendRangeRef(out, v.ranges[i]);
    }
    out.append("\">");
    if (!v.formula1.empty()) {
      out.append("<formula1>");
      AppendXmlEscaped(out, v.formula1);
      out.append("</formula1>");
    }
    if (!v.formula2.empty()) {
      out.append("<formula2>");
      AppendXmlEscaped(out, v.formula2);
      out.append("</formula2>");
    }
    out.append("</dataValidation>");
  }
  out.append("</dataValidations>");
  return out;
}

/// Builds the `<hyperlinks>` block. `rid_for_index` returns the rels-file
/// rId integer assigned to the hyperlink at index `i`, or 0 when no rId
/// was assigned (defensive — every external hyperlink gets one).
std::string BuildHyperlinksBlock(const Sheet& sheet, const std::vector<std::string>& rid_per_hyperlink) {
  if (sheet.hyperlinks().empty()) {
    return {};
  }
  std::string out;
  out.reserve(48 + sheet.hyperlinks().size() * 96);
  out.append("<hyperlinks>");
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& h = sheet.hyperlinks()[i];
    out.append("<hyperlink ref=\"");
    AppendCellRefForRef(out, h.row, h.col);
    out.append("\"");
    if (i < rid_per_hyperlink.size() && !rid_per_hyperlink[i].empty()) {
      out.append(" r:id=\"");
      AppendXmlEscaped(out, rid_per_hyperlink[i]);
      out.append("\"");
    }
    if (!h.location.empty()) {
      out.append(" location=\"");
      AppendXmlEscaped(out, h.location);
      out.append("\"");
    }
    if (!h.tooltip.empty()) {
      out.append(" tooltip=\"");
      AppendXmlEscaped(out, h.tooltip);
      out.append("\"");
    }
    if (!h.display.empty()) {
      out.append(" display=\"");
      AppendXmlEscaped(out, h.display);
      out.append("\"");
    }
    out.append("/>");
  }
  out.append("</hyperlinks>");
  return out;
}

std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& hyperlink_rids) {
  const std::string sheet_view_xml = BuildSheetViewXml(sheet.view());
  const std::string cols_xml = BuildColsXml(sheet.layout());
  const std::string sheet_data = BuildSheetDataXml(sheet);
  // Conditional-format blocks live between <sheetData> and <tableParts>
  // in ECMA-376 document order. Empty list => empty string, no
  // wrapper.
  const std::string cf_xml = write_conditional_formattings(sheet.conditional_formats());
  const std::string merges_xml = BuildMergeCellsBlock(sheet);
  const std::string dv_xml = BuildDataValidationsBlock(sheet);
  const std::string hl_xml = BuildHyperlinksBlock(sheet, hyperlink_rids);
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + cols_xml.size() + sheet_data.size() + cf_xml.size() + merges_xml.size() +
              dv_xml.size() + hl_xml.size() + sheet_tables.size() * 96);
  out.append(kXmlDecl);
  out.append(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  // ECMA-376 element order: sheetPr -> dimension -> sheetViews ->
  // sheetFormatPr -> cols -> sheetData -> conditionalFormatting ->
  // tableParts. We currently emit a subset; the helpers stay quiet
  // when their underlying field is at default values so absent
  // metadata yields no extra bytes.
  if (!sheet_view_xml.empty()) {
    out.append("  ");
    out.append(sheet_view_xml);
    out.push_back('\n');
  }
  if (!cols_xml.empty()) {
    out.append("  ");
    out.append(cols_xml);
    out.push_back('\n');
  }
  out.append("  ");
  out.append(sheet_data);
  out.push_back('\n');
  // <sheetProtection> sits between <sheetData> and <mergeCells> per
  // ECMA-376 document order. Helper returns "" when protection is
  // disabled, leaving no trailing whitespace in that case.
  {
    const std::string sp_xml = BuildSheetProtectionXml(sheet.protection());
    if (!sp_xml.empty()) {
      out.append("  ");
      out.append(sp_xml);
      out.push_back('\n');
    }
  }
  // Merge cells precede CF in ECMA-376 document order.
  if (!merges_xml.empty()) {
    out.append("  ");
    out.append(merges_xml);
    out.push_back('\n');
  }
  if (!cf_xml.empty()) {
    out.append("  ");
    out.append(cf_xml);
    out.push_back('\n');
  }
  if (!dv_xml.empty()) {
    out.append("  ");
    out.append(dv_xml);
    out.push_back('\n');
  }
  if (!hl_xml.empty()) {
    out.append("  ");
    out.append(hl_xml);
    out.push_back('\n');
  }
  if (!sheet_tables.empty()) {
    out.append("  <tableParts count=\"");
    out.append(std::to_string(sheet_tables.size()));
    out.append("\">\n");
    for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
      out.append("    <tablePart r:id=\"rId");
      out.append(std::to_string(i + 1));
      out.append("\"/>\n");
    }
    out.append("  </tableParts>\n");
  }
  out.append("</worksheet>\n");
  return out;
}

/// Returned per-sheet rels result: the serialised XML plus the rId
/// strings (`"rIdN"`) the writer assigned to each hyperlink in
/// document order. The caller threads the rId vector into the sheet
/// part's `<hyperlinks>` block so the two stay in sync.
struct SheetRelsResult {
  std::string xml;
  std::vector<std::string> hyperlink_rids;
};

SheetRelsResult BuildSheetRels(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                               const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables,
                               const EmissionPlan::CommentsPlan& comments_plan) {
  SheetRelsResult res;
  std::string& out = res.xml;
  out.reserve(256 + (sheet_tables.size() + sheet_pivot_tables.size() + sheet.hyperlinks().size()) * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::uint32_t next_rid = 1;
  for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(next_rid++));
    out.append("\" Type=\"");
    out.append(kRelTable);
    out.append("\" Target=\"../tables/table");
    out.append(std::to_string(sheet_tables[i].numeric_id));
    out.append(".xml\"/>\n");
  }
  // Pivot-table relationships follow the table relationships, with rId
  // numbering continuing in sequence so each rel id is unique within
  // the sheet rels file.
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(next_rid++));
    out.append("\" Type=\"");
    out.append(kRelPivotTable);
    out.append("\" Target=\"../pivotTables/pivotTable");
    out.append(std::to_string(sheet_pivot_tables[i].numeric_id));
    out.append(".xml\"/>\n");
  }
  // Hyperlink relationships. Each external hyperlink target gets one
  // entry. Relative ordering is preserved so the round-trip writes the
  // rIds in the same order the reader observed.
  res.hyperlink_rids.reserve(sheet.hyperlinks().size());
  for (const Hyperlink& h : sheet.hyperlinks()) {
    if (h.target.empty()) {
      // Pure internal links carry their target through the inline
      // `location=` attribute instead of a rels entry.
      res.hyperlink_rids.emplace_back();
      continue;
    }
    // Reuse the source `rid` when present so byte-identical round-trips
    // are possible; otherwise mint a fresh rIdN counter.
    std::string rid;
    if (!h.rid.empty()) {
      rid = h.rid;
    } else {
      rid = "rId" + std::to_string(next_rid++);
    }
    res.hyperlink_rids.push_back(rid);
    out.append("  <Relationship Id=\"");
    AppendXmlEscaped(out, rid);
    out.append("\" Type=\"");
    out.append(kRelHyperlink);
    out.append("\" Target=\"");
    AppendXmlEscaped(out, h.target);
    out.append("\" TargetMode=\"External\"/>\n");
  }
  // Comments + VML relationships when the sheet has comments. The
  // comments rel comes first; the VML rel follows so the two ids are
  // adjacent and readers that scan top-down see them as a pair.
  if (comments_plan.numeric_id != 0) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(next_rid++));
    out.append("\" Type=\"");
    out.append(kRelComments);
    out.append("\" Target=\"../comments");
    out.append(std::to_string(comments_plan.numeric_id));
    out.append(".xml\"/>\n");
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(next_rid++));
    out.append("\" Type=\"");
    out.append(kRelVmlDrawing);
    out.append("\" Target=\"../drawings/vmlDrawing");
    out.append(std::to_string(comments_plan.numeric_id));
    out.append(".vml\"/>\n");
  }
  out.append("</Relationships>\n");
  return res;
}

/// Emits the `_rels` document for a pivotCacheDefinition part: a single
/// relationship of type `pivotCacheRecords` pointing at the matching
/// records part. The records target lives in the same directory as the
/// definition, so the `Target` is just the basename (e.g.
/// `"pivotCacheRecords1.xml"`).
std::string BuildPivotCacheDefinitionRels(std::string_view records_filename) {
  std::string out;
  out.reserve(256 + records_filename.size());
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append("  <Relationship Id=\"rId1\" Type=\"");
  out.append(kRelPivotCacheRecords);
  out.append("\" Target=\"");
  out.append(records_filename);
  out.append("\"/>\n");
  out.append("</Relationships>\n");
  return out;
}

std::string BuildTableXml(const TableMetadata& t, std::uint32_t numeric_id) {
  std::string out;
  out.reserve(256 + t.columns.size() * 64);
  out.append(kXmlDecl);
  out.append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
  out.append(std::to_string(numeric_id));
  out.append("\" name=\"");
  AppendXmlEscaped(out, t.name);
  out.append("\" displayName=\"");
  AppendXmlEscaped(out, t.display_name);
  out.append("\" ref=\"");
  AppendXmlEscaped(out, t.ref);
  out.push_back('"');
  // headerRowCount: emit explicit "0" when disabled; default (1) is
  // implicit via OOXML schema and we omit it.
  if (!t.header_row) {
    out.append(" headerRowCount=\"0\"");
  }
  if (t.totals_row) {
    out.append(" totalsRowCount=\"1\"");
  }
  out.append(">\n");
  out.append("  <tableColumns count=\"");
  out.append(std::to_string(t.columns.size()));
  out.append("\">\n");
  for (const TableColumn& col : t.columns) {
    out.append("    <tableColumn id=\"");
    out.append(std::to_string(col.id));
    out.append("\" name=\"");
    AppendXmlEscaped(out, col.name);
    out.push_back('"');
    if (!col.totals_label.empty()) {
      out.append(" totalsRowLabel=\"");
      AppendXmlEscaped(out, col.totals_label);
      out.push_back('"');
    }
    if (!col.totals_function.empty()) {
      out.append(" totalsRowFunction=\"");
      AppendXmlEscaped(out, col.totals_function);
      out.push_back('"');
    }
    // <calculatedColumnFormula> is the only child of <tableColumn> we
    // emit. When the field is empty we omit the element entirely (Excel
    // never emits empty calc-column elements) and keep the self-closing
    // <tableColumn/> form.
    if (col.calculated_column_formula.empty()) {
      out.append("/>\n");
    } else {
      out.append(">\n");
      out.append("      <calculatedColumnFormula>");
      AppendXmlEscaped(out, col.calculated_column_formula);
      out.append("</calculatedColumnFormula>\n");
      out.append("    </tableColumn>\n");
    }
  }
  out.append("  </tableColumns>\n");
  out.append("</table>\n");
  return out;
}

// ---------------------------------------------------------------------------
// miniz helpers
// ---------------------------------------------------------------------------

/// RAII guard around an initialised `mz_zip_archive` writer. The destructor
/// releases any heap buffer retained by miniz when the writer is abandoned
/// mid-flight (e.g. an `mz_zip_writer_add_mem` call failed and we early-
/// returned an error).
class ZipWriterGuard {
 public:
  ZipWriterGuard() = default;
  ZipWriterGuard(const ZipWriterGuard&) = delete;
  ZipWriterGuard& operator=(const ZipWriterGuard&) = delete;
  ZipWriterGuard(ZipWriterGuard&&) = delete;
  ZipWriterGuard& operator=(ZipWriterGuard&&) = delete;

  ~ZipWriterGuard() {
    if (active_) {
      // Best-effort cleanup; we're already on an error path.
      mz_zip_writer_end(&archive_);
    }
  }

  bool init() {
    if (mz_zip_writer_init_heap(&archive_, /*size_to_reserve_at_beginning=*/0,
                                /*initial_allocation_size=*/8 * 1024) == MZ_FALSE) {
      return false;
    }
    active_ = true;
    return true;
  }

  mz_zip_archive* get() noexcept { return &archive_; }

  /// Releases ownership of the underlying archive to the caller. Subsequent
  /// destruction no longer touches miniz state.
  void release() noexcept { active_ = false; }

 private:
  mz_zip_archive archive_{};
  bool active_ = false;
};

/// Adds a single text part to the archive. Returns an `Error` tagged
/// with the part path when miniz refuses the write.
Expected<void, Error> AddPart(mz_zip_archive* archive, std::string_view path, const std::string& body) {
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed", std::move(context));
  }
  return Expected<void, Error>::Ok();
}

/// Adds a binary part (passthrough). Same error contract as `AddPart`.
Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body) {
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed (passthrough)",
                      std::move(context));
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// View / layout part builders
// ---------------------------------------------------------------------------
//
// These helpers each return either an empty string (when the underlying
// field is at its default value, meaning the worksheet XML omits the
// element entirely) or a self-contained XML chunk that the caller
// inserts inside the `<worksheet>` element. Keeping them local to this
// translation unit avoids cross-bundle collisions; they consume only
// the public Sheet accessors documented in `src/sheet.h`.

/// Emits `<sheetViews><sheetView>...</sheetView></sheetViews>` for the
/// sheet's view state. Returns an empty string when every field is at
/// its default (zoom 100, no freeze panes, tab not hidden) — Excel
/// itself omits the element in that case.
std::string BuildSheetViewXml(const SheetView& view) {
  const bool zoom_default = view.zoom_scale == SheetView::kDefaultZoomScale;
  const bool no_freeze = view.freeze_rows == 0U && view.freeze_cols == 0U;
  const bool tab_default = !view.tab_hidden;
  if (zoom_default && no_freeze && tab_default) {
    return std::string();
  }
  std::string out;
  out.reserve(192);
  out.append("<sheetViews><sheetView workbookViewId=\"0\"");
  if (!zoom_default) {
    out.append(" zoomScale=\"");
    out.append(std::to_string(view.zoom_scale));
    out.push_back('"');
  }
  if (no_freeze) {
    out.append("/></sheetViews>");
    return out;
  }
  out.push_back('>');
  // OOXML attribute order for `<pane>`: xSplit, ySplit, topLeftCell,
  // activePane, state. We emit only the fields we own; the writer
  // does not yet model `topLeftCell` or `activePane`, so they are
  // absent. Excel gracefully accepts a freeze record without them.
  out.append("<pane");
  if (view.freeze_cols != 0U) {
    out.append(" xSplit=\"");
    out.append(std::to_string(view.freeze_cols));
    out.push_back('"');
  }
  if (view.freeze_rows != 0U) {
    out.append(" ySplit=\"");
    out.append(std::to_string(view.freeze_rows));
    out.push_back('"');
  }
  out.append(" state=\"frozen\"/></sheetView></sheetViews>");
  return out;
}

/// Emits `<sheetProtection .../>` matching the structure ECMA-376
/// §18.3.1.85 prescribes. Returns an empty string when
/// `p.enabled == false` so the caller can drop the surrounding
/// indentation cleanly. Boolean attributes are emitted only when their
/// value is `true`; the spec defaults absent attributes to `false`.
std::string BuildSheetProtectionXml(const SheetProtection& p) {
  if (!p.enabled) {
    return std::string();
  }
  std::string out;
  out.reserve(256);
  out.append("<sheetProtection");
  if (!p.algorithm_name.empty()) {
    out.append(" algorithmName=\"");
    AppendXmlEscaped(out, p.algorithm_name);
    out.push_back('"');
  }
  if (!p.hash_value.empty()) {
    out.append(" hashValue=\"");
    AppendXmlEscaped(out, p.hash_value);
    out.push_back('"');
  }
  if (!p.salt_value.empty()) {
    out.append(" saltValue=\"");
    AppendXmlEscaped(out, p.salt_value);
    out.push_back('"');
  }
  if (p.spin_count != 0U) {
    out.append(" spinCount=\"");
    out.append(std::to_string(p.spin_count));
    out.push_back('"');
  }
  if (!p.legacy_password.empty()) {
    out.append(" password=\"");
    AppendXmlEscaped(out, p.legacy_password);
    out.push_back('"');
  }
  // Boolean attributes — only emit when `true`. Order mirrors Excel's
  // own emission order so byte-identical round-trips are achievable
  // for the common cases.
  const auto append_bool = [&out](const char* name, bool v) {
    if (!v) {
      return;
    }
    out.push_back(' ');
    out.append(name);
    out.append("=\"1\"");
  };
  append_bool("sheet", p.sheet);
  append_bool("objects", p.objects);
  append_bool("scenarios", p.scenarios);
  append_bool("formatCells", p.format_cells);
  append_bool("formatColumns", p.format_columns);
  append_bool("formatRows", p.format_rows);
  append_bool("insertColumns", p.insert_columns);
  append_bool("insertRows", p.insert_rows);
  append_bool("insertHyperlinks", p.insert_hyperlinks);
  append_bool("deleteColumns", p.delete_columns);
  append_bool("deleteRows", p.delete_rows);
  append_bool("selectLockedCells", p.select_locked_cells);
  append_bool("sort", p.sort);
  append_bool("autoFilter", p.auto_filter);
  append_bool("pivotTables", p.pivot_tables);
  append_bool("selectUnlockedCells", p.select_unlocked_cells);
  out.append("/>");
  return out;
}

/// Emits `<cols>...</cols>` containing one `<col>` per
/// `ColumnLayout` entry. Returns an empty string when there are no
/// column layout overrides.
std::string BuildColsXml(const SheetLayout& layout) {
  if (layout.columns.empty()) {
    return std::string();
  }
  std::string out;
  out.reserve(32U + layout.columns.size() * 96U);
  out.append("<cols>");
  for (const ColumnLayout& col : layout.columns) {
    out.append("<col min=\"");
    out.append(std::to_string(col.first + 1U));
    out.append("\" max=\"");
    out.append(std::to_string(col.last + 1U));
    out.push_back('"');
    if (col.width > 0.0) {
      out.append(" width=\"");
      char buf[32];
      std::snprintf(buf, sizeof(buf), "%.6g", col.width);
      out.append(buf);
      // Excel emits `customWidth="1"` whenever an explicit `width` is
      // present so a reload preserves the column metric.
      out.append("\" customWidth=\"1\"");
    }
    if (col.hidden) {
      out.append(" hidden=\"1\"");
    }
    if (col.outline_level != 0U) {
      out.append(" outlineLevel=\"");
      out.append(std::to_string(static_cast<unsigned int>(col.outline_level)));
      out.push_back('"');
    }
    out.append("/>");
  }
  out.append("</cols>");
  return out;
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> write_ooxml(const Workbook& wb) {
  const std::size_t sheet_count = wb.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "workbook has zero sheets", "context=write_ooxml");
  }

  const EmissionPlan plan = BuildEmissionPlan(wb);

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_ooxml");
  }

  // 1. [Content_Types].xml
  {
    auto result = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(wb, plan));
    if (!result) {
      return result.error();
    }
  }

  // 2. _rels/.rels
  {
    auto result = AddPart(writer.get(), "_rels/.rels", BuildPackageRels());
    if (!result) {
      return result.error();
    }
  }

  // 3. xl/workbook.xml
  {
    auto result = AddPart(writer.get(), "xl/workbook.xml", BuildWorkbookXml(wb, plan));
    if (!result) {
      return result.error();
    }
  }

  // 4. xl/_rels/workbook.xml.rels
  {
    auto result = AddPart(writer.get(), "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheet_count, plan));
    if (!result) {
      return result.error();
    }
  }

  // 5. Per-sheet: worksheet, sheet rels (when the sheet owns tables,
  // pivot tables, hyperlinks, or comments).
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const auto& sheet_tables = plan.tables_by_sheet[i];
    const auto& sheet_pivot_tables = plan.pivot_tables_by_sheet[i];
    const auto& comments_plan = plan.comments_by_sheet[i];
    const bool has_hyperlinks = !wb.sheet(i).hyperlinks().empty();
    const bool has_comments = comments_plan.numeric_id != 0;
    const bool has_rels = !sheet_tables.empty() || !sheet_pivot_tables.empty() || has_hyperlinks || has_comments;
    // Build the rels first because the hyperlink rId vector feeds into
    // the worksheet's <hyperlinks> block. When the sheet has no rels we
    // still call BuildSheetRels with an empty comments plan to get a
    // (possibly-empty) hyperlink_rids vector.
    SheetRelsResult rels_result = BuildSheetRels(wb.sheet(i), sheet_tables, sheet_pivot_tables, comments_plan);
    std::string part_path("xl/worksheets/sheet");
    part_path.append(std::to_string(i + 1));
    part_path.append(".xml");
    auto wresult =
        AddPart(writer.get(), part_path, BuildWorksheetXml(wb.sheet(i), sheet_tables, rels_result.hyperlink_rids));
    if (!wresult) {
      return wresult.error();
    }
    if (has_rels) {
      std::string rels_path("xl/worksheets/_rels/sheet");
      rels_path.append(std::to_string(i + 1));
      rels_path.append(".xml.rels");
      auto rels_add = AddPart(writer.get(), rels_path, rels_result.xml);
      if (!rels_add) {
        return rels_add.error();
      }
    }
  }

  // 6. xl/styles.xml
  {
    auto result = AddPart(writer.get(), "xl/styles.xml", write_styles(wb.styles()));
    if (!result) {
      return result.error();
    }
  }

  // 7. xl/tables/tableN.xml — one per planned table.
  for (const auto& per_sheet : plan.tables_by_sheet) {
    for (const EmissionPlan::PerSheetTable& t : per_sheet) {
      auto result = AddPart(writer.get(), t.path, BuildTableXml(*t.table, t.numeric_id));
      if (!result) {
        return result.error();
      }
    }
  }

  // 8. Pivot caches — for each planned cache, emit the definition, the
  // records, and the definition's own rels file (which points at the
  // matching records part). The records target stored on the rels file
  // is just the basename because both parts live in the same package
  // directory (`xl/pivotCache/`).
  for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
    {
      auto result = AddPart(writer.get(), c.definition_path, write_pivot_cache_definition(*c.cache));
      if (!result) {
        return result.error();
      }
    }
    {
      auto result = AddPart(writer.get(), c.records_path, write_pivot_cache_records(*c.cache));
      if (!result) {
        return result.error();
      }
    }
    {
      // Records target is relative to the definition's directory, so
      // pass the basename of `records_path` (everything after the last
      // `/`). The basename is guaranteed to be present given the path
      // template, but defend against unexpected reshapes anyway.
      const std::size_t slash = c.records_path.find_last_of('/');
      const std::string_view records_filename = slash == std::string::npos
                                                    ? std::string_view(c.records_path)
                                                    : std::string_view(c.records_path).substr(slash + 1);
      auto result = AddPart(writer.get(), c.definition_rels_path, BuildPivotCacheDefinitionRels(records_filename));
      if (!result) {
        return result.error();
      }
    }
  }

  // 9. Pivot tables — one per planned pivot-table entry, package-wide.
  // Sheet-rels emission in step 5 already wired a rId to each part.
  for (const auto& per_sheet : plan.pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      auto result = AddPart(writer.get(), t.path, write_pivot_table_definition(*t.table));
      if (!result) {
        return result.error();
      }
    }
  }

  // 9.25. External link rels — one per `wb.external_links()` record
  // with a captured target URL. The body part itself rides through
  // passthrough; we only generate the rels file pointing at the remote
  // workbook URL. Records without a target are skipped (Excel writers
  // never emit a relationship-less rels file and would treat one as
  // malformed).
  for (const EmissionPlan::ExternalLinkPlan& e : plan.external_links) {
    std::string rels_xml = BuildExternalLinkRels(*e.record);
    if (rels_xml.empty()) {
      continue;
    }
    const std::size_t slash = e.record->part_path.find_last_of('/');
    std::string rels_path;
    if (slash == std::string::npos) {
      rels_path = "_rels/" + e.record->part_path + ".rels";
    } else {
      rels_path.append(e.record->part_path.substr(0, slash));
      rels_path.append("/_rels/");
      rels_path.append(e.record->part_path.substr(slash + 1));
      rels_path.append(".rels");
    }
    auto result = AddPart(writer.get(), rels_path, std::move(rels_xml));
    if (!result) {
      return result.error();
    }
  }

  // 9.5. Comments + VML drawings — one pair per sheet that has at
  // least one comment. The VML companion uses the passthrough bytes
  // when available so unchanged round-trips stay byte-identical;
  // otherwise the writer's stub keeps Excel happy on a fresh comment.
  for (std::size_t i = 0; i < plan.comments_by_sheet.size(); ++i) {
    const EmissionPlan::CommentsPlan& cplan = plan.comments_by_sheet[i];
    if (cplan.numeric_id == 0) {
      continue;
    }
    auto cresult = AddPart(writer.get(), cplan.comments_path, write_comments(wb.sheet(i).comments()));
    if (!cresult) {
      return cresult.error();
    }
    if (cplan.vml_source != nullptr) {
      auto vresult = AddPartBytes(writer.get(), cplan.vml_path, cplan.vml_source->bytes);
      if (!vresult) {
        return vresult.error();
      }
    } else {
      auto vresult = AddPart(writer.get(), cplan.vml_path, write_vml_drawing_stub());
      if (!vresult) {
        return vresult.error();
      }
    }
  }

  // 10. Passthrough parts — bytes from the original archive, written
  // verbatim. Their `<Override>` registration was already emitted in
  // step 1 (when content_type was non-empty).
  for (const PassthroughPart* part : plan.passthrough_kept) {
    auto result = AddPartBytes(writer.get(), part->path, part->bytes);
    if (!result) {
      return result.error();
    }
  }

  // Finalise into a heap buffer, then copy into a std::vector so the caller
  // owns the bytes through normal RAII.
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(writer.get(), &archive_ptr, &archive_size) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_finalize_heap_archive failed",
                      "context=write_ooxml");
  }
  if (mz_zip_writer_end(writer.get()) == MZ_FALSE) {
    // finalize succeeded but end failed — still free the buffer miniz handed
    // us before surfacing the error.
    if (archive_ptr != nullptr) {
      mz_free(archive_ptr);
    }
    writer.release();
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_end failed", "context=write_ooxml");
  }
  writer.release();

  std::vector<std::uint8_t> bytes;
  bytes.resize(archive_size);
  if (archive_size > 0 && archive_ptr != nullptr) {
    std::memcpy(bytes.data(), archive_ptr, archive_size);
  }
  if (archive_ptr != nullptr) {
    mz_free(archive_ptr);
  }
  return bytes;
}

}  // namespace io
}  // namespace formulon
