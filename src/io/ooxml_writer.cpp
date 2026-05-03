// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML (.xlsx) package writer. The writer emits the minimum spreadsheet
// surface that Excel 365 will open without complaint, and additionally
// round-trips the metadata Bundles 2.3 and 2.4 wired into the reader:
// defined names, table parts, and Override-listed parts the reader did
// not consume (unknown-part passthrough). Cells are still emitted with
// inline strings (`t="inlineStr"`); SST emission would force every text
// cell to walk a side table for no observable gain — the inline form
// round-trips cleanly already.
//
// Design references:
//   * backup/plans/04-xlsx-io.md §4.2 (package structure)
//   * backup/plans/04-xlsx-io.md §4.3 (part classification)
//   * backup/plans/26-implementation-plan.md (Phase 2.5)

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

#include "io/cf_writer.h"
#include "io/defined_names.h"
#include "io/ooxml_writer_cell.h"
#include "io/passthrough_part.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/xml_escape.h"
#include "miniz.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
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

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

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
    const std::vector<std::vector<EmissionPlan::PivotTablePlan>>& pivot_tables_by_sheet) {
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

  // Collision detection between generated paths and passthrough paths.
  // Generated paths win; passthrough copy is dropped with a warning.
  std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, flat_tables, plan.pivot_caches, plan.pivot_tables_by_sheet);
  // Sheet rels for sheets that own tables OR pivot tables are also
  // generated.
  for (std::size_t i = 0; i < plan.tables_by_sheet.size(); ++i) {
    const bool has_tables = !plan.tables_by_sheet[i].empty();
    const bool has_pivots = i < plan.pivot_tables_by_sheet.size() && !plan.pivot_tables_by_sheet[i].empty();
    if (has_tables || has_pivots) {
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
  out.append("  <Override PartName=\"/xl/workbook.xml\" ContentType=\"");
  out.append(workbook_kind_content_type(wb.kind()));
  out.append("\"/>\n");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out.append("  <Override PartName=\"/xl/worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".xml\" ContentType=\"");
    out.append(kCtWorksheet);
    out.append("\"/>\n");
  }
  out.append("  <Override PartName=\"/xl/styles.xml\" ContentType=\"");
  out.append(kCtStyles);
  out.append("\"/>\n");
  // Per-table overrides (one per emitted table part, regardless of
  // owning sheet).
  for (const auto& per_sheet : plan.tables_by_sheet) {
    for (const EmissionPlan::PerSheetTable& t : per_sheet) {
      out.append("  <Override PartName=\"/");
      out.append(t.path);
      out.append("\" ContentType=\"");
      out.append(kCtTable);
      out.append("\"/>\n");
    }
  }
  // Pivot caches: one Override per definition + one per records part.
  for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
    out.append("  <Override PartName=\"/");
    out.append(c.definition_path);
    out.append("\" ContentType=\"");
    out.append(kCtPivotCacheDefinition);
    out.append("\"/>\n");
    out.append("  <Override PartName=\"/");
    out.append(c.records_path);
    out.append("\" ContentType=\"");
    out.append(kCtPivotCacheRecords);
    out.append("\"/>\n");
  }
  // Pivot tables: one Override per part.
  for (const auto& per_sheet : plan.pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      out.append("  <Override PartName=\"/");
      out.append(t.path);
      out.append("\" ContentType=\"");
      out.append(kCtPivotTable);
      out.append("\"/>\n");
    }
  }
  // Passthrough overrides: only for entries that carried an explicit
  // ContentType in the source archive. Default-typed parts (empty
  // content_type) must NOT appear as Overrides — the package's
  // `<Default Extension=...>` entries already cover them.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part->content_type.empty()) {
      continue;
    }
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
  // <definedNames> sits between <sheets> and <calcPr>/end-of-workbook
  // per OOXML schema (cf. ECMA-376 sheet ordering).
  AppendDefinedNamesBlock(out, wb.defined_names());
  // <pivotCaches> follows <definedNames> in the ECMA-376 schema element
  // order: sheets, functionGroups, externalReferences, definedNames,
  // calcPr, oleSize, customWorkbookViews, pivotCaches, ...
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
  out.append("</Relationships>\n");
  return out;
}

// Forward declarations for the helpers defined at the bottom of this
// file. Each builder is fenced into its own translation-unit-local
// helper so concurrent Wave 1 bundles can extend the worksheet XML
// shape without colliding here.
std::string BuildSheetViewXml(const SheetView& view);
std::string BuildColsXml(const SheetLayout& layout);

std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables) {
  const std::string sheet_view_xml = BuildSheetViewXml(sheet.view());
  const std::string cols_xml = BuildColsXml(sheet.layout());
  const std::string sheet_data = BuildSheetDataXml(sheet);
  // Conditional-format blocks live between <sheetData> and <tableParts>
  // in ECMA-376 document order. Empty list => empty string, no
  // wrapper.
  const std::string cf_xml = write_conditional_formattings(sheet.conditional_formats());
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + cols_xml.size() + sheet_data.size() + cf_xml.size() +
              sheet_tables.size() * 96);
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
  if (!cf_xml.empty()) {
    out.append("  ");
    out.append(cf_xml);
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

std::string BuildSheetRels(const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                           const std::vector<EmissionPlan::PivotTablePlan>& sheet_pivot_tables) {
  std::string out;
  out.reserve(256 + (sheet_tables.size() + sheet_pivot_tables.size()) * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(i + 1));
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
    out.append(std::to_string(sheet_tables.size() + i + 1));
    out.append("\" Type=\"");
    out.append(kRelPivotTable);
    out.append("\" Target=\"../pivotTables/pivotTable");
    out.append(std::to_string(sheet_pivot_tables[i].numeric_id));
    out.append(".xml\"/>\n");
  }
  out.append("</Relationships>\n");
  return out;
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

std::string BuildStylesXml() {
  std::string out;
  out.reserve(512);
  out.append(kXmlDecl);
  out.append("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n");
  out.append("  <fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/></font></fonts>\n");
  out.append("  <fills count=\"1\"><fill><patternFill patternType=\"none\"/></fill></fills>\n");
  out.append("  <borders count=\"1\"><border/></borders>\n");
  out.append(
      "  <cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>\n");
  out.append(
      "  <cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>\n");
  out.append("</styleSheet>\n");
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

  // 5. Per-sheet: worksheet, sheet rels (when the sheet owns tables or
  // pivot tables).
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const auto& sheet_tables = plan.tables_by_sheet[i];
    const auto& sheet_pivot_tables = plan.pivot_tables_by_sheet[i];
    std::string part_path("xl/worksheets/sheet");
    part_path.append(std::to_string(i + 1));
    part_path.append(".xml");
    auto result = AddPart(writer.get(), part_path, BuildWorksheetXml(wb.sheet(i), sheet_tables));
    if (!result) {
      return result.error();
    }
    if (!sheet_tables.empty() || !sheet_pivot_tables.empty()) {
      std::string rels_path("xl/worksheets/_rels/sheet");
      rels_path.append(std::to_string(i + 1));
      rels_path.append(".xml.rels");
      auto rels_result = AddPart(writer.get(), rels_path, BuildSheetRels(sheet_tables, sheet_pivot_tables));
      if (!rels_result) {
        return rels_result.error();
      }
    }
  }

  // 6. xl/styles.xml
  {
    auto result = AddPart(writer.get(), "xl/styles.xml", BuildStylesXml());
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
