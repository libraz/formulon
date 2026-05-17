// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Workbook-level OOXML part builders: `[Content_Types].xml`,
// `_rels/.rels`, `xl/workbook.xml`, and `xl/_rels/workbook.xml.rels`.
// See header for the caller contract.
//
// The writer-only content-type and relationship-type URI constants live
// at file scope (anonymous namespace) below: keeping them with internal
// linkage avoids ODR-emit overhead in WASM for sibling TUs that do not
// consume the strings.

#include "io/ooxml/workbook_xml_builder.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "eval/iterative_solver.h"
#include "io/defined_names.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml_defs.h"
#include "io/passthrough_part.h"
#include "io/unknown_relationship.h"
#include "io/workbook_kind.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "utils/double_format.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------
// `kXmlDecl` lives in `io/xml_utils.h` and is shared with comments / cf
// writers (single source of truth for the XML 1.0 prologue).
//
// Relationship type URIs (the `kRel*` family that is also consumed by
// the reader) live in `io/ooxml_defs.h`; the writer-only relationship
// URIs (`kRelCalcChain`, `kRelTheme`, `kRelCoreProperties`,
// `kRelExtendedProperties`) stay below.

constexpr std::string_view kCtPackageRels = "application/vnd.openxmlformats-package.relationships+xml";
constexpr std::string_view kCtXml = "application/xml";
constexpr std::string_view kCtWorksheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
constexpr std::string_view kCtStyles = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
constexpr std::string_view kCtTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml";
constexpr std::string_view kCtPivotCacheDefinition =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml";
constexpr std::string_view kCtPivotCacheRecords =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml";
constexpr std::string_view kCtPivotTable = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml";
constexpr std::string_view kCtComments = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
constexpr std::string_view kCtVmlDrawing = "application/vnd.openxmlformats-officedocument.vmlDrawing";

// Writer-only relationship URIs (no reader consumer).
constexpr std::string_view kRelCalcChain =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
constexpr std::string_view kRelTheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
constexpr std::string_view kRelCoreProperties =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
constexpr std::string_view kRelExtendedProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

/// Escapes `text` and appends it as the body of an XML element. Callers
/// that need attribute escaping should use `AppendXmlEscaped` directly.
inline void AppendEscaped(std::string& out, std::string_view text) {
  AppendXmlEscaped(out, text);
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

}  // namespace

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

std::string BuildPackageRels(const EmissionPlan& plan) {
  std::string out;
  out.reserve(256);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  AppendRelationship(out, 1, kRelOfficeDocument, "xl/workbook.xml");
  std::uint32_t next_rid = 2;
  if (HasPassthroughPart(plan, "docProps/core.xml")) {
    AppendRelationship(out, next_rid++, kRelCoreProperties, "docProps/core.xml");
  }
  if (HasPassthroughPart(plan, "docProps/app.xml")) {
    AppendRelationship(out, next_rid++, kRelExtendedProperties, "docProps/app.xml");
  }
  out.append("</Relationships>\n");
  return out;
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

std::string BuildWorkbookRels(std::size_t sheet_count, const EmissionPlan& plan, const Workbook& wb) {
  std::string out;
  out.reserve(256 + sheet_count * 192 + plan.pivot_caches.size() * 192 + wb.unknown_workbook_rels().size() * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const std::string target = "worksheets/sheet" + std::to_string(i + 1) + ".xml";
    AppendRelationship(out, static_cast<std::uint32_t>(i + 1), kRelWorksheet, target);
  }
  // Styles relationship follows the worksheet relationships.
  AppendRelationship(out, static_cast<std::uint32_t>(sheet_count + 1), kRelStyles, "styles.xml");
  // Pivot-cache definition relationships, one per planned cache. Targets
  // are relative to the workbook directory (`xl/`); we strip the `xl/`
  // prefix from `definition_path` so the form matches what Excel emits
  // (e.g. `Target="pivotCache/pivotCacheDefinition1.xml"`).
  for (const EmissionPlan::PivotCachePlan& c : plan.pivot_caches) {
    AppendRelationship(out, c.workbook_rid, kRelPivotCacheDefinition, WithoutXlPrefix(c.definition_path));
  }
  // External link relationships. Same `xl/` prefix stripping as pivot
  // caches above; targets land as `Target="externalLinks/externalLink1.xml"`.
  for (const EmissionPlan::ExternalLinkPlan& e : plan.external_links) {
    AppendRelationship(out, e.workbook_rid, kRelExternalLink, WithoutXlPrefix(e.record->part_path),
                       /*target_external=*/false, /*escape_target=*/true);
  }
  // Round-tripped relationships whose Type URI the reader did not
  // recognise (theme, calcChain, vbaProject, customXml, ...). Without
  // these the passthrough parts they point at become orphans and Excel
  // opens the package in "needs repair" mode. Fresh deterministic rIds
  // sit past the worksheet / styles / pivot / external-link numbering.
  std::uint32_t next_rid =
      static_cast<std::uint32_t>(sheet_count + 2 + plan.pivot_caches.size() + plan.external_links.size());
  auto has_unknown_rel = [&wb](std::string_view type, std::string_view resolved_target) {
    for (const UnknownRelationship& r : wb.unknown_workbook_rels()) {
      if (!r.target_external && r.type == type && r.target == resolved_target) {
        return true;
      }
    }
    return false;
  };
  if (HasPassthroughPart(plan, "xl/calcChain.xml") && !has_unknown_rel(kRelCalcChain, "xl/calcChain.xml")) {
    AppendRelationship(out, next_rid++, kRelCalcChain, "calcChain.xml");
  }
  if (HasPassthroughPart(plan, "xl/theme/theme1.xml") && !has_unknown_rel(kRelTheme, "xl/theme/theme1.xml")) {
    AppendRelationship(out, next_rid++, kRelTheme, "theme/theme1.xml");
  }
  if (HasPassthroughPart(plan, "xl/sharedStrings.xml")) {
    AppendRelationship(out, next_rid++, kRelSharedStrings, "sharedStrings.xml");
  }
  for (const UnknownRelationship& r : wb.unknown_workbook_rels()) {
    const std::string_view target =
        r.target_external ? std::string_view(r.target) : WithoutXlPrefix(std::string_view(r.target));
    AppendRelationship(out, next_rid++, std::string_view(r.type), target, r.target_external,
                       /*escape_target=*/true);
  }
  out.append("</Relationships>\n");
  return out;
}

}  // namespace io
}  // namespace formulon
