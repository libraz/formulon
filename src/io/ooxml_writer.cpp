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
//
// This TU is now a thin orchestrator: emission planning, relationship
// emission, miniz wrappers, and cell-reference formatting all live in
// sibling TUs under `src/io/ooxml/`. The XML body builders (workbook,
// worksheet, table, hyperlinks, data validations, sheet views, ...)
// stay here because they each touch the package only through
// `AddPart()` and are not consumed by any other TU.

#include "io/ooxml_writer.h"

#include <cstddef>
#include <cstdint>
#include <cstdio>
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
#include "io/ooxml/cell_ref_writer.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/ooxml_writer_cell.h"
#include "io/passthrough_part.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/styles_writer.h"
#include "io/tables_reader.h"
#include "io/unknown_relationship.h"
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

// ---------------------------------------------------------------------------
// XML helpers
// ---------------------------------------------------------------------------
//
// Relationship emission helpers (`AppendOverride`, `AppendRelationship`,
// `RelsPathForPart`, `WithoutXlPrefix`, `TargetRelativeToWorksheet`)
// live in `io/ooxml/relationship_writer.h`. Cell-reference formatters
// (`AppendColumnLettersForRef`, `AppendCellRefForRef`,
// `AppendRangeRef`) live in `io/ooxml/cell_ref_writer.h`. Emission
// planning (`EmissionPlan`, `BuildEmissionPlan`, `HasPassthroughPart`)
// lives in `io/ooxml/emission_plan.h`. Miniz wrappers (`ZipWriterGuard`,
// `AddPart`, `AddPartBytes`) live in `io/ooxml/zip_part_writer.h`.

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
  std::string_view type = kRelExternalLinkPath;
  switch (rec.kind) {
    case ExternalLinkRecord::Kind::kOleLink:
      type = kRelOleLink;
      break;
    case ExternalLinkRecord::Kind::kDdeLink:
      type = kRelDdeLink;
      break;
    case ExternalLinkRecord::Kind::kExternalBook:
    case ExternalLinkRecord::Kind::kUnknown:
    default:
      break;
  }
  AppendRelationship(out, rec.body_rel_id.empty() ? std::string_view("rId1") : std::string_view(rec.body_rel_id), type,
                     rec.target, rec.target_external, /*escape_target=*/true);
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

std::string PageSetupWithRelationshipId(std::string page_setup_xml, std::string_view rid) {
  if (page_setup_xml.empty() || rid.empty()) {
    return page_setup_xml;
  }
  auto replace_attr = [&](std::string_view attr_name) {
    const std::string needle = std::string(attr_name) + "=\"";
    const std::size_t pos = page_setup_xml.find(needle);
    if (pos == std::string::npos) {
      return false;
    }
    const std::size_t value_start = pos + needle.size();
    const std::size_t value_end = page_setup_xml.find('"', value_start);
    if (value_end == std::string::npos) {
      return false;
    }
    page_setup_xml.replace(value_start, value_end - value_start, rid);
    return true;
  };
  if (replace_attr("r:id") || replace_attr("id")) {
    return page_setup_xml;
  }
  const std::size_t insert_pos = page_setup_xml.rfind("/>");
  const std::size_t fallback_pos = page_setup_xml.rfind('>');
  const std::size_t pos = insert_pos != std::string::npos ? insert_pos : fallback_pos;
  if (pos == std::string::npos) {
    return page_setup_xml;
  }
  std::string attr(" r:id=\"");
  attr.append(rid);
  attr.push_back('"');
  page_setup_xml.insert(pos, attr);
  return page_setup_xml;
}

/// Emits a `<rowBreaks>` or `<colBreaks>` block for the given manual
/// breaks. Returns an empty string when `breaks` is empty so the caller
/// adds no bytes. The stored 0-based break index is converted back to
/// OOXML's 1-based form; `count` and `manualBreakCount` mirror the entry
/// count.
std::string BuildPageBreaksXml(std::string_view element, const std::vector<ManualBreak>& breaks) {
  if (breaks.empty()) {
    return {};
  }
  // Rough size estimate so the buffer rarely reallocates: a fixed
  // allowance for the wrapper element plus one allowance per `<brk>`.
  constexpr std::size_t kBreaksWrapperReserveBytes = 48U;
  constexpr std::size_t kPerBreakReserveBytes = 48U;
  std::string out;
  out.reserve(kBreaksWrapperReserveBytes + breaks.size() * kPerBreakReserveBytes);
  const std::string count = std::to_string(breaks.size());
  out.push_back('<');
  out.append(element);
  out.append(" count=\"");
  out.append(count);
  out.append("\" manualBreakCount=\"");
  out.append(count);
  out.append("\">");
  for (const ManualBreak& brk : breaks) {
    out.append("<brk id=\"");
    out.append(std::to_string(static_cast<std::uint64_t>(brk.id) + 1U));
    out.append("\" min=\"");
    out.append(std::to_string(brk.min));
    out.append("\" max=\"");
    out.append(std::to_string(brk.max));
    out.append("\"");
    if (brk.manual) {
      out.append(" man=\"1\"");
    }
    out.append("/>");
  }
  out.append("</");
  out.append(element);
  out.push_back('>');
  return out;
}

std::string BuildWorksheetXml(const Sheet& sheet, const std::vector<EmissionPlan::PerSheetTable>& sheet_tables,
                              const std::vector<std::string>& hyperlink_rids, std::string_view printer_settings_rid) {
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
  const SheetPrintSettings& print = sheet.print_settings();
  const std::string page_setup_xml = PageSetupWithRelationshipId(print.page_setup_xml, printer_settings_rid);
  std::string out;
  out.reserve(192U + sheet_view_xml.size() + cols_xml.size() + sheet_data.size() + cf_xml.size() + merges_xml.size() +
              dv_xml.size() + hl_xml.size() + print.sheet_pr_xml.size() + print.page_margins_xml.size() +
              page_setup_xml.size() + sheet_tables.size() * 96);
  out.append(kXmlDecl);
  out.append(
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n");
  // ECMA-376 element order: sheetPr -> dimension -> sheetViews ->
  // sheetFormatPr -> cols -> sheetData -> conditionalFormatting ->
  // pageMargins -> pageSetup -> rowBreaks -> colBreaks -> tableParts.
  // We currently emit a subset; the helpers stay quiet when their
  // underlying field is at default values so absent metadata yields no
  // extra bytes.
  if (!print.sheet_pr_xml.empty()) {
    out.append("  ");
    out.append(print.sheet_pr_xml);
    out.push_back('\n');
  }
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
  if (!print.page_margins_xml.empty()) {
    out.append("  ");
    out.append(print.page_margins_xml);
    out.push_back('\n');
  }
  if (!page_setup_xml.empty()) {
    out.append("  ");
    out.append(page_setup_xml);
    out.push_back('\n');
  }
  // Manual page breaks. ECMA-376 places <rowBreaks>/<colBreaks> after
  // <pageSetup> and before drawing parts / <tableParts>.
  {
    const std::string row_breaks_xml = BuildPageBreaksXml("rowBreaks", print.manual_row_breaks);
    if (!row_breaks_xml.empty()) {
      out.append("  ");
      out.append(row_breaks_xml);
      out.push_back('\n');
    }
    const std::string col_breaks_xml = BuildPageBreaksXml("colBreaks", print.manual_col_breaks);
    if (!col_breaks_xml.empty()) {
      out.append("  ");
      out.append(col_breaks_xml);
      out.push_back('\n');
    }
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
  std::string printer_settings_rid;
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
  std::unordered_set<std::string> used_rids;
  auto next_unique_rid = [&]() {
    std::string rid;
    do {
      rid = "rId" + std::to_string(next_rid++);
    } while (used_rids.count(rid) != 0U);
    used_rids.insert(rid);
    return rid;
  };
  for (std::size_t i = 0; i < sheet_tables.size(); ++i) {
    const std::string target = "../tables/table" + std::to_string(sheet_tables[i].numeric_id) + ".xml";
    AppendRelationship(out, next_unique_rid(), kRelTable, target);
  }
  // Pivot-table relationships follow the table relationships, with rId
  // numbering continuing in sequence so each rel id is unique within
  // the sheet rels file.
  for (std::size_t i = 0; i < sheet_pivot_tables.size(); ++i) {
    const std::string target = "../pivotTables/pivotTable" + std::to_string(sheet_pivot_tables[i].numeric_id) + ".xml";
    AppendRelationship(out, next_unique_rid(), kRelPivotTable, target);
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
    if (!h.rid.empty() && used_rids.count(h.rid) == 0U) {
      rid = h.rid;
      used_rids.insert(rid);
    } else {
      rid = next_unique_rid();
    }
    res.hyperlink_rids.push_back(rid);
    AppendRelationship(out, rid, kRelHyperlink, h.target, /*target_external=*/true, /*escape_target=*/true);
  }
  const SheetPrintSettings& print = sheet.print_settings();
  if (!print.printer_settings_path.empty()) {
    if (!print.printer_settings_rid.empty() && used_rids.count(print.printer_settings_rid) == 0U) {
      res.printer_settings_rid = print.printer_settings_rid;
      used_rids.insert(res.printer_settings_rid);
    } else {
      res.printer_settings_rid = next_unique_rid();
    }
    AppendRelationship(out, res.printer_settings_rid, kRelPrinterSettings,
                       TargetRelativeToWorksheet(print.printer_settings_path));
  }
  // Comments + VML relationships when the sheet has comments. The
  // comments rel comes first; the VML rel follows so the two ids are
  // adjacent and readers that scan top-down see them as a pair.
  if (comments_plan.numeric_id != 0) {
    const std::string comments_target = "../comments" + std::to_string(comments_plan.numeric_id) + ".xml";
    const std::string vml_target = "../drawings/vmlDrawing" + std::to_string(comments_plan.numeric_id) + ".vml";
    AppendRelationship(out, next_unique_rid(), kRelComments, comments_target);
    AppendRelationship(out, next_unique_rid(), kRelVmlDrawing, vml_target);
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
  AppendRelationship(out, 1, kRelPivotCacheRecords, records_filename);
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
    auto result = AddPart(writer.get(), "_rels/.rels", BuildPackageRels(plan));
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
    auto result = AddPart(writer.get(), "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheet_count, plan, wb));
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
    const bool has_print_settings = !wb.sheet(i).print_settings().printer_settings_path.empty();
    const bool has_rels =
        !sheet_tables.empty() || !sheet_pivot_tables.empty() || has_hyperlinks || has_comments || has_print_settings;
    // Build the rels first because the hyperlink rId vector feeds into
    // the worksheet's <hyperlinks> block. When the sheet has no rels we
    // still call BuildSheetRels with an empty comments plan to get a
    // (possibly-empty) hyperlink_rids vector.
    SheetRelsResult rels_result = BuildSheetRels(wb.sheet(i), sheet_tables, sheet_pivot_tables, comments_plan);
    std::string part_path("xl/worksheets/sheet");
    part_path.append(std::to_string(i + 1));
    part_path.append(".xml");
    auto wresult = AddPart(
        writer.get(), part_path,
        BuildWorksheetXml(wb.sheet(i), sheet_tables, rels_result.hyperlink_rids, rels_result.printer_settings_rid));
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
    std::string rels_path = RelsPathForPart(e.record->part_path);
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
