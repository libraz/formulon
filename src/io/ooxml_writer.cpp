//
// OOXML (.xlsx) package writer. The writer emits the minimum spreadsheet
// surface that Excel 365 will open without complaint, and additionally
// round-trips the metadata Bundles 2.3 and 2.4 wired into the reader:
// defined names, table parts, and Override-listed parts the reader did
// not consume (unknown-part passthrough). Literal text cells are interned in
// xl/sharedStrings.xml so repeated values are emitted once per package.
//
// This TU is now a thin orchestrator. Emission planning, relationship
// emission, miniz wrappers, and cell-reference formatting live in
// sibling TUs under `src/io/ooxml/`. Workbook-level XML body builders
// (Content Types, package rels, workbook part, workbook rels) live in
// `src/io/ooxml/workbook_xml_builder.{h,cpp}`. Per-sheet XML body
// builders (worksheet part, sheet views, cols, merge cells, data
// validations, hyperlinks, sheet protection, page breaks, sheet rels)
// live in `src/io/ooxml/sheet_xml_builder.{h,cpp}`. The pipeline-only
// pieces — `BuildTableXml`, `BuildExternalLinkRels`,
// `BuildPivotCacheDefinitionRels`, and `write_ooxml()` itself — stay
// here because each is consumed only by `write_ooxml()`.

#include "io/ooxml_writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/comments_writer.h"
#include "io/external_links.h"
#include "io/ooxml/emission_plan.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/shared_strings_writer.h"
#include "io/ooxml/sheet_xml_builder.h"
#include "io/ooxml/workbook_xml_builder.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/passthrough_part.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/styles_writer.h"
#include "io/tables_reader.h"
#include "io/xml_escape.h"
#include "io/xml_utils.h"
#include "miniz.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

/// Builds the `_rels` document for a pivotCacheDefinition part: a single
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

/// Builds the `_rels` document for a pivotTable part: a single
/// relationship of type `pivotCacheDefinition` pointing at the cache the
/// table draws from. ECMA-376 §12.3.19 requires this relationship even
/// though the table also names the cache by `cacheId`, so consumers that
/// navigate the package by relationship alone can still find it. The
/// target steps out of `xl/pivotTables/` into `xl/pivotCache/`.
std::string BuildPivotTableRels(std::string_view cache_definition_target) {
  std::string out;
  out.reserve(256 + cache_definition_target.size());
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  AppendRelationship(out, 1, kRelPivotCacheDefinition, cache_definition_target);
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

std::string BuildTableXml(const TableMetadata& t, std::uint32_t numeric_id) {
  std::string out;
  out.reserve(256 + t.columns.size() * 64);
  out.append(kXmlDecl);
  out.append("<table xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" id=\"");
  out.append(std::to_string(numeric_id));
  out.append("\" name=\"");
  AppendXmlAttrEscaped(out, t.name);
  out.append("\" displayName=\"");
  AppendXmlAttrEscaped(out, t.display_name);
  out.append("\" ref=\"");
  AppendXmlAttrEscaped(out, t.ref);
  out.push_back('"');
  out.append(t.root_extra_attrs);
  // headerRowCount: emit explicit "0" when disabled; default (1) is
  // implicit via OOXML schema and we omit it.
  if (!t.header_row) {
    out.append(" headerRowCount=\"0\"");
  }
  if (t.totals_row) {
    out.append(" totalsRowCount=\"1\"");
  }
  out.append(">\n");
  // CT_Table orders the optional filter/sort payload before the required
  // tableColumns collection. Keep these fragments opaque: they may carry
  // filter criteria, extension attributes, or other fields the table model
  // does not interpret.
  if (!t.auto_filter_xml.empty()) {
    out.append("  ");
    out.append(t.auto_filter_xml);
    out.push_back('\n');
  }
  if (!t.sort_state_xml.empty()) {
    out.append("  ");
    out.append(t.sort_state_xml);
    out.push_back('\n');
  }
  out.append("  <tableColumns count=\"");
  out.append(std::to_string(t.columns.size()));
  out.append("\">\n");
  for (const TableColumn& col : t.columns) {
    out.append("    <tableColumn id=\"");
    out.append(std::to_string(col.id));
    out.append("\" name=\"");
    AppendXmlAttrEscaped(out, col.name);
    out.push_back('"');
    if (!col.totals_label.empty()) {
      out.append(" totalsRowLabel=\"");
      AppendXmlAttrEscaped(out, col.totals_label);
      out.push_back('"');
    }
    if (!col.totals_function.empty()) {
      out.append(" totalsRowFunction=\"");
      AppendXmlAttrEscaped(out, col.totals_function);
      out.push_back('"');
    }
    out.append(col.extra_attrs);
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
  // `<tableStyleInfo>` follows `<tableColumns>` in the CT_Table schema.
  // Re-emit the captured element verbatim so banded-row / style-name
  // metadata survives the round trip.
  if (!t.table_style_info_xml.empty()) {
    out.append("  ");
    out.append(t.table_style_info_xml);
    out.push_back('\n');
  }
  if (!t.ext_lst_xml.empty()) {
    out.append("  ");
    out.append(t.ext_lst_xml);
    out.push_back('\n');
  }
  out.append("</table>\n");
  return out;
}

}  // namespace

Expected<std::vector<std::uint8_t>, Error> write_ooxml(const Workbook& wb) {
  const std::size_t sheet_count = wb.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "workbook has zero sheets", "context=write_ooxml");
  }
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index >= sheet_count) {
      return make_error(FormulonErrorCode::kIoWriteFailed, "table references a missing worksheet",
                        "context=write_ooxml table=" + table.name + " sheet_index=" +
                            std::to_string(table.sheet_index) + " sheet_count=" + std::to_string(sheet_count));
    }
  }
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const Sheet& sheet = wb.sheet(i);
    if (!sheet.is_opaque_ooxml_sheet()) {
      continue;
    }
    const auto it =
        std::find_if(wb.passthrough_parts().begin(), wb.passthrough_parts().end(),
                     [&sheet](const PassthroughPart& part) { return part.path == sheet.opaque_ooxml_part_path(); });
    if (it == wb.passthrough_parts().end()) {
      return make_error(FormulonErrorCode::kIoWriteFailed, "opaque sheet has no passthrough payload",
                        "context=write_ooxml sheet=" + sheet.name() + " part=" + sheet.opaque_ooxml_part_path());
    }
  }

  const SharedStrings shared_strings = BuildSharedStrings(wb);
  const EmissionPlan plan = BuildEmissionPlan(wb, !shared_strings.empty());

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_ooxml");
  }

  // Tracks every path written so far. `AddPart` / `AddPartBytes` refuse
  // a repeat instead of silently producing a duplicate zip entry, which
  // readers resolve inconsistently (see `EmissionPlan`'s collision
  // detection for the deliberate-drop path; this guard catches anything
  // that slips past it).
  std::unordered_set<std::string> written_paths;

  // 1. [Content_Types].xml
  {
    auto result = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(wb, plan), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 2. _rels/.rels
  {
    auto result = AddPart(writer.get(), "_rels/.rels", BuildPackageRels(wb, plan), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 3. xl/workbook.xml
  {
    auto result = AddPart(writer.get(), "xl/workbook.xml", BuildWorkbookXml(wb, plan), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 4. xl/_rels/workbook.xml.rels
  {
    auto result =
        AddPart(writer.get(), "xl/_rels/workbook.xml.rels", BuildWorkbookRels(sheet_count, plan, wb), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 5. Per-sheet: worksheet, sheet rels (when the sheet's rels file
  // declares at least one relationship).
  for (std::size_t i = 0; i < sheet_count; ++i) {
    if (wb.sheet(i).is_opaque_ooxml_sheet()) {
      continue;
    }
    const auto& sheet_tables = plan.tables_by_sheet[i];
    // Built once in `BuildEmissionPlan`; reused verbatim here so the
    // worksheet body's `r:id` references and the rels file this writer
    // emits always agree on which id names which relationship.
    const SheetRelsResult& rels_result = plan.sheet_rels[i];
    std::string part_path("xl/worksheets/sheet");
    part_path.append(std::to_string(i + 1));
    part_path.append(".xml");
    auto wresult =
        AddPart(writer.get(), part_path,
                BuildWorksheetXml(wb.sheet(i), sheet_tables, rels_result.table_rids, rels_result.hyperlink_rids,
                                  rels_result.printer_settings_rid, rels_result.drawing_rid,
                                  rels_result.legacy_drawing_rid, &shared_strings, wb.styles().dxfs.size()),
                &written_paths);
    if (!wresult) {
      return wresult.error();
    }
    // An empty rels file is invalid OOXML; Excel repairs the package and
    // drops it. `relationship_count` is the single predicate for whether
    // this sheet's rels part is worth writing at all — for example, a
    // sheet whose only hyperlinks are purely internal (target carried in
    // the inline `location=` attribute) declares zero relationships and
    // gets no rels part.
    if (rels_result.relationship_count > 0) {
      std::string rels_path("xl/worksheets/_rels/sheet");
      rels_path.append(std::to_string(i + 1));
      rels_path.append(".xml.rels");
      auto rels_add = AddPart(writer.get(), rels_path, rels_result.xml, &written_paths);
      if (!rels_add) {
        return rels_add.error();
      }
    }
  }

  // 6. xl/styles.xml
  {
    auto result = AddPart(writer.get(), "xl/styles.xml", write_styles(wb.styles()), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 6.5. xl/sharedStrings.xml — emitted only when literal text cells are
  // present. The matching relationship and content-type Override are
  // generated from the same plan flag above.
  if (!shared_strings.empty()) {
    auto result = AddPart(writer.get(), "xl/sharedStrings.xml", WriteSharedStrings(shared_strings), &written_paths);
    if (!result) {
      return result.error();
    }
  }

  // 7. xl/tables/tableN.xml — one per planned table.
  for (const auto& per_sheet : plan.tables_by_sheet) {
    for (const EmissionPlan::PerSheetTable& t : per_sheet) {
      auto result = AddPart(writer.get(), t.path, BuildTableXml(*t.table, t.numeric_id), &written_paths);
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
      auto result = AddPart(writer.get(), c.definition_path, write_pivot_cache_definition(*c.cache), &written_paths);
      if (!result) {
        return result.error();
      }
    }
    {
      auto result = AddPart(writer.get(), c.records_path, write_pivot_cache_records(*c.cache), &written_paths);
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
      auto result = AddPart(writer.get(), c.definition_rels_path, BuildPivotCacheDefinitionRels(records_filename),
                            &written_paths);
      if (!result) {
        return result.error();
      }
    }
  }

  // 9. Pivot tables — one per planned pivot-table entry, package-wide.
  // Sheet-rels emission in step 5 already wired a rId to each part.
  for (const auto& per_sheet : plan.pivot_tables_by_sheet) {
    for (const EmissionPlan::PivotTablePlan& t : per_sheet) {
      auto result = AddPart(writer.get(), t.path, write_pivot_table_definition(*t.table), &written_paths);
      if (!result) {
        return result.error();
      }
      // The table's own rels file names the cache definition it draws
      // from. A table whose cache id resolved to no planned cache gets
      // no rels part, so the package never carries a dangling target.
      if (!t.cache_definition_target.empty()) {
        auto rels_result =
            AddPart(writer.get(), t.rels_path, BuildPivotTableRels(t.cache_definition_target), &written_paths);
        if (!rels_result) {
          return rels_result.error();
        }
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
    std::string rels_path = ooxml::rels_path_for_part(e.record->part_path);
    auto result = AddPart(writer.get(), rels_path, rels_xml, &written_paths);
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
    auto cresult = AddPart(writer.get(), cplan.comments_path, write_comments(wb.sheet(i).comments()), &written_paths);
    if (!cresult) {
      return cresult.error();
    }
    if (cplan.vml_source != nullptr) {
      auto vresult = AddPartBytes(writer.get(), cplan.vml_path, cplan.vml_source->bytes, &written_paths);
      if (!vresult) {
        return vresult.error();
      }
    } else {
      auto vresult = AddPart(writer.get(), cplan.vml_path, write_vml_drawing_stub(), &written_paths);
      if (!vresult) {
        return vresult.error();
      }
    }
  }

  // 10. Passthrough parts — bytes from the original archive, written
  // verbatim. Their `<Override>` registration was already emitted in
  // step 1 (when content_type was non-empty).
  for (const PassthroughPart* part : plan.passthrough_kept) {
    // Never emit a traversal-shaped part name, even if one reached the model
    // through a path other than the reader (which already rejects them).
    if (!ooxml::is_safe_part_name(part->path)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "passthrough part name escapes package root; refusing to write",
                        "context=write_ooxml part=" + part->path);
    }
    auto result = AddPartBytes(writer.get(), part->path, part->bytes, &written_paths);
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
