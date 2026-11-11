//
// Implementation of the MS-XLSB package writer. See `io/xlsb/writer.h`
// for the contract. The implementation mirrors the structure of
// `io/ooxml_writer.cpp`: build an emission plan (`io/xlsb/emission_plan.h`:
// which sheet owns which numeric id, which passthrough parts survive
// collision detection), then compose the parts and pipe them through miniz.

#include "io/xlsb/writer.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "calc_settings.h"
#include "cell.h"
#include "cf/cf_types.h"
#include "chart/chart_cache.h"
#include "external_link.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/workbook_xml_builder.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/passthrough_part_index.h"
#include "io/rich_value_errors.h"
#include "io/writer_common.h"
#include "io/xlsb/drawing_record.h"
#include "io/xlsb/emission_plan.h"
#include "io/xlsb/external_link_writer.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/part_paths.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_properties_records.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/slicer_records.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/styles_writer.h"
#include "io/xlsb/table_records.h"
#include "io/xlsb/view_records.h"
#include "io/xlsb/workbook_bin_writer.h"
#include "io/xlsb/workbook_props_records.h"
#include "miniz.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

void ReportDeferred(std::uint32_t* count, std::string_view kind, std::size_t items, std::size_t sheet_index) {
  if (items == 0U)
    return;
  const std::uint64_t total = static_cast<std::uint64_t>(*count) + items;
  *count = static_cast<std::uint32_t>(std::min<std::uint64_t>(total, std::numeric_limits<std::uint32_t>::max()));
  StructuredLog("xlsb.writer.deferred")
      .field("kind", kind)
      .field("count", static_cast<std::int64_t>(items))
      .field("sheet_index", static_cast<std::int64_t>(sheet_index))
      .warn();
}

/// True when `workbook`'s passthrough set already carries at least one
/// XLSB-native pivot table part (`xl/pivotTables/*.bin`) that is still
/// fresh against the current model.
///
/// Unlike every other feature `ReportDeferredSheetFeatures` checks, a
/// pivot table's *source* part can survive a save untouched: the XLSB
/// pivot reader deliberately never marks a decoded-or-skipped pivot part
/// as consumed (see `LoadPivotParts` in xlsb/reader.cpp), specifically so
/// it keeps round-tripping through passthrough even though this writer
/// has no pivot output of its own. `sheet.pivot_tables()` is non-empty
/// whenever the source was `.xlsb` with intact pivots, which is exactly
/// the case where nothing is actually lost -- so the deferred count below
/// only fires when there is no such surviving, fresh part (the source was
/// `.xlsx`, whose pivot parts are plain XML with no `.bin` counterpart for
/// this passthrough set to carry, or the retained part is stale, in which
/// case the save fails closed before this diagnostic ever reaches a
/// caller).
bool HasPivotPassthroughPart(const Workbook& workbook) {
  for (const PassthroughPart& part : workbook.passthrough_parts()) {
    if (part.retained_origin != PassthroughPart::RetainedOrigin::kPivotTable) {
      continue;
    }
    if (CheckRetainedPartFreshness(workbook, part)) {
      return true;
    }
  }
  return false;
}

std::uint32_t ReportDeferredSheetFeatures(const Workbook& workbook) {
  std::uint32_t count = 0U;
  const bool pivots_survive_via_passthrough = HasPivotPassthroughPart(workbook);
  const std::uint32_t views = workbook_view_count(workbook.book_views_xml());
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    const Sheet& sheet = workbook.sheet(i);
    ReportDeferred(&count, "auto_filter", sheet.has_auto_filter() ? 1U : 0U, i);
    ReportDeferred(&count, "comments", sheet.comments().size(), i);
    ReportDeferred(&count, "threaded_comments", sheet.threaded_comments().size(), i);
    ReportDeferred(&count, "pivot_tables", pivots_survive_via_passthrough ? 0U : sheet.pivot_tables().size(), i);
    const SheetPrintSettings& print = sheet.print_settings();
    const bool has_print = !print.sheet_pr_xml.empty() || !print.page_margins_xml.empty() ||
                           !print.page_setup_xml.empty() || !print.print_options_xml.empty() ||
                           !print.header_footer_xml.empty() || !print.manual_row_breaks.empty() ||
                           !print.manual_col_breaks.empty();
    ReportDeferred(&count, "print_settings", has_print ? 1U : 0U, i);
    const SheetFormatDefaults& defaults = sheet.format_defaults();
    std::size_t invalid_defaults = 0U;
    if (!IsRepresentableBaseColumnWidth(defaults.base_col_width)) {
      ++invalid_defaults;
    }
    if (defaults.has_default_col_width && !IsRepresentableDefaultColumnWidth(defaults.default_col_width)) {
      ++invalid_defaults;
    }
    if (defaults.has_default_row_height && !IsRepresentableDefaultRowHeight(defaults.default_row_height)) {
      ++invalid_defaults;
    }
    ReportDeferred(&count, "sheet_format_defaults", invalid_defaults, i);
    ReportDeferred(&count, "sheet_views", unwritten_view_part_count(sheet.view(), views), i);
    ReportDeferred(&count, "protected_ranges", unwritten_protected_range_count(sheet), i);
  }
  // Workbook globals the BrtBookProtection / BrtWbProp / BrtBookView records
  // cannot hold.
  std::vector<std::uint8_t> scratch;
  const auto book_protection = emit_book_protection(scratch, workbook.workbook_protection_xml());
  ReportDeferred(&count, "workbook_protection", book_protection && !book_protection.value() ? 1U : 0U, 0U);
  const bool wb_prop_carried = emit_wb_prop(scratch, workbook.workbook_pr_xml(), workbook.date1904());
  ReportDeferred(&count, "workbook_properties", wb_prop_carried ? 0U : 1U, 0U);
  const bool book_views_carried =
      emit_book_views(scratch, workbook.book_views_xml(), workbook.active_sheet(), workbook.sheet_count());
  ReportDeferred(&count, "book_views", book_views_carried ? 0U : 1U, 0U);
  // An .xlsx slicer part has no binary form here and is left out.
  std::size_t xml_slicers = 0U;
  for (const PassthroughPart& part : workbook.passthrough_parts()) {
    xml_slicers += is_xml_slicer_part(part.path) && strings::starts_with(part.path, "xl/slicers/") ? 1U : 0U;
  }
  ReportDeferred(&count, "slicers", xml_slicers, 0U);
  // No binary form is written for custom table / PivotTable styles or for
  // default style names other than Excel's own.
  const TableStylesSection& table_styles = workbook.styles().table_styles;
  const bool custom_defaults =
      (!table_styles.default_table_style.empty() && table_styles.default_table_style != kBuiltinDefaultTableStyle) ||
      (!table_styles.default_pivot_style.empty() && table_styles.default_pivot_style != kBuiltinDefaultPivotStyle);
  ReportDeferred(&count, "table_styles", !table_styles.styles.empty() || custom_defaults ? 1U : 0U, 0U);
  return count;
}

/// Whether the model drawing of `sheet` (which has a target) can be linked
/// from the sheet. A missing part is a dropped relationship; a part the XLSB
/// package cannot carry, or a tail the record cannot be spliced into, is a
/// deferred drawing.
bool DrawingIsLinkable(const Workbook& wb, const Sheet& sheet, std::size_t sheet_index, WriteDiagnostics* diagnostics) {
  const std::string& target = sheet.drawing_rel_target();
  const PassthroughPart* part = ooxml::is_safe_part_name(target) ? ooxml::find_passthrough_part(wb, target) : nullptr;
  if (part == nullptr) {
    RecordDroppedRelationship("xlsb.writer.sheet_rel_skipped", kRelDrawing, target, diagnostics);
    return false;
  }
  const std::string_view xml(reinterpret_cast<const char*>(part->bytes.data()), part->bytes.size());
  if (part->content_type != kCtDrawing || !drawing_frames_are_writable(xml) ||
      !scan_drawing_records(sheet.xlsb_tail().after_hyperlinks).parsed) {
    ReportDeferred(&diagnostics->deferred_feature_count, "drawings", 1U, sheet_index);
    return false;
  }
  return true;
}

/// Smallest `rId<N>` not in `taken`.
std::string FreshRelId(const std::unordered_set<std::string>& taken) {
  for (std::size_t n = 1U;; ++n) {
    std::string id = "rId" + std::to_string(n);
    if (taken.count(id) == 0U) {
      return id;
    }
  }
}

// The metadata part the package ships: whether it is generated, which
// cell-metadata entry a spill anchor's `BrtCellMeta` record must name, and
// the rich errors the cells' `BrtValueMeta` records name.
//
// `ifmd` is 1-based, and 0 means no anchor may emit a `BrtCellMeta` record at
// all: either the workbook has no spill anchors, or the metadata part that
// ships is a retained passthrough whose dynamic-array entry could not be
// identified. Both the index and the part are decided here, once, so the
// worksheet bodies and the package can never disagree about which entry a
// cell names. A retained part carries a numbering of its own, so no rich
// error is generated over it.
struct MetadataPlan {
  bool generate_part = false;
  bool dynamic_arrays = false;
  std::uint32_t ifmd = 0;
  RichErrorTable rich;
};

// True when the package carries the pivot table part that
// `sheet.pivot_tables()[index]` was loaded from.
bool PivotPartIsWritten(const Workbook& wb, std::size_t sheet, std::size_t index) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.retained_origin == PassthroughPart::RetainedOrigin::kPivotTable && part.origin_sheet_index == sheet &&
        part.origin_pivot_index == index) {
      return true;
    }
  }
  return false;
}

MetadataPlan BuildMetadataPlan(const Workbook& wb) {
  MetadataPlan plan;
  plan.dynamic_arrays = has_dynamic_array_formula(wb);
  // The first passthrough part wins the path; `BuildEmissionPlan` drops any
  // later duplicate, so that is the one whose numbering ships.
  const PassthroughPart* retained = nullptr;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == kMetadataBinPath) {
      retained = &part;
      break;
    }
  }
  if (retained == nullptr) {
    // A blocked pivot's anchor is written plain unless its pivot part is: only a part kept from an XLSB source says so.
    plan.rich = build_rich_error_table(wb, PivotPartIsWritten);
    plan.generate_part = plan.dynamic_arrays || !plan.rich.empty();
    plan.ifmd = plan.dynamic_arrays ? 1U : 0U;  // The generated part's XLDAPR entry comes first.
    return plan;
  }
  if (plan.dynamic_arrays) {
    // A retained part carries its own numbering, and its first entry is not
    // necessarily the dynamic-array one — it may hold rich-value or
    // cube-function metadata instead, or no dynamic-array type at all.
    plan.ifmd = find_dynamic_array_cell_meta_index(ByteSpan{retained->bytes.data(), retained->bytes.size()});
  }
  return plan;
}

}  // namespace

Expected<XlsbWriteResult, Error> write_xlsb_with_result(const Workbook& workbook) {
  const std::size_t sheet_count = workbook.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kInvalidArgument, "Workbook has zero sheets", kWriteXlsbContext);
  }

  // Pre-pass: emit each sheet body so we know whether the SST will be
  // non-empty. We hold the resulting bytes until after we write the
  // envelope so the order of `mz_zip_writer_add_mem` calls matches
  // what the reader expects (it does not, but we keep symmetry with
  // `write_ooxml`).
  // Ordered sheet-name list: the Ptg encoder maps a qualified
  // reference's sheet to its 0-based `ixti` through this list.
  std::vector<std::string> sheet_names = sheet_names_of(workbook);

  // PtgName record order: every genuine defined name plus every
  // future-function callee / `NameRef` any cell's formula needs, built
  // once so `ilbl` assignments are shared between the sheet bodies below
  // and the `BrtName` table `BuildWorkbookBin` emits into
  // `xl/workbook.bin`. The name -> `ilbl` map itself is per sheet: which
  // record a bare name resolves to depends on the sheet the formula sits
  // on (see `BuildNameTableForScope`).
  std::vector<OrderedName> ordered_names;
  BuildOrderedNames(workbook, ordered_names);

  // ExternSheet table: every distinct sheet-qualified reference span
  // (single- or multi-sheet) any cell or defined-name formula needs an
  // `ixti` for, built once so the assignment is shared between the
  // sheet bodies below and the `BrtExternSheet` record `BuildWorkbookBin`
  // emits into `xl/workbook.bin`.
  auto sheet_ranges_or = BuildSheetRangeTable(workbook, sheet_names);
  if (!sheet_ranges_or) {
    return std::move(sheet_ranges_or.error());
  }
  const SheetRangeTable& sheet_ranges = sheet_ranges_or.value();

  SstBuilder sst;
  const MetadataPlan metadata = BuildMetadataPlan(workbook);
  WriteDiagnostics diagnostics;
  std::uint32_t downgraded_formula_count = 0;
  diagnostics.deferred_feature_count = ReportDeferredSheetFeatures(workbook);
  std::vector<std::vector<std::uint8_t>> sheet_bodies;
  sheet_bodies.reserve(sheet_count);
  // Table parts are generated from the model, numbered clear of every part
  // the package already carries. A table on a missing sheet fails the save,
  // as it does for `.xlsx`.
  RETURN_IF_ERROR(check_tables_have_sheets(workbook, kWriteXlsbContext));
  std::unordered_set<std::string> taken_paths;
  for (const PassthroughPart& part : workbook.passthrough_parts()) {
    taken_paths.insert(part.path);
  }
  XlsbTableAllocator table_allocator(std::move(taken_paths));
  std::vector<std::vector<XlsbTablePart>> sheet_tables(sheet_count);
  std::vector<std::vector<std::string>> sheet_hyperlink_rids;
  sheet_hyperlink_rids.reserve(sheet_count);
  std::vector<std::string> sheet_drawing_rids(sheet_count);
  std::vector<bool> sheet_owns_drawing(sheet_count, false);
  const std::uint32_t views = workbook_view_count(workbook.book_views_xml());
  // One folder for the sheet bodies (display) and the sheet rels (target).
  const std::string_view book_directory = workbook.location() && workbook.location()->directory
                                              ? std::string_view(*workbook.location()->directory)
                                              : std::string_view{};
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const NameTable sheet_name_table = BuildNameTableForScope(workbook, ordered_names, static_cast<std::int32_t>(i));
    std::unordered_set<std::string> taken_rel_ids;
    for (const UnknownRelationship& rel : workbook.sheet(i).unknown_relationships()) {
      taken_rel_ids.insert(rel.id);
    }
    sheet_hyperlink_rids.push_back(hyperlink_relationship_ids(workbook.sheet(i)));
    for (const std::string& rel_id : sheet_hyperlink_rids.back()) {
      taken_rel_ids.insert(rel_id);
    }
    std::uint32_t tables_lost = 0;
    sheet_tables[i] =
        encode_sheet_tables(workbook, i, FeatureFormulaWriteContext{sheet_names, sheet_ranges, sheet_name_table},
                            taken_rel_ids, table_allocator, &tables_lost);
    ReportDeferred(&diagnostics.deferred_feature_count, "tables", tables_lost, i);
    std::vector<std::string> table_rel_ids;
    for (const XlsbTablePart& table : sheet_tables[i]) {
      table_rel_ids.push_back(table.rel_id);
      taken_rel_ids.insert(table.rel_id);
    }
    // The drawing id avoids retained ids, a suppressed drawing's included.
    sheet_owns_drawing[i] = !workbook.sheet(i).drawing_rel_target().empty();
    if (sheet_owns_drawing[i] && DrawingIsLinkable(workbook, workbook.sheet(i), i, &diagnostics)) {
      sheet_drawing_rids[i] = FreshRelId(taken_rel_ids);
    }
    std::uint32_t sparklines_lost = 0;
    auto sheet_body_or =
        emit_sheet(workbook.sheet(i), sst, sheet_names, sheet_ranges, sheet_name_table, sheet_hyperlink_rids[i],
                   &downgraded_formula_count, metadata.ifmd, name_shapes(workbook, i), table_rel_ids, views,
                   &sparklines_lost, sheet_drawing_rids[i], sheet_owns_drawing[i],
                   metadata.rich.empty() ? nullptr : &metadata.rich.vm_by_sheet[i], book_directory);
    ReportDeferred(&diagnostics.deferred_feature_count, "sparklines", sparklines_lost, i);
    if (!sheet_body_or) {
      return std::move(sheet_body_or.error());
    }
    sheet_bodies.push_back(std::move(sheet_body_or.value()));
  }
  const bool emit_sst_part = !sst.empty();

  auto plan_or = BuildEmissionPlan(workbook, emit_sst_part, metadata.generate_part, !metadata.rich.empty(),
                                   sheet_ranges.links.size(), &diagnostics);
  if (!plan_or) {
    return std::move(plan_or.error());
  }
  EmissionPlan plan = plan_or.take();
  plan.sheet_tables = std::move(sheet_tables);
  plan.sheet_drawing_rids = std::move(sheet_drawing_rids);
  plan.sheet_owns_drawing = std::move(sheet_owns_drawing);
  plan.book_directory = std::string(book_directory);
  const std::vector<const ExternalLinkRecord*> links = written_external_links(workbook);
  std::vector<std::string> link_rel_ids;
  link_rel_ids.reserve(links.size());
  for (std::size_t i = 0; i < links.size(); ++i) {
    link_rel_ids.push_back("rId" + std::to_string(sheet_count + i + 1U));
  }

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", kWriteXlsbContext);
  }

  // 1. [Content_Types].xml
  RETURN_IF_ERROR(AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(workbook, plan)));
  // 2. _rels/.rels
  RETURN_IF_ERROR(
      AddPart(writer.get(), "_rels/.rels",
              BuildPackageRels(workbook, kWorkbookBinPath, plan.passthrough_kept, plan.passthrough_path_index,
                               "xlsb.writer.package_rel_skipped", &diagnostics)));
  // 3. xl/_rels/workbook.bin.rels
  // Passthrough parts (styles.bin, theme, metadata, sharedStrings) are only
  // discoverable through workbook relationships; BuildWorkbookRels emits a rel
  // for each part actually present so none dangle on reload.
  std::vector<std::pair<std::string, std::string>> workbook_rel_ids;
  RETURN_IF_ERROR(AddPart(
      writer.get(), kWorkbookBinRelsPath,
      BuildWorkbookRels(sheet_count, emit_sst_part, plan, workbook, link_rel_ids, &diagnostics, &workbook_rel_ids)));
  // 4. xl/workbook.bin
  {
    ASSIGN_OR_RETURN(auto wb_bytes, BuildWorkbookBin(workbook, ordered_names, sheet_ranges, sheet_names, link_rel_ids));
    splice_book_slicer_cache_ids(wb_bytes, workbook.xlsb_book_frt(), workbook_rel_ids);
    RETURN_IF_ERROR(AddPartBytes(writer.get(), kWorkbookBinPath, wb_bytes));
  }
  // 4b. xl/styles.bin, when the source was not already an XLSB package with
  // an opaque style payload to preserve.  Its relationship is emitted above
  // in lockstep with this part.
  if (plan.has_generated_styles) {
    const std::vector<std::uint8_t> styles_bytes = write_styles_bin(workbook.styles());
    RETURN_IF_ERROR(AddPartBytes(writer.get(), kStylesBinPath, styles_bytes));
  }
  // 4c. xl/metadata.bin for dynamic-array spill anchors and rich errors, with
  // the rich values themselves. The worksheet BrtCellMeta / BrtValueMeta
  // records and workbook relationships are emitted in lockstep.
  if (plan.has_generated_metadata) {
    const std::vector<std::uint8_t> metadata_bytes =
        build_metadata_bin(metadata.dynamic_arrays, metadata.rich.values.size());
    RETURN_IF_ERROR(AddPartBytes(writer.get(), kMetadataBinPath, metadata_bytes));
  }
  if (plan.has_generated_rich_values) {
    RETURN_IF_ERROR(AddPart(writer.get(), kRdRichValuePartPath, serialize_rdrichvalue(metadata.rich)));
    RETURN_IF_ERROR(
        AddPart(writer.get(), kRdRichValueStructurePartPath, serialize_rdrichvaluestructure(metadata.rich)));
    RETURN_IF_ERROR(AddPart(writer.get(), kRdRichValueTypesPartPath, std::string(rich_value_types_xml())));
  }
  // 5. xl/worksheets/sheet<N>.bin, plus its rels when the sheet's retained
  // records reference other parts (hyperlinks, drawings, table definitions).
  for (std::size_t i = 0; i < sheet_count; ++i) {
    RETURN_IF_ERROR(AddPartBytes(writer.get(), sheet_part_path(i + 1U), sheet_bodies[i]));
    for (const XlsbTablePart& table : plan.sheet_tables[i]) {
      RETURN_IF_ERROR(AddPartBytes(writer.get(), table.path, table.bytes));
    }
    const std::string sheet_rels =
        BuildSheetRels(workbook.sheet(i), plan, plan.sheet_tables[i], sheet_hyperlink_rids[i],
                       plan.sheet_owns_drawing[i], plan.sheet_drawing_rids[i], &diagnostics);
    if (!sheet_rels.empty()) {
      RETURN_IF_ERROR(AddPart(writer.get(), sheet_rels_part_path(i + 1U), sheet_rels));
    }
  }
  // 5b. xl/externalLinks/externalLink<N>.bin and its rels, one per link, in
  // the order `BrtSupBookSrc` lists them.
  for (std::size_t i = 0; i < links.size(); ++i) {
    const std::vector<std::uint8_t> link_bytes = build_external_link_bin(*links[i], sheet_ranges.links[i]);
    RETURN_IF_ERROR(AddPartBytes(writer.get(), external_link_part_path(i + 1U), link_bytes));
    RETURN_IF_ERROR(AddPart(writer.get(), ExternalLinkRelsPath(i + 1U), build_external_link_rels(*links[i])));
  }
  // 6. xl/sharedStrings.bin (conditional)
  if (emit_sst_part) {
    ASSIGN_OR_RETURN(auto sst_body, emit_sst(sst));
    RETURN_IF_ERROR(AddPartBytes(writer.get(), kSharedStringsBinPath, sst_body));
  }
  // 7. Passthrough parts. Chart caches are refreshed into replacement bytes;
  // the model is untouched.
  ASSIGN_OR_RETURN(auto refreshed, refreshed_chart_parts(workbook));
  const io::PassthroughPartIndex refreshed_index(refreshed);
  for (const PassthroughPart* part : plan.passthrough_kept) {
    // Never emit a traversal-shaped part name, even if one reached the model
    // through a path other than the reader (which already rejects them).
    if (!ooxml::is_safe_part_name(part->path)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "passthrough part name escapes package root; refusing to write",
                        std::string(kWriteXlsbContext) + " part=" + part->path);
    }
    // Retained bytes must not silently drop a mutation made since load.
    if (auto fresh = CheckRetainedPartFreshness(workbook, *part); !fresh) {
      return std::move(fresh.error());
    }
    const PassthroughPart* replacement = refreshed_index.find(part->path);
    const auto& part_bytes = replacement != nullptr ? replacement->bytes : part->bytes;
    RETURN_IF_ERROR(AddPartBytes(writer.get(), part->path, part_bytes));
  }

  ASSIGN_OR_RETURN(auto bytes, FinalizeArchive(writer, kWriteXlsbContext));
  diagnostics.downgraded_formula_count = downgraded_formula_count;
  return XlsbWriteResult{std::move(bytes), diagnostics};
}

Expected<std::vector<std::uint8_t>, Error> write_xlsb(const Workbook& workbook) {
  ASSIGN_OR_RETURN(auto result, write_xlsb_with_result(workbook));
  return std::move(result.bytes);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
