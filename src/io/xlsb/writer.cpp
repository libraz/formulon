//
// Implementation of the MS-XLSB package writer. See `io/xlsb/writer.h`
// for the contract. The implementation mirrors the structure of
// `io/ooxml_writer.cpp`: build an emission plan (which sheet owns
// which numeric id, which passthrough parts survive collision
// detection), then compose the parts and pipe them through miniz.

#include "io/xlsb/writer.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "calc_settings.h"
#include "cell.h"
#include "cf/cf_types.h"
#include "default_content_type.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/relationship_writer.h"
#include "io/ooxml/workbook_xml_builder.h"
#include "io/ooxml/zip_part_writer.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_targets.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/styles_writer.h"
#include "io/xlsb/workbook_bin_writer.h"
#include "io/xml_escape.h"
#include "miniz.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// ---------------------------------------------------------------------------
// Constants
// ---------------------------------------------------------------------------

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

// The reader accepts both `application/vnd.ms-excel.sheet.binary.macroEnabled.main`
// (used by `.xlsm` and the `.xlsb` corpus xlwings emits on macOS) and
// `application/vnd.ms-excel.sheet.macroEnabled.main` (the alternative
// some non-macro xlsb writers ship). We emit the first form because
// (a) it's what the reader's primary fixture uses and (b) it's the
// content type Excel for Mac actually writes — the other variant is
// accepted for compatibility on input only.
constexpr std::string_view kCtPackageRels = "application/vnd.openxmlformats-package.relationships+xml";
constexpr std::string_view kCtXml = "application/xml";
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
// Worksheet part content type. This is the worksheet data part, NOT the
// binary-index part (`application/vnd.ms-excel.binIndexWs`); mislabelling it as
// the index type makes Excel treat the package as having zero worksheets and
// reject it (-50).
constexpr std::string_view kCtWorksheetXlsb = "application/vnd.ms-excel.worksheet";
constexpr std::string_view kCtSharedStringsXlsb = "application/vnd.ms-excel.sharedStrings";
constexpr std::string_view kCtStylesXlsb = "application/vnd.ms-excel.styles";
constexpr std::string_view kCtSheetMetadataXlsb = "application/vnd.ms-excel.sheetMetadata";

// ---------------------------------------------------------------------------
// Emission plan: where do passthrough parts land, do any collide?
// ---------------------------------------------------------------------------

struct EmissionPlan {
  std::vector<const PassthroughPart*> passthrough_kept;
  /// Source Default registrations, lower-case by extension and unique. They
  /// remain available even when a source archive contains no current part of
  /// a given extension: keeping the registry avoids changing the meaning of
  /// any Default-typed part a caller may add before the next save.
  std::vector<DefaultContentType> default_content_types;
  /// When a source `bin` Default describes an embedded OLE/VBA payload (or
  /// any other non-workbook binary), the generated workbook must override
  /// that Default with the canonical XLSB workbook type.
  bool workbook_bin_override = false;
  bool has_text_cells = false;  // gates emission of xl/sharedStrings.bin
  bool has_generated_styles = false;
  bool has_generated_dynamic_metadata = false;
};

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

bool IsRepresentableBaseColumnWidth(double value) {
  return std::isfinite(value) && value >= 0.0 && value <= 255.0 && std::floor(value) == value;
}

bool IsRepresentableDefaultColumnWidth(double value) {
  constexpr double kMaxDefaultColumnWidth = 65535.0 / 256.0;
  return std::isfinite(value) && value >= 0.0 && value <= kMaxDefaultColumnWidth;
}

bool IsRepresentableDefaultRowHeight(double value) {
  constexpr double kMaxDefaultRowHeight = 65535.0 / 20.0;
  if (!std::isfinite(value) || value < 0.0 || value > kMaxDefaultRowHeight) {
    return false;
  }
  return std::isfinite(std::round(value * 20.0));
}

/// Fails a retained part (pivot table, pivot cache, styles) whose model
/// state changed since load, including the model object having gone.
/// Untracked passthrough parts always pass.
Expected<void, Error> CheckRetainedPartFreshness(const Workbook& workbook, const PassthroughPart& part) {
  if (part.retained_origin == PassthroughPart::RetainedOrigin::kNone) {
    return Expected<void, Error>::Ok();
  }
  const std::optional<std::uint64_t> current = current_retained_part_fingerprint(workbook, part);
  if (!current.has_value() || !part.model_fingerprint.has_value() || *current != *part.model_fingerprint) {
    return make_error(FormulonErrorCode::kIoXlsbRetainedPartStale,
                      "retained XLSB part no longer matches the model state it was loaded from",
                      "context=write_xlsb part=" + part.path);
  }
  return Expected<void, Error>::Ok();
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
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    const Sheet& sheet = workbook.sheet(i);
    ReportDeferred(&count, "auto_filter", sheet.has_auto_filter() ? 1U : 0U, i);
    ReportDeferred(&count, "comments", sheet.comments().size(), i);
    ReportDeferred(&count, "threaded_comments", sheet.threaded_comments().size(), i);
    // A drawing read from (or inserted into) an OOXML package has no
    // binary relationship here; only `.xlsb`-retained drawings survive.
    ReportDeferred(&count, "drawings", sheet.drawing_rel_target().empty() ? 0U : 1U, i);
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
  }
  if (!workbook.tables().empty()) {
    ReportDeferred(&count, "tables", workbook.tables().size(), 0U);
  }
  // Workbook-level calc settings: this writer emits no calc-properties
  // record at all, so a non-default mode or an enabled iterative solve
  // silently reverts to automatic / non-iterative recalculation on open.
  // lockWindows / lockRevision / a revisions password have no measured
  // BrtBookProtection field and are left out of the saved workbook.
  std::vector<std::uint8_t> protection_scratch;
  const auto book_protection = emit_book_protection(protection_scratch, workbook.workbook_protection_xml());
  ReportDeferred(&count, "workbook_protection", book_protection && !book_protection.value() ? 1U : 0U, 0U);
  ReportDeferred(&count, "calc_mode", workbook.calc_mode() == Workbook::CalcMode::kAuto ? 0U : 1U, 0U);
  ReportDeferred(&count, "iterative_calc", workbook.iterative_options().enabled ? 1U : 0U, 0U);
  return count;
}

// True for a worksheet binary-index part (`xl/worksheets/binaryIndex<N>.bin`).
// These describe the original sheet bodies' row layout and are stale once we
// regenerate the sheets; Excel opens without them.
bool IsBinaryIndexPart(const std::string& path) {
  constexpr std::string_view kPrefix = "xl/worksheets/binaryIndex";
  return path.size() > kPrefix.size() && path.compare(0, kPrefix.size(), kPrefix) == 0;
}

// `xl/metadata.xml` is an XLSX-only dynamic-array metadata part.  Its XLSB
// counterpart is `xl/metadata.bin`; copying XML bytes into a binary workbook
// makes Excel repair the file on open.  Dynamic-array metadata that the
// model can represent is regenerated below; other XLSX-only metadata remains
// deliberately out of the XLSB package.
bool IsXlsxOnlyMetadataPart(const std::string& path) {
  return path == "xl/metadata.xml";
}

bool HasRawStylesPart(const Workbook& wb) {
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == "xl/styles.bin")
      return true;
  }
  return false;
}

bool HasModelledStyles(const Workbook& wb) {
  const StylesTable& styles = wb.styles();
  // An empty table is the unstyled shape a package without a styles part
  // parses into; both workbook factories seed one.  Any populated collection,
  // including an explicitly-created default XF, needs a styles part because
  // worksheet iStyleRef values resolve only through its relationship.
  return !styles.fonts.empty() || !styles.fills.empty() || !styles.borders.empty() || !styles.num_fmts.empty() ||
         !styles.cell_xfs.empty() || !styles.cell_style_xfs.empty() || !styles.cell_styles.empty() ||
         !styles.dxfs.empty();
}

bool HasDynamicArrayMetadata(const Workbook& wb) {
  return has_dynamic_array_formula(wb);
}

// Whether a dynamic-array metadata part is needed, and which cell-metadata
// entry a spill anchor's `BrtCellMeta` record must name.
//
// `ifmd` is 1-based, and 0 means no anchor may emit a `BrtCellMeta` record at
// all: either the workbook has no spill anchors, or the metadata part that
// ships is a retained passthrough whose dynamic-array entry could not be
// identified. Both the index and the part are decided here, once, so the
// worksheet bodies and the package can never disagree about which entry a
// cell names.
struct DynamicArrayMetadataPlan {
  bool generate_part = false;
  std::uint32_t ifmd = 0;
};

DynamicArrayMetadataPlan BuildDynamicArrayMetadataPlan(const Workbook& wb) {
  DynamicArrayMetadataPlan plan;
  if (!HasDynamicArrayMetadata(wb)) {
    return plan;
  }
  // The first passthrough part wins the path; `BuildEmissionPlan` drops any
  // later duplicate, so that is the one whose numbering ships.
  const PassthroughPart* retained = nullptr;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (part.path == "xl/metadata.bin") {
      retained = &part;
      break;
    }
  }
  if (retained == nullptr) {
    plan.generate_part = true;
    plan.ifmd = 1U;  // The generated part declares exactly one entry.
    return plan;
  }
  // A retained part carries its own numbering, and its first entry is not
  // necessarily the dynamic-array one — it may hold rich-value or
  // cube-function metadata instead, or no dynamic-array type at all.
  plan.ifmd = find_dynamic_array_cell_meta_index(ByteSpan{retained->bytes.data(), retained->bytes.size()});
  return plan;
}

std::unordered_set<std::string> BuildGeneratedPathSet(const Workbook& wb, bool emit_sst_part, bool emit_styles_part,
                                                      bool emit_dynamic_metadata_part) {
  std::unordered_set<std::string> paths;
  paths.insert("[Content_Types].xml");
  paths.insert("_rels/.rels");
  paths.insert("xl/workbook.bin");
  paths.insert("xl/_rels/workbook.bin.rels");
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    paths.insert("xl/worksheets/sheet" + std::to_string(i + 1) + ".bin");
    paths.insert("xl/worksheets/_rels/sheet" + std::to_string(i + 1) + ".bin.rels");
  }
  if (emit_sst_part) {
    paths.insert("xl/sharedStrings.bin");
  }
  if (emit_styles_part) {
    paths.insert("xl/styles.bin");
  }
  if (emit_dynamic_metadata_part) {
    paths.insert("xl/metadata.bin");
  }
  return paths;
}

Expected<EmissionPlan, Error> BuildEmissionPlan(const Workbook& wb, bool sst_present, bool generate_dynamic_metadata,
                                                WriteDiagnostics* diagnostics) {
  EmissionPlan plan;
  plan.has_text_cells = sst_present;
  // Existing XLSB packages retain their original styles bytes verbatim.  This
  // protects style features which are not represented by StylesTable yet
  // (notably differential formats).  XLSX/native workbooks have no raw part,
  // so their modelled style table becomes a fresh styles.bin instead.
  plan.has_generated_styles = !HasRawStylesPart(wb) && HasModelledStyles(wb);
  // Decided by `BuildDynamicArrayMetadataPlan`, which the sheet bodies were
  // already emitted against.
  plan.has_generated_dynamic_metadata = generate_dynamic_metadata;

  const std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, sst_present, plan.has_generated_styles, plan.has_generated_dynamic_metadata);

  // The reader rejects conflicting defaults, but Workbook is also a public
  // construction surface. Validate that hand-built workbooks cannot produce
  // ambiguous extension semantics on write.
  std::unordered_map<std::string, std::string> source_defaults;
  source_defaults.reserve(wb.default_content_types().size());
  for (const DefaultContentType& source : wb.default_content_types()) {
    if (source.extension.empty() || source.content_type.empty()) {
      continue;
    }
    const std::string extension = ooxml::lowercase_extension(source.extension);
    auto [it, inserted] = source_defaults.emplace(extension, source.content_type);
    if (!inserted && it->second != source.content_type) {
      return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                        "workbook has conflicting Default content types for extension " + extension,
                        "context=write_xlsb extension=" + extension);
    }
  }

  std::unordered_set<std::string> kept_paths;
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (!ooxml::is_safe_part_name(part.path)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "passthrough part name escapes package root; refusing to write",
                        "context=write_xlsb part=" + part.path);
    }
    if (generated.count(part.path) != 0U) {
      StructuredLog("xlsb.writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("generated_path_wins"))
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_part_count;
      }
      continue;
    }
    // Drop parts that would be stale or orphaned after we regenerate the
    // worksheets. The binary-index parts (`xl/worksheets/binaryIndex*.bin`)
    // describe the ROW layout of the original sheet bodies; once we re-emit
    // sheets they no longer match, and Excel opens fine without them. The
    // calc chain likewise references the original formula graph; a stale
    // chain is rejected, and Excel rebuilds it on load when absent (same
    // policy as the OOXML writer). Neither carries a workbook relationship
    // here, so keeping them would leave dangling / mismatched parts.
    if (IsBinaryIndexPart(part.path) || part.path == "xl/calcChain.bin" || IsXlsxOnlyMetadataPart(part.path)) {
      if (IsXlsxOnlyMetadataPart(part.path)) {
        StructuredLog("xlsb.writer.deferred")
            .field("kind", std::string_view("xlsx_metadata"))
            .field("count", static_cast<std::int64_t>(1))
            .warn();
      }
      continue;
    }
    if (!kept_paths.insert(part.path).second) {
      StructuredLog("xlsb.writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("duplicate_passthrough_path"))
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_part_count;
      }
      continue;
    }
    if (part.content_type.empty()) {
      const std::string extension = ooxml::extension_of_part(part.path);
      const auto default_it = source_defaults.find(extension);
      if (extension.empty() || default_it == source_defaults.end()) {
        return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                          "Default-typed passthrough part has no matching Default registration",
                          "context=write_xlsb part=" + part.path + " extension=" + extension);
      }
    }
    plan.passthrough_kept.push_back(&part);
  }
  for (const auto& [extension, content_type] : source_defaults) {
    plan.default_content_types.push_back(DefaultContentType{extension, content_type});
    if (extension == "bin" && content_type != kCtWorkbookXlsb) {
      plan.workbook_bin_override = true;
    }
  }
  sort_by_index(plan.default_content_types, [](const DefaultContentType& lhs, const DefaultContentType& rhs) {
    return lhs.extension < rhs.extension;
  });
  return plan;
}

// ---------------------------------------------------------------------------
// XML part builders: the package envelope is XML even in xlsb.
// ---------------------------------------------------------------------------

std::string BuildContentTypes(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(512 + wb.sheet_count() * 128 + plan.passthrough_kept.size() * 128);
  out.append(kXmlDecl);
  out.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n");
  auto append_default = [&out](std::string_view extension, std::string_view content_type) {
    out.append("  <Default Extension=\"");
    AppendXmlAttrEscaped(out, extension);
    out.append("\" ContentType=\"");
    AppendXmlAttrEscaped(out, content_type);
    out.append("\"/>\n");
  };
  auto append_override = [&out](std::string_view path, std::string_view content_type) {
    out.append("  <Override PartName=\"/");
    AppendXmlAttrEscaped(out, path);
    out.append("\" ContentType=\"");
    AppendXmlAttrEscaped(out, content_type);
    out.append("\"/>\n");
  };
  std::string_view rels_default = kCtPackageRels;
  std::string_view xml_default = kCtXml;
  // `bin` is commonly registered as the workbook type, but a real package
  // may use the same extension for OLE/VBA payloads. Preserve the source
  // Default when a kept binary passthrough needs it, and add a canonical
  // workbook Override so `xl/workbook.bin` remains unambiguous.
  std::string_view bin_default = kCtWorkbookXlsb;
  for (const DefaultContentType& def : plan.default_content_types) {
    if (def.extension == "rels") {
      rels_default = def.content_type;
    } else if (def.extension == "xml") {
      xml_default = def.content_type;
    } else if (def.extension == "bin") {
      bin_default = def.content_type;
    }
  }
  append_default("rels", rels_default);
  append_default("xml", xml_default);
  append_default("bin", bin_default);
  for (const DefaultContentType& def : plan.default_content_types) {
    if (def.extension == "rels" || def.extension == "xml" || def.extension == "bin") {
      continue;
    }
    append_default(def.extension, def.content_type);
  }
  if (plan.workbook_bin_override) {
    out.append("  <Override PartName=\"/xl/workbook.bin\" ContentType=\"");
    AppendXmlAttrEscaped(out, kCtWorkbookXlsb);
    out.append("\"/>\n");
  }
  // A source `rels` Default may use a vendor-specific type. Generated
  // relationship parts still need their canonical OPC semantics, so pin
  // those generated paths with Overrides while leaving the source Default
  // available for retained opaque parts of the same extension. The special
  // `[Content_Types].xml` part intentionally remains Default-typed when the
  // source registry uses a noncanonical `xml` type; adding an Override for
  // that part would change the source registry's meaning.
  if (rels_default != kCtPackageRels) {
    append_override("_rels/.rels", kCtPackageRels);
    append_override("xl/_rels/workbook.bin.rels", kCtPackageRels);
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      bool has_emitted_sheet_rels = std::any_of(wb.sheet(i).hyperlinks().begin(), wb.sheet(i).hyperlinks().end(),
                                                [](const Hyperlink& hyperlink) { return !hyperlink.target.empty(); });
      for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
        if (rel.target_external) {
          has_emitted_sheet_rels = true;
          break;
        }
        if (std::any_of(plan.passthrough_kept.begin(), plan.passthrough_kept.end(),
                        [&rel](const PassthroughPart* part) { return part->path == rel.target; })) {
          has_emitted_sheet_rels = true;
          break;
        }
      }
      if (has_emitted_sheet_rels) {
        append_override("xl/worksheets/_rels/sheet" + std::to_string(i + 1U) + ".bin.rels", kCtPackageRels);
      }
    }
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    out.append("  <Override PartName=\"/xl/worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".bin\" ContentType=\"");
    out.append(kCtWorksheetXlsb);
    out.append("\"/>\n");
  }
  if (plan.has_text_cells) {
    out.append("  <Override PartName=\"/xl/sharedStrings.bin\" ContentType=\"");
    out.append(kCtSharedStringsXlsb);
    out.append("\"/>\n");
  }
  if (plan.has_generated_dynamic_metadata) {
    out.append("  <Override PartName=\"/xl/metadata.bin\" ContentType=\"");
    out.append(kCtSheetMetadataXlsb);
    out.append("\"/>\n");
  }
  if (plan.has_generated_styles) {
    out.append("  <Override PartName=\"/xl/styles.bin\" ContentType=\"");
    out.append(kCtStylesXlsb);
    out.append("\"/>\n");
  }
  // Passthrough overrides: only for entries that carried an explicit
  // ContentType in the source archive. Default-typed parts (empty
  // content_type) must NOT appear as Overrides.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part->content_type.empty()) {
      continue;
    }
    out.append("  <Override PartName=\"/");
    AppendXmlAttrEscaped(out, part->path);
    out.append("\" ContentType=\"");
    AppendXmlAttrEscaped(out, part->content_type);
    out.append("\"/>\n");
  }
  out.append("</Types>\n");
  return out;
}

// Builds `xl/worksheets/_rels/sheet<N>.bin.rels`, or an empty string when the
// sheet has nothing to relate to.
//
// Retained relationship ids remain verbatim so opaque worksheet-tail records
// (drawings, tables, and similar features) continue to resolve. Model-owned
// BrtHLink records are emitted separately and get ids from the same collision
// set; a relationship whose target part is not in the package is dropped so
// the emitted rels never dangle.
std::string BuildSheetRels(const Sheet& sheet, const EmissionPlan& plan, WriteDiagnostics* diagnostics) {
  std::string entries;
  for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
    if (!rel.target_external && !HasPassthroughPart(plan.passthrough_kept, rel.target)) {
      StructuredLog("xlsb.writer.sheet_rel_skipped")
          .field("reason", std::string_view("target_part_absent"))
          .field("type", rel.type)
          .field("target", rel.target)
          .warn();
      // Counted alongside the package- and workbook-scope drops: a dropped
      // sheet relationship is a dropped relationship, and OOXML having no
      // sheet-scope equivalent must not make it invisible.
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_relationship_count;
      }
      continue;
    }
    const std::string target = rel.target_external ? rel.target : TargetRelativeToWorksheet(rel.target);
    entries.append("  <Relationship Id=\"");
    AppendXmlAttrEscaped(entries, rel.id);
    entries.append("\" Type=\"");
    AppendXmlAttrEscaped(entries, rel.type);
    entries.append("\" Target=\"");
    AppendXmlAttrEscaped(entries, target);
    if (rel.target_external) {
      entries.append("\" TargetMode=\"External\"/>\n");
    } else {
      entries.append("\"/>\n");
    }
  }
  const std::vector<std::string> hyperlink_rids = hyperlink_relationship_ids(sheet);
  std::unordered_set<std::string> emitted_hyperlink_rids;
  emitted_hyperlink_rids.reserve(hyperlink_rids.size());
  for (std::size_t i = 0; i < sheet.hyperlinks().size(); ++i) {
    const Hyperlink& hyperlink = sheet.hyperlinks()[i];
    if (hyperlink.target.empty() || i >= hyperlink_rids.size() || hyperlink_rids[i].empty()) {
      continue;
    }
    if (!emitted_hyperlink_rids.insert(hyperlink_rids[i]).second) {
      continue;
    }
    entries.append("  <Relationship Id=\"");
    AppendXmlAttrEscaped(entries, hyperlink_rids[i]);
    entries.append("\" Type=\"");
    AppendXmlAttrEscaped(entries, kRelHyperlink);
    entries.append("\" Target=\"");
    AppendXmlAttrEscaped(entries, hyperlink.target);
    entries.append("\" TargetMode=\"External\"/>\n");
  }
  if (entries.empty()) {
    return {};
  }
  std::string out;
  out.reserve(entries.size() + 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  out.append(entries);
  out.append("</Relationships>\n");
  return out;
}

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst, const EmissionPlan& plan, const Workbook& wb,
                              WriteDiagnostics* diagnostics) {
  std::string out;
  out.reserve(256 + sheet_count * 192 + wb.unknown_workbook_rels().size() * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::uint32_t next_rid = 1;
  for (std::size_t i = 0; i < sheet_count; ++i) {
    out.append("  <Relationship Id=\"rId");
    out.append(std::to_string(next_rid++));
    out.append("\" Type=\"");
    out.append(kRelWorksheet);
    out.append("\" Target=\"worksheets/sheet");
    out.append(std::to_string(i + 1));
    out.append(".bin\"/>\n");
  }
  if (emit_sst) {
    AppendRelationship(out, next_rid++, kRelSharedStrings, "sharedStrings.bin", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  // The styles / theme / metadata parts ride the passthrough path, but the
  // reader (and Excel) locate them only through these workbook relationships.
  // Without them a styled cell's iStyleRef dangles against an empty style
  // table, and the theme / metadata parts are treated as orphans. Emit a rel
  // for each part that is actually present in the package.
  if (plan.has_generated_styles || HasPassthroughPart(plan.passthrough_kept, "xl/styles.bin")) {
    AppendRelationship(out, next_rid++, kRelStyles, "styles.bin", /*target_external=*/false, /*escape_target=*/true);
  }
  if (HasPassthroughPart(plan.passthrough_kept, "xl/theme/theme1.xml")) {
    AppendRelationship(out, next_rid++, kRelTheme, "theme/theme1.xml", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  if (plan.has_generated_dynamic_metadata || HasPassthroughPart(plan.passthrough_kept, "xl/metadata.bin")) {
    AppendRelationship(out, next_rid++, kRelSheetMetadata, "metadata.bin", /*target_external=*/false,
                       /*escape_target=*/true);
  }
  // Preserve relationships to raw XLSB parts that the reader does not model
  // (drawings, VBA, custom XML, etc.).  Internal targets are stored as
  // package paths on Workbook, while workbook relationships are relative to
  // `xl/`.  Do not duplicate relationships which the generated package
  // already owns above.
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (!rel.target_external && !HasPassthroughPart(plan.passthrough_kept, rel.target)) {
      StructuredLog("xlsb.writer.workbook_rel_skipped")
          .field("reason", std::string_view("target_part_absent"))
          .field("type", rel.type)
          .field("target", rel.target)
          .warn();
      if (diagnostics != nullptr) {
        ++diagnostics->dropped_relationship_count;
      }
      continue;
    }
    if (!rel.target_external && ((rel.type == kRelTheme && rel.target == "xl/theme/theme1.xml") ||
                                 (rel.type == kRelSheetMetadata && rel.target == "xl/metadata.bin"))) {
      continue;
    }
    const std::string target = rel.target_external ? rel.target : std::string(WithoutXlPrefix(rel.target));
    AppendRelationship(out, next_rid++, rel.type, target, rel.target_external, /*escape_target=*/true);
  }
  out.append("</Relationships>\n");
  return out;
}

}  // namespace

Expected<XlsbWriteResult, Error> write_xlsb_with_result(const Workbook& workbook) {
  const std::size_t sheet_count = workbook.sheet_count();
  if (sheet_count == 0) {
    return make_error(FormulonErrorCode::kInvalidArgument, "workbook has zero sheets", "context=write_xlsb");
  }

  // Pre-pass: emit each sheet body so we know whether the SST will be
  // non-empty. We hold the resulting bytes until after we write the
  // envelope so the order of `mz_zip_writer_add_mem` calls matches
  // what the reader expects (it does not, but we keep symmetry with
  // `write_ooxml`).
  // Ordered sheet-name list: the Ptg encoder maps a qualified
  // reference's sheet to its 0-based `ixti` through this list.
  std::vector<std::string> sheet_names;
  sheet_names.reserve(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    sheet_names.push_back(workbook.sheet(i).name());
  }

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
    return sheet_ranges_or.error();
  }
  const SheetRangeTable& sheet_ranges = sheet_ranges_or.value();

  SstBuilder sst;
  const DynamicArrayMetadataPlan dynamic_array = BuildDynamicArrayMetadataPlan(workbook);
  WriteDiagnostics diagnostics;
  std::uint32_t downgraded_formula_count = 0;
  diagnostics.deferred_feature_count = ReportDeferredSheetFeatures(workbook);
  std::vector<std::vector<std::uint8_t>> sheet_bodies;
  sheet_bodies.reserve(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    const NameTable sheet_name_table = BuildNameTableForScope(workbook, ordered_names, static_cast<std::int32_t>(i));
    auto sheet_body_or = emit_sheet(workbook.sheet(i), sst, sheet_names, sheet_ranges, sheet_name_table,
                                    &downgraded_formula_count, dynamic_array.ifmd, name_shapes(workbook, i));
    if (!sheet_body_or) {
      return sheet_body_or.error();
    }
    sheet_bodies.push_back(std::move(sheet_body_or.value()));
  }
  const bool emit_sst_part = !sst.empty();

  auto plan_or = BuildEmissionPlan(workbook, emit_sst_part, dynamic_array.generate_part, &diagnostics);
  if (!plan_or) {
    return plan_or.error();
  }
  const EmissionPlan plan = plan_or.take();

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_xlsb");
  }

  // 1. [Content_Types].xml
  if (auto r = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(workbook, plan)); !r) {
    return r.error();
  }
  // 2. _rels/.rels
  if (auto r = AddPart(writer.get(), "_rels/.rels",
                       BuildPackageRels(workbook, "xl/workbook.bin", plan.passthrough_kept,
                                        "xlsb.writer.package_rel_skipped", &diagnostics));
      !r) {
    return r.error();
  }
  // 3. xl/_rels/workbook.bin.rels
  // Passthrough parts (styles.bin, theme, metadata, sharedStrings) are only
  // discoverable through workbook relationships; BuildWorkbookRels emits a rel
  // for each part actually present so none dangle on reload.
  if (auto r = AddPart(writer.get(), "xl/_rels/workbook.bin.rels",
                       BuildWorkbookRels(sheet_count, emit_sst_part, plan, workbook, &diagnostics));
      !r) {
    return r.error();
  }
  // 4. xl/workbook.bin
  {
    auto wb_bytes_or = BuildWorkbookBin(workbook, ordered_names, sheet_ranges, sheet_names);
    if (!wb_bytes_or) {
      return wb_bytes_or.error();
    }
    if (auto r = AddPartBytes(writer.get(), "xl/workbook.bin", wb_bytes_or.value()); !r) {
      return r.error();
    }
  }
  // 4b. xl/styles.bin, when the source was not already an XLSB package with
  // an opaque style payload to preserve.  Its relationship is emitted above
  // in lockstep with this part.
  if (plan.has_generated_styles) {
    const std::vector<std::uint8_t> styles_bytes = write_styles_bin(workbook.styles());
    if (auto r = AddPartBytes(writer.get(), "xl/styles.bin", styles_bytes); !r) {
      return r.error();
    }
  }
  // 4c. xl/metadata.bin for dynamic-array spill anchors. The worksheet
  // BrtCellMeta records and workbook relationship are emitted in lockstep.
  if (plan.has_generated_dynamic_metadata) {
    const std::vector<std::uint8_t> metadata_bytes = build_dynamic_array_metadata_bin();
    if (auto r = AddPartBytes(writer.get(), "xl/metadata.bin", metadata_bytes); !r) {
      return r.error();
    }
  }
  // 5. xl/worksheets/sheet<N>.bin, plus its rels when the sheet's retained
  // records reference other parts (hyperlinks, drawings, table definitions).
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::string path("xl/worksheets/sheet");
    path.append(std::to_string(i + 1));
    path.append(".bin");
    if (auto r = AddPartBytes(writer.get(), path, sheet_bodies[i]); !r) {
      return r.error();
    }
    const std::string sheet_rels = BuildSheetRels(workbook.sheet(i), plan, &diagnostics);
    if (!sheet_rels.empty()) {
      std::string rels_path("xl/worksheets/_rels/sheet");
      rels_path.append(std::to_string(i + 1));
      rels_path.append(".bin.rels");
      if (auto r = AddPart(writer.get(), rels_path, sheet_rels); !r) {
        return r.error();
      }
    }
  }
  // 6. xl/sharedStrings.bin (conditional)
  if (emit_sst_part) {
    auto sst_body_or = emit_sst(sst);
    if (!sst_body_or) {
      return sst_body_or.error();
    }
    if (auto r = AddPartBytes(writer.get(), "xl/sharedStrings.bin", sst_body_or.value()); !r) {
      return r.error();
    }
  }
  // 7. Passthrough parts.
  for (const PassthroughPart* part : plan.passthrough_kept) {
    // Never emit a traversal-shaped part name, even if one reached the model
    // through a path other than the reader (which already rejects them).
    if (!ooxml::is_safe_part_name(part->path)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "passthrough part name escapes package root; refusing to write",
                        "context=write_xlsb part=" + part->path);
    }
    // Retained bytes must not silently drop a mutation made since load.
    if (auto fresh = CheckRetainedPartFreshness(workbook, *part); !fresh) {
      return fresh.error();
    }
    if (auto r = AddPartBytes(writer.get(), part->path, part->bytes); !r) {
      return r.error();
    }
  }

  auto bytes_or = FinalizeArchive(writer, "context=write_xlsb");
  if (!bytes_or) {
    return bytes_or.error();
  }
  diagnostics.downgraded_formula_count = downgraded_formula_count;
  return XlsbWriteResult{std::move(bytes_or.value()), diagnostics};
}

Expected<std::vector<std::uint8_t>, Error> write_xlsb(const Workbook& workbook) {
  auto result_or = write_xlsb_with_result(workbook);
  if (!result_or) {
    return result_or.error();
  }
  return std::move(result_or.value().bytes);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
