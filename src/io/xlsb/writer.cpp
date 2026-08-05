//
// Implementation of the MS-XLSB package writer. See `io/xlsb/writer.h`
// for the contract. The implementation mirrors the structure of
// `io/ooxml_writer.cpp`: build an emission plan (which sheet owns
// which numeric id, which passthrough parts survive collision
// detection), then compose the parts and pipe them through miniz.

#include "io/xlsb/writer.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "io/ooxml/package_validator.h"
#include "io/passthrough_part.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_writer.h"
#include "io/xlsb/sst_writer.h"
#include "io/xlsb/styles_writer.h"
#include "io/xml_escape.h"
#include "miniz.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
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
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
// Worksheet part content type. This is the worksheet data part, NOT the
// binary-index part (`application/vnd.ms-excel.binIndexWs`); mislabelling it as
// the index type makes Excel treat the package as having zero worksheets and
// reject it (-50).
constexpr std::string_view kCtWorksheetXlsb = "application/vnd.ms-excel.worksheet";
constexpr std::string_view kCtSharedStringsXlsb = "application/vnd.ms-excel.sharedStrings";
constexpr std::string_view kCtStylesXlsb = "application/vnd.ms-excel.styles";
constexpr std::string_view kCtSheetMetadataXlsb = "application/vnd.ms-excel.sheetMetadata";

constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
constexpr std::string_view kRelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
constexpr std::string_view kRelTheme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
constexpr std::string_view kRelSheetMetadata =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata";
constexpr std::string_view kRelCoreProps =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
constexpr std::string_view kRelExtendedProps =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
constexpr std::string_view kRelCustomProps =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

// Workbook-globals record ids the reader does not consume (so they are absent
// from `XlsbRecordType`) but the writer must emit for a well-formed stream.
// A `BrtExternSheet` must live inside a `BrtBeginExternals ... BrtEndExternals`
// block with a `BrtSupSelf` (self-referencing supporting book); Excel rejects a
// bare `BrtExternSheet`. Ids per [MS-XLSB] §2.4.
constexpr std::uint16_t kBrtBeginExternals = 353;
constexpr std::uint16_t kBrtSupSelf = 357;
constexpr std::uint16_t kBrtEndExternals = 354;
// Workbook-globals structural records Excel expects before the sheet bundle.
constexpr std::uint16_t kBrtFileVersion = 128;
constexpr std::uint16_t kBrtBeginBookViews = 135;
constexpr std::uint16_t kBrtWbView = 158;
constexpr std::uint16_t kBrtEndBookViews = 136;

// ---------------------------------------------------------------------------
// Emission plan: where do passthrough parts land, do any collide?
// ---------------------------------------------------------------------------

struct EmissionPlan {
  std::vector<const PassthroughPart*> passthrough_kept;
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

std::uint32_t ReportDeferredSheetFeatures(const Workbook& workbook) {
  std::uint32_t count = 0U;
  for (std::size_t i = 0; i < workbook.sheet_count(); ++i) {
    const Sheet& sheet = workbook.sheet(i);
    ReportDeferred(&count, "conditional_formats", sheet.conditional_formats().size(), i);
    ReportDeferred(&count, "data_validations", sheet.validations().size(), i);
    ReportDeferred(&count, "hyperlinks", sheet.hyperlinks().size(), i);
    ReportDeferred(&count, "auto_filter", sheet.auto_filter_xml().empty() ? 0U : 1U, i);
    const SheetPrintSettings& print = sheet.print_settings();
    const bool has_print = !print.sheet_pr_xml.empty() || !print.page_margins_xml.empty() ||
                           !print.page_setup_xml.empty() || !print.print_options_xml.empty() ||
                           !print.header_footer_xml.empty() || !print.manual_row_breaks.empty() ||
                           !print.manual_col_breaks.empty();
    ReportDeferred(&count, "print_settings", has_print ? 1U : 0U, i);
  }
  if (!workbook.tables().empty()) {
    ReportDeferred(&count, "tables", workbook.tables().size(), 0U);
  }
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
  // An empty table is the intentionally unstyled `Workbook::create_empty()`
  // shape.  Any populated collection, including an explicitly-created default
  // XF, needs a styles part because worksheet iStyleRef values resolve only
  // through its relationship.
  return !styles.fonts.empty() || !styles.fills.empty() || !styles.borders.empty() || !styles.num_fmts.empty() ||
         !styles.cell_xfs.empty() || !styles.cell_style_xfs.empty() || !styles.cell_styles.empty();
}

bool HasDynamicArrayMetadata(const Workbook& wb) {
  for (std::size_t sheet_index = 0; sheet_index < wb.sheet_count(); ++sheet_index) {
    const Sheet& sheet = wb.sheet(sheet_index);
    for (const auto& [row, cells] : sheet.rows()) {
      for (std::uint32_t col = 0; col < cells.size(); ++col) {
        if (!cells[col].formula_text.empty()) {
          if (sheet.spill_region_at_anchor(row, col) != nullptr) {
            return true;
          }
        }
      }
    }
  }
  return false;
}

// Emits the XLDAPR metadata stream Excel uses to mark spilled dynamic-array
// anchors. The record sequence and constant GUID are the canonical dynamic
// array metadata shape Excel writes in XLSB workbooks. Cell records refer to
// its sole entry with BrtCellMeta index 1.
std::vector<std::uint8_t> BuildDynamicArrayMetadataBin() {
  constexpr std::uint16_t kBrtBeginMetadata = 332;
  constexpr std::uint16_t kBrtEndMetadata = 333;
  constexpr std::uint16_t kBrtBeginEsmdtinfo = 334;
  constexpr std::uint16_t kBrtMdtinfo = 335;
  constexpr std::uint16_t kBrtEndEsmdtinfo = 336;
  constexpr std::uint16_t kBrtBeginEsfmd = 337;
  constexpr std::uint16_t kBrtEndEsfmd = 338;
  constexpr std::uint16_t kBrtBeginEsfmdInfo = 339;
  constexpr std::uint16_t kBrtEndEsfmdInfo = 340;
  constexpr std::uint16_t kBrtBeginExt = 52;
  constexpr std::uint16_t kBrtBeginFRT = 35;
  constexpr std::uint16_t kBrtBeginDynamicArrayExt = 4096;
  constexpr std::uint16_t kBrtDynamicArrayProperties = 4097;
  constexpr std::uint16_t kBrtEndFRT = 36;
  constexpr std::uint16_t kBrtEndExt = 53;
  constexpr std::uint16_t kBrtCellMetaEntry = 51;

  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_record(out, kBrtBeginMetadata, ByteSpan{});

  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsmdtinfo, payload);
  payload.clear();
  emit_u32(payload, 0xD86AC0B0U);
  emit_u32(payload, 0x0001D4C0U);
  emit_xlwidestring(payload, "XLDAPR");
  emit_record(out, kBrtMdtinfo, payload);
  emit_record(out, kBrtEndEsmdtinfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, 1U);
  emit_xlwidestring(payload, "XLDAPR");
  emit_record(out, kBrtBeginEsfmdInfo, payload);
  emit_record(out, kBrtBeginExt, ByteSpan{});
  payload.clear();
  emit_u32(payload, 0x00020002U);
  emit_record(out, kBrtBeginFRT, payload);
  emit_record(out, kBrtBeginDynamicArrayExt, ByteSpan{});
  payload.clear();
  emit_u16(payload, 1U);
  emit_record(out, kBrtDynamicArrayProperties, payload);
  emit_record(out, kBrtEndFRT, ByteSpan{});
  emit_record(out, kBrtEndExt, ByteSpan{});
  emit_record(out, kBrtEndEsfmdInfo, ByteSpan{});

  payload.clear();
  emit_u32(payload, 1U);
  emit_u32(payload, 1U);
  emit_record(out, kBrtBeginEsfmd, payload);
  payload.clear();
  emit_u32(payload, 1U);
  emit_u32(payload, 1U);
  emit_u32(payload, 0U);
  emit_record(out, kBrtCellMetaEntry, payload);
  emit_record(out, kBrtEndEsfmd, ByteSpan{});
  emit_record(out, kBrtEndMetadata, ByteSpan{});
  return out;
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

EmissionPlan BuildEmissionPlan(const Workbook& wb, bool sst_present) {
  EmissionPlan plan;
  plan.has_text_cells = sst_present;
  // Existing XLSB packages retain their original styles bytes verbatim.  This
  // protects style features which are not represented by StylesTable yet
  // (notably differential formats).  XLSX/native workbooks have no raw part,
  // so their modelled style table becomes a fresh styles.bin instead.
  plan.has_generated_styles = !HasRawStylesPart(wb) && HasModelledStyles(wb);
  plan.has_generated_dynamic_metadata =
      HasDynamicArrayMetadata(wb) &&
      !std::any_of(wb.passthrough_parts().begin(), wb.passthrough_parts().end(),
                   [](const PassthroughPart& part) { return part.path == "xl/metadata.bin"; });

  const std::unordered_set<std::string> generated =
      BuildGeneratedPathSet(wb, sst_present, plan.has_generated_styles, plan.has_generated_dynamic_metadata);
  for (const PassthroughPart& part : wb.passthrough_parts()) {
    if (generated.count(part.path) != 0U) {
      StructuredLog("xlsb.writer.passthrough_collision")
          .field("path", part.path)
          .field("reason", std::string_view("generated_path_wins"))
          .warn();
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
    plan.passthrough_kept.push_back(&part);
  }
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
  out.append("  <Default Extension=\"rels\" ContentType=\"");
  out.append(kCtPackageRels);
  out.append("\"/>\n");
  out.append("  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n");
  // Default for `bin` is the workbook content type — matches what the
  // reader test fixture uses and lets passthrough binary parts (e.g.
  // `xl/theme/theme1.xml` is XML, but binary parts like images would
  // ride this Default in a real xlsb).
  out.append("  <Default Extension=\"bin\" ContentType=\"");
  out.append(kCtWorkbookXlsb);
  out.append("\"/>\n");
  // No Override for `/xl/workbook.bin`: it already resolves to the `bin`
  // Default above, and real Excel omits the redundant Override.
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
    AppendXmlEscaped(out, part->path);
    out.append("\" ContentType=\"");
    AppendXmlEscaped(out, part->content_type);
    out.append("\"/>\n");
  }
  out.append("</Types>\n");
  return out;
}

// Appends one `<Relationship>` with a fresh rId drawn from `*next_rid`.
void AppendRelationship(std::string& out, std::size_t* next_rid, std::string_view type, std::string_view target,
                        bool target_external = false) {
  out.append("  <Relationship Id=\"rId");
  out.append(std::to_string((*next_rid)++));
  out.append("\" Type=\"");
  AppendXmlEscaped(out, type);
  out.append("\" Target=\"");
  AppendXmlEscaped(out, target);
  if (target_external) {
    out.append("\" TargetMode=\"External\"/>\n");
  } else {
    out.append("\"/>\n");
  }
}

// Returns true when a passthrough part with `path` will be emitted.
bool HasPassthrough(const EmissionPlan& plan, std::string_view path) {
  for (const PassthroughPart* part : plan.passthrough_kept) {
    if (part->path == path) {
      return true;
    }
  }
  return false;
}

std::string BuildPackageRels(const Workbook& wb, const EmissionPlan& plan) {
  std::string out;
  out.reserve(384);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::size_t next_rid = 1;
  AppendRelationship(out, &next_rid, kRelOfficeDocument, "xl/workbook.bin");
  // docProps parts ride the passthrough path but need package-level rels or
  // Excel treats them as orphaned (and drops the document properties).
  if (HasPassthrough(plan, "docProps/core.xml")) {
    AppendRelationship(out, &next_rid, kRelCoreProps, "docProps/core.xml");
  }
  if (HasPassthrough(plan, "docProps/app.xml")) {
    AppendRelationship(out, &next_rid, kRelExtendedProps, "docProps/app.xml");
  }
  if (HasPassthrough(plan, "docProps/custom.xml")) {
    AppendRelationship(out, &next_rid, kRelCustomProps, "docProps/custom.xml");
  }
  for (const UnknownRelationship& r : wb.unknown_package_rels()) {
    if (!r.target_external && !HasPassthrough(plan, r.target)) {
      StructuredLog("xlsb.writer.package_rel_skipped")
          .field("reason", std::string_view("target_part_absent"))
          .field("type", r.type)
          .field("target", r.target)
          .warn();
      continue;
    }
    AppendRelationship(out, &next_rid, r.type, r.target, r.target_external);
  }
  out.append("</Relationships>\n");
  return out;
}

std::string WorkbookRelativeTarget(std::string_view package_path) {
  constexpr std::string_view kWorkbookDirectory = "xl/";
  if (package_path.size() >= kWorkbookDirectory.size() &&
      package_path.substr(0U, kWorkbookDirectory.size()) == kWorkbookDirectory) {
    return std::string(package_path.substr(kWorkbookDirectory.size()));
  }
  return std::string(package_path);
}

std::string BuildWorkbookRels(std::size_t sheet_count, bool emit_sst, const EmissionPlan& plan, const Workbook& wb) {
  std::string out;
  out.reserve(256 + sheet_count * 192 + wb.unknown_workbook_rels().size() * 192);
  out.append(kXmlDecl);
  out.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n");
  std::size_t next_rid = 1;
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
    AppendRelationship(out, &next_rid, kRelSharedStrings, "sharedStrings.bin");
  }
  // The styles / theme / metadata parts ride the passthrough path, but the
  // reader (and Excel) locate them only through these workbook relationships.
  // Without them a styled cell's iStyleRef dangles against an empty style
  // table, and the theme / metadata parts are treated as orphans. Emit a rel
  // for each part that is actually present in the package.
  if (plan.has_generated_styles || HasPassthrough(plan, "xl/styles.bin")) {
    AppendRelationship(out, &next_rid, kRelStyles, "styles.bin");
  }
  if (HasPassthrough(plan, "xl/theme/theme1.xml")) {
    AppendRelationship(out, &next_rid, kRelTheme, "theme/theme1.xml");
  }
  if (plan.has_generated_dynamic_metadata || HasPassthrough(plan, "xl/metadata.bin")) {
    AppendRelationship(out, &next_rid, kRelSheetMetadata, "metadata.bin");
  }
  // Preserve relationships to raw XLSB parts that the reader does not model
  // (drawings, VBA, custom XML, etc.).  Internal targets are stored as
  // package paths on Workbook, while workbook relationships are relative to
  // `xl/`.  Do not duplicate relationships which the generated package
  // already owns above.
  for (const UnknownRelationship& rel : wb.unknown_workbook_rels()) {
    if (!rel.target_external && !HasPassthrough(plan, rel.target)) {
      StructuredLog("xlsb.writer.workbook_rel_skipped")
          .field("reason", std::string_view("target_part_absent"))
          .field("type", rel.type)
          .field("target", rel.target)
          .warn();
      continue;
    }
    if (!rel.target_external && ((rel.type == kRelTheme && rel.target == "xl/theme/theme1.xml") ||
                                 (rel.type == kRelSheetMetadata && rel.target == "xl/metadata.bin"))) {
      continue;
    }
    const std::string target = rel.target_external ? rel.target : WorkbookRelativeTarget(rel.target);
    AppendRelationship(out, &next_rid, rel.type, target, rel.target_external);
  }
  out.append("</Relationships>\n");
  return out;
}

// ---------------------------------------------------------------------------
// PtgName table (BrtName): defined names + future-function callees
// ---------------------------------------------------------------------------

/// Parses `formula` (with or without a leading `=`) and folds every name
/// `collect_ptg_names` finds into `names` / `seen`. Parse failures are
/// silently skipped here — `EncodeCellFormula` / the defined-name
/// encode pass below surface the same failure as a proper `Expected`
/// error when the formula is actually encoded.
void CollectNamesFromFormula(std::string_view formula, std::vector<std::string>& names,
                             std::unordered_set<std::string>& seen) {
  if (!formula.empty() && formula.front() == '=') {
    formula.remove_prefix(1);
  }
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr || !p.errors().empty()) {
    return;
  }
  collect_ptg_names(*root, names, seen);
}

/// Builds the `PtgName` table for the whole workbook: every genuine
/// defined name (`Workbook::defined_names()`, in declaration order) is
/// registered first, then every future-function callee / `NameRef`
/// `collect_ptg_names` discovers across defined-name and sheet formulas (in
/// first-encounter order) that is not already a defined name. Returns
/// the ordered name list (index `i` <-> 1-based `ilbl` `i + 1`) and the
/// name -> ilbl map `encode_ptgs` consults.
void BuildNameTable(const Workbook& wb, std::vector<std::string>& ordered_names, NameTable& name_table) {
  std::unordered_set<std::string> seen;
  for (const io::DefinedName& dn : wb.defined_names()) {
    if (seen.insert(dn.name).second) {
      ordered_names.push_back(dn.name);
    }
  }
  for (const io::DefinedName& dn : wb.defined_names()) {
    CollectNamesFromFormula(dn.formula, ordered_names, seen);
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const auto& [row, cells] : wb.sheet(i).rows()) {
      (void)row;
      for (const Cell& cell : cells) {
        if (cell.formula_text.empty()) {
          continue;
        }
        CollectNamesFromFormula(cell.formula_text, ordered_names, seen);
      }
    }
  }
  for (std::size_t i = 0; i < ordered_names.size(); ++i) {
    name_table.emplace(ordered_names[i], static_cast<std::uint32_t>(i + 1));
  }
}

// ---------------------------------------------------------------------------
// ExternSheet table (BrtExternSheet): every sheet-qualified reference's
// (itabFirst, itabLast) span, single- and multi-sheet alike.
// ---------------------------------------------------------------------------

/// Parses `formula` (no leading `=`) and folds every distinct
/// `(itabFirst, itabLast)` span `collect_ptg_sheet_ranges` finds into
/// `ranges` / `seen`. Parse failures are silently skipped here --
/// `EncodeCellFormula` surfaces the same failure as a proper `Expected`
/// error when the formula is actually encoded.
void CollectSheetRangesFromFormula(std::string_view formula, const std::vector<std::string>& sheet_names,
                                   SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen) {
  Arena arena;
  parser::Parser p(formula, arena);
  parser::AstNode* root = p.parse();
  if (root == nullptr || !p.errors().empty()) {
    return;
  }
  collect_ptg_sheet_ranges(*root, sheet_names, ranges, seen);
}

/// Builds the `BrtExternSheet` table for the whole workbook: every
/// distinct sheet-qualified reference span (single-sheet `(itab, itab)`
/// or a genuine 3-D range `(itabFirst, itabLast)`) any cell formula or
/// defined-name formula needs, in first-encounter order. Every
/// `PtgRef3d` / `PtgArea3d` token this writer emits -- single- or
/// multi-sheet alike -- resolves its `ixti` through this one table:
/// once the workbook emits any `BrtExternSheet` entry, the reader
/// interprets *every* `ixti` as an index into it rather than a bare
/// sheet index (see `ptg_reader.cpp`'s `sheet_for_ixti` /
/// `sheet_range_for_ixti`), so single- and multi-sheet references
/// cannot use two different numbering schemes in the same file.
SheetRangeTable BuildSheetRangeTable(const Workbook& wb, const std::vector<std::string>& sheet_names) {
  SheetRangeTable ranges;
  std::unordered_set<std::uint64_t> seen;
  for (const io::DefinedName& dn : wb.defined_names()) {
    CollectSheetRangesFromFormula(dn.formula, sheet_names, ranges, seen);
  }
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    for (const auto& [row, cells] : wb.sheet(i).rows()) {
      (void)row;
      for (const Cell& cell : cells) {
        if (cell.formula_text.empty()) {
          continue;
        }
        std::string_view body(cell.formula_text);
        if (!body.empty() && body.front() == '=') {
          body.remove_prefix(1);
        }
        CollectSheetRangesFromFormula(body, sheet_names, ranges, seen);
      }
    }
  }
  return ranges;
}

/// Emits one `BrtExternSheet` record. Byte layout verified against a
/// real Excel-365-produced `xl/workbook.bin` (see `ptg_reader.cpp`'s
/// `DecodeExternSheet`, the decoder counterpart): `count(u32)` followed
/// by `count` entries of `(iSupBook, itabFirst, itabLast)` as three
/// `i32`s each. `iSupBook` is always 0 here (internal-workbook sheets
/// only; external-workbook 3-D ranges are out of scope for this
/// writer).
void EmitExternSheet(std::vector<std::uint8_t>& body, const SheetRangeTable& ranges) {
  if (ranges.empty()) {
    return;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, static_cast<std::uint32_t>(ranges.size()));
  for (const auto& [itab_first, itab_last] : ranges) {
    emit_u32(p, 0);  // iSupBook
    emit_u32(p, static_cast<std::uint32_t>(itab_first));
    emit_u32(p, static_cast<std::uint32_t>(itab_last));
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet), p);
}

/// Emits one `BrtName` record. Byte layout verified against a real
/// Excel-365-produced `xl/workbook.bin` (see `ptg_reader.cpp`'s
/// `DecodeWorkbookNames`, the decoder counterpart):
///   flags (u16: bit 0 = fHidden) + 3 reserved bytes + itab (i32,
///   `-1` = workbook scope) + cch (u32) + cch x UTF-16LE name +
///   cce (u32) + cce bytes rgce + cb (u32) + cb bytes rgcb.
/// `formula` is the name's own bare expression text (no leading `=`,
/// matching `<definedName>`'s XML text content); an empty string emits
/// a zero-length `rgce` (the hidden future-function / LET-parameter
/// placeholders this writer registers carry no real formula body — real
/// Excel stores a `#NAME?` placeholder there, which is not required for
/// this writer's own reader to round-trip the name table).
Expected<void, Error> EmitName(std::vector<std::uint8_t>& body, const std::string& name, std::string_view formula,
                               std::int32_t itab, bool hidden, const std::vector<std::string>& sheet_names,
                               const SheetRangeTable& sheet_ranges, const NameTable& name_table) {
  // Names carrying Excel's hidden storage prefixes are not ordinary
  // defined names: `_xlfn.<FN>` registers a post-2007 "future function"
  // and `_xlpm.<param>` a LET / LAMBDA parameter. Real Excel stores each
  // with a specific flag word (fHidden | fFunc | fFutureFunction for the
  // former, the proc-parameter flags for the latter), a PtgErr(#NAME?)
  // placeholder body, and five trailing null strings. A cell's future-
  // function call resolves through the matching BrtName's ilbl, so these
  // flags and the placeholder body must match Excel byte-for-byte or the
  // callee shows up as #NAME? on load.
  const bool is_param = name.rfind("_xlpm.", 0) == 0;
  const bool is_future_fn = name.rfind("_xlfn.", 0) == 0;
  const bool is_placeholder = is_param || is_future_fn;
  std::uint32_t flags;
  if (is_param) {
    flags = 0x00020019U;
  } else if (is_future_fn) {
    flags = 0x0002000bU;
  } else {
    flags = hidden ? 0x00000001U : 0x00000000U;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, flags);  // grbit (flags word)
  emit_u8(p, 0);       // chKey
  emit_u32(p, static_cast<std::uint32_t>(itab));
  emit_xlwidestring(p, name);
  if (is_placeholder) {
    emit_u32(p, 2);    // cce
    emit_u8(p, 0x1C);  // PtgErr
    emit_u8(p, 0x1D);  // #NAME? error code
    emit_u32(p, 0);    // cb
  } else if (formula.empty()) {
    emit_u32(p, 0);  // cce
    emit_u32(p, 0);  // cb
  } else {
    Arena arena;
    parser::Parser parser(formula, arena);
    parser::AstNode* root = parser.parse();
    if (root == nullptr || !parser.errors().empty()) {
      return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg,
                        "xlsb writer: defined-name formula failed to parse for Ptg encoding",
                        std::string("context=xlsb_writer name=") + name);
    }
    auto encoded_or = encode_ptgs(*root, sheet_names, sheet_ranges, name_table);
    if (!encoded_or) {
      return encoded_or.error();
    }
    const EncodedFormula& encoded = encoded_or.value();
    emit_u32(p, static_cast<std::uint32_t>(encoded.rgce.size()));
    p.insert(p.end(), encoded.rgce.begin(), encoded.rgce.end());
    emit_u32(p, static_cast<std::uint32_t>(encoded.rgcb.size()));
    p.insert(p.end(), encoded.rgcb.begin(), encoded.rgcb.end());
  }
  // Trailing BrtName strings ([MS-XLSB] §2.4.649), each a null
  // XLNullableWideString (cchCharacters = 0xFFFFFFFF, no character data).
  // Excel rejects a BrtName that stops after the formula. A plain defined
  // name carries just the comment; a future-function / proc-parameter
  // placeholder carries five strings (comment plus four further unused
  // strings), matching real Excel output.
  const int trailing_strings = is_placeholder ? 5 : 1;
  for (int i = 0; i < trailing_strings; ++i) {
    emit_u32(p, 0xFFFFFFFFU);  // null XLNullableWideString
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtName), p);
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// Workbook stream (xl/workbook.bin)
// ---------------------------------------------------------------------------

Expected<std::vector<std::uint8_t>, Error> BuildWorkbookBin(const Workbook& wb,
                                                            const std::vector<std::string>& ordered_names,
                                                            const NameTable& name_table,
                                                            const SheetRangeTable& sheet_ranges,
                                                            const std::vector<std::string>& sheet_names) {
  std::vector<std::uint8_t> body;
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBook), ByteSpan{});

  // Workbook globals Excel expects before the sheet bundle. These are
  // fixed-shape records with no dependency on workbook content, so we emit
  // known-valid default payloads (byte layout captured from a real Excel 365
  // `xl/workbook.bin`):
  //   * BrtFileVersion : appName "xl", lastEdited/lowestEdited "7", build.
  //   * BrtWbProp      : default flags + defaultThemeVersion 202300.
  //   * BrtWbView      : a single window view; itabCur is left at 0 (first
  //                      sheet) since the model does not track an active tab.
  // Omitting these produced a workbook stream Excel rejected.
  static const std::vector<std::uint8_t> kFileVersionPayload = {
      0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02,
      0x00, 0x00, 0x00, 0x78, 0x00, 0x6c, 0x00, 0x01, 0x00, 0x00, 0x00, 0x37, 0x00, 0x01, 0x00, 0x00, 0x00,
      0x37, 0x00, 0x05, 0x00, 0x00, 0x00, 0x31, 0x00, 0x30, 0x00, 0x36, 0x00, 0x32, 0x00, 0x38, 0x00};
  static const std::vector<std::uint8_t> kDefaultWbPropPayload = {0x20, 0x00, 0x01, 0x00, 0x3c, 0x16,
                                                                  0x03, 0x00, 0x00, 0x00, 0x00, 0x00};
  static const std::vector<std::uint8_t> kWbViewPayload = {0x10, 0x4f, 0x00, 0x00, 0x18, 0x15, 0x00, 0x00, 0x8c, 0x6e,
                                                           0x00, 0x00, 0xd0, 0x43, 0x00, 0x00, 0x58, 0x02, 0x00, 0x00,
                                                           0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x78};
  emit_record(body, kBrtFileVersion, kFileVersionPayload);
  std::vector<std::uint8_t> wb_prop_payload = kDefaultWbPropPayload;
  // BrtWbProp ([MS-XLSB] §2.4.866): f1904 is bit 0 of the leading
  // little-endian u32 grbit. Preserve Excel's captured default flags.
  if (wb.date1904()) {
    wb_prop_payload[0] |= 0x01U;
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp), wb_prop_payload);
  emit_record(body, kBrtBeginBookViews, ByteSpan{});
  emit_record(body, kBrtWbView, kWbViewPayload);
  emit_record(body, kBrtEndBookViews, ByteSpan{});

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBeginBundleShs), ByteSpan{});
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    // BrtBundleSh ([MS-XLSB] §2.4.304):
    //   hsState    : u32 (0 = visible, 1 = hidden)
    //   iTabID     : u32 (sheet id; 1-based)
    //   strRelID   : XLNullableWideString
    //   strName    : XLWideString
    std::vector<std::uint8_t> p;
    // The Sheet model tracks a single `tab_hidden` bit (mirroring the
    // OOXML `state="hidden"` path), so a hidden sheet emits hsState=1.
    // Very-hidden (hsState=2) is not separately modelled and folds to 1.
    const std::uint32_t hs_state = wb.sheet(i).view().tab_hidden ? 1U : 0U;
    emit_u32(p, hs_state);                            // hsState
    emit_u32(p, static_cast<std::uint32_t>(i + 1U));  // iTabID
    const std::string rid = std::string("rId") + std::to_string(i + 1U);
    emit_xlnullablewidestring(p, std::optional<std::string_view>{rid});
    emit_xlwidestring(p, wb.sheet(i).name());
    emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh), p);
  }
  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBundleShs), ByteSpan{});

  // BrtExternSheet: every distinct sheet-qualified reference span this
  // workbook's formulas need an `ixti` for (single- and multi-sheet
  // alike; see `BuildSheetRangeTable`'s doc comment for why both share
  // this one table). Omitted entirely when no formula uses a qualified
  // reference, matching the fallback the reader's `sheet_for_ixti` /
  // `sheet_range_for_ixti` already implement for that case.
  //
  // The record MUST be wrapped in the externals block: a bare
  // `BrtExternSheet` outside `BrtBeginExternals ... BrtEndExternals` is an
  // out-of-place record that makes Excel reject the package. A workbook that
  // only references its own sheets still needs a single `BrtSupSelf`
  // supporting-book entry inside the block.
  if (!sheet_ranges.empty()) {
    emit_record(body, kBrtBeginExternals, ByteSpan{});
    emit_record(body, kBrtSupSelf, ByteSpan{});
    EmitExternSheet(body, sheet_ranges);
    emit_record(body, kBrtEndExternals, ByteSpan{});
  }

  // BrtName table: every genuine defined name (hidden per its own OOXML
  // `hidden` flag) followed by every future-function callee / defined
  // name `collect_ptg_names` needed a `PtgName` reference for that
  // wasn't already a defined name (always hidden — these are never
  // user-visible). `ordered_names` / `name_table` are built once by
  // `BuildNameTable` and shared with every sheet's cell encoder so
  // `ilbl` assignments stay consistent workbook-wide.
  const std::size_t defined_count = wb.defined_names().size();
  for (std::size_t i = 0; i < ordered_names.size(); ++i) {
    if (i < defined_count) {
      const io::DefinedName& dn = wb.defined_names()[i];
      if (auto r =
              EmitName(body, dn.name, dn.formula, dn.local_sheet_id, dn.hidden, sheet_names, sheet_ranges, name_table);
          !r) {
        return r.error();
      }
    } else {
      if (auto r = EmitName(body, ordered_names[i], /*formula=*/{}, /*itab=*/-1, /*hidden=*/true, sheet_names,
                            sheet_ranges, name_table);
          !r) {
        return r.error();
      }
    }
  }

  emit_record(body, static_cast<std::uint16_t>(XlsbRecordType::BrtEndBook), ByteSpan{});
  return body;
}

// ---------------------------------------------------------------------------
// miniz helpers (mirrors `ooxml_writer.cpp`)
// ---------------------------------------------------------------------------

class ZipWriterGuard {
 public:
  ZipWriterGuard() = default;
  ZipWriterGuard(const ZipWriterGuard&) = delete;
  ZipWriterGuard& operator=(const ZipWriterGuard&) = delete;
  ZipWriterGuard(ZipWriterGuard&&) = delete;
  ZipWriterGuard& operator=(ZipWriterGuard&&) = delete;

  ~ZipWriterGuard() {
    if (active_) {
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

  void release() noexcept { active_ = false; }

 private:
  mz_zip_archive archive_{};
  bool active_ = false;
};

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

Expected<void, Error> AddPartBytes(mz_zip_archive* archive, std::string_view path,
                                   const std::vector<std::uint8_t>& body) {
  const mz_bool ok = mz_zip_writer_add_mem(archive, std::string(path).c_str(), body.data(), body.size(),
                                           static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION));
  if (ok == MZ_FALSE) {
    std::string context("part=");
    context.append(path);
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_add_mem failed (binary)",
                      std::move(context));
  }
  return Expected<void, Error>::Ok();
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

  // PtgName table: every genuine defined name plus every future-function
  // callee / `NameRef` any cell's formula needs, built once so `ilbl`
  // assignments are shared between the sheet bodies below and the
  // `BrtName` table `BuildWorkbookBin` emits into `xl/workbook.bin`.
  std::vector<std::string> ordered_names;
  NameTable name_table;
  BuildNameTable(workbook, ordered_names, name_table);

  // ExternSheet table: every distinct sheet-qualified reference span
  // (single- or multi-sheet) any cell or defined-name formula needs an
  // `ixti` for, built once so the assignment is shared between the
  // sheet bodies below and the `BrtExternSheet` record `BuildWorkbookBin`
  // emits into `xl/workbook.bin`.
  const SheetRangeTable sheet_ranges = BuildSheetRangeTable(workbook, sheet_names);

  SstBuilder sst;
  const bool emit_dynamic_metadata = HasDynamicArrayMetadata(workbook);
  std::uint32_t downgraded_formula_count = 0;
  const std::uint32_t deferred_feature_count = ReportDeferredSheetFeatures(workbook);
  std::vector<std::vector<std::uint8_t>> sheet_bodies;
  sheet_bodies.reserve(sheet_count);
  for (std::size_t i = 0; i < sheet_count; ++i) {
    auto sheet_body_or = emit_sheet(workbook.sheet(i), sst, sheet_names, sheet_ranges, name_table,
                                    &downgraded_formula_count, emit_dynamic_metadata);
    if (!sheet_body_or) {
      return sheet_body_or.error();
    }
    sheet_bodies.push_back(std::move(sheet_body_or.value()));
  }
  const bool emit_sst_part = !sst.empty();

  const EmissionPlan plan = BuildEmissionPlan(workbook, emit_sst_part);

  ZipWriterGuard writer;
  if (!writer.init()) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_init_heap failed", "context=write_xlsb");
  }

  // 1. [Content_Types].xml
  if (auto r = AddPart(writer.get(), "[Content_Types].xml", BuildContentTypes(workbook, plan)); !r) {
    return r.error();
  }
  // 2. _rels/.rels
  if (auto r = AddPart(writer.get(), "_rels/.rels", BuildPackageRels(workbook, plan)); !r) {
    return r.error();
  }
  // 3. xl/_rels/workbook.bin.rels
  // Passthrough parts (styles.bin, theme, metadata, sharedStrings) are only
  // discoverable through workbook relationships; BuildWorkbookRels emits a rel
  // for each part actually present so none dangle on reload.
  if (auto r = AddPart(writer.get(), "xl/_rels/workbook.bin.rels",
                       BuildWorkbookRels(sheet_count, emit_sst_part, plan, workbook));
      !r) {
    return r.error();
  }
  // 4. xl/workbook.bin
  {
    auto wb_bytes_or = BuildWorkbookBin(workbook, ordered_names, name_table, sheet_ranges, sheet_names);
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
    const std::vector<std::uint8_t> metadata_bytes = BuildDynamicArrayMetadataBin();
    if (auto r = AddPartBytes(writer.get(), "xl/metadata.bin", metadata_bytes); !r) {
      return r.error();
    }
  }
  // 5. xl/worksheets/sheet<N>.bin
  for (std::size_t i = 0; i < sheet_count; ++i) {
    std::string path("xl/worksheets/sheet");
    path.append(std::to_string(i + 1));
    path.append(".bin");
    if (auto r = AddPartBytes(writer.get(), path, sheet_bodies[i]); !r) {
      return r.error();
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
    if (auto r = AddPartBytes(writer.get(), part->path, part->bytes); !r) {
      return r.error();
    }
  }

  // Finalise into a heap buffer, then copy into a std::vector.
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  if (mz_zip_writer_finalize_heap_archive(writer.get(), &archive_ptr, &archive_size) == MZ_FALSE) {
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_finalize_heap_archive failed",
                      "context=write_xlsb");
  }
  if (mz_zip_writer_end(writer.get()) == MZ_FALSE) {
    if (archive_ptr != nullptr) {
      mz_free(archive_ptr);
    }
    writer.release();
    return make_error(FormulonErrorCode::kIoWriteFailed, "miniz mz_zip_writer_end failed", "context=write_xlsb");
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
  return XlsbWriteResult{std::move(bytes), downgraded_formula_count, deferred_feature_count};
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
