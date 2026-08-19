//
// OOXML reader orchestrator. Walks the in-memory package via
// `ZipReader`, dispatches each part to a focused helper module, and
// assembles the resulting Workbook. The per-part logic lives under
// `src/io/ooxml/` (package validation + path normalisation, workbook
// rels, sheet aux rels, external links, pivot cache target resolution)
// and under sibling readers (`sheet_reader`, `cf_reader`, `tables_reader`,
// `sst_reader`, `styles_reader`, `comments_reader`, `pivot_cache_reader`,
// `pivot_table_reader`); this file owns only the read pipeline order
// and the bookkeeping that ties the loaded parts back onto the
// workbook.
//
// Shared-strings (`xl/sharedStrings.xml`) resolution is wired in: the
// SST is loaded ahead of the per-sheet read loop, each sheet queues its
// `(row, col, sst_index)` tuples in a per-sheet `SheetReadContext`, and
// a final resolution pass replaces the placeholder `Text("")` values
// with views into the SST. Styles (`xl/styles.xml`) is parsed for
// validation only — the runtime style model lands when the formatter
// pipeline begins consuming it.

#include "io/ooxml_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "eval/iterative_solver.h"
#include "io/cf_reader.h"
#include "io/comments_reader.h"
#include "io/default_content_type.h"
#include "io/defined_names.h"
#include "io/defined_names_internal.h"
#include "io/ooxml/external_link_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/pivot_target_reader.h"
#include "io/ooxml/print_settings_parse.h"
#include "io/ooxml/sheet_aux_rels_reader.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/ooxml_defs.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_table_reader.h"
#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr std::string_view kRelCoreProperties =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
constexpr std::string_view kRelExtendedProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
constexpr std::string_view kRelCustomProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";

Expected<std::vector<UnknownRelationship>, Error> ReadUnknownPackageRels(const std::vector<std::uint8_t>& bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, bytes, "ooxml_reader", "package-level rels"));
  const pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=ooxml_reader part=_rels/.rels");
  }
  std::vector<UnknownRelationship> result;
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type(rel.attribute("Type").value());
    if (type == kRelOfficeDocument || type == kRelCoreProperties || type == kRelExtendedProperties ||
        type == kRelCustomProperties) {
      continue;
    }
    const bool external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    std::string target(rel.attribute("Target").value());
    if (type.empty() || target.empty()) {
      continue;
    }
    if (!external) {
      if (target.front() == '/')
        target.erase(0, 1);
      if (!ooxml::is_safe_part_name(target)) {
        return make_error(FormulonErrorCode::kIoZipSlip, "package relationship target escapes package root",
                          "context=ooxml_reader part=_rels/.rels target=" + target);
      }
    }
    result.push_back(
        UnknownRelationship{std::string(rel.attribute("Id").value()), std::string(type), std::move(target), external});
  }
  return result;
}

// Content type for the binary printer-settings part; the orchestrator
// stamps it onto the passthrough record when it captures a sheet's
// printer-settings bytes (the sheet aux-rels reader only resolves the
// path; the orchestrator owns the round-trip wrapping).
constexpr std::string_view kCtPrinterSettings =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";

// The structured `PageSetup` / `PageMargins` / `ManualBreak` derivations
// live in `io/ooxml/print_settings_parse.h`: the C-ABI print setters
// re-derive them from a freshly written raw fragment, and both sides must
// agree on what the XML means.
using ooxml::apply_structured_page_margins;
using ooxml::apply_structured_page_setup;
using ooxml::read_fit_to_page;
using ooxml::read_manual_breaks;

/// Returns a copy of `sheet_xml` with the `<sheetData>` element's children
/// removed (the open / close tags are kept as an empty element). This lets
/// the SAX path parse the small non-cell worksheet metadata as a DOM
/// without materialising the full cell tree the SAX scanner exists to
/// avoid. Returns the input unchanged when there is no `<sheetData>` (or
/// it is already empty / self-closing).
std::vector<std::uint8_t> BuildWorksheetShellBytes(const std::vector<std::uint8_t>& sheet_xml) {
  const std::string_view sv(reinterpret_cast<const char*>(sheet_xml.data()), sheet_xml.size());
  constexpr std::string_view kOpen = "<sheetData";
  constexpr std::string_view kClose = "</sheetData>";
  // Locate the `<sheetData` open tag, requiring a real name boundary after
  // it so `<sheetDataX>` (hypothetical) does not match.
  std::size_t open = std::string_view::npos;
  for (std::size_t from = 0;;) {
    const std::size_t hit = sv.find(kOpen, from);
    if (hit == std::string_view::npos) {
      break;
    }
    const std::size_t after = hit + kOpen.size();
    const char d = after < sv.size() ? sv[after] : '\0';
    if (d == ' ' || d == '\t' || d == '\r' || d == '\n' || d == '>' || d == '/') {
      open = hit;
      break;
    }
    from = hit + 1;
  }
  if (open == std::string_view::npos) {
    return sheet_xml;
  }
  const std::size_t gt = sv.find('>', open);
  if (gt == std::string_view::npos || sv[gt - 1] == '/') {
    // Malformed, or a self-closing `<sheetData/>` with no children.
    return sheet_xml;
  }
  const std::size_t close = sv.find(kClose, gt + 1);
  if (close == std::string_view::npos) {
    return sheet_xml;
  }
  const std::size_t close_end = close + kClose.size();
  std::vector<std::uint8_t> out;
  out.reserve((gt + 1) + kClose.size() + (sheet_xml.size() - close_end));
  out.insert(out.end(), sheet_xml.begin(), sheet_xml.begin() + static_cast<std::ptrdiff_t>(gt + 1));
  out.insert(out.end(), kClose.begin(), kClose.end());
  out.insert(out.end(), sheet_xml.begin() + static_cast<std::ptrdiff_t>(close_end), sheet_xml.end());
  return out;
}

/// Reads every non-cell worksheet element (siblings of `<sheetData>`) from
/// `doc` into sheet `i`: conditional formats, view / layout, merges,
/// hyperlinks, data validations, sheet protection, and the raw print
/// settings (sheetPr / pageMargins / pageSetup / printOptions /
/// headerFooter / autoFilter / row + col breaks). Shared between the DOM
/// path (full document) and the SAX path (metadata shell). Per-row
/// overrides (`<row ht=/hidden=>`) live inside `<sheetData>` and are only
/// populated on the DOM path; on the SAX shell that content is stripped.
/// Every overlay entry these readers drop lands in `diagnostics`.
Expected<void, Error> ApplyWorksheetMetadata(const pugi::xml_document& doc, std::size_t i, Workbook& wb,
                                             ReadDiagnostics* diagnostics) {
  const pugi::xml_node worksheet = doc.child("worksheet");
  auto cfs_or = read_conditional_formats(worksheet, diagnostics);
  if (!cfs_or) {
    return cfs_or.error();
  }
  wb.sheet(i).mutable_conditional_formats() = std::move(cfs_or.value());
  auto view_layout_or = read_sheet_view_and_layout(doc, i, wb);
  if (!view_layout_or) {
    return view_layout_or.error();
  }
  auto merges_or = read_merges(worksheet, diagnostics);
  if (!merges_or) {
    return merges_or.error();
  }
  wb.sheet(i).mutable_merges() = std::move(merges_or.value());
  auto hls_or = read_hyperlinks(worksheet, diagnostics);
  if (!hls_or) {
    return hls_or.error();
  }
  wb.sheet(i).mutable_hyperlinks() = std::move(hls_or.value());
  auto dvs_or = read_data_validations(worksheet, diagnostics);
  if (!dvs_or) {
    return dvs_or.error();
  }
  wb.sheet(i).mutable_validations() = std::move(dvs_or.value());
  wb.sheet(i).mutable_protection() = read_sheet_protection(worksheet);
  WorksheetRawExtensions& raw_extensions = wb.sheet(i).mutable_raw_extensions();
  auto capture_raw_extension = [&worksheet](const char* name) -> std::string {
    if (const pugi::xml_node node = worksheet.child(name)) {
      return raw_xml(node);
    }
    return {};
  };
  raw_extensions.protected_ranges_xml = capture_raw_extension("protectedRanges");
  raw_extensions.scenarios_xml = capture_raw_extension("scenarios");
  raw_extensions.custom_sheet_views_xml = capture_raw_extension("customSheetViews");
  raw_extensions.phonetic_pr_xml = capture_raw_extension("phoneticPr");
  raw_extensions.ignored_errors_xml = capture_raw_extension("ignoredErrors");
  raw_extensions.legacy_drawing_hf_xml = capture_raw_extension("legacyDrawingHF");
  raw_extensions.picture_xml = capture_raw_extension("picture");
  raw_extensions.ole_objects_xml = capture_raw_extension("oleObjects");
  raw_extensions.controls_xml = capture_raw_extension("controls");

  SheetPrintSettings& print = wb.sheet(i).mutable_print_settings();
  if (pugi::xml_node sheet_pr = worksheet.child("sheetPr"); sheet_pr) {
    // Capture the whole `<sheetPr>` verbatim whenever it exists — it may
    // carry only `tabColor` / `codeName` (VBA binding) with no
    // `<pageSetUpPr>` child, and gating on that child dropped such sheets'
    // `<sheetPr>` entirely on save. The structured `fit_to_page` view is
    // populated additionally when `<pageSetUpPr>` is present.
    print.sheet_pr_xml = raw_xml(sheet_pr);
    print.page_setup.fit_to_page = read_fit_to_page(sheet_pr);
  }
  if (pugi::xml_node page_margins = worksheet.child("pageMargins")) {
    print.page_margins_xml = raw_xml(page_margins);
    apply_structured_page_margins(page_margins, print.page_margins);
  }
  if (pugi::xml_node page_setup = worksheet.child("pageSetup")) {
    print.page_setup_xml = raw_xml(page_setup);
    print.printer_settings_rid = ooxml::relationship_ref_id(page_setup);
    apply_structured_page_setup(page_setup, print.page_setup);
  }
  if (pugi::xml_node print_options = worksheet.child("printOptions")) {
    print.print_options_xml = raw_xml(print_options);
  }
  if (pugi::xml_node header_footer = worksheet.child("headerFooter")) {
    print.header_footer_xml = raw_xml(header_footer);
  }
  if (pugi::xml_node auto_filter = worksheet.child("autoFilter")) {
    wb.sheet(i).set_auto_filter_xml(raw_xml(auto_filter));
  }
  // Worksheet-level `<extLst>` holds the *data* for 2010+ extensions —
  // notably `x14:conditionalFormattings` (DataBar negative-fill / axis /
  // gradient), linked to legacy `cfRule`s by the base `id` attribute
  // (see `cf_reader.h`, which decodes the DataBar fields out of this
  // block into `cf::DataBarSpec`). Capture the block raw regardless, so
  // any other 2010+ extension content it carries (unrelated to CF)
  // survives a save cycle unchanged. Living in this shared helper means
  // the SAX path recovers it too, via the metadata shell.
  if (pugi::xml_node ext_lst = worksheet.child("extLst")) {
    wb.sheet(i).set_ext_lst_xml(raw_xml(ext_lst));
  }
  // Capture the worksheet root's extra namespace declarations so any
  // prefixed attribute carried inside a raw capture above resolves when
  // re-emitted (mirrors the workbook-root handling; keeps the output
  // well-formed).
  wb.sheet(i).set_root_extra_ns_attrs(capture_root_extra_ns_attrs(worksheet));
  read_manual_breaks(worksheet.child("rowBreaks"), print.manual_row_breaks);
  read_manual_breaks(worksheet.child("colBreaks"), print.manual_col_breaks);
  return Expected<void, Error>::Ok();
}

// Encrypted OOXML packages (password-protected .xlsx/.xlsb produced by Excel)
// are wrapped in an OLE/CDFV2 compound-file container, not a ZIP. The container
// begins with the fixed 8-byte signature `D0 CF 11 E0 A1 B1 1A E1`. Without this
// check the bytes reach `ZipReader::open`, which fails with a generic
// "corrupt zip" diagnostic that misleads callers into thinking the file is
// damaged rather than encrypted.
bool IsCdfv2Container(ByteSpan bytes) noexcept {
  static constexpr std::uint8_t kCdfv2Magic[8] = {0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1};
  if (bytes.data == nullptr || bytes.size < sizeof(kCdfv2Magic)) {
    return false;
  }
  return std::memcmp(bytes.data, kCdfv2Magic, sizeof(kCdfv2Magic)) == 0;
}

}  // namespace

namespace internal {

Expected<std::string, Error> ResolveRelativePathForTesting(std::string_view base_dir, std::string_view target) {
  return ooxml::resolve_relative_path(base_dir, target);
}

}  // namespace internal

// Shared implementation behind the public `read_ooxml` and the test-only
// threshold-injection seam. `sax_threshold` is the byte size at or above
// which a sheet routes through the streaming SAX scanner instead of the
// pugixml DOM; production passes `kSaxThresholdBytes`, tests pass a tiny
// value to force the SAX branch on ordinary-size sheets.
static Expected<OoxmlReadResult, Error> ReadOoxmlWithThreshold(ByteSpan bytes, std::size_t sax_threshold) {
  ReadDiagnostics diagnostics;
  // Surface a precise "encrypted" diagnostic before the ZIP layer reports the
  // CDFV2 container as a corrupt archive.
  if (IsCdfv2Container(bytes)) {
    return make_error(FormulonErrorCode::kIoZipEncrypted,
                      "package is an encrypted OLE/CDFV2 container; decrypt before loading", "context=ooxml_reader");
  }
  ZipReader zip;
  {
    auto open_result = zip.open(bytes);
    if (!open_result) {
      return open_result.error();
    }
  }

  // 1. [Content_Types].xml — sanity check + listing for unknown_parts.
  if (!zip.has_entry("[Content_Types].xml")) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing from package",
                      "context=ooxml_reader");
  }
  auto ct_bytes_or = zip.read_entry("[Content_Types].xml");
  if (!ct_bytes_or) {
    return ct_bytes_or.error();
  }
  const std::vector<std::uint8_t>& ct_bytes = ct_bytes_or.value();
  auto kind_or = ooxml::verify_content_types(ct_bytes, &diagnostics);
  if (!kind_or) {
    return kind_or.error();
  }
  const io::WorkbookKind workbook_kind = kind_or.value();
  auto override_part_entries_or = ooxml::list_override_part_entries(ct_bytes);
  if (!override_part_entries_or) {
    return override_part_entries_or.error();
  }
  const std::vector<ooxml::OverrideEntry> override_part_entries = std::move(override_part_entries_or.value());
  auto default_content_types_or = ooxml::list_default_content_types(ct_bytes);
  if (!default_content_types_or) {
    return default_content_types_or.error();
  }
  std::vector<DefaultContentType> default_content_types = std::move(default_content_types_or.value());

  // 2. _rels/.rels — locate the workbook part path.
  if (!zip.has_entry("_rels/.rels")) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "_rels/.rels: missing from package",
                      "context=ooxml_reader");
  }
  auto root_rels_or = zip.read_entry("_rels/.rels");
  if (!root_rels_or) {
    return root_rels_or.error();
  }
  auto wb_path_or = ooxml::resolve_office_document_path(root_rels_or.value());
  if (!wb_path_or) {
    return wb_path_or.error();
  }
  const std::string workbook_path = wb_path_or.value();
  auto package_rels_or = ReadUnknownPackageRels(root_rels_or.value());
  if (!package_rels_or) {
    return package_rels_or.error();
  }

  // 3. xl/_rels/workbook.xml.rels — load and validate. We need both the
  // sheet rId -> part-path map (for the per-sheet read loop below) and
  // the resolved paths for the sharedStrings / styles parts so we can
  // load them at the right point in the pipeline.
  auto wb_rels_or = ooxml::load_workbook_rels(zip, workbook_path);
  if (!wb_rels_or) {
    return wb_rels_or.error();
  }
  ooxml::WorkbookRels& wb_rels = wb_rels_or.value();

  // 4. xl/workbook.xml — the <sheets> list (in document order).
  if (!zip.has_entry(workbook_path)) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: part not found at relationship target",
                      "context=ooxml_reader workbook_path=" + workbook_path);
  }
  auto wb_bytes_or = zip.read_entry(workbook_path);
  if (!wb_bytes_or) {
    return wb_bytes_or.error();
  }
  const std::vector<std::uint8_t>& wb_bytes = wb_bytes_or.value();

  pugi::xml_document wb_doc;
  RETURN_IF_ERROR(load_xml_buffer(wb_doc, wb_bytes, "ooxml_reader", "workbook.xml"));
  pugi::xml_node wb_root = wb_doc.child("workbook");
  if (!wb_root) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: missing <workbook> root",
                      "context=ooxml_reader part=" + workbook_path);
  }
  pugi::xml_node sheets_node = wb_root.child("sheets");
  if (!sheets_node) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: missing <sheets>",
                      "context=ooxml_reader part=" + workbook_path);
  }

  // Build the workbook bottom-up: empty container, then append sheets in
  // document order. `create_empty()` keeps this loop simple — no implicit
  // Sheet1 to overwrite or remove. The workbook kind is set up-front so
  // any subsequent error path (e.g. corrupt sheet) still returns metadata
  // consistent with the `[Content_Types].xml` declaration we observed.
  Workbook wb = Workbook::create_empty();
  wb.set_kind(workbook_kind);

  // Capture the `<workbook>` root's extra namespace declarations (and
  // `mc:Ignorable`) so the writer can re-emit them. The raw `<bookViews>`
  // capture below can carry namespaced attributes (e.g. `xr2:uid`) whose
  // prefixes are declared only on this root; without re-declaring them the
  // re-emitted fragment is malformed XML and Excel refuses the file. The
  // same helper handles the `<worksheet>` root (see ApplyWorksheetMetadata).
  wb.set_workbook_root_extra_attrs(capture_root_extra_ns_attrs(wb_root));

  // Track which sheet relationships were consumed, so the unknown-parts
  // computation can subtract them.
  std::unordered_set<std::string> consumed_parts;
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  consumed_parts.insert(ooxml::rels_path_for_part(workbook_path));
  std::vector<PassthroughPart> extra_passthrough_parts;

  // Collect (display_name, part_path) per sheet in document order. The
  // sheet's `r:id` attribute resolves to a part path through
  // `sheet_rels`. We need both pieces in sync so the sheet at workbook
  // index `i` is read from `sheet_part_paths[i]`.
  const std::unordered_map<std::string, ooxml::WorkbookRels::SheetTarget>& sheet_rels = wb_rels.sheet_targets;
  std::vector<std::string> sheet_part_paths;
  for (pugi::xml_node sn = sheets_node.child("sheet"); sn; sn = sn.next_sibling("sheet")) {
    std::string name = sn.attribute("name").value();
    // OOXML requires a non-empty name; treat blank as corrupt rather than
    // silently appending an unnamed sheet (which would round-trip to an
    // invalid workbook).
    if (name.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: <sheet> with empty name attribute",
                        "context=ooxml_reader part=" + workbook_path);
    }
    // Resolve r:id -> sheet part path. Excel always emits the attribute;
    // accept both the Office-namespaced ("r:id") and the unprefixed
    // ("id") variants because pugixml exposes namespace-prefixed names
    // verbatim and writers in the wild diverge on which one they use.
    std::string rid = ooxml::relationship_ref_id(sn);
    if (rid.empty()) {
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: <sheet> missing r:id attribute",
                        "context=ooxml_reader part=" + workbook_path);
    }
    auto it = sheet_rels.find(rid);
    if (it == sheet_rels.end()) {
      std::string ctx("context=ooxml_reader part=");
      ctx.append(workbook_path);
      ctx.append(" rid=");
      ctx.append(rid);
      return make_error(FormulonErrorCode::kIoRelationshipBroken,
                        "workbook.xml: r:id has no matching workbook relationship", std::move(ctx));
    }
    // Workbook-side `<sheet state="hidden">` (or `state="veryHidden"`)
    // attribute — Excel records sheet visibility here, not on the
    // worksheet part. Mirror it onto `Sheet::view().tab_hidden` so
    // either signal flips the bit at load time. The worksheet-side
    // `<sheetPr><tabHidden/>` form is handled by the per-sheet reader
    // call below; the merge across both paths is OR-style.
    const std::string_view state = sn.attribute("state").value();
    const bool workbook_hides = (state == "hidden") || (state == "veryHidden");
    // Validate the name at the boundary instead of trusting it. Sheet
    // lookup resolves to the first match, so a workbook carrying two
    // sheets whose names Unicode-simple-fold together would answer every reference from
    // one of them and produce a confident wrong number with no ambiguity
    // signal a caller could act on. Excel treats such a file as needing
    // repair rather than opening it; renaming a sheet the author wrote
    // would be a worse repair than refusing the load.
    auto added = wb.add_sheet_validated(name);
    if (!added) {
      if (added.error().code != FormulonErrorCode::kInvalidSheetName) {
        return added.error();
      }
      std::string ctx("context=ooxml_reader part=");
      ctx.append(workbook_path);
      ctx.append(" sheet=\"").append(name).append("\"");
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "workbook.xml: <sheet> name is invalid or collides with an earlier sheet", std::move(ctx));
    }
    if (workbook_hides) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().tab_hidden = true;
    }
    const ooxml::WorkbookRels::SheetTarget& target = it->second;
    if (target.relationship_type != kRelWorksheet) {
      // Keep non-worksheet sheet types in their original position, but do
      // not feed their incompatible XML through the worksheet reader. The
      // raw part and any descendants remain unconsumed and are captured by
      // the normal passthrough sweep below.
      wb.sheet(wb.sheet_count() - 1U).set_opaque_ooxml_sheet(target.path, target.relationship_type);
    } else {
      consumed_parts.insert(target.path);
    }
    sheet_part_paths.push_back(target.path);
  }

  if (sheet_part_paths.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: empty <sheets> list",
                      "context=ooxml_reader part=" + workbook_path);
  }

  // <calcPr> — workbook-level calc mode and iterative-calc options. The
  // element is optional; absence means Excel defaults (auto + iterative
  // off). When present, accept the documented `calcMode` values
  // (`auto` / `manual` / `autoNoTable`) and the iterative trio
  // (`iterate`, `iterateCount`, `iterateDelta`). Unknown calcMode
  // strings fall back to `auto` rather than failing the load.
  if (pugi::xml_node calc_pr = wb_root.child("calcPr"); calc_pr) {
    const std::string_view calc_mode_attr = calc_pr.attribute("calcMode").value();
    if (calc_mode_attr == "manual") {
      wb.set_calc_mode(Workbook::CalcMode::kManual);
    } else if (calc_mode_attr == "autoNoTable") {
      wb.set_calc_mode(Workbook::CalcMode::kAutoNoTable);
    } else {
      wb.set_calc_mode(Workbook::CalcMode::kAuto);
    }
    eval::IterativeOptions opts;
    if (pugi::xml_attribute iterate = calc_pr.attribute("iterate"); iterate) {
      opts.enabled = parse_xml_bool(iterate.value());
    }
    if (pugi::xml_attribute count = calc_pr.attribute("iterateCount"); count) {
      const long long parsed = count.as_llong(static_cast<long long>(eval::kDefaultMaxIterations));
      opts.max_iterations = parsed < 1 ? 1U : static_cast<std::uint32_t>(parsed);
    }
    // The solver stops once the largest change falls below `max_change`.
    // A NaN tolerance makes that comparison false forever, so the workbook
    // silently burns the whole iteration budget and reports the
    // unconverged values; a negative one can never be reached either.
    double delta_value = 0.0;
    if (parse_xsd_nonneg_double(attr_str(calc_pr, "iterateDelta"), &delta_value)) {
      opts.max_change = delta_value;
    }
    wb.set_iterative_options(opts);
  }

  // Workbook-level elements captured raw for verbatim re-emission:
  // `<fileVersion>`, `<fileSharing>`, `<workbookPr>`, `<workbookProtection>`,
  // `<bookViews>`, and trailing `<extLst>`. Without this,
  // the writer regenerates only `<sheets>` / `<definedNames>` / `<calcPr>`
  // / `<pivotCaches>` and silently drops the date system, tab-selection
  // state, and workbook protection. `<workbookPr date1904>` additionally
  // seeds the model-level `date1904` flag the date-serial conversions read.
  if (pugi::xml_node file_version = wb_root.child("fileVersion"); file_version) {
    wb.set_file_version_xml(raw_xml(file_version));
  }
  if (pugi::xml_node file_sharing = wb_root.child("fileSharing"); file_sharing) {
    wb.set_file_sharing_xml(raw_xml(file_sharing));
  }
  if (pugi::xml_node workbook_pr = wb_root.child("workbookPr"); workbook_pr) {
    wb.set_workbook_pr_xml(raw_xml(workbook_pr));
    // Excel emits the attribute as `date1904`; some legacy producers use
    // the bare `1904` spelling. Both default to false when absent.
    const bool from_date1904 = read_xsd_bool(workbook_pr, "date1904", false);
    const bool from_legacy = read_xsd_bool(workbook_pr, "1904", false);
    wb.set_date1904(from_date1904 || from_legacy);
  }
  if (pugi::xml_node workbook_protection = wb_root.child("workbookProtection"); workbook_protection) {
    wb.set_workbook_protection_xml(raw_xml(workbook_protection));
  }
  if (pugi::xml_node book_views = wb_root.child("bookViews"); book_views) {
    wb.set_book_views_xml(raw_xml(book_views));
  }
  if (pugi::xml_node ext_lst = wb_root.child("extLst"); ext_lst) {
    wb.set_workbook_ext_lst_xml(raw_xml(ext_lst));
  }

  // The workbook owns the text-storage deque; readers append directly
  // into `wb.mutable_text_storage()`. This keeps `Value::text` views
  // valid for the full workbook lifetime — including after the caller
  // moves `wb` out of the read result and discards the result. A
  // `std::deque` (pointer-stable across appends) is required so the
  // views handed to cells remain valid as later strings are appended.
  std::deque<std::string>& result_text_storage = wb.mutable_text_storage();

  // 4a. Shared strings — must load BEFORE the sheet read loop because
  // the sheet reader queues `(row, col, sst_index)` tuples that we
  // resolve after the loop. The relationship is optional; a workbook
  // with no text-via-SST cells legally omits it.
  SharedStringTable sst;
  if (!wb_rels.sst_path.empty()) {
    if (!zip.has_entry(wb_rels.sst_path)) {
      std::string ctx("context=ooxml_reader sst_path=");
      ctx.append(wb_rels.sst_path);
      return make_error(FormulonErrorCode::kIoRelationshipBroken, "sharedStrings: rel target missing from package",
                        std::move(ctx));
    }
    auto sst_bytes_or = zip.read_entry(wb_rels.sst_path);
    if (!sst_bytes_or) {
      return sst_bytes_or.error();
    }
    auto sst_or = read_shared_strings(std::move(sst_bytes_or.value()), result_text_storage);
    if (!sst_or) {
      return sst_or.error();
    }
    sst = std::move(sst_or.value());
    consumed_parts.insert(wb_rels.sst_path);
  }
  // Always swallow the canonical SST path if present in the archive,
  // even when no relationship pointed to it (some writers emit the part
  // via Override only). This keeps `unknown_parts` from surfacing it.
  if (zip.has_entry("xl/sharedStrings.xml")) {
    consumed_parts.insert("xl/sharedStrings.xml");
  }

  // Swallow calcChain.xml rather than passing it through. It is purely a
  // recalculation-order cache; because the writer rewrites cell values it
  // is necessarily stale after any save, and a stale calcChain makes real
  // Excel reject / "repair" the workbook. Dropping it is safe (Excel
  // rebuilds the chain on demand) and it consuming the part here means the
  // passthrough sweep skips it, the workbook-rels calcChain relationship is
  // not re-emitted (its target is no longer a passthrough part, so the
  // unknown-rel guard drops it), and no `<Override>` is written for it.
  if (zip.has_entry("xl/calcChain.xml")) {
    consumed_parts.insert("xl/calcChain.xml");
  }

  // 4b. Styles — read for validation only at this slice. The full
  // numFmt/font/fill runtime model lands later when the formatter
  // pipeline begins consuming it. Mark consumed regardless of whether
  // the rel was present, mirroring the SST treatment above.
  if (!wb_rels.styles_path.empty()) {
    if (!zip.has_entry(wb_rels.styles_path)) {
      std::string ctx("context=ooxml_reader styles_path=");
      ctx.append(wb_rels.styles_path);
      return make_error(FormulonErrorCode::kIoRelationshipBroken, "styles: rel target missing from package",
                        std::move(ctx));
    }
    auto styles_bytes_or = zip.read_entry(wb_rels.styles_path);
    if (!styles_bytes_or) {
      return styles_bytes_or.error();
    }
    auto styles_or = read_styles(styles_bytes_or.value());
    if (!styles_or) {
      return styles_or.error();
    }
    wb.set_styles(std::move(styles_or.value()));
    consumed_parts.insert(wb_rels.styles_path);
  }
  if (zip.has_entry("xl/styles.xml")) {
    consumed_parts.insert("xl/styles.xml");
  }

  // 5. Read each sheet's <sheetData> via the cell-aware sheet reader.
  // The sheet reader appends every inline-string payload directly into
  // the workbook-owned text-storage deque, a pointer-stable
  // `std::deque` whose lifetime is the workbook's. Cells therefore
  // hold `Value::text` views that remain valid for the workbook's
  // lifetime — even after the `OoxmlReadResult` is destroyed and the
  // workbook is moved out.
  //
  // We retain each sheet's `pending_sst_cells` so the SST resolution
  // pass below can rewrite the placeholder text values without rerunning
  // the sheet walk.
  //
  // The same loop also walks each sheet's `_rels/sheetN.xml.rels` (when
  // present) to discover and load referenced table parts. Tables are
  // accumulated into `tables_metadata` and stashed on the workbook
  // after the loop; this is passive metadata for round-trip and does
  // not influence cell evaluation at this layer.
  std::vector<SheetReadContext> sheet_contexts(sheet_part_paths.size());
  std::vector<TableMetadata> tables_metadata;
  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string& sheet_path = sheet_part_paths[i];
    if (wb.sheet(i).is_opaque_ooxml_sheet()) {
      // Ensure the relationship does not silently become dangling. The raw
      // part stays unconsumed so it (and its own rels/dependencies) flows
      // into passthrough unchanged.
      if (!zip.has_entry(sheet_path)) {
        std::string ctx("context=ooxml_reader opaque_sheet_path=");
        ctx.append(sheet_path);
        return make_error(FormulonErrorCode::kIoSheetCorrupt, "opaque sheet part missing from package", std::move(ctx));
      }
      continue;
    }
    if (!zip.has_entry(sheet_path)) {
      // The relationship resolved to a path the package does not
      // contain. Treat as a structural error rather than silently
      // skipping: a missing sheet part is data loss.
      std::string ctx("context=ooxml_reader sheet_path=");
      ctx.append(sheet_path);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "sheet part missing from package", std::move(ctx));
    }
    auto sheet_bytes_or = zip.read_entry(sheet_path);
    if (!sheet_bytes_or) {
      return sheet_bytes_or.error();
    }
    // Bound to a mutable reference because the DOM path below parses it
    // in place; nothing else in this iteration reads the sheet bytes
    // after that point, and `sheet_bytes_or` outlives `sheet_doc`.
    std::vector<std::uint8_t>& sheet_bytes = sheet_bytes_or.value();
    // Choose the read path by raw XML size: small sheets stay on the
    // pugixml DOM path (familiar code, well-validated); large sheets
    // (>= `kSaxThresholdBytes`) stream through the SAX scanner so a
    // 1M-cell worksheet does not need to materialise as a DOM in
    // memory. Both paths produce identical Workbook output.
    //
    // Whether the SAX path is compiled in at all is a compile-time
    // decision: on WASM `kSaxThresholdBytes` is `SIZE_MAX` (see
    // `sheet_reader.h`) so the branch is statically dead and the linker
    // removes the streaming scanner entirely — saving ~17 KiB of `.wasm`.
    // The `if constexpr` makes the elimination explicit so it is robust
    // under -O0 / -Og too. The *runtime* threshold is `sax_threshold`
    // (defaults to `kSaxThresholdBytes`; tests inject a tiny value).
    constexpr bool kSaxEnabled = kSaxThresholdBytes != static_cast<std::size_t>(-1);
    bool sax_used = false;
    if constexpr (kSaxEnabled) {
      if (sheet_bytes.size() >= sax_threshold) {
        ByteSpan sheet_span{sheet_bytes.data(), sheet_bytes.size()};
        auto rs = read_sheet_data_sax(sheet_span, i, wb, sheet_contexts[i], result_text_storage);
        if (!rs) {
          return rs.error();
        }
        sax_used = true;
      }
    }
    // Parse the worksheet's non-cell metadata (siblings of <sheetData>)
    // from a DOM. On the DOM path this is the full sheet DOM the cell
    // reader already consumed; on the SAX path we parse a lightweight
    // shell with <sheetData> stripped so the streamed cell tree is never
    // materialised — recovering conditional formats, view / layout,
    // merges, hyperlinks, data validations, protection, and print
    // settings that the SAX path used to drop. Row overrides
    // (<row ht=/hidden=>) live inside <sheetData> and stay DOM-only.
    //
    // Both parses below are in place: the DOM aliases the buffer it was
    // built from instead of pugixml holding a second full-size copy,
    // which is what made a large sheet cost twice its own size to open.
    // The buffers are single-use — the shell is built here and read
    // nowhere else, `sheet_bytes` is dead once the cell reader has run.
    // `shell` is declared ahead of `sheet_doc` so it is destroyed after
    // it; `sheet_bytes_or` already sits further out for the same reason.
    std::vector<std::uint8_t> shell;
    pugi::xml_document sheet_doc;
    if (sax_used) {
      shell = BuildWorksheetShellBytes(sheet_bytes);
      RETURN_IF_ERROR(load_xml_buffer_inplace(sheet_doc, shell, "ooxml_reader", "sheet*.xml (metadata shell)"));
    } else {
      RETURN_IF_ERROR(load_xml_buffer_inplace(sheet_doc, sheet_bytes, "ooxml_reader", "sheet*.xml"));
      auto rs = read_sheet_data(sheet_doc, i, wb, sheet_contexts[i], result_text_storage);
      if (!rs) {
        return rs.error();
      }
    }
    RETURN_IF_ERROR(ApplyWorksheetMetadata(sheet_doc, i, wb, &diagnostics));

    // Sheet rels file (`xl/worksheets/_rels/sheetN.xml.rels`) — drives
    // the table-part and pivot-table lookups. Optional: most sheets have
    // no rels at all. Tables and pivot tables are read in two passes
    // through the same rels file (separate helpers, each scoped to one
    // relationship type) so each consumer site reads linearly.
    const std::string sheet_rels_path = ooxml::rels_path_for_part(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      const std::string sheet_dir = ooxml::dir_of(sheet_path);
      auto targets_or = ooxml::load_sheet_table_targets(zip, sheet_rels_path, sheet_dir);
      if (!targets_or) {
        return targets_or.error();
      }
      consumed_parts.insert(sheet_rels_path);
      for (const std::string& table_path : targets_or.value()) {
        if (!zip.has_entry(table_path)) {
          std::string ctx("context=ooxml_reader sheet_index=");
          ctx.append(std::to_string(i));
          ctx.append(" table_path=").append(table_path);
          return make_error(FormulonErrorCode::kIoRelationshipBroken, "table: rel target missing from package",
                            std::move(ctx));
        }
        auto table_bytes_or = zip.read_entry(table_path);
        if (!table_bytes_or) {
          return table_bytes_or.error();
        }
        auto table_or = read_table(table_bytes_or.value(), i);
        if (!table_or) {
          return table_or.error();
        }
        tables_metadata.push_back(std::move(table_or.value()));
        consumed_parts.insert(table_path);
      }

      // Pivot tables anchored on this sheet. Each part feeds into the
      // pivot-table reader and is attached to the owning sheet; the
      // workbook-level pivot caches are loaded after the sheet loop.
      auto pivot_targets_or = ooxml::load_sheet_pivot_table_targets(zip, sheet_rels_path, sheet_dir);
      if (!pivot_targets_or) {
        return pivot_targets_or.error();
      }
      for (const std::string& pivot_table_path : pivot_targets_or.value()) {
        if (!zip.has_entry(pivot_table_path)) {
          std::string ctx("context=ooxml_reader sheet_index=");
          ctx.append(std::to_string(i));
          ctx.append(" pivot_table_path=").append(pivot_table_path);
          return make_error(FormulonErrorCode::kIoRelationshipBroken, "pivotTable: rel target missing from package",
                            std::move(ctx));
        }
        auto pt_bytes_or = zip.read_entry(pivot_table_path);
        if (!pt_bytes_or) {
          return pt_bytes_or.error();
        }
        auto pt_or = read_pivot_table_definition(pt_bytes_or.value());
        if (!pt_or) {
          return pt_or.error();
        }
        wb.sheet(i).add_pivot_table(std::make_unique<pivot::PivotTable>(std::move(pt_or.value())));
        consumed_parts.insert(pivot_table_path);
        // The pivot-table part may carry its own rels file pointing back
        // at the parent cache definition. We do not need to re-resolve
        // it (the table already carries `pivot_cache_id`), but we mark
        // the rels file as consumed so it does not surface as an
        // unknown part.
        const std::string pt_rels_path = ooxml::rels_path_for_part(pivot_table_path);
        if (zip.has_entry(pt_rels_path)) {
          consumed_parts.insert(pt_rels_path);
        }
      }

      // Hyperlink / comments / VML auxiliary parts. The rels walker
      // surfaces all three in one pass; missing entries are simply
      // empty in the result. The worksheet body's own `<legacyDrawing>`
      // element (comment geometry) names which `kRelVmlDrawing`
      // relationship is the modelled comment-VML slot, distinguishing it
      // from a second `<legacyDrawingHF>` (header/footer image)
      // relationship the sheet may also carry.
      const std::string legacy_drawing_body_rid =
          ooxml::relationship_ref_id(sheet_doc.child("worksheet").child("legacyDrawing"));
      auto aux_or = ooxml::load_sheet_aux_rels(zip, sheet_rels_path, sheet_dir, legacy_drawing_body_rid);
      if (!aux_or) {
        return aux_or.error();
      }
      const ooxml::SheetAuxRels& aux = aux_or.value();
      wb.sheet(i).set_unknown_relationships(aux.unknown_rels);
      // Stitch each hyperlink's `target` from the rels lookup.
      apply_hyperlink_rels(wb.sheet(i).mutable_hyperlinks(), aux.hyperlink_rid_to_target);

      if (!aux.printer_settings_path.empty()) {
        SheetPrintSettings& print = wb.sheet(i).mutable_print_settings();
        if (print.printer_settings_rid.empty()) {
          print.printer_settings_rid = aux.printer_settings_rid;
        }
        print.printer_settings_path = aux.printer_settings_path;
        if (zip.has_entry(aux.printer_settings_path)) {
          auto pb_or = zip.read_entry(aux.printer_settings_path);
          if (!pb_or) {
            return pb_or.error();
          }
          PassthroughPart part;
          part.path = aux.printer_settings_path;
          part.content_type = std::string(kCtPrinterSettings);
          part.bytes = std::move(pb_or.value());
          extra_passthrough_parts.push_back(std::move(part));
          consumed_parts.insert(aux.printer_settings_path);
        }
      }

      // Comments part: load + attach. The VML drawing companion is
      // intentionally NOT consumed here so the bytes flow through the
      // unknown-parts passthrough mechanism unchanged.
      if (!aux.comments_path.empty() && zip.has_entry(aux.comments_path)) {
        auto cb_or = zip.read_entry(aux.comments_path);
        if (!cb_or) {
          return cb_or.error();
        }
        auto comments_or = read_comments(cb_or.value());
        if (!comments_or) {
          return comments_or.error();
        }
        wb.sheet(i).mutable_comments() = std::move(comments_or.value());
        wb.sheet(i).set_comment_vml_path(aux.vml_path);
        consumed_parts.insert(aux.comments_path);
      }

      // Drawing (DrawingML) reference. The reader does not model the
      // drawing part; it records the target path so the writer can
      // re-emit the `<drawing>` element and its sheet-rels relationship.
      // The part body, its own rels, and any anchored media round-trip
      // through the Default-typed passthrough capture below.
      if (!aux.drawing_path.empty()) {
        wb.sheet(i).set_drawing_rel_target(aux.drawing_path);
      }
    }
  }

  // 6. Resolve every queued SST reference: replace each cell's
  // `Text("")` placeholder with a view into the SST entry. We use
  // `Sheet::set_cell_cached_value` directly rather than
  // `Workbook::set_cell_value` because (a) these cells are pure data
  // (no formula text to disturb) and (b) the workbook is freshly built
  // from disk, so there is no live dep-graph state to dirty.
  std::uint32_t pending_sst_count = 0;
  for (std::size_t i = 0; i < sheet_contexts.size(); ++i) {
    const SheetReadContext& sctx = sheet_contexts[i];
    if (sctx.pending_sst_cells.empty()) {
      continue;
    }
    if (wb_rels.sst_path.empty() && sst.entries.empty()) {
      // Sheet referenced an SST index but the package did not declare a
      // sharedStrings part. The first such reference is the diagnostic
      // anchor.
      const auto& first = sctx.pending_sst_cells.front();
      std::string ctx("context=ooxml_reader sheet_index=");
      ctx.append(std::to_string(i));
      ctx.append(" row=").append(std::to_string(std::get<0>(first)));
      ctx.append(" col=").append(std::to_string(std::get<1>(first)));
      ctx.append(" sst_index=").append(std::to_string(std::get<2>(first)));
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "sheet references SST but package has no sharedStrings part", std::move(ctx));
    }
    for (const auto& triple : sctx.pending_sst_cells) {
      const std::uint32_t row = std::get<0>(triple);
      const std::uint32_t col = std::get<1>(triple);
      const std::uint32_t idx = std::get<2>(triple);
      if (idx >= sst.entries.size()) {
        std::string ctx("context=ooxml_reader sheet_index=");
        ctx.append(std::to_string(i));
        ctx.append(" row=").append(std::to_string(row));
        ctx.append(" col=").append(std::to_string(col));
        ctx.append(" sst_index=").append(std::to_string(idx));
        ctx.append(" sst_size=").append(std::to_string(sst.entries.size()));
        return make_error(FormulonErrorCode::kIoSheetCorrupt, "shared-string index out of range", std::move(ctx));
      }
      wb.sheet(i).set_cell_cached_value_borrowed(row, col, Value::text(sst.entries[idx]));
      // Propagate any <rPh> annotation from the SST entry onto the cell
      // so PHONETIC() can surface the kana. The SST reader keeps
      // `phonetic_for_entries` parallel to `entries`, with empty views
      // for unannotated entries; we only commit a non-empty annotation
      // to avoid allocating into `Cell::phonetic_text` for the common
      // unannotated case.
      if (idx < sst.phonetic_for_entries.size() && !sst.phonetic_for_entries[idx].empty()) {
        wb.sheet(i).set_cell_phonetic(row, col, sst.phonetic_for_entries[idx]);
      }
      ++pending_sst_count;
    }
  }

  // 7. Defined names — parsed from the same `wb_doc` we already loaded
  // above. Pure metadata extraction; the writer slice (Bundle 2.5) will
  // emit them back. Resolution at evaluation time arrives in Phase 4.
  // We run this after sheet construction so future name validation can
  // cross-reference sheet indices without re-shuffling the call order.
  auto defined_names_or = read_defined_names(wb_doc);
  if (!defined_names_or) {
    return defined_names_or.error();
  }
  wb.set_defined_names(std::move(defined_names_or.value()));

  // 8. Tables — already accumulated in the per-sheet loop above. Move
  // the workbook-scope vector onto the workbook for round-trip.
  wb.set_tables(std::move(tables_metadata));

  // 8b. External links. Joins `<externalReferences>` against the
  // workbook rels and the per-link rels files. The body parts continue
  // to round-trip through `passthrough_parts()`; only the per-link rels
  // files are marked as consumed (their content is regenerated by the
  // writer from the captured records). Missing parts and successfully-read
  // malformed XML remain failure-tolerant, while extraction failures are
  // returned unchanged.
  {
    auto ext_or = ooxml::load_external_links(zip, wb_root, wb_rels);
    if (!ext_or) {
      return ext_or.error();
    }
    ooxml::ExternalLinkLoadResult ext = ext_or.take();
    for (const std::string& rels_path : ext.consumed_rels_paths) {
      consumed_parts.insert(rels_path);
    }
    wb.set_external_links(std::move(ext.records));
  }

  // 9. Pivot caches. The workbook XML's `<pivotCaches>` element pairs
  // each `cacheId` with a workbook-scoped relationship id; the
  // workbook rels file resolves that id to the part path of the
  // `pivotCacheDefinition*.xml`. Each definition's own rels file (if
  // present) points at the matching `pivotCacheRecords*.xml`. We load
  // both parts here so the workbook owns a fully populated
  // `pivot::PivotCache` ready for evaluation; the per-sheet pivot-table
  // loading above attaches `PivotTable`s that reference these caches by
  // id.
  pugi::xml_node pivot_caches_node = wb_root.child("pivotCaches");
  for (pugi::xml_node pc = pivot_caches_node.child("pivotCache"); pc; pc = pc.next_sibling("pivotCache")) {
    const std::uint32_t cache_id = pc.attribute("cacheId").as_uint(0U);
    // Accept both "r:id" (Office-namespaced) and bare "id" — same
    // forgiveness as the sheet relationship walk above.
    std::string rid = ooxml::relationship_ref_id(pc);
    if (rid.empty()) {
      return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook.xml: <pivotCache> missing r:id attribute",
                        "context=ooxml_reader part=" + workbook_path);
    }
    auto def_it = wb_rels.pivot_cache_definition_paths_by_rid.find(rid);
    if (def_it == wb_rels.pivot_cache_definition_paths_by_rid.end()) {
      std::string ctx("context=ooxml_reader part=");
      ctx.append(workbook_path);
      ctx.append(" rid=");
      ctx.append(rid);
      return make_error(FormulonErrorCode::kIoRelationshipBroken,
                        "workbook.xml: <pivotCache> r:id has no matching workbook relationship", std::move(ctx));
    }
    const std::string& definition_path = def_it->second;
    if (!zip.has_entry(definition_path)) {
      std::string ctx("context=ooxml_reader cache_id=");
      ctx.append(std::to_string(cache_id));
      ctx.append(" definition_path=");
      ctx.append(definition_path);
      return make_error(FormulonErrorCode::kIoRelationshipBroken,
                        "pivotCacheDefinition: rel target missing from package", std::move(ctx));
    }
    auto def_bytes_or = zip.read_entry(definition_path);
    if (!def_bytes_or) {
      return def_bytes_or.error();
    }
    auto cache_or = read_pivot_cache_definition(def_bytes_or.value());
    if (!cache_or) {
      return cache_or.error();
    }
    pivot::PivotCache cache = std::move(cache_or.value());
    cache.set_cache_id(cache_id);
    consumed_parts.insert(definition_path);

    auto records_target_or = ooxml::load_pivot_cache_records_target(zip, definition_path);
    if (!records_target_or) {
      return records_target_or.error();
    }
    const std::string& records_path = records_target_or.value();
    if (!records_path.empty()) {
      if (!zip.has_entry(records_path)) {
        std::string ctx("context=ooxml_reader cache_id=");
        ctx.append(std::to_string(cache_id));
        ctx.append(" records_path=");
        ctx.append(records_path);
        return make_error(FormulonErrorCode::kIoRelationshipBroken,
                          "pivotCacheRecords: rel target missing from package", std::move(ctx));
      }
      auto rec_bytes_or = zip.read_entry(records_path);
      if (!rec_bytes_or) {
        return rec_bytes_or.error();
      }
      auto rec_status = read_pivot_cache_records(std::move(rec_bytes_or.value()), cache);
      if (!rec_status) {
        return rec_status.error();
      }
      consumed_parts.insert(records_path);
    }

    // The cache definition may carry its own rels file (it does when
    // there's a records part); mark it consumed so it does not surface
    // as an unknown part.
    const std::string def_rels_path = ooxml::rels_path_for_part(definition_path);
    if (zip.has_entry(def_rels_path)) {
      consumed_parts.insert(def_rels_path);
    }

    wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(std::move(cache)));
  }

  // Both pivot tables (per sheet) and their caches (workbook-level) are
  // now in memory. Resolve each table's field / item names against its
  // bound cache so GETPIVOTDATA can match a field by its source-column
  // name (the pivot-table part links to the cache only by index). The
  // loop body lives in the pivot layer to keep this reader hook minimal.
  pivot::resolve_all_pivot_names(wb);

  // Compute unknown_parts: every part the reader did not consume,
  // captured raw so the writer can re-emit it verbatim. Two sources:
  //   (1) `<Override>`-listed parts the reader did not model — captured
  //       with their declared content type so the writer replicates the
  //       `<Override>` registration.
  //   (2) Default-typed parts (vbaProject.bin, xl/media/*, drawings,
  //       VML, their rels, ...) declared only via `<Default Extension>`
  //       — captured with an empty content type; the writer relies on
  //       the round-tripped `<Default>` registration.
  std::vector<PassthroughPart> unknown_parts;
  unknown_parts.reserve(override_part_entries.size());
  for (const ooxml::OverrideEntry& entry : override_part_entries) {
    if (consumed_parts.find(entry.part_name) != consumed_parts.end()) {
      continue;
    }
    // Refuse to carry a hostile part name through passthrough: re-emitting
    // a `../` or absolute-shaped name would hand a downstream extractor a
    // zip-slip primitive on the round-tripped package.
    if (!ooxml::is_safe_part_name(entry.part_name)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "Override part name escapes package root; refusing to load",
                        "context=ooxml_reader part=" + entry.part_name);
    }
    // Read the bytes once. Failures here are propagated as ZIP errors;
    // the part was advertised in `[Content_Types].xml`, so if miniz
    // cannot extract it the package itself is corrupt.
    if (!zip.has_entry(entry.part_name)) {
      // Override referenced a part that does not exist in the archive.
      // We treat this as "nothing to passthrough" rather than fail: the
      // part is missing and there's no way to round-trip it. A future
      // bundle may surface a structured warning.
      continue;
    }
    auto bytes_or = zip.read_entry(entry.part_name);
    if (!bytes_or) {
      return bytes_or.error();
    }
    PassthroughPart part;
    part.path = entry.part_name;
    part.content_type = entry.content_type;
    part.bytes = std::move(bytes_or.value());
    unknown_parts.push_back(std::move(part));
  }
  for (PassthroughPart& part : extra_passthrough_parts) {
    auto duplicate = std::find_if(unknown_parts.begin(), unknown_parts.end(),
                                  [&part](const PassthroughPart& existing) { return existing.path == part.path; });
    if (duplicate == unknown_parts.end()) {
      unknown_parts.push_back(std::move(part));
    }
  }

  // Sweep every remaining archive entry the reader neither modelled nor
  // already captured as an Override passthrough. These are Default-typed
  // parts declared only via `<Default Extension>` — vbaProject.bin,
  // xl/media/* images, xl/drawings/* (bodies + their rels), VML
  // companions, and so on. Without this, real Excel-authored .xlsm /
  // .xlsx packages silently lose macros, images, shapes, and note
  // geometry on the first round-trip. Captured with an empty content
  // type; the writer relies on the round-tripped `<Default>` entries.
  {
    std::unordered_set<std::string> captured;
    captured.reserve(unknown_parts.size());
    for (const PassthroughPart& part : unknown_parts) {
      captured.insert(part.path);
    }
    for (const std::string& name : zip.list_entries()) {
      // Directory markers and empty names are not parts.
      if (name.empty() || name.back() == '/') {
        continue;
      }
      // Same zip-slip guard as the Override sweep: a Default-typed archive
      // entry with a traversal-shaped name must not be round-tripped.
      if (!ooxml::is_safe_part_name(name)) {
        return make_error(FormulonErrorCode::kIoZipSlip, "archive entry name escapes package root; refusing to load",
                          "context=ooxml_reader part=" + name);
      }
      if (consumed_parts.find(name) != consumed_parts.end()) {
        continue;
      }
      if (captured.find(name) != captured.end()) {
        continue;
      }
      auto bytes_or = zip.read_entry(name);
      if (!bytes_or) {
        return bytes_or.error();
      }
      PassthroughPart part;
      part.path = name;
      // Empty content type: the part is Default-typed, so the writer
      // must not emit a per-part `<Override>` for it.
      part.bytes = std::move(bytes_or.value());
      unknown_parts.push_back(std::move(part));
      captured.insert(name);
    }
  }

  // Stable order so callers / tests can compare deterministically.
  std::sort(unknown_parts.begin(), unknown_parts.end(),
            [](const PassthroughPart& a, const PassthroughPart& b) { return a.path < b.path; });

  // The workbook is the sole owner of the passthrough payload; the read
  // result does not mirror it. Handing it over by move keeps a package
  // with an 80 MB embedded image at one resident copy rather than two.
  wb.set_passthrough_parts(std::move(unknown_parts));
  wb.set_unknown_workbook_rels(std::move(wb_rels.unknown_rels));
  wb.set_unknown_package_rels(std::move(package_rels_or.value()));
  wb.set_default_content_types(std::move(default_content_types));

  OoxmlReadResult result{std::move(wb), pending_sst_count, diagnostics};
  return result;
}

Expected<OoxmlReadResult, Error> read_ooxml(ByteSpan bytes) {
  return ReadOoxmlWithThreshold(bytes, kSaxThresholdBytes);
}

namespace internal {

Expected<OoxmlReadResult, Error> ReadOoxmlWithSaxThresholdForTesting(ByteSpan bytes, std::size_t sax_threshold) {
  return ReadOoxmlWithThreshold(bytes, sax_threshold);
}

}  // namespace internal

}  // namespace io
}  // namespace formulon
