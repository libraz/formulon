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

#include "calc_settings.h"
#include "default_content_type.h"
#include "defined_name.h"
#include "io/cf_reader.h"
#include "io/comments_reader.h"
#include "io/defined_names_internal.h"
#include "io/dynamic_array_formula.h"
#include "io/ooxml/external_link_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/pivot_target_reader.h"
#include "io/ooxml/sheet_aux_rels_reader.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/ooxml/worksheet_metadata_reader.h"
#include "io/ooxml_defs.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_table_reader.h"
#include "io/sheet_overlay_reader.h"
#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/threaded_comments_io.h"
#include "io/workbook_kind_ooxml.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "io/xsd_double.h"
#include "io/zip_reader.h"
#include "passthrough_part.h"
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
  const WorkbookKind workbook_kind = kind_or.value();
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
  ASSIGN_OR_RETURN(auto root_rels, zip.read_entry("_rels/.rels"));
  ASSIGN_OR_RETURN(auto wb_path, ooxml::resolve_office_document_path(root_rels));
  const std::string workbook_path = wb_path;
  ASSIGN_OR_RETURN(auto package_rels, ReadUnknownPackageRels(root_rels));

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
  // The package's style table is authoritative, including when the package
  // has none: back-filling the factory's seeded defaults would invent style
  // records the file never carried and change how a dangling `s=` resolves.
  wb.set_styles(StylesTable{});

  // Capture the `<workbook>` root's extra namespace declarations (and
  // `mc:Ignorable`) so the writer can re-emit them. The raw `<bookViews>`
  // capture below can carry namespaced attributes (e.g. `xr2:uid`) whose
  // prefixes are declared only on this root; without re-declaring them the
  // re-emitted fragment is malformed XML and Excel refuses the file. The
  // same helper handles the `<worksheet>` root (see `ooxml::apply_worksheet_metadata`).
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
    // Workbook-side `<sheet state="...">` attribute — Excel records sheet
    // visibility here, not on the worksheet part. All three `ST_SheetState`
    // values are carried across, because `veryHidden` is what keeps a sheet
    // out of Excel's "Unhide" dialog and folding it to `hidden` would put
    // it back within a user's reach. The worksheet-side
    // `<sheetPr><tabHidden/>` form is handled by the per-sheet reader call
    // below; the merge across both paths is OR-style, and only ever
    // strengthens the state the workbook part stated.
    const std::string_view state = sn.attribute("state").value();
    const SheetVisibility workbook_visibility = state == "veryHidden" ? SheetVisibility::kVeryHidden
                                                : state == "hidden"   ? SheetVisibility::kHidden
                                                                      : SheetVisibility::kVisible;
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
    if (workbook_visibility != SheetVisibility::kVisible) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().set_visibility(workbook_visibility);
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
    IterativeOptions opts;
    if (pugi::xml_attribute iterate = calc_pr.attribute("iterate"); iterate) {
      opts.enabled = parse_xml_bool(iterate.value());
    }
    if (pugi::xml_attribute count = calc_pr.attribute("iterateCount"); count) {
      const long long parsed = count.as_llong(static_cast<long long>(kDefaultMaxIterations));
      // Clamp into Excel's own dialog range. The upper bound matters more
      // than the lower one: the file decides how much work the first
      // `recalc()` performs, the solver has no wall-clock limit, and the
      // cancellation callback is null unless the host opted in. Without
      // this, `iterateCount="4294967295"` on a two-cell cycle is an
      // unrecoverable hang rather than a slow load.
      opts.max_iterations = parsed < 1                                           ? 1U
                            : parsed > static_cast<long long>(kMaxIterationsCap) ? kMaxIterationsCap
                                                                                 : static_cast<std::uint32_t>(parsed);
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
    ASSIGN_OR_RETURN(auto sst_bytes, zip.read_entry(wb_rels.sst_path));
    auto sst_or = read_shared_strings(std::move(sst_bytes), result_text_storage);
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

  // Persons: the person list the sheets' threaded comments name. The part
  // is model-owned, so its relationship leaves the unknown list with it.
  for (auto it = wb_rels.unknown_rels.begin(); it != wb_rels.unknown_rels.end(); ++it) {
    if (it->type != kRelPerson || it->target_external || !zip.has_entry(it->target)) {
      continue;
    }
    ASSIGN_OR_RETURN(auto pb, zip.read_entry(it->target));
    ASSIGN_OR_RETURN(auto persons, read_persons(pb));
    wb.mutable_persons() = std::move(persons);
    consumed_parts.insert(it->target);
    wb_rels.unknown_rels.erase(it);
    break;
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
    ASSIGN_OR_RETURN(auto styles_bytes, zip.read_entry(wb_rels.styles_path));
    ASSIGN_OR_RETURN(auto styles, read_styles(styles_bytes));
    wb.set_styles(std::move(styles));
    consumed_parts.insert(wb_rels.styles_path);
  }
  if (zip.has_entry("xl/styles.xml")) {
    consumed_parts.insert("xl/styles.xml");
  }

  // 4c. External links. Joins `<externalReferences>` against the
  // workbook rels and the per-link rels files. Loaded before any formula,
  // because ingestion spells each `[N]` from these records and appends a
  // record for a book none of them matches. The body parts continue to
  // round-trip through `passthrough_parts()`; only the per-link rels files
  // are marked as consumed (their content is regenerated by the writer from
  // the captured records). Missing parts and successfully-read malformed XML
  // remain failure-tolerant, while extraction failures are returned
  // unchanged.
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
        auto rs = read_sheet_data_sax(sheet_span, i, wb, sheet_contexts[i], result_text_storage, &diagnostics);
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
      shell = ooxml::build_worksheet_shell_bytes(sheet_bytes);
      RETURN_IF_ERROR(load_xml_buffer_inplace(sheet_doc, shell, "ooxml_reader", "sheet*.xml (metadata shell)"));
    } else {
      RETURN_IF_ERROR(load_xml_buffer_inplace(sheet_doc, sheet_bytes, "ooxml_reader", "sheet*.xml"));
      auto rs = read_sheet_data(sheet_doc, i, wb, sheet_contexts[i], result_text_storage, &diagnostics);
      if (!rs) {
        return rs.error();
      }
    }
    RETURN_IF_ERROR(ooxml::apply_worksheet_metadata(sheet_doc, i, wb, &diagnostics));

    // Sheet rels file (`xl/worksheets/_rels/sheetN.xml.rels`) — drives
    // the table-part and pivot-table lookups. Optional: most sheets have
    // no rels at all. Tables and pivot tables are read in two passes
    // through the same rels file (separate helpers, each scoped to one
    // relationship type) so each consumer site reads linearly.
    const std::string sheet_rels_path = ooxml::rels_path_for_part(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      const std::string sheet_dir = ooxml::dir_of(sheet_path);
      ASSIGN_OR_RETURN(auto targets, ooxml::load_sheet_table_targets(zip, sheet_rels_path, sheet_dir));
      consumed_parts.insert(sheet_rels_path);
      for (const std::string& table_path : targets) {
        if (!zip.has_entry(table_path)) {
          std::string ctx("context=ooxml_reader sheet_index=");
          ctx.append(std::to_string(i));
          ctx.append(" table_path=").append(table_path);
          return make_error(FormulonErrorCode::kIoRelationshipBroken, "table: rel target missing from package",
                            std::move(ctx));
        }
        ASSIGN_OR_RETURN(auto table_bytes, zip.read_entry(table_path));
        ASSIGN_OR_RETURN(auto table, read_table(table_bytes, i));
        tables_metadata.push_back(std::move(table));
        consumed_parts.insert(table_path);
      }

      // Pivot tables anchored on this sheet. Each part feeds into the
      // pivot-table reader and is attached to the owning sheet; the
      // workbook-level pivot caches are loaded after the sheet loop.
      ASSIGN_OR_RETURN(auto pivot_targets, ooxml::load_sheet_pivot_table_targets(zip, sheet_rels_path, sheet_dir));
      for (const std::string& pivot_table_path : pivot_targets) {
        if (!zip.has_entry(pivot_table_path)) {
          std::string ctx("context=ooxml_reader sheet_index=");
          ctx.append(std::to_string(i));
          ctx.append(" pivot_table_path=").append(pivot_table_path);
          return make_error(FormulonErrorCode::kIoRelationshipBroken, "pivotTable: rel target missing from package",
                            std::move(ctx));
        }
        ASSIGN_OR_RETURN(auto pt_bytes, zip.read_entry(pivot_table_path));
        ASSIGN_OR_RETURN(auto pt, read_pivot_table_definition(pt_bytes));
        wb.sheet(i).add_pivot_table(std::make_unique<pivot::PivotTable>(std::move(pt)));
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
          ASSIGN_OR_RETURN(auto pb, zip.read_entry(aux.printer_settings_path));
          PassthroughPart part;
          part.path = aux.printer_settings_path;
          part.content_type = std::string(kCtPrinterSettings);
          part.bytes = std::move(pb);
          extra_passthrough_parts.push_back(std::move(part));
          consumed_parts.insert(aux.printer_settings_path);
        }
      }

      // Threaded comments load before the legacy comments part so the
      // thread stubs in it can be told apart from notes.
      if (!aux.threaded_comments_path.empty() && zip.has_entry(aux.threaded_comments_path)) {
        ASSIGN_OR_RETURN(auto tb, zip.read_entry(aux.threaded_comments_path));
        ASSIGN_OR_RETURN(auto threads, read_threaded_comments(tb));
        wb.sheet(i).mutable_threaded_comments() = std::move(threads);
        consumed_parts.insert(aux.threaded_comments_path);
      }

      // Comments part: load + attach. The VML drawing companion is
      // intentionally NOT consumed here so the bytes flow through the
      // unknown-parts passthrough mechanism unchanged; the anchors it was
      // read with let the writer tell whether those bytes are still current.
      if (!aux.comments_path.empty() && zip.has_entry(aux.comments_path)) {
        ASSIGN_OR_RETURN(auto cb, zip.read_entry(aux.comments_path));
        ASSIGN_OR_RETURN(auto comments, read_comments(cb));
        drop_thread_stubs(comments, wb.sheet(i).threaded_comments());
        wb.sheet(i).mutable_comments() = std::move(comments);
        wb.sheet(i).set_comment_vml_path(aux.vml_path);
        wb.sheet(i).set_comment_vml_anchors(wb.sheet(i).comment_anchor_set());
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

  // 5b. Conditional-format and data-validation formulas name other books
  // through the same `[N]` cell formulas do.
  ingest_feature_formulas(wb);

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
      // so PHONETIC() can surface the kana over the characters it
      // covers. The SST reader keeps `phonetic_for_entries` parallel to
      // `entries`, empty for unannotated entries; we only commit a
      // non-empty annotation to avoid touching `Cell::phonetic_runs` for
      // the common unannotated case. Several cells may share one `<si>`,
      // so the runs are copied rather than moved out.
      if (idx < sst.phonetic_for_entries.size() && !sst.phonetic_for_entries[idx].empty()) {
        wb.sheet(i).set_cell_phonetic_runs(row, col, sst.phonetic_for_entries[idx]);
        if (idx < sst.phonetic_props_for_entries.size()) {
          wb.sheet(i).set_cell_phonetic_props(row, col, sst.phonetic_props_for_entries[idx]);
        }
      }
      ++pending_sst_count;
    }
  }

  // 6b. Register each sheet's dynamic-array spill regions, now that every
  // SST reference is resolved. This must run after step 6: a spill
  // region's phantom cells capture their footprint's `cached_value` at
  // registration time, so registering before SST resolution would freeze
  // an SST-typed phantom's unresolved `Text("")` placeholder into the
  // region rather than the string it actually holds (see
  // `SheetReadContext::array_anchors`).
  for (std::size_t i = 0; i < sheet_contexts.size(); ++i) {
    if (auto r = RegisterArraySpills(wb.sheet(i), sheet_contexts[i].array_anchors); !r) {
      return r.error();
    }
  }

  // 6c. A formula is a dynamic-array formula exactly when its `cm=` names
  // the XLDAPR entry of `xl/metadata.xml`; any other (a legacy CSE block, an
  // implicit-intersection formula) keeps its legacy meaning.
  {
    std::uint32_t xldapr = 0U;
    if (auto metadata = zip.read_entry("xl/metadata.xml"); metadata) {
      xldapr = xldapr_cell_metadata_index(metadata.value());
    }
    for (std::size_t i = 0; i < sheet_contexts.size(); ++i) {
      std::unordered_set<std::uint64_t> marked;
      for (const auto& [row, col, cm] : sheet_contexts[i].cell_metadata) {
        if (xldapr != 0U && cm == xldapr) {
          marked.insert(dynamic_array_cell_key(row, col));
        }
      }
      apply_loaded_dynamic_array_marks(wb.sheet(i), marked);
    }
  }

  // 7. Defined names, read from `wb_doc` after sheet construction and resolved lazily at evaluation time.
  ASSIGN_OR_RETURN(auto defined_names, read_defined_names(wb_doc));
  for (DefinedName& entry : defined_names) {
    entry.formula = wb.ingest_stored_formula(entry.formula);
  }
  wb.set_defined_names(std::move(defined_names));

  // 8. Tables — already accumulated in the per-sheet loop above. Move
  // the workbook-scope vector onto the workbook for round-trip.
  wb.set_tables(std::move(tables_metadata));

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
    ASSIGN_OR_RETURN(auto def_bytes, zip.read_entry(definition_path));
    auto cache_or = read_pivot_cache_definition(def_bytes);
    if (!cache_or) {
      return cache_or.error();
    }
    pivot::PivotCache cache = std::move(cache_or.value());
    cache.set_cache_id(cache_id);
    consumed_parts.insert(definition_path);

    ASSIGN_OR_RETURN(auto records_target, ooxml::load_pivot_cache_records_target(zip, definition_path));
    const std::string& records_path = records_target;
    if (!records_path.empty()) {
      if (!zip.has_entry(records_path)) {
        std::string ctx("context=ooxml_reader cache_id=");
        ctx.append(std::to_string(cache_id));
        ctx.append(" records_path=");
        ctx.append(records_path);
        return make_error(FormulonErrorCode::kIoRelationshipBroken,
                          "pivotCacheRecords: rel target missing from package", std::move(ctx));
      }
      ASSIGN_OR_RETURN(auto rec_bytes, zip.read_entry(records_path));
      auto rec_status = read_pivot_cache_records(std::move(rec_bytes), cache);
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
    ASSIGN_OR_RETURN(auto part_bytes, zip.read_entry(entry.part_name));
    PassthroughPart part;
    part.path = entry.part_name;
    part.content_type = entry.content_type;
    part.bytes = std::move(part_bytes);
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
      ASSIGN_OR_RETURN(auto part_bytes, zip.read_entry(name));
      PassthroughPart part;
      part.path = name;
      // Empty content type: the part is Default-typed, so the writer
      // must not emit a per-part `<Override>` for it.
      part.bytes = std::move(part_bytes);
      unknown_parts.push_back(std::move(part));
      captured.insert(name);
    }
  }

  // Stable order so callers / tests can compare deterministically.
  sort_passthrough_parts(unknown_parts);

  // The workbook is the sole owner of the passthrough payload; the read
  // result does not mirror it. Handing it over by move keeps a package
  // with an 80 MB embedded image at one resident copy rather than two.
  wb.set_passthrough_parts(std::move(unknown_parts));
  wb.set_unknown_workbook_rels(std::move(wb_rels.unknown_rels));
  wb.set_unknown_package_rels(std::move(package_rels));
  wb.set_default_content_types(std::move(default_content_types));

  wb.apply_legacy_implicit_intersections();

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
