// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
#include "io/defined_names.h"
#include "io/defined_names_internal.h"
#include "io/ooxml/external_link_reader.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/pivot_target_reader.h"
#include "io/ooxml/sheet_aux_rels_reader.h"
#include "io/ooxml/workbook_rels_reader.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_table_reader.h"
#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/xml_utils.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
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

// Content type for the binary printer-settings part; the orchestrator
// stamps it onto the passthrough record when it captures a sheet's
// printer-settings bytes (the sheet aux-rels reader only resolves the
// path; the orchestrator owns the round-trip wrapping).
constexpr std::string_view kCtPrinterSettings =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";

struct StringXmlWriter final : pugi::xml_writer {
  std::string* dst = nullptr;
  void write(const void* data, size_t size) override {
    if (dst != nullptr) {
      dst->append(static_cast<const char*>(data), size);
    }
  }
};

std::string RawXml(const pugi::xml_node& node) {
  std::string out;
  StringXmlWriter sink;
  sink.dst = &out;
  node.print(sink, /*indent=*/"", pugi::format_raw);
  return out;
}

/// Populates the structured `PageSetup` fields from a `<pageSetup>` node.
/// The raw XML string remains the writer's source of truth; this is an
/// additive parse for consumers that need typed access. Missing
/// attributes keep the struct defaults. `fit_to_page` is set separately
/// from `<sheetPr><pageSetUpPr>`.
void ApplyStructuredPageSetup(const pugi::xml_node& page_setup, PageSetup& out) {
  if (pugi::xml_attribute attr = page_setup.attribute("orientation"); attr) {
    const std::string_view value = attr.value();
    if (value == "portrait") {
      out.orientation = Orientation::kPortrait;
    } else if (value == "landscape") {
      out.orientation = Orientation::kLandscape;
    } else {
      out.orientation = Orientation::kDefault;
    }
  }
  if (pugi::xml_attribute attr = page_setup.attribute("paperSize"); attr) {
    out.paper_size = static_cast<std::uint32_t>(attr.as_uint(out.paper_size));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("scale"); attr) {
    out.scale = static_cast<std::uint32_t>(attr.as_uint(out.scale));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("fitToWidth"); attr) {
    out.fit_to_width = static_cast<std::uint32_t>(attr.as_uint(out.fit_to_width));
  }
  if (pugi::xml_attribute attr = page_setup.attribute("fitToHeight"); attr) {
    out.fit_to_height = static_cast<std::uint32_t>(attr.as_uint(out.fit_to_height));
  }
}

/// Populates the structured `PageMargins` fields from a `<pageMargins>`
/// node. Additive alongside the raw XML string; missing attributes keep
/// the struct defaults.
void ApplyStructuredPageMargins(const pugi::xml_node& page_margins, PageMargins& out) {
  out.left = page_margins.attribute("left").as_double(out.left);
  out.right = page_margins.attribute("right").as_double(out.right);
  out.top = page_margins.attribute("top").as_double(out.top);
  out.bottom = page_margins.attribute("bottom").as_double(out.bottom);
  out.header = page_margins.attribute("header").as_double(out.header);
  out.footer = page_margins.attribute("footer").as_double(out.footer);
}

/// Reads the `<brk>` children of a `<rowBreaks>` / `<colBreaks>` node
/// into `out`. OOXML stores the break index 1-based in the `id`
/// attribute; this normalises to 0-based (clamping at 0). The
/// `count` / `manualBreakCount` wrapper attributes are ignored — only
/// the `<brk>` entries themselves are honoured.
void ReadManualBreaks(const pugi::xml_node& breaks_node, std::vector<ManualBreak>& out) {
  if (!breaks_node) {
    return;
  }
  for (pugi::xml_node brk = breaks_node.child("brk"); brk; brk = brk.next_sibling("brk")) {
    ManualBreak entry;
    const unsigned int raw_id = brk.attribute("id").as_uint(0);
    entry.id = raw_id > 0U ? raw_id - 1U : 0U;
    entry.min = static_cast<std::uint32_t>(brk.attribute("min").as_uint(0));
    entry.max = static_cast<std::uint32_t>(brk.attribute("max").as_uint(0));
    if (pugi::xml_attribute man = brk.attribute("man"); man) {
      entry.manual = man.as_bool(true);
    }
    out.push_back(entry);
  }
}

}  // namespace

namespace internal {

Expected<std::string, Error> ResolveRelativePathForTesting(std::string_view base_dir, std::string_view target) {
  return ooxml::resolve_relative_path(base_dir, target);
}

}  // namespace internal

Expected<OoxmlReadResult, Error> read_ooxml(ByteSpan bytes) {
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
  auto kind_or = ooxml::verify_content_types(ct_bytes);
  if (!kind_or) {
    return kind_or.error();
  }
  const io::WorkbookKind workbook_kind = kind_or.value();
  auto override_part_entries_or = ooxml::list_override_part_entries(ct_bytes);
  if (!override_part_entries_or) {
    return override_part_entries_or.error();
  }
  const std::vector<ooxml::OverrideEntry> override_part_entries = std::move(override_part_entries_or.value());

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
  const std::unordered_map<std::string, std::string>& sheet_rels = wb_rels.sheet_targets;
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
    wb.add_sheet(std::move(name));
    if (workbook_hides) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().tab_hidden = true;
    }
    sheet_part_paths.push_back(it->second);
    consumed_parts.insert(it->second);
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
    if (pugi::xml_attribute delta = calc_pr.attribute("iterateDelta"); delta) {
      opts.max_change = delta.as_double(eval::kDefaultMaxChange);
    }
    wb.set_iterative_options(opts);
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
    auto sst_or = read_shared_strings(sst_bytes_or.value(), result_text_storage);
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
    const std::vector<std::uint8_t>& sheet_bytes = sheet_bytes_or.value();
    // Choose the read path by raw XML size: small sheets stay on the
    // pugixml DOM path (familiar code, well-validated); large sheets
    // (>= `kSaxThresholdBytes`) stream through the SAX scanner so a
    // 1M-cell worksheet does not need to materialise as a DOM in
    // memory. Both paths produce identical Workbook output.
    //
    // The threshold is a compile-time constant. On WASM the value is
    // `SIZE_MAX` (see `sheet_reader.h`) so the SAX branch is
    // statically dead and the linker removes the streaming scanner
    // entirely — saving ~17 KiB of `.wasm`. The `if constexpr` makes
    // the elimination explicit so this is robust under -O0 / -Og too.
    constexpr bool kSaxEnabled = kSaxThresholdBytes != static_cast<std::size_t>(-1);
    bool sax_used = false;
    if constexpr (kSaxEnabled) {
      if (sheet_bytes.size() >= kSaxThresholdBytes) {
        ByteSpan sheet_span{sheet_bytes.data(), sheet_bytes.size()};
        auto rs = read_sheet_data_sax(sheet_span, i, wb, sheet_contexts[i], result_text_storage);
        if (!rs) {
          return rs.error();
        }
        sax_used = true;
      }
    }
    if (!sax_used) {
      pugi::xml_document sheet_doc;
      RETURN_IF_ERROR(load_xml_buffer(sheet_doc, sheet_bytes, "ooxml_reader", "sheet*.xml"));
      auto rs = read_sheet_data(sheet_doc, i, wb, sheet_contexts[i], result_text_storage);
      if (!rs) {
        return rs.error();
      }
      // Conditional-format blocks live at the top level of <worksheet>
      // (siblings to <sheetData>), so they ride the same DOM the cell
      // reader just consumed. The SAX path skips this scan; loading a
      // second DOM purely for CF on multi-MB sheets would defeat the
      // SAX optimisation, so SAX-side CF reading is deferred to a
      // follow-up PR. In practice the sheets that exercise the SAX
      // threshold (>256 KiB) are dominated by row data, not CF blocks,
      // so the missed coverage is small.
      auto cfs_or = read_conditional_formats(sheet_doc.child("worksheet"));
      if (!cfs_or) {
        return cfs_or.error();
      }
      wb.sheet(i).mutable_conditional_formats() = std::move(cfs_or.value());
      // View / layout metadata (`<sheetView>`, `<sheetPr>`, `<cols>`,
      // per-row overrides) lives at the same DOM level — process it
      // here so the round-trip writer can reproduce it. SAX-side
      // coverage is deferred for the same reason as CF.
      auto view_layout_or = read_sheet_view_and_layout(sheet_doc, i, wb);
      if (!view_layout_or) {
        return view_layout_or.error();
      }

      // Merge / hyperlink / validation blocks all live at the top
      // level of <worksheet>, same as CF. Same DOM-only reasoning
      // applies.
      auto merges_or = read_merges(sheet_doc.child("worksheet"));
      if (!merges_or) {
        return merges_or.error();
      }
      wb.sheet(i).mutable_merges() = std::move(merges_or.value());

      auto hls_or = read_hyperlinks(sheet_doc.child("worksheet"));
      if (!hls_or) {
        return hls_or.error();
      }
      wb.sheet(i).mutable_hyperlinks() = std::move(hls_or.value());

      auto dvs_or = read_data_validations(sheet_doc.child("worksheet"));
      if (!dvs_or) {
        return dvs_or.error();
      }
      wb.sheet(i).mutable_validations() = std::move(dvs_or.value());

      // `<sheetProtection>` is a single optional element; the reader
      // never fails for it (default-on-error). Empty = enabled false.
      wb.sheet(i).mutable_protection() = read_sheet_protection(sheet_doc.child("worksheet"));

      SheetPrintSettings& print = wb.sheet(i).mutable_print_settings();
      pugi::xml_node worksheet = sheet_doc.child("worksheet");
      if (pugi::xml_node sheet_pr = worksheet.child("sheetPr"); sheet_pr) {
        if (pugi::xml_node page_setup_pr = sheet_pr.child("pageSetUpPr"); page_setup_pr) {
          print.sheet_pr_xml = RawXml(sheet_pr);
          print.page_setup.fit_to_page = page_setup_pr.attribute("fitToPage").as_bool(false);
        }
      }
      if (pugi::xml_node page_margins = worksheet.child("pageMargins")) {
        print.page_margins_xml = RawXml(page_margins);
        ApplyStructuredPageMargins(page_margins, print.page_margins);
      }
      if (pugi::xml_node page_setup = worksheet.child("pageSetup")) {
        print.page_setup_xml = RawXml(page_setup);
        print.printer_settings_rid = ooxml::relationship_ref_id(page_setup);
        ApplyStructuredPageSetup(page_setup, print.page_setup);
      }
      // Manual page breaks. `<rowBreaks>` / `<colBreaks>` are otherwise
      // dropped; capture them structurally so a save cycle preserves them.
      ReadManualBreaks(worksheet.child("rowBreaks"), print.manual_row_breaks);
      ReadManualBreaks(worksheet.child("colBreaks"), print.manual_col_breaks);
    }

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
      // empty in the result.
      auto aux_or = ooxml::load_sheet_aux_rels(zip, sheet_rels_path, sheet_dir);
      if (!aux_or) {
        return aux_or.error();
      }
      const ooxml::SheetAuxRels& aux = aux_or.value();
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
        consumed_parts.insert(aux.comments_path);
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
      wb.sheet(i).set_cell_cached_value(row, col, Value::text(sst.entries[idx]));
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
  // writer from the captured records). Failure-tolerant: malformed
  // sections produce `kUnknown` records rather than failing the load.
  {
    ooxml::ExternalLinkLoadResult ext = ooxml::load_external_links(zip, wb_root, wb_rels);
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
      auto rec_status = read_pivot_cache_records(rec_bytes_or.value(), cache);
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

  // Compute unknown_parts: every Override-listed part the reader did
  // not consume, captured raw so the writer can re-emit it verbatim.
  // Default-typed parts (images, OLE) are out of scope at this layer;
  // see `passthrough_part.h` for the rationale.
  std::vector<PassthroughPart> unknown_parts;
  unknown_parts.reserve(override_part_entries.size());
  for (const ooxml::OverrideEntry& entry : override_part_entries) {
    if (consumed_parts.find(entry.part_name) != consumed_parts.end()) {
      continue;
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
  // Stable order so callers / tests can compare deterministically.
  std::sort(unknown_parts.begin(), unknown_parts.end(),
            [](const PassthroughPart& a, const PassthroughPart& b) { return a.path < b.path; });

  // Hand the same payload to the workbook so the writer can find it
  // even when the caller only retains `result.workbook`. The
  // `OoxmlReadResult::unknown_parts` view stays populated for tests and
  // tooling that want to inspect what was preserved.
  wb.set_passthrough_parts(unknown_parts);
  wb.set_unknown_workbook_rels(std::move(wb_rels.unknown_rels));

  OoxmlReadResult result{std::move(wb), std::move(unknown_parts), pending_sst_count};
  return result;
}

}  // namespace io
}  // namespace formulon
