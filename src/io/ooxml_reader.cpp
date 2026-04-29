// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML reader implementation. Walks the in-memory package via
// `ZipReader`, parses the four package-structure parts with pugixml,
// builds a Workbook whose sheets reflect the `<sheets>` order from
// `xl/workbook.xml`, and then drives the per-sheet cell parser
// (`io::read_sheet_data`) so each `<c>` lands in the workbook with its
// formula registered against the recalc engine.
//
// Shared-strings (`xl/sharedStrings.xml`) resolution is wired in: the
// SST is loaded ahead of the per-sheet read loop, each sheet queues its
// `(row, col, sst_index)` tuples in a per-sheet `SheetReadContext`, and
// a final resolution pass replaces the placeholder `Text("")` values
// with views into the SST. Styles (`xl/styles.xml`) is parsed for
// validation only — the runtime style model lands when the formatter
// pipeline begins consuming it. Defined names and tables remain
// deferred.

#include "io/ooxml_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/styles_reader.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace {

// Relationship type URIs used by Excel-produced packages.
constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
constexpr std::string_view kRelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// Content types we expect to see referenced from `[Content_Types].xml`.
// We only look up the workbook content type to verify the package is
// well-formed; the full content-type registry is built in a later bundle
// (which will recognise the worksheet, styles, sharedStrings, etc.
// content types as well).
constexpr std::string_view kCtWorkbook = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

/// Returns the part path the package-level rels file points at for
/// `OfficeDocument`. The OOXML spec allows arbitrary placement (Excel
/// always uses `/xl/workbook.xml`), so we follow the relationship rather
/// than hard-coding the path. Path is normalised to drop any leading
/// slash so the result is directly consumable as a ZIP entry name.
Expected<std::string, Error> ResolveOfficeDocumentPath(const std::vector<std::uint8_t>& rels_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=_rels/.rels desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "package-level rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=ooxml_reader part=_rels/.rels");
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    if (type == kRelOfficeDocument) {
      std::string target = rel.attribute("Target").value();
      if (target.empty()) {
        return make_error(FormulonErrorCode::kIoRelationshipBroken,
                          "package-level rels: empty Target for OfficeDocument",
                          "context=ooxml_reader part=_rels/.rels");
      }
      // Normalise: relationship targets may be absolute (`/xl/...`) or
      // relative (`xl/...`). The ZIP catalogue stores names relative to
      // the package root with no leading slash.
      if (!target.empty() && target.front() == '/') {
        target.erase(0, 1);
      }
      return target;
    }
  }
  return make_error(FormulonErrorCode::kIoRelationshipBroken,
                    "package-level rels: no OfficeDocument relationship found",
                    "context=ooxml_reader part=_rels/.rels");
}

/// Verifies that `[Content_Types].xml` references the workbook content
/// type at least once. The full content-type registry is deferred to a
/// later bundle — for now we simply gate on "looks like a spreadsheet
/// package" so callers do not get half-built workbooks for non-Excel
/// archives that happen to contain a workbook.xml.
Expected<void, Error> VerifyContentTypes(const std::vector<std::uint8_t>& ct_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(ct_bytes.data(), ct_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=[Content_Types].xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "[Content_Types].xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=ooxml_reader part=[Content_Types].xml");
  }
  bool saw_workbook = false;
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    if (std::string_view(node.name()) == "Override") {
      const std::string_view ct = node.attribute("ContentType").value();
      if (ct == kCtWorkbook) {
        saw_workbook = true;
        break;
      }
    }
  }
  if (!saw_workbook) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "[Content_Types].xml: no workbook content-type override",
                      "context=ooxml_reader part=[Content_Types].xml");
  }
  return Expected<void, Error>::Ok();
}

/// Lists every part name advertised by `[Content_Types].xml`'s `<Override>`
/// elements. Used to compute the unknown-parts set; defaults are ignored
/// because they describe extensions, not specific parts.
std::vector<std::string> ListOverridePartNames(const std::vector<std::uint8_t>& ct_bytes) {
  std::vector<std::string> out;
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(ct_bytes.data(), ct_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    return out;
  }
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return out;
  }
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    if (std::string_view(node.name()) == "Override") {
      std::string part_name = node.attribute("PartName").value();
      if (!part_name.empty() && part_name.front() == '/') {
        part_name.erase(0, 1);
      }
      if (!part_name.empty()) {
        out.push_back(std::move(part_name));
      }
    }
  }
  return out;
}

/// Builds a path relative to `base_dir`. OOXML rels Target attributes are
/// relative to the part that owns the rels file (e.g. `xl/_rels/workbook.xml.rels`
/// has its Targets resolved against `xl/`). We collapse `..` segments so
/// `worksheets/sheet1.xml` resolved against `xl/` yields `xl/worksheets/sheet1.xml`,
/// and `../theme/theme1.xml` resolved against `xl/` would yield `theme/theme1.xml`.
std::string ResolveRelativePath(std::string_view base_dir, std::string_view target) {
  // Absolute-path targets short-circuit the relative resolution.
  if (!target.empty() && target.front() == '/') {
    return std::string(target.substr(1));
  }

  std::vector<std::string> stack;
  // Seed the stack from `base_dir`.
  std::size_t start = 0;
  for (std::size_t i = 0; i <= base_dir.size(); ++i) {
    if (i == base_dir.size() || base_dir[i] == '/') {
      if (i > start) {
        stack.emplace_back(base_dir.substr(start, i - start));
      }
      start = i + 1;
    }
  }
  // Append the target, applying `.` / `..` normalisation.
  start = 0;
  for (std::size_t i = 0; i <= target.size(); ++i) {
    if (i == target.size() || target[i] == '/') {
      if (i > start) {
        std::string_view seg = target.substr(start, i - start);
        if (seg == ".") {
          // skip
        } else if (seg == "..") {
          if (!stack.empty()) {
            stack.pop_back();
          }
        } else {
          stack.emplace_back(seg);
        }
      }
      start = i + 1;
    }
  }
  std::string out;
  for (std::size_t i = 0; i < stack.size(); ++i) {
    if (i > 0) {
      out.push_back('/');
    }
    out.append(stack[i]);
  }
  return out;
}

/// Returns the directory portion of `path` (everything up to the last
/// `/`, exclusive). Empty for top-level paths like `_rels/.rels`. Used as
/// the base directory for resolving relative relationship targets.
std::string DirOf(std::string_view path) {
  const std::size_t pos = path.find_last_of('/');
  if (pos == std::string_view::npos) {
    return {};
  }
  return std::string(path.substr(0, pos));
}

/// Aggregated workbook-relationship lookup: per-sheet `rId -> path` map
/// plus optional resolved paths for the `sharedStrings` and `styles`
/// parts. Empty `sst_path` / `styles_path` mean "no such relationship",
/// which is legal — the package can omit either part.
struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;
  std::string sst_path;
  std::string styles_path;
};

/// Loads `<workbook_dir>/_rels/<workbook_filename>.rels` (if present) and
/// returns the relationship-id -> resolved target-path map for each
/// worksheet relationship, plus the resolved target paths for the
/// shared-strings and styles relationships when present.
Expected<WorkbookRels, Error> LoadWorkbookRels(const ZipReader& zip, std::string_view workbook_path) {
  WorkbookRels rels;

  // workbook_path = "xl/workbook.xml" => rels = "xl/_rels/workbook.xml.rels"
  const std::size_t slash = workbook_path.find_last_of('/');
  std::string rels_path;
  if (slash == std::string_view::npos) {
    rels_path.append("_rels/").append(workbook_path).append(".rels");
  } else {
    rels_path.append(workbook_path.substr(0, slash));
    rels_path.append("/_rels/");
    rels_path.append(workbook_path.substr(slash + 1));
    rels_path.append(".rels");
  }

  if (!zip.has_entry(rels_path)) {
    // Excel always emits this; treat absence as a broken package.
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=ooxml_reader rels_path=" + rels_path);
  }
  auto rels_bytes_or = zip.read_entry(rels_path);
  if (!rels_bytes_or) {
    return rels_bytes_or.error();
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "workbook rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(rels_path);
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: missing <Relationships>",
                      std::move(ctx));
  }

  const std::string base_dir = DirOf(workbook_path);
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    if (target.empty()) {
      continue;
    }
    if (type == kRelWorksheet) {
      const std::string id = rel.attribute("Id").value();
      if (id.empty()) {
        continue;
      }
      rels.sheet_targets.emplace(id, ResolveRelativePath(base_dir, target));
    } else if (type == kRelSharedStrings) {
      // Last writer wins on duplicates (Excel never emits more than one,
      // but defending against malformed inputs costs almost nothing).
      rels.sst_path = ResolveRelativePath(base_dir, target);
    } else if (type == kRelStyles) {
      rels.styles_path = ResolveRelativePath(base_dir, target);
    }
  }
  return rels;
}

}  // namespace

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
  if (auto v = VerifyContentTypes(ct_bytes); !v) {
    return v.error();
  }
  const std::vector<std::string> override_part_names = ListOverridePartNames(ct_bytes);

  // 2. _rels/.rels — locate the workbook part path.
  if (!zip.has_entry("_rels/.rels")) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "_rels/.rels: missing from package",
                      "context=ooxml_reader");
  }
  auto root_rels_or = zip.read_entry("_rels/.rels");
  if (!root_rels_or) {
    return root_rels_or.error();
  }
  auto wb_path_or = ResolveOfficeDocumentPath(root_rels_or.value());
  if (!wb_path_or) {
    return wb_path_or.error();
  }
  const std::string workbook_path = wb_path_or.value();

  // 3. xl/_rels/workbook.xml.rels — load and validate. We need both the
  // sheet rId -> part-path map (for the per-sheet read loop below) and
  // the resolved paths for the sharedStrings / styles parts so we can
  // load them at the right point in the pipeline.
  auto wb_rels_or = LoadWorkbookRels(zip, workbook_path);
  if (!wb_rels_or) {
    return wb_rels_or.error();
  }
  const WorkbookRels& wb_rels = wb_rels_or.value();

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
  pugi::xml_parse_result wb_parse =
      wb_doc.load_buffer(wb_bytes.data(), wb_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!wb_parse) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(workbook_path);
    ctx.append(" desc=");
    ctx.append(wb_parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "workbook.xml: pugixml parse failed", std::move(ctx));
  }
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
  // Sheet1 to overwrite or remove.
  Workbook wb = Workbook::create_empty();

  // Track which sheet relationships were consumed, so the unknown-parts
  // computation can subtract them.
  std::unordered_set<std::string> consumed_parts;
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  // Also mark the workbook rels file as consumed.
  {
    const std::size_t slash = workbook_path.find_last_of('/');
    std::string rels_path;
    if (slash == std::string::npos) {
      rels_path.append("_rels/").append(workbook_path).append(".rels");
    } else {
      rels_path.append(workbook_path.substr(0, slash));
      rels_path.append("/_rels/");
      rels_path.append(workbook_path.substr(slash + 1));
      rels_path.append(".rels");
    }
    consumed_parts.insert(std::move(rels_path));
  }

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
    std::string rid = sn.attribute("r:id").value();
    if (rid.empty()) {
      rid = sn.attribute("id").value();
    }
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
    wb.add_sheet(std::move(name));
    sheet_part_paths.push_back(it->second);
    consumed_parts.insert(it->second);
  }

  if (sheet_part_paths.empty()) {
    return make_error(FormulonErrorCode::kIoSheetCorrupt, "workbook.xml: empty <sheets> list",
                      "context=ooxml_reader part=" + workbook_path);
  }

  // result_text_storage is the workbook-lifetime backing store for every
  // string the reader owns: inline-string `<is>` payloads decoded by the
  // sheet reader AND every entry in the shared-string table loaded
  // below. A `std::deque` (pointer-stable across appends) is required so
  // the `string_view`s handed to cells remain valid as later strings are
  // appended.
  std::deque<std::string> result_text_storage;

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
    // Result is intentionally discarded; see read_styles() docs.
    (void)styles_or.value();
    consumed_parts.insert(wb_rels.styles_path);
  }
  if (zip.has_entry("xl/styles.xml")) {
    consumed_parts.insert("xl/styles.xml");
  }

  // 5. Read each sheet's <sheetData> via the cell-aware sheet reader.
  // The sheet reader appends every inline-string payload directly into
  // `result_text_storage`, a pointer-stable `std::deque` whose lifetime
  // is the read result. Cells therefore hold `Value::text` views that
  // remain valid as long as the `OoxmlReadResult` is alive — no
  // post-load repointing is needed.
  //
  // We retain each sheet's `pending_sst_cells` so the SST resolution
  // pass below can rewrite the placeholder text values without rerunning
  // the sheet walk.
  std::vector<SheetReadContext> sheet_contexts(sheet_part_paths.size());
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
    pugi::xml_document sheet_doc;
    pugi::xml_parse_result sheet_parse =
        sheet_doc.load_buffer(sheet_bytes.data(), sheet_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
    if (!sheet_parse) {
      std::string ctx("context=ooxml_reader part=");
      ctx.append(sheet_path);
      ctx.append(" desc=");
      ctx.append(sheet_parse.description());
      return make_error(FormulonErrorCode::kIoXmlParse, "sheet*.xml: pugixml parse failed", std::move(ctx));
    }
    auto rs = read_sheet_data(sheet_doc, i, wb, sheet_contexts[i], result_text_storage);
    if (!rs) {
      return rs.error();
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
      ++pending_sst_count;
    }
  }

  // Compute unknown_parts: every Override-listed part name we did not
  // touch above. Bundle 2.5 will refine the categorisation; for now this
  // keeps round-trip work in scope without trying to be clever.
  std::vector<std::string> unknown_parts;
  unknown_parts.reserve(override_part_names.size());
  for (const std::string& part : override_part_names) {
    if (consumed_parts.find(part) == consumed_parts.end()) {
      unknown_parts.push_back(part);
    }
  }
  std::sort(unknown_parts.begin(), unknown_parts.end());

  OoxmlReadResult result{std::move(wb), std::move(unknown_parts), pending_sst_count, std::move(result_text_storage)};
  return result;
}

}  // namespace io
}  // namespace formulon
