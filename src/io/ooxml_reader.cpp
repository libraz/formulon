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
#include <memory>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/cf_reader.h"
#include "io/defined_names.h"
#include "io/pivot_cache_reader.h"
#include "io/pivot_table_reader.h"
#include "io/sheet_reader.h"
#include "io/sst_reader.h"
#include "io/styles_reader.h"
#include "io/tables_reader.h"
#include "io/workbook_kind.h"
#include "io/zip_reader.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/structured_log.h"
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
constexpr std::string_view kRelTable = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
constexpr std::string_view kRelPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
constexpr std::string_view kRelPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
constexpr std::string_view kRelPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

// Content types we expect to see referenced from `[Content_Types].xml`.
// We only look up the workbook content type to verify the package is
// well-formed and to discriminate between `.xlsx` / `.xlsm` / `.xltx` /
// `.xltm` packages; the full content-type registry is built in a later
// bundle (which will recognise the worksheet, styles, sharedStrings,
// etc. content types as well).
constexpr std::string_view kCtWorkbookXlsx =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
constexpr std::string_view kCtWorkbookXlsm = "application/vnd.ms-excel.sheet.macroEnabled.main+xml";
constexpr std::string_view kCtWorkbookXltx =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml";
constexpr std::string_view kCtWorkbookXltm = "application/vnd.ms-excel.template.macroEnabled.main+xml";

/// Returns the matching `WorkbookKind` for `content_type`, or
/// `std::nullopt` when the string does not match any of the four
/// recognised workbook variants. Comparison is case-sensitive — Excel
/// emits the canonical lowercase form and we follow [OPC] §10's strict
/// matching rules.
struct DetectedKind {
  io::WorkbookKind kind;
  bool recognised;
};
DetectedKind DetectWorkbookKind(std::string_view content_type) {
  if (content_type == kCtWorkbookXlsx) {
    return {io::WorkbookKind::kXlsx, true};
  }
  if (content_type == kCtWorkbookXlsm) {
    return {io::WorkbookKind::kXlsm, true};
  }
  if (content_type == kCtWorkbookXltx) {
    return {io::WorkbookKind::kXltx, true};
  }
  if (content_type == kCtWorkbookXltm) {
    return {io::WorkbookKind::kXltm, true};
  }
  return {io::WorkbookKind::kXlsx, false};
}

/// Returns true if `content_type` references one of the four known
/// workbook variants. Used to gate "looks like a spreadsheet package"
/// without committing to the kind discriminator yet.
bool IsKnownWorkbookContentType(std::string_view content_type) {
  return content_type == kCtWorkbookXlsx || content_type == kCtWorkbookXlsm || content_type == kCtWorkbookXltx ||
         content_type == kCtWorkbookXltm;
}

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

/// Verifies that `[Content_Types].xml` references a workbook content
/// type at least once and returns the corresponding `WorkbookKind`.
///
/// The package is considered well-formed if any `<Override>` declares
/// either one of the four known workbook content types (xlsx / xlsm /
/// xltx / xltm) OR a content type that ends in
/// `spreadsheetml.sheet.main+xml` / similar — we are intentionally
/// strict here and only accept the four canonical strings. Anything
/// else yields a structured-log warning and a `kXlsx` fallback so
/// Excel-compatibility-first behaviour is preserved (the engine still
/// reads cells from non-canonical packages).
///
/// The full content-type registry is deferred to a later bundle.
Expected<io::WorkbookKind, Error> VerifyContentTypes(const std::vector<std::uint8_t>& ct_bytes) {
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
  // We need to recognise one of two situations:
  //   (a) An Override carries one of the four canonical workbook
  //       content types — we accept the package and pin the kind.
  //   (b) An Override exists for `/xl/workbook.xml` (or whatever the
  //       package-level rels file points at) but with an unfamiliar
  //       content type. We still accept the package for round-trip
  //       reads but log a warning and fall back to `kXlsx`.
  // The package-level rels file is parsed later in the pipeline; at
  // this point we walk every Override looking for a recognised kind
  // first, and only if we find none do we record the first
  // workbook-shaped override (heuristic: ContentType ends in
  // `+xml` and the part name is a candidate workbook). Keeping the
  // logic simple here matches the original "saw_workbook" heuristic
  // while extending it to four kinds.
  bool any_workbook_like = false;
  std::string first_unknown_ct;
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    if (std::string_view(node.name()) != "Override") {
      continue;
    }
    const std::string_view ct = node.attribute("ContentType").value();
    if (IsKnownWorkbookContentType(ct)) {
      return DetectWorkbookKind(ct).kind;
    }
    // Heuristic for case (b): part name targets `xl/workbook.xml`
    // (the canonical Excel placement) but with a content type we do
    // not recognise. Surface the first such occurrence so the warning
    // names the actual offender.
    if (first_unknown_ct.empty()) {
      const std::string_view part_name = node.attribute("PartName").value();
      if (part_name == "/xl/workbook.xml" || part_name == "xl/workbook.xml") {
        first_unknown_ct.assign(ct);
        any_workbook_like = true;
      }
    }
  }
  if (any_workbook_like) {
    StructuredLog("ooxml.reader.unknown_workbook_content_type")
        .field("content_type", first_unknown_ct)
        .field("fallback_kind", std::string_view("kXlsx"))
        .warn();
    return io::WorkbookKind::kXlsx;
  }
  return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: no workbook content-type override",
                    "context=ooxml_reader part=[Content_Types].xml");
}

/// One entry from `[Content_Types].xml`'s `<Override>` list, paired with
/// its declared content type. The reader uses the content type (a) to
/// decide whether the part is interesting at all (we only consume
/// recognised content types) and (b) so the writer slice can re-emit
/// the `<Override>` for passthrough parts verbatim.
struct OverrideEntry {
  std::string part_name;     // package-relative, no leading slash
  std::string content_type;  // verbatim ContentType= attribute value
};

/// Lists every part name advertised by `[Content_Types].xml`'s `<Override>`
/// elements together with its content type. `<Default>` entries are
/// ignored: they describe extensions rather than specific parts and the
/// passthrough flow only carries Override-registered parts.
std::vector<OverrideEntry> ListOverridePartEntries(const std::vector<std::uint8_t>& ct_bytes) {
  std::vector<OverrideEntry> out;
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
      if (part_name.empty()) {
        continue;
      }
      std::string content_type = node.attribute("ContentType").value();
      out.push_back(OverrideEntry{std::move(part_name), std::move(content_type)});
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
///
/// `pivot_cache_definition_paths_by_rid` carries the resolved part path
/// for every `<Relationship Type=".../pivotCacheDefinition">` entry,
/// keyed by relationship id. The workbook's `<pivotCaches>` element
/// (parsed by `read_ooxml`) joins each `cacheId` to its definition path
/// through this map.
struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;
  std::string sst_path;
  std::string styles_path;
  std::unordered_map<std::string, std::string> pivot_cache_definition_paths_by_rid;
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
    } else if (type == kRelPivotCacheDefinition) {
      const std::string id = rel.attribute("Id").value();
      if (id.empty()) {
        continue;
      }
      rels.pivot_cache_definition_paths_by_rid.emplace(id, ResolveRelativePath(base_dir, target));
    }
  }
  return rels;
}

/// Returns the sheet rels file path corresponding to `sheet_path`. For
/// example, `xl/worksheets/sheet1.xml` becomes
/// `xl/worksheets/_rels/sheet1.xml.rels`. Sheet parts at the package
/// root (no directory component) similarly produce `_rels/<file>.rels`.
std::string SheetRelsPath(std::string_view sheet_path) {
  const std::size_t slash = sheet_path.find_last_of('/');
  std::string out;
  if (slash == std::string_view::npos) {
    out.append("_rels/").append(sheet_path).append(".rels");
  } else {
    out.append(sheet_path.substr(0, slash));
    out.append("/_rels/");
    out.append(sheet_path.substr(slash + 1));
    out.append(".rels");
  }
  return out;
}

/// Loads `sheet_rels_path` and returns the resolved table-part paths it
/// references. The caller is expected to check `zip.has_entry(...)`
/// before invoking us; absent rels files are not an error (most sheets
/// have none) and are handled at the call site.
///
/// Each returned path is resolved relative to the sheet's directory so
/// `Target="../tables/table1.xml"` from `xl/worksheets/_rels/sheet1.xml.rels`
/// becomes `xl/tables/table1.xml`. Non-table relationships are silently
/// ignored at this layer; Bundle 2.5+ will widen the categorisation.
Expected<std::vector<std::string>, Error> LoadSheetTableTargets(const ZipReader& zip, std::string_view sheet_rels_path,
                                                                std::string_view sheet_dir) {
  std::vector<std::string> targets;
  auto rels_bytes_or = zip.read_entry(sheet_rels_path);
  if (!rels_bytes_or) {
    return rels_bytes_or.error();
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(sheet_rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "sheet rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(sheet_rels_path);
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "sheet rels: missing <Relationships>", std::move(ctx));
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    if (target.empty()) {
      continue;
    }
    if (type == kRelTable) {
      targets.push_back(ResolveRelativePath(sheet_dir, target));
    }
    // Other rel types (printerSettings, drawings, comments, ...) are
    // out of scope for Bundle 2.4. A future bundle may widen this
    // dispatch.
  }
  return targets;
}

/// Walks `sheet_rels_path` for `kRelPivotTable` entries and returns the
/// resolved part paths in document order. Mirrors `LoadSheetTableTargets`
/// in shape; the two helpers stay separate so each consumer site reads
/// linearly. Non-pivot relationships are silently ignored — the caller
/// decides which families it cares about.
Expected<std::vector<std::string>, Error> LoadSheetPivotTableTargets(const ZipReader& zip,
                                                                     std::string_view sheet_rels_path,
                                                                     std::string_view sheet_dir) {
  std::vector<std::string> targets;
  auto rels_bytes_or = zip.read_entry(sheet_rels_path);
  if (!rels_bytes_or) {
    return rels_bytes_or.error();
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(sheet_rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "sheet rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(sheet_rels_path);
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "sheet rels: missing <Relationships>", std::move(ctx));
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    if (target.empty()) {
      continue;
    }
    if (type == kRelPivotTable) {
      targets.push_back(ResolveRelativePath(sheet_dir, target));
    }
  }
  return targets;
}

/// Resolves the records-part target referenced from a
/// pivotCacheDefinition's own rels file (e.g.
/// `xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels`). Returns the
/// resolved path (relative to the package root) of the matching
/// `kRelPivotCacheRecords` entry, or an empty string when the rels file
/// is absent or carries no records relationship — both are valid OOXML
/// states (definition-only caches are uncommon but legal).
Expected<std::string, Error> LoadPivotCacheRecordsTarget(const ZipReader& zip, std::string_view definition_path) {
  // Build the rels path: <dir>/_rels/<filename>.rels.
  const std::size_t slash = definition_path.find_last_of('/');
  std::string rels_path;
  if (slash == std::string_view::npos) {
    rels_path.append("_rels/").append(definition_path).append(".rels");
  } else {
    rels_path.append(definition_path.substr(0, slash));
    rels_path.append("/_rels/");
    rels_path.append(definition_path.substr(slash + 1));
    rels_path.append(".rels");
  }
  if (!zip.has_entry(rels_path)) {
    return std::string{};
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
    return make_error(FormulonErrorCode::kIoXmlParse, "pivotCache rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=ooxml_reader part=");
    ctx.append(rels_path);
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "pivotCache rels: missing <Relationships>",
                      std::move(ctx));
  }
  const std::string base_dir = DirOf(definition_path);
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    if (target.empty()) {
      continue;
    }
    if (type == kRelPivotCacheRecords) {
      return ResolveRelativePath(base_dir, target);
    }
  }
  return std::string{};
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
  auto kind_or = VerifyContentTypes(ct_bytes);
  if (!kind_or) {
    return kind_or.error();
  }
  const io::WorkbookKind workbook_kind = kind_or.value();
  const std::vector<OverrideEntry> override_part_entries = ListOverridePartEntries(ct_bytes);

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
    }

    // Sheet rels file (`xl/worksheets/_rels/sheetN.xml.rels`) — drives
    // the table-part and pivot-table lookups. Optional: most sheets have
    // no rels at all. Tables and pivot tables are read in two passes
    // through the same rels file (separate helpers, each scoped to one
    // relationship type) so each consumer site reads linearly.
    const std::string sheet_rels_path = SheetRelsPath(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      const std::string sheet_dir = DirOf(sheet_path);
      auto targets_or = LoadSheetTableTargets(zip, sheet_rels_path, sheet_dir);
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
      auto pivot_targets_or = LoadSheetPivotTableTargets(zip, sheet_rels_path, sheet_dir);
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
        const std::string pt_rels_path = SheetRelsPath(pivot_table_path);
        if (zip.has_entry(pt_rels_path)) {
          consumed_parts.insert(pt_rels_path);
        }
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
    std::string rid = pc.attribute("r:id").value();
    if (rid.empty()) {
      rid = pc.attribute("id").value();
    }
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

    auto records_target_or = LoadPivotCacheRecordsTarget(zip, definition_path);
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
    const std::string def_rels_path = SheetRelsPath(definition_path);
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
  for (const OverrideEntry& entry : override_part_entries) {
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
  // Stable order so callers / tests can compare deterministically.
  std::sort(unknown_parts.begin(), unknown_parts.end(),
            [](const PassthroughPart& a, const PassthroughPart& b) { return a.path < b.path; });

  // Hand the same payload to the workbook so the writer can find it
  // even when the caller only retains `result.workbook`. The
  // `OoxmlReadResult::unknown_parts` view stays populated for tests and
  // tooling that want to inspect what was preserved.
  wb.set_passthrough_parts(unknown_parts);

  OoxmlReadResult result{std::move(wb), std::move(unknown_parts), pending_sst_count, std::move(result_text_storage)};
  return result;
}

}  // namespace io
}  // namespace formulon
