
#include "io/ooxml/package_validator.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/ooxml_defs.h"
#include "io/workbook_kind.h"
#include "io/xml_utils.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"

namespace formulon {
namespace io {
namespace ooxml {
namespace {

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
  WorkbookKind kind;
  bool recognised;
};
DetectedKind DetectWorkbookKind(std::string_view content_type) {
  if (content_type == kCtWorkbookXlsx) {
    return {WorkbookKind::kXlsx, true};
  }
  if (content_type == kCtWorkbookXlsm) {
    return {WorkbookKind::kXlsm, true};
  }
  if (content_type == kCtWorkbookXltx) {
    return {WorkbookKind::kXltx, true};
  }
  if (content_type == kCtWorkbookXltm) {
    return {WorkbookKind::kXltm, true};
  }
  return {WorkbookKind::kXlsx, false};
}

/// Returns true if `content_type` references one of the four known
/// workbook variants. Used to gate "looks like a spreadsheet package"
/// without committing to the kind discriminator yet.
bool IsKnownWorkbookContentType(std::string_view content_type) {
  return content_type == kCtWorkbookXlsx || content_type == kCtWorkbookXlsm || content_type == kCtWorkbookXltx ||
         content_type == kCtWorkbookXltm;
}

}  // namespace

Expected<std::string, Error> resolve_office_document_path(const std::vector<std::uint8_t>& rels_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, rels_bytes, "ooxml_reader", "package-level rels"));
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=ooxml_reader part=_rels/.rels");
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type = attr_str(rel, "Type");
    if (type == kRelOfficeDocument) {
      std::string target(attr_str(rel, "Target"));
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
      // The workbook part name becomes an archive key and the base
      // directory every downstream rels target resolves against, so it
      // goes through the same gate as every other package-rels target.
      if (!is_safe_part_name(target)) {
        return make_error(FormulonErrorCode::kIoZipSlip,
                          "package-level rels: OfficeDocument target escapes package root",
                          "context=ooxml_reader part=_rels/.rels target=" + target);
      }
      return target;
    }
  }
  return make_error(FormulonErrorCode::kIoRelationshipBroken,
                    "package-level rels: no OfficeDocument relationship found",
                    "context=ooxml_reader part=_rels/.rels");
}

Expected<WorkbookKind, Error> verify_content_types(const std::vector<std::uint8_t>& ct_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, ct_bytes, "ooxml_reader", "[Content_Types].xml"));
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
    const std::string_view ct = attr_str(node, "ContentType");
    if (IsKnownWorkbookContentType(ct)) {
      return DetectWorkbookKind(ct).kind;
    }
    // Heuristic for case (b): part name targets `xl/workbook.xml`
    // (the canonical Excel placement) but with a content type we do
    // not recognise. Surface the first such occurrence so the warning
    // names the actual offender.
    if (first_unknown_ct.empty()) {
      const std::string_view part_name = attr_str(node, "PartName");
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
    return WorkbookKind::kXlsx;
  }
  return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: no workbook content-type override",
                    "context=ooxml_reader part=[Content_Types].xml");
}

Expected<std::vector<OverrideEntry>, Error> list_override_part_entries(const std::vector<std::uint8_t>& ct_bytes) {
  std::vector<OverrideEntry> out;
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, ct_bytes, "ooxml_reader", "[Content_Types].xml"));
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=ooxml_reader part=[Content_Types].xml");
  }
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    if (std::string_view(node.name()) == "Override") {
      std::string part_name(attr_str(node, "PartName"));
      if (!part_name.empty() && part_name.front() == '/') {
        part_name.erase(0, 1);
      }
      if (part_name.empty()) {
        continue;
      }
      std::string content_type(attr_str(node, "ContentType"));
      out.push_back(OverrideEntry{std::move(part_name), std::move(content_type)});
    }
  }
  return out;
}

Expected<std::vector<DefaultContentType>, Error> list_default_content_types(const std::vector<std::uint8_t>& ct_bytes) {
  std::vector<DefaultContentType> out;
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, ct_bytes, "ooxml_reader", "[Content_Types].xml"));
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=ooxml_reader part=[Content_Types].xml");
  }
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    if (std::string_view(node.name()) != "Default") {
      continue;
    }
    std::string extension(attr_str(node, "Extension"));
    if (extension.empty()) {
      continue;
    }
    // Normalise to lower case so lookups against archive entry suffixes
    // stay case-insensitive (OPC treats extensions case-insensitively).
    for (char& c : extension) {
      if (c >= 'A' && c <= 'Z') {
        c = static_cast<char>(c - 'A' + 'a');
      }
    }
    out.push_back(DefaultContentType{std::move(extension), std::string(attr_str(node, "ContentType"))});
  }
  return out;
}

Expected<std::string, Error> resolve_relative_path(std::string_view base_dir, std::string_view target) {
  // Absolute-path targets are not permitted in well-formed OOXML rels
  // files. A writer that emits `Target="/xl/..."` is either misbehaving
  // or attempting path traversal; in either case the result is not safe
  // to consume because the path no longer participates in `..`-segment
  // accounting that defends the package root.
  if (!target.empty() && target.front() == '/') {
    std::string ctx("context=ooxml_reader base_dir=");
    ctx.append(base_dir);
    ctx.append(" target=");
    ctx.append(target);
    return make_error(FormulonErrorCode::kIoZipSlip, "rels target uses package-absolute path; refusing to resolve",
                      std::move(ctx));
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
  // Append the target, applying `.` / `..` normalisation. Track whether
  // any `..` segment failed to find a directory to pop so we can refuse
  // a target that escapes the package root.
  start = 0;
  for (std::size_t i = 0; i <= target.size(); ++i) {
    if (i == target.size() || target[i] == '/') {
      if (i > start) {
        std::string_view seg = target.substr(start, i - start);
        if (seg == ".") {
          // skip
        } else if (seg == "..") {
          if (stack.empty()) {
            std::string ctx("context=ooxml_reader base_dir=");
            ctx.append(base_dir);
            ctx.append(" target=");
            ctx.append(target);
            return make_error(FormulonErrorCode::kIoZipSlip, "rels target escapes package root via '..' traversal",
                              std::move(ctx));
          }
          stack.pop_back();
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
  if (out.empty()) {
    std::string ctx("context=ooxml_reader base_dir=");
    ctx.append(base_dir);
    ctx.append(" target=");
    ctx.append(target);
    return make_error(FormulonErrorCode::kIoZipSlip, "rels target resolves to empty path", std::move(ctx));
  }
  return out;
}

bool is_safe_part_name(std::string_view part_name) noexcept {
  // Passthrough part names are stored package-relative with the leading
  // slash already stripped, so a residual leading '/' means the raw name
  // was package-absolute (or, worse, filesystem-absolute).
  if (part_name.empty() || part_name.front() == '/') {
    return false;
  }
  // Reject any character that lets a name masquerade as a different path
  // shape once written to disk: backslash separators and drive/scheme
  // colons. Control bytes and NUL are likewise never valid part names.
  for (const char c : part_name) {
    const auto uc = static_cast<unsigned char>(c);
    if (c == '\\' || c == ':' || uc < 0x20) {
      return false;
    }
  }
  // Walk '/'-separated segments and refuse `.`, `..`, and empty segments
  // (the latter come from `//`, a leading slash we already caught, or a
  // trailing slash — none of which name a real part).
  std::size_t start = 0;
  for (std::size_t i = 0; i <= part_name.size(); ++i) {
    if (i == part_name.size() || part_name[i] == '/') {
      const std::string_view seg = part_name.substr(start, i - start);
      if (seg.empty() || seg == "." || seg == "..") {
        return false;
      }
      start = i + 1;
    }
  }
  return true;
}

std::string dir_of(std::string_view path) {
  const std::size_t pos = path.find_last_of('/');
  if (pos == std::string_view::npos) {
    return {};
  }
  return std::string(path.substr(0, pos));
}

std::string rels_path_for_part(std::string_view part_path) {
  const std::size_t slash = part_path.find_last_of('/');
  std::string out;
  if (slash == std::string_view::npos) {
    out.append("_rels/").append(part_path).append(".rels");
  } else {
    out.append(part_path.substr(0, slash));
    out.append("/_rels/");
    out.append(part_path.substr(slash + 1));
    out.append(".rels");
  }
  return out;
}

std::string relationship_ref_id(const pugi::xml_node& node) {
  std::string rid = node.attribute("r:id").value();
  if (rid.empty()) {
    rid = node.attribute("id").value();
  }
  return rid;
}

std::string lowercase_extension(std::string_view extension) {
  std::string out(extension);
  for (char& c : out) {
    if (c >= 'A' && c <= 'Z') {
      c = static_cast<char>(c - 'A' + 'a');
    }
  }
  return out;
}

std::string extension_of_part(std::string_view path) {
  const std::size_t slash = path.find_last_of('/');
  const std::size_t dot = path.find_last_of('.');
  if (dot == std::string_view::npos || dot == path.size() - 1U || (slash != std::string_view::npos && dot < slash)) {
    return {};
  }
  return lowercase_extension(path.substr(dot + 1U));
}

}  // namespace ooxml
}  // namespace io
}  // namespace formulon
