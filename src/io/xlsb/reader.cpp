//
// Implementation of the XLSB read pipeline. See `io/xlsb/reader.h` for
// the design context.
//
// The reader reuses the OOXML zip envelope plus three XML parts
// (`[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`).
// The binary parts (`xl/workbook.bin`, `xl/worksheets/sheet*.bin`,
// `xl/sharedStrings.bin`) are decoded via the record framing in
// `io/xlsb/record.h`. Formulas are decoded through a full
// `Ptg → AST → Excel-formula-text` pipeline (`DecodeFormulaText`
// below); a Ptg stream outside the supported token set logs a
// structured warning and leaves the formula text empty instead.

#include "io/xlsb/reader.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "io/default_content_type.h"
#include "io/ooxml/package_validator.h"
#include "io/passthrough_part.h"
#include "io/unknown_relationship.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/styles_reader.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/strings.h"
#include "utils/structured_log.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

// ---------------------------------------------------------------------------
// XML envelope: same shape as the OOXML reader, but the relationship and
// content types target binary parts (`*.bin`) instead of XML.
// ---------------------------------------------------------------------------

constexpr std::string_view kRelOfficeDocument =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
constexpr std::string_view kRelCoreProperties =
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
constexpr std::string_view kRelExtendedProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
constexpr std::string_view kRelCustomProperties =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties";
constexpr std::string_view kRelWorksheet =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
constexpr std::string_view kRelSharedStrings =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
constexpr std::string_view kRelStyles = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";

// `[Content_Types].xml` registers the workbook part under the binary
// content type for `.xlsb`. The reader gates on this content type to
// avoid acting on a `.xlsx` archive that happened to be passed in.
constexpr std::string_view kCtWorkbookXlsb = "application/vnd.ms-excel.sheet.binary.macroEnabled.main";
// Some non-macro xlsb writers emit the alternative content type below;
// we accept either flavor as "this is an xlsb workbook".
constexpr std::string_view kCtWorkbookXlsbAlt = "application/vnd.ms-excel.sheet.macroEnabled.main";

/// Drops the leading `/` from a part path (relationship targets can be
/// absolute or relative, but the ZIP catalogue uses relative names).
std::string NormalisePartName(std::string_view raw) {
  if (!raw.empty() && raw.front() == '/') {
    return std::string(raw.substr(1));
  }
  return std::string(raw);
}

/// Resolves an OOXML rels target relative to a base directory, normalising
/// `.` / `..` segments. Thin wrapper over the OOXML package validator's
/// shared helper so both readers share a single path-traversal defence.
///
/// Path-traversal hardening: a `Target` with excess `..` segments or a
/// leading `/` (package-absolute) is refused with `kIoZipSlip`. The
/// surfaced error context is the validator's standard `context=ooxml_reader`
/// flavour; xlsb's xml envelope is structurally identical to xlsx, so the
/// shared diagnostic stays accurate for both.
Expected<std::string, Error> ResolveRelativePath(std::string_view base_dir, std::string_view target) {
  return ooxml::resolve_relative_path(base_dir, target);
}

std::string DirOf(std::string_view path) {
  const std::size_t pos = path.find_last_of('/');
  if (pos == std::string_view::npos) {
    return {};
  }
  return std::string(path.substr(0, pos));
}

std::string RelsPathForPart(std::string_view part_path) {
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

template <typename Fn>
Expected<void, Error> VisitRelationshipNodes(const ZipReader& zip, std::string_view rels_path, std::string_view label,
                                             Fn&& fn) {
  auto rels_bytes_or = zip.read_entry(rels_path);
  if (!rels_bytes_or) {
    return rels_bytes_or.error();
  }
  const std::vector<std::uint8_t>& rels_bytes = rels_bytes_or.value();

  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=xlsb_reader part=");
    ctx.append(rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    std::string message(label);
    message.append(": pugixml parse failed");
    return make_error(FormulonErrorCode::kIoXmlParse, std::move(message), std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    std::string ctx("context=xlsb_reader part=");
    ctx.append(rels_path);
    std::string message(label);
    message.append(": missing <Relationships>");
    return make_error(FormulonErrorCode::kIoRelationshipBroken, std::move(message), std::move(ctx));
  }

  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    auto status = fn(rel);
    if (!status) {
      return status.error();
    }
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// XML envelope walkers (return `Expected` on every failure path).
// ---------------------------------------------------------------------------

/// Walks `[Content_Types].xml` and asserts at least one Override is
/// for the xlsb workbook content type. Also returns the full Override
/// list so callers can compute passthrough parts.
struct ContentTypesView {
  std::vector<std::pair<std::string, std::string>> overrides;  // (part_name, content_type)
  std::vector<DefaultContentType> defaults;
};

std::string LowercaseExtension(std::string_view extension) {
  std::string out(extension);
  for (char& c : out) {
    if (c >= 'A' && c <= 'Z') {
      c = static_cast<char>(c - 'A' + 'a');
    }
  }
  return out;
}

std::string ExtensionOfPart(std::string_view path) {
  const std::size_t slash = path.find_last_of('/');
  const std::size_t dot = path.find_last_of('.');
  if (dot == std::string_view::npos || dot == path.size() - 1U || (slash != std::string_view::npos && dot < slash)) {
    return {};
  }
  return LowercaseExtension(path.substr(dot + 1U));
}

Expected<ContentTypesView, Error> LoadContentTypes(const std::vector<std::uint8_t>& ct_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(ct_bytes.data(), ct_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=xlsb_reader part=[Content_Types].xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "[Content_Types].xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Types");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing <Types> root",
                      "context=xlsb_reader part=[Content_Types].xml");
  }
  ContentTypesView view;
  bool saw_workbook = false;
  for (pugi::xml_node node = root.first_child(); node; node = node.next_sibling()) {
    const std::string_view node_name(node.name());
    if (node_name == "Default") {
      // Real Excel-365 output does not always carry a per-part
      // `<Override>` for `xl/workbook.bin`: the workbook part can rely
      // on the blanket `<Default Extension="bin">` registration alone
      // (verified against a real Excel-365-produced `.xlsb`). Accept
      // that form too, rather than requiring an Override that a
      // genuinely valid package may not emit.
      const std::string_view ext(node.attribute("Extension").value());
      const std::string_view ct(node.attribute("ContentType").value());
      if (strings::case_insensitive_eq(ext, "bin") && (ct == kCtWorkbookXlsb || ct == kCtWorkbookXlsbAlt)) {
        saw_workbook = true;
      }
      if (ext.empty()) {
        continue;
      }
      const std::string normalized_ext = LowercaseExtension(ext);
      auto existing = std::find_if(
          view.defaults.begin(), view.defaults.end(),
          [&normalized_ext](const DefaultContentType& value) { return value.extension == normalized_ext; });
      if (existing != view.defaults.end()) {
        if (existing->content_type != ct) {
          return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                            "[Content_Types].xml: conflicting Default content types for extension " + normalized_ext,
                            "context=xlsb_reader part=[Content_Types].xml extension=" + normalized_ext);
        }
        continue;
      }
      view.defaults.push_back(DefaultContentType{normalized_ext, std::string(ct)});
      continue;
    }
    if (node_name != "Override") {
      continue;
    }
    std::string part_name = NormalisePartName(node.attribute("PartName").value());
    if (part_name.empty()) {
      continue;
    }
    std::string ct = node.attribute("ContentType").value();
    if (ct == kCtWorkbookXlsb || ct == kCtWorkbookXlsbAlt) {
      saw_workbook = true;
    }
    auto existing = std::find_if(
        view.overrides.begin(), view.overrides.end(),
        [&part_name](const std::pair<std::string, std::string>& value) { return value.first == part_name; });
    if (existing != view.overrides.end()) {
      if (existing->second != ct) {
        return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                          "[Content_Types].xml: conflicting Override content types for part " + part_name,
                          "context=xlsb_reader part=[Content_Types].xml override=" + part_name);
      }
      continue;
    }
    view.overrides.emplace_back(std::move(part_name), std::move(ct));
  }
  if (!saw_workbook) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "[Content_Types].xml: no xlsb-workbook content-type override",
                      "context=xlsb_reader part=[Content_Types].xml");
  }
  return view;
}

Expected<std::string, Error> ResolveOfficeDocumentPath(const std::vector<std::uint8_t>& rels_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(rels_bytes.data(), rels_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=xlsb_reader part=_rels/.rels desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "package-level rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=xlsb_reader part=_rels/.rels");
  }
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    if (std::string_view(rel.attribute("Type").value()) == kRelOfficeDocument) {
      std::string target = NormalisePartName(rel.attribute("Target").value());
      if (target.empty()) {
        return make_error(FormulonErrorCode::kIoRelationshipBroken,
                          "package-level rels: empty Target for OfficeDocument",
                          "context=xlsb_reader part=_rels/.rels");
      }
      return target;
    }
  }
  return make_error(FormulonErrorCode::kIoRelationshipBroken,
                    "package-level rels: no OfficeDocument relationship found", "context=xlsb_reader part=_rels/.rels");
}

Expected<std::vector<UnknownRelationship>, Error> ReadUnknownPackageRels(const std::vector<std::uint8_t>& bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse = doc.load_buffer(bytes.data(), bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    return make_error(FormulonErrorCode::kIoXmlParse, "package-level rels: pugixml parse failed",
                      "context=xlsb_reader part=_rels/.rels desc=" + std::string(parse.description()));
  }
  const pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "package-level rels: missing <Relationships>",
                      "context=xlsb_reader part=_rels/.rels");
  }
  std::vector<UnknownRelationship> result;
  for (pugi::xml_node rel = root.child("Relationship"); rel; rel = rel.next_sibling("Relationship")) {
    const std::string_view type(rel.attribute("Type").value());
    if (type == kRelOfficeDocument || type == kRelCoreProperties || type == kRelExtendedProperties ||
        type == kRelCustomProperties) {
      continue;
    }
    const bool external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    std::string target = NormalisePartName(rel.attribute("Target").value());
    if (type.empty() || target.empty())
      continue;
    if (!external && !ooxml::is_safe_part_name(target)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "package relationship target escapes package root",
                        "context=xlsb_reader part=_rels/.rels target=" + target);
    }
    result.push_back(
        UnknownRelationship{std::string(rel.attribute("Id").value()), std::string(type), std::move(target), external});
  }
  return result;
}

struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;  // rId -> resolved part path
  std::string sst_path;
  std::string styles_path;
  std::vector<UnknownRelationship> unknown_rels;
};

Expected<WorkbookRels, Error> LoadWorkbookRels(const ZipReader& zip, std::string_view workbook_path) {
  WorkbookRels rels;
  const std::string rels_path = RelsPathForPart(workbook_path);
  if (!zip.has_entry(rels_path)) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=xlsb_reader rels_path=" + rels_path);
  }
  const std::string base_dir = DirOf(workbook_path);
  auto visit_status =
      VisitRelationshipNodes(zip, rels_path, "workbook rels", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        const std::string_view type = rel.attribute("Type").value();
        const std::string_view target = rel.attribute("Target").value();
        if (target.empty()) {
          return Expected<void, Error>::Ok();
        }
        if (type == kRelWorksheet) {
          const std::string id = rel.attribute("Id").value();
          if (id.empty()) {
            return Expected<void, Error>::Ok();
          }
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.sheet_targets.emplace(id, std::move(resolved).value());
        } else if (type == kRelSharedStrings) {
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.sst_path = std::move(resolved).value();
        } else if (type == kRelStyles) {
          auto resolved = ResolveRelativePath(base_dir, target);
          if (!resolved) {
            return resolved.error();
          }
          rels.styles_path = std::move(resolved).value();
        } else {
          UnknownRelationship unknown;
          unknown.id.assign(rel.attribute("Id").value());
          unknown.type.assign(type);
          unknown.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
          if (unknown.target_external) {
            unknown.target.assign(target);
          } else {
            auto resolved = ResolveRelativePath(base_dir, target);
            if (!resolved) {
              return resolved.error();
            }
            unknown.target = std::move(resolved).value();
          }
          if (!unknown.type.empty()) {
            rels.unknown_rels.push_back(std::move(unknown));
          }
        }
        return Expected<void, Error>::Ok();
      });
  if (!visit_status) {
    return visit_status.error();
  }
  return rels;
}

// ---------------------------------------------------------------------------
// Binary part decoders.
// ---------------------------------------------------------------------------

/// One decoded entry of the workbook's sheet bundle.
///
///   * `name`  — display name of the sheet.
///   * `rid`   — workbook-rels relationship id pointing at the sheet
///               binary part.
struct SheetBundleEntry {
  std::string name;
  std::string rid;
  /// True when the sheet's `hsState` is hidden or very-hidden. The Sheet
  /// model carries a single `tab_hidden` bit (mirroring the OOXML path),
  /// so both XLSB visibility states fold to "hidden" here.
  bool hidden = false;
};

/// Workbook-global fields needed while constructing the model.
struct WorkbookBinInfo {
  std::vector<SheetBundleEntry> sheets;
  bool date1904 = false;
};

/// Decodes `xl/workbook.bin` to extract the ordered sheet-bundle list and
/// workbook date system. Other records are skipped.
Expected<WorkbookBinInfo, Error> DecodeWorkbookBin(const std::vector<std::uint8_t>& body) {
  WorkbookBinInfo info;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp)) {
      // BrtWbProp ([MS-XLSB] §2.4.866) begins with a u32 grbit; bit 0
      // is f1904. The following theme-version and optional code-name
      // fields are irrelevant to the workbook model.
      ByteSpan p = rec.payload;
      auto flags_or = read_u32(p);
      if (!flags_or) {
        return flags_or.error();
      }
      info.date1904 = (flags_or.value() & 0x00000001U) != 0U;
      continue;
    }
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtBundleSh)) {
      continue;
    }
    // BrtBundleSh ([MS-XLSB] §2.4.304):
    //   hsState    : u32 (visibility)
    //   iTabID     : u32
    //   strRelID   : XLNullableWideString
    //   strName    : XLWideString
    ByteSpan p = rec.payload;
    // hsState: 0 = visible, 1 = hidden, 2 = very hidden.
    auto hs_state_or = read_u32(p);
    if (!hs_state_or) {
      return hs_state_or.error();
    }
    auto skip2 = read_u32(p);  // iTabID
    if (!skip2) {
      return skip2.error();
    }
    auto rid_or = read_xlnullablewidestring(p);
    if (!rid_or) {
      return rid_or.error();
    }
    auto name_or = read_xlwidestring(p);
    if (!name_or) {
      return name_or.error();
    }
    SheetBundleEntry entry;
    entry.rid = std::move(rid_or.value());
    entry.name = std::move(name_or.value());
    entry.hidden = hs_state_or.value() != 0U;
    info.sheets.push_back(std::move(entry));
  }
  if (info.sheets.empty()) {
    return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: no BrtBundleSh records",
                      "context=xlsb_reader part=xl/workbook.bin");
  }
  return info;
}

/// Decodes a UTF-16LE name of `units` code units starting at `cursor`,
/// advancing it past the name. `units` is caller-known (from a fixed-size
/// header field), unlike `read_xlwidestring`'s self-describing length. This
/// path intentionally avoids Expected<std::string, Error>: malformed
/// workbook names are common fuzz inputs, and libc++'s variant dispatch
/// under UBSan must not turn recoverable input errors into a process abort.
bool ReadFixedWideString(ByteSpan& cursor, std::uint32_t units, std::string& out) {
  // Check in the destination width before multiplying. On wasm32, a hostile
  // u32 `units` value can wrap `units * 2` back below cursor.size otherwise.
  if (units > cursor.size / 2U) {
    return false;
  }
  const std::size_t byte_len = static_cast<std::size_t>(units) * 2U;
  out.clear();
  out.reserve(byte_len);
  for (std::uint32_t i = 0; i < units; ++i) {
    const std::size_t offset = static_cast<std::size_t>(i) * 2U;
    const std::uint16_t cu = static_cast<std::uint16_t>(static_cast<std::uint16_t>(cursor.data[offset]) |
                                                        (static_cast<std::uint16_t>(cursor.data[offset + 1U]) << 8));
    std::uint32_t cp = cu;
    if (cu >= 0xD800U && cu <= 0xDBFFU && i + 1 < units) {
      const std::size_t low_offset = static_cast<std::size_t>(i + 1U) * 2U;
      const std::uint16_t low =
          static_cast<std::uint16_t>(static_cast<std::uint16_t>(cursor.data[low_offset]) |
                                     (static_cast<std::uint16_t>(cursor.data[low_offset + 1U]) << 8));
      if (low >= 0xDC00U && low <= 0xDFFFU) {
        cp = 0x10000U + ((static_cast<std::uint32_t>(cu) - 0xD800U) << 10) + (low - 0xDC00U);
        ++i;
      }
    }
    if (cp < 0x80U) {
      out.push_back(static_cast<char>(cp));
    } else if (cp < 0x800U) {
      out.push_back(static_cast<char>(0xC0U | (cp >> 6)));
      out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    } else if (cp < 0x10000U) {
      out.push_back(static_cast<char>(0xE0U | (cp >> 12)));
      out.push_back(static_cast<char>(0x80U | ((cp >> 6) & 0x3FU)));
      out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    } else {
      out.push_back(static_cast<char>(0xF0U | (cp >> 18)));
      out.push_back(static_cast<char>(0x80U | ((cp >> 12) & 0x3FU)));
      out.push_back(static_cast<char>(0x80U | ((cp >> 6) & 0x3FU)));
      out.push_back(static_cast<char>(0x80U | (cp & 0x3FU)));
    }
  }
  cursor.data += byte_len;
  cursor.size -= byte_len;
  return true;
}

/// Decodes the workbook-scope `BrtName` table from `xl/workbook.bin`:
/// ordinary defined names and the hidden `_xlfn.*` / `_xlpm.*`
/// future-function / LET-parameter placeholders `PtgName` resolves by
/// 1-based declaration order. Byte layout verified against a real
/// Excel-365-produced `xl/workbook.bin`:
///   flags (u16) + 3 reserved bytes + itab (i32) + cch (u32) +
///   cch x UTF-16LE code units + <formula body, not consumed here>.
/// Records other than `BrtName` are skipped.
Expected<std::vector<XlsbName>, Error> DecodeWorkbookNames(const std::vector<std::uint8_t>& body) {
  std::vector<XlsbName> names;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtName)) {
      continue;
    }
    ByteSpan p = rec.payload;
    if (p.size < 5) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "workbook.bin: BrtName header truncated",
                        "context=xlsb_reader");
    }
    auto flags_or = read_u16(p);
    if (!flags_or) {
      return flags_or.error();
    }
    p.data += 3;  // 3 reserved bytes between flags and itab.
    p.size -= 3;
    auto itab_or = read_u32(p);
    if (!itab_or) {
      return itab_or.error();
    }
    auto cch_or = read_u32(p);
    if (!cch_or) {
      return cch_or.error();
    }
    std::string name;
    if (!ReadFixedWideString(p, cch_or.value(), name)) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb fixed-length wide string truncated",
                        "context=xlsb_reader");
    }
    XlsbName entry;
    entry.itab = static_cast<std::int32_t>(itab_or.value());
    entry.name = std::move(name);
    entry.hidden = (flags_or.value() & 0x0001U) != 0;
    names.push_back(std::move(entry));
  }
  return names;
}

/// Decodes the `BrtExternSheet` table from `xl/workbook.bin`: resolves a
/// `PtgRef3d` / `PtgArea3d` `ixti` (0-based index into the returned
/// vector) to a `(itabFirst, itabLast)` sheet-index range. Byte layout
/// verified against a real Excel-365-produced `xl/workbook.bin`: u32
/// count, followed by `count` entries of `(iSupBook, itabFirst,
/// itabLast)` as 3 x i32 each. `iSupBook` is not consulted — external-
/// workbook 3-D references (a non-zero `iSupBook`) are out of scope;
/// only the internal (`iSupBook == 0`, "this workbook") case is
/// resolved. A workbook with no qualified references at all carries no
/// `BrtExternSheet` record, so an empty result is a normal outcome, not
/// an error.
Expected<std::vector<XlsbSheetRange>, Error> DecodeExternSheet(const std::vector<std::uint8_t>& body) {
  std::vector<XlsbSheetRange> ranges;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      continue;
    }
    ByteSpan p = rec.payload;
    auto count_or = read_u32(p);
    if (!count_or) {
      return count_or.error();
    }
    // Each entry consumes three u32s. Never reserve beyond what the
    // remaining payload can actually hold, so an attacker-controlled
    // count (up to 4 billion) cannot force a multi-GB reservation before
    // the per-entry reads run out of bytes and fail.
    constexpr std::size_t kExternSheetEntryBytes = 12U;
    const std::size_t reservable =
        static_cast<std::size_t>(std::min<std::uint64_t>(count_or.value(), p.size / kExternSheetEntryBytes));
    ranges.reserve(reservable);
    for (std::uint32_t i = 0; i < count_or.value(); ++i) {
      auto sup_book_or = read_u32(p);
      if (!sup_book_or) {
        return sup_book_or.error();
      }
      auto first_or = read_u32(p);
      if (!first_or) {
        return first_or.error();
      }
      auto last_or = read_u32(p);
      if (!last_or) {
        return last_or.error();
      }
      (void)sup_book_or;
      ranges.push_back(
          XlsbSheetRange{static_cast<std::int32_t>(first_or.value()), static_cast<std::int32_t>(last_or.value())});
    }
    break;  // Exactly one BrtExternSheet record per workbook.
  }
  return ranges;
}

/// Decodes `xl/sharedStrings.bin` into an in-order list of string
/// payloads, appending each one into `text_storage` so cells can take
/// non-owning views. Returns the list of `string_view`s parallel to the
/// SST index.
Expected<std::vector<std::string_view>, Error> DecodeSharedStringsBin(const std::vector<std::uint8_t>& body,
                                                                      std::deque<std::string>& text_storage) {
  std::vector<std::string_view> entries;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtSSTItem)) {
      continue;
    }
    // BrtSSTItem ([MS-XLSB] §2.4.293):
    //   richStr   : RichStr (we only need the string — first byte is a
    //               flags byte; rich-format runs follow when fRichStr is
    //               set, but those are not decoded here).
    ByteSpan p = rec.payload;
    auto flags_or = read_u8(p);
    if (!flags_or) {
      return flags_or.error();
    }
    (void)flags_or.value();
    auto str_or = read_xlwidestring(p);
    if (!str_or) {
      return str_or.error();
    }
    text_storage.push_back(std::move(str_or.value()));
    entries.push_back(text_storage.back());
  }
  return entries;
}

/// Attempts to decode the `rgce` Ptg byte stream into an Excel formula
/// text (with a leading `=`). Returns the formula text on success, or an
/// empty string when the stream uses a token outside the supported set
/// (a structured-log diagnostic records the reason). On the empty-string
/// path the caller PRESERVES the cell's cached value and stores no
/// formula, so the cell still displays the correct cached result instead
/// of a fabricated formula that would recalc to `#NAME?`.
std::string DecodeFormulaText(ByteSpan ptg_bytes, ByteSpan rgcb, const std::vector<std::string>& sheet_names,
                              const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                              std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                              std::uint32_t* undecoded_formula_count) {
  Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
  auto ast_or = decode_ptgs(ptg_bytes, rgcb, arena, sheet_names, name_table, sheet_ranges);
  if (!ast_or) {
    StructuredLog("xlsb.formula.not_decoded")
        .field("sheet_index", static_cast<std::int64_t>(sheet_index))
        .field("row", static_cast<std::int64_t>(row))
        .field("col", static_cast<std::int64_t>(col))
        .field("ptg_bytes", static_cast<std::int64_t>(ptg_bytes.size))
        .field("reason", ast_or.error().message)
        .warn();
    if (undecoded_formula_count != nullptr) {
      ++*undecoded_formula_count;
    }
    return {};
  }
  std::string out("=");
  out.append(parser::format_formula(*ast_or.value()));
  return out;
}

/// Registers every non-hidden `BrtName` entry as a workbook defined
/// name (`Workbook::set_defined_name_scoped`). Hidden entries (the
/// `_xlfn.*` / `_xlpm.*` future-function and LET/LAMBDA-parameter
/// placeholders `PtgName` resolves during Ptg decode — see
/// `name_table`) are skipped; they are not user-visible names. Walks
/// `xl/workbook.bin`'s `BrtName` records a second time (after
/// `DecodeWorkbookNames` has already built the complete `name_table`),
/// decoding each entry's own formula body so a qualified or
/// name-referencing formula (e.g. `Rate` defined as a cell reference)
/// resolves against the full table rather than a partially-built one.
/// A name whose formula uses a Ptg token outside the supported set logs
/// a structured warning and is skipped rather than failing the whole
/// read — the OOXML reader's `DecodeFormulaText` contract mirrored at
/// the workbook-name level.
Expected<void, Error> RegisterDefinedNames(const std::vector<std::uint8_t>& body,
                                           const std::vector<XlsbName>& name_table,
                                           const std::vector<std::string>& sheet_names,
                                           const std::vector<XlsbSheetRange>& sheet_ranges, Workbook& wb,
                                           std::uint32_t* undecoded_defined_name_count) {
  ByteSpan cursor{body.data(), body.size()};
  std::size_t name_index = 0;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtName)) {
      continue;
    }
    if (name_index >= name_table.size()) {
      break;  // Defensive: should be unreachable (same records, same order).
    }
    const XlsbName& entry = name_table[name_index];
    ++name_index;
    if (entry.hidden) {
      continue;
    }
    ByteSpan p = rec.payload;
    if (p.size < 9) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        "workbook.bin: BrtName header truncated (defined-name pass)", "context=xlsb_reader");
    }
    p.data += 9;  // flags (2) + 3 reserved bytes + itab (4).
    p.size -= 9;
    auto cch_or = read_u32(p);
    if (!cch_or) {
      return cch_or.error();
    }
    const std::size_t name_bytes = static_cast<std::size_t>(cch_or.value()) * 2;
    if (name_bytes > p.size) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        "workbook.bin: BrtName name truncated (defined-name pass)", "context=xlsb_reader");
    }
    p.data += name_bytes;
    p.size -= name_bytes;
    auto cce_or = read_u32(p);
    if (!cce_or) {
      return cce_or.error();
    }
    const std::uint32_t cce = cce_or.value();
    if (cce > p.size) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        "workbook.bin: BrtName formula rgce length exceeds payload", "context=xlsb_reader");
    }
    if (cce == 0) {
      continue;
    }
    ByteSpan rgce{p.data, cce};
    Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
    auto ast_or = decode_ptgs(rgce, ByteSpan{}, arena, sheet_names, name_table, sheet_ranges);
    if (!ast_or) {
      StructuredLog("xlsb.defined_name.not_decoded")
          .field("name", entry.name)
          .field("reason", ast_or.error().message)
          .warn();
      if (undecoded_defined_name_count != nullptr) {
        ++*undecoded_defined_name_count;
      }
      continue;
    }
    // Defined-name formulas store the bare expression text (no leading
    // `=`), matching the OOXML `<definedName>` element's text content.
    const std::string formula_text = parser::format_formula(*ast_or.value());
    auto reg = wb.set_defined_name_scoped(entry.name, formula_text, entry.itab);
    if (!reg) {
      return reg.error();
    }
  }
  return Expected<void, Error>::Ok();
}

/// A dynamic-array formula's anchor cell plus its footprint, decoded
/// from a `BrtArrFmla` record's `RfX`. `row` / `col` is the anchor
/// (`rwFirst`, `colFirst`); `last_row` / `last_col` is the footprint's
/// bottom-right corner (`rwLast`, `colLast`). One-cell array anchors are
/// retained as well so dynamic-array metadata survives an XLSB rewrite.
struct ArrayAnchor {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
};

/// Per-sheet decode state. The reader walks records in order and
/// updates `current_row` whenever it sees a `BrtRowHdr`. Cell records
/// then resolve `(current_row, col)` to the absolute cell address.
struct SheetDecodeState {
  std::uint32_t current_row = 0;
  bool row_seen = false;
  std::uint32_t cells_decoded = 0;
  /// Dynamic-array anchors recorded while walking `BrtArrFmla` records.
  /// Registered as spill regions in a second pass after the whole sheet
  /// has been decoded (see `RegisterArraySpills`'s doc comment for why
  /// this cannot happen inline).
  std::vector<ArrayAnchor> array_anchors;
  /// Worksheet-tail records retained verbatim (see `XlsbSheetTail`).
  XlsbSheetTail tail;
  /// True once `BrtEndSheetData` has been seen: everything from there to
  /// `BrtEndSheet` is tail.
  bool in_tail = false;
  /// True once the merged-cell block has been passed, which selects which
  /// of `tail`'s two buffers subsequent records append to.
  bool merges_seen = false;
};

/// True for a tail record the writer re-emits from the model, which must
/// therefore not also be retained verbatim (or it would be emitted twice).
bool IsModelOwnedTailRecord(XlsbRecordType type) {
  switch (type) {
    case XlsbRecordType::BrtEndSheet:
    case XlsbRecordType::BrtBeginMergeCells:
    case XlsbRecordType::BrtMergeCell:
    case XlsbRecordType::BrtEndMergeCells:
      return true;
    default:
      return false;
  }
}

/// Appends the framed bytes of one worksheet-tail record to whichever
/// `XlsbSheetTail` buffer matches its position relative to the merged-cell
/// block. `framed` spans the record header *and* payload, so re-emission is
/// a plain byte copy rather than a re-encode.
void RetainTailRecord(SheetDecodeState& state, XlsbRecordType type, const std::uint8_t* framed, std::size_t size) {
  if (IsModelOwnedTailRecord(type)) {
    return;
  }
  std::vector<std::uint8_t>& dst = state.merges_seen ? state.tail.after_merges : state.tail.before_merges;
  dst.insert(dst.end(), framed, framed + size);
}

/// Reads every entry of a worksheet's `.rels` part, keeping the original
/// relationship ids so the retained tail records still resolve. Internal
/// targets are normalised to package-relative paths (matching what the OOXML
/// reader stores); external targets stay verbatim.
Expected<std::vector<io::UnknownRelationship>, Error> LoadSheetRelationships(const ZipReader& zip,
                                                                             std::string_view rels_path,
                                                                             std::string_view sheet_dir) {
  std::vector<io::UnknownRelationship> out;
  auto status = VisitRelationshipNodes(zip, rels_path, "sheet rels", [&](const pugi::xml_node& rel) {
    io::UnknownRelationship entry;
    entry.id = rel.attribute("Id").value();
    if (entry.id.empty()) {
      return Expected<void, Error>::Ok();
    }
    entry.type = rel.attribute("Type").value();
    const std::string_view target = rel.attribute("Target").value();
    entry.target_external = std::string_view(rel.attribute("TargetMode").value()) == "External";
    if (entry.target_external) {
      entry.target = std::string(target);
    } else {
      auto resolved = ResolveRelativePath(sheet_dir, target);
      if (!resolved) {
        return Expected<void, Error>(resolved.error());
      }
      entry.target = std::move(resolved).value();
    }
    out.push_back(std::move(entry));
    return Expected<void, Error>::Ok();
  });
  if (!status) {
    return status.error();
  }
  return out;
}

/// Column + style-table index decoded by `ReadCellHeader`.
struct CellHeaderInfo {
  std::uint32_t col = 0;
  /// 0-based index into `StylesTable::cell_xfs`. `0` is the default xf
  /// and is never stored on a cell (mirrors the OOXML reader's
  /// `xf_index != 0` guard before calling `set_cell_xf_index`).
  std::uint32_t xf_index = 0;
};

/// Reads the eight-byte cell header common to every cell record:
///   * column   : u32 (zero-based)
///   * iStyleRef: u24 (index into `StylesTable::cell_xfs`)
///   * fPhShow  : u8  (Phonetic-text flag; ignored)
///
/// Returns the column index and style-xf index, and advances the
/// cursor past the header.
Expected<CellHeaderInfo, Error> ReadCellHeader(ByteSpan& cursor) {
  auto col_or = read_u32(cursor);
  if (!col_or) {
    return col_or.error();
  }
  if (col_or.value() >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb cell header column out of range",
                      "context=xlsb_reader");
  }
  // iStyleRef (3 bytes, little-endian) + fPhShow (1 byte) = 4 bytes.
  if (cursor.size < 4) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb cell header truncated (style/phonetic)",
                      "context=xlsb_reader");
  }
  const std::uint32_t xf_index = static_cast<std::uint32_t>(cursor.data[0]) |
                                 (static_cast<std::uint32_t>(cursor.data[1]) << 8) |
                                 (static_cast<std::uint32_t>(cursor.data[2]) << 16);
  cursor.data += 4;
  cursor.size -= 4;
  return CellHeaderInfo{col_or.value(), xf_index};
}

/// Stores `xf_index` on `(sheet_index, row, col)` when it differs from
/// the default xf (`0`), mirroring the OOXML sheet reader's
/// `xf_index != 0` guard before calling `Workbook::set_cell_xf_index`.
Expected<void, Error> ApplyXfIndex(Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                   std::uint32_t xf_index) {
  if (xf_index == 0) {
    return Expected<void, Error>::Ok();
  }
  return wb.set_cell_xf_index(sheet_index, row, col, xf_index);
}

/// Registers each recorded dynamic-array anchor as a spill region so the
/// footprint's non-anchor cells' raw literal payload (Excel writes a
/// plain `BrtCellRk` / `BrtCellReal` / ... for the cached spill targets
/// of e.g. `=SEQUENCE(3)` in F6, spilling into F7:F8) does not read back
/// as independent literals that block the anchor's own re-spill on
/// recalc. Mirrors the OOXML reader's `RegisterArraySpills`
/// (`io/sheet_reader.cpp`) byte-for-byte in intent: capture the
/// footprint's current cached values, blank the non-anchor cells, then
/// commit the spill region. Must run only after the entire sheet has
/// been decoded -- see the `BrtArrFmla` case's comment in
/// `DecodeSheetBin` for why registering inline (while later rows in the
/// footprint have not been decoded yet) does not work.
void RegisterArraySpills(Workbook& wb, std::size_t sheet_index, const std::vector<ArrayAnchor>& anchors) {
  for (const ArrayAnchor& a : anchors) {
    // Reversed rects are rejected at decode time (`BrtArrFmla` case in
    // `DecodeSheetBin`); guard again here so the size math below can
    // never underflow-wrap.
    if (a.last_row < a.row || a.last_col < a.col) {
      continue;
    }
    const std::uint32_t rows = a.last_row - a.row + 1U;
    const std::uint32_t cols = a.last_col - a.col + 1U;
    std::vector<Value> values;
    values.reserve(static_cast<std::size_t>(rows) * cols);
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        const Cell* cell = wb.sheet(sheet_index).cell_at(r, c);
        values.push_back(cell != nullptr ? cell->cached_value : Value::blank());
      }
    }
    for (std::uint32_t r = a.row; r <= a.last_row; ++r) {
      for (std::uint32_t c = a.col; c <= a.last_col; ++c) {
        if (r == a.row && c == a.col) {
          continue;
        }
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(r, c, Value::blank());
      }
    }
    wb.sheet(sheet_index).commit_spill(a.row, a.col, rows, cols, std::move(values));
  }
}

/// Decodes one sheet binary part. Cells (literal + formula) flow into
/// `wb.sheet(sheet_index)`. SST indices are resolved against
/// `sst_entries`; out-of-range indices are returned as
/// `kIoXlsbCorrupt`.
Expected<SheetDecodeState, Error> DecodeSheetBin(const std::vector<std::uint8_t>& body, std::size_t sheet_index,
                                                 Workbook& wb, const std::vector<std::string_view>& sst_entries,
                                                 std::deque<std::string>& text_storage,
                                                 const std::vector<std::string>& sheet_names,
                                                 const std::vector<XlsbName>& name_table,
                                                 const std::vector<XlsbSheetRange>& sheet_ranges,
                                                 std::uint32_t* undecoded_formula_count) {
  SheetDecodeState state;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    const std::uint8_t* const framed = cursor.data;
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    const auto type = static_cast<XlsbRecordType>(rec.type);
    // Retain worksheet-tail records before dispatching: the tail carries the
    // sheet-level features the model does not express, and the sheet part is
    // consumed whole so package passthrough cannot rescue them.
    if (state.in_tail) {
      RetainTailRecord(state, type, framed, static_cast<std::size_t>(cursor.data - framed));
    }
    if (type == XlsbRecordType::BrtEndSheetData) {
      state.in_tail = true;
    } else if (type == XlsbRecordType::BrtEndMergeCells) {
      state.merges_seen = true;
    }
    switch (type) {
      case XlsbRecordType::BrtBeginWsView: {
        // BrtBeginWsView ([MS-XLSB] §2.4.141) stores the SheetView fields
        // modelled by Formulon in its flags word and wScale.  Other viewport
        // state is intentionally not represented by SheetView yet.
        ByteSpan p = rec.payload;
        auto flags_or = read_u16(p);
        auto skip_view = read_u32(p);
        auto skip_top = read_u32(p);
        auto skip_left = read_u32(p);
        auto skip_color = read_u8(p);
        auto skip_reserved8 = read_u8(p);
        auto skip_reserved16 = read_u16(p);
        auto zoom_or = read_u16(p);
        if (!flags_or || !skip_view || !skip_top || !skip_left || !skip_color || !skip_reserved8 || !skip_reserved16 ||
            !zoom_or) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtBeginWsView truncated",
                            "context=xlsb_reader");
        }
        SheetView& view = wb.sheet(sheet_index).mutable_view();
        const std::uint16_t flags = flags_or.value();
        view.show_grid_lines = (flags & 0x0004U) != 0U;
        view.show_row_col_headers = (flags & 0x0008U) != 0U;
        view.show_zeros = (flags & 0x0010U) != 0U;
        view.right_to_left = (flags & 0x0020U) != 0U;
        view.tab_selected = (flags & 0x0040U) != 0U;
        if (zoom_or.value() >= 10U && zoom_or.value() <= 400U) {
          view.zoom_scale = zoom_or.value();
        }
        break;
      }
      case XlsbRecordType::BrtPane: {
        // BrtPane ([MS-XLSB] §2.4.723): two Xnum split positions, then
        // the lower-right pane's top-left cell and the frozen-state flags.
        // In a frozen pane the Xnum fields are integral row/column counts.
        ByteSpan p = rec.payload;
        if (p.size < 29U) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane truncated", "context=xlsb_reader");
        }
        double rows = 0.0;
        double cols = 0.0;
        std::memcpy(&rows, p.data, sizeof(rows));
        std::memcpy(&cols, p.data + sizeof(rows), sizeof(cols));
        p.data += 16U;
        p.size -= 16U;
        auto top_row_or = read_u32(p);
        auto left_col_or = read_u32(p);
        auto active_pane_or = read_u32(p);
        auto flags_or = read_u8(p);
        if (!top_row_or || !left_col_or || !active_pane_or || !flags_or) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtPane fields truncated",
                            "context=xlsb_reader");
        }
        const std::uint8_t frozen_flags = static_cast<std::uint8_t>(flags_or.value() & 0x03U);
        if (frozen_flags == 0U) {
          break;
        }
        if (frozen_flags == 0x03U || !std::isfinite(rows) || !std::isfinite(cols) || rows < 0.0 || cols < 0.0 ||
            std::trunc(rows) != rows || std::trunc(cols) != cols || rows >= Sheet::kMaxRows ||
            cols >= Sheet::kMaxCols) {
          return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtPane frozen counts invalid",
                            "context=xlsb_reader");
        }
        SheetView& view = wb.sheet(sheet_index).mutable_view();
        view.freeze_rows = static_cast<std::uint32_t>(rows);
        view.freeze_cols = static_cast<std::uint32_t>(cols);
        break;
      }
      case XlsbRecordType::BrtRowHdr: {
        // BrtRowHdr ([MS-XLSB] §2.4.770): rw, ixfe, miyRw, two flag
        // bytes, fPhShow, then ccolspan. Older minimal producers may provide only
        // rw, which remains enough for cell decoding.
        ByteSpan p = rec.payload;
        auto row_or = read_u32(p);
        if (!row_or) {
          return row_or.error();
        }
        if (row_or.value() >= Sheet::kMaxRows) {
          return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtRowHdr row out of range",
                            "context=xlsb_reader");
        }
        state.current_row = row_or.value();
        state.row_seen = true;
        if (p.size >= 9U) {
          auto skip_xf = read_u32(p);
          auto height_or = read_u16(p);
          auto flags1_or = read_u8(p);
          auto flags2_or = read_u8(p);
          auto skip_phonetic_show = read_u8(p);
          if (!skip_xf || !height_or || !flags1_or || !flags2_or || !skip_phonetic_show) {
            return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtRowHdr layout truncated",
                              "context=xlsb_reader");
          }
          const std::uint8_t flags = flags2_or.value();
          if (height_or.value() != 0U || (flags & 0x10U) != 0U || (flags & 0x07U) != 0U) {
            RowLayout layout;
            layout.row = state.current_row;
            layout.height = static_cast<double>(height_or.value()) / 20.0;
            layout.has_height = height_or.value() != 0U || (flags & 0x20U) != 0U;
            layout.hidden = (flags & 0x10U) != 0U;
            layout.outline_level = static_cast<std::uint8_t>(flags & 0x07U);
            wb.sheet(sheet_index).mutable_layout().row_overrides.push_back(std::move(layout));
          }
        }
        break;
      }
      case XlsbRecordType::BrtColInfo: {
        // BrtColInfo ([MS-XLSB] §2.4.336): colFirst, colLast, coldx,
        // ixfe, flags(u16). coldx is in 1/256 standard digits.
        ByteSpan p = rec.payload;
        auto first_or = read_u32(p);
        auto last_or = read_u32(p);
        auto width_or = read_u32(p);
        auto skip_xf = read_u32(p);
        auto flags_or = read_u16(p);
        if (!first_or || !last_or || !width_or || !skip_xf || !flags_or) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtColInfo truncated",
                            "context=xlsb_reader");
        }
        if (first_or.value() >= Sheet::kMaxCols || last_or.value() >= Sheet::kMaxCols ||
            first_or.value() > last_or.value()) {
          return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtColInfo range out of bounds",
                            "context=xlsb_reader");
        }
        ColumnLayout layout;
        layout.first = first_or.value();
        layout.last = last_or.value();
        layout.width = static_cast<double>(width_or.value()) / 256.0;
        layout.hidden = (flags_or.value() & 0x0001U) != 0U;
        layout.outline_level = static_cast<std::uint8_t>((flags_or.value() >> 8U) & 0x07U);
        wb.sheet(sheet_index).mutable_layout().columns.push_back(std::move(layout));
        break;
      }
      case XlsbRecordType::BrtMergeCell: {
        // BrtMergeCell ([MS-XLSB] §2.4.713): RfX = first/last row and
        // first/last column, all u32.  Containers are structural only; a
        // record may appear after sheet data and therefore must not depend on
        // the active BrtRowHdr state.
        ByteSpan p = rec.payload;
        auto first_row_or = read_u32(p);
        auto last_row_or = read_u32(p);
        auto first_col_or = read_u32(p);
        auto last_col_or = read_u32(p);
        if (!first_row_or || !last_row_or || !first_col_or || !last_col_or) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtMergeCell truncated",
                            "context=xlsb_reader");
        }
        if (first_row_or.value() >= Sheet::kMaxRows || last_row_or.value() >= Sheet::kMaxRows ||
            first_col_or.value() >= Sheet::kMaxCols || last_col_or.value() >= Sheet::kMaxCols ||
            first_row_or.value() > last_row_or.value() || first_col_or.value() > last_col_or.value()) {
          return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtMergeCell range out of bounds",
                            "context=xlsb_reader");
        }
        wb.sheet(sheet_index)
            .mutable_merges()
            .push_back(
                MergeRange{first_row_or.value(), first_col_or.value(), last_row_or.value(), last_col_or.value()});
        break;
      }
      case XlsbRecordType::BrtCellBlank: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::blank());
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellRk: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        auto rk_or = read_u32(p);
        if (!rk_or) {
          return rk_or.error();
        }
        const double v = decode_rk_number(rk_or.value());
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::number(v));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellReal: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        if (p.size < 8) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtCellReal payload truncated",
                            "context=xlsb_reader");
        }
        double v;
        std::memcpy(&v, p.data, sizeof(v));
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::number(v));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellBool: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        auto b_or = read_u8(p);
        if (!b_or) {
          return b_or.error();
        }
        wb.sheet(sheet_index)
            .set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::boolean(b_or.value() != 0));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellError: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        auto code_or = read_u8(p);
        if (!code_or) {
          return code_or.error();
        }
        // Map the OOXML wire code to `ErrorCode` via the single
        // `kErrorTable`-backed lookup shared with the writer, so this
        // path stays symmetric with `ooxml_code()` (see
        // `error_from_ooxml_code` in `value.h`).
        const ErrorCode ec = error_from_ooxml_code(static_cast<std::int32_t>(code_or.value()));
        wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::error(ec));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellSt: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        auto str_or = read_xlwidestring(p);
        if (!str_or) {
          return str_or.error();
        }
        text_storage.push_back(std::move(str_or.value()));
        wb.sheet(sheet_index)
            .set_cell_cached_value_borrowed(state.current_row, col_or.value().col, Value::text(text_storage.back()));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtCellIsst: {
        if (!state.row_seen) {
          break;
        }
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        auto idx_or = read_u32(p);
        if (!idx_or) {
          return idx_or.error();
        }
        if (idx_or.value() >= sst_entries.size()) {
          std::string ctx("context=xlsb_reader sheet_index=");
          ctx.append(std::to_string(sheet_index));
          ctx.append(" row=").append(std::to_string(state.current_row));
          ctx.append(" col=").append(std::to_string(col_or.value().col));
          ctx.append(" sst_index=").append(std::to_string(idx_or.value()));
          ctx.append(" sst_size=").append(std::to_string(sst_entries.size()));
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb sst index out of range", std::move(ctx));
        }
        wb.sheet(sheet_index)
            .set_cell_cached_value_borrowed(state.current_row, col_or.value().col,
                                            Value::text(sst_entries[idx_or.value()]));
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtFmlaNum:
      case XlsbRecordType::BrtFmlaString:
      case XlsbRecordType::BrtFmlaBool:
      case XlsbRecordType::BrtFmlaError: {
        if (!state.row_seen) {
          break;
        }
        // All four formula records share the same prefix:
        //   cell-header (8 bytes)
        //   value       (8 bytes for Num, XLWideString for String, 2
        //                bytes for Bool/Error: u8 result + u8 flags)
        //   flags       (u16 grbitFlags)
        //   ptgs        (CellParsedFormula = u32 cce + cce bytes of
        //                 Rgce + ... we just slice the remainder as
        //                 opaque bytes for now)
        ByteSpan p = rec.payload;
        auto col_or = ReadCellHeader(p);
        if (!col_or) {
          return col_or.error();
        }
        // Decode the formula's cached result so we can PRESERVE it on the
        // cell even when the Ptg stream cannot be decoded to a formula.
        Value cached = Value::blank();
        switch (type) {
          case XlsbRecordType::BrtFmlaNum: {
            if (p.size < 8) {
              return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFmlaNum value truncated",
                                "context=xlsb_reader");
            }
            double v;
            std::memcpy(&v, p.data, sizeof(v));
            cached = Value::number(v);
            p.data += 8;
            p.size -= 8;
            break;
          }
          case XlsbRecordType::BrtFmlaString: {
            auto s = read_xlwidestring(p);
            if (!s) {
              return s.error();
            }
            text_storage.push_back(std::move(s.value()));
            cached = Value::text(text_storage.back());
            break;
          }
          case XlsbRecordType::BrtFmlaBool: {
            auto b = read_u8(p);
            if (!b) {
              return b.error();
            }
            cached = Value::boolean(b.value() != 0);
            break;
          }
          case XlsbRecordType::BrtFmlaError: {
            auto b = read_u8(p);
            if (!b) {
              return b.error();
            }
            // Same `kErrorTable`-backed lookup the literal path uses.
            cached = Value::error(error_from_ooxml_code(static_cast<std::int32_t>(b.value())));
            break;
          }
          default:
            break;
        }
        // grbitFlags (u16).
        if (p.size < 2) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula flags truncated",
                            "context=xlsb_reader");
        }
        p.data += 2;
        p.size -= 2;
        // CellParsedFormula: u32 cce (rgce byte length) + cce bytes of
        // Ptg stream + u32 cb + cb bytes of rgcb (the array-constant
        // extra-data area `PtgArray` consumes; empty for formulas with
        // no array literals).
        auto cce_or = read_u32(p);
        if (!cce_or) {
          return cce_or.error();
        }
        const std::uint32_t cce = cce_or.value();
        if (cce > p.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula rgce length exceeds payload",
                            "context=xlsb_reader");
        }
        ByteSpan rgce{p.data, cce};
        p.data += cce;
        p.size -= cce;
        ByteSpan rgcb{};
        if (p.size >= 4) {
          auto cb_or = read_u32(p);
          if (!cb_or) {
            return cb_or.error();
          }
          const std::uint32_t cb = cb_or.value();
          if (cb > p.size) {
            return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb formula rgcb length exceeds payload",
                              "context=xlsb_reader");
          }
          rgcb = ByteSpan{p.data, cb};
        }
        const std::string formula_text =
            DecodeFormulaText(rgce, rgcb, sheet_names, name_table, sheet_ranges, sheet_index, state.current_row,
                              col_or.value().col, undecoded_formula_count);
        if (!formula_text.empty()) {
          // Register the real formula via the workbook-level entry so the
          // dep graph tracks it (matching the OOXML reader). The cached
          // value is preserved separately below.
          auto wf = wb.set_cell_formula(sheet_index, state.current_row, col_or.value().col, formula_text);
          if (!wf) {
            return wf.error();
          }
        }
        // Always preserve the cached value (for undecodable formulas this
        // is the only correct datum the cell carries; for decoded ones it
        // matches Excel's stored result until the next recalc).
        if (!cached.is_blank()) {
          wb.sheet(sheet_index).set_cell_cached_value_borrowed(state.current_row, col_or.value().col, cached);
        }
        if (auto r = ApplyXfIndex(wb, sheet_index, state.current_row, col_or.value().col, col_or.value().xf_index);
            !r) {
          return r.error();
        }
        ++state.cells_decoded;
        break;
      }
      case XlsbRecordType::BrtArrFmla: {
        if (!state.row_seen) {
          break;
        }
        // BrtArrFmla ([MS-XLSB], verified against a real Excel-produced
        // `xl/worksheets/sheetN.bin`): RfX (4 x u32: rwFirst, rwLast,
        // colFirst, colLast) + 1 reserved/flag byte + CellParsedFormula
        // (u32 cce + cce bytes rgce + u32 cb + cb bytes rgcb). This
        // record supplies the REAL Ptg tokens for a CSE / dynamic-array
        // formula whose anchor cell's own formula-result "shell" record
        // (BrtFmlaNum/String/Bool/Error, processed above) carries only a
        // `PtgExp` placeholder. Only the anchor cell (`rwFirst`,
        // `colFirst`) gets a formula string; the rest of the array range
        // (if any) has no formula of its own — matching the OOXML
        // reader's treatment of `t="array"` / dynamic-array spill
        // formulas, where only the anchor cell stores `<f>`.
        ByteSpan p = rec.payload;
        auto rw_first_or = read_u32(p);
        if (!rw_first_or) {
          return rw_first_or.error();
        }
        auto rw_last_or = read_u32(p);
        if (!rw_last_or) {
          return rw_last_or.error();
        }
        auto col_first_or = read_u32(p);
        if (!col_first_or) {
          return col_first_or.error();
        }
        auto col_last_or = read_u32(p);
        if (!col_last_or) {
          return col_last_or.error();
        }
        // The RfX rect must lie inside the grid and be well-ordered on
        // BOTH axes before any of it is used: the anchor guard below is
        // an OR, so a rect reversed on only one axis would otherwise
        // still be recorded and wrap the size math in
        // `RegisterArraySpills`.
        if (rw_first_or.value() >= Sheet::kMaxRows || rw_last_or.value() >= Sheet::kMaxRows ||
            col_first_or.value() >= Sheet::kMaxCols || col_last_or.value() >= Sheet::kMaxCols ||
            rw_last_or.value() < rw_first_or.value() || col_last_or.value() < col_first_or.value()) {
          return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtArrFmla range out of bounds",
                            "context=xlsb_reader");
        }
        if (p.size < 1) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla flag truncated",
                            "context=xlsb_reader");
        }
        p.data += 1;
        p.size -= 1;
        auto cce_or = read_u32(p);
        if (!cce_or) {
          return cce_or.error();
        }
        const std::uint32_t cce = cce_or.value();
        if (cce > p.size) {
          return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla rgce length exceeds payload",
                            "context=xlsb_reader");
        }
        ByteSpan rgce{p.data, cce};
        p.data += cce;
        p.size -= cce;
        ByteSpan rgcb{};
        if (p.size >= 4) {
          auto cb_or = read_u32(p);
          if (!cb_or) {
            return cb_or.error();
          }
          const std::uint32_t cb = cb_or.value();
          if (cb > p.size) {
            return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtArrFmla rgcb length exceeds payload",
                              "context=xlsb_reader");
          }
          rgcb = ByteSpan{p.data, cb};
        }
        const std::string formula_text =
            DecodeFormulaText(rgce, rgcb, sheet_names, name_table, sheet_ranges, sheet_index, rw_first_or.value(),
                              col_first_or.value(), undecoded_formula_count);
        if (formula_text.empty()) {
          break;
        }
        auto wf = wb.set_cell_formula(sheet_index, rw_first_or.value(), col_first_or.value(), formula_text);
        if (!wf) {
          return wf.error();
        }
        // Record the footprint for a second pass after the whole sheet
        // has been decoded (see `RegisterArraySpills`, called at the end
        // of this function). `BrtArrFmla` for the anchor `(rwFirst,
        // colFirst)` appears within the anchor's OWN row group in the
        // record stream, i.e. BEFORE the non-anchor rows' own cell
        // records (`rwLast > rwFirst` spans into rows not yet decoded).
        // Registering the spill immediately here would have those later
        // records' `set_cell_cached_value` calls silently repopulate the
        // phantom cells with their raw literal payload, which a
        // subsequent recalc's spill-commit would then see as "already
        // occupied" and surface `#SPILL!` instead of the real result.
        state.array_anchors.push_back(
            ArrayAnchor{rw_first_or.value(), col_first_or.value(), rw_last_or.value(), col_last_or.value()});
        break;
      }
      default:
        break;
    }
  }
  RegisterArraySpills(wb, sheet_index, state.array_anchors);
  if (!state.tail.empty()) {
    wb.sheet(sheet_index).set_xlsb_tail(state.tail);
  }
  return state;
}

}  // namespace

Expected<XlsbReadResult, Error> read_xlsb(ByteSpan bytes) {
  ZipReader zip;
  if (auto open_result = zip.open(bytes); !open_result) {
    return open_result.error();
  }

  // 1. [Content_Types].xml — gate + Override list.
  if (!zip.has_entry("[Content_Types].xml")) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "[Content_Types].xml: missing from package",
                      "context=xlsb_reader");
  }
  auto ct_bytes_or = zip.read_entry("[Content_Types].xml");
  if (!ct_bytes_or) {
    return ct_bytes_or.error();
  }
  auto ct_view_or = LoadContentTypes(ct_bytes_or.value());
  if (!ct_view_or) {
    return ct_view_or.error();
  }
  ContentTypesView ct_view = ct_view_or.take();

  // 2. _rels/.rels — locate the workbook part path.
  if (!zip.has_entry("_rels/.rels")) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "_rels/.rels: missing from package",
                      "context=xlsb_reader");
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
  auto package_rels_or = ReadUnknownPackageRels(root_rels_or.value());
  if (!package_rels_or) {
    return package_rels_or.error();
  }

  // 3. xl/_rels/workbook.xml.rels (still XML in xlsb).
  auto wb_rels_or = LoadWorkbookRels(zip, workbook_path);
  if (!wb_rels_or) {
    return wb_rels_or.error();
  }
  const WorkbookRels& wb_rels = wb_rels_or.value();

  // 4. xl/workbook.bin — sheet bundle list.
  if (!zip.has_entry(workbook_path)) {
    return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: not found at relationship target",
                      "context=xlsb_reader workbook_path=" + workbook_path);
  }
  auto wb_bytes_or = zip.read_entry(workbook_path);
  if (!wb_bytes_or) {
    return wb_bytes_or.error();
  }
  auto bundle_or = DecodeWorkbookBin(wb_bytes_or.value());
  if (!bundle_or) {
    return bundle_or.error();
  }
  const WorkbookBinInfo& workbook_info = bundle_or.value();
  const std::vector<SheetBundleEntry>& bundle = workbook_info.sheets;

  // `BrtName` (defined names + hidden future-function / LET-parameter
  // placeholders) and `BrtExternSheet` (qualified-reference sheet
  // ranges) both live in `xl/workbook.bin` globals; decode them once so
  // every sheet's `PtgName` / `PtgRef3d` / `PtgArea3d` tokens can
  // resolve against the same tables.
  auto name_table_or = DecodeWorkbookNames(wb_bytes_or.value());
  if (!name_table_or) {
    return name_table_or.error();
  }
  const std::vector<XlsbName>& name_table = name_table_or.value();
  auto sheet_ranges_or = DecodeExternSheet(wb_bytes_or.value());
  if (!sheet_ranges_or) {
    return sheet_ranges_or.error();
  }
  const std::vector<XlsbSheetRange>& sheet_ranges = sheet_ranges_or.value();

  // 5. Build the workbook bottom-up.
  Workbook wb = Workbook::create_empty();
  wb.set_date1904(workbook_info.date1904);
  std::vector<std::string> sheet_part_paths;
  sheet_part_paths.reserve(bundle.size());
  for (const SheetBundleEntry& b : bundle) {
    if (b.name.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: BrtBundleSh with empty name",
                        "context=xlsb_reader");
    }
    if (b.rid.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: BrtBundleSh with empty rId",
                        "context=xlsb_reader sheet=" + b.name);
    }
    auto it = wb_rels.sheet_targets.find(b.rid);
    if (it == wb_rels.sheet_targets.end()) {
      std::string ctx("context=xlsb_reader rid=");
      ctx.append(b.rid);
      ctx.append(" sheet=").append(b.name);
      return make_error(FormulonErrorCode::kIoRelationshipBroken,
                        "workbook.bin: BrtBundleSh rId has no matching workbook relationship", std::move(ctx));
    }
    wb.add_sheet(b.name);
    if (b.hidden) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().tab_hidden = true;
    }
    sheet_part_paths.push_back(it->second);
  }

  // Ordered sheet display names — the Ptg decoder maps a 3-D reference's
  // `ixti` (0-based sheet index) to the qualifying sheet name through
  // this list.
  std::vector<std::string> sheet_names;
  sheet_names.reserve(bundle.size());
  for (const SheetBundleEntry& b : bundle) {
    sheet_names.push_back(b.name);
  }

  // 5b. Register every non-hidden `BrtName` entry as a defined name.
  // Needs `sheet_names` (for qualified references inside a name's own
  // formula) and the complete `name_table` / `sheet_ranges` (for
  // self-consistent `PtgName` / `PtgRef3d` resolution), so this can
  // only run after both are fully built above.
  std::uint32_t undecoded_formula_count = 0;
  std::uint32_t undecoded_defined_name_count = 0;
  if (auto r = RegisterDefinedNames(wb_bytes_or.value(), name_table, sheet_names, sheet_ranges, wb,
                                    &undecoded_defined_name_count);
      !r) {
    return r.error();
  }

  // 6. xl/sharedStrings.bin — load before the per-sheet decode loop so
  // BrtCellIsst can resolve indices in-pipeline. The text deque is
  // owned by the workbook itself so `Value::text` views remain valid
  // after the caller moves the workbook out of the read result.
  std::deque<std::string>& text_storage = wb.mutable_text_storage();
  std::vector<std::string_view> sst_entries;
  if (!wb_rels.sst_path.empty() && zip.has_entry(wb_rels.sst_path)) {
    auto sst_bytes_or = zip.read_entry(wb_rels.sst_path);
    if (!sst_bytes_or) {
      return sst_bytes_or.error();
    }
    auto sst_or = DecodeSharedStringsBin(sst_bytes_or.value(), text_storage);
    if (!sst_or) {
      return sst_or.error();
    }
    sst_entries = std::move(sst_or.value());
  }

  // 6b. xl/styles.bin — numFmt + cellXfs/cellStyleXfs (see
  // `io/xlsb/styles_reader.h`). Deliberately NOT added to
  // `consumed_parts` below: leaving it out lets step 8's passthrough
  // loop capture the raw bytes (with the correct `[Content_Types].xml`
  // content-type) alongside the parsed `StylesTable`, so a
  // read-modify-write cycle keeps the original font/fill/border detail
  // this reader does not model in-memory.
  if (!wb_rels.styles_path.empty() && zip.has_entry(wb_rels.styles_path)) {
    auto styles_bytes_or = zip.read_entry(wb_rels.styles_path);
    if (!styles_bytes_or) {
      return styles_bytes_or.error();
    }
    const std::vector<std::uint8_t>& styles_bytes = styles_bytes_or.value();
    ByteSpan styles_span{styles_bytes.data(), styles_bytes.size()};
    auto styles_or = read_styles_bin(styles_span);
    if (!styles_or) {
      return styles_or.error();
    }
    wb.set_styles(std::move(styles_or.value()));
  }

  // 7. Each sheet binary.
  std::uint32_t cells_read = 0;
  std::unordered_set<std::string> consumed_parts;
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  consumed_parts.insert(RelsPathForPart(workbook_path));
  if (!wb_rels.sst_path.empty()) {
    consumed_parts.insert(wb_rels.sst_path);
  }

  for (std::size_t i = 0; i < sheet_part_paths.size(); ++i) {
    const std::string& sheet_path = sheet_part_paths[i];
    if (!zip.has_entry(sheet_path)) {
      std::string ctx("context=xlsb_reader sheet_path=");
      ctx.append(sheet_path);
      return make_error(FormulonErrorCode::kIoXlsbCorrupt, "sheet binary part missing from package", std::move(ctx));
    }
    auto sheet_bytes_or = zip.read_entry(sheet_path);
    if (!sheet_bytes_or) {
      return sheet_bytes_or.error();
    }
    auto state_or = DecodeSheetBin(sheet_bytes_or.value(), i, wb, sst_entries, text_storage, sheet_names, name_table,
                                   sheet_ranges, &undecoded_formula_count);
    if (!state_or) {
      return state_or.error();
    }
    cells_read += state_or.value().cells_decoded;
    consumed_parts.insert(sheet_path);

    // The sheet's own rels file resolves the relationship ids carried by the
    // retained tail records (hyperlink targets, drawing and table parts). It
    // is `rels`-Default-typed, so the Override-driven passthrough loop below
    // never sees it; keeping every entry — none of them are modelled by the
    // binary reader — preserves both the ids and their targets.
    const std::string sheet_rels_path = RelsPathForPart(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      auto rels_or = LoadSheetRelationships(zip, sheet_rels_path, DirOf(sheet_path));
      if (!rels_or) {
        return rels_or.error();
      }
      wb.sheet(i).set_unknown_relationships(std::move(rels_or.value()));
      consumed_parts.insert(sheet_rels_path);
    }
  }

  // 8. Passthrough parts. First capture every Override-listed part the
  // reader did not consume, retaining its explicit content type. A second
  // residual sweep resolves the remaining entries through the source
  // Default registry. Default-typed parts intentionally carry an empty
  // content_type; the registry itself travels on the Workbook and the
  // writer re-emits the matching Default entry.
  std::vector<PassthroughPart> unknown_parts;
  unknown_parts.reserve(ct_view.overrides.size());
  std::unordered_set<std::string> captured_parts;
  captured_parts.reserve(ct_view.overrides.size());
  for (const auto& [part_name, content_type] : ct_view.overrides) {
    if (consumed_parts.find(part_name) != consumed_parts.end()) {
      continue;
    }
    if (captured_parts.find(part_name) != captured_parts.end()) {
      continue;
    }
    // Refuse a traversal-shaped passthrough name so a round-tripped .xlsb
    // never hands a downstream extractor a zip-slip primitive.
    if (!ooxml::is_safe_part_name(part_name)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "Override part name escapes package root; refusing to load",
                        "context=xlsb_reader part=" + part_name);
    }
    if (!zip.has_entry(part_name)) {
      continue;
    }
    auto bytes_or = zip.read_entry(part_name);
    if (!bytes_or) {
      return bytes_or.error();
    }
    PassthroughPart part;
    part.path = part_name;
    part.content_type = content_type;
    part.bytes = std::move(bytes_or.value());
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(part_name);
  }

  auto default_content_type_for = [&ct_view](std::string_view path) -> const DefaultContentType* {
    const std::string extension = ExtensionOfPart(path);
    if (extension.empty()) {
      return nullptr;
    }
    const auto it =
        std::find_if(ct_view.defaults.begin(), ct_view.defaults.end(),
                     [&extension](const DefaultContentType& value) { return value.extension == extension; });
    return it == ct_view.defaults.end() ? nullptr : &*it;
  };

  std::uint32_t dropped_part_count = 0;
  std::string first_dropped;
  for (const std::string& entry : zip.list_entries()) {
    // Directory markers are not OPC parts. Skip them before path validation:
    // ZIP producers commonly record a trailing-slash directory entry, and
    // the trailing slash is intentionally not a canonical OPC part name.
    // Payload entries are validated below, including consumed paths, so a
    // hostile ZIP catalogue cannot hide a traversal-shaped name behind the
    // modelled path set.
    if (entry.empty() || entry.back() == '/') {
      continue;
    }
    if (!ooxml::is_safe_part_name(entry)) {
      return make_error(FormulonErrorCode::kIoZipSlip, "archive entry name escapes package root; refusing to load",
                        "context=xlsb_reader part=" + entry);
    }
    if (consumed_parts.find(entry) != consumed_parts.end() || captured_parts.find(entry) != captured_parts.end()) {
      continue;
    }
    const DefaultContentType* default_type = default_content_type_for(entry);
    if (default_type == nullptr || default_type->content_type.empty()) {
      if (dropped_part_count == 0U) {
        first_dropped = entry;
      }
      ++dropped_part_count;
      continue;
    }
    auto bytes_or = zip.read_entry(entry);
    if (!bytes_or) {
      return bytes_or.error();
    }
    PassthroughPart part;
    part.path = entry;
    // Default-typed parts deliberately do not copy the effective content
    // type into the per-part record. This keeps Override and Default
    // semantics distinguishable on the write side.
    part.bytes = std::move(bytes_or.value());
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(entry);
  }

  std::sort(unknown_parts.begin(), unknown_parts.end(),
            [](const PassthroughPart& a, const PassthroughPart& b) { return a.path < b.path; });

  if (dropped_part_count != 0U) {
    StructuredLog("xlsb.package.parts_dropped")
        .field("count", static_cast<std::int64_t>(dropped_part_count))
        .field("first_part", first_dropped)
        .field("reason", std::string_view("part content type could not be resolved from Override or Default"))
        .warn();
  }

  // The workbook is the sole owner; the read result does not mirror the
  // payload. See `XlsbReadResult`.
  wb.set_passthrough_parts(std::move(unknown_parts));
  wb.set_default_content_types(std::move(ct_view.defaults));
  wb.set_unknown_package_rels(std::move(package_rels_or.value()));
  wb.set_unknown_workbook_rels(std::move(wb_rels_or.value().unknown_rels));

  XlsbReadResult result{std::move(wb), cells_read, undecoded_formula_count, undecoded_defined_name_count,
                        dropped_part_count};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
