//
// Implementation of the XLSB read pipeline. See `io/xlsb/reader.h` for
// the design context.
//
// The reader reuses the OOXML zip envelope plus three XML parts
// (`[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`).
// The binary parts (`xl/workbook.bin`, `xl/worksheets/sheet*.bin`,
// `xl/sharedStrings.bin`) are decoded via the record framing in
// `io/xlsb/record.h`; worksheets decode in `io/xlsb/sheet_reader.cpp`,
// the shared-string table in `io/xlsb/sst_reader.cpp`. Formulas are
// decoded through a full `Ptg → AST → Excel-formula-text` pipeline
// (`DecodeFormulaText` in `io/xlsb/sheet_reader.cpp`); a Ptg stream
// outside the supported token set logs a structured warning and leaves
// the formula text empty instead.

#include "io/xlsb/reader.h"

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <deque>
#include <limits>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "default_content_type.h"
#include "io/dynamic_array_formula.h"
#include "io/future_functions.h"
#include "io/ooxml/package_validator.h"
#include "io/ooxml/rels_walker.h"
#include "io/ooxml_defs.h"
#include "io/xlsb/external_link_reader.h"
#include "io/xlsb/metadata_bin.h"
#include "io/xlsb/pivot_reader.h"
#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/retained_part_fingerprint.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sst_reader.h"
#include "io/xlsb/styles_reader.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "parser/formula_prefix.h"
#include "passthrough_part.h"
#include "phonetic.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_table.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "unknown_relationship.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "utils/structured_log.h"
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
constexpr std::string_view kRelHyperlink =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
constexpr std::string_view kRelPivotTable =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";
constexpr std::string_view kRelPivotCacheDefinition =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
constexpr std::string_view kRelPivotCacheRecords =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";

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
      const std::string normalized_ext = ooxml::lowercase_extension(ext);
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
  const std::string rels_path = ooxml::rels_path_for_part(workbook_path);
  if (!zip.has_entry(rels_path)) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=xlsb_reader rels_path=" + rels_path);
  }
  const std::string base_dir = ooxml::dir_of(workbook_path);
  auto visit_status = ooxml::visit_relationship_nodes(
      zip, rels_path, "workbook rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
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
  /// The sheet's `hsState`, which shares its numbering with
  /// `SheetVisibility` and so is carried across whole. Very-hidden stays
  /// distinct from hidden: it is the state that keeps a sheet out of
  /// Excel's "Unhide" dialog.
  SheetVisibility visibility = SheetVisibility::kVisible;
};

/// Workbook-global fields needed while constructing the model.
struct WorkbookBinInfo {
  std::vector<SheetBundleEntry> sheets;
  bool date1904 = false;
  /// `<workbookProtection>` rebuilt from BrtBookProtection(Iso); empty
  /// when the book is unprotected or the records were not decodable.
  std::string protection_xml;
};

/// Decodes `xl/workbook.bin` to extract the ordered sheet-bundle list and
/// workbook date system and protection. Other records are skipped.
Expected<WorkbookBinInfo, Error> DecodeWorkbookBin(const std::vector<std::uint8_t>& body) {
  WorkbookBinInfo info;
  ByteSpan protection_iso{};
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
    if (rec.type == kBrtBookProtectionIso) {
      protection_iso = rec.payload;
      continue;
    }
    if (rec.type == kBrtBookProtection) {
      if (!decode_book_protection(rec.payload, protection_iso, info.protection_xml)) {
        info.protection_xml.clear();
        StructuredLog("xlsb.book_protection.not_decoded").warn();
      }
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
    // An hsState outside the three defined values is not a visibility this
    // model can name; treat anything non-zero it cannot place as plain
    // hidden, which is the conservative direction (the sheet stays out of
    // sight rather than appearing unbidden).
    switch (hs_state_or.value()) {
      case 0U:
        entry.visibility = SheetVisibility::kVisible;
        break;
      case 2U:
        entry.visibility = SheetVisibility::kVeryHidden;
        break;
      default:
        entry.visibility = SheetVisibility::kHidden;
        break;
    }
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

/// True when `name` carries one of Excel's hidden storage prefixes, i.e.
/// the record is a `_xlfn.<FN>` future-function or `_xlpm.<param>`
/// LET / LAMBDA-parameter placeholder rather than a user-visible defined
/// name. Matched case-insensitively, the same way `ptg_reader.cpp`
/// resolves these names during Ptg decode.
///
/// This is deliberately NOT the `fHidden` bit: Excel sets `fHidden` on
/// the placeholders, but it also sets it on an ordinary defined name the
/// user chose to hide from the Name Manager, and both this reader's
/// writer counterpart and Excel itself store the two the same way.
bool IsStoragePlaceholderName(std::string_view name) {
  constexpr std::string_view kPrefixes[] = {"_xlfn.", "_xlpm."};
  for (const std::string_view prefix : kPrefixes) {
    if (name.size() < prefix.size()) {
      continue;
    }
    bool match = true;
    for (std::size_t i = 0; i < prefix.size(); ++i) {
      const char lhs = static_cast<char>(std::tolower(static_cast<unsigned char>(name[i])));
      if (lhs != prefix[i]) {
        match = false;
        break;
      }
    }
    if (match) {
      return true;
    }
  }
  return false;
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

/// Decodes the supporting-book list from `xl/workbook.bin` into one
/// entry per book, in `BrtExternSheet`'s `iSupBook` order. The value is
/// `0` for this workbook and `N >= 1` for the N-th external book in
/// package order, which is the number the equivalent xlsx formula text
/// spells as `[N]`.
///
/// Layout verified against Excel-365-produced `xl/workbook.bin` files:
/// the books are the records between `BrtBeginExternals` and
/// `BrtEndExternals`, one record each, and the record id alone says
/// which kind it is. Only `BrtSupSelf` is this workbook; `BrtSupAddin`
/// and `BrtSupSame` are counted so later entries keep their index, but
/// neither names a resolvable external package, so both are reported as
/// external and therefore refused downstream.
///
/// A workbook can legitimately have no self entry: one whose only
/// qualified references are cross-workbook carries a single
/// `BrtSupBookSrc`, so `iSupBook == 0` is then an external book. That
/// case is why this list has to be read rather than assumed.
///
/// A `BrtSupBookSrc` also names the relationship its external link part
/// hangs off, which is how a `[N]` is turned back into a package part.
/// That id is retained on the entry; the other kinds leave it empty.
struct XlsbSupBook {
  std::uint32_t external_book = 0;
  std::string rel_id;
};

std::vector<XlsbSupBook> DecodeSupBooks(const std::vector<std::uint8_t>& body) {
  std::vector<XlsbSupBook> books;
  ByteSpan cursor{body.data(), body.size()};
  bool inside = false;
  std::uint32_t next_external = 1;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      break;
    }
    const XlsbRecord& rec = rec_or.value();
    const auto type = static_cast<XlsbRecordType>(rec.type);
    if (type == XlsbRecordType::BrtBeginExternals) {
      inside = true;
      continue;
    }
    if (type == XlsbRecordType::BrtEndExternals) {
      break;
    }
    if (!inside) {
      continue;
    }
    switch (type) {
      case XlsbRecordType::BrtSupSelf:
        books.push_back(XlsbSupBook{});
        break;
      case XlsbRecordType::BrtSupBookSrc: {
        ByteSpan payload = rec.payload;
        auto rid_or = read_xlwidestring(payload);
        books.push_back(XlsbSupBook{next_external, rid_or ? std::move(rid_or.value()) : std::string()});
        ++next_external;
        break;
      }
      case XlsbRecordType::BrtSupAddin:
      case XlsbRecordType::BrtSupSame:
        books.push_back(XlsbSupBook{next_external, std::string()});
        ++next_external;
        break;
      default:
        // `BrtSupTabs` and the future-record framing that decorates a
        // book entry sit inside the same block without adding a book.
        break;
    }
  }
  return books;
}

/// Decodes the `BrtExternSheet` table from `xl/workbook.bin`: resolves a
/// `PtgRef3d` / `PtgArea3d` `ixti` (0-based index into the returned
/// vector) to a `(itabFirst, itabLast)` sheet-index range. Byte layout
/// verified against a real Excel-365-produced `xl/workbook.bin`: u32
/// count, followed by `count` entries of `(iSupBook, itabFirst,
/// itabLast)` as 3 x i32 each. `iSupBook` is resolved through
/// `DecodeSupBooks` so `decode_ptgs` can tell an internal range from an
/// external-workbook one, whose sheet indices index the supporting book
/// rather than this workbook's `sheet_names`.
/// A workbook with no qualified references at all carries no
/// `BrtExternSheet` record, so an empty result is a normal outcome, not
/// an error.
Expected<std::vector<XlsbSheetRange>, Error> DecodeExternSheet(const std::vector<std::uint8_t>& body,
                                                               const std::vector<XlsbSupBook>& books) {
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
      // An `iSupBook` past the end of the list cannot be resolved. Treat
      // it as external: that refuses the reference, where assuming
      // "internal" would bind it to a local sheet by index.
      //
      // An *entirely* absent list is a different situation: the file
      // carries qualified references but no supporting-book block to
      // resolve them against, which no Excel-written file observed here
      // does. Refusing every reference would regress such a file from
      // working to undecodable, so it keeps the weaker rule that index
      // 0 is this workbook.
      const std::uint32_t sup_book = sup_book_or.value();
      std::uint32_t external_book = std::numeric_limits<std::uint32_t>::max();
      if (books.empty()) {
        external_book = sup_book == 0U ? 0U : sup_book;
      } else if (sup_book < books.size()) {
        external_book = books[sup_book].external_book;
      }
      ranges.push_back(XlsbSheetRange{static_cast<std::int32_t>(first_or.value()),
                                      static_cast<std::int32_t>(last_or.value()), external_book});
    }
    break;  // Exactly one BrtExternSheet record per workbook.
  }
  return ranges;
}

/// Registers every user-visible `BrtName` entry as a workbook defined
/// name via a single bulk `Workbook::set_defined_names` call (mirroring
/// the OOXML reader's `ooxml_reader.cpp` pattern -- a load-time
/// population pass, not the incremental single-name edit API
/// `set_defined_name_scoped` guards with dedup/dep-graph-rebuild logic
/// that only matters for post-load mutation). Only the storage
/// placeholders (the `_xlfn.*` / `_xlpm.*` future-function and
/// LET/LAMBDA-parameter names `PtgName` resolves during Ptg decode — see
/// `name_table`) are skipped; they are not user-visible names and never
/// carry a Name Manager comment. A name Excel merely hides from the Name
/// Manager keeps its `fHidden` bit on the produced `DefinedName` and
/// is registered like any other. Walks `xl/workbook.bin`'s `BrtName`
/// records a second time (after `DecodeWorkbookNames` has already built
/// the complete `name_table`), decoding each entry's own formula body so
/// a qualified or name-referencing formula (e.g. `Rate` defined as a
/// cell reference) resolves against the full table rather than a
/// partially-built one. A name whose formula uses a Ptg token outside
/// the supported set logs a structured warning and is skipped rather
/// than failing the whole read — the OOXML reader's `DecodeFormulaText`
/// contract mirrored at the workbook-name level.
Expected<void, Error> RegisterDefinedNames(const std::vector<std::uint8_t>& body,
                                           const std::vector<XlsbName>& name_table,
                                           const std::vector<std::string>& sheet_names,
                                           const std::vector<XlsbSheetRange>& sheet_ranges,
                                           const XlsbExternalBooks& external_books, Workbook& wb,
                                           std::uint32_t* undecoded_defined_name_count) {
  ByteSpan cursor{body.data(), body.size()};
  std::size_t name_index = 0;
  std::vector<DefinedName> out;
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
    if (IsStoragePlaceholderName(entry.name)) {
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
    p.data += cce;
    p.size -= cce;
    // `cb` + `rgcb` ([MS-XLSB] §2.4.649's `CellParsedFormula`): the
    // array-constant / mem-area extra data for this formula's own Ptg
    // tokens. Must be skipped correctly (not just assumed absent) to
    // land on the trailing comment string below.
    auto cb_or = read_u32(p);
    if (!cb_or) {
      return cb_or.error();
    }
    const std::uint32_t cb = cb_or.value();
    if (cb > p.size) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        "workbook.bin: BrtName formula rgcb length exceeds payload", "context=xlsb_reader");
    }
    ByteSpan rgcb{p.data, cb};
    p.data += cb;
    p.size -= cb;
    Arena arena(/*initial_chunk_bytes=*/4096, kMaxLoadArenaBytes);
    auto ast_or = decode_ptgs(rgce, rgcb, arena, sheet_names, name_table, sheet_ranges, external_books);
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
    // Trailing BrtName strings: a plain name carries exactly one, the
    // optional Name Manager comment, as a null `XLNullableWideString`
    // when unset (`read_xlnullablewidestring` maps that to an empty
    // string). This mirrors the writer's `EmitName`.
    auto comment_or = read_xlnullablewidestring(p);
    if (!comment_or) {
      return comment_or.error();
    }
    // Defined-name formulas store the bare expression text (no leading
    // `=`), matching the OOXML `<definedName>` element's text content.
    DefinedName dn;
    dn.name = entry.name;
    // The decoder names a hidden-name callee with its storage prefix; the
    // text reads back in formula-bar spelling, like a cell's.
    dn.formula = parser::spell_storage_operators(
        parser::strip_storage_prefixes(parser::format_formula(*ast_or.value()), &has_storage_prefix));
    dn.local_sheet_id = entry.itab;
    dn.hidden = entry.hidden;
    dn.comment = std::move(comment_or.value());
    out.push_back(std::move(dn));
  }
  wb.set_defined_names(std::move(out));
  return Expected<void, Error>::Ok();
}

/// Reads every entry of a worksheet's `.rels` part, keeping the original
/// relationship ids so the retained tail records still resolve. Internal
/// targets are normalised to package-relative paths (matching what the OOXML
/// reader stores); external targets stay verbatim.
Expected<std::vector<UnknownRelationship>, Error> LoadSheetRelationships(const ZipReader& zip,
                                                                         std::string_view rels_path,
                                                                         std::string_view sheet_dir) {
  std::vector<UnknownRelationship> out;
  auto status =
      ooxml::visit_relationship_nodes(zip, rels_path, "sheet rels", "xlsb_reader", [&](const pugi::xml_node& rel) {
        UnknownRelationship entry;
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

/// Returns the package-relative target of the first internal relationship
/// of `type` declared by `part_path`, or an empty string when the part has
/// no rels file or no such relationship.
Expected<std::string, Error> FindRelationshipTarget(const ZipReader& zip, std::string_view part_path,
                                                    std::string_view type) {
  const std::string rels_path = ooxml::rels_path_for_part(part_path);
  if (!zip.has_entry(rels_path)) {
    return std::string();
  }
  const std::string dir = ooxml::dir_of(part_path);
  std::string found;
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, "pivot rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (!found.empty()) {
          return Expected<void, Error>::Ok();
        }
        if (std::string_view(rel.attribute("Type").value()) != type) {
          return Expected<void, Error>::Ok();
        }
        if (std::string_view(rel.attribute("TargetMode").value()) == "External") {
          return Expected<void, Error>::Ok();
        }
        auto resolved = ResolveRelativePath(dir, rel.attribute("Target").value());
        if (!resolved) {
          return Expected<void, Error>(resolved.error());
        }
        found = std::move(resolved).value();
        return Expected<void, Error>::Ok();
      });
  if (!status) {
    return status.error();
  }
  return found;
}

/// Links a retained passthrough part's package path to the specific model
/// object `PassthroughPart::model_fingerprint` must track for it. Built
/// while pivot parts and styles load below; applied once, by the
/// passthrough capture sweep, to tag each `PassthroughPart` as it is
/// created (see `io/xlsb/retained_part_fingerprint.h`).
struct RetainedPartOriginTag {
  PassthroughPart::RetainedOrigin origin = PassthroughPart::RetainedOrigin::kNone;
  std::size_t sheet_index = 0;
  std::size_t pivot_index = 0;
  std::uint32_t cache_id = 0;
};

/// Package path -> `RetainedPartOriginTag`. The path index rides on the
/// `string -> uint32_t` hash map the pivot-cache loader already uses, and
/// the tags live beside it; a later `set` for the same path replaces the tag.
struct RetainedPartOrigins {
  std::unordered_map<std::string, std::uint32_t> index;
  std::vector<RetainedPartOriginTag> tags;

  void set(const std::string& path, const RetainedPartOriginTag& tag) {
    const auto [it, inserted] = index.emplace(path, static_cast<std::uint32_t>(tags.size()));
    if (inserted) {
      tags.push_back(tag);
    } else {
      tags[it->second] = tag;
    }
  }

  const RetainedPartOriginTag* find(const std::string& path) const {
    const auto it = index.find(path);
    return it == index.end() ? nullptr : &tags[it->second];
  }
};

/// Decodes the cache a pivot-table part binds to, returning the model id
/// it was registered under. Caches already loaded are reused so two
/// tables over one cache share it, as they do in the file.
///
/// The id is ours to assign: the workbook's own cache-id table is not
/// decoded, and nothing outside the model consults these ids -- the parts
/// round-trip through passthrough rather than being rewritten from the
/// model.
std::optional<std::uint32_t> LoadPivotCacheFor(const ZipReader& zip, Workbook& wb, std::string_view pivot_table_path,
                                               std::unordered_map<std::string, std::uint32_t>& loaded,
                                               RetainedPartOrigins& origins) {
  auto def_path_or = FindRelationshipTarget(zip, pivot_table_path, kRelPivotCacheDefinition);
  if (!def_path_or || def_path_or.value().empty()) {
    return std::nullopt;
  }
  const std::string def_path = std::move(def_path_or).value();
  const auto seen = loaded.find(def_path);
  if (seen != loaded.end()) {
    return seen->second;
  }
  auto rec_path_or = FindRelationshipTarget(zip, def_path, kRelPivotCacheRecords);
  if (!rec_path_or || rec_path_or.value().empty()) {
    return std::nullopt;
  }
  const std::string rec_path = std::move(rec_path_or).value();
  if (!zip.has_entry(def_path) || !zip.has_entry(rec_path)) {
    return std::nullopt;
  }
  auto def_bytes_or = zip.read_entry(def_path);
  if (!def_bytes_or) {
    return std::nullopt;
  }
  auto rec_bytes_or = zip.read_entry(rec_path);
  if (!rec_bytes_or) {
    return std::nullopt;
  }
  const std::vector<std::uint8_t>& def_bytes = def_bytes_or.value();
  const std::vector<std::uint8_t>& rec_bytes = rec_bytes_or.value();
  auto cache_or =
      read_pivot_cache_bin(ByteSpan{def_bytes.data(), def_bytes.size()}, ByteSpan{rec_bytes.data(), rec_bytes.size()});
  if (!cache_or) {
    return std::nullopt;
  }
  const auto cache_id = static_cast<std::uint32_t>(wb.pivot_caches().size());
  pivot::PivotCache cache = std::move(cache_or.value());
  cache.set_cache_id(cache_id);
  origins.set(def_path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheDefinition, 0, 0, cache_id});
  origins.set(rec_path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotCacheRecords, 0, 0, cache_id});
  wb.add_pivot_cache(std::make_unique<pivot::PivotCache>(std::move(cache)));
  loaded.emplace(def_path, cache_id);
  return cache_id;
}

/// Decodes every pivot table reachable from a sheet's relationships.
///
/// A pivot this reader cannot account for is skipped, not failed. Before
/// this existed no XLSB pivot was decoded at all, so refusing one leaves
/// that workbook loading exactly as it used to, whereas propagating the
/// error would turn a working open into a failure. Skipping is also why
/// the parts are left out of `consumed_parts`: the XLSB writer has no
/// pivot output, so a decoded-but-not-consumed part still round-trips
/// verbatim, and a skipped one is indistinguishable from before.
/// Returns the package-relative target of the relationship `rel_id`
/// declared by `part_path`, or an empty string when it is absent or
/// points outside the package.
Expected<std::string, Error> FindRelationshipById(const ZipReader& zip, std::string_view part_path,
                                                  std::string_view rel_id) {
  const std::string rels_path = ooxml::rels_path_for_part(part_path);
  if (rel_id.empty() || !zip.has_entry(rels_path)) {
    return std::string();
  }
  const std::string dir = ooxml::dir_of(part_path);
  std::string found;
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, "external link rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (!found.empty() || std::string_view(rel.attribute("Id").value()) != rel_id) {
          return Expected<void, Error>::Ok();
        }
        if (std::string_view(rel.attribute("TargetMode").value()) == "External") {
          return Expected<void, Error>::Ok();
        }
        auto resolved = ResolveRelativePath(dir, rel.attribute("Target").value());
        if (!resolved) {
          return Expected<void, Error>(resolved.error());
        }
        found = std::move(resolved).value();
        return Expected<void, Error>::Ok();
      });
  if (!status) {
    return status.error();
  }
  return found;
}

/// Reads the remote workbook URL recorded in an external link part's own
/// rels file. Empty when the part has no rels or no external-path
/// relationship, which is what an unresolvable link already looks like.
std::string ReadExternalLinkTarget(const ZipReader& zip, std::string_view part_path) {
  const std::string rels_path = ooxml::rels_path_for_part(part_path);
  if (!zip.has_entry(rels_path)) {
    return std::string();
  }
  std::string found;
  auto status = ooxml::visit_relationship_nodes(
      zip, rels_path, "external link rels", "xlsb_reader", [&](const pugi::xml_node& rel) -> Expected<void, Error> {
        if (found.empty() && std::string_view(rel.attribute("Type").value()) == kRelExternalLinkPath) {
          found = rel.attribute("Target").value();
        }
        return Expected<void, Error>::Ok();
      });
  (void)status;
  return found;
}

/// Loads every external link part the supporting-book list names, in
/// `[N]` order, so `external_links()[N - 1]` is the book a formula's
/// `[N]` prefix selects.
///
/// A part that is missing, unreadable or uses an encoding the decoder
/// has not measured contributes an entry with an empty cache rather than
/// none: the position has to be held or every later `[N]` would shift by
/// one and resolve against the wrong workbook. A reference into an empty
/// cache reads `#REF!`, which is the behaviour before this existed.
///
/// The parts stay out of `consumed_parts` for the same reason the pivot
/// parts do -- the XLSB writer has no external-link output, so they must
/// continue to round-trip verbatim through passthrough.
void LoadExternalLinkParts(const ZipReader& zip, Workbook& wb, const std::vector<XlsbSupBook>& books) {
  std::vector<ExternalLinkRecord> links;
  for (const XlsbSupBook& sup : books) {
    if (sup.external_book == 0) {
      continue;  // This workbook.
    }
    ExternalLinkRecord record;
    record.index = sup.external_book;
    record.rel_id = sup.rel_id;
    auto path_or = FindRelationshipById(zip, "xl/workbook.bin", sup.rel_id);
    if (path_or && !path_or.value().empty() && zip.has_entry(path_or.value())) {
      record.part_path = path_or.value();
      record.target = ReadExternalLinkTarget(zip, record.part_path);
      auto bytes_or = zip.read_entry(record.part_path);
      if (bytes_or) {
        const std::vector<std::uint8_t>& bytes = bytes_or.value();
        auto book_or = read_external_link_bin(ByteSpan{bytes.data(), bytes.size()});
        if (book_or) {
          record.kind = ExternalLinkRecord::Kind::kExternalBook;
          record.book = std::move(book_or.value());
        }
      }
    }
    links.push_back(std::move(record));
  }
  if (!links.empty()) {
    wb.set_external_links(std::move(links));
  }
}

/// Projects the loaded external links onto the two name tables
/// `decode_ptgs` consults. The projection exists so the Ptg decoder does
/// not have to know about `Workbook` or the link records at all.
XlsbExternalBooks CollectExternalBookTables(const Workbook& wb) {
  XlsbExternalBooks books;
  books.reserve(wb.external_links().size());
  for (const ExternalLinkRecord& link : wb.external_links()) {
    XlsbExternalBook entry;
    entry.sheet_names = link.book.sheet_names;
    entry.names.reserve(link.book.names.size());
    for (const ExternalBookName& name : link.book.names) {
      entry.names.push_back(name.name);
    }
    books.push_back(std::move(entry));
  }
  return books;
}

void LoadPivotParts(const ZipReader& zip, Workbook& wb, RetainedPartOrigins& origins) {
  std::unordered_map<std::string, std::uint32_t> loaded_caches;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    std::vector<std::string> pivot_table_paths;
    for (const UnknownRelationship& rel : wb.sheet(i).unknown_relationships()) {
      if (rel.type == kRelPivotTable && !rel.target_external && !rel.target.empty()) {
        pivot_table_paths.push_back(rel.target);
      }
    }
    for (const std::string& path : pivot_table_paths) {
      if (!zip.has_entry(path)) {
        continue;
      }
      auto bytes_or = zip.read_entry(path);
      if (!bytes_or) {
        continue;
      }
      const std::vector<std::uint8_t>& bytes = bytes_or.value();
      auto table_or = read_pivot_table_bin(ByteSpan{bytes.data(), bytes.size()});
      if (!table_or) {
        continue;
      }
      const std::optional<std::uint32_t> cache_id = LoadPivotCacheFor(zip, wb, path, loaded_caches, origins);
      if (!cache_id) {
        continue;
      }
      pivot::PivotTable table = std::move(table_or.value());
      table.set_pivot_cache_id(*cache_id);
      origins.set(path, RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kPivotTable, i,
                                              wb.sheet(i).pivot_tables().size(), 0});
      wb.sheet(i).add_pivot_table(std::make_unique<pivot::PivotTable>(std::move(table)));
    }
  }
  // Backfills each field's source-column name and each item's label from
  // the bound cache; without it GETPIVOTDATA cannot match a field by name.
  pivot::resolve_all_pivot_names(wb);
}

/// Resolves the relationship ids carried by model-owned BrtHLink records.
/// Only relationships actually consumed by a hyperlink are removed from the
/// sheet's unknown relationship list; unrelated drawing/table/custom rels,
/// including unused hyperlink rels, remain available to the writer.
Expected<void, Error> ResolveSheetHyperlinks(Sheet& sheet, std::vector<UnknownRelationship>& relationships) {
  std::unordered_set<std::string> consumed;
  for (Hyperlink& hyperlink : sheet.mutable_hyperlinks()) {
    if (hyperlink.rid.empty()) {
      // A present-but-empty RelID is the internal target form. Its location
      // string carries the destination and no relationship is consumed.
      continue;
    }
    const UnknownRelationship* matched = nullptr;
    for (const UnknownRelationship& relationship : relationships) {
      if (relationship.id != hyperlink.rid) {
        continue;
      }
      if (matched != nullptr) {
        return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb BrtHLink relationship id is duplicated",
                          "context=xlsb_reader rid=" + hyperlink.rid);
      }
      matched = &relationship;
    }
    if (matched == nullptr || matched->type != kRelHyperlink || !matched->target_external || matched->target.empty()) {
      return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                        matched == nullptr ? "xlsb BrtHLink relationship id is dangling"
                                           : "xlsb BrtHLink relationship has wrong type or target mode",
                        "context=xlsb_reader rid=" + hyperlink.rid);
    }
    hyperlink.target = matched->target;
    consumed.insert(hyperlink.rid);
  }
  if (!consumed.empty()) {
    relationships.erase(std::remove_if(relationships.begin(), relationships.end(),
                                       [&consumed](const UnknownRelationship& relationship) {
                                         return consumed.count(relationship.id) != 0U;
                                       }),
                        relationships.end());
  }
  return Expected<void, Error>::Ok();
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
  const std::vector<XlsbSupBook> sup_books = DecodeSupBooks(wb_bytes_or.value());
  auto sheet_ranges_or = DecodeExternSheet(wb_bytes_or.value(), sup_books);
  if (!sheet_ranges_or) {
    return sheet_ranges_or.error();
  }
  const std::vector<XlsbSheetRange>& sheet_ranges = sheet_ranges_or.value();

  // 5. Build the workbook bottom-up.
  Workbook wb = Workbook::create_empty();
  // The package's style table is authoritative, including when the package
  // has none: back-filling the factory's seeded defaults would invent style
  // records the file never carried and add an unwanted styles part on write.
  wb.set_styles(StylesTable{});
  wb.set_date1904(workbook_info.date1904);
  wb.set_workbook_protection_xml(workbook_info.protection_xml);
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
    // Same boundary validation the OOXML reader applies, and the same
    // error code: a duplicate sheet name makes every lookup resolve to the
    // first match, so the workbook would compute from the wrong sheet with
    // no ambiguity signal. Callers should not have to switch on the source
    // format to recognise that condition.
    auto added = wb.add_sheet_validated(b.name);
    if (!added) {
      if (added.error().code != FormulonErrorCode::kInvalidSheetName) {
        return added.error();
      }
      return make_error(FormulonErrorCode::kIoSheetCorrupt,
                        "workbook.bin: BrtBundleSh name is invalid or collides with an earlier sheet",
                        "context=xlsb_reader sheet=\"" + b.name + "\"");
    }
    if (b.visibility != SheetVisibility::kVisible) {
      wb.sheet(wb.sheet_count() - 1U).mutable_view().set_visibility(b.visibility);
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

  // Runs before both the defined names and the sheets, because a formula
  // in either can carry a `PtgNameX` that only the supporting workbook's
  // own name table can bind.
  LoadExternalLinkParts(zip, wb, sup_books);
  const XlsbExternalBooks external_books = CollectExternalBookTables(wb);

  // 5b. Register every non-hidden `BrtName` entry as a defined name.
  // Needs `sheet_names` (for qualified references inside a name's own
  // formula) and the complete `name_table` / `sheet_ranges` (for
  // self-consistent `PtgName` / `PtgRef3d` resolution), so this can
  // only run after both are fully built above.
  std::uint32_t undecoded_formula_count = 0;
  std::uint32_t undecoded_defined_name_count = 0;
  if (auto r = RegisterDefinedNames(wb_bytes_or.value(), name_table, sheet_names, sheet_ranges, external_books, wb,
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
  // Parallel to `sst_entries`, one (possibly empty) run list per index.
  std::vector<std::vector<PhoneticRun>> sst_phonetic;
  std::vector<PhoneticProperties> sst_phonetic_props;
  if (!wb_rels.sst_path.empty() && zip.has_entry(wb_rels.sst_path)) {
    auto sst_bytes_or = zip.read_entry(wb_rels.sst_path);
    if (!sst_bytes_or) {
      return sst_bytes_or.error();
    }
    auto sst_or = DecodeSharedStringsBin(sst_bytes_or.value(), text_storage, sst_phonetic, sst_phonetic_props);
    if (!sst_or) {
      return sst_or.error();
    }
    sst_entries = std::move(sst_or.value());
  }

  // Links each retained pivot / styles part's package path to the model
  // object its `PassthroughPart::model_fingerprint` must track, so step 8
  // below can tag the `PassthroughPart` it captures for that path.
  RetainedPartOrigins retained_part_origins;

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
    retained_part_origins.set(wb_rels.styles_path,
                              RetainedPartOriginTag{PassthroughPart::RetainedOrigin::kStyles, 0, 0, 0});
    wb.set_styles(std::move(styles_or.value()));
  }

  // 7. Each sheet binary.
  std::uint32_t cells_read = 0;
  std::uint32_t dropped_record_count = 0;
  std::unordered_set<std::string> consumed_parts;
  consumed_parts.insert("[Content_Types].xml");
  consumed_parts.insert("_rels/.rels");
  consumed_parts.insert(workbook_path);
  consumed_parts.insert(ooxml::rels_path_for_part(workbook_path));
  if (!wb_rels.sst_path.empty()) {
    consumed_parts.insert(wb_rels.sst_path);
  }

  // A formula is a dynamic-array formula exactly when a `BrtCellMeta`
  // naming this entry of `xl/metadata.bin` precedes it.
  std::uint32_t dynamic_array_ifmd = 0U;
  if (zip.has_entry("xl/metadata.bin")) {
    auto metadata_or = zip.read_entry("xl/metadata.bin");
    if (metadata_or) {
      dynamic_array_ifmd =
          find_dynamic_array_cell_meta_index(ByteSpan{metadata_or.value().data(), metadata_or.value().size()});
    }
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
    auto state_or =
        DecodeSheetBin(sheet_bytes_or.value(), i, wb, sst_entries, sst_phonetic, sst_phonetic_props, text_storage,
                       sheet_names, name_table, sheet_ranges, external_books, &undecoded_formula_count);
    if (!state_or) {
      return state_or.error();
    }
    cells_read += state_or.value().cells_decoded;
    dropped_record_count += state_or.value().dropped_records;
    std::unordered_set<std::uint64_t> dynamic_cells;
    for (const auto& [row, col, ifmd] : state_or.value().cell_metadata) {
      if (dynamic_array_ifmd != 0U && ifmd == dynamic_array_ifmd) {
        dynamic_cells.insert(dynamic_array_cell_key(row, col));
      }
    }
    apply_loaded_dynamic_array_marks(wb.sheet(i), dynamic_cells);
    consumed_parts.insert(sheet_path);

    // The sheet's own rels file resolves the relationship ids carried by the
    // retained tail records (hyperlink targets, drawing and table parts). It
    // is `rels`-Default-typed, so the Override-driven passthrough loop below
    // never sees it; keeping every entry — none of them are modelled by the
    // binary reader — preserves both the ids and their targets.
    const std::string sheet_rels_path = ooxml::rels_path_for_part(sheet_path);
    if (zip.has_entry(sheet_rels_path)) {
      auto rels_or = LoadSheetRelationships(zip, sheet_rels_path, ooxml::dir_of(sheet_path));
      if (!rels_or) {
        return rels_or.error();
      }
      auto relationships = std::move(rels_or.value());
      auto hyperlinks = ResolveSheetHyperlinks(wb.sheet(i), relationships);
      if (!hyperlinks) {
        return hyperlinks.error();
      }
      wb.sheet(i).set_unknown_relationships(std::move(relationships));
      consumed_parts.insert(sheet_rels_path);
    } else {
      std::vector<UnknownRelationship> no_relationships;
      auto hyperlinks = ResolveSheetHyperlinks(wb.sheet(i), no_relationships);
      if (!hyperlinks) {
        return hyperlinks.error();
      }
    }
  }

  // Pivot caches and tables, reached through the sheet relationships the
  // loop above stored. Runs after every sheet is decoded because a table
  // is attached to its sheet and its cache is workbook-level.
  LoadPivotParts(zip, wb, retained_part_origins);

  // 8. Passthrough parts. First capture every Override-listed part the
  // reader did not consume, retaining its explicit content type. A second
  // residual sweep resolves the remaining entries through the source
  // Default registry. Default-typed parts intentionally carry an empty
  // content_type; the registry itself travels on the Workbook and the
  // writer re-emits the matching Default entry.
  //
  // A part named in `retained_part_origins` (pivot table/cache, styles) is
  // tagged with the model object it depends on and fingerprinted against
  // the now-fully-loaded workbook, so `write_xlsb` can tell an unmodified
  // retained part from one whose model twin has since been mutated.
  auto apply_retained_origin = [&retained_part_origins, &wb](PassthroughPart& part) {
    const RetainedPartOriginTag* found = retained_part_origins.find(part.path);
    if (found == nullptr) {
      return;
    }
    const RetainedPartOriginTag& tag = *found;
    part.retained_origin = tag.origin;
    part.origin_sheet_index = tag.sheet_index;
    part.origin_pivot_index = tag.pivot_index;
    part.origin_cache_id = tag.cache_id;
    part.model_fingerprint = current_retained_part_fingerprint(wb, part);
  };

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
    apply_retained_origin(part);
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(part_name);
  }

  auto default_content_type_for = [&ct_view](std::string_view path) -> const DefaultContentType* {
    const std::string extension = ooxml::extension_of_part(path);
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
    apply_retained_origin(part);
    unknown_parts.push_back(std::move(part));
    captured_parts.insert(entry);
  }

  sort_passthrough_parts(unknown_parts);

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

  wb.apply_legacy_implicit_intersections();

  XlsbReadResult result{std::move(wb),      cells_read,          undecoded_formula_count, undecoded_defined_name_count,
                        dropped_part_count, dropped_record_count};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
