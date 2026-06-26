// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the XLSB read pipeline skeleton. See
// `io/xlsb/reader.h` for the design context.
//
// The reader reuses the OOXML zip envelope plus three XML parts
// (`[Content_Types].xml`, `_rels/.rels`, `xl/_rels/workbook.xml.rels`).
// The binary parts (`xl/workbook.bin`, `xl/worksheets/sheet*.bin`,
// `xl/sharedStrings.bin`) are decoded via the record framing in
// `io/xlsb/record.h`. Formulas are not yet AST-decoded; we emit a stub
// formula text plus a structured-log warning so callers can see the
// gap. Bundle 4.2 lands the real decoder.

#include "io/xlsb/reader.h"

#include <algorithm>
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

#include "io/ooxml/package_validator.h"
#include "io/passthrough_part.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
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
    if (std::string_view(node.name()) != "Override") {
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

struct WorkbookRels {
  std::unordered_map<std::string, std::string> sheet_targets;  // rId -> resolved part path
  std::string sst_path;
  std::string styles_path;
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

/// Decodes `xl/workbook.bin` to extract the ordered sheet-bundle list.
/// Records other than `BrtBundleSh` are skipped — Bundle 4.1 only
/// needs the (name, rId) tuples to build sheets in document order.
Expected<std::vector<SheetBundleEntry>, Error> DecodeWorkbookBin(const std::vector<std::uint8_t>& body) {
  std::vector<SheetBundleEntry> entries;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
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
    entries.push_back(std::move(entry));
  }
  if (entries.empty()) {
    return make_error(FormulonErrorCode::kIoXlsbCorrupt, "workbook.bin: no BrtBundleSh records",
                      "context=xlsb_reader part=xl/workbook.bin");
  }
  return entries;
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
    //               set, but we ignore those for the skeleton).
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
std::string DecodeFormulaText(ByteSpan ptg_bytes, const std::vector<std::string>& sheet_names, std::size_t sheet_index,
                              std::uint32_t row, std::uint32_t col) {
  Arena arena;
  auto ast_or = decode_ptgs(ptg_bytes, arena, sheet_names);
  if (!ast_or) {
    StructuredLog("xlsb.formula.not_decoded")
        .field("sheet_index", static_cast<std::int64_t>(sheet_index))
        .field("row", static_cast<std::int64_t>(row))
        .field("col", static_cast<std::int64_t>(col))
        .field("ptg_bytes", static_cast<std::int64_t>(ptg_bytes.size))
        .field("reason", ast_or.error().message)
        .warn();
    return {};
  }
  std::string out("=");
  out.append(parser::format_formula(*ast_or.value()));
  return out;
}

/// Per-sheet decode state. The reader walks records in order and
/// updates `current_row` whenever it sees a `BrtRowHdr`. Cell records
/// then resolve `(current_row, col)` to the absolute cell address.
struct SheetDecodeState {
  std::uint32_t current_row = 0;
  bool row_seen = false;
  std::uint32_t cells_decoded = 0;
};

/// Reads the eight-byte cell header common to every cell record:
///   * column   : u32 (zero-based)
///   * iStyleRef: u24 (we ignore the style for now — Bundle 4.x will
///                consume it)
///   * fPhShow  : u8  (Phonetic-text flag; ignored)
///
/// Returns the column index and advances the cursor past the header.
Expected<std::uint32_t, Error> ReadCellHeader(ByteSpan& cursor) {
  auto col_or = read_u32(cursor);
  if (!col_or) {
    return col_or.error();
  }
  // iStyleRef (3 bytes) + fPhShow (1 byte) = 4 bytes total.
  if (cursor.size < 4) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb cell header truncated (style/phonetic)",
                      "context=xlsb_reader");
  }
  cursor.data += 4;
  cursor.size -= 4;
  return col_or.value();
}

/// Decodes one sheet binary part. Cells (literal + formula) flow into
/// `wb.sheet(sheet_index)`. SST indices are resolved against
/// `sst_entries`; out-of-range indices are returned as
/// `kIoXlsbCorrupt`.
Expected<SheetDecodeState, Error> DecodeSheetBin(const std::vector<std::uint8_t>& body, std::size_t sheet_index,
                                                 Workbook& wb, const std::vector<std::string_view>& sst_entries,
                                                 std::deque<std::string>& text_storage,
                                                 const std::vector<std::string>& sheet_names) {
  SheetDecodeState state;
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    const auto type = static_cast<XlsbRecordType>(rec.type);
    switch (type) {
      case XlsbRecordType::BrtRowHdr: {
        // BrtRowHdr ([MS-XLSB] §2.4.660): first u32 is the row index.
        ByteSpan p = rec.payload;
        auto row_or = read_u32(p);
        if (!row_or) {
          return row_or.error();
        }
        state.current_row = row_or.value();
        state.row_seen = true;
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
        wb.sheet(sheet_index).set_cell_cached_value(state.current_row, col_or.value(), Value::blank());
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
        wb.sheet(sheet_index).set_cell_cached_value(state.current_row, col_or.value(), Value::number(v));
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
        wb.sheet(sheet_index).set_cell_cached_value(state.current_row, col_or.value(), Value::number(v));
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
            .set_cell_cached_value(state.current_row, col_or.value(), Value::boolean(b_or.value() != 0));
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
        // Map OOXML wire codes to ErrorCode. Anything outside the
        // documented set surfaces as Unknown.
        ErrorCode ec = ErrorCode::Unknown;
        switch (code_or.value()) {
          case 0:
            ec = ErrorCode::Null;
            break;
          case 7:
            ec = ErrorCode::Div0;
            break;
          case 15:
            ec = ErrorCode::Value;
            break;
          case 23:
            ec = ErrorCode::Ref;
            break;
          case 29:
            ec = ErrorCode::Name;
            break;
          case 36:
            ec = ErrorCode::Num;
            break;
          case 42:
            ec = ErrorCode::NA;
            break;
          case 43:
            ec = ErrorCode::GettingData;
            break;
          default:
            ec = ErrorCode::Unknown;
            break;
        }
        wb.sheet(sheet_index).set_cell_cached_value(state.current_row, col_or.value(), Value::error(ec));
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
            .set_cell_cached_value(state.current_row, col_or.value(), Value::text(text_storage.back()));
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
          ctx.append(" col=").append(std::to_string(col_or.value()));
          ctx.append(" sst_index=").append(std::to_string(idx_or.value()));
          ctx.append(" sst_size=").append(std::to_string(sst_entries.size()));
          return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb sst index out of range", std::move(ctx));
        }
        wb.sheet(sheet_index)
            .set_cell_cached_value(state.current_row, col_or.value(), Value::text(sst_entries[idx_or.value()]));
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
            // Reuse the same wire-code mapping the literal path uses.
            ErrorCode ec = ErrorCode::Unknown;
            switch (b.value()) {
              case 0:
                ec = ErrorCode::Null;
                break;
              case 7:
                ec = ErrorCode::Div0;
                break;
              case 15:
                ec = ErrorCode::Value;
                break;
              case 23:
                ec = ErrorCode::Ref;
                break;
              case 29:
                ec = ErrorCode::Name;
                break;
              case 36:
                ec = ErrorCode::Num;
                break;
              case 42:
                ec = ErrorCode::NA;
                break;
              case 43:
                ec = ErrorCode::GettingData;
                break;
              default:
                ec = ErrorCode::Unknown;
                break;
            }
            cached = Value::error(ec);
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
        // Ptg stream + u32 cb + cb bytes of rgcb. Slice exactly the rgce
        // span and hand it to the decoder.
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
        const std::string formula_text =
            DecodeFormulaText(rgce, sheet_names, sheet_index, state.current_row, col_or.value());
        if (!formula_text.empty()) {
          // Register the real formula via the workbook-level entry so the
          // dep graph tracks it (matching the OOXML reader). The cached
          // value is preserved separately below.
          auto wf = wb.set_cell_formula(sheet_index, state.current_row, col_or.value(), formula_text);
          if (!wf) {
            return wf.error();
          }
        }
        // Always preserve the cached value (for undecodable formulas this
        // is the only correct datum the cell carries; for decoded ones it
        // matches Excel's stored result until the next recalc).
        if (!cached.is_blank()) {
          wb.sheet(sheet_index).set_cell_cached_value(state.current_row, col_or.value(), cached);
        }
        ++state.cells_decoded;
        break;
      }
      default:
        break;
    }
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
  const ContentTypesView& ct_view = ct_view_or.value();

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
  const std::vector<SheetBundleEntry>& bundle = bundle_or.value();

  // 5. Build the workbook bottom-up.
  Workbook wb = Workbook::create_empty();
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
  if (!wb_rels.styles_path.empty()) {
    consumed_parts.insert(wb_rels.styles_path);
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
    auto state_or = DecodeSheetBin(sheet_bytes_or.value(), i, wb, sst_entries, text_storage, sheet_names);
    if (!state_or) {
      return state_or.error();
    }
    cells_read += state_or.value().cells_decoded;
    consumed_parts.insert(sheet_path);
  }

  // 8. Passthrough parts: every Override-listed part the reader did
  // not consume, captured raw so a future writer can re-emit it
  // verbatim. Default-typed binary parts are out of scope per the
  // OOXML reader's contract.
  std::vector<PassthroughPart> unknown_parts;
  unknown_parts.reserve(ct_view.overrides.size());
  for (const auto& [part_name, content_type] : ct_view.overrides) {
    if (consumed_parts.find(part_name) != consumed_parts.end()) {
      continue;
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
  }
  std::sort(unknown_parts.begin(), unknown_parts.end(),
            [](const PassthroughPart& a, const PassthroughPart& b) { return a.path < b.path; });
  wb.set_passthrough_parts(unknown_parts);

  XlsbReadResult result{std::move(wb), std::move(unknown_parts), cells_read};
  return result;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
