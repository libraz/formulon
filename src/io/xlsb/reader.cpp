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

#include "io/passthrough_part.h"
#include "io/xlsb/record.h"
#include "io/zip_reader.h"
#include "pugixml.hpp"
#include "sheet.h"
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
/// `.` / `..` segments. Mirror of the OOXML reader's helper, restated
/// here so we don't drag the OOXML reader header into this TU.
///
/// Path-traversal hardening matches the OOXML variant: a `Target` with
/// excess `..` segments or a leading `/` (package-absolute) is refused
/// with `kIoZipSlip`.
Expected<std::string, Error> ResolveRelativePath(std::string_view base_dir, std::string_view target) {
  if (!target.empty() && target.front() == '/') {
    std::string ctx("context=xlsb_reader base_dir=");
    ctx.append(base_dir);
    ctx.append(" target=");
    ctx.append(target);
    return make_error(FormulonErrorCode::kIoZipSlip, "rels target uses package-absolute path; refusing to resolve",
                      std::move(ctx));
  }
  std::vector<std::string> stack;
  std::size_t start = 0;
  for (std::size_t i = 0; i <= base_dir.size(); ++i) {
    if (i == base_dir.size() || base_dir[i] == '/') {
      if (i > start) {
        stack.emplace_back(base_dir.substr(start, i - start));
      }
      start = i + 1;
    }
  }
  start = 0;
  for (std::size_t i = 0; i <= target.size(); ++i) {
    if (i == target.size() || target[i] == '/') {
      if (i > start) {
        std::string_view seg = target.substr(start, i - start);
        if (seg == ".") {
          // skip
        } else if (seg == "..") {
          if (stack.empty()) {
            std::string ctx("context=xlsb_reader base_dir=");
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
    std::string ctx("context=xlsb_reader base_dir=");
    ctx.append(base_dir);
    ctx.append(" target=");
    ctx.append(target);
    return make_error(FormulonErrorCode::kIoZipSlip, "rels target resolves to empty path", std::move(ctx));
  }
  return out;
}

std::string DirOf(std::string_view path) {
  const std::size_t pos = path.find_last_of('/');
  if (pos == std::string_view::npos) {
    return {};
  }
  return std::string(path.substr(0, pos));
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
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: part not found",
                      "context=xlsb_reader rels_path=" + rels_path);
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
    std::string ctx("context=xlsb_reader part=");
    ctx.append(rels_path);
    ctx.append(" desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "workbook rels: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("Relationships");
  if (!root) {
    return make_error(FormulonErrorCode::kIoRelationshipBroken, "workbook rels: missing <Relationships>",
                      "context=xlsb_reader part=" + rels_path);
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
    auto skip = read_u32(p);
    if (!skip) {
      return skip.error();
    }
    auto skip2 = read_u32(p);
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

/// Best-effort hex-stub for un-decoded formula bytes. Bundle 4.2
/// replaces this with a real `Ptg → Excel-formula-text` round-trip.
/// The leading `=` keeps callers' "treat as formula" predicates
/// (`!formula_text.empty()` and friends) consistent with the OOXML
/// path; the body is a dotted hex sequence so the writer can later
/// recognise stub formulas if it needs to.
std::string FormulaStubFromBytes(ByteSpan ptg_bytes, std::size_t sheet_index, std::uint32_t row, std::uint32_t col) {
  StructuredLog("xlsb.formula.not_decoded")
      .field("sheet_index", static_cast<std::int64_t>(sheet_index))
      .field("row", static_cast<std::int64_t>(row))
      .field("col", static_cast<std::int64_t>(col))
      .field("ptg_bytes", static_cast<std::int64_t>(ptg_bytes.size))
      .warn();
  std::string out;
  // `ptg_bytes.size` is bounded by the enclosing record's payload size
  // (currently <= 64 KiB on the wire); cap defensively so a future record
  // size growth never lets `size * 3 + 16` overflow `std::size_t` on a
  // 32-bit WASM build. The cap only constrains the reservation hint -
  // the per-byte loop below still emits every input byte verbatim.
  constexpr std::size_t kMaxPtgBytesForStub = 65535U;
  const std::size_t bounded = std::min(ptg_bytes.size, kMaxPtgBytesForStub);
  out.reserve(bounded * 3 + 16);
  out.append("=__FORMULON_XLSB_PTG__(");
  static constexpr char kHex[] = "0123456789ABCDEF";
  for (std::size_t i = 0; i < ptg_bytes.size; ++i) {
    if (i > 0) {
      out.push_back('.');
    }
    out.push_back(kHex[(ptg_bytes.data[i] >> 4) & 0xF]);
    out.push_back(kHex[ptg_bytes.data[i] & 0xF]);
  }
  out.push_back(')');
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
                                                 std::deque<std::string>& text_storage) {
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
        // Skip the result payload according to record kind.
        switch (type) {
          case XlsbRecordType::BrtFmlaNum: {
            if (p.size < 8) {
              return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtFmlaNum value truncated",
                                "context=xlsb_reader");
            }
            double v;
            std::memcpy(&v, p.data, sizeof(v));
            (void)v;
            p.data += 8;
            p.size -= 8;
            break;
          }
          case XlsbRecordType::BrtFmlaString: {
            auto s = read_xlwidestring(p);
            if (!s) {
              return s.error();
            }
            (void)s.value();
            break;
          }
          case XlsbRecordType::BrtFmlaBool:
          case XlsbRecordType::BrtFmlaError: {
            auto b = read_u8(p);
            if (!b) {
              return b.error();
            }
            (void)b.value();
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
        // The remainder is the CellParsedFormula payload — captured as
        // opaque bytes for the Bundle 4.1 stub. Bundle 4.2 will decode
        // this via `io::xlsb::ptg`.
        const std::string stub = FormulaStubFromBytes(p, sheet_index, state.current_row, col_or.value());
        // The OOXML reader uses the workbook-level set_cell_formula
        // entry to register the cell with the dep graph. For the XLSB
        // skeleton we don't yet have a real formula text to parse, so
        // we route through Sheet::set_cell_formula directly to avoid
        // burdening the dep graph with unparseable stub strings.
        wb.sheet(sheet_index).set_cell_formula(state.current_row, col_or.value(), stub);
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
    sheet_part_paths.push_back(it->second);
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
  // workbook rels file path:
  {
    const std::size_t slash = workbook_path.find_last_of('/');
    std::string p;
    if (slash == std::string::npos) {
      p.append("_rels/").append(workbook_path).append(".rels");
    } else {
      p.append(workbook_path.substr(0, slash));
      p.append("/_rels/");
      p.append(workbook_path.substr(slash + 1));
      p.append(".rels");
    }
    consumed_parts.insert(std::move(p));
  }
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
    auto state_or = DecodeSheetBin(sheet_bytes_or.value(), i, wb, sst_entries, text_storage);
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
