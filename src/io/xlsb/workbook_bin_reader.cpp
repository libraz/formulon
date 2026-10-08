//
// Implementation of the `xl/workbook.bin` global-record decoders. See
// `io/xlsb/workbook_bin_reader.h`.

#include "io/xlsb/workbook_bin_reader.h"

#include <algorithm>
#include <cctype>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/xlsb/protection_records.h"
#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes `xl/workbook.bin` to extract the ordered sheet-bundle list and
/// workbook date system and protection. Other records are skipped.
Expected<WorkbookBinInfo, Error> DecodeWorkbookBin(const std::vector<std::uint8_t>& body) {
  WorkbookBinInfo info;
  ByteSpan protection_iso{};
  ByteSpan cursor{body.data(), body.size()};
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type == static_cast<std::uint16_t>(XlsbRecordType::BrtWbProp)) {
      // BrtWbProp ([MS-XLSB] §2.4.866) begins with a u32 grbit; bit 0
      // is f1904. The following theme-version and optional code-name
      // fields are irrelevant to the workbook model.
      ByteSpan p = rec.payload;
      ASSIGN_OR_RETURN(auto flags, read_u32(p));
      info.date1904 = (flags & 0x00000001U) != 0U;
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
    ASSIGN_OR_RETURN(auto hs_state, read_u32(p));
    auto skip2 = read_u32(p);  // iTabID
    if (!skip2) {
      return std::move(skip2.error());
    }
    ASSIGN_OR_RETURN(auto rid, read_xlnullablewidestring(p));
    auto name_or = read_xlwidestring(p);
    if (!name_or) {
      return std::move(name_or.error());
    }
    SheetBundleEntry entry;
    entry.rid = std::move(rid);
    entry.name = std::move(name_or.value());
    // An hsState outside the three defined values is not a visibility this
    // model can name; treat anything non-zero it cannot place as plain
    // hidden, which is the conservative direction (the sheet stays out of
    // sight rather than appearing unbidden).
    switch (hs_state) {
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

namespace {

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
/// the record is a `_xlfn.<FN>` future-function, `_xlpm.<param>`
/// LET / LAMBDA-parameter or `_xleta.<FN>` function-value placeholder
/// rather than a user-visible defined
/// name. Matched case-insensitively, the same way `ptg_reader.cpp`
/// resolves these names during Ptg decode.
///
/// This is deliberately NOT the `fHidden` bit: Excel sets `fHidden` on
/// the placeholders, but it also sets it on an ordinary defined name the
/// user chose to hide from the Name Manager, and both this reader's
/// writer counterpart and Excel itself store the two the same way.
bool IsStoragePlaceholderName(std::string_view name) {
  constexpr std::string_view kPrefixes[] = {"_xlfn.", "_xlpm.", "_xleta."};
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

}  // namespace

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
      return std::move(rec_or.error());
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
      return std::move(flags_or.error());
    }
    p.data += 3;  // 3 reserved bytes between flags and itab.
    p.size -= 3;
    ASSIGN_OR_RETURN(auto itab, read_u32(p));
    ASSIGN_OR_RETURN(auto cch, read_u32(p));
    std::string name;
    if (!ReadFixedWideString(p, cch, name)) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb fixed-length wide string truncated",
                        "context=xlsb_reader");
    }
    XlsbName entry;
    entry.itab = static_cast<std::int32_t>(itab);
    entry.name = std::move(name);
    entry.hidden = (flags_or.value() & 0x0001U) != 0;
    names.push_back(std::move(entry));
  }
  return names;
}

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
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != static_cast<std::uint16_t>(XlsbRecordType::BrtExternSheet)) {
      continue;
    }
    ByteSpan p = rec.payload;
    auto count_or = read_u32(p);
    if (!count_or) {
      return std::move(count_or.error());
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
      ASSIGN_OR_RETURN(const std::uint32_t sup_book, read_u32(p));
      ASSIGN_OR_RETURN(auto first, read_u32(p));
      ASSIGN_OR_RETURN(auto last, read_u32(p));
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
      std::uint32_t external_book = std::numeric_limits<std::uint32_t>::max();
      if (books.empty()) {
        external_book = sup_book == 0U ? 0U : sup_book;
      } else if (sup_book < books.size()) {
        external_book = books[sup_book].external_book;
      }
      ranges.push_back(
          XlsbSheetRange{static_cast<std::int32_t>(first), static_cast<std::int32_t>(last), external_book});
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
      return std::move(rec_or.error());
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
    ASSIGN_OR_RETURN(auto cch, read_u32(p));
    const std::size_t name_bytes = static_cast<std::size_t>(cch) * 2;
    if (name_bytes > p.size) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        "workbook.bin: BrtName name truncated (defined-name pass)", "context=xlsb_reader");
    }
    p.data += name_bytes;
    p.size -= name_bytes;
    ASSIGN_OR_RETURN(const std::uint32_t cce, read_u32(p));
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
    ASSIGN_OR_RETURN(const std::uint32_t cb, read_u32(p));
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
    ASSIGN_OR_RETURN(auto comment, read_xlnullablewidestring(p));
    // Defined-name formulas store the bare expression text (no leading
    // `=`), matching the OOXML `<definedName>` element's text content.
    DefinedName dn;
    dn.name = entry.name;
    // The decoder names a hidden-name callee with its storage prefix and a
    // book by its link index; the text reads back in formula-bar spelling,
    // like a cell's.
    dn.formula = wb.ingest_stored_formula(parser::format_formula(*ast_or.value()));
    dn.local_sheet_id = entry.itab;
    dn.hidden = entry.hidden;
    dn.comment = std::move(comment);
    out.push_back(std::move(dn));
  }
  wb.set_defined_names(std::move(out));
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
