// MS-XLSB worksheet cell-record decoding: the common cell header, literal
// value records, and Ptg-stream-to-formula-text conversion for formula cells.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_ids.h"
#include "io/xlsb/sheet_reader.h"
#include "io/xlsb/sheet_reader_internal.h"
#include "parser/ast.h"
#include "parser/ast_format.h"
#include "phonetic.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "utils/structured_log.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// True when `ptgs` is the placeholder Excel stores in a cell whose real
/// tokens live in a separate record: a lone `PtgExp` (`0x01`) naming the
/// owning anchor. Both the anchor and the phantoms of an array or
/// dynamic-array formula carry it, and `BrtArrFmla` supplies the tokens
/// out-of-band, so the shell is the expected encoding rather than a Ptg
/// the decoder failed on.
///
/// It is only ever the whole stream — it takes no operands and cannot
/// combine with another token — so testing the first byte classifies
/// the record.
bool IsAnchorShellFormula(ByteSpan ptgs) {
  return ptgs.size != 0U && ptgs.data[0] == 0x01U;
}
}  // namespace

namespace detail {

/// Attempts to decode the `rgce` Ptg byte stream into an Excel formula
/// text (with a leading `=`). Returns the formula text on success, or an
/// empty string when the stream uses a token outside the supported set
/// (a structured-log diagnostic records the reason). On the empty-string
/// path the caller PRESERVES the cell's cached value and stores no
/// formula, so the cell still displays the correct cached result instead
/// of a fabricated formula that would recalc to `#NAME?`.
///
/// An anchor shell returns empty as well, but neither logs nor counts:
/// `undecoded_formula_count` states that the reader could not recover a
/// formula the source carried, and for a shell there is no formula at
/// that position to recover. The record that does carry the tokens is
/// decoded on its own. A shell whose owning record the reader does not
/// model instead surfaces through that record's own disposition, so the
/// loss is still reported — just not as a decode failure here.
std::string DecodeFormulaText(ByteSpan ptg_bytes, ByteSpan rgcb, const std::vector<std::string>& sheet_names,
                              const std::vector<XlsbName>& name_table, const std::vector<XlsbSheetRange>& sheet_ranges,
                              const XlsbExternalBooks& external_books, std::size_t sheet_index, std::uint32_t row,
                              std::uint32_t col, std::uint32_t* undecoded_formula_count,
                              const std::vector<TableMetadata>* tables) {
  if (IsAnchorShellFormula(ptg_bytes)) {
    return {};
  }
  Arena arena(kDefaultArenaChunkBytes, kMaxLoadArenaBytes);
  parser::FormulaWhitespace whitespace;
  auto ast_or = decode_ptgs(ptg_bytes, rgcb, arena, sheet_names, name_table, sheet_ranges, external_books,
                            static_cast<std::int32_t>(sheet_index), std::nullopt, &whitespace, tables);
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
  out.append(parser::format_formula(*ast_or.value(), &whitespace));
  return out;
}

/// Reads the eight-byte cell header common to every cell record:
///   * column   : u32 (zero-based)
///   * iStyleRef: u24 (index into `StylesTable::cell_xfs`)
///   * fPhShow  : u8  (Phonetic-text flag; bit 0)
///
/// Returns the column index and style-xf index, and advances the
/// cursor past the header.
Expected<CellHeaderInfo, Error> ReadCellHeader(ByteSpan& cursor) {
  ASSIGN_OR_RETURN(auto col, read_u32(cursor));
  if (col >= Sheet::kMaxCols) {
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
  const bool show_phonetic = (cursor.data[3] & 0x01U) != 0U;
  cursor.data += 4;
  cursor.size -= 4;
  return CellHeaderInfo{col, xf_index, show_phonetic};
}

}  // namespace detail

namespace {

/// Stores `xf_index` on `(sheet_index, row, col)` when it differs from
/// the default xf (`0`); a value record's default style is implicit.
Expected<void, Error> ApplyXfIndex(Workbook& wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                                   std::uint32_t xf_index) {
  if (xf_index == 0) {
    return Expected<void, Error>::Ok();
  }
  return wb.set_cell_xf_index(sheet_index, row, col, xf_index);
}
}  // namespace

namespace detail {

/// Stores the cell's xf and counts it as decoded. Value records skip the
/// default xf (`0`); blank and formula records store it explicitly.
Expected<void, Error> CommitCell(Workbook& wb, std::size_t sheet_index, SheetDecodeState& state,
                                 const CellHeaderInfo& header, bool explicit_xf) {
  if (explicit_xf) {
    RETURN_IF_ERROR(wb.set_cell_xf_index(sheet_index, state.current_row, header.col, header.xf_index));
  } else {
    RETURN_IF_ERROR(ApplyXfIndex(wb, sheet_index, state.current_row, header.col, header.xf_index));
  }
  if (header.show_phonetic) {
    wb.sheet(sheet_index).set_cell_show_phonetic(state.current_row, header.col, true);
  }
  ++state.cells_decoded;
  return Expected<void, Error>::Ok();
}

/// Slices a `CellParsedFormula` (u32 cce + rgce, then an optional u32 cb + rgcb
/// when at least four bytes remain); `what` names the record in error text.
Expected<void, Error> ReadCellParsedFormula(ByteSpan& p, const char* what, ByteSpan* rgce, ByteSpan* rgcb) {
  ASSIGN_OR_RETURN(const std::uint32_t cce, read_u32(p));
  if (cce > p.size) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                      std::string("xlsb ") + what + " rgce length exceeds payload", "context=xlsb_reader");
  }
  *rgce = ByteSpan{p.data, cce};
  p.data += cce;
  p.size -= cce;
  *rgcb = ByteSpan{};
  if (p.size >= 4) {
    ASSIGN_OR_RETURN(const std::uint32_t cb, read_u32(p));
    if (cb > p.size) {
      return make_error(FormulonErrorCode::kIoXlsbRecordTruncated,
                        std::string("xlsb ") + what + " rgcb length exceeds payload", "context=xlsb_reader");
    }
    *rgcb = ByteSpan{p.data, cb};
  }
  return Expected<void, Error>::Ok();
}

/// Decodes one literal cell record (blank, RK, real, bool, error, inline or
/// shared string): the cell header, the value, then the style commit. A
/// blank stores its xf explicitly; the value records skip the default one.
Expected<RecordDisposition, Error> DecodeLiteralCell(const XlsbRecord& rec, XlsbRecordType type,
                                                     SheetDecodeState& state, std::size_t sheet_index, Workbook& wb,
                                                     const std::vector<std::string_view>& sst_entries,
                                                     const std::vector<std::vector<PhoneticRun>>& sst_phonetic,
                                                     const std::vector<PhoneticProperties>& sst_phonetic_props,
                                                     std::deque<std::string>& text_storage) {
  if (!state.row_seen) {
    return RecordDisposition::kAccounted;
  }
  ByteSpan p = rec.payload;
  ASSIGN_OR_RETURN(auto col, ReadCellHeader(p));
  Value value = Value::blank();
  std::uint32_t sst_index = 0;
  switch (type) {
    case XlsbRecordType::BrtCellRk: {
      ASSIGN_OR_RETURN(auto rk, read_u32(p));
      value = Value::number(decode_rk_number(rk));
      break;
    }
    case XlsbRecordType::BrtCellReal: {
      if (p.size < 8) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb BrtCellReal payload truncated",
                          "context=xlsb_reader");
      }
      double v;
      std::memcpy(&v, p.data, sizeof(v));
      value = Value::number(v);
      break;
    }
    case XlsbRecordType::BrtCellBool: {
      ASSIGN_OR_RETURN(auto b, read_u8(p));
      value = Value::boolean(b != 0);
      break;
    }
    case XlsbRecordType::BrtCellError: {
      ASSIGN_OR_RETURN(auto code, read_u8(p));
      // Map the OOXML wire code to `ErrorCode` via the single
      // `kErrorTable`-backed lookup shared with the writer, so this
      // path stays symmetric with `ooxml_code()` (see
      // `error_from_ooxml_code` in `value.h`).
      value = Value::error(error_from_ooxml_code(static_cast<std::int32_t>(code)));
      break;
    }
    case XlsbRecordType::BrtCellSt: {
      ASSIGN_OR_RETURN(auto str, read_xlwidestring(p));
      text_storage.push_back(std::move(str));
      value = Value::text(text_storage.back());
      break;
    }
    case XlsbRecordType::BrtCellIsst: {
      ASSIGN_OR_RETURN(sst_index, read_u32(p));
      if (sst_index >= sst_entries.size()) {
        std::string ctx("context=xlsb_reader sheet_index=");
        ctx.append(std::to_string(sheet_index));
        ctx.append(" row=").append(std::to_string(state.current_row));
        ctx.append(" col=").append(std::to_string(col.col));
        ctx.append(" sst_index=").append(std::to_string(sst_index));
        ctx.append(" sst_size=").append(std::to_string(sst_entries.size()));
        return make_error(FormulonErrorCode::kIoXlsbCorrupt, "xlsb sst index out of range", std::move(ctx));
      }
      value = Value::text(sst_entries[sst_index]);
      break;
    }
    default:
      break;
  }
  Sheet& sheet = wb.sheet(sheet_index);
  sheet.set_cell_cached_value_borrowed(state.current_row, col.col, std::move(value));
  // Attached after the value, mirroring the OOXML reader: every
  // value-mutating setter clears the annotation, so the order is
  // load-bearing. Skipped when the entry carries no guide so an
  // unannotated cell keeps its default-constructed run vector.
  if (type == XlsbRecordType::BrtCellIsst && sst_index < sst_phonetic.size() && !sst_phonetic[sst_index].empty()) {
    sheet.set_cell_phonetic_runs(state.current_row, col.col, sst_phonetic[sst_index]);
    if (sst_index < sst_phonetic_props.size()) {
      sheet.set_cell_phonetic_props(state.current_row, col.col, sst_phonetic_props[sst_index]);
    }
  }
  RETURN_IF_ERROR(CommitCell(wb, sheet_index, state, col, /*explicit_xf=*/type == XlsbRecordType::BrtCellBlank));
  return RecordDisposition::kModelled;
}

}  // namespace detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
