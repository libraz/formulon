//
// Implementation of the per-cell XLSB record dispatcher. See
// `io/xlsb/cell_writer.h` for the contract.

#include "io/xlsb/cell_writer.h"

#include <cmath>
#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "io/stored_cell_error.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sheet_record_flags.h"
#include "io/xlsb/sst_writer.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "phonetic.h"
#include "utils/arena.h"
#include "utils/formula_text.h"
#include "utils/structured_log.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// The style word handed to `EmitCellHeader` carries `fPhShow` in bit 24,
/// above the 24-bit style index.
constexpr std::uint32_t kShowPhoneticShift = 24U;

std::uint32_t StyleWord(const Cell& cell, std::uint32_t xf_index) {
  // Excel keeps fPhShow on constants and blanks only.
  const bool show_phonetic = cell.show_phonetic && cell.formula_text.empty();
  return (xf_index & 0xFFFFFFU) | (show_phonetic ? (1U << kShowPhoneticShift) : 0U);
}

/// Mask of the style word `EmitCellHeader` writes: the 24-bit style index and
/// the fPhShow flag in bit 24.
constexpr std::uint32_t kStyleWordMask = (1U << (kShowPhoneticShift + 1U)) - 1U;

/// Emits the standard XLSB cell-header (column, iStyleRef, fPhShow):
///   * column    : u32
///   * iStyleRef : 3 bytes, little-endian (index into the xf table)
///   * fPhShow   : u8 (bit 0 = the style word's bit 24)
///
/// `xf_index` is the cell's style word; it must round-trip so a styled cell
/// keeps its formatting through `write_xlsb -> read_xlsb`. Bits above the
/// flag are dropped so they cannot reach the fourth byte's reserved bits.
/// `ReadCellHeader` decodes the same layout.
void EmitCellHeader(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index) {
  emit_u32(dst, col);
  emit_u32(dst, xf_index & kStyleWordMask);
}

/// Returns the `BErr` byte a cell value of `e` is stored as (see
/// `io::stored_cell_error`). Mirrors the inverse mapping `read_xlsb`
/// performs in `BrtCellError`.
std::uint8_t ErrorWireCode(ErrorCode e) {
  return static_cast<std::uint8_t>(ooxml_code(stored_cell_error(e)));
}

/// How a cached scalar is stored: a number, or a single byte (boolean value
/// or error wire code). A non-finite number is stored as `#NUM!`.
struct StoredScalar {
  enum class Shape { kNumber, kBool, kError, kOther };
  Shape shape = Shape::kOther;
  double number = 0.0;
  std::uint8_t byte = 0;
};

StoredScalar StoreScalar(const Value& v) {
  StoredScalar out;
  switch (v.kind()) {
    case ValueKind::Number:
      if (std::isfinite(v.as_number())) {
        out.shape = StoredScalar::Shape::kNumber;
        out.number = v.as_number();
      } else {
        out.shape = StoredScalar::Shape::kError;
        out.byte = ErrorWireCode(ErrorCode::Num);
      }
      break;
    case ValueKind::Bool:
      out.shape = StoredScalar::Shape::kBool;
      out.byte = v.as_boolean() ? 1U : 0U;
      break;
    case ValueKind::Error:
      out.shape = StoredScalar::Shape::kError;
      out.byte = ErrorWireCode(v.as_error());
      break;
    default:
      break;
  }
  return out;
}

/// Parses `cell.formula_text` into an AST and encodes it as a Ptg
/// (`rgce`) byte stream. The formula text starts with `=`; the parser
/// consumes the body after it. Returns the encoded bytes, or an Error
/// when the formula cannot be parsed or lowered to the supported Ptg
/// token set. The caller surfaces that error through `write_xlsb` rather
/// than silently dropping the formula. `array_formula` marks the tokens of
/// an array block's `BrtArrFmla`. The body's whitespace is stored where
/// Excel stores it (see `encode_ptgs`).
Expected<EncodedFormula, Error> EncodeCellFormula(const Cell& cell, const std::vector<std::string>& sheet_names,
                                                  const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                                  bool array_formula, const NameShapes& name_shapes) {
  const std::string_view body = strip_formula_prefix(cell.formula_text);
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb writer: formula failed to parse for Ptg encoding",
                      std::string("context=xlsb_cell_writer formula=") + cell.formula_text);
  }
  // Cell formulas: a bare reference/range as the whole formula promotes.
  return encode_ptgs(*root, sheet_names, sheet_ranges, name_table, PtgRootClass::kValue, std::nullopt,
                     cell.dynamic_array ? PtgEvaluation::kDynamicArray
                     : array_formula    ? PtgEvaluation::kLegacyArray
                                        : PtgEvaluation::kLegacy,
                     name_shapes, body);
}

/// Emits a `BrtFmla*` record matching `cached`'s kind, with the encoded
/// `rgce` / `rgcb` bytes spliced into the `CellParsedFormula` payload.
/// Layout per [MS-XLSB] §2.4.x:
///
///   cell-header (8 bytes)
///   value       (kind-specific)
///   grbitFlags  (u16: fAlwaysCalc 0x0002, as Excel sets it where .xlsx has `ca="1"`)
///   cce         (u32 byte length of rgce)
///   rgce        (cce bytes)         — the Ptg stream
///   cb          (u32 byte length of rgcb)
///   rgcb        (cb bytes)          — array-constant extra data
void EmitFormulaCellRecord(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, std::uint32_t col,
                           std::uint32_t xf_index, const Value& cached, const EncodedFormula& formula,
                           bool always_calculates) {
  scratch.clear();
  std::vector<std::uint8_t>& p = scratch;
  EmitCellHeader(p, col, xf_index);

  // A number, and the stand-in for a cached value that is none of the four scalar kinds.
  XlsbRecordType type = XlsbRecordType::BrtFmlaNum;
  const StoredScalar scalar = StoreScalar(cached);
  switch (scalar.shape) {
    case StoredScalar::Shape::kNumber:
      emit_double(p, scalar.number);
      break;
    case StoredScalar::Shape::kBool:
      type = XlsbRecordType::BrtFmlaBool;
      emit_u8(p, scalar.byte);
      break;
    case StoredScalar::Shape::kError:
      type = XlsbRecordType::BrtFmlaError;
      emit_u8(p, scalar.byte);
      break;
    case StoredScalar::Shape::kOther:
      if (cached.is_text()) {
        type = XlsbRecordType::BrtFmlaString;
        emit_xlwidestring(p, cached.as_text());
      } else {
        // Blank / Array / Lambda / Ref: Excel always materialises a
        // formula cell's cached value as one of the four scalar kinds
        // above. For round-trip we default to BrtFmlaNum with a 0.0
        // payload — the read-back path uses the decoded formula, and this
        // keeps the wire format well-formed.
        emit_double(p, 0.0);
      }
      break;
  }
  emit_u16(p, always_calculates ? 0x0002U : 0U);  // grbitFlags: fAlwaysCalc

  emit_cell_parsed_formula(p, formula.rgce, formula.rgcb);

  emit_record(dst, static_cast<std::uint16_t>(type), p);
}

/// Builds the `CellParsedFormula` of a dynamic-array shell cell: a single
/// `PtgExp` token (opcode 0x01 + the anchor row as u32) in `rgce`, with
/// the anchor column carried as a u32 in `rgcb`. Verified against a real
/// Excel-365-produced `xl/worksheets/sheetN.bin`: every cell of a spilled
/// footprint (anchor and phantoms alike) stores this shell, and the real
/// tokens live once in the anchor's `BrtArrFmla`.
EncodedFormula MakePtgExpShell(std::uint32_t anchor_row, std::uint32_t anchor_col) {
  EncodedFormula shell;
  emit_u8(shell.rgce, 0x01U);  // PtgExp
  emit_u32(shell.rgce, anchor_row);
  emit_u32(shell.rgcb, anchor_col);
  return shell;
}

/// Emits a `BrtArrFmla` record. Layout ([MS-XLSB], matching the reader's
/// decode): RfX (rwFirst, rwLast, colFirst, colLast as u32) + 1 flag byte
/// + `CellParsedFormula` (cce + rgce + cb + rgcb).
void EmitArrayFormulaRecord(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, std::uint32_t rw_first,
                            std::uint32_t rw_last, std::uint32_t col_first, std::uint32_t col_last,
                            const EncodedFormula& formula, bool always_calculates) {
  scratch.clear();
  std::vector<std::uint8_t>& p = scratch;
  emit_rfx(p, MergeRange{rw_first, col_first, rw_last, col_last});
  emit_u8(p, always_calculates ? 0x01U : 0U);  // flags: fAlwaysCalc, the `aca="1"` of .xlsx
  emit_cell_parsed_formula(p, formula.rgce, formula.rgcb);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla), p);
}

/// Emits a literal record for `cached`. Used for plain literal cells.
///
/// `phonetic` travels with a Text payload into the shared-string table:
/// the guide belongs to the string entry rather than the cell record, so
/// two cells reading the same kanji differently need distinct entries.
void EmitLiteralCellRecord(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, std::uint32_t col,
                           std::uint32_t xf_index, const Value& cached, const std::vector<PhoneticRun>& phonetic,
                           PhoneticProperties phonetic_props, SstBuilder& sst) {
  scratch.clear();
  std::vector<std::uint8_t>& p = scratch;
  EmitCellHeader(p, col, xf_index);
  XlsbRecordType type = XlsbRecordType::BrtCellBlank;
  switch (cached.kind()) {
    case ValueKind::Blank:
      break;
    case ValueKind::Number:
    case ValueKind::Bool:
    case ValueKind::Error: {
      const StoredScalar scalar = StoreScalar(cached);
      if (scalar.shape == StoredScalar::Shape::kNumber) {
        if (rk_round_trips_value(scalar.number)) {
          type = XlsbRecordType::BrtCellRk;
          emit_rk_number(p, scalar.number);
        } else {
          type = XlsbRecordType::BrtCellReal;
          emit_double(p, scalar.number);
        }
      } else {
        type = scalar.shape == StoredScalar::Shape::kBool ? XlsbRecordType::BrtCellBool : XlsbRecordType::BrtCellError;
        emit_u8(p, scalar.byte);
      }
      break;
    }
    case ValueKind::Text:
      type = XlsbRecordType::BrtCellIsst;
      emit_u32(p, sst.intern(cached.as_text(), phonetic, phonetic_props));
      break;
    case ValueKind::Array:
    case ValueKind::Lambda:
    case ValueKind::Ref:
      // These kinds never appear as a plain cell `cached_value` after
      // recalc — the recalc engine spills `Array` results into
      // SpillRegions, and `Lambda` / `Ref` are not legal cell payload.
      // Defensive fallback: emit a blank cell so the file stays valid.
      StructuredLog("xlsb.writer.unsupported_cached_value")
          .field("col", static_cast<std::int64_t>(col))
          .field("kind", static_cast<std::int64_t>(cached.kind()))
          .warn();
      break;
  }
  emit_record(dst, static_cast<std::uint16_t>(type), p);
}

}  // namespace

void emit_blank(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, std::uint32_t col,
                std::uint32_t xf_index) {
  scratch.clear();
  // The style index alone: a run column never sets the fPhShow bit.
  EmitCellHeader(scratch, col, xf_index & 0xFFFFFFU);
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank), scratch);
}

Expected<void, Error> emit_cell(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, const Cell& cell,
                                std::uint32_t row, std::uint32_t col, std::uint32_t xf_index, bool always_calculates,
                                SstBuilder& sst, const std::vector<std::string>& sheet_names,
                                const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                std::uint32_t* downgraded_formula_count, const NameShapes& name_shapes) {
  // Formula cells take precedence: even if the cached_value is blank, we
  // still emit a BrtFmla* record so the formula round-trips.
  if (!cell.formula_text.empty()) {
    auto formula_or = EncodeCellFormula(cell, sheet_names, sheet_ranges, name_table, false, name_shapes);
    if (!formula_or) {
      StructuredLog("xlsb.writer.formula_downgraded")
          .field("row", static_cast<std::int64_t>(row))
          .field("col", static_cast<std::int64_t>(col))
          .field("reason", formula_or.error().message)
          .warn();
      if (downgraded_formula_count != nullptr) {
        ++*downgraded_formula_count;
      }
      EmitLiteralCellRecord(dst, scratch, col, StyleWord(cell, xf_index), cell.cached_value, cell.phonetic_runs,
                            cell.phonetic_props, sst);
      return Expected<void, Error>::Ok();
    }
    EmitFormulaCellRecord(dst, scratch, col, StyleWord(cell, xf_index), cell.cached_value, formula_or.value(),
                          always_calculates);
    return Expected<void, Error>::Ok();
  }

  EmitLiteralCellRecord(dst, scratch, col, StyleWord(cell, xf_index), cell.cached_value, cell.phonetic_runs,
                        cell.phonetic_props, sst);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> emit_array_anchor(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch,
                                        const Cell& cell, const Value& anchor_value, bool always_calculates,
                                        std::uint32_t col, std::uint32_t xf_index, std::uint32_t anchor_row,
                                        std::uint32_t last_row, std::uint32_t last_col,
                                        const std::vector<std::string>& sheet_names,
                                        const SheetRangeTable& sheet_ranges, const NameTable& name_table,
                                        SstBuilder& sst, std::uint32_t* downgraded_formula_count,
                                        bool* downgraded_to_literal, const NameShapes& name_shapes) {
  auto formula_or = EncodeCellFormula(cell, sheet_names, sheet_ranges, name_table, true, name_shapes);
  if (!formula_or) {
    StructuredLog("xlsb.writer.array_formula_downgraded")
        .field("row", static_cast<std::int64_t>(anchor_row))
        .field("col", static_cast<std::int64_t>(col))
        .field("reason", formula_or.error().message)
        .warn();
    if (downgraded_formula_count != nullptr) {
      ++*downgraded_formula_count;
    }
    if (downgraded_to_literal != nullptr) {
      *downgraded_to_literal = true;
    }
    EmitLiteralCellRecord(dst, scratch, col, StyleWord(cell, xf_index), anchor_value, cell.phonetic_runs,
                          cell.phonetic_props, sst);
    return Expected<void, Error>::Ok();
  }
  // The anchor's own cell record is a PtgExp shell typed by the spilled
  // value at the anchor; the real tokens go into the following BrtArrFmla.
  const EncodedFormula shell = MakePtgExpShell(anchor_row, col);
  EmitFormulaCellRecord(dst, scratch, col, StyleWord(cell, xf_index), anchor_value, shell, always_calculates);
  EmitArrayFormulaRecord(dst, scratch, anchor_row, last_row, col, last_col, formula_or.value(), always_calculates);
  return Expected<void, Error>::Ok();
}

void emit_data_table_anchor(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, const Cell& cell,
                            std::uint32_t col, std::uint32_t xf_index, const DataTable& table, SstBuilder& sst) {
  // TABLECELL = BrtTable CELLMETA (BrtCellRk / BrtCellError / BrtCellBool / BrtCellReal / BrtCellSt).
  scratch.clear();
  std::vector<std::uint8_t>& p = scratch;
  emit_rfx(p, table.ref);
  for (const DataTableInput* input : {&table.r1, &table.r2}) {
    const bool usable = input->present && !input->deleted;
    emit_u32(p, usable ? input->row : 0U);
    emit_u32(p, usable ? input->col : 0U);
  }
  emit_u8(p, static_cast<std::uint8_t>(
                 (table.row_oriented ? kTableRowInput : 0U) | (table.two_dimensional ? kTableTwoDimensional : 0U) |
                 (table.r1.deleted ? kTableDeletedInput1 : 0U) | (table.r2.deleted ? kTableDeletedInput2 : 0U) |
                 (table.always_calc ? kTableAlwaysCalc : 0U)));
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtTable), p);
  const Value& v = cell.cached_value;
  if (v.is_text()) {
    p.clear();
    EmitCellHeader(p, col, StyleWord(cell, xf_index));
    emit_xlwidestring(p, v.as_text());
    emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellSt), p);
    return;
  }
  // A blank result reads as 0 in a data table, and the grammar has no blank anchor.
  const Value scalar = v.is_number() || v.is_boolean() || v.is_error() ? v : Value::number(0.0);
  EmitLiteralCellRecord(dst, scratch, col, StyleWord(cell, xf_index), scalar, cell.phonetic_runs, cell.phonetic_props,
                        sst);
}

void emit_array_phantom(std::vector<std::uint8_t>& dst, std::vector<std::uint8_t>& scratch, std::uint32_t col,
                        std::uint32_t xf_index, const Value& cached, std::uint32_t anchor_row, std::uint32_t anchor_col,
                        bool always_calculates) {
  const EncodedFormula shell = MakePtgExpShell(anchor_row, anchor_col);
  EmitFormulaCellRecord(dst, scratch, col, xf_index, cached, shell, always_calculates);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
