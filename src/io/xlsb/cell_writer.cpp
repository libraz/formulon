// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the per-cell XLSB record dispatcher. See
// `io/xlsb/cell_writer.h` for the contract.

#include "io/xlsb/cell_writer.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <vector>

#include "cell.h"
#include "io/xlsb/ptg_writer.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sst_writer.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "utils/arena.h"
#include "utils/structured_log.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Emits the standard XLSB cell-header (column, iStyleRef, fPhShow):
///   * column    : u32
///   * iStyleRef : 3 bytes, little-endian (index into the xf table)
///   * fPhShow   : u8 (0)
///
/// `xf_index` is the cell's style-table index; it must round-trip so a
/// styled cell keeps its formatting through `write_xlsb -> read_xlsb`.
/// `ReadCellHeader` decodes the same 3-byte little-endian layout.
void EmitCellHeader(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index) {
  emit_u32(dst, col);
  // iStyleRef: 24-bit little-endian style index. Excel xf indices always
  // fit in 24 bits; mask defensively so a corrupt oversized index cannot
  // bleed into the phonetic flag.
  emit_u8(dst, static_cast<std::uint8_t>(xf_index & 0xFFU));
  emit_u8(dst, static_cast<std::uint8_t>((xf_index >> 8) & 0xFFU));
  emit_u8(dst, static_cast<std::uint8_t>((xf_index >> 16) & 0xFFU));
  emit_u8(dst, 0);  // fPhShow
}

/// Returns the OOXML wire code for `e`. Mirrors the inverse mapping
/// `read_xlsb` performs in `BrtCellError`.
std::uint8_t ErrorWireCode(ErrorCode e) {
  const std::int32_t code = ooxml_code(e);
  if (code < 0 || code > 0xFF) {
    return 0x09;  // `#UNKNOWN!` wire code
  }
  return static_cast<std::uint8_t>(code);
}

/// Parses `cell.formula_text` into an AST and encodes it as a Ptg
/// (`rgce`) byte stream. The formula text starts with `=`; the parser
/// consumes the body after it. Returns the encoded bytes, or an Error
/// when the formula cannot be parsed or lowered to the supported Ptg
/// token set. The caller surfaces that error through `write_xlsb` rather
/// than silently dropping the formula.
Expected<EncodedFormula, Error> EncodeCellFormula(const Cell& cell, const std::vector<std::string>& sheet_names,
                                                  const SheetRangeTable& sheet_ranges, const NameTable& name_table) {
  std::string_view body(cell.formula_text);
  if (!body.empty() && body.front() == '=') {
    body.remove_prefix(1);
  }
  Arena arena;
  parser::Parser parser(body, arena);
  parser::AstNode* root = parser.parse();
  if (root == nullptr || !parser.errors().empty()) {
    return make_error(FormulonErrorCode::kIoXlsbUnsupportedPtg, "xlsb writer: formula failed to parse for Ptg encoding",
                      std::string("context=xlsb_cell_writer formula=") + cell.formula_text);
  }
  return encode_ptgs(*root, sheet_names, sheet_ranges, name_table);
}

/// Emits a `BrtFmla*` record matching `cached`'s kind, with the encoded
/// `rgce` / `rgcb` bytes spliced into the `CellParsedFormula` payload.
/// Layout per [MS-XLSB] §2.4.x:
///
///   cell-header (8 bytes)
///   value       (kind-specific)
///   grbitFlags  (u16, written as zero)
///   cce         (u32 byte length of rgce)
///   rgce        (cce bytes)         — the Ptg stream
///   cb          (u32 byte length of rgcb)
///   rgcb        (cb bytes)          — array-constant extra data
void EmitFormulaCellRecord(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index,
                           const Value& cached, const EncodedFormula& formula) {
  std::vector<std::uint8_t> p;
  EmitCellHeader(p, col, xf_index);

  XlsbRecordType type = XlsbRecordType::BrtFmlaNum;
  switch (cached.kind()) {
    case ValueKind::Number:
      type = XlsbRecordType::BrtFmlaNum;
      emit_double(p, cached.as_number());
      break;
    case ValueKind::Bool:
      type = XlsbRecordType::BrtFmlaBool;
      emit_u8(p, cached.as_boolean() ? 1U : 0U);
      break;
    case ValueKind::Error:
      type = XlsbRecordType::BrtFmlaError;
      emit_u8(p, ErrorWireCode(cached.as_error()));
      break;
    case ValueKind::Text:
      type = XlsbRecordType::BrtFmlaString;
      emit_xlwidestring(p, cached.as_text());
      break;
    default:
      // Blank / Array / Lambda / Ref: Excel always materialises a
      // formula cell's cached value as one of the four scalar kinds
      // above. For round-trip we default to BrtFmlaNum with a 0.0
      // payload — the read-back path uses the decoded formula, and this
      // keeps the wire format well-formed.
      type = XlsbRecordType::BrtFmlaNum;
      emit_double(p, 0.0);
      break;
  }
  emit_u16(p, 0);  // grbitFlags

  // CellParsedFormula: cce + rgce + cb + rgcb.
  emit_u32(p, static_cast<std::uint32_t>(formula.rgce.size()));
  p.insert(p.end(), formula.rgce.begin(), formula.rgce.end());
  emit_u32(p, static_cast<std::uint32_t>(formula.rgcb.size()));
  p.insert(p.end(), formula.rgcb.begin(), formula.rgcb.end());

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
  shell.rgce.push_back(0x01);  // PtgExp
  shell.rgce.push_back(static_cast<std::uint8_t>(anchor_row & 0xFFU));
  shell.rgce.push_back(static_cast<std::uint8_t>((anchor_row >> 8) & 0xFFU));
  shell.rgce.push_back(static_cast<std::uint8_t>((anchor_row >> 16) & 0xFFU));
  shell.rgce.push_back(static_cast<std::uint8_t>((anchor_row >> 24) & 0xFFU));
  shell.rgcb.push_back(static_cast<std::uint8_t>(anchor_col & 0xFFU));
  shell.rgcb.push_back(static_cast<std::uint8_t>((anchor_col >> 8) & 0xFFU));
  shell.rgcb.push_back(static_cast<std::uint8_t>((anchor_col >> 16) & 0xFFU));
  shell.rgcb.push_back(static_cast<std::uint8_t>((anchor_col >> 24) & 0xFFU));
  return shell;
}

/// Emits a `BrtArrFmla` record. Layout ([MS-XLSB], matching the reader's
/// decode): RfX (rwFirst, rwLast, colFirst, colLast as u32) + 1 flag byte
/// + `CellParsedFormula` (cce + rgce + cb + rgcb).
void EmitArrayFormulaRecord(std::vector<std::uint8_t>& dst, std::uint32_t rw_first, std::uint32_t rw_last,
                            std::uint32_t col_first, std::uint32_t col_last, const EncodedFormula& formula) {
  std::vector<std::uint8_t> p;
  emit_u32(p, rw_first);
  emit_u32(p, rw_last);
  emit_u32(p, col_first);
  emit_u32(p, col_last);
  emit_u8(p, 0);  // flags (fAlwaysCalc etc.); zero matches Excel's output here.
  emit_u32(p, static_cast<std::uint32_t>(formula.rgce.size()));
  p.insert(p.end(), formula.rgce.begin(), formula.rgce.end());
  emit_u32(p, static_cast<std::uint32_t>(formula.rgcb.size()));
  p.insert(p.end(), formula.rgcb.begin(), formula.rgcb.end());
  emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtArrFmla), p);
}

/// Emits a literal record for `cached`. Used for plain literal cells.
void EmitLiteralCellRecord(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index,
                           const Value& cached, SstBuilder& sst) {
  switch (cached.kind()) {
    case ValueKind::Blank: {
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank), p);
      return;
    }
    case ValueKind::Number: {
      const double v = cached.as_number();
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      if (rk_round_trips_value(v)) {
        emit_rk_number(p, v);
        emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellRk), p);
      } else {
        emit_double(p, v);
        emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellReal), p);
      }
      return;
    }
    case ValueKind::Bool: {
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      emit_u8(p, cached.as_boolean() ? 1U : 0U);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBool), p);
      return;
    }
    case ValueKind::Error: {
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      emit_u8(p, ErrorWireCode(cached.as_error()));
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellError), p);
      return;
    }
    case ValueKind::Text: {
      const std::uint32_t idx = sst.intern(cached.as_text());
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      emit_u32(p, idx);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellIsst), p);
      return;
    }
    case ValueKind::Array:
    case ValueKind::Lambda:
    case ValueKind::Ref: {
      // These kinds never appear as a plain cell `cached_value` after
      // recalc — the recalc engine spills `Array` results into
      // SpillRegions, and `Lambda` / `Ref` are not legal cell payload.
      // Defensive fallback: emit a blank cell so the file stays valid.
      StructuredLog("xlsb.writer.unsupported_cached_value")
          .field("col", static_cast<std::int64_t>(col))
          .field("kind", static_cast<std::int64_t>(cached.kind()))
          .warn();
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col, xf_index);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank), p);
      return;
    }
  }
}

}  // namespace

Expected<void, Error> emit_cell(std::vector<std::uint8_t>& dst, const Cell& cell, std::uint32_t row, std::uint32_t col,
                                SstBuilder& sst, const std::vector<std::string>& sheet_names,
                                const SheetRangeTable& sheet_ranges, const NameTable& name_table) {
  // Formula cells take precedence: even if the cached_value is blank, we
  // still emit a BrtFmla* record so the formula round-trips.
  if (!cell.formula_text.empty()) {
    auto formula_or = EncodeCellFormula(cell, sheet_names, sheet_ranges, name_table);
    if (!formula_or) {
      // A formula we cannot encode must NOT be silently dropped to a
      // literal: that would lose the formula. Surface the failure so
      // `write_xlsb` returns it to the caller.
      StructuredLog("xlsb.writer.formula_encode_failed")
          .field("row", static_cast<std::int64_t>(row))
          .field("col", static_cast<std::int64_t>(col))
          .field("reason", formula_or.error().message)
          .warn();
      return formula_or.error();
    }
    EmitFormulaCellRecord(dst, col, cell.xf_index, cell.cached_value, formula_or.value());
    return Expected<void, Error>::Ok();
  }

  EmitLiteralCellRecord(dst, col, cell.xf_index, cell.cached_value, sst);
  return Expected<void, Error>::Ok();
}

Expected<void, Error> emit_array_anchor(std::vector<std::uint8_t>& dst, const Cell& cell, const Value& anchor_value,
                                        std::uint32_t col, std::uint32_t anchor_row, std::uint32_t last_row,
                                        std::uint32_t last_col, const std::vector<std::string>& sheet_names,
                                        const SheetRangeTable& sheet_ranges, const NameTable& name_table) {
  auto formula_or = EncodeCellFormula(cell, sheet_names, sheet_ranges, name_table);
  if (!formula_or) {
    return formula_or.error();
  }
  // The anchor's own cell record is a PtgExp shell typed by the spilled
  // value at the anchor; the real tokens go into the following BrtArrFmla.
  const EncodedFormula shell = MakePtgExpShell(anchor_row, col);
  EmitFormulaCellRecord(dst, col, cell.xf_index, anchor_value, shell);
  EmitArrayFormulaRecord(dst, anchor_row, last_row, col, last_col, formula_or.value());
  return Expected<void, Error>::Ok();
}

void emit_array_phantom(std::vector<std::uint8_t>& dst, std::uint32_t col, std::uint32_t xf_index, const Value& cached,
                        std::uint32_t anchor_row, std::uint32_t anchor_col) {
  const EncodedFormula shell = MakePtgExpShell(anchor_row, anchor_col);
  EmitFormulaCellRecord(dst, col, xf_index, cached, shell);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
