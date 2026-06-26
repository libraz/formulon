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
///   * iStyleRef : 3 bytes (we always write zero)
///   * fPhShow   : u8 (0)
void EmitCellHeader(std::vector<std::uint8_t>& dst, std::uint32_t col) {
  emit_u32(dst, col);
  // iStyleRef is a 3-byte little-endian field; emit as three zero
  // bytes. The reader reads it as part of a 4-byte block (3 style + 1
  // fPhShow), so we follow it with the phonetic flag.
  emit_u8(dst, 0);
  emit_u8(dst, 0);
  emit_u8(dst, 0);
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
Expected<std::vector<std::uint8_t>, Error> EncodeCellFormula(const Cell& cell,
                                                             const std::vector<std::string>& sheet_names) {
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
  return encode_ptgs(*root, sheet_names);
}

/// Emits a `BrtFmla*` record matching `cached`'s kind, with the encoded
/// `rgce` bytes spliced into the `CellParsedFormula` payload. Layout per
/// [MS-XLSB] §2.4.x:
///
///   cell-header (8 bytes)
///   value       (kind-specific)
///   grbitFlags  (u16, written as zero)
///   cce         (u32 byte length of rgce)
///   rgce        (cce bytes)         — the Ptg stream
///   cb          (u32 byte length of rgcb, written as zero)
void EmitFormulaCellRecord(std::vector<std::uint8_t>& dst, std::uint32_t col, const Value& cached,
                           const std::vector<std::uint8_t>& rgce) {
  std::vector<std::uint8_t> p;
  EmitCellHeader(p, col);

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

  // CellParsedFormula: cce + rgce + cb (+ rgcb, empty).
  emit_u32(p, static_cast<std::uint32_t>(rgce.size()));
  p.insert(p.end(), rgce.begin(), rgce.end());
  emit_u32(p, 0);  // cb (rgcb length)

  emit_record(dst, static_cast<std::uint16_t>(type), p);
}

/// Emits a literal record for `cached`. Used for plain literal cells.
void EmitLiteralCellRecord(std::vector<std::uint8_t>& dst, std::uint32_t col, const Value& cached, SstBuilder& sst) {
  switch (cached.kind()) {
    case ValueKind::Blank: {
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank), p);
      return;
    }
    case ValueKind::Number: {
      const double v = cached.as_number();
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col);
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
      EmitCellHeader(p, col);
      emit_u8(p, cached.as_boolean() ? 1U : 0U);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBool), p);
      return;
    }
    case ValueKind::Error: {
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col);
      emit_u8(p, ErrorWireCode(cached.as_error()));
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellError), p);
      return;
    }
    case ValueKind::Text: {
      const std::uint32_t idx = sst.intern(cached.as_text());
      std::vector<std::uint8_t> p;
      EmitCellHeader(p, col);
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
      EmitCellHeader(p, col);
      emit_record(dst, static_cast<std::uint16_t>(XlsbRecordType::BrtCellBlank), p);
      return;
    }
  }
}

}  // namespace

Expected<void, Error> emit_cell(std::vector<std::uint8_t>& dst, const Cell& cell, std::uint32_t row, std::uint32_t col,
                                SstBuilder& sst, const std::vector<std::string>& sheet_names) {
  // Formula cells take precedence: even if the cached_value is blank, we
  // still emit a BrtFmla* record so the formula round-trips.
  if (!cell.formula_text.empty()) {
    auto rgce_or = EncodeCellFormula(cell, sheet_names);
    if (!rgce_or) {
      // A formula we cannot encode must NOT be silently dropped to a
      // literal: that would lose the formula. Surface the failure so
      // `write_xlsb` returns it to the caller.
      StructuredLog("xlsb.writer.formula_encode_failed")
          .field("row", static_cast<std::int64_t>(row))
          .field("col", static_cast<std::int64_t>(col))
          .field("reason", rgce_or.error().message)
          .warn();
      return rgce_or.error();
    }
    EmitFormulaCellRecord(dst, col, cell.cached_value, rgce_or.value());
    return Expected<void, Error>::Ok();
  }

  EmitLiteralCellRecord(dst, col, cell.cached_value, sst);
  return Expected<void, Error>::Ok();
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
