// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the per-cell XLSB record dispatcher. See
// `io/xlsb/cell_writer.h` for the contract.

#include "io/xlsb/cell_writer.h"

#include <cstdint>
#include <cstring>
#include <string_view>
#include <vector>

#include "cell.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xlsb/sst_writer.h"
#include "utils/structured_log.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Stub prefix emitted by Bundle 4.1's reader for every formula cell
/// whose Ptg payload could not yet be decoded into Excel-formula
/// text. The writer recognises the same prefix to splice the captured
/// Ptg bytes back into a `BrtFmla*` record.
constexpr std::string_view kFormulaStubPrefix = "=__FORMULON_XLSB_PTG__(";
constexpr char kFormulaStubSuffix = ')';

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

/// Decodes one nibble of a hex digit. Returns -1 on invalid input.
int HexNibble(char c) {
  if (c >= '0' && c <= '9') {
    return c - '0';
  }
  if (c >= 'A' && c <= 'F') {
    return 10 + (c - 'A');
  }
  if (c >= 'a' && c <= 'f') {
    return 10 + (c - 'a');
  }
  return -1;
}

/// Tries to parse the dot-separated hex byte sequence inside the
/// Bundle 4.1 stub. Returns the captured Ptg bytes on success, or an
/// empty optional when the body is not a clean dotted-hex sequence
/// (the writer then falls back to the literal path).
struct StubParseResult {
  bool ok = false;
  std::vector<std::uint8_t> ptg_bytes;
};

StubParseResult ParseStubBody(std::string_view body) {
  StubParseResult out;
  if (body.empty()) {
    out.ok = true;
    return out;
  }
  // Body is `HH.HH.HH...`: pairs of hex digits separated by dots.
  std::size_t i = 0;
  while (i < body.size()) {
    if (i + 2 > body.size()) {
      return {};  // dangling single nibble
    }
    const int hi = HexNibble(body[i]);
    const int lo = HexNibble(body[i + 1]);
    if (hi < 0 || lo < 0) {
      return {};
    }
    out.ptg_bytes.push_back(static_cast<std::uint8_t>((hi << 4) | lo));
    i += 2;
    if (i == body.size()) {
      break;
    }
    if (body[i] != '.') {
      return {};
    }
    ++i;  // consume separator
  }
  out.ok = true;
  return out;
}

/// Identifies a stub formula and extracts the captured Ptg bytes.
/// Returns `{true, bytes}` for clean stubs and `{false, {}}` when the
/// formula text is not a Bundle 4.1 stub (or is malformed).
StubParseResult MatchFormulaStub(std::string_view formula_text) {
  if (formula_text.size() < kFormulaStubPrefix.size() + 1U) {
    return {};
  }
  if (formula_text.substr(0, kFormulaStubPrefix.size()) != kFormulaStubPrefix) {
    return {};
  }
  if (formula_text.back() != kFormulaStubSuffix) {
    return {};
  }
  const std::string_view body =
      formula_text.substr(kFormulaStubPrefix.size(), formula_text.size() - kFormulaStubPrefix.size() - 1);
  return ParseStubBody(body);
}

/// Emits a `BrtFmla*` record matching `cached`'s kind. The captured
/// `ptg_bytes` go into the `CellParsedFormula` (rgce) suffix; the
/// reader treats anything past the grbitFlags as opaque, so we just
/// concatenate the bytes as-is. Layout per [MS-XLSB] §2.4.x:
///
///   cell-header (8 bytes)
///   value       (kind-specific)
///   grbitFlags  (u16, written as zero)
///   cce         (u32 byte length of rgce)        — part of CellParsedFormula
///   rgce        (cce bytes)                       — Ptg byte-stream
///   cb          (u32 byte length of rgcb)        — extra-data byte length
///   rgcb        (cb bytes, written as zero bytes) — extras
///
/// The reader's Bundle 4.1 path does not parse rgce / rgcb separately;
/// it slices everything past grbitFlags as opaque bytes and re-stubs
/// the entire suffix on the next read. So as long as we put the
/// captured `ptg_bytes` back at the same offset, the round-trip is
/// stable.
void EmitFormulaCellRecord(std::vector<std::uint8_t>& dst, std::uint32_t col, const Value& cached,
                           const std::vector<std::uint8_t>& ptg_bytes) {
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
      // payload — the read-back path discards it anyway, and this
      // keeps the wire format well-formed.
      type = XlsbRecordType::BrtFmlaNum;
      emit_double(p, 0.0);
      break;
  }
  emit_u16(p, 0);  // grbitFlags

  // rgce + rgcb. The reader stubs everything past grbitFlags as
  // opaque bytes, so we put the original Ptg bytes verbatim. The
  // captured bytes already include any `cce`/`cb` length prefixes
  // that were present in the source archive.
  if (!ptg_bytes.empty()) {
    p.insert(p.end(), ptg_bytes.begin(), ptg_bytes.end());
  }

  emit_record(dst, static_cast<std::uint16_t>(type), p);
}

/// Emits a literal record for `cached`. Used for plain literal cells
/// and for the formula-lost fallback path.
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

void emit_cell(std::vector<std::uint8_t>& dst, const Cell& cell, std::uint32_t row, std::uint32_t col,
               SstBuilder& sst) {
  // Formula cells take precedence: even if the cached_value is blank,
  // we still want to emit a BrtFmla* record so the formula text round-
  // trips. (The reader resets cached_value to blank when it sees a
  // formula cell, which would otherwise turn into an empty literal
  // here.)
  if (!cell.formula_text.empty()) {
    const StubParseResult parsed = MatchFormulaStub(cell.formula_text);
    if (parsed.ok) {
      EmitFormulaCellRecord(dst, col, cell.cached_value, parsed.ptg_bytes);
      return;
    }
    // Authored-in-engine formula or malformed stub: we have no Ptg
    // bytes to round-trip. Surface the loss explicitly so callers can
    // see it, then drop to the literal path. This is the v1 contract;
    // a later bundle replaces it with AST→Ptg encoding.
    StructuredLog("xlsb.writer.formula_lost")
        .field("row", static_cast<std::int64_t>(row))
        .field("col", static_cast<std::int64_t>(col))
        .field("formula_size", static_cast<std::int64_t>(cell.formula_text.size()))
        .warn();
    EmitLiteralCellRecord(dst, col, cell.cached_value, sst);
    return;
  }

  EmitLiteralCellRecord(dst, col, cell.cached_value, sst);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
