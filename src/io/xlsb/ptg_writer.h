// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Parser AST -> MS-XLSB Ptg stream encoder. The exact inverse of
// `ptg_reader.h`: a post-order walk of the AST emits operand tokens
// first, then the operator / function token that consumes them, so the
// resulting `rgce` byte stream decodes back to a structurally
// equivalent AST.
//
// The encoder is matched byte-for-byte with the decoder in this module
// (the two are the engine's own private round-trip pair); the on-wire
// shapes follow [MS-XLSB] §2.5.97 for the common token set but the
// decoder is the authoritative consumer, so the encoder need only stay
// consistent with it.
//
// Tokens the AST can carry but the encoder cannot lower (defined-name
// refs, structured refs, external refs, lambda / let forms, spilled
// refs, implicit-intersection) return `kIoXlsbUnsupportedPtg` rather
// than silently dropping data; the cell writer surfaces that as a hard
// failure through `write_xlsb`'s `Expected` return.

#ifndef FORMULON_IO_XLSB_PTG_WRITER_H_
#define FORMULON_IO_XLSB_PTG_WRITER_H_

#include <cstdint>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "parser/ast.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Maps a name (a genuine defined name, or a hidden `_xlfn.*` /
/// `_xlpm.*` future-function / LET-parameter placeholder) to its
/// 1-based `BrtName` declaration index. Shared between the encoder
/// (`encode_ptgs`, which emits `PtgName(ilbl)` for any lookup hit) and
/// the top-level writer (which emits the matching `BrtName` records
/// into `xl/workbook.bin`'s globals).
using NameTable = std::unordered_map<std::string, std::uint32_t>;

/// Ordered `(itabFirst, itabLast)` pairs, mirroring `xl/workbook.bin`'s
/// `BrtExternSheet` table: entry `i` is the range a `PtgRef3d` /
/// `PtgArea3d` token resolves via `ixti == i`. A single-sheet qualified
/// reference (`Sheet1!A1`) stores `itabFirst == itabLast`; a genuine
/// 3-D range (`Sheet1:Sheet3!A1`) stores the full span. Built once per
/// workbook by `collect_ptg_sheet_ranges` (mirroring `NameTable` /
/// `collect_ptg_names`) so every sheet's cell encoder and the
/// `BrtExternSheet` record the top-level writer emits agree on `ixti`
/// assignments.
using SheetRangeTable = std::vector<std::pair<std::int32_t, std::int32_t>>;

/// Result of `encode_ptgs`: the main token stream plus the array-
/// constant extra-data area a `CellParsedFormula` appends after it.
struct EncodedFormula {
  /// The `rgce` Ptg token stream.
  std::vector<std::uint8_t> rgce;
  /// The `rgcb` extra-data area `PtgArray` tokens in `rgce` reference
  /// (rows/cols + tagged elements, consumed by the decoder in
  /// encounter order). Empty when the formula carries no array
  /// constants. The caller emits this as the `CellParsedFormula`'s
  /// `cb` (byte length) + `rgcb` (payload) trailer, after `cce` + `rgce`.
  std::vector<std::uint8_t> rgcb;
};

/// Walks `node`'s AST appending, in encounter order, every name a
/// `PtgName` reference will be needed for while encoding it:
///
///   * A `Call` node whose name is not resolvable via `func_id_table`
///     (a post-2007 "future function": XLOOKUP, TEXTJOIN, CONCAT, IFS,
///     SEQUENCE, ...).
///   * A `NameRef` node (an ordinary defined-name reference, e.g.
///     `Rate`).
///
/// Names already present in `seen` are skipped (both to dedupe and so
/// callers can pre-seed `seen` with names that already have an assigned
/// `ilbl`, e.g. from the workbook's existing defined-name table).
/// `LetBinding` / `Lambda` nodes are not walked into by this pass (the
/// encoder does not yet lower them — see `ptg_writer.cpp`'s `emit`).
void collect_ptg_names(const parser::AstNode& node, std::vector<std::string>& names,
                       std::unordered_set<std::string>& seen);

/// Walks `node`'s AST appending, in encounter order, every distinct
/// `(itabFirst, itabLast)` sheet-range pair a qualified reference will
/// need an `ixti` for while encoding it: a single-sheet `Ref` whose
/// `sheet` is non-empty contributes `(itab, itab)`; a `Ref3D` node
/// contributes its full `(begin, end)` span. `sheet_names` resolves a
/// sheet display name to its 0-based index; a name absent from
/// `sheet_names` is skipped here (the encode fails later with a precise
/// error instead of silently fabricating an entry). `seen` dedupes
/// (both across one call and across callers pre-seeding it), and
/// `ranges`' index order becomes the `ixti` assignment `encode_ptgs`
/// consults via `sheet_ranges`.
void collect_ptg_sheet_ranges(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                              SheetRangeTable& ranges, std::unordered_set<std::uint64_t>& seen);

/// Encodes the AST rooted at `node` into an `rgce` Ptg byte stream (plus
/// its `rgcb` array-constant extra data, see `EncodedFormula`).
/// `sheet_names` maps a sheet display name to its 0-based index.
/// `sheet_ranges` maps a resolved `(itabFirst, itabLast)` pair to its
/// `ixti` (its index in the table, built by `collect_ptg_sheet_ranges`);
/// a qualified reference whose sheet(s) are not present in
/// `sheet_ranges` fails the encode rather than fabricating an entry. (In
/// practice the cell writer passes a table built from every formula in
/// the workbook, so any sheet referenced by a live formula resolves.)
/// `name_table` resolves `NameRef` nodes and future-function `Call`
/// callees to a `PtgName` `ilbl` (see `collect_ptg_names`); a name
/// absent from the table fails the encode.
///
/// Returns `kIoXlsbUnsupportedPtg` for any node kind outside the
/// supported set (see header banner). The error context names the
/// offending node kind.
Expected<EncodedFormula, Error> encode_ptgs(const parser::AstNode& node, const std::vector<std::string>& sheet_names,
                                            const SheetRangeTable& sheet_ranges, const NameTable& name_table);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_WRITER_H_
