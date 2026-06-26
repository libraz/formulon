// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// MS-XLSB Ptg stream -> parser AST decoder.
//
// The XLSB `CellParsedFormula` body (the `rgce` Ptg byte stream) is
// reverse-Polish: operand tokens push onto a stack, operator and
// function tokens pop their operands and push a result. This module
// runs that stack machine and rebuilds the same `parser::AstNode` tree
// the Pratt parser would produce for the equivalent A1 formula text, so
// the rest of the engine (formatter, evaluator) consumes a single AST
// representation regardless of the source container (xlsx text vs xlsb
// Ptg bytes).
//
// Coverage is the common token set that appears in the overwhelming
// majority of real formulas (see `decode_ptgs` doc for the full list).
// Tokens outside that set are surfaced as `kIoXlsbUnsupportedPtg` so the
// reader can fall back to preserving the cached value rather than
// inventing a wrong formula. Every byte read is bounds-checked against
// the record payload — this is binary parsing of untrusted input.
//
// Design references:
//   * [MS-XLSB] §2.5.97 (Ptg) and the per-Ptg sub-sections.

#ifndef FORMULON_IO_XLSB_PTG_READER_H_
#define FORMULON_IO_XLSB_PTG_READER_H_

#include <string>
#include <vector>

#include "io/zip_reader.h"
#include "parser/ast.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Decodes the `rgce` Ptg byte stream `ptgs` into a `parser::AstNode`
/// tree allocated in `arena`. The returned node is the formula root
/// (the single value left on the operand stack after the RPN walk).
///
/// Supported tokens (status `Full` in `ptg.h`):
///   * Operands: PtgInt, PtgNum, PtgStr, PtgBool, PtgErr, PtgMissArg,
///     PtgArray (constant arrays).
///   * References: PtgRef, PtgArea, PtgRef3d, PtgArea3d and their
///     RefErr / AreaErr forms (decoded to `#REF!`).
///   * Operators: Add/Sub/Mul/Div/Power/Concat, the six comparisons,
///     Uplus/Uminus/Percent, Paren, Range(`:`), Union(`,`),
///     Isect(space).
///   * Functions: PtgFunc (fixed arity) and PtgFuncVar (variable arity)
///     resolved via `func_id_table`.
///   * PtgAttr sub-kinds that are structurally transparent (Space is
///     dropped, Sum collapses to a unary SUM call, If/Choose/Goto
///     control jumps are consumed without changing the operand stack).
///
/// `sheet_names` maps a 0-based sheet index to a display name; it is
/// consulted for the 3-D reference forms (PtgRef3d / PtgArea3d), whose
/// wire encoding carries an `ixti` index rather than a name. When the
/// index is out of range the decoder emits `#REF!` for the reference
/// (Excel's own behaviour for a dangling sheet index) rather than
/// failing the whole formula.
///
/// Errors:
///   * `kIoXlsbUnsupportedPtg` — a token outside the supported set
///     (PtgExp/PtgTbl/PtgName/PtgNameX/PtgMemFunc materialisation/...).
///   * `kIoXlsbRecordTruncated` — a token's payload would overrun
///     `ptgs`.
///   * `kIoXlsbCorrupt` — the operand stack is unbalanced (too few
///     operands for an operator, or not exactly one value remaining).
///   * `kOutOfMemory` — arena allocation failed.
Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, Arena& arena, const std::vector<std::string>& sheet_names);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_READER_H_
