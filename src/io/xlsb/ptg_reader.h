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

#include <cstdint>
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

/// One workbook-scope `BrtName` entry: a `<definedName>`-equivalent or a
/// hidden `_xlfn.*` / `_xlpm.*` future-function / LET-LAMBDA-parameter
/// placeholder. `PtgName`'s `ilbl` (1-based) indexes this table in
/// declaration order.
struct XlsbName {
  /// `-1` = workbook scope; otherwise the 0-based sheet index the name
  /// is local to.
  std::int32_t itab = -1;
  std::string name;
  /// `BrtName`'s `fHidden` bit. Set for the `_xlfn.*` / `_xlpm.*`
  /// future-function and LET/LAMBDA-parameter placeholders `PtgName`
  /// resolves internally; clear for genuine user-visible defined names
  /// that should also be registered on the workbook's defined-name
  /// table (see `RegisterDefinedNames`).
  bool hidden = false;
};

/// One `BrtExternSheet` entry: resolves a `PtgRef3d` / `PtgArea3d`
/// `ixti` (a direct 0-based index into this table) to the 0-based sheet
/// range it qualifies. `itab_first == itab_last` is an ordinary
/// single-sheet qualified reference (`Sheet2!A1`); `itab_first !=
/// itab_last` is a genuine 3-D range (`Sheet1:Sheet3!A1`).
struct XlsbSheetRange {
  std::int32_t itab_first = -1;
  std::int32_t itab_last = -1;
};

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
///     resolved via `func_id_table`; PtgFuncVar with the `id == 255`
///     future-function sentinel resolves the callee from a preceding
///     `PtgName` operand instead (post-2007 functions such as XLOOKUP,
///     LET, TEXTJOIN, CONCAT, IFS, SEQUENCE store their name this way —
///     see `name_table`).
///   * PtgAttr sub-kinds that are structurally transparent (Space is
///     dropped, Sum collapses to a unary SUM call, If/Choose/Goto
///     control jumps are consumed without changing the operand stack).
///   * PtgName: resolved through `name_table` into a `NameRef` (ordinary
///     defined-name reference) unless the resolved name is consumed by
///     the future-function dispatch above.
///
/// `rgcb` is the `CellParsedFormula`'s extra-data area (the bytes after
/// `rgce`, i.e. `cb` + its payload in the caller's framing). `PtgArray`
/// stores only an 8-byte placeholder in `rgce`; the real dimensions and
/// element values live in `rgcb`, consumed in encounter order (each
/// `PtgArray` token advances an internal cursor into `rgcb` by exactly
/// its own array's worth of bytes). Pass an empty span when the caller
/// knows the formula carries no array constants.
///
/// `sheet_names` maps a 0-based sheet index to a display name.
///
/// `name_table` is the workbook's `BrtName` list in declaration order;
/// `PtgName`'s 1-based `ilbl` indexes it (`name_table[ilbl - 1]`).
///
/// `sheet_ranges` is the workbook's `BrtExternSheet` list; `PtgRef3d` /
/// `PtgArea3d`'s `ixti` directly indexes it (0-based) to resolve the
/// qualified sheet range. When `sheet_ranges` is empty (no
/// `BrtExternSheet` record — e.g. a workbook with no qualified
/// references at all) the decoder falls back to treating `ixti` as a
/// direct 0-based index into `sheet_names`, matching pre-ExternSheet-
/// aware behaviour. When an index is out of range for either table the
/// decoder emits `#REF!` for the reference (Excel's own behaviour for a
/// dangling sheet index) rather than failing the whole formula.
///
/// Errors:
///   * `kIoXlsbUnsupportedPtg` — a token outside the supported set
///     (PtgTbl/PtgNameX/PtgMemFunc materialisation/...).
///   * `kIoXlsbRecordTruncated` — a token's payload would overrun
///     `ptgs` or `rgcb`.
///   * `kIoXlsbCorrupt` — the operand stack is unbalanced (too few
///     operands for an operator, or not exactly one value remaining).
///   * `kOutOfMemory` — arena allocation failed.
Expected<parser::AstNode*, Error> decode_ptgs(ByteSpan ptgs, ByteSpan rgcb, Arena& arena,
                                              const std::vector<std::string>& sheet_names,
                                              const std::vector<XlsbName>& name_table,
                                              const std::vector<XlsbSheetRange>& sheet_ranges);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_READER_H_
