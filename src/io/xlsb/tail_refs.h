//
// Coordinate rewriting inside a sheet's retained `.xlsb` tail records for
// row / column structural edits -- the binary counterpart of
// `io/ext_lst_refs.h`.
//
// The future records that carry cell coordinates the model does not own
// share one header, measured against Excel's .xlsx twins: u32 flags (0x02
// range list, 0x04 formula list), then the range list (u32 count; per range
// u32 flags and an Sqrfx), then the formula list (u32 count; per formula
// u32 flags, `cce`, `cb`, `rgce`, `rgcb`). Ranges live on
// BrtBeginConditionalFormatting14 (1046) and BrtSparkline (1043); formulas
// on BrtSparkline, BrtBeginCFRule14 (1048) and BrtCFVO14 (1050). A CF
// formula's PtgRefN / PtgAreaN are relative to its block's range top-left,
// as in a legacy rule. Records of any other shape are left as they are.

#ifndef FORMULON_IO_XLSB_TAIL_REFS_H_
#define FORMULON_IO_XLSB_TAIL_REFS_H_

#include "io/ext_lst_refs.h"
#include "parser/ast_shift.h"
#include "sheet_passthrough.h"

namespace formulon::io::xlsb {

/// Maps the range of every retained x14 CF block and sparkline through
/// `remap`. An emptied range drops the block or sparkline, then a
/// sparkline group left without sparklines, then a container left empty.
void remap_tail_sqrefs(XlsbSheetTail& tail, const SqrefRemap& remap);

/// Maps every reference in the retained sparkline, x14 CF rule and x14
/// threshold formulas through `transform`, naming 3-D sheets through
/// `tail.extern_sheets` and re-anchoring relative ones to their block's
/// moved range. A reference the transform removes becomes the error form of
/// the same Ptg, as a cell formula's does. A formula holding a token this
/// walk cannot step over is left as it is.
void remap_tail_formulas(XlsbSheetTail& tail, const parser::RefTransform& transform);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_TAIL_REFS_H_
