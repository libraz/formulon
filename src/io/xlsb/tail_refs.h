//
// Coordinate rewriting inside a sheet's retained `.xlsb` tail records for
// row / column structural edits -- the binary counterpart of
// `io/ext_lst_refs.h`.
//
// The future records that carry cell coordinates the model does not own
// share one header: u32 flags (0x02
// range list, 0x04 formula list), then the range list (u32 count; per range
// u32 flags and an Sqrfx), then the formula list (u32 count; per formula
// u32 flags, `cce`, `cb`, `rgce`, `rgcb`). Ranges live on
// BrtBeginConditionalFormatting14 (1046) and BrtSparkline (1043); formulas
// on BrtSparkline, BrtBeginCFRule14 (1048) and BrtCFVO14 (1050). A CF
// formula's PtgRefN / PtgAreaN are relative to its block's range top-left,
// as in a legacy rule. Records of any other shape are left as they are.

#ifndef FORMULON_IO_XLSB_TAIL_REFS_H_
#define FORMULON_IO_XLSB_TAIL_REFS_H_

#include <cstddef>
#include <cstdint>
#include <optional>

#include "io/ext_lst_refs.h"
#include "io/worksheet_child_refs.h"
#include "parser/ast_shift.h"
#include "sheet_passthrough.h"

namespace formulon::io::xlsb {

/// The cbFmla Excel stores beside a BrtBeginCFRule14 formula: `cce` plus 2
/// per PtgRefN / PtgAreaN (3-D and absolute tokens add nothing). Nothing
/// when `rgce` holds a token this walk cannot step over.
std::optional<std::uint32_t> frt_formula_size_field(const std::uint8_t* rgce, std::size_t cce);

/// The cbFmla Excel stores beside a legacy BrtBeginCFRule / BrtCFVO formula: `cce` plus 2 per PtgRef / PtgArea /
/// PtgRefN / PtgAreaN, absolute or relative, plus 4 per PtgName. Nothing when `rgce` holds a token this walk cannot
/// step over.
std::optional<std::uint32_t> legacy_cf_formula_size_field(const std::uint8_t* rgce, std::size_t cce);

/// Maps the range of every retained x14 CF block and sparkline through
/// `remap`. An emptied range drops the block or sparkline, then a
/// sparkline group left without sparklines, then a container left empty.
/// A non-null `cf_remap` maps each x14 CF block's range instead of `remap`.
void remap_tail_sqrefs(XlsbSheetTail& tail, const SqrefRemap& remap, const SqrefRemap* cf_remap = nullptr);

/// Moves the coordinates of the other retained records that name cells for
/// `edit`: the AutoFilter block and the manual breaks by the rules the model
/// applies to its own copies (`shift_auto_filter`, the break shift), and the
/// records behind the children `remap_worksheet_children` covers --
/// BrtRangeProtection, the scenario and sort-state blocks, BrtCellWatch,
/// BrtCellIgnoreEC and a range-sourced web-publish item -- by its rules.
/// What an edit empties is dropped with its container. Under a band edit only
/// what lies inside the band moves, and the breaks stay.
void remap_tail_child_refs(XlsbSheetTail& tail, const StructuralEdit& edit);

/// Maps every reference in the retained sparkline, x14 CF rule and x14
/// threshold formulas through `transform`, naming 3-D sheets through
/// `tail.extern_sheets` and re-anchoring relative ones to their block's
/// moved range. A reference the transform removes becomes the error form of
/// the same Ptg, as a cell formula's does. A formula holding a token this
/// walk cannot step over is left as it is.
void remap_tail_formulas(XlsbSheetTail& tail, const parser::RefTransform& transform);

/// Appends the copies `clone` describes after each retained BrtSparkline
/// hosted next to an inserted band, translating the source formula as the
/// `.xlsx` path does. A sparkline whose formula this walk cannot translate is
/// not copied.
void clone_tail_sparklines(XlsbSheetTail& tail, const SparklineClone& clone);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_TAIL_REFS_H_
