//
// Dynamic-array formulas in the file formats. `Cell::dynamic_array` is the
// single source: set at entry (`Workbook::set_cell_formula`) and at load
// from the file's own mark (`cm=` naming the XLDAPR cell metadata in .xlsx,
// a preceding `BrtCellMeta` in .xlsb). A writer stores a marked formula in
// the dynamic-array form (`cm=` + `<f t="array">`, `BrtCellMeta` +
// `BrtArrFmla`); an unmarked one keeps its legacy meaning, so a loaded
// implicit-intersection formula or fixed-size CSE block is not upgraded.

#ifndef FORMULON_IO_DYNAMIC_ARRAY_FORMULA_H_
#define FORMULON_IO_DYNAMIC_ARRAY_FORMULA_H_

#include <cstdint>
#include <unordered_set>
#include <vector>

#include "cell.h"
#include "io/xlsb/ptg_writer.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// True when `cell` is stored as a dynamic-array formula.
bool is_dynamic_array_formula(const Cell& cell);

/// True when any formula cell of `wb` is stored as a dynamic-array formula.
bool has_dynamic_array_formula(const Workbook& wb);

/// The 1-based `<cellMetadata>/<bk>` index of `xl/metadata.xml`'s XLDAPR
/// (dynamic-array) entry, which a dynamic-array cell's `cm=` names; 0 when
/// the part is malformed or has none.
std::uint32_t xldapr_cell_metadata_index(const std::vector<std::uint8_t>& metadata_xml);

/// Key of `(row, col)` in the set `apply_loaded_dynamic_array_marks` takes.
inline std::uint64_t dynamic_array_cell_key(std::uint32_t row, std::uint32_t col) {
  return (static_cast<std::uint64_t>(row) << 32) | col;
}

/// Sets every formula cell of `sheet` to be dynamic exactly when the file
/// marked it (`marked` holds `dynamic_array_cell_key`s), replacing what
/// formula entry inferred: a loaded formula keeps the meaning the file gives.
void apply_loaded_dynamic_array_marks(Sheet& sheet, const std::unordered_set<std::uint64_t>& marked);

/// `xlsb::NameShape` of every defined name of `wb`, in `wb.defined_names()` order.
std::vector<xlsb::NameShape> defined_name_shapes(const Workbook& wb);

/// What a formula on sheet `sheet_index` of `wb` knows of each defined name
/// it references (sheet-local before workbook-wide), from the name's own
/// formula: whether it is one cell or a constant, the rectangle it covers,
/// and the positive integer it is.
xlsb::NameShapes name_shapes(const Workbook& wb, std::size_t sheet_index);

/// True when the formula in `cell` at (`row`, `col`) of `sheet` is stored as
/// recalculated every time (`.xlsx` `ca="1"`, `.xlsb` fAlwaysCalc), as Excel
/// 365 stores it: `xlsb::formula_always_calculates` holds for it, or it is a
/// dynamic-array formula whose spill is blocked. False for a non-formula cell.
bool formula_cell_always_calculates(const Sheet& sheet, std::uint32_t row, std::uint32_t col, const Cell& cell,
                                    const xlsb::NameShapes& names);

/// True when Excel 365 marks `root`, a formula typed into sheet
/// `sheet_index` of `wb`, as a dynamic-array formula
/// (`xlsb::formula_is_dynamic_array` over `wb`'s defined names).
bool entered_as_dynamic_array(const Workbook& wb, std::size_t sheet_index, const parser::AstNode& root);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DYNAMIC_ARRAY_FORMULA_H_
