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

/// Which names a formula on sheet `sheet_index` of `wb` reads as one value
/// (`xlsb::legacy_intersections`): a defined name, sheet-local before
/// workbook-wide, whose body is a one-cell reference or a constant.
xlsb::NameIsScalar legacy_name_shapes(const Workbook& wb, std::size_t sheet_index);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DYNAMIC_ARRAY_FORMULA_H_
