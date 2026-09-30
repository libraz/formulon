//
// Rewrites every formula-bearing holder in a workbook through one
// `parser::RefTransform`: cell formulas, defined names, and the sheet-attached
// and workbook-level metadata that carries a formula or a sheet reference.
// Shared by the sheet rename / remove and row / column edit paths of
// `Workbook`.

#ifndef FORMULON_WORKBOOK_REF_REWRITE_H_
#define FORMULON_WORKBOOK_REF_REWRITE_H_

#include <cstdint>
#include <memory>
#include <string_view>
#include <vector>

#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "parser/ast_shift.h"
#include "pivot/pivot_cache.h"
#include "sheet.h"
#include "table.h"

namespace formulon {

// Applies a per-sheet AST transform to every non-cell formula holder:
// conditional-format rules and their formula thresholds, hyperlink locations,
// data validations, table refs and calculated columns, and pivot-cache
// worksheet sources.
//
// `per_sheet[i]` governs the metadata owned by `sheets[i]`. A row/column edit
// needs `local_means_target` enabled only on the edited sheet — an unqualified
// reference on any other sheet names that other sheet and must be left alone —
// while a rename or removal passes the same transform in every slot.
//
// Tables and pivot caches hang off the workbook rather than off a `Sheet`, so
// each resolves its own owner: a table through its `sheet_index`, a cache
// through the sheet name recorded in its worksheet source. An owner that does
// not resolve falls back to the first slot, which is the whole table whenever
// the transform is uniform.
//
// `direct_sheet_old/new` covers OOXML's worksheetSource/@sheet field, which is
// a plain metadata string rather than a formula AST. For removal,
// `removed_sheet_name` additionally reports cache ids whose real worksheet
// source became unavailable.
//
// Callers pass one slot per sheet. A short `per_sheet` degrades the excess
// sheets to the first slot rather than skipping the rewrite, so no exit path
// leaves a formula holder describing pre-edit coordinates.
void rewrite_sheet_metadata_formulas(std::vector<Sheet>& sheets,
                                     const std::vector<const parser::RefTransform*>& per_sheet,
                                     std::vector<TableMetadata>& tables,
                                     std::vector<std::unique_ptr<pivot::PivotCache>>& pivot_caches,
                                     std::string_view direct_sheet_old, std::string_view direct_sheet_new,
                                     std::string_view removed_sheet_name,
                                     std::vector<std::uint32_t>& dropped_cache_ids);

// Applies one AST transform to every formula-bearing holder in the workbook.
// The caller supplies the transform policy, so rename and removal share the
// exact same parse/identity/error behaviour. `direct_sheet_old/new` covers
// OOXML's worksheetSource/@sheet field, which is a plain metadata string
// rather than a formula AST. For removal, `removed_sheet_name` additionally
// reports cache ids whose real worksheet source became unavailable.
//
// Precondition: caller holds the engine mutex. Formula writes use
// `Sheet::set_cell_formula` directly and never route through the public
// Workbook setter (which would attempt to acquire this mutex again).
//
// Structural sheet mutations use this one path for cells and every
// formula-bearing metadata holder, so string literals, unresolved names, and
// external references do not accidentally participate in a sheet mutation.
void rewrite_workbook_references(std::vector<Sheet>& sheets, std::vector<DefinedName>& defined_names,
                                 std::vector<TableMetadata>& tables,
                                 std::vector<std::unique_ptr<pivot::PivotCache>>& pivot_caches,
                                 const parser::RefTransform& transform,
                                 const eval::RecalcEngine::LockedMutator& mutator, std::string_view direct_sheet_old,
                                 std::string_view direct_sheet_new, std::string_view removed_sheet_name,
                                 std::vector<std::uint32_t>& dropped_cache_ids, bool& defined_names_changed);

/// Applies `transform` to every defined name's formula. Returns true when at
/// least one definition changed, which means every formula that references a
/// name now resolves to a different range than the dep graph was built from.
bool rewrite_defined_names(std::vector<DefinedName>& names, const parser::RefTransform& transform);

// Moves the ranges inside `sheet`'s retained worksheet extensions -- the
// raw `<extLst>` and the `.xlsb` tail records -- for a row/column edit, by
// the rule `shift_sqref_ranges` applies to the modelled sqrefs. The sheet
// cannot do this in its own move: finding those ranges takes the
// file-format readers. Formulas in the same content move with
// `rewrite_sheet_metadata_formulas` against pre-edit coordinates, so call
// this after the sheet's own move.
void shift_retained_extension_ranges(Sheet& sheet, std::uint32_t index, std::uint32_t count, bool is_delete,
                                     bool row_axis);

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_REF_REWRITE_H_
