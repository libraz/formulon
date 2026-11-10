// Totals-row cell <-> column metadata mapping.
//
// A totals cell is the source of truth: its content decides the column's
// totals label, built-in function or custom formula. One mapping serves the
// API write path and the post-load normalization of both readers.

#ifndef FORMULON_TABLE_TOTALS_SYNC_H_
#define FORMULON_TABLE_TOTALS_SYNC_H_

#include <cstddef>
#include <string>

#include "cell.h"
#include "eval/recalc_engine.h"
#include "table.h"

namespace formulon {

class Workbook;

/// What a totals-row cell stands for.
struct TotalsCellClass {
  enum class Kind {
    kNone,      ///< Blank cell.
    kLabel,     ///< A value; `text` is its General-format text.
    kFunction,  ///< The engine-generated SUBTOTAL of a built-in function; `text` is its `totalsRowFunction` token.
    kCustom,    ///< Any other formula; `text` is the stored text without `=`.
  };
  Kind kind = Kind::kNone;
  std::string text;
  /// Whether a `kCustom` formula is an array formula.
  bool array = false;
};

/// Classifies `cell`, the totals-row cell of column `column` of `table`.
/// A formula is a built-in function exactly when its parsed, re-formatted text
/// equals that of `=SUBTOTAL(c,<table>[<column>])` for one of the eight built-in
/// function codes.
TotalsCellClass classify_totals_cell(const Workbook& wb, const TableMetadata& table, std::size_t column,
                                     const Cell& cell);

/// Classifies the totals cell of column `column` of table `table_index` and
/// writes the result into the column's totals metadata, clearing the other
/// fields. A label cell holding a number, bool or error is rewritten as a
/// string. Does nothing when the table shows no totals row or an index is out
/// of range. The caller holds the engine mutex and passes its mutator.
void apply_totals_cell_locked(Workbook& wb, const eval::RecalcEngine::LockedMutator& mutator, std::size_t table_index,
                              std::size_t column);

/// Applies `apply_totals_cell_locked` to every column of every table that
/// shows a totals row and stores a cell in it; a column with no stored totals
/// cell keeps the settings the file declares.
void normalize_table_totals_cells_locked(Workbook& wb, const eval::RecalcEngine::LockedMutator& mutator);

}  // namespace formulon

#endif  // FORMULON_TABLE_TOTALS_SYNC_H_
