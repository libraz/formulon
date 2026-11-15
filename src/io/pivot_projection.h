//
// The projected report of one pivot table, which both workbook writers save:
// the grid it occupies and, for a built or edited table, the row and column
// line items of that report.

#ifndef FORMULON_IO_PIVOT_PROJECTION_H_
#define FORMULON_IO_PIVOT_PROJECTION_H_

#include <optional>

#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "pivot/pivot_table.h"
#include "workbook.h"

namespace formulon::io {

/// What a save writes for one pivot table from its projected report.
struct ProjectedPivot {
  /// The grid the report occupies; nullopt when it cannot be established.
  std::optional<PivotRenderedSpan> span;
  /// `<rowItems>` / `<colItems>` generated for an edited table.
  std::optional<PivotItemsXml> items;
  /// The table was built in memory or edited, so no items read with it describe it.
  bool needs_items = false;
};

/// Projects `table`: the grid it will occupy, for its location, and for an
/// edited table the line items of that report.
///
/// The saved location has to describe the report Excel draws, and only the
/// layout layer knows how many rows and columns that is. This is the one
/// place a writer reaches past the pivot model into evaluation, which leaves
/// the model-to-bytes writers pure.
///
/// The layout options bear on the *extent*, not merely on label text:
/// `pivot::layout` selects Excel's compact shape only when the profile
/// supplies axis-label captions, and the compact and legacy shapes differ in
/// header-row and grand-total-strip counts. Projecting under the workbook's
/// own profile is what makes the saved location agree with the grid Excel
/// renders.
///
/// Lines are generated only for a built or edited table whose projection is not
/// pending, from the same projection while its result is alive; a refused
/// shape or an item the written items lack leaves them out. `plan` is the write
/// plan of the table's cache, whose items the lines address (null for a writer
/// that saves each field's items as modelled), and the caller holds the
/// workbook's `ScopedEvalProfile`.
///
/// A failed evaluation or projection leaves the span and the lines unset:
/// saving then falls back to the table's own span, and a workbook must not
/// fail to save because one pivot could not be evaluated.
ProjectedPivot project_pivot(const Workbook& wb, const pivot::PivotTable& table, const PivotCacheWritePlan* plan);

}  // namespace formulon::io

#endif  // FORMULON_IO_PIVOT_PROJECTION_H_
