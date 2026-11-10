//
// Where a pivot cache's worksheet source lies in its workbook now.
//
// A source names a table, a defined name, or a sheet plus an A1 range. A
// table resolves to its current range, header row through the last data row
// (the totals row is not part of it); a defined name to its current
// definition, which must be one static rectangle on one sheet. A range is
// taken as written and never grows. Anything else — no source, a source in
// another workbook or of another kind, a missing sheet, table or name, a
// `#REF!` or computed name body — does not resolve.

#ifndef FORMULON_PIVOT_PIVOT_SOURCE_H_
#define FORMULON_PIVOT_PIVOT_SOURCE_H_

#include <cstddef>

#include "merge_range.h"
#include "pivot/pivot_cache.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Workbook;
}  // namespace formulon

namespace formulon::pivot {

/// A resolved source: the sheet and the rectangle, header row first.
struct WorksheetSourceRange {
  std::size_t sheet = 0;
  MergeRange rect;
};

/// Resolves `cache`'s worksheet source in `wb`. `kPreconditionFailed` when it
/// does not resolve.
Expected<WorksheetSourceRange, Error> resolve_worksheet_source(const Workbook& wb, const PivotCache& cache);

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_SOURCE_H_
