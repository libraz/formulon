//
// Reference formula -> cell rectangles. The formula is resolved to a set of
// areas, never evaluated.

#ifndef FORMULON_CHART_CHART_SOURCE_H_
#define FORMULON_CHART_CHART_SOURCE_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_model.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Sheet;
class Workbook;

namespace eval {
class FunctionRegistry;
}

struct ResolvedSource {
  std::vector<SheetRect> areas;  ///< In point order; a parenthesised union keeps its element order.
};

/// Resolves a series / title reference to rectangles. Unqualified references
/// and bare names belong to `host_sheet`; `kChartNoHostSheet` rejects the
/// former and sees workbook-scoped names only. Names (`[0]!N`, `Sheet!N`,
/// hidden `_xlchart.v1.N`) resolve through their bodies, so a union, spill
/// or `OFFSET` body works. Whole columns / rows keep their full extent.
/// Fails with `kIoChartInvalidReference` for unparseable text, `#REF!`,
/// external or non-reference results.
Expected<ResolvedSource, Error> resolve_chart_source(const Workbook& wb, std::size_t host_sheet,
                                                     std::string_view formula, const eval::FunctionRegistry& registry);

/// Whole-column / whole-row rectangles stop at the sheet's populated extent; an
/// empty extent yields an inverted rectangle (`top > bottom`).
SheetRect trim_to_populated(const Workbook& wb, SheetRect r);

/// Whether `row` / `col` is hidden by the sheet's layout.
bool row_hidden(const Sheet& sheet, std::uint32_t row);
bool col_hidden(const Sheet& sheet, std::uint32_t col);

/// Decimal text of a point count or index.
std::string count_text(std::uint64_t n);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_SOURCE_H_
