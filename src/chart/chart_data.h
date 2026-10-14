//
// `ChartModel` -> `ChartData`: applies Excel's series-resolution rules
// (hidden rows, blanks, errors, text categories) to the cells the chart
// references. Reads the workbook; never changes it.

#ifndef FORMULON_CHART_CHART_DATA_H_
#define FORMULON_CHART_CHART_DATA_H_

#include <cstddef>

#include "chart/chart_model.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Resolves every series of `model` (hosted on `host_sheet`, or
/// `kChartNoHostSheet` for a chartsheet). An unresolvable source is recorded
/// in `ChartData::unresolved` and falls back to the part's cache when it has
/// one.
Expected<ChartData, Error> resolve_chart_data(const Workbook& wb, std::size_t host_sheet, const ChartModel& model);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_DATA_H_
