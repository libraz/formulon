//
// Refresh of chart caches (`c:numCache` / `c:strCache` / `c:multiLvlStrCache`)
// from the cells their references name, in the form Excel writes on save.

#ifndef FORMULON_CHART_CHART_CACHE_H_
#define FORMULON_CHART_CHART_CACHE_H_

#include <cstddef>
#include <vector>

#include "passthrough_part.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Chart parts whose caches differ from the cells they reference, with the
/// updated bytes. Parts already in sync, and refs that are `#REF!`, external
/// or unparseable, are left out; so are pivot charts, chartex parts and parts
/// that do not parse. Does not change `wb`.
Expected<std::vector<PassthroughPart>, Error> refreshed_chart_parts(const Workbook& wb);

/// Rewrites, in `doc`, the cache of every series reference (`c:tx`, `c:cat`,
/// `c:val`, `c:xVal`, `c:yVal`, `c:bubbleSize`, filtered series included)
/// that differs from its cells. `host_sheet` is the sheet holding the chart,
/// or `kChartNoHostSheet` for a chartsheet. Returns whether `doc` changed;
/// a pivot chart or chartex `doc` is never changed.
bool refresh_chart_caches(const Workbook& wb, std::size_t host_sheet, pugi::xml_document& doc);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_CACHE_H_
