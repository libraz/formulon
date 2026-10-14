//
// Chart part DOM -> `ChartModel`. Handles `c:` charts and `cx:` chartex parts;
// elements are matched by local name. The DOM is never rewritten here.

#ifndef FORMULON_CHART_CHART_XML_H_
#define FORMULON_CHART_CHART_XML_H_

#include <cstdint>
#include <vector>

#include "chart/chart_model.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Parses chart part bytes. Fails with `kIoChartUnparseable` when they are not
/// well-formed XML or the root is neither `c:chartSpace` nor `cx:chartSpace`.
Expected<void, Error> parse_chart_part(const std::vector<std::uint8_t>& bytes, pugi::xml_document& doc);

/// Builds the typed view of a parsed chart part. `wb` supplies colour
/// resolution and the date system. Identity fields (`object_id`, `name`,
/// `descr`, anchor, `on_chartsheet`) are filled by the caller from the
/// owning `graphicFrame`; `source` is derived through `chart_template`.
Expected<ChartModel, Error> read_chart_model(const Workbook& wb, const pugi::xml_document& doc);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_XML_H_
