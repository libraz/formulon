//
// Series splitting and XML generation for new charts, and the inverse
// (deriving a chart's data range from its series references).

#ifndef FORMULON_CHART_CHART_TEMPLATE_H_
#define FORMULON_CHART_CHART_TEMPLATE_H_

#include <cstddef>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "chart/chart_model.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// One series produced by splitting a source rectangle (reference formulas,
/// empty when the series has none).
struct ChartSeriesRefs {
  std::string name;
  std::string cat;  ///< Scatter: unused (see `x`).
  std::string val;
  std::string x;  ///< Scatter only.
  std::string y;  ///< Scatter only.
};

struct ChartSplit {
  std::vector<ChartSeriesRefs> series;
  ChartSeriesDirection direction = ChartSeriesDirection::kColumns;  ///< Resolved, never kAuto.
  bool header_row = false;
  bool header_column = false;
};

/// Splits `source` (a single rectangle; unqualified means `host_sheet`) into
/// series for `kind`. Fails with `kIoChartInvalidReference` for a bad source.
Expected<ChartSplit, Error> split_chart_source(const Workbook& wb, std::size_t host_sheet, std::string_view source,
                                               ChartInsertKind kind, ChartSeriesDirection direction,
                                               ChartHeaderMode header_row, ChartHeaderMode header_column);

/// Generates the chart part XML for `options` over an already split source,
/// caches filled from the cells, with automatic colours from colour style
/// `color_style` (a `cs:colorStyle` element).
Expected<std::string, Error> build_chart_xml(const Workbook& wb, const ChartInsertOptions& options,
                                             const ChartSplit& split, const pugi::xml_node& color_style);

/// `build_chart_xml` for a new chart, in the colour style Excel gives one.
Expected<std::string, Error> build_chart_xml(const Workbook& wb, const ChartInsertOptions& options,
                                             const ChartSplit& split);

/// Derives the shared data range of `model`'s series by re-splitting
/// candidate rectangles; unset when the references do not come from one.
std::optional<ChartSource> derive_chart_source(const Workbook& wb, std::size_t host_sheet, const ChartModel& model);

/// The colour style part (`cs:colorStyle`, cycle method) Excel gives a new chart, and falls back to for a
/// styled chart without one.
extern const char* const kDefaultChartColorStyle;

/// The automatic colour that colour style `style` gives series (or pie point) `idx`, as `a:` colour XML: base
/// colour `idx % n` carrying the children of variation `(idx / n) % V`. Empty when `style` has no base colour.
std::string chart_cycle_color(const pugi::xml_node& style, std::size_t idx);

/// Absolute A1 reference to a rectangle on sheet `sheet_name`, in the form
/// Excel writes into `c:f` (`Sheet1!$A$2:$A$5`, `'My Data'!$B$1`).
std::string chart_rect_ref(const std::string& sheet_name, std::uint32_t top, std::uint32_t left, std::uint32_t bottom,
                           std::uint32_t right);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_TEMPLATE_H_
