//
// Chart reads, creation and typed edits, addressed by `(sheet, object_id)`
// as for other drawing objects. Edits patch the chart part's DOM and write
// nothing when the result is unchanged. Chartex charts refuse edits with
// `kIoChartUnsupportedKind`; a part that does not parse fails with
// `kIoChartUnparseable`. A chart on a drawing retained from xlsb is read but
// not edited (`kIoDrawingUnparseable`), and a pivot chart refuses the series
// and source edits (`kIoChartUnsupportedKind`). New elements go at their
// schema position, in the forms Excel writes; reference edits refresh the
// part's caches. Unknown ids, plots and series are `kInvalidArgument`.

#ifndef FORMULON_CHART_CHART_EDIT_H_
#define FORMULON_CHART_CHART_EDIT_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "chart/chart_model.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

Expected<ChartModel, Error> get_chart(const Workbook& wb, std::size_t sheet, std::uint32_t object_id);

Expected<ChartData, Error> get_chart_data(const Workbook& wb, std::size_t sheet, std::uint32_t object_id);

/// `get_chart_data` for a copy of the chart with `overrides` applied by the
/// edit rules; `wb` is not changed. Source, direction and header overrides
/// re-split as `set_chart_source` does: over the overridden range (unset
/// direction / headers auto), else over the derived range (unset direction /
/// headers as derived).
Expected<ChartData, Error> preview_chart_data(const Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                              const ChartPreviewOverrides& overrides);

/// Creates a chart with filled caches and returns its `cNvPr` id. A zero
/// width or height takes Excel's default 360 x 216 pt frame.
Expected<std::uint32_t, Error> insert_chart(Workbook& wb, std::size_t sheet, const ChartInsertOptions& options);

/// Replaces plot `plot` with `kind` / `grouping` from the insert template,
/// keeping each series' idx, order, name, references, unique id and explicit
/// RGB colour. Axes stay when the new kind uses the same axis types.
Expected<void, Error> set_chart_type(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                     ChartInsertKind kind, ChartGrouping grouping);

Expected<void, Error> set_chart_title(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                      const ChartTitle& title);

Expected<void, Error> set_chart_legend(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const ChartLegendPatch& legend);

Expected<void, Error> set_chart_axis(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::uint32_t axis_id,
                                     const ChartAxisPatch& patch);

/// Writes both elements; `kAbsent` removes the element.
Expected<void, Error> set_chart_display(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                        ChartBlanksAs blanks_as, Tri plot_vis_only);

/// `series_idx` is the series' `c:idx`.
Expected<void, Error> set_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                       std::uint32_t series_idx, const ChartSeriesPatch& patch);

Expected<void, Error> set_chart_point(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                      std::uint32_t series_idx, std::uint32_t point_idx, const ChartPointPatch& patch);

Expected<void, Error> set_chart_plot(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                     const ChartPlotPatch& patch);

/// Hides series (moved to `c15:filtered*Series`) and categories (positions of
/// the full category range), replacing the current filter.
Expected<void, Error> set_chart_filter(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const std::vector<std::uint32_t>& hidden_series_idx,
                                       const std::vector<std::uint32_t>& hidden_category_idx);

/// Returns the new series' `c:idx`.
Expected<std::uint32_t, Error> add_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                                std::size_t plot, const ChartSeriesPatch& patch);

Expected<void, Error> remove_chart_series(Workbook& wb, std::size_t sheet, std::uint32_t object_id, std::size_t plot,
                                          std::uint32_t series_idx);

/// Rebuilds every series from `source` by the insert splitting rules,
/// keeping chart kind and per-series formatting by idx; clears both filters.
Expected<void, Error> set_chart_source(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                       const std::string& source, ChartSeriesDirection direction,
                                       ChartHeaderMode header_row, ChartHeaderMode header_column);

}  // namespace formulon

#endif  // FORMULON_CHART_CHART_EDIT_H_
