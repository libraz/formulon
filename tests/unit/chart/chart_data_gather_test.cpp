#include <string>
#include <utility>
#include <vector>

#include "chart/chart_data.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "styles.h"
#include "workbook.h"

namespace formulon {
namespace {

ChartDataSource Ref(const std::string& formula) {
  ChartDataSource source;
  source.kind = ChartDataSource::Kind::kRef;
  source.ref_element = ChartDataSource::RefElement::kNumRef;
  source.formula = formula;
  return source;
}

ChartModel OneSeries(const std::string& values, const std::string& categories = {}) {
  ChartModel model;
  ChartPlot plot;
  plot.type = ChartPlotType::kBar;
  ChartSeries series;
  series.val = Ref(values);
  if (!categories.empty()) {
    series.cat = Ref(categories);
  }
  plot.series.push_back(std::move(series));
  model.plots.push_back(std::move(plot));
  return model;
}

ChartData Resolve(const Workbook& wb, const ChartModel& model) {
  auto data = resolve_chart_data(wb, 0, model);
  EXPECT_TRUE(static_cast<bool>(data)) << (data ? "" : data.error().message);
  return data ? std::move(data.value()) : ChartData{};
}

void SetPercentFormat(Workbook& wb, std::uint32_t row, std::uint32_t col) {
  CellXf percent;
  percent.num_fmt_id = 10;
  percent.apply_number_format = true;
  wb.mutable_styles().cell_xfs.push_back(percent);
  wb.sheet(0).set_cell_xf_index(row, col, static_cast<std::uint32_t>(wb.styles().cell_xfs.size() - 1U));
}

std::vector<double> PointValues(const ChartDataSeries& series) {
  std::vector<double> values;
  for (const ChartDataPoint& point : series.points) {
    values.push_back(point.value);
  }
  return values;
}

TEST(ChartDataGather, SecondUnfilteredPointSetsFormatEvenWhenHidden) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(2, 1, Value::number(2));
  sheet.set_cell_value(1, 2, Value::number(3));
  sheet.set_cell_value(2, 2, Value::number(4));
  SetPercentFormat(wb, 2, 1);
  RowLayout hidden_row;
  hidden_row.row = 2;
  hidden_row.hidden = true;
  sheet.mutable_layout().row_overrides.push_back(hidden_row);

  ChartModel model = OneSeries("Sheet1!$B$2:$C$3");
  model.plot_vis_only = Tri::kTrue;
  const ChartData data = Resolve(wb, model);

  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(PointValues(data.series[0]), (std::vector<double>{1, 3}));
  EXPECT_EQ(data.series[0].format_code, "0.00%");
}

TEST(ChartDataGather, SecondPointAcrossAreasSetsFormatBeforeHiddenFiltering) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(1, 2, Value::number(3));
  SetPercentFormat(wb, 1, 2);
  ColumnLayout hidden_column;
  hidden_column.first = 2;
  hidden_column.last = 2;
  hidden_column.hidden = true;
  sheet.mutable_layout().columns.push_back(hidden_column);

  ChartModel model = OneSeries("(Sheet1!$B$2,Sheet1!$C$2)");
  model.plot_vis_only = Tri::kTrue;
  const ChartData data = Resolve(wb, model);

  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(PointValues(data.series[0]), (std::vector<double>{1}));
  EXPECT_EQ(data.series[0].format_code, "0.00%");
}

TEST(ChartDataGather, OverlappingAreasKeepDuplicatePointsInSourceOrder) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(1, 1, Value::number(1));
  sheet.set_cell_value(1, 2, Value::number(2));
  sheet.set_cell_value(1, 3, Value::number(3));
  SetPercentFormat(wb, 1, 2);

  const ChartData data = Resolve(wb, OneSeries("(Sheet1!$B$2:$C$2,Sheet1!$C$2:$D$2)"));

  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(PointValues(data.series[0]), (std::vector<double>{1, 2, 2, 3}));
  EXPECT_EQ(data.series[0].format_code, "0.00%");
}

TEST(ChartDataGather, LonePointUsesItsFormatAndEmptyTrimmedRangeHasNone) {
  Workbook wb = Workbook::create();
  wb.sheet(0).set_cell_value(1, 1, Value::number(0.5));
  SetPercentFormat(wb, 1, 1);

  const ChartData one = Resolve(wb, OneSeries("Sheet1!$B$2"));
  ASSERT_EQ(one.series.size(), 1U);
  EXPECT_EQ(PointValues(one.series[0]), (std::vector<double>{0.5}));
  EXPECT_EQ(one.series[0].format_code, "0.00%");

  const Workbook empty_wb = Workbook::create();
  const ChartData empty = Resolve(empty_wb, OneSeries("Sheet1!$A:$A"));
  ASSERT_EQ(empty.series.size(), 1U);
  EXPECT_TRUE(empty.series[0].points.empty());
  EXPECT_TRUE(empty.series[0].format_code.empty());
}

TEST(ChartDataGather, MultiLevelFormatsUseSecondLowestPointInBothOrientations) {
  {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    sheet.set_cell_text(0, 0, "A");
    sheet.set_cell_text(1, 0, "B");
    sheet.set_cell_text(2, 0, "C");
    sheet.set_cell_value(0, 1, Value::number(1));
    sheet.set_cell_value(1, 1, Value::number(2));
    sheet.set_cell_value(2, 1, Value::number(3));
    sheet.set_cell_value(0, 3, Value::number(1));
    sheet.set_cell_value(1, 3, Value::number(2));
    sheet.set_cell_value(2, 3, Value::number(3));
    SetPercentFormat(wb, 1, 1);

    const ChartData vertical = Resolve(wb, OneSeries("Sheet1!$D$1:$D$3", "Sheet1!$A$1:$B$3"));
    EXPECT_EQ(vertical.categories.labels, (std::vector<std::string>{"1", "200.00%", "3"}));
    ASSERT_EQ(vertical.categories.levels.size(), 1U);
    EXPECT_EQ(vertical.categories.levels[0], (std::vector<std::string>{"A", "B", "C"}));
    EXPECT_EQ(vertical.categories.format_code, "0.00%");
  }

  {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    for (std::uint32_t col = 0; col < 3; ++col) {
      sheet.set_cell_text(0, col, std::string("G") + std::to_string(col));
      sheet.set_cell_value(1, col, Value::number(static_cast<double>(col + 1U)));
    }
    sheet.set_cell_value(0, 3, Value::number(1));
    sheet.set_cell_value(1, 3, Value::number(2));
    SetPercentFormat(wb, 1, 1);

    const ChartData horizontal = Resolve(wb, OneSeries("Sheet1!$D$1:$D$2", "Sheet1!$A$1:$C$2"));
    EXPECT_EQ(horizontal.categories.labels, (std::vector<std::string>{"1", "200.00%", "3"}));
    ASSERT_EQ(horizontal.categories.levels.size(), 1U);
    EXPECT_EQ(horizontal.categories.levels[0], (std::vector<std::string>{"G0", "G1", "G2"}));
    EXPECT_EQ(horizontal.categories.format_code, "0.00%");
  }
}

}  // namespace
}  // namespace formulon
