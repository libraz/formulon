#include "chart/chart_data.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <limits>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_xml.h"
#include "defined_name.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "eval/text_format/display_text.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "styles.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

std::vector<std::uint8_t> ReadFileBytes(const std::string& path) {
  std::vector<std::uint8_t> out;
  FILE* file = std::fopen(path.c_str(), "rb");
  if (file == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(file, 0, SEEK_END);
  const long size = std::ftell(file);
  std::fseek(file, 0, SEEK_SET);
  if (size > 0) {
    out.resize(static_cast<std::size_t>(size));
    if (std::fread(out.data(), 1, out.size(), file) != out.size()) {
      ADD_FAILURE() << "short read on fixture: " << path;
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

/// Loads an Excel-saved fixture under the Mac ja-JP profile it was saved with.
Workbook LoadFixture(const std::string& name) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
  Workbook wb = Workbook::create_empty();
  if (!bytes.empty()) {
    const io::ByteSpan span{bytes.data(), bytes.size()};
    if (name.size() > 5 && name.substr(name.size() - 5) == ".xlsb") {
      auto read = io::xlsb::read_xlsb(span);
      EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
      if (read) {
        wb = std::move(read.value().workbook);
      }
    } else {
      auto read = io::read_ooxml(span);
      EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
      if (read) {
        wb = std::move(read.value().workbook);
      }
    }
  }
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

struct LoadedChart {
  std::string name;
  std::size_t host = 0;
  ChartModel model;
};

/// Every chart of the workbook with its drawing-object name and host sheet.
std::vector<LoadedChart> LoadCharts(const Workbook& wb) {
  std::vector<LoadedChart> out;
  auto refs = all_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(refs));
  if (!refs) {
    return out;
  }
  for (const SheetChartPartRef& item : refs.value()) {
    LoadedChart chart;
    auto objects = list_drawing_objects(wb, item.sheet);
    if (objects) {
      for (const DrawingObject& obj : objects.value()) {
        if (obj.object_id == item.ref.object_id) {
          chart.name = obj.name;
        }
      }
    }
    chart.host = wb.sheet(item.sheet).is_opaque_ooxml_sheet() ? kChartNoHostSheet : item.sheet;
    const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, item.ref.part_path);
    EXPECT_NE(part, nullptr) << item.ref.part_path;
    if (part == nullptr) {
      continue;
    }
    pugi::xml_document doc;
    EXPECT_TRUE(static_cast<bool>(parse_chart_part(part->bytes, doc)));
    auto model = read_chart_model(wb, doc);
    EXPECT_TRUE(static_cast<bool>(model)) << item.ref.part_path;
    if (!model) {
      continue;
    }
    chart.model = std::move(model.value());
    out.push_back(std::move(chart));
  }
  return out;
}

std::string Num(double value) {
  char buf[40];
  if (value == std::floor(value) && std::fabs(value) < 1e15) {
    std::snprintf(buf, sizeof(buf), "%.0f", value);
  } else {
    std::snprintf(buf, sizeof(buf), "%.17g", value);
  }
  return buf;
}

std::string State(ChartPointState state, double value) {
  switch (state) {
    case ChartPointState::kGap:
      return "gap";
    case ChartPointState::kConnect:
      return "conn";
    case ChartPointState::kValue:
      break;
  }
  return Num(value);
}

std::string Join(const std::vector<std::string>& parts, const char* sep = ",") {
  std::string out;
  for (std::size_t i = 0; i < parts.size(); ++i) {
    out += (i == 0 ? "" : sep) + parts[i];
  }
  return out;
}

/// `name|points` (category domain) or `name|x:y ...` (XY domain).
std::string SeriesText(const ChartDataSeries& s) {
  std::vector<std::string> pts;
  for (const ChartDataPoint& p : s.points) {
    pts.push_back(State(p.state, p.value));
  }
  for (const ChartXYPoint& p : s.xy) {
    pts.push_back(Num(p.x) + ":" + State(p.state, p.y));
  }
  return (s.name ? *s.name : "<null>") + (s.filtered ? "(filtered)" : "") + "|" + Join(pts, " ");
}

/// `cats=...|levels=...|series...`, one line per chart.
std::string DataText(const ChartData& d) {
  std::string out = "cats=" + Join(d.categories.labels);
  if (d.categories.generated) {
    out += "(gen)";
  }
  for (const std::vector<std::string>& level : d.categories.levels) {
    out += "|lvl=" + Join(level);
  }
  for (const ChartDataSeries& s : d.series) {
    out += "|" + SeriesText(s);
  }
  return out;
}

ChartData Resolve(const Workbook& wb, const LoadedChart& chart) {
  auto data = resolve_chart_data(wb, chart.host, chart.model);
  EXPECT_TRUE(static_cast<bool>(data)) << chart.name;
  return data ? std::move(data.value()) : ChartData{};
}

void ExpectFixtureData(const std::string& fixture, const std::vector<std::pair<std::string, std::string>>& expected) {
  const Workbook wb = LoadFixture(fixture);
  const std::vector<LoadedChart> charts = LoadCharts(wb);
  for (const auto& [name, text] : expected) {
    bool found = false;
    for (const LoadedChart& chart : charts) {
      if (chart.name == name) {
        found = true;
        EXPECT_EQ(DataText(Resolve(wb, chart)), text) << fixture << " " << name;
      }
    }
    EXPECT_TRUE(found) << fixture << " " << name;
  }
}

ChartDataSource Ref(const std::string& formula) {
  ChartDataSource src;
  src.kind = ChartDataSource::Kind::kRef;
  src.ref_element = ChartDataSource::RefElement::kNumRef;
  src.formula = formula;
  return src;
}

ChartModel OneSeries(ChartPlotType type, const std::string& cat, const std::string& val) {
  ChartModel model;
  ChartPlot plot;
  plot.type = type;
  ChartSeries s;
  if (!cat.empty()) {
    s.cat = Ref(cat);
  }
  if (type == ChartPlotType::kScatter) {
    s.y = Ref(val);
  } else {
    s.val = Ref(val);
  }
  plot.series.push_back(std::move(s));
  model.plots.push_back(std::move(plot));
  return model;
}

ChartData ResolveSynthetic(const Workbook& wb, const ChartModel& model) {
  auto data = resolve_chart_data(wb, 0, model);
  EXPECT_TRUE(static_cast<bool>(data));
  return data ? std::move(data.value()) : ChartData{};
}

// ---------------------------------------------------------------------------
// Rule interactions (pairwise over chart type x dispBlanksAs x plotVisOnly x
// cell content x hidden row / column x dispNaAsBlank)
// ---------------------------------------------------------------------------

struct RuleCase {
  ChartPlotType type;
  ChartBlanksAs blanks;
  Tri vis_only;
  const char* content;  ///< number / blank / na / div0 / text / bool, at the middle point.
  const char* hidden;   ///< none / row (the middle point's row) / col (the value column).
  Tri na_as_blank;
  const char* expected;  ///< Middle point: a value, gap, conn; or removed / dropped.
};

TEST(ChartData, PairwiseRuleInteractions) {
  using T = ChartPlotType;
  using B = ChartBlanksAs;
  const Tri A = Tri::kAbsent;
  const Tri F = Tri::kFalse;
  const Tri Y = Tri::kTrue;
  const std::vector<RuleCase> cases = {
      {T::kBar, B::kGap, Y, "bool", "none", F, "1"},
      {T::kArea, B::kAbsent, Y, "blank", "col", A, "dropped"},
      {T::kPie, B::kZero, A, "div0", "col", Y, "dropped"},
      {T::kScatter, B::kZero, F, "number", "none", A, "5"},
      {T::kBar, B::kSpan, F, "na", "row", Y, "gap"},
      {T::kRadar, B::kGap, A, "text", "row", A, "removed"},
      {T::kRadar, B::kSpan, F, "div0", "col", F, "0"},
      {T::kPie, B::kSpan, Y, "text", "none", A, "0"},
      {T::kLine, B::kGap, F, "text", "col", Y, "0"},
      {T::kRadar, B::kZero, Y, "bool", "row", Y, "removed"},
      {T::kPie, B::kAbsent, F, "blank", "row", F, "0"},
      {T::kArea, B::kSpan, A, "blank", "none", Y, "gap"},
      {T::kLine, B::kAbsent, A, "na", "none", F, "conn"},
      {T::kRadar, B::kAbsent, Y, "number", "row", Y, "removed"},
      {T::kArea, B::kZero, F, "text", "row", F, "0"},
      {T::kLine, B::kZero, Y, "na", "row", A, "removed"},
      {T::kBar, B::kAbsent, F, "div0", "none", A, "0"},
      {T::kScatter, B::kZero, F, "blank", "row", Y, "0"},
      {T::kScatter, B::kGap, Y, "div0", "row", F, "removed"},
      {T::kBar, B::kGap, A, "number", "col", F, "dropped"},
      {T::kPie, B::kGap, F, "na", "col", F, "gap"},
      {T::kScatter, B::kAbsent, F, "bool", "col", A, "1"},
      {T::kRadar, B::kSpan, A, "bool", "none", A, "1"},
      {T::kArea, B::kGap, F, "div0", "none", Y, "0"},
      {T::kLine, B::kSpan, A, "number", "row", F, "removed"},
      {T::kRadar, B::kGap, Y, "blank", "col", F, "dropped"},
      {T::kRadar, B::kAbsent, Y, "na", "row", Y, "removed"},
      {T::kBar, B::kZero, F, "blank", "none", F, "0"},
      {T::kLine, B::kZero, F, "blank", "none", Y, "0"},
      {T::kLine, B::kZero, F, "bool", "none", Y, "1"},
      {T::kScatter, B::kAbsent, A, "text", "col", Y, "dropped"},
      {T::kPie, B::kAbsent, A, "bool", "col", A, "dropped"},
      {T::kPie, B::kZero, F, "number", "col", Y, "5"},
      {T::kArea, B::kZero, F, "bool", "none", F, "1"},
      {T::kLine, B::kZero, Y, "div0", "col", A, "dropped"},
      {T::kScatter, B::kZero, A, "na", "col", Y, "dropped"},
      {T::kScatter, B::kSpan, Y, "blank", "row", Y, "removed"},
      {T::kBar, B::kAbsent, Y, "text", "row", F, "removed"},
      {T::kArea, B::kZero, F, "na", "none", Y, "0"},
      {T::kArea, B::kZero, F, "number", "row", A, "5"},
      // Hand-added: the connect outcomes the generated rows leave unexercised.
      {T::kLine, B::kSpan, F, "blank", "none", A, "conn"},
      {T::kScatter, B::kGap, F, "na", "none", F, "conn"},
      {T::kRadar, B::kSpan, F, "na", "none", Y, "conn"},
      {T::kLine, B::kGap, F, "na", "none", Y, "gap"},
  };
  for (const RuleCase& c : cases) {
    Workbook wb = Workbook::create();
    Sheet& sheet = wb.sheet(0);
    sheet.set_cell_text(0, 0, "x");
    sheet.set_cell_text(1, 0, "y");
    sheet.set_cell_text(2, 0, "z");
    sheet.set_cell_value(0, 1, Value::number(1));
    sheet.set_cell_value(2, 1, Value::number(3));
    const std::string content = c.content;
    if (content == "number") {
      sheet.set_cell_value(1, 1, Value::number(5));
    } else if (content == "na") {
      sheet.set_cell_value(1, 1, Value::error(ErrorCode::NA));
    } else if (content == "div0") {
      sheet.set_cell_value(1, 1, Value::error(ErrorCode::Div0));
    } else if (content == "text") {
      sheet.set_cell_text(1, 1, "t");
    } else if (content == "bool") {
      sheet.set_cell_value(1, 1, Value::boolean(true));
    }
    const std::string hidden = c.hidden;
    if (hidden == "row") {
      RowLayout row;
      row.row = 1;
      row.hidden = true;
      insert_row_override(sheet.mutable_layout(), row);
    } else if (hidden == "col") {
      ColumnLayout col;
      col.first = 1;
      col.last = 1;
      col.hidden = true;
      sheet.mutable_layout().columns.push_back(col);
    }
    ChartModel model = OneSeries(c.type, c.type == T::kScatter ? "" : "Sheet1!$A$1:$A$3", "Sheet1!$B$1:$B$3");
    model.disp_blanks_as = c.blanks;
    model.plot_vis_only = c.vis_only;
    model.disp_na_as_blank = c.na_as_blank;
    const ChartData data = ResolveSynthetic(wb, model);
    const std::string label = std::to_string(static_cast<int>(c.type)) + "/" +
                              std::to_string(static_cast<int>(c.blanks)) + "/" +
                              std::to_string(static_cast<int>(c.vis_only)) + "/" + content + "/" + hidden + "/" +
                              std::to_string(static_cast<int>(c.na_as_blank));
    const std::string expected = c.expected;
    if (expected == "dropped") {
      EXPECT_TRUE(data.series.empty()) << label;
      continue;
    }
    ASSERT_EQ(data.series.size(), 1U) << label;
    std::vector<std::string> got;
    for (const ChartDataPoint& p : data.series[0].points) {
      got.push_back(State(p.state, p.value));
    }
    for (const ChartXYPoint& p : data.series[0].xy) {
      got.push_back(State(p.state, p.y));
    }
    if (expected == "removed") {
      EXPECT_EQ(Join(got, " "), "1 3") << label;
    } else {
      EXPECT_EQ(Join(got, " "), "1 " + expected + " 3") << label;
    }
  }
}

// ---------------------------------------------------------------------------
// Excel-saved fixtures: the cells must give what Excel's Series.Values gave
// ---------------------------------------------------------------------------

TEST(ChartData, DataFixtureMatchesSeriesValues) {
  ExpectFixtureData("chart_data.xlsx",
                    {
                        // dispNaAsBlank=1 makes #N/A a blank, plotted as a gap.
                        {"mixed_line", "cats=a,b,c,d,e,f,g,h|Mixed|1 gap 3 gap 0 1 0 8"},
                        {"mixed_col", "cats=a,b,c,d,e,f,g,h|Mixed|1 gap 3 gap 0 1 0 8"},
                        {"mixed_scatter", "cats=|Mixed|3:1 1:gap 4:3 1:gap 5:0 9:1 2:0 6:8"},
                        {"hid_vis", "cats=a,b,d,f,g,h|Hid|10 20 40 60 70 80"},
                        {"hid_all", "cats=a,b,c,d,e,f,g,h|Hid|10 20 30 40 50 60 70 80"},
                        // The value column is hidden: Excel drops the series.
                        {"hidcol_vis", "cats=a,b,c,d,e,f,g,h"},
                        {"hidcol_all", "cats=a,b,c,d,e,f,g,h|HidCol|1 2 3 4 5 6 7 8"},
                        {"blank_zero", "cats=a,b,c,d,e,f,g,h|Mixed|1 0 3 0 0 1 0 8"},
                        {"blank_span", "cats=a,b,c,d,e,f,g,h|Mixed|1 conn 3 conn 0 1 0 8"},
                        {"multilvl", "cats=i1,i2,i3,i4,i5,i6,i7,i8|lvl=G1,,,G2,,G3,,|MV|5 6 7 8 9 10 11 12"},
                        // A text X makes the X values 1..n.
                        {"scatter_textx", "cats=|Ys|1:3 2:1 3:4 4:1 5:5 6:9 7:2 8:6"},
                    });
}

TEST(ChartData, FormatFixtureKeepsExactDoubles) {
  const Workbook wb = LoadFixture("chart_data.xlsx");
  for (const LoadedChart& chart : LoadCharts(wb)) {
    if (chart.name != "fmt_col") {
      continue;
    }
    const ChartData data = Resolve(wb, chart);
    ASSERT_EQ(data.series.size(), 1U);
    const std::vector<double> expected = {0.1 + 0.2, 1.0 / 3.0, 1e-20,  123456789012345678.0,
                                          -0.5,      45293.0,   1234.5, 2.0 / 3.0};
    ASSERT_EQ(data.series[0].points.size(), expected.size());
    for (std::size_t i = 0; i < expected.size(); ++i) {
      EXPECT_EQ(data.series[0].points[i].value, expected[i]) << i;
    }
    // The second point's format.
    EXPECT_EQ(data.series[0].format_code, "General");
    return;
  }
  ADD_FAILURE() << "fmt_col not found";
}

TEST(ChartData, CacheFixtureMatchesSeriesValues) {
  ExpectFixtureData("chart_cache.xlsx", {
                                            {"pct_all",
                                             "cats=a,5,,#N/A|Pct|0.10000000000000001 0.20000000000000001 "
                                             "0.29999999999999999 0.40000000000000002"},
                                            {"first_fmt", "cats=a,5,,#N/A|FirstFmt|1.5 2.5 3.5 4.5"},
                                            {"name_num", "cats=a,5,,#N/A|2024|1 2 3 4"},
                                            {"name_blank", "cats=a,5,,#N/A||1 2 3 4"},
                                            {"name_literal", "cats=a,5,,#N/A|Lit|1 2 3 4"},
                                            {"vals_text", "cats=a,5,,#N/A|AllText|0 0 0 0"},
                                            {"date_cat", "cats=2024/1/1,2024/1/2,2024/1/5,2024/2/1|DV|3 1 4 1"},
                                            {"num_cat", "cats=10,20,30,40|NV|2 7 1 8"},
                                            {"no_cat", "cats=1,2,3,4(gen)|NV|2 7 1 8"},
                                            {"lit_vals", "cats=p,q|L|1 2"},
                                        });
}

TEST(ChartData, DateCategoryFormatCodeIsTheLocaleForm) {
  Workbook wb = LoadFixture("chart_cache.xlsx");
  wb.sheet(0).set_cell_xf_index(2, 7, 1);  // built-in 14 on the second category cell (xf 1)
  bool seen = false;
  for (const LoadedChart& chart : LoadCharts(wb)) {
    if (chart.name == "date_cat") {
      seen = true;
      EXPECT_EQ(Resolve(wb, chart).categories.format_code, "yyyy/m/d");  // ja-JP form, not the cached m/d/yy
    }
  }
  EXPECT_TRUE(seen);
}

TEST(ChartData, CacheFixtureCategoryValuesAndFormats) {
  const Workbook wb = LoadFixture("chart_cache.xlsx");
  int seen = 0;
  for (const LoadedChart& chart : LoadCharts(wb)) {
    const ChartData data = Resolve(wb, chart);
    if (chart.name == "catmix") {
      // The date label is the cell's display text (built-in format 14).
      const std::string date = text_format::format_cell_for_display(wb, wb.sheet(0), 6, 0).text;
      EXPECT_EQ(DataText(data), "cats=a,5,,#N/A,TRUE," + date + "|V|1 2 3 4 5 6");
    }
    if (chart.name == "date_cat") {
      ++seen;
      EXPECT_EQ(data.categories.values, (std::vector<double>{45292, 45293, 45296, 45323}));
    } else if (chart.name == "num_cat") {
      ++seen;
      EXPECT_EQ(data.categories.values, (std::vector<double>{10, 20, 30, 40}));
    } else if (chart.name == "catmix") {
      ++seen;
      EXPECT_TRUE(data.categories.values.empty());
    } else if (chart.name == "pct_all") {
      ++seen;
      ASSERT_EQ(data.series.size(), 1U);
      EXPECT_EQ(data.series[0].format_code, "0.0%");
    } else if (chart.name == "first_fmt") {
      ++seen;
      ASSERT_EQ(data.series.size(), 1U);
      EXPECT_EQ(data.series[0].format_code, "General");
    }
  }
  EXPECT_EQ(seen, 5);
}

TEST(ChartData, RefsFixtureMatchesSeriesValues) {
  ExpectFixtureData("chart_refs.xlsx", {
                                           {"wb_name", "cats=a,b,c,d|V1|1 2 3 7"},
                                           {"sheet_name", "cats=a,b,c,d|V2|4 5 6 8"},
                                           {"dyn_name", "cats=1,2,3,4,5,6,7(gen)|V1|1 2 3 7 gap 11 12"},
                                           {"table_ref", "cats=1,2,3(gen)|<null>|1 3 5"},
                                           // Trimmed to the used range (rows 1-9); Excel reports 1048576 values.
                                           {"whole_col", "cats=1,2,3,4,5,6,7,8,9(gen)|V1|0 1 2 3 7 gap 11 12 13"},
                                           {"other_sheet", "cats=p,q,r|OV|10 20 30"},
                                           {"hidden_sheet", "cats=x,y|HV|100 200"},
                                           {"merged_cat", "cats=m1,,m3|V1|11 12 13"},
                                           {"two_d_val", "cats=1,2(gen)|<null>|1 2"},
                                           {"multi_name", "cats=a,b,c,d|V1 V2|1 2 3 7"},
                                           {"union", "cats=a,b,d|V1|1 2 7"},
                                       });
}

TEST(ChartData, SpecialFixtures) {
  ExpectFixtureData("chart_filter.xlsx", {
                                             {"col_clustered", "cats=a,b,c,d|S1|1 2 3 7|S2(filtered)|4 5 6 8"},
                                             {"col_stacked", "cats=a,b,c,d|S1|1 2 3 7|S2|4 5 6 8"},
                                         });
  // A pivot chart names its series from the pivot ("集計"), not from the cell.
  ExpectFixtureData("chart_pivot.xlsx", {{"pivot_chart", "cats=East,North,West|\xE9\x9B\x86\xE8\xA8\x88|40 7 25"}});
  // Chartex references go through hidden `_xlchart.v1.N` names.
  ExpectFixtureData("chart_chartex.xlsx", {{"cx1", "cats=start,up,down,end|Amt|10 5 -3 12"}});
  ExpectFixtureData("chart_chartsheet.xlsx",
                    {
                        {"\xE3\x82\xB0\xE3\x83\xA9\xE3\x83\x95 1", "cats=a,b,c,d|S1|1 2 -3 7|S2|4 5 6 8"},
                        {"embedded_line", "cats=a,b,c,d|S1|1 2 -3 7|S2|4 5 6 8"},
                    });
  ExpectFixtureData("chart_combo.xlsx", {
                                            {"combo",
                                             "cats=a,b,c,d|Sales|100 200 150 300|Rate|0.10000000000000001 "
                                             "0.25 0.20000000000000001 0.40000000000000002"},
                                            {"doughnut", "cats=a,b,c,d|Sales|100 200 150 300"},
                                        });
}

/// Every fixture: each resolved series agrees point-for-point with the cache
/// Excel saved (modulo whole-column trimming), and nothing is unresolved.
TEST(ChartData, EveryFixtureAgreesWithItsCache) {
  const std::vector<std::string> fixtures = {
      "chart_column.xlsx",   "chart_bar.xlsx",        "chart_line.xlsx",   "chart_area.xlsx",
      "chart_pie.xlsx",      "chart_scatter.xlsx",    "chart_format.xlsx", "chart_filter.xlsx",
      "chart_data.xlsx",     "chart_cache.xlsx",      "chart_refs.xlsx",   "chart_combo.xlsx",
      "chart_doughnut.xlsx", "chart_chartsheet.xlsx", "chart_pivot.xlsx",  "chart_xlsb.xlsb"};
  int compared = 0;
  for (const std::string& fixture : fixtures) {
    const Workbook wb = LoadFixture(fixture);
    for (const LoadedChart& chart : LoadCharts(wb)) {
      const ChartData data = Resolve(wb, chart);
      EXPECT_TRUE(data.unresolved.empty()) << fixture << " " << chart.name;
      // `#N/A` shown as a blank, and blanks plotted as zero.
      const bool na_is_zero =
          chart.model.disp_na_as_blank == Tri::kTrue &&
          (chart.model.disp_blanks_as == ChartBlanksAs::kZero || chart.model.disp_blanks_as == ChartBlanksAs::kAbsent);
      for (const ChartPlot& plot : chart.model.plots) {
        for (const ChartSeries& s : plot.series) {
          const bool xy = plot.type == ChartPlotType::kScatter;
          const ChartDataSource& val = xy ? s.y : s.val;
          const ChartDataSeries* got = nullptr;
          for (const ChartDataSeries& d : data.series) {
            if (d.idx == s.idx) {
              got = &d;
            }
          }
          if (val.kind != ChartDataSource::Kind::kRef) {
            continue;
          }
          if (val.point_count == 0) {
            EXPECT_EQ(got, nullptr) << fixture << " " << chart.name << " hidden series";
            continue;
          }
          ASSERT_NE(got, nullptr) << fixture << " " << chart.name;
          // The data keeps filtered categories; Excel's cache holds only the shown ones.
          std::vector<std::size_t> shown;
          const std::size_t total = xy ? got->xy.size() : got->points.size();
          for (std::size_t i = 0; i < total; ++i) {
            if (std::find(chart.model.category_filter.begin(), chart.model.category_filter.end(), i) ==
                chart.model.category_filter.end()) {
              shown.push_back(i);
            }
          }
          const std::size_t n = shown.size();
          if (val.point_count != Sheet::kMaxRows) {
            EXPECT_EQ(n, val.point_count) << fixture << " " << chart.name;
          }
          for (const ChartValue& v : val.cache) {
            if (v.idx >= n) {
              ADD_FAILURE() << fixture << " " << chart.name << " idx " << v.idx;
              continue;
            }
            const std::size_t at = shown[v.idx];
            const double value = xy ? got->xy[at].y : got->points[at].value;
            const ChartPointState state = xy ? got->xy[at].state : got->points[at].state;
            if (v.kind == ChartValue::Kind::kNumber) {
              EXPECT_EQ(value, v.number) << fixture << " " << chart.name << " idx " << v.idx;
              EXPECT_EQ(state, ChartPointState::kValue) << fixture << " " << chart.name << " idx " << v.idx;
              ++compared;
            } else if (na_is_zero) {
              EXPECT_EQ(value, 0.0) << fixture << " " << chart.name << " idx " << v.idx;
            } else {
              EXPECT_NE(state, ChartPointState::kValue) << fixture << " " << chart.name << " idx " << v.idx;
            }
          }
          if (s.name.kind == ChartDataSource::Kind::kRef && !s.name.cache.empty()) {
            EXPECT_EQ(got->name.value_or("<null>"), s.name.cache.front().text) << fixture << " " << chart.name;
          }
        }
      }
    }
  }
  EXPECT_GT(compared, 300);
}

// ---------------------------------------------------------------------------
// Synthetic: shapes and rules the fixtures do not pin
// ---------------------------------------------------------------------------

Workbook GridBook() {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  // B2:C3 = 1 3 / 2 4 (column-major gives 1 2 3 4).
  s.set_cell_value(1, 1, Value::number(1));
  s.set_cell_value(2, 1, Value::number(2));
  s.set_cell_value(1, 2, Value::number(3));
  s.set_cell_value(2, 2, Value::number(4));
  s.set_cell_text(0, 1, "N1");
  s.set_cell_text(0, 3, "N3");
  return wb;
}

TEST(ChartData, TwoDimensionalValuesFlattenColumnMajor) {
  const Workbook wb = GridBook();
  const ChartData data = ResolveSynthetic(wb, OneSeries(ChartPlotType::kBar, "", "Sheet1!$B$2:$C$3"));
  EXPECT_EQ(DataText(data), "cats=1,2,3,4(gen)|<null>|1 2 3 4");
}

TEST(ChartData, MultiCellNameJoinsNonBlankTextsWithOneSpace) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "Sheet1!$B$2:$B$3");
  model.plots[0].series[0].name = Ref("Sheet1!$B$1:$D$1");
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2(gen)|N1 N3|1 2");
}

TEST(ChartData, RowOrientedMultiLevelCategories) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 0, "G1");
  s.set_cell_text(0, 2, "G2");
  for (std::uint32_t c = 0; c < 3; ++c) {
    s.set_cell_text(1, c, std::string("i") + std::to_string(c + 1));
    s.set_cell_value(2, c, Value::number(c + 1));
  }
  const ChartData data = ResolveSynthetic(wb, OneSeries(ChartPlotType::kBar, "Sheet1!$A$1:$C$2", "Sheet1!$A$3:$C$3"));
  EXPECT_EQ(DataText(data), "cats=i1,i2,i3|lvl=G1,,G2|<null>|1 2 3");
}

TEST(ChartData, UnresolvedSourcesFallBackToTheCache) {
  const Workbook wb = GridBook();
  ChartModel model;
  ChartPlot plot;
  plot.type = ChartPlotType::kLine;
  ChartSeries s;
  s.idx = 4;
  // External: the cache stands in for the cells.
  s.name = Ref("[1]Data!$A$1");
  s.name.cache.push_back(ChartValue{0, ChartValue::Kind::kText, 0.0, "Ext", ""});
  s.name.point_count = 1;
  s.cat = Ref("[1]Data!$A$2:$A$4");
  s.cat.point_count = 3;
  s.cat.cache.push_back(ChartValue{0, ChartValue::Kind::kText, 0.0, "p", ""});
  s.cat.cache.push_back(ChartValue{2, ChartValue::Kind::kText, 0.0, "r", ""});
  s.val = Ref("[1]Data!$B$2:$B$4");
  s.val.point_count = 3;
  s.val.cache.push_back(ChartValue{0, ChartValue::Kind::kNumber, 7.0, "", ""});
  s.val.cache.push_back(ChartValue{1, ChartValue::Kind::kText, 0.0, "#N/A", ""});
  plot.series.push_back(std::move(s));
  // `#REF!` after a sheet delete: name `#REF!`, no categories, one value of 1.
  ChartSeries r;
  r.idx = 5;
  r.name = Ref("#REF!");
  r.name.cache.push_back(ChartValue{0, ChartValue::Kind::kText, 0.0, "#REF!", ""});
  r.name.point_count = 1;
  r.val = Ref("#REF!");
  r.val.point_count = 1;
  r.val.cache.push_back(ChartValue{0, ChartValue::Kind::kNumber, 1.0, "", ""});
  plot.series.push_back(std::move(r));
  model.plots.push_back(std::move(plot));
  model.disp_blanks_as = ChartBlanksAs::kGap;
  const ChartData data = ResolveSynthetic(wb, model);
  EXPECT_EQ(DataText(data), "cats=p,,r|Ext|7 conn gap|#REF!|1");
  std::vector<std::string> unresolved;
  for (const ChartUnresolved& u : data.unresolved) {
    unresolved.push_back(std::to_string(u.series_idx) + ":" + u.source);
  }
  EXPECT_EQ(Join(unresolved), "4:name,4:cat,4:val,5:name,5:val");
  EXPECT_TRUE(data.series[0].val_ranges.empty());
}

TEST(ChartData, CacheDateCategoryRendersInTheWorkbookLocale) {
  Workbook wb = GridBook();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  ChartModel model = OneSeries(ChartPlotType::kBar, "[1]Data!$A$2:$A$2", "[1]Data!$B$2:$B$2");
  ChartDataSource& cat = model.plots[0].series[0].cat;
  cat.point_count = 1;
  cat.cache_format_code = "m/d/yy";  // built-in 14 as a cache spells it
  cat.cache.push_back(ChartValue{0, ChartValue::Kind::kNumber, 45355.0, "", ""});
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)).substr(0, 14), "cats=2024/3/4|");
}

TEST(ChartData, WholeColumnCacheOfAnUnresolvedSourceIsTrimmed) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "[1]Data!$B:$B");
  ChartDataSource& val = model.plots[0].series[0].val;
  val.point_count = Sheet::kMaxRows;
  val.cache.push_back(ChartValue{0, ChartValue::Kind::kNumber, 2.0, "", ""});
  val.cache.push_back(ChartValue{2, ChartValue::Kind::kNumber, 5.0, "", ""});
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2,3(gen)|<null>|2 0 5");
}

TEST(ChartData, RejectsSparseCacheWithOutOfGridIndex) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "[1]Data!$B$1");
  ChartDataSource& val = model.plots[0].series[0].val;
  val.point_count = 1U;
  val.cache.push_back(ChartValue{std::numeric_limits<std::uint32_t>::max(), ChartValue::Kind::kNumber, 7.0, "", ""});

  auto data = resolve_chart_data(wb, 0, model);
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, RejectsOutOfGridPointFromXmlLiteralBeforeDenseAllocation) {
  const Workbook wb = GridBook();
  const std::string xml =
      "<?xml version=\"1.0\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
      "<c:chart><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:val><c:numLit><c:ptCount val=\"1\"/><c:pt idx=\"4294967295\"><c:v>7</c:v></c:pt>"
      "</c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
  pugi::xml_document doc;
  auto parsed = parse_chart_part(std::vector<std::uint8_t>(xml.begin(), xml.end()), doc);
  ASSERT_TRUE(static_cast<bool>(parsed));
  auto model = read_chart_model(wb, doc);
  ASSERT_TRUE(static_cast<bool>(model));

  auto data = resolve_chart_data(wb, 0U, model.value());
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, RejectsOutOfGridPointFromXmlCacheBeforeDenseAllocation) {
  const Workbook wb = GridBook();
  const std::string xml =
      "<?xml version=\"1.0\"?><c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\">"
      "<c:chart><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:val><c:numRef><c:f>[1]Data!$B$1</c:f><c:numCache><c:ptCount val=\"1\"/>"
      "<c:pt idx=\"4294967295\"><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:val>"
      "</c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>";
  pugi::xml_document doc;
  auto parsed = parse_chart_part(std::vector<std::uint8_t>(xml.begin(), xml.end()), doc);
  ASSERT_TRUE(static_cast<bool>(parsed));
  auto model = read_chart_model(wb, doc);
  ASSERT_TRUE(static_cast<bool>(model));

  auto data = resolve_chart_data(wb, 0U, model.value());
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, RejectsRectangularReferenceBeforePositionAllocation) {
  const Workbook wb = GridBook();
  // 16,383 columns by 1,000 rows exceeds the shared expansion budget. The
  // source must fail before positions.reserve() or the nested walk.
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "Sheet1!$A$1:$XFC$1000");
  auto data = resolve_chart_data(wb, 0U, model);
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, RejectsExcessiveSparsePointCountBeforeAllocation) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "[1]Data!$B$1");
  ChartDataSource& val = model.plots[0].series[0].val;
  val.point_count = Sheet::kMaxRows + 1U;

  auto data = resolve_chart_data(wb, 0, model);
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, RejectsMultiLevelCacheBeyondCellBudgetBeforeAllocation) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "[1]Data!$A$1", "Sheet1!$B$2:$B$3");
  ChartDataSource& cat = model.plots[0].series[0].cat;
  cat.point_count = Sheet::kMaxRows - 1U;
  cat.cache_levels.assign(10U, {});

  auto data = resolve_chart_data(wb, 0, model);
  ASSERT_FALSE(static_cast<bool>(data));
  EXPECT_EQ(data.error().code, FormulonErrorCode::kSecResourceLimit);
}

TEST(ChartData, DenseCacheScansUnsortedIndicesWithinExcelBounds) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "[1]Data!$B$1");
  ChartDataSource& val = model.plots[0].series[0].val;
  val.point_count = 2U;
  val.cache.push_back(ChartValue{2U, ChartValue::Kind::kNumber, 5.0, "", ""});
  val.cache.push_back(ChartValue{0U, ChartValue::Kind::kNumber, 2.0, "", ""});

  auto data = resolve_chart_data(wb, 0, model);
  ASSERT_TRUE(static_cast<bool>(data));
  ASSERT_EQ(data.value().series.size(), 1U);
  EXPECT_EQ(SeriesText(data.value().series[0]), "<null>|2 0 5");
}

TEST(ChartData, ResolvedRangesKeepTheReferencedExtent) {
  const Workbook wb = GridBook();
  ChartModel model = OneSeries(ChartPlotType::kBar, "Sheet1!$A$2:$A$3", "Sheet1!$B:$B");
  model.plots[0].series[0].name = Ref("Sheet1!$B$1");
  const ChartData data = ResolveSynthetic(wb, model);
  ASSERT_EQ(data.series.size(), 1U);
  const ChartDataSeries& s = data.series[0];
  ASSERT_EQ(s.val_ranges.size(), 1U);
  EXPECT_EQ(s.val_ranges[0].bottom, Sheet::kMaxRows - 1);
  ASSERT_EQ(s.cat_ranges.size(), 1U);
  EXPECT_EQ(s.cat_ranges[0].top, 1U);
  ASSERT_EQ(s.name_ranges.size(), 1U);
  EXPECT_EQ(s.name_ranges[0].left, 1U);
  // Points stop at the populated extent (row 3), blanks plotted as zero by default.
  EXPECT_EQ(SeriesText(s), "N1|0 1 2");
}

TEST(ChartData, BubbleSizesAndXYDomain) {
  const Workbook wb = GridBook();
  ChartModel model;
  ChartPlot plot;
  plot.type = ChartPlotType::kBubble;
  ChartSeries s;
  s.x = Ref("Sheet1!$B$2:$B$3");
  s.y = Ref("Sheet1!$C$2:$C$3");
  s.size = Ref("Sheet1!$B$2:$B$3");
  plot.series.push_back(std::move(s));
  model.plots.push_back(std::move(plot));
  const ChartData data = ResolveSynthetic(wb, model);
  EXPECT_EQ(data.domain, ChartData::Domain::kXY);
  EXPECT_EQ(data.point_count, 2U);
  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(SeriesText(data.series[0]), "<null>|1:3 2:4");
  EXPECT_EQ(data.series[0].sizes, (std::vector<double>{1, 2}));
}

TEST(ChartData, XYSeriesCarriesTheXAndYFormatCodes) {
  Workbook wb = GridBook();
  CellXf pct;
  pct.num_fmt_id = 10;
  pct.apply_number_format = true;
  wb.mutable_styles().cell_xfs.push_back(pct);
  const auto xf = static_cast<std::uint32_t>(wb.styles().cell_xfs.size() - 1);
  wb.sheet(0).set_cell_xf_index(1, 1, xf);
  wb.sheet(0).set_cell_xf_index(2, 1, xf);
  ChartModel model;
  ChartPlot plot;
  plot.type = ChartPlotType::kScatter;
  ChartSeries s;
  s.x = Ref("Sheet1!$B$2:$B$3");
  s.y = Ref("Sheet1!$C$2:$C$3");
  plot.series.push_back(std::move(s));
  model.plots.push_back(std::move(plot));
  const ChartData data = ResolveSynthetic(wb, model);
  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(data.series[0].x_format_code, "0.00%");
  EXPECT_EQ(data.series[0].format_code, "General");
}

TEST(ChartData, TitleAndAxisTitleTexts) {
  Workbook wb = GridBook();
  wb.sheet(0).set_cell_value(5, 0, Value::number(2024));
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "Sheet1!$B$2:$B$3");
  model.title.kind = ChartTitleKind::kRef;
  model.title.ref_formula = "Sheet1!$A$6";
  ChartAxis cat_axis;
  cat_axis.id = 10;
  cat_axis.title.kind = ChartTitleKind::kText;
  cat_axis.title.text = "Month";
  ChartAxis val_axis;
  val_axis.id = 20;
  val_axis.title.kind = ChartTitleKind::kRef;
  val_axis.title.ref_formula = "Sheet1!$B$1";
  ChartAxis bad_axis;
  bad_axis.id = 30;
  bad_axis.title.kind = ChartTitleKind::kRef;
  bad_axis.title.ref_formula = "#REF!";
  ChartAxis auto_axis;
  auto_axis.id = 40;
  auto_axis.title.kind = ChartTitleKind::kAuto;
  model.axes = {cat_axis, val_axis, bad_axis, auto_axis};
  const ChartData data = ResolveSynthetic(wb, model);
  EXPECT_EQ(data.title_text.value_or("<null>"), "2024");
  std::vector<std::string> axes;
  for (const ChartAxisTitleText& t : data.axis_title_texts) {
    axes.push_back(std::to_string(t.axis_id) + "=" + t.text);
  }
  EXPECT_EQ(Join(axes), "10=Month,20=N1");
  ASSERT_EQ(data.unresolved.size(), 1U);
  EXPECT_EQ(data.unresolved[0].axis_id, 30U);
  EXPECT_EQ(data.unresolved[0].source, "axis_title");

  model.title.kind = ChartTitleKind::kText;
  model.title.text = "Sales";
  EXPECT_EQ(ResolveSynthetic(wb, model).title_text.value_or("<null>"), "Sales");
  model.title.kind = ChartTitleKind::kAuto;
  EXPECT_FALSE(ResolveSynthetic(wb, model).title_text.has_value());
  model.title.kind = ChartTitleKind::kRef;
  model.title.ref_formula = "[1]X!$A$1";
  const ChartData bad = ResolveSynthetic(wb, model);
  EXPECT_FALSE(bad.title_text.has_value());
  ASSERT_FALSE(bad.unresolved.empty());
  EXPECT_EQ(bad.unresolved[0].source, "title");
}

TEST(ChartData, FixtureTitlesResolve) {
  const Workbook wb = LoadFixture("chart_format.xlsx");
  for (const LoadedChart& chart : LoadCharts(wb)) {
    const ChartData data = Resolve(wb, chart);
    // No fixture title carries text or a reference; auto / none give no text.
    EXPECT_FALSE(data.title_text.has_value()) << chart.name;
  }
}

void HideRow(Sheet& sheet, std::uint32_t row) {
  RowLayout entry;
  entry.row = row;
  entry.hidden = true;
  insert_row_override(sheet.mutable_layout(), entry);
}

void HideCol(Sheet& sheet, std::uint32_t col) {
  ColumnLayout span;
  span.first = col;
  span.last = col;
  span.hidden = true;
  sheet.mutable_layout().columns.push_back(span);
}

void SetEmptyString(Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  sheet.set_cell_formula(row, col, "=\"\"");
  sheet.set_cell_cached_value(row, col, Value::text(""));
}

TEST(ChartData, NameInAHiddenRowOrColumnIsKept) {
  Workbook wb = GridBook();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(5, 0, "HidRowName");
  s.set_cell_text(0, 4, "HidColName");
  HideRow(s, 5);
  HideCol(s, 4);
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "Sheet1!$B$2:$B$3");
  model.plots[0].series[0].name = Ref("Sheet1!$A$6");
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2(gen)|HidRowName|1 2");
  model.plots[0].series[0].name = Ref("Sheet1!$E$1");
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2(gen)|HidColName|1 2");
}

TEST(ChartData, EmptyStringPlotsAsZeroAndLabelsAsBlank) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 0, "a");
  SetEmptyString(s, 1, 0);
  s.set_cell_text(2, 0, "c");
  s.set_cell_value(0, 1, Value::number(1));
  SetEmptyString(s, 1, 1);
  s.set_cell_value(2, 1, Value::number(3));
  ChartModel model = OneSeries(ChartPlotType::kLine, "Sheet1!$A$1:$A$3", "Sheet1!$B$1:$B$3");
  // Under gap a true blank would read `gap`; `=""` plots as text does.
  model.disp_blanks_as = ChartBlanksAs::kGap;
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=a,,c|<null>|1 0 3");
}

TEST(ChartData, BlankScatterXIsNotDrawn) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_value(0, 0, Value::number(1));
  s.set_cell_value(2, 0, Value::error(ErrorCode::Div0));
  s.set_cell_value(3, 0, Value::boolean(true));
  s.set_cell_value(4, 0, Value::number(5));
  for (std::uint32_t r = 0; r < 5; ++r) {
    s.set_cell_value(r, 1, Value::number(10 + r));
  }
  ChartModel model = OneSeries(ChartPlotType::kScatter, "", "Sheet1!$B$1:$B$5");
  model.plots[0].series[0].x = Ref("Sheet1!$A$1:$A$5");
  const ChartData data = ResolveSynthetic(wb, model);
  ASSERT_EQ(data.series.size(), 1U);
  EXPECT_EQ(SeriesText(data.series[0]), "<null>|1:10 0:gap 0:12 1:13 5:14");
}

TEST(ChartData, MultiCellNameSkipsBlankCells) {
  Workbook wb = GridBook();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(6, 0, "X");
  s.set_cell_text(6, 2, "Y");
  s.set_cell_text(7, 0, "P");
  s.set_cell_text(7, 1, "Q");
  ChartModel model = OneSeries(ChartPlotType::kBar, "", "Sheet1!$B$2:$B$3");
  model.plots[0].series[0].name = Ref("Sheet1!$A$7:$C$7");
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2(gen)|X Y|1 2");
  model.plots[0].series[0].name = Ref("Sheet1!$A$8:$C$8");
  EXPECT_EQ(DataText(ResolveSynthetic(wb, model)), "cats=1,2(gen)|P Q|1 2");
}

TEST(ChartData, AllCategoriesHiddenDropsTheSeries) {
  Workbook wb = Workbook::create();
  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 1, "Kept");
  for (std::uint32_t r = 1; r <= 3; ++r) {
    s.set_cell_text(r, 0, "g");
    s.set_cell_text(r, 1, "i" + std::to_string(r));
    s.set_cell_value(r, 2, Value::number(r));
    HideRow(s, r);
  }
  for (const char* cat : {"Sheet1!$B$2:$B$4", "Sheet1!$A$2:$B$4"}) {
    ChartModel model = OneSeries(ChartPlotType::kBar, cat, "Sheet1!$C$2:$C$4");
    model.plots[0].series[0].name = Ref("Sheet1!$B$1");
    const ChartData data = ResolveSynthetic(wb, model);
    EXPECT_TRUE(data.series.empty()) << cat;
    EXPECT_TRUE(data.categories.labels.empty()) << cat;
  }
}

}  // namespace
}  // namespace formulon
