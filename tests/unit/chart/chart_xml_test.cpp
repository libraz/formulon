#include "chart/chart_xml.h"

#include <cmath>
#include <cstdint>
#include <cstdio>
#include <cstdlib>
#include <string>
#include <utility>
#include <vector>

#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "utils/error.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::string Wrap(const std::string& chart_children, const std::string& space_children);
ChartModel Read(const std::string& xml, const Workbook& wb);

// ---------------------------------------------------------------------------
// Loading and describing
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

Workbook LoadFixture(const std::string& name) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
  if (bytes.empty()) {
    return Workbook::create_empty();
  }
  const io::ByteSpan span{bytes.data(), bytes.size()};
  if (name.size() > 5 && name.substr(name.size() - 5) == ".xlsb") {
    auto read = io::xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
    return read ? std::move(read.value().workbook) : Workbook::create_empty();
  }
  auto read = io::read_ooxml(span);
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

std::string Hex(std::uint32_t argb) {
  char buf[16];
  std::snprintf(buf, sizeof(buf), "%08X", static_cast<unsigned>(argb));
  return buf;
}

/// Shortest decimal text that reads back as `value`.
std::string Num(double value) {
  char buf[40];
  if (value == std::floor(value) && std::fabs(value) < 1e15) {
    std::snprintf(buf, sizeof(buf), "%.0f", value);
    return buf;
  }
  for (int precision = 1; precision <= 17; ++precision) {
    std::snprintf(buf, sizeof(buf), "%.*g", precision, value);
    if (std::strtod(buf, nullptr) == value) {
      break;
    }
  }
  return buf;
}

std::string Pts(const std::vector<ChartValue>& values) {
  std::string out;
  for (const ChartValue& v : values) {
    out += " " + std::to_string(v.idx) + ":" + (v.kind == ChartValue::Kind::kNumber ? Num(v.number) : v.text);
    if (!v.format_code.empty()) {
      out += "[" + v.format_code + "]";
    }
  }
  return out;
}

std::string Src(const ChartDataSource& s) {
  switch (s.kind) {
    case ChartDataSource::Kind::kNone:
      return "absent";
    case ChartDataSource::Kind::kLiteral:
      return "lit n=" + std::to_string(s.point_count) + Pts(s.literal);
    case ChartDataSource::Kind::kRef:
      break;
  }
  const char* element = "ref";
  switch (s.ref_element) {
    case ChartDataSource::RefElement::kNumRef:
      element = "numRef";
      break;
    case ChartDataSource::RefElement::kStrRef:
      element = "strRef";
      break;
    case ChartDataSource::RefElement::kMultiLvlStrRef:
      element = "multiLvlStrRef";
      break;
    case ChartDataSource::RefElement::kCxDim:
      element = "cxDim";
      break;
    case ChartDataSource::RefElement::kNone:
      break;
  }
  std::string out = std::string(element) + " " + s.formula;
  if (!s.cache_format_code.empty()) {
    out += " fc=" + s.cache_format_code;
  }
  if (s.cache.empty() && s.point_count == 0) {
    return out + " no cache";
  }
  out += " n=" + std::to_string(s.point_count) + Pts(s.cache);
  for (const std::vector<ChartValue>& level : s.cache_levels) {
    out += " up=[" + Pts(level).substr(1) + "]";
  }
  return out;
}

std::string FillStr(const ChartFill& f) {
  if (!f.color) {
    return "absent";
  }
  return Hex(*f.color) + (f.alpha ? " alpha=" + std::to_string(*f.alpha) : "");
}

std::string LineStr(const ChartLine& l) {
  if (l.none) {
    return "none";
  }
  if (!l.color && !l.width_emu && !l.dash) {
    return "absent";
  }
  return (l.color ? Hex(*l.color) : std::string("-")) + (l.alpha ? " alpha=" + std::to_string(*l.alpha) : "") +
         " w=" + (l.width_emu ? std::to_string(*l.width_emu) : "absent") + " dash=" + (l.dash ? l.dash_name : "absent");
}

std::string MarkerStr(const std::optional<ChartMarker>& marker) {
  if (!marker) {
    return "absent";
  }
  std::string out =
      "symbol=" + marker->symbol_name + " size=" + (marker->size ? std::to_string(*marker->size) : "absent");
  if (marker->fill.color) {
    out += " fill=" + FillStr(marker->fill);
  }
  if (marker->line.color || marker->line.width_emu || marker->line.dash || marker->line.none) {
    out += " line=" + LineStr(marker->line);
  }
  return out;
}

std::string PointsStr(const std::vector<ChartPoint>& points) {
  if (points.empty()) {
    return "-";
  }
  std::string out;
  for (const ChartPoint& p : points) {
    out += out.empty() ? "" : "; ";
    out += std::to_string(p.idx) + ":";
    out += p.fill.color ? "fill=" + FillStr(p.fill) : "marker(" + MarkerStr(p.marker) + ")";
  }
  return out;
}

std::string LabelsStr(const std::optional<ChartLabels>& labels) {
  if (!labels) {
    return "absent";
  }
  if (labels->deleted) {
    return "deleted";
  }
  std::string shown;
  const auto add = [&shown](bool on, const char* name) {
    if (on) {
      shown += shown.empty() ? name : std::string("+") + name;
    }
  };
  add(labels->show_val, "Val");
  add(labels->show_cat, "CatName");
  add(labels->show_ser, "SerName");
  add(labels->show_pct, "Percent");
  add(labels->show_legend_key, "LegendKey");
  std::string out = "show=" + shown;
  if (labels->separator) {
    out += " sep='" + *labels->separator + "'";
  }
  if (labels->position) {
    out += " pos=" + std::to_string(static_cast<int>(*labels->position));
  }
  if (labels->num_fmt) {
    out += " numFmt=" + labels->num_fmt->code + (labels->num_fmt->source_linked ? "!" : "");
  }
  if (labels->font_size_hundredths_pt) {
    out += " sz=" + std::to_string(*labels->font_size_hundredths_pt);
  }
  if (labels->font_color) {
    out += " color=" + Hex(*labels->font_color);
  }
  return out;
}

const char* TriStr(Tri t) {
  return t == Tri::kAbsent ? "absent" : t == Tri::kFalse ? "0" : "1";
}

std::string SeriesRow(std::size_t plot_index, const ChartPlot& plot, const ChartSeries& s) {
  const bool xy = plot.type == ChartPlotType::kScatter || plot.type == ChartPlotType::kBubble;
  return std::to_string(plot_index) + "|" + std::to_string(s.idx) + "/" + std::to_string(s.order) +
         (s.filtered ? " filtered" : "") + "|name=" + Src(s.name) + "|cat=" + Src(xy ? s.x : s.cat) +
         "|val=" + Src(xy ? s.y : s.val) + "|fill=" + FillStr(s.fill) + "|line=" + LineStr(s.line) +
         "|marker=" + MarkerStr(s.marker) + "|smooth=" + TriStr(s.smooth) + "|dpt=" + PointsStr(s.points) +
         "|labels=" + LabelsStr(s.labels);
}

const char* PlotName(ChartPlotType t) {
  switch (t) {
    case ChartPlotType::kBar:
      return "barChart";
    case ChartPlotType::kLine:
      return "lineChart";
    case ChartPlotType::kArea:
      return "areaChart";
    case ChartPlotType::kPie:
      return "pieChart";
    case ChartPlotType::kDoughnut:
      return "doughnutChart";
    case ChartPlotType::kScatter:
      return "scatterChart";
    case ChartPlotType::kChartex:
      return "chartex";
    default:
      return "other";
  }
}

const char* GroupingStr(const std::optional<ChartGrouping>& g) {
  if (!g)
    return "absent";
  switch (*g) {
    case ChartGrouping::kClustered:
      return "clustered";
    case ChartGrouping::kStacked:
      return "stacked";
    case ChartGrouping::kPercentStacked:
      return "percentStacked";
    case ChartGrouping::kStandard:
      return "standard";
  }
  return "absent";
}

template <typename T>
std::string OptStr(const std::optional<T>& v) {
  return v ? std::to_string(*v) : "absent";
}

const char* LegendStr(const ChartLegend& l) {
  if (!l.present)
    return "-";
  switch (l.position) {
    case ChartLegendPosition::kBottom:
      return "b";
    case ChartLegendPosition::kTop:
      return "t";
    case ChartLegendPosition::kLeft:
      return "l";
    case ChartLegendPosition::kRight:
      return "r";
    case ChartLegendPosition::kTopRight:
      return "tr";
    case ChartLegendPosition::kNone:
      return "-";
  }
  return "-";
}

/// "id name|plots|barDir|grouping|vary|gap|overlap|ang|title|atd|legend|pvo/dba[|cf=..][|pivot]" of the first plot.
std::string Header(const ChartModel& m, const std::string& id_name) {
  std::string plots;
  for (const ChartPlot& p : m.plots) {
    plots += plots.empty() ? "" : "+";
    plots += PlotName(p.type);
  }
  const ChartPlot& p = m.plots.front();
  const char* bar_dir = p.bar_dir == ChartBarDirection::kCol   ? "col"
                        : p.bar_dir == ChartBarDirection::kBar ? "bar"
                                                               : "absent";
  const char* title = m.title.kind == ChartTitleKind::kNone   ? "none"
                      : m.title.kind == ChartTitleKind::kAuto ? "auto"
                      : m.title.kind == ChartTitleKind::kText ? "text"
                                                              : "ref";
  const char* blanks = m.disp_blanks_as == ChartBlanksAs::kGap    ? "gap"
                       : m.disp_blanks_as == ChartBlanksAs::kZero ? "zero"
                       : m.disp_blanks_as == ChartBlanksAs::kSpan ? "span"
                                                                  : "absent";
  std::string out = id_name + "|" + plots + "|" + bar_dir + "|" + GroupingStr(p.grouping) + "|" +
                    TriStr(p.vary_colors) + "|" + OptStr(p.gap_width) + "|" + OptStr(p.overlap) + "|" +
                    OptStr(p.first_slice_angle) + "|" + title + "|" + TriStr(m.auto_title_deleted) + "|" +
                    LegendStr(m.legend) + "|" + TriStr(m.plot_vis_only) + "/" + blanks + "/" +
                    TriStr(m.disp_na_as_blank);
  if (p.hole_size) {
    out += "|hole=" + std::to_string(*p.hole_size);
  }
  if (!m.category_filter.empty()) {
    out += "|cf=";
    for (std::size_t i = 0; i < m.category_filter.size(); ++i) {
      out += (i == 0 ? "" : ",") + std::to_string(m.category_filter[i]);
    }
  }
  if (m.is_pivot_chart) {
    out += "|pivot";
  }
  return out;
}

/// One description per chart part of the workbook, in `all_chart_parts` order.
std::vector<std::string> DescribeAll(const Workbook& wb) {
  std::vector<std::string> out;
  auto refs = all_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(refs));
  if (!refs) {
    return out;
  }
  for (const SheetChartPartRef& item : refs.value()) {
    std::string id_name = std::to_string(item.ref.object_id);
    auto objects = list_drawing_objects(wb, item.sheet);
    if (objects) {
      for (const DrawingObject& obj : objects.value()) {
        if (obj.object_id == item.ref.object_id) {
          id_name += " " + obj.name;
        }
      }
    }
    const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, item.ref.part_path);
    EXPECT_NE(part, nullptr) << item.ref.part_path;
    if (part == nullptr) {
      continue;
    }
    pugi::xml_document doc;
    auto parsed = parse_chart_part(part->bytes, doc);
    EXPECT_TRUE(static_cast<bool>(parsed)) << item.ref.part_path;
    auto model = read_chart_model(wb, doc);
    EXPECT_TRUE(static_cast<bool>(model)) << item.ref.part_path;
    if (!parsed || !model) {
      continue;
    }
    std::string text = Header(model.value(), id_name);
    for (std::size_t pi = 0; pi < model.value().plots.size(); ++pi) {
      for (const ChartSeries& s : model.value().plots[pi].series) {
        text += "\n" + SeriesRow(pi, model.value().plots[pi], s);
      }
    }
    out.push_back(std::move(text));
  }
  return out;
}

void ExpectFixture(const std::string& file, const std::vector<std::string>& expected) {
  SCOPED_TRACE(file);
  const Workbook wb = LoadFixture(file);
  const std::vector<std::string> actual = DescribeAll(wb);
  ASSERT_EQ(actual.size(), expected.size());
  for (std::size_t i = 0; i < actual.size(); ++i) {
    EXPECT_EQ(actual[i], expected[i]);
  }
}

std::string Join(std::initializer_list<std::string> lines) {
  std::string out;
  for (const std::string& line : lines) {
    out += out.empty() ? "" : "\n";
    out += line;
  }
  return out;
}

/// "plot|idx/order|name|cat|val|fill|line|marker|smooth|dpt|labels" in the described shape.
std::string Row(int plot, const std::string& io, const std::string& name, const std::string& cat,
                const std::string& val, const std::string& fill, const std::string& line,
                const std::string& marker = "absent", const std::string& smooth = "absent",
                const std::string& dpt = "-", const std::string& labels = "absent") {
  return std::to_string(plot) + "|" + io + "|name=" + name + "|cat=" + cat + "|val=" + val + "|fill=" + fill +
         "|line=" + line + "|marker=" + marker + "|smooth=" + smooth + "|dpt=" + dpt + "|labels=" + labels;
}

// Frequently repeated fragments of the measured tables.
const std::string kNameS1 = "strRef Sheet1!$B$1 n=1 0:S1";
const std::string kNameS2 = "strRef Sheet1!$C$1 n=1 0:S2";
const std::string kCatAbcd = "strRef Sheet1!$A$2:$A$5 n=4 0:a 1:b 2:c 3:d";
const std::string kValB = "numRef Sheet1!$B$2:$B$5 fc=General n=4 0:1 1:2 2:3 3:7";
const std::string kValC = "numRef Sheet1!$C$2:$C$5 fc=General n=4 0:4 1:5 2:6 3:8";
const std::string kAccent1 = "FF156082";
const std::string kAccent2 = "FFE97132";
const std::string kBarLine = "none";
const std::string kLine1 = "FF156082 w=28575 dash=absent";
const std::string kLine2 = "FFE97132 w=28575 dash=absent";
const std::string kNoneMarker = "symbol=none size=absent";
const std::string kCircle1 = "symbol=circle size=5 fill=FF156082 line=FF156082 w=9525 dash=absent";
const std::string kCircle2 = "symbol=circle size=5 fill=FFE97132 line=FFE97132 w=9525 dash=absent";
const std::string kPieDots = "0:fill=FF156082; 1:fill=FFE97132; 2:fill=FF196B24; 3:fill=FF0F9ED5";

std::string BarPair(const std::string& header) {
  return Join({header, Row(0, "0/0", kNameS1, kCatAbcd, kValB, kAccent1, kBarLine),
               Row(0, "1/1", kNameS2, kCatAbcd, kValC, kAccent2, kBarLine)});
}

std::string LinePair(const std::string& header) {
  return Join({header, Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", kLine1, kNoneMarker, "0"),
               Row(0, "1/1", kNameS2, kCatAbcd, kValC, "absent", kLine2, kNoneMarker, "0")});
}

std::string AreaPair(const std::string& header) {
  return BarPair(header);
}

// ---------------------------------------------------------------------------
// Fixture models
// ---------------------------------------------------------------------------

TEST(ChartXml, ColumnFixtureMatchesMeasuredModel) {
  const std::vector<std::string> expected = {
      BarPair("2 col_clustered|barChart|col|clustered|0|219|absent|absent|none|0|b|1/gap/1"),
      BarPair("3 col_stacked|barChart|col|stacked|0|219|100|absent|none|0|b|1/gap/1"),
      BarPair("4 col_percent|barChart|col|percentStacked|0|219|100|absent|none|0|b|1/gap/1"),
      Join({"5 col_single|barChart|col|clustered|0|219|absent|absent|auto|0|b|1/gap/1",
            Row(0, "0/0", kNameS1, kCatAbcd, kValB, kAccent1, kBarLine)}),
  };
  ExpectFixture("chart_column.xlsx", expected);
  // The xlsb fixture holds byte-identical chart parts.
  ExpectFixture("chart_xlsb.xlsb", expected);
}

TEST(ChartXml, BarFixtureMatchesMeasuredModel) {
  ExpectFixture("chart_bar.xlsx",
                {BarPair("2 bar_clustered|barChart|bar|clustered|0|182|absent|absent|none|0|b|1/gap/1"),
                 BarPair("3 bar_stacked|barChart|bar|stacked|0|182|100|absent|none|0|b|1/gap/1"),
                 BarPair("4 bar_percent|barChart|bar|percentStacked|0|182|100|absent|none|0|b|1/gap/1")});
}

TEST(ChartXml, LineFixtureMatchesMeasuredModel) {
  ExpectFixture("chart_line.xlsx",
                {LinePair("2 line_standard|lineChart|absent|standard|0|absent|absent|absent|none|0|b|1/gap/1"),
                 LinePair("3 line_stacked|lineChart|absent|stacked|0|absent|absent|absent|none|0|b|1/gap/1"),
                 LinePair("4 line_percent|lineChart|absent|percentStacked|0|absent|absent|absent|none|0|b|1/gap/1"),
                 Join({"5 line_markers|lineChart|absent|standard|0|absent|absent|absent|none|0|b|1/gap/1",
                       Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", kLine1, kCircle1, "0"),
                       Row(0, "1/1", kNameS2, kCatAbcd, kValC, "absent", kLine2, kCircle2, "0")})});
}

TEST(ChartXml, AreaFixtureMatchesMeasuredModel) {
  ExpectFixture("chart_area.xlsx",
                {AreaPair("2 area_standard|areaChart|absent|standard|0|absent|absent|absent|none|0|b|1/gap/1"),
                 AreaPair("3 area_stacked|areaChart|absent|stacked|0|absent|absent|absent|none|0|b|1/gap/1"),
                 AreaPair("4 area_percent|areaChart|absent|percentStacked|0|absent|absent|absent|none|0|b|1/gap/1")});
}

TEST(ChartXml, PieFixtureCarriesPointColoursOnly) {
  ExpectFixture("chart_pie.xlsx",
                {Join({"2 pie_single|pieChart|absent|absent|1|absent|absent|0|auto|0|b|1/gap/1",
                       Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", "absent", "absent", "absent", kPieDots)}),
                 Join({"3 pie_two|pieChart|absent|absent|1|absent|absent|0|auto|0|b|1/gap/1",
                       Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", "absent", "absent", "absent", kPieDots),
                       Row(0, "1/1", kNameS2, kCatAbcd, kValC, "absent", "absent", "absent", "absent", kPieDots)})});
}

TEST(ChartXml, DoughnutFixtureReadsHoleSizeAndTypeChanges) {
  const auto pie_row = [&](const char* io, const std::string& name, const std::string& val) {
    return Row(0, io, name, kCatAbcd, val, "absent", "absent", "absent", "absent", kPieDots);
  };
  const std::string pie_head = "|absent|absent|1|absent|absent|0|";
  const std::string tail = "|0|b|1/gap/1";
  ExpectFixture(
      "chart_doughnut.xlsx",
      {Join({"2 doughnut|doughnutChart" + pie_head + "auto" + tail + "|hole=50", pie_row("0/0", kNameS1, kValB)}),
       Join({"3 doughnut_two|doughnutChart" + pie_head + "none" + tail + "|hole=50", pie_row("0/0", kNameS1, kValB),
             pie_row("1/1", kNameS2, kValC)}),
       Join({"4 pie_one_ref|pieChart" + pie_head + "auto" + tail, pie_row("0/0", kNameS1, kValB)}),
       Join({"5 pie_to_doughnut|doughnutChart" + pie_head + "none" + tail + "|hole=50", pie_row("0/0", kNameS1, kValB),
             pie_row("1/1", kNameS2, kValC)}),
       Join({"6 doughnut_to_column|barChart|col|clustered|0|150|absent|absent|none" + tail,
             Row(0, "0/0", kNameS1, kCatAbcd, kValB, kAccent1, "FFFFFFFF w=19050 dash=absent"),
             Row(0, "1/1", kNameS2, kCatAbcd, kValC, kAccent2, "FFFFFFFF w=19050 dash=absent")}),
       Join({"7 column_to_doughnut|doughnutChart" + pie_head + "none" + tail + "|hole=50",
             pie_row("0/0", kNameS1, kValB), pie_row("1/1", kNameS2, kValC)})});
}

TEST(ChartXml, ScatterFixtureReadsXAndYSources) {
  const std::string x = "numRef Sheet1!$E$2:$E$5 fc=General n=4 0:1 1:2 2:3 3:5";
  const std::string y1 = "numRef Sheet1!$F$2:$F$5 fc=General n=4 0:2 1:4 2:1 3:3";
  const std::string y2 = "numRef Sheet1!$G$2:$G$5 fc=General n=4 0:3 1:1 2:5 3:2";
  const std::string n1 = "strRef Sheet1!$F$1 n=1 0:Y1";
  const std::string n2 = "strRef Sheet1!$G$1 n=1 0:Y2";
  ExpectFixture("chart_scatter.xlsx",
                {Join({"2 scatter_markers|scatterChart|absent|absent|0|absent|absent|absent|none|0|b|1/gap/1",
                       Row(0, "0/0", n1, x, y1, "absent", "none", kCircle1, "0"),
                       Row(0, "1/1", n2, x, y2, "absent", "none", kCircle2, "0")}),
                 Join({"3 scatter_lines|scatterChart|absent|absent|0|absent|absent|absent|none|0|b|1/gap/1",
                       Row(0, "0/0", n1, x, y1, "absent", "FF156082 w=19050 dash=absent", kCircle1, "0"),
                       Row(0, "1/1", n2, x, y2, "absent", "FFE97132 w=19050 dash=absent", kCircle2, "0")}),
                 Join({"4 scatter_single|scatterChart|absent|absent|0|absent|absent|absent|auto|0|b|1/gap/1",
                       Row(0, "0/0", n1, x, y1, "absent", "none", kCircle1, "0")})});
}

TEST(ChartXml, FormatFixtureReadsLinesMarkersPointsAndLabels) {
  const std::string label = "show=Val sz=900 color=FF404040";
  const std::string wide_dash = "FFFF0000 alpha=70000 w=44450 dash=dash";
  ExpectFixture(
      "chart_format.xlsx",
      {Join({"2 fmt_line|lineChart|absent|standard|0|absent|absent|absent|none|0|b|1/gap/1",
             Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", wide_dash,
                 "symbol=circle size=9 fill=FF156082 line=" + wide_dash, "1",
                 "1:marker(symbol=diamond size=9 fill=FF0000FF line=" + wide_dash + ")",
                 "show=Val+CatName sep='; ' sz=900 color=FF404040"),
             Row(0, "1/1", kNameS2, kCatAbcd, kValC, "absent", "FFE97132 w=28575 dash=dot", kNoneMarker, "0")}),
       Join(
           {"3 fmt_col|barChart|col|clustered|0|80|20|absent|none|0|b|1/gap/1",
            Row(0, "0/0", kNameS1, kCatAbcd, kValB, "FF008000 alpha=50000", "FF000000 w=25400 dash=sysDash", "absent",
                "absent", "-", label),
            Row(0, "1/1", kNameS2, kCatAbcd, kValC, kAccent2, kBarLine, "absent", "absent", "2:fill=FF0000FF", label)}),
       Join({"4 fmt_pie|pieChart|absent|absent|1|absent|absent|45|auto|0|b|1/gap/1",
             Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", "absent", "absent", "absent",
                 "0:fill=FFFFC000; 1:fill=FFE97132; 2:fill=FF196B24; 3:fill=FF0F9ED5",
                 "show=Percent sz=900 color=FF404040")}),
       Join({"5 fmt_vary|barChart|col|clustered|1|219|-27|absent|auto|0|b|1/gap/1",
             Row(0, "0/0", kNameS1, kCatAbcd, kValB, "absent", "absent", "absent", "absent", kPieDots)}),
       Join({"6 fmt_scatter|scatterChart|absent|absent|0|absent|absent|absent|auto|0|b|1/gap/1",
             Row(0, "0/0", "strRef Sheet1!$F$1 n=1 0:Y1", "numRef Sheet1!$E$2:$E$5 fc=General n=4 0:1 1:2 2:3 3:5",
                 "numRef Sheet1!$F$2:$F$5 fc=General n=4 0:2 1:4 2:1 3:3", "absent", "FF000000 w=38100 dash=dashDot",
                 "symbol=square size=7 fill=FF156082 line=FF156082 w=9525 dash=dashDot", "0")})});
}

TEST(ChartXml, FilterFixtureReadsFilteredSeriesAndCategoryFilter) {
  const std::string cat_visible = "strRef (Sheet1!$A$2,Sheet1!$A$4:$A$5) n=3 0:a 1:c 2:d";
  ExpectFixture("chart_filter.xlsx",
                {Join({"2 col_clustered|barChart|col|clustered|0|219|absent|absent|auto|0|b|1/gap/1",
                       Row(0, "0/0", kNameS1, kCatAbcd, kValB, kAccent1, kBarLine),
                       Row(0, "1/1 filtered", kNameS2, kCatAbcd, kValC, kAccent2, kBarLine)}),
                 Join({"3 col_stacked|barChart|col|stacked|0|219|100|absent|none|0|b|1/gap/1|cf=1",
                       Row(0, "0/0", kNameS1, cat_visible,
                           "numRef (Sheet1!$B$2,Sheet1!$B$4:$B$5) fc=General n=3 0:1 1:3 2:7", kAccent1, kBarLine),
                       Row(0, "1/1", kNameS2, cat_visible,
                           "numRef (Sheet1!$C$2,Sheet1!$C$4:$C$5) fc=General n=3 0:4 1:6 2:8", kAccent2, kBarLine)}),
                 BarPair("4 col_percent|barChart|col|percentStacked|0|219|100|absent|none|0|b|1/gap/1"),
                 Join({"5 col_single|barChart|col|clustered|0|219|absent|absent|auto|0|b|1/gap/1",
                       Row(0, "0/0", kNameS1, kCatAbcd, kValB, kAccent1, kBarLine)})});
}

TEST(ChartXml, DataFixtureReadsCachesAndDisplayFlags) {
  const std::string cat8 = "strRef Sheet1!$A$2:$A$9 n=8 0:a 1:b 2:c 3:d 4:e 5:f 6:g 7:h";
  const std::string mixed = "numRef Sheet1!$B$2:$B$9 fc=General n=8 0:1 2:3 3:#N/A 4:0 5:1 6:0 7:8";
  const std::string name_mixed = "strRef Sheet1!$B$1 n=1 0:Mixed";
  const std::string scatter_x = "numRef Sheet1!$J$2:$J$9 fc=General n=8 0:3 1:1 2:4 3:1 4:5 5:9 6:2 7:6";
  const std::string kHead = "|lineChart|absent|standard|0|absent|absent|absent|auto|0|b|";
  const auto line_chart = [&](const std::string& id_name, const std::string& flags, const std::string& name,
                              const std::string& cat, const std::string& val) {
    return Join({id_name + kHead + flags, Row(0, "0/0", name, cat, val, "absent", kLine1, kNoneMarker, "0")});
  };
  const std::string bar_head = "|barChart|col|clustered|0|219|-27|absent|auto|0|b|1/gap/1";
  ExpectFixture(
      "chart_data.xlsx",
      {line_chart("2 mixed_line", "1/gap/1", name_mixed, cat8, mixed),
       Join({"3 mixed_col" + bar_head, Row(0, "0/0", name_mixed, cat8, mixed, kAccent1, kBarLine)}),
       Join({"4 mixed_scatter|scatterChart|absent|absent|0|absent|absent|absent|auto|0|b|1/gap/1",
             Row(0, "0/0", name_mixed, scatter_x, mixed, "absent", "none", kCircle1, "0")}),
       Join({"5 fmt_col" + bar_head,
             Row(0, "0/0", "strRef Sheet1!$C$1 n=1 0:Fmt", cat8,
                 "numRef Sheet1!$C$2:$C$9 fc=General n=8 0:0.30000000000000004 1:0.3333333333333333 2:1e-20 "
                 "3:1.2345678901234568e+17 4:-0.5[0.00%] 5:45293[m/d/yy] 6:1234.5[#,##0.00] 7:0.6666666666666666[0.0]",
                 kAccent1, kBarLine)}),
       line_chart("6 hid_vis", "1/gap/1", "strRef Sheet1!$B$11 n=1 0:Hid",
                  "strRef Sheet1!$A$12:$A$19 n=6 0:a 1:b 2:d 3:f 4:g 5:h",
                  "numRef Sheet1!$B$12:$B$19 fc=General n=6 0:10 1:20 2:40 3:60 4:70 5:80"),
       line_chart("7 hid_all", "0/gap/1", "strRef Sheet1!$B$11 n=1 0:Hid",
                  "strRef Sheet1!$A$12:$A$19 n=8 0:a 1:b 2:c 3:d 4:e 5:f 6:g 7:h",
                  "numRef Sheet1!$B$12:$B$19 fc=General n=8 0:10 1:20 2:30 3:40 4:50 5:60 6:70 7:80"),
       line_chart("8 hidcol_vis", "1/gap/1", "strRef Sheet1!$E$1 n=1 0:HidCol", cat8,
                  "numRef Sheet1!$E$2:$E$9 no cache"),
       line_chart("9 hidcol_all", "0/gap/1", "strRef Sheet1!$E$1 n=1 0:HidCol", cat8,
                  "numRef Sheet1!$E$2:$E$9 fc=General n=8 0:1 1:2 2:3 3:4 4:5 5:6 6:7 7:8"),
       line_chart("10 blank_zero", "1/zero/1", name_mixed, cat8, mixed),
       line_chart("11 blank_span", "1/span/1", name_mixed, cat8, mixed),
       Join({"12 multilvl" + bar_head,
             Row(0, "0/0", "strRef Sheet1!$H$1 n=1 0:MV",
                 "multiLvlStrRef Sheet1!$F$2:$G$9 n=8 0:i1 1:i2 2:i3 3:i4 4:i5 5:i6 6:i7 7:i8 up=[0:G1 3:G2 5:G3]",
                 "numRef Sheet1!$H$2:$H$9 fc=General n=8 0:5 1:6 2:7 3:8 4:9 5:10 6:11 7:12", kAccent1, kBarLine)}),
       Join({"13 scatter_textx|scatterChart|absent|absent|0|absent|absent|absent|auto|0|b|1/gap/1",
             Row(0, "0/0", "strRef Sheet1!$J$1 n=1 0:Ys", "strRef Sheet1!$I$2:$I$9 n=8 0:1 1:2 2:x 3:4 4:5 5:6 6:7 7:8",
                 "numRef Sheet1!$J$2:$J$9 fc=General n=8 0:3 1:1 2:4 3:1 4:5 5:9 6:2 7:6", "absent", "none", kCircle1,
                 "0")})});
}

TEST(ChartXml, CacheFixtureReadsNamesLiteralsAndFormatCodes) {
  const std::string cat_mix = "strRef Sheet1!$A$2:$A$5 n=4 0:a 1:5 3:#N/A";
  const std::string val_b = "numRef Sheet1!$B$2:$B$5 fc=General n=4 0:1 1:2 2:3 3:4";
  const std::string bar_head = "|barChart|col|clustered|0|219|-27|absent|auto|0|b|1/gap/1";
  const auto bar = [&](const std::string& id_name, const std::string& name, const std::string& cat,
                       const std::string& val) {
    return Join({id_name + bar_head, Row(0, "0/0", name, cat, val, kAccent1, kBarLine)});
  };
  ExpectFixture(
      "chart_cache.xlsx",
      {Join(
           {"2 catmix|lineChart|absent|standard|0|absent|absent|absent|auto|0|b|1/gap/1",
            Row(0, "0/0", "strRef Sheet1!$B$1 n=1 0:V", "strRef Sheet1!$A$2:$A$7 n=6 0:a 1:5 3:#N/A 4:TRUE 5:2024/3/4",
                "numRef Sheet1!$B$2:$B$7 fc=General n=6 0:1 1:2 2:3 3:4 4:5 5:6", "absent", kLine1, kNoneMarker, "0")}),
       bar("3 pct_all", "strRef Sheet1!$C$1 n=1 0:Pct", cat_mix,
           "numRef Sheet1!$C$2:$C$5 fc=0.0% n=4 0:0.1 1:0.2 2:0.3 3:0.4"),
       bar("4 first_fmt", "strRef Sheet1!$D$1 n=1 0:FirstFmt", cat_mix,
           "numRef Sheet1!$D$2:$D$5 fc=General n=4 0:1.5[0.00] 1:2.5 2:3.5 3:4.5"),
       bar("5 name_num", "strRef Sheet1!$E$1 n=1 0:2024", cat_mix, val_b),
       bar("6 name_blank", "strRef Sheet1!$F$1 n=1", cat_mix, val_b),
       bar("7 name_literal", "lit n=1 0:Lit", cat_mix, val_b),
       bar("8 vals_text", "strRef Sheet1!$G$1 n=1 0:AllText", cat_mix,
           "numRef Sheet1!$G$2:$G$5 fc=General n=4 0:0 1:0 2:0 3:0"),
       Join({"9 date_cat|lineChart|absent|standard|0|absent|absent|absent|auto|0|b|1/gap/1",
             Row(0, "0/0", "strRef Sheet1!$I$1 n=1 0:DV",
                 "numRef Sheet1!$H$2:$H$5 fc=m/d/yy n=4 0:45292 1:45293 2:45296 3:45323",
                 "numRef Sheet1!$I$2:$I$5 fc=General n=4 0:3 1:1 2:4 3:1", "absent", kLine1, kNoneMarker, "0")}),
       bar("10 num_cat", "strRef Sheet1!$K$1 n=1 0:NV", "numRef Sheet1!$J$2:$J$5 fc=General n=4 0:10 1:20 2:30 3:40",
           "numRef Sheet1!$K$2:$K$5 fc=General n=4 0:2 1:7 2:1 3:8"),
       bar("11 no_cat", "strRef Sheet1!$K$1 n=1 0:NV", "absent",
           "numRef Sheet1!$K$2:$K$5 fc=General n=4 0:2 1:7 2:1 3:8"),
       bar("12 lit_vals", "lit n=1 0:L", "lit n=2 0:p 1:q", "lit n=2 0:1 1:2")});
}

TEST(ChartXml, LegacyStringValuesDecodeOnceWhilePlainTextStaysRaw) {
  const Workbook wb = Workbook::create();
  const std::string encoded_value = "literal _x0041_ _x005F_x0041_ &amp; &lt;猫&gt; _x000D__x0001_";
  const std::string chart =
      "<c:title><c:tx><c:rich><a:p><a:r><a:t>plain _x0041_ _x005F_x0041_ &amp; &lt;猫&gt;&#13;</a:t>"
      "</a:r></a:p></c:rich></c:tx></c:title>"
      "<c:plotArea><c:barChart><c:barDir val=\"col\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:tx><c:v>" +
      encoded_value +
      "</c:v></c:tx>"
      "<c:cat><c:strRef><c:f>_x0041_!$A$1:$A$2</c:f><c:strCache><c:ptCount val=\"2\"/>"
      "<c:pt idx=\"0\"><c:v>_x0041_</c:v></c:pt>"
      "<c:pt idx=\"1\"><c:v>_x005F_x0041_</c:v></c:pt>"
      "</c:strCache></c:strRef></c:cat>"
      "<c:val><c:numRef><c:f>_x0041_!$B$1:$B$2</c:f><c:numCache><c:ptCount val=\"0\"/>"
      "</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea>";
  const ChartModel model = Read(Wrap(chart, ""), wb);

  ASSERT_EQ(model.title.kind, ChartTitleKind::kText);
  EXPECT_EQ(model.title.text, "plain _x0041_ _x005F_x0041_ & <猫>\r");
  ASSERT_EQ(model.plots.size(), 1U);
  ASSERT_EQ(model.plots[0].series.size(), 1U);
  const ChartSeries& series = model.plots[0].series[0];
  ASSERT_EQ(series.name.kind, ChartDataSource::Kind::kLiteral);
  ASSERT_EQ(series.name.literal.size(), 1U);
  EXPECT_EQ(series.name.literal[0].text, "literal A _x0041_ & <猫> \r\x01");
  EXPECT_EQ(series.cat.formula, "_x0041_!$A$1:$A$2");
  ASSERT_EQ(series.cat.cache.size(), 2U);
  EXPECT_EQ(series.cat.cache[0].text, "A");
  EXPECT_EQ(series.cat.cache[1].text, "_x0041_");
}

TEST(ChartXml, ChartexPlainTextAndPointsKeepOoxmlLookingSpellings) {
  const Workbook wb = Workbook::create();
  const std::string raw =
      "<cx:chartSpace xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\">"
      "<cx:chartData><cx:data id=\"1\"><cx:strDim type=\"cat\"><cx:f>_x0041_!$A$1</cx:f>"
      "<cx:lvl><cx:pt idx=\"0\">_x0041_</cx:pt></cx:lvl></cx:strDim></cx:data></cx:chartData>"
      "<cx:chart><cx:title><cx:tx><cx:txData><cx:v>title _x0041_ &amp; _x005F_x0041_</cx:v>"
      "</cx:txData></cx:tx></cx:title><cx:plotArea><cx:plotAreaRegion><cx:series layoutId=\"line\">"
      "<cx:dataId val=\"1\"/><cx:tx><cx:txData><cx:v>name _x0041_ &amp; _x005F_x0041_</cx:v>"
      "</cx:txData></cx:tx></cx:series></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>";
  const ChartModel model = Read(raw, wb);

  ASSERT_TRUE(model.is_chartex);
  EXPECT_EQ(model.title.text, "title _x0041_ & _x005F_x0041_");
  ASSERT_EQ(model.plots.size(), 1U);
  ASSERT_EQ(model.plots[0].series.size(), 1U);
  const ChartSeries& series = model.plots[0].series[0];
  ASSERT_EQ(series.name.literal.size(), 1U);
  EXPECT_EQ(series.name.literal[0].text, "name _x0041_ & _x005F_x0041_");
  EXPECT_EQ(series.cat.formula, "_x0041_!$A$1");
  ASSERT_EQ(series.cat.cache.size(), 1U);
  EXPECT_EQ(series.cat.cache[0].text, "_x0041_");
}

TEST(ChartXml, RefsFixtureReadsNamesUnionsAndWholeColumns) {
  const std::string name_v1 = "strRef Sheet1!$B$1 n=1 0:V1";
  const std::string bar_head = "|barChart|col|clustered|0|219|-27|absent|auto|0|b|1/gap/1";
  const auto bar = [&](const std::string& id_name, const std::string& head, const std::string& name,
                       const std::string& cat, const std::string& val) {
    return Join({id_name + head, Row(0, "0/0", name, cat, val, kAccent1, kBarLine)});
  };
  ExpectFixture(
      "chart_refs.xlsx",
      {bar("2 wb_name", bar_head, name_v1, kCatAbcd, "numRef [0]!WbVals fc=General n=4 0:1 1:2 2:3 3:7"),
       bar("3 sheet_name", bar_head, "strRef Sheet1!$C$1 n=1 0:V2", kCatAbcd,
           "numRef Sheet1!LocVals fc=General n=4 0:4 1:5 2:6 3:8"),
       bar("4 dyn_name", bar_head, name_v1, "absent", "numRef [0]!DynVals fc=General n=7 0:1 1:2 2:3 3:7 5:11 6:12"),
       bar("5 table_ref", bar_head, "absent", "absent", "numRef Sheet1!$E$2:$E$4 fc=General n=3 0:1 1:3 2:5"),
       bar("6 whole_col", bar_head, name_v1, "absent",
           "numRef Sheet1!$B:$B fc=General n=1048576 0:0 1:1 2:2 3:3 4:7 6:11 7:12 8:13"),
       bar("7 other_sheet", bar_head, "strRef 'My Data'!$B$1 n=1 0:OV", "strRef 'My Data'!$A$2:$A$4 n=3 0:p 1:q 2:r",
           "numRef 'My Data'!$B$2:$B$4 fc=General n=3 0:10 1:20 2:30"),
       bar("8 hidden_sheet", bar_head, "strRef Hid!$B$1 n=1 0:HV", "strRef Hid!$A$2:$A$3 n=2 0:x 1:y",
           "numRef Hid!$B$2:$B$3 fc=General n=2 0:100 1:200"),
       bar("9 merged_cat", bar_head, name_v1, "strRef Sheet1!$A$7:$A$9 n=3 0:m1 2:m3",
           "numRef Sheet1!$B$7:$B$9 fc=General n=3 0:11 1:12 2:13"),
       bar("10 two_d_val", bar_head, "absent", "absent", "numRef Sheet1!$B$2:$B$3 fc=General n=2 0:1 1:2"),
       bar("11 multi_name", bar_head, "strRef Sheet1!$B$1:$C$1 n=1 0:V1 V2", kCatAbcd, kValB),
       bar("12 union", bar_head, name_v1, "strRef (Sheet1!$A$2:$A$3,Sheet1!$A$5) n=3 0:a 1:b 2:d",
           "numRef (Sheet1!$B$2:$B$3,Sheet1!$B$5) fc=General n=3 0:1 1:2 2:7")});
}

TEST(ChartXml, ComboFixtureHasTwoPlotsAndDoughnut) {
  ExpectFixture(
      "chart_combo.xlsx",
      {Join({"2 combo|barChart+lineChart|col|clustered|0|219|-27|absent|none|0|b|1/gap/1",
             Row(0, "0/0", "strRef Sheet1!$B$1 n=1 0:Sales", kCatAbcd,
                 "numRef Sheet1!$B$2:$B$5 fc=General n=4 0:100 1:200 2:150 3:300", kAccent1, kBarLine),
             Row(1, "1/1", "strRef Sheet1!$C$1 n=1 0:Rate", kCatAbcd,
                 "numRef Sheet1!$C$2:$C$5 fc=General n=4 0:0.1 1:0.25 2:0.2 3:0.4", "absent", kLine2, kCircle2, "0")}),
       Join({"3 doughnut|doughnutChart|absent|absent|1|absent|absent|0|auto|0|b|1/gap/1|hole=75",
             Row(0, "0/0", "strRef Sheet1!$B$1 n=1 0:Sales", kCatAbcd,
                 "numRef Sheet1!$B$2:$B$5 fc=General n=4 0:100 1:200 2:150 3:300", "absent", "absent", "absent",
                 "absent", kPieDots)})});
}

TEST(ChartXml, ChartsheetFixtureListsBothCharts) {
  const std::string val_b = "numRef Sheet1!$B$2:$B$5 fc=General n=4 0:1 1:2 2:-3 3:7";
  ExpectFixture(
      "chart_chartsheet.xlsx",
      {Join({"2 \xE3\x82\xB0\xE3\x83\xA9\xE3\x83\x95 1|barChart|col|clustered|0|219|-27|absent|none|0|b|1/gap/1",
             Row(0, "0/0", kNameS1, kCatAbcd, val_b, kAccent1, kBarLine),
             Row(0, "1/1", kNameS2, kCatAbcd, kValC, kAccent2, kBarLine)}),
       Join({"3 embedded_line|lineChart|absent|standard|0|absent|absent|absent|none|0|b|1/gap/1",
             Row(0, "0/0", kNameS1, kCatAbcd, val_b, "absent", kLine1, kNoneMarker, "0"),
             Row(0, "1/1", kNameS2, kCatAbcd, kValC, "absent", kLine2, kNoneMarker, "0")})});
}

TEST(ChartXml, PivotFixtureIsFlaggedAndHasNoSeriesColour) {
  ExpectFixture("chart_pivot.xlsx",
                {Join({"2 pivot_chart|barChart|col|clustered|0|150|absent|absent|auto|0|r|1/gap/1|pivot",
                       Row(0, "0/0", "strRef Pivot!$B$3 n=1 0:\xE9\x9B\x86\xE8\xA8\x88",
                           "strRef Pivot!$A$4:$A$7 n=3 0:East 1:North 2:West",
                           "numRef Pivot!$B$4:$B$7 fc=General n=3 0:40 1:7 2:25", "absent", "absent")})});
}

TEST(ChartXml, ChartexFixtureBindsDataThroughDataId) {
  const Workbook wb = LoadFixture("chart_chartex.xlsx");
  auto refs = all_chart_parts(wb);
  ASSERT_TRUE(static_cast<bool>(refs));
  ASSERT_EQ(refs.value().size(), 1U);
  EXPECT_TRUE(refs.value()[0].ref.is_chartex);
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, refs.value()[0].ref.part_path);
  ASSERT_NE(part, nullptr);
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_chart_part(part->bytes, doc)));
  auto model = read_chart_model(wb, doc);
  ASSERT_TRUE(static_cast<bool>(model));
  const ChartModel& m = model.value();
  EXPECT_TRUE(m.is_chartex);
  EXPECT_FALSE(m.is_pivot_chart);
  EXPECT_EQ(m.title.kind, ChartTitleKind::kNone);
  ASSERT_TRUE(m.legend.present);
  EXPECT_EQ(m.legend.position, ChartLegendPosition::kTop);
  EXPECT_FALSE(m.legend.overlay);
  ASSERT_EQ(m.plots.size(), 1U);
  EXPECT_EQ(m.plots[0].type, ChartPlotType::kChartex);
  EXPECT_EQ(m.plots[0].chartex_layout, "waterfall");
  ASSERT_EQ(m.plots[0].series.size(), 1U);
  const ChartSeries& s = m.plots[0].series[0];
  EXPECT_EQ(Src(s.name), "cxDim _xlchart.v1.1 n=1 0:Amt");
  EXPECT_EQ(Src(s.cat), "cxDim _xlchart.v1.0 no cache");
  EXPECT_EQ(Src(s.val), "cxDim _xlchart.v1.2 no cache");
  EXPECT_EQ(Src(s.x), "absent");
  ASSERT_TRUE(s.labels.has_value());
  EXPECT_EQ(LabelsStr(s.labels), "show=Val pos=" + std::to_string(static_cast<int>(ChartLabelPosition::kOutsideEnd)));
  ASSERT_EQ(m.axes.size(), 2U);
  EXPECT_EQ(m.axes[0].id, 0U);
  EXPECT_EQ(m.axes[0].kind, ChartAxisKind::kCategory);
  EXPECT_EQ(m.axes[1].id, 1U);
  EXPECT_EQ(m.axes[1].kind, ChartAxisKind::kValue);
  EXPECT_TRUE(m.axes[1].major_gridlines);
  EXPECT_FALSE(m.axes[0].major_gridlines);
}

TEST(ChartXml, FixtureAxesLegendAndStyleAreRead) {
  const Workbook wb = LoadFixture("chart_format.xlsx");
  auto refs = all_chart_parts(wb);
  ASSERT_TRUE(static_cast<bool>(refs));
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, refs.value()[0].ref.part_path);
  ASSERT_NE(part, nullptr);
  pugi::xml_document doc;
  ASSERT_TRUE(static_cast<bool>(parse_chart_part(part->bytes, doc)));
  auto model = read_chart_model(wb, doc);
  ASSERT_TRUE(static_cast<bool>(model));
  const ChartModel& m = model.value();
  EXPECT_FALSE(m.date1904);
  EXPECT_EQ(m.style, 102);
  EXPECT_EQ(m.disp_na_as_blank, Tri::kTrue);
  ASSERT_TRUE(m.plots[0].labels.has_value());
  EXPECT_EQ(LabelsStr(m.plots[0].labels), "show=");
  EXPECT_EQ(m.plots[0].axis_ids, (std::vector<std::uint32_t>{1978249471U, 1978241407U}));
  ASSERT_EQ(m.axes.size(), 2U);
  const ChartAxis& cat = m.axes[0];
  EXPECT_EQ(cat.kind, ChartAxisKind::kCategory);
  EXPECT_EQ(cat.id, 1978249471U);
  EXPECT_EQ(cat.cross_ax, 1978241407U);
  EXPECT_EQ(cat.position, ChartAxisPosition::kBottom);
  EXPECT_EQ(cat.crosses, ChartAxisCrosses::kAutoZero);
  EXPECT_FALSE(cat.deleted);
  EXPECT_FALSE(cat.reversed);
  EXPECT_TRUE(std::isnan(cat.min));
  EXPECT_TRUE(std::isnan(cat.log_base));
  EXPECT_TRUE(std::isnan(cat.crosses_at));
  ASSERT_TRUE(cat.num_fmt.has_value());
  EXPECT_EQ(cat.num_fmt->code, "General");
  EXPECT_TRUE(cat.num_fmt->source_linked);
  EXPECT_EQ(cat.tick_label_pos, "nextTo");
  EXPECT_EQ(cat.major_tick, "none");
  EXPECT_FALSE(cat.major_gridlines);
  const ChartAxis& val = m.axes[1];
  EXPECT_EQ(val.kind, ChartAxisKind::kValue);
  EXPECT_EQ(val.position, ChartAxisPosition::kLeft);
  EXPECT_TRUE(val.major_gridlines);
  EXPECT_EQ(val.cross_ax, 1978249471U);
}

// ---------------------------------------------------------------------------
// Constructed parts
// ---------------------------------------------------------------------------

std::string Wrap(const std::string& chart_children, const std::string& space_children = "") {
  return "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
         "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\" "
         "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" "
         "xmlns:c15=\"http://schemas.microsoft.com/office/drawing/2012/chart\">" +
         space_children + "<c:chart>" + chart_children + "</c:chart></c:chartSpace>";
}

ChartModel Read(const std::string& xml, const Workbook& wb) {
  pugi::xml_document doc;
  auto parsed = parse_chart_part(std::vector<std::uint8_t>(xml.begin(), xml.end()), doc);
  EXPECT_TRUE(static_cast<bool>(parsed));
  auto model = read_chart_model(wb, doc);
  EXPECT_TRUE(static_cast<bool>(model));
  return model ? std::move(model.value()) : ChartModel{};
}

TEST(ChartXml, RejectsPartsThatAreNotChartSpaces) {
  pugi::xml_document doc;
  const std::string broken = "<c:chartSpace";
  auto bad_xml = parse_chart_part(std::vector<std::uint8_t>(broken.begin(), broken.end()), doc);
  ASSERT_FALSE(static_cast<bool>(bad_xml));
  EXPECT_EQ(bad_xml.error().code, FormulonErrorCode::kIoChartUnparseable);

  const std::string wrong_root = "<xdr:wsDr xmlns:xdr=\"urn:x\"/>";
  auto bad_root = parse_chart_part(std::vector<std::uint8_t>(wrong_root.begin(), wrong_root.end()), doc);
  ASSERT_FALSE(static_cast<bool>(bad_root));
  EXPECT_EQ(bad_root.error().code, FormulonErrorCode::kIoChartUnparseable);

  const std::string empty_space = "<c:chartSpace xmlns:c=\"urn:c\"/>";
  EXPECT_TRUE(
      static_cast<bool>(parse_chart_part(std::vector<std::uint8_t>(empty_space.begin(), empty_space.end()), doc)));
}

TEST(ChartXml, SingleVisibleSeriesWithTitleNotDeletedHasAutomaticTitle) {
  const Workbook wb = Workbook::create();
  const std::string one =
      "<c:autoTitleDeleted val=\"0\"/><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/>"
      "<c:order val=\"0\"/></c:ser></c:barChart></c:plotArea>";
  EXPECT_EQ(Read(Wrap(one), wb).title.kind, ChartTitleKind::kAuto);

  const std::string two =
      "<c:autoTitleDeleted val=\"0\"/><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/>"
      "<c:order val=\"0\"/></c:ser><c:ser><c:idx val=\"1\"/><c:order val=\"1\"/></c:ser>"
      "</c:barChart></c:plotArea>";
  EXPECT_EQ(Read(Wrap(two), wb).title.kind, ChartTitleKind::kNone);

  const std::string deleted =
      "<c:autoTitleDeleted val=\"1\"/><c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/>"
      "<c:order val=\"0\"/></c:ser></c:barChart></c:plotArea>";
  EXPECT_EQ(Read(Wrap(deleted), wb).title.kind, ChartTitleKind::kNone);
}

TEST(ChartXml, AbsentElementsStayAbsent) {
  const Workbook wb = Workbook::create();
  const ChartModel m = Read(Wrap("<c:plotArea><c:barChart><c:ser><c:idx val=\"3\"/><c:order val=\"4\"/></c:ser>"
                                 "</c:barChart></c:plotArea>"),
                            wb);
  EXPECT_EQ(m.title.kind, ChartTitleKind::kNone);
  EXPECT_EQ(m.auto_title_deleted, Tri::kAbsent);
  EXPECT_EQ(m.plot_vis_only, Tri::kAbsent);
  EXPECT_EQ(m.disp_blanks_as, ChartBlanksAs::kAbsent);
  EXPECT_FALSE(m.legend.present);
  EXPECT_FALSE(m.style.has_value());
  EXPECT_FALSE(m.date1904);
  ASSERT_EQ(m.plots.size(), 1U);
  const ChartPlot& p = m.plots[0];
  EXPECT_EQ(p.bar_dir, ChartBarDirection::kAbsent);
  EXPECT_FALSE(p.grouping.has_value());
  EXPECT_EQ(p.vary_colors, Tri::kAbsent);
  EXPECT_FALSE(p.gap_width.has_value());
  EXPECT_FALSE(p.overlap.has_value());
  ASSERT_EQ(p.series.size(), 1U);
  const ChartSeries& s = p.series[0];
  EXPECT_EQ(s.idx, 3U);
  EXPECT_EQ(s.order, 4U);
  EXPECT_EQ(s.name.kind, ChartDataSource::Kind::kNone);
  EXPECT_FALSE(s.fill.color.has_value());
  EXPECT_FALSE(s.line.width_emu.has_value());
  EXPECT_FALSE(s.line.none);
  EXPECT_FALSE(s.marker.has_value());
  EXPECT_EQ(s.smooth, Tri::kAbsent);
  EXPECT_FALSE(s.labels.has_value());
}

TEST(ChartXml, ChartLevelFlagsTitleLegendAndStyle) {
  const Workbook wb = Workbook::create();
  const std::string space =
      "<c:date1904 val=\"1\"/><mc:AlternateContent "
      "xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\">"
      "<mc:Choice xmlns:c14=\"urn:c14\" Requires=\"c14\"><c14:style val=\"104\"/></mc:Choice>"
      "<mc:Fallback><c:style val=\"4\"/></mc:Fallback></mc:AlternateContent>";
  const ChartModel m = Read(
      Wrap("<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Sales </a:t></a:r><a:r><a:t>2024</a:t></a:r></a:p>"
           "<a:p><a:r><a:t>second</a:t></a:r></a:p></c:rich></c:tx></c:title>"
           "<c:autoTitleDeleted val=\"1\"/><c:plotArea><c:pieChart/></c:plotArea>"
           "<c:legend><c:legendPos val=\"tr\"/><c:legendEntry><c:idx val=\"2\"/><c:delete val=\"1\"/></c:legendEntry>"
           "<c:legendEntry><c:idx val=\"5\"/><c:delete val=\"0\"/></c:legendEntry><c:overlay val=\"1\"/></c:legend>"
           "<c:plotVisOnly val=\"0\"/><c:dispBlanksAs val=\"span\"/>",
           space),
      wb);
  EXPECT_TRUE(m.date1904);
  EXPECT_EQ(m.style, 104);
  EXPECT_EQ(m.title.kind, ChartTitleKind::kText);
  EXPECT_EQ(m.title.text, "Sales 2024\nsecond");
  EXPECT_EQ(m.auto_title_deleted, Tri::kTrue);
  EXPECT_EQ(m.plot_vis_only, Tri::kFalse);
  EXPECT_EQ(m.disp_blanks_as, ChartBlanksAs::kSpan);
  ASSERT_TRUE(m.legend.present);
  EXPECT_EQ(m.legend.position, ChartLegendPosition::kTopRight);
  EXPECT_TRUE(m.legend.overlay);
  EXPECT_EQ(m.legend.hidden_series_idx, (std::vector<std::uint32_t>{2U}));

  const ChartModel plain =
      Read(Wrap("<c:title><c:tx><c:strRef><c:f>Sheet1!$A$1</c:f></c:strRef></c:tx></c:title><c:plotArea><c:pieChart/>"
                "</c:plotArea><c:legend/>"),
           wb);
  EXPECT_EQ(plain.title.kind, ChartTitleKind::kRef);
  EXPECT_EQ(plain.title.ref_formula, "Sheet1!$A$1");
  EXPECT_EQ(plain.legend.position, ChartLegendPosition::kRight);
  EXPECT_FALSE(plain.style.has_value());

  // A style written only as c:style, and the workbook's date system when the chart omits c:date1904.
  Workbook wb1904 = Workbook::create();
  wb1904.set_date1904(true);
  const ChartModel inherited = Read(Wrap("<c:plotArea/>", "<c:style val=\"7\"/>"), wb1904);
  EXPECT_TRUE(inherited.date1904);
  EXPECT_EQ(inherited.style, 7);
  EXPECT_TRUE(inherited.plots.empty());
}

TEST(ChartXml, AxisScalingAndKinds) {
  const Workbook wb = Workbook::create();
  const ChartModel m = Read(
      Wrap("<c:plotArea><c:lineChart><c:axId val=\"10\"/><c:axId val=\"20\"/></c:lineChart>"
           "<c:dateAx><c:axId val=\"10\"/><c:scaling><c:orientation val=\"maxMin\"/></c:scaling><c:delete val=\"1\"/>"
           "<c:axPos val=\"t\"/><c:title><c:tx><c:rich><a:p><a:r><a:t>Day</a:t></a:r></a:p></c:rich></c:tx></c:title>"
           "<c:numFmt formatCode=\"d-mmm\" sourceLinked=\"0\"/><c:crossAx val=\"20\"/><c:crosses val=\"max\"/>"
           "<c:baseTimeUnit val=\"days\"/><c:minorGridlines/></c:dateAx>"
           "<c:valAx><c:axId val=\"20\"/><c:scaling><c:logBase val=\"10\"/><c:orientation val=\"minMax\"/>"
           "<c:max val=\"1000\"/><c:min val=\"1\"/></c:scaling><c:axPos val=\"r\"/><c:majorGridlines/>"
           "<c:tickLblPos val=\"high\"/><c:crossAx val=\"10\"/><c:crossesAt val=\"2.5\"/><c:majorUnit val=\"10\"/>"
           "<c:minorUnit val=\"2\"/></c:valAx><c:serAx><c:axId val=\"30\"/></c:serAx></c:plotArea>"),
      wb);
  ASSERT_EQ(m.axes.size(), 3U);
  const ChartAxis& date = m.axes[0];
  EXPECT_EQ(date.kind, ChartAxisKind::kDate);
  EXPECT_TRUE(date.reversed);
  EXPECT_TRUE(date.deleted);
  EXPECT_EQ(date.position, ChartAxisPosition::kTop);
  EXPECT_EQ(date.title.kind, ChartTitleKind::kText);
  EXPECT_EQ(date.title.text, "Day");
  ASSERT_TRUE(date.num_fmt.has_value());
  EXPECT_EQ(date.num_fmt->code, "d-mmm");
  EXPECT_FALSE(date.num_fmt->source_linked);
  EXPECT_EQ(date.crosses, ChartAxisCrosses::kMax);
  EXPECT_EQ(date.base_time_unit, "days");
  EXPECT_TRUE(date.minor_gridlines);
  const ChartAxis& val = m.axes[1];
  EXPECT_EQ(val.kind, ChartAxisKind::kValue);
  EXPECT_EQ(val.log_base, 10.0);
  EXPECT_EQ(val.min, 1.0);
  EXPECT_EQ(val.max, 1000.0);
  EXPECT_EQ(val.major_unit, 10.0);
  EXPECT_EQ(val.minor_unit, 2.0);
  EXPECT_EQ(val.crosses, ChartAxisCrosses::kAt);
  EXPECT_EQ(val.crosses_at, 2.5);
  EXPECT_EQ(val.position, ChartAxisPosition::kRight);
  EXPECT_EQ(val.tick_label_pos, "high");
  EXPECT_TRUE(val.major_gridlines);
  EXPECT_EQ(m.axes[2].kind, ChartAxisKind::kSeries);
  EXPECT_EQ(m.axes[2].id, 30U);
  EXPECT_EQ(m.axes[2].crosses, ChartAxisCrosses::kAbsent);
}

TEST(ChartXml, SeriesFormatForms) {
  const Workbook wb = Workbook::create();
  const std::string xml = Wrap(
      "<c:plotArea><c:lineChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:spPr><a:solidFill><a:schemeClr val=\"accent3\"><a:alpha val=\"25000\"/></a:schemeClr></a:solidFill>"
      "<a:ln w=\"12700\"><a:solidFill><a:sysClr val=\"windowText\" lastClr=\"112233\"/></a:solidFill>"
      "<a:prstDash val=\"sysDot\"/></a:ln></c:spPr>"
      "<c:marker><c:symbol val=\"star\"/><c:size val=\"12\"/></c:marker>"
      "<c:dPt><c:idx val=\"2\"/><c:spPr><a:solidFill><a:srgbClr val=\"AABBCC\"/></a:solidFill>"
      "<a:ln><a:noFill/></a:ln></c:spPr></c:dPt>"
      "<c:dLbls><c:numFmt formatCode=\"0.0%\" sourceLinked=\"0\"/><c:spPr><a:noFill/></c:spPr>"
      "<c:txPr><a:bodyPr/><a:p><a:pPr><a:defRPr sz=\"1200\"><a:solidFill><a:srgbClr val=\"FF0000\"/></a:solidFill>"
      "</a:defRPr></a:pPr></a:p></c:txPr><c:dLblPos val=\"outEnd\"/><c:showLegendKey val=\"1\"/><c:showVal val=\"0\"/>"
      "<c:showCatName val=\"0\"/><c:showSerName val=\"1\"/><c:showPercent val=\"1\"/><c:separator> </c:separator>"
      "</c:dLbls><c:smooth val=\"1\"/></c:ser>"
      "<c:ser><c:idx val=\"1\"/><c:order val=\"1\"/><c:spPr><a:ln><a:noFill/></a:ln></c:spPr>"
      "<c:marker><c:symbol val=\"auto\"/></c:marker><c:dLbls><c:delete val=\"1\"/></c:dLbls></c:ser>"
      "<c:ser><c:idx val=\"2\"/><c:order val=\"2\"/><c:marker><c:size val=\"3\"/></c:marker><c:smooth/></c:ser>"
      "</c:lineChart></c:plotArea>");
  const ChartModel m = Read(xml, wb);
  ASSERT_EQ(m.plots.size(), 1U);
  ASSERT_EQ(m.plots[0].series.size(), 3U);
  const ChartSeries& s0 = m.plots[0].series[0];
  EXPECT_EQ(FillStr(s0.fill), "FF196B24 alpha=25000");
  EXPECT_EQ(s0.line.width_emu, 12700);
  EXPECT_EQ(s0.line.color, 0xFF112233U);
  EXPECT_EQ(s0.line.dash_name, "sysDot");
  EXPECT_EQ(s0.line.dash, ChartLineDash::kOther);
  ASSERT_TRUE(s0.marker.has_value());
  EXPECT_EQ(s0.marker->symbol, ChartMarkerSymbol::kStar);
  EXPECT_EQ(s0.marker->symbol_name, "star");
  EXPECT_EQ(s0.marker->size, 12U);
  EXPECT_EQ(s0.smooth, Tri::kTrue);
  ASSERT_EQ(s0.points.size(), 1U);
  EXPECT_EQ(s0.points[0].idx, 2U);
  EXPECT_EQ(s0.points[0].fill.color, 0xFFAABBCCU);
  EXPECT_TRUE(s0.points[0].line.none);
  EXPECT_FALSE(s0.points[0].marker.has_value());
  ASSERT_TRUE(s0.labels.has_value());
  const ChartLabels& l = *s0.labels;
  EXPECT_FALSE(l.show_val);
  EXPECT_FALSE(l.show_cat);
  EXPECT_TRUE(l.show_ser);
  EXPECT_TRUE(l.show_pct);
  EXPECT_TRUE(l.show_legend_key);
  EXPECT_EQ(l.separator, " ");
  EXPECT_EQ(l.position, ChartLabelPosition::kOutsideEnd);
  ASSERT_TRUE(l.num_fmt.has_value());
  EXPECT_EQ(l.num_fmt->code, "0.0%");
  EXPECT_FALSE(l.num_fmt->source_linked);
  EXPECT_EQ(l.font_color, 0xFFFF0000U);
  EXPECT_EQ(l.font_size_hundredths_pt, 1200U);
  EXPECT_FALSE(l.deleted);

  const ChartSeries& s1 = m.plots[0].series[1];
  EXPECT_TRUE(s1.line.none);
  ASSERT_TRUE(s1.marker.has_value());
  EXPECT_EQ(s1.marker->symbol, ChartMarkerSymbol::kAuto);
  ASSERT_TRUE(s1.labels.has_value());
  EXPECT_TRUE(s1.labels->deleted);

  const ChartSeries& s2 = m.plots[0].series[2];
  ASSERT_TRUE(s2.marker.has_value());
  EXPECT_EQ(s2.marker->symbol, ChartMarkerSymbol::kAuto);
  EXPECT_TRUE(s2.marker->symbol_name.empty());
  EXPECT_EQ(s2.smooth, Tri::kTrue);
}

TEST(ChartXml, DataSourceForms) {
  const Workbook wb = Workbook::create();
  const std::string xml = Wrap(
      "<c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:tx><c:v>Typed</c:v></c:tx>"
      "<c:cat><c:strLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>x</c:v></c:pt><c:pt idx=\"2\"><c:v>z</c:v></c:pt>"
      "</c:strLit></c:cat>"
      "<c:val><c:numLit><c:formatCode>0.00</c:formatCode><c:ptCount val=\"2\"/><c:pt idx=\"0\"><c:v>1.5</c:v></c:pt>"
      "<c:pt idx=\"1\"><c:v>#N/A</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea>");
  const ChartModel m = Read(xml, wb);
  const ChartSeries& s = m.plots[0].series[0];
  EXPECT_EQ(Src(s.name), "lit n=1 0:Typed");
  EXPECT_EQ(Src(s.cat), "lit n=3 0:x 2:z");
  EXPECT_EQ(Src(s.val), "lit n=2 0:1.5 1:#N/A");
  EXPECT_EQ(s.val.cache_format_code, "0.00");
  EXPECT_EQ(s.cat.point_count, 3U);
  ASSERT_EQ(s.cat.literal.size(), 2U);
  EXPECT_EQ(s.cat.literal[1].idx, 2U);
  EXPECT_EQ(s.val.literal[1].kind, ChartValue::Kind::kText);
}

TEST(ChartXml, FilteredScatterSeriesAndCategoryFilterByColumn) {
  const Workbook wb = Workbook::create();
  const std::string ref_ext =
      "<c:extLst><c:ext uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\"><c15:fullRef><c15:sqref>Sheet1!$A$1:$A$4"
      "</c15:sqref></c15:fullRef></c:ext></c:extLst>";
  const std::string xml = Wrap(
      "<c:plotArea><c:scatterChart><c:scatterStyle val=\"lineMarker\"/><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/>"
      "<c:xVal><c:numRef>" +
      ref_ext +
      "<c:f>(Sheet1!$A$1:$A$2,Sheet1!$A$4)</c:f></c:numRef></c:xVal>"
      "<c:yVal><c:numRef><c:f>Sheet1!$B$1:$B$4</c:f></c:numRef></c:yVal></c:ser>"
      "<c:extLst><c:ext uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\"><c15:filteredScatterSeries>"
      "<c15:ser><c:idx val=\"1\"/><c:order val=\"1\"/><c:yVal><c:numRef><c:extLst><c:ext "
      "uri=\"{02D57815-91ED-43cb-92C2-25804820EDAC}\">"
      "<c15:formulaRef><c15:sqref>Sheet1!$C$1:$C$4</c15:sqref></c15:formulaRef></c:ext></c:extLst></c:numRef></"
      "c:yVal></c15:ser>"
      "</c15:filteredScatterSeries></c:ext></c:extLst></c:scatterChart></c:plotArea>");
  const ChartModel m = Read(xml, wb);
  ASSERT_EQ(m.plots.size(), 1U);
  EXPECT_EQ(m.plots[0].scatter_style, "lineMarker");
  ASSERT_EQ(m.plots[0].series.size(), 2U);
  EXPECT_FALSE(m.plots[0].series[0].filtered);
  EXPECT_TRUE(m.plots[0].series[1].filtered);
  EXPECT_EQ(m.plots[0].series[1].y.formula, "Sheet1!$C$1:$C$4");
  EXPECT_EQ(m.plots[0].series[0].x.formula, "(Sheet1!$A$1:$A$2,Sheet1!$A$4)");
  EXPECT_EQ(m.category_filter, (std::vector<std::uint32_t>{2U}));
}

TEST(ChartXml, ThemeColoursResolveThroughTheWorkbookTheme) {
  Workbook wb = Workbook::create();
  ThemeColors colors = wb.load_theme().theme.colors;
  colors[4] = 0xFF102030U;  // accent1 in scheme order is index 4
  ASSERT_TRUE(static_cast<bool>(wb.set_theme_colors(colors)));
  const ChartModel m = Read(
      Wrap("<c:plotArea><c:barChart><c:ser><c:idx val=\"0\"/><c:order val=\"0\"/><c:spPr><a:solidFill>"
           "<a:schemeClr val=\"accent1\"/></a:solidFill><a:ln><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill>"
           "</a:ln></c:spPr></c:ser></c:barChart></c:plotArea>"),
      wb);
  const ChartSeries& s = m.plots[0].series[0];
  EXPECT_EQ(s.fill.color, 0xFF102030U);
  EXPECT_EQ(s.line.color, wb.load_theme().theme.colors[1]);  // bg1 = lt1
}

}  // namespace
}  // namespace formulon
