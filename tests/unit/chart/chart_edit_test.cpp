#include "chart/chart_edit.h"

#include <algorithm>
#include <cmath>
#include <cstdint>
#include <cstdio>
#include <regex>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_cache.h"
#include "chart/chart_model.h"
#include "chart/chart_template.h"
#include "drawing/drawing_edit.h"
#include "drawing/drawing_xml.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/xlsb/reader.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

using Code = FormulonErrorCode;
using D = ChartSeriesDirection;
using H = ChartHeaderMode;
using K = ChartInsertKind;

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
      out.clear();
    }
  }
  std::fclose(file);
  return out;
}

Workbook LoadFixture(const std::string& name) {
  const std::vector<std::uint8_t> bytes = ReadFileBytes(std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name);
  Workbook wb = Workbook::create_empty();
  const io::ByteSpan span{bytes.data(), bytes.size()};
  if (name.substr(name.size() - 5) == ".xlsb") {
    auto read = io::xlsb::read_xlsb(span);
    EXPECT_TRUE(static_cast<bool>(read));
    if (read) {
      wb = std::move(read.value().workbook);
    }
  } else {
    auto read = io::read_ooxml(span);
    EXPECT_TRUE(static_cast<bool>(read));
    if (read) {
      wb = std::move(read.value().workbook);
    }
  }
  wb.set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

struct Located {
  std::size_t sheet = 0;
  std::uint32_t id = 0;
  std::string path;
};

/// The chart whose part is `part` (`xl/charts/chartN.xml`), or the first chart when empty.
Located Find(const Workbook& wb, const std::string& part = {}) {
  auto all = all_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(all));
  if (all) {
    for (const SheetChartPartRef& e : all.value()) {
      if (part.empty() || e.ref.part_path == part) {
        return Located{e.sheet, e.ref.object_id, e.ref.part_path};
      }
    }
  }
  ADD_FAILURE() << "no chart " << part;
  return {};
}

std::string PartText(const Workbook& wb, const std::string& path) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  EXPECT_NE(part, nullptr) << path;
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

std::vector<std::pair<std::string, std::vector<std::uint8_t>>> Parts(const Workbook& wb) {
  std::vector<std::pair<std::string, std::vector<std::uint8_t>>> out;
  for (const PassthroughPart& p : wb.passthrough_parts()) {
    out.emplace_back(p.path, p.bytes);
  }
  return out;
}

/// The `n`th substring from `open` through `close` (inclusive).
std::string Between(const std::string& text, const std::string& open, const std::string& close, int n = 0) {
  std::size_t at = 0;
  for (int i = 0; i <= n; ++i) {
    at = text.find(open, i == 0 ? 0 : at + 1);
    if (at == std::string::npos) {
      return {};
    }
  }
  const std::size_t end = text.find(close, at);
  return end == std::string::npos ? std::string() : text.substr(at, end + close.size() - at);
}

std::string NoUid(const std::string& text) {
  return std::regex_replace(text, std::regex("<c16:uniqueId val=\"[^\"]*\"/>"), "<c16:uniqueId/>");
}

/// Equal and non-empty, so a missing element cannot pass vacuously.
void ExpectSameForm(const std::string& got, const std::string& want) {
  EXPECT_FALSE(want.empty());
  EXPECT_EQ(got, want);
}

bool NoRefresh(const Workbook& wb) {
  auto got = refreshed_chart_parts(wb);
  return got && got.value().empty();
}

template <typename T>
T Must(Expected<T, Error> got) {
  EXPECT_TRUE(static_cast<bool>(got)) << (got ? "" : got.error().message);
  return got ? std::move(got.value()) : T{};
}

void MustOk(const Expected<void, Error>& got) {
  EXPECT_TRUE(static_cast<bool>(got)) << (got ? "" : got.error().message);
}

template <typename E>
void ExpectCode(const E& got, Code code) {
  ASSERT_FALSE(static_cast<bool>(got));
  EXPECT_EQ(got.error().code, code) << got.error().message;
}

/// Sheet1: A1 blank, B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs.
Workbook GridBook() {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  const char* names[] = {"S1", "S2", "S3"};
  const char* cats[] = {"a", "b", "c", "d"};
  for (std::uint32_t c = 0; c < 3; ++c) {
    s.set_cell_text(0, c + 1, names[c]);
  }
  for (std::uint32_t r = 0; r < 4; ++r) {
    s.set_cell_text(r + 1, 0, cats[r]);
    for (std::uint32_t c = 0; c < 3; ++c) {
      s.set_cell_value(r + 1, c + 1, Value::number(static_cast<double>((r + 1) * (c + 2))));
    }
    s.set_cell_value(r + 1, 5, Value::number(static_cast<double>(r * 3)));
    s.set_cell_value(r + 1, 6, Value::number(static_cast<double>(r * r)));
  }
  s.set_cell_text(0, 5, "X");
  s.set_cell_text(0, 6, "Y");
  return wb;
}

std::uint32_t Insert(Workbook& wb, K kind, ChartGrouping grouping = ChartGrouping::kClustered,
                     const std::string& source = "Sheet1!$A$1:$D$5") {
  ChartInsertOptions o;
  o.kind = kind;
  o.grouping = grouping;
  o.source = kind == K::kScatter && source == "Sheet1!$A$1:$D$5" ? "Sheet1!$F$1:$G$5" : source;
  return Must(insert_chart(wb, 0, o));
}

ChartPlotType PlotType(K kind) {
  static const ChartPlotType kTypes[] = {ChartPlotType::kBar,     ChartPlotType::kBar, ChartPlotType::kLine,
                                         ChartPlotType::kArea,    ChartPlotType::kPie, ChartPlotType::kScatter,
                                         ChartPlotType::kDoughnut};
  return kTypes[static_cast<int>(kind)];
}

void ExpectKind(const ChartPlot& plot, K kind, ChartGrouping grouping) {
  EXPECT_EQ(plot.type, PlotType(kind));
  if (kind == K::kColumn || kind == K::kBar) {
    EXPECT_EQ(plot.bar_dir, kind == K::kBar ? ChartBarDirection::kBar : ChartBarDirection::kCol);
    EXPECT_EQ(plot.grouping, grouping == ChartGrouping::kStandard ? ChartGrouping::kClustered : grouping);
  } else if (kind == K::kLine || kind == K::kArea) {
    EXPECT_EQ(plot.grouping, grouping == ChartGrouping::kClustered ? ChartGrouping::kStandard : grouping);
  } else {
    EXPECT_FALSE(plot.grouping.has_value());
  }
}

std::vector<std::string> ValueRefs(const ChartModel& m) {
  std::vector<std::string> out;
  for (const ChartSeries& s : m.plots[0].series) {
    out.push_back(m.plots[0].type == ChartPlotType::kScatter ? s.y.formula : s.val.formula);
  }
  return out;
}

void ExpectSameData(const ChartData& a, const ChartData& b) {
  ASSERT_EQ(a.series.size(), b.series.size());
  EXPECT_EQ(a.categories.labels, b.categories.labels);
  EXPECT_EQ(a.point_count, b.point_count);
  for (std::size_t i = 0; i < a.series.size(); ++i) {
    EXPECT_EQ(a.series[i].name, b.series[i].name);
    ASSERT_EQ(a.series[i].points.size(), b.series[i].points.size());
    for (std::size_t p = 0; p < a.series[i].points.size(); ++p) {
      EXPECT_EQ(a.series[i].points[p].value, b.series[i].points[p].value);
      EXPECT_EQ(a.series[i].points[p].state, b.series[i].points[p].state);
    }
    ASSERT_EQ(a.series[i].xy.size(), b.series[i].xy.size());
  }
}

// ---------------------------------------------------------------------------
// Insert and type change (pairwise: kind x grouping x direction x headers x target kind)
// ---------------------------------------------------------------------------

TEST(ChartEdit, InsertAndChangeTypePairwise) {
  static const int kRows[][6] = {
      {4, 2, 2, 2, 2, 3}, {1, 2, 0, 1, 1, 0}, {0, 0, 2, 1, 0, 2}, {0, 1, 1, 0, 1, 3}, {0, 0, 0, 2, 1, 4},
      {0, 2, 1, 0, 2, 1}, {5, 1, 0, 0, 0, 2}, {5, 1, 2, 1, 2, 5}, {1, 0, 1, 2, 0, 5}, {2, 2, 1, 1, 0, 4},
      {3, 1, 2, 2, 1, 4}, {3, 0, 1, 0, 2, 0}, {0, 1, 2, 0, 0, 0}, {0, 0, 1, 1, 0, 1}, {2, 0, 0, 0, 2, 3},
      {5, 0, 1, 2, 1, 0}, {4, 1, 2, 1, 0, 1}, {0, 0, 0, 0, 1, 5}, {3, 2, 0, 1, 2, 3}, {1, 1, 2, 0, 2, 4},
      {4, 0, 1, 2, 2, 2}, {2, 1, 0, 2, 1, 1}, {2, 2, 2, 0, 1, 2}, {5, 2, 1, 2, 0, 3}, {2, 1, 0, 1, 0, 0},
      {2, 2, 2, 0, 0, 5}, {4, 1, 2, 1, 0, 0}, {4, 0, 0, 0, 1, 4}, {4, 2, 1, 1, 2, 5}, {3, 0, 0, 1, 0, 2},
      {1, 1, 2, 1, 2, 2}, {3, 0, 1, 1, 0, 1}, {5, 2, 1, 1, 2, 4}, {5, 0, 1, 1, 0, 1}, {1, 2, 0, 0, 1, 1},
      {3, 2, 2, 0, 0, 5}, {1, 0, 2, 1, 2, 3}};
  for (const auto& row : kRows) {
    SCOPED_TRACE(testing::Message() << row[0] << row[1] << row[2] << row[3] << row[4] << row[5]);
    Workbook wb = GridBook();
    ChartInsertOptions o;
    o.kind = static_cast<K>(row[0]);
    o.grouping = static_cast<ChartGrouping>(row[1]);
    o.direction = static_cast<D>(row[2]);
    o.header_row = static_cast<H>(row[3]);
    o.header_column = static_cast<H>(row[4]);
    o.source = "Sheet1!$A$1:$D$5";
    const auto split = split_chart_source(wb, 0, o.source, o.kind, o.direction, o.header_row, o.header_column);
    ASSERT_TRUE(static_cast<bool>(split));
    const std::uint32_t id = Must(insert_chart(wb, 0, o));
    EXPECT_TRUE(NoRefresh(wb));

    const ChartModel m = Must(get_chart(wb, 0, id));
    ASSERT_EQ(m.plots.size(), 1U);
    ExpectKind(m.plots[0], o.kind, o.grouping);
    EXPECT_EQ(m.plots[0].series.size(), split.value().series.size());
    EXPECT_EQ(m.anchor_kind, AnchorKind::kTwoCell);
    EXPECT_EQ(m.cx, 4572000);
    EXPECT_EQ(m.cy, 2743200);
    EXPECT_EQ(m.title.kind, ChartTitleKind::kAuto);
    ASSERT_TRUE(m.source.has_value());
    EXPECT_EQ(m.source->range, "Sheet1!$A$1:$D$5");
    EXPECT_EQ(m.source->direction, split.value().direction);
    EXPECT_EQ(Must(get_chart_data(wb, 0, id)).series.size(), split.value().series.size());

    const K to = static_cast<K>(row[5]);
    MustOk(set_chart_type(wb, 0, id, 0, to, o.grouping));
    const ChartModel after = Must(get_chart(wb, 0, id));
    ASSERT_EQ(after.plots.size(), 1U);
    ExpectKind(after.plots[0], to, o.grouping);
    EXPECT_EQ(ValueRefs(after), ValueRefs(m));
    EXPECT_EQ(after.axes.size(), to == K::kPie ? 0U : 2U);
    for (const ChartAxis& a : after.axes) {
      EXPECT_EQ(a.kind, to == K::kScatter || a.id != after.plots[0].axis_ids[0] ? ChartAxisKind::kValue
                                                                                : ChartAxisKind::kCategory);
    }
    EXPECT_TRUE(NoRefresh(wb));
  }
}

TEST(ChartEdit, InsertTitleLegendAndRefusals) {
  Workbook wb = GridBook();
  ChartInsertOptions o;
  o.source = "Sheet1!$A$1:$C$5";
  o.title = std::string();
  o.legend = ChartLegendPosition::kNone;
  o.placement.width_emu = 1000000;
  o.placement.height_emu = 500000;
  o.name = "mine";
  const std::uint32_t id = Must(insert_chart(wb, 0, o));
  const ChartModel m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.title.kind, ChartTitleKind::kNone);
  EXPECT_EQ(m.auto_title_deleted, Tri::kTrue);
  EXPECT_FALSE(m.legend.present);
  EXPECT_EQ(m.cx, 1000000);
  EXPECT_EQ(m.name, "mine");
  o.title = std::string("Sales");
  const std::uint32_t second = Must(insert_chart(wb, 0, o));
  EXPECT_NE(second, id);
  EXPECT_EQ(Must(get_chart(wb, 0, second)).title.text, "Sales");

  o.source = "Sheet1!#REF!";
  ExpectCode(insert_chart(wb, 0, o), Code::kIoChartInvalidReference);
  ExpectCode(insert_chart(wb, 9, o), Code::kInvalidArgument);
  ExpectCode(get_chart(wb, 0, 999), Code::kInvalidArgument);

  Workbook xlsb = LoadFixture("chart_xlsb.xlsb");
  const Located x = Find(xlsb);
  EXPECT_TRUE(static_cast<bool>(get_chart(xlsb, x.sheet, x.id)));
  o.source = "Sheet1!$A$1:$C$5";
  ExpectCode(insert_chart(xlsb, x.sheet, o), Code::kIoDrawingUnparseable);
  ExpectCode(set_chart_title(xlsb, x.sheet, x.id, ChartTitle{}), Code::kIoDrawingUnparseable);
}

TEST(ChartEdit, ChartexAndPivotRefusals) {
  Workbook cx = LoadFixture("chart_chartex.xlsx");
  const Located c = Find(cx);
  EXPECT_TRUE(Must(get_chart(cx, c.sheet, c.id)).is_chartex);
  ExpectCode(set_chart_title(cx, c.sheet, c.id, ChartTitle{}), Code::kIoChartUnsupportedKind);
  ExpectCode(set_chart_display(cx, c.sheet, c.id, ChartBlanksAs::kGap, Tri::kTrue), Code::kIoChartUnsupportedKind);
  ExpectCode(preview_chart_data(cx, c.sheet, c.id, ChartPreviewOverrides{}), Code::kIoChartUnsupportedKind);

  Workbook pv = LoadFixture("chart_pivot.xlsx");
  const Located p = Find(pv);
  ASSERT_TRUE(Must(get_chart(pv, p.sheet, p.id)).is_pivot_chart);
  ChartSeriesPatch rename;
  rename.name = std::string("Sheet1!$A$1");
  ExpectCode(set_chart_series(pv, p.sheet, p.id, 0, 0, rename), Code::kIoChartUnsupportedKind);
  ExpectCode(add_chart_series(pv, p.sheet, p.id, 0, rename), Code::kIoChartUnsupportedKind);
  ExpectCode(remove_chart_series(pv, p.sheet, p.id, 0, 0), Code::kIoChartUnsupportedKind);
  ExpectCode(set_chart_filter(pv, p.sheet, p.id, {0}, {}), Code::kIoChartUnsupportedKind);
  ExpectCode(set_chart_source(pv, p.sheet, p.id, "Sheet1!$A$1:$B$3", D::kAuto, H::kAuto, H::kAuto),
             Code::kIoChartUnsupportedKind);
  ExpectCode(set_chart_type(pv, p.sheet, p.id, 0, K::kScatter, ChartGrouping::kClustered),
             Code::kIoChartUnsupportedKind);
  ChartTitle title;
  title.kind = ChartTitleKind::kText;
  title.text = "Pivot";
  MustOk(set_chart_title(pv, p.sheet, p.id, title));
  EXPECT_EQ(Must(get_chart(pv, p.sheet, p.id)).title.text, "Pivot");
  ChartSeriesPatch color;
  color.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF102030U);
  MustOk(set_chart_series(pv, p.sheet, p.id, 0, 0, color));
}

// ---------------------------------------------------------------------------
// No-op edits write nothing; a bad reference writes nothing
// ---------------------------------------------------------------------------

TEST(ChartEdit, NoOpAndFailedEditsWriteNothing) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const Located c = Find(wb);
  const ChartModel m = Must(get_chart(wb, c.sheet, c.id));
  const auto before = Parts(wb);
  MustOk(set_chart_display(wb, c.sheet, c.id, m.disp_blanks_as, m.plot_vis_only));
  ChartPlotPatch plot;
  plot.gap_width = m.plots[0].gap_width;
  plot.vary_colors = m.plots[0].vary_colors == Tri::kTrue;
  MustOk(set_chart_plot(wb, c.sheet, c.id, 0, plot));
  MustOk(set_chart_type(wb, c.sheet, c.id, 0, K::kColumn, *m.plots[0].grouping));
  ChartLegendPatch legend;
  legend.position = m.legend.position;
  legend.overlay = m.legend.overlay;
  MustOk(set_chart_legend(wb, c.sheet, c.id, legend));
  MustOk(set_chart_point(wb, c.sheet, c.id, 0, m.plots[0].series[0].idx, 3, ChartPointPatch{}));
  MustOk(set_chart_series(wb, c.sheet, c.id, 0, m.plots[0].series[0].idx, ChartSeriesPatch{}));
  MustOk(set_chart_filter(wb, c.sheet, c.id, {}, {}));
  MustOk(set_chart_axis(wb, c.sheet, c.id, m.axes[0].id, ChartAxisPatch{}));
  ChartTitle same = m.title;
  if (same.kind == ChartTitleKind::kAuto) {
    MustOk(set_chart_title(wb, c.sheet, c.id, same));
  }
  ChartSeriesPatch bad;
  bad.val = std::string("Nope!$A$1:$A$4");
  bad.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF000000U);
  ExpectCode(set_chart_series(wb, c.sheet, c.id, 0, m.plots[0].series[0].idx, bad), Code::kIoChartInvalidReference);
  bad.val = std::string();
  ExpectCode(set_chart_series(wb, c.sheet, c.id, 0, m.plots[0].series[0].idx, bad), Code::kInvalidArgument);
  ChartSeriesPatch marker;
  marker.marker = ChartPatchField<ChartMarkerPatch>::Set(ChartMarkerPatch{});
  ExpectCode(set_chart_series(wb, c.sheet, c.id, 0, m.plots[0].series[0].idx, marker), Code::kInvalidArgument);
  ChartAxisPatch unit;
  unit.major_unit = ChartPatchField<double>::Set(2.0);
  ExpectCode(
      set_chart_axis(wb, c.sheet, c.id, m.axes[0].kind == ChartAxisKind::kCategory ? m.axes[0].id : m.axes[1].id, unit),
      Code::kInvalidArgument);
  ExpectCode(set_chart_axis(wb, c.sheet, c.id, 12345U, unit), Code::kInvalidArgument);
  ExpectCode(set_chart_filter(wb, c.sheet, c.id, {99}, {}), Code::kInvalidArgument);
  ExpectCode(set_chart_filter(wb, c.sheet, c.id, {}, {0, 1, 2, 3}), Code::kInvalidArgument);
  ChartPlotPatch angle;
  angle.first_slice_angle = 10;
  ExpectCode(set_chart_plot(wb, c.sheet, c.id, 0, angle), Code::kInvalidArgument);
  EXPECT_EQ(Parts(wb), before);
}

// ---------------------------------------------------------------------------
// Set -> get round trips of every editable field
// ---------------------------------------------------------------------------

TEST(ChartEdit, TitleLegendDisplayPlotRoundTrip) {
  Workbook wb = GridBook();
  const std::uint32_t id = Insert(wb, K::kColumn);
  ChartTitle t;
  t.kind = ChartTitleKind::kText;
  t.text = "Two\nlines";
  MustOk(set_chart_title(wb, 0, id, t));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).title.text, "Two\nlines");
  t.kind = ChartTitleKind::kRef;
  t.ref_formula = "Sheet1!$B$1";
  MustOk(set_chart_title(wb, 0, id, t));
  ChartModel m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.title.kind, ChartTitleKind::kRef);
  EXPECT_EQ(m.title.ref_formula, "Sheet1!$B$1");
  EXPECT_EQ(Must(get_chart_data(wb, 0, id)).title_text, std::optional<std::string>("S1"));
  t.kind = ChartTitleKind::kNone;
  MustOk(set_chart_title(wb, 0, id, t));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.title.kind, ChartTitleKind::kNone);
  EXPECT_EQ(m.auto_title_deleted, Tri::kTrue);
  t.kind = ChartTitleKind::kAuto;
  MustOk(set_chart_title(wb, 0, id, t));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.title.kind, ChartTitleKind::kAuto);
  EXPECT_EQ(m.auto_title_deleted, Tri::kFalse);

  // Setting a text title reproduces the insert template's text title exactly.
  Workbook other = GridBook();
  ChartInsertOptions o;
  o.source = "Sheet1!$A$1:$D$5";
  o.title = std::string("Sales");
  const std::uint32_t titled = Must(insert_chart(other, 0, o));
  const std::uint32_t plain = Insert(other, K::kColumn);
  t.kind = ChartTitleKind::kText;
  t.text = "Sales";
  MustOk(set_chart_title(other, 0, plain, t));
  const Located a = Find(other, "xl/charts/chart1.xml");
  const Located b = Find(other, "xl/charts/chart2.xml");
  EXPECT_EQ(a.id, titled);
  ExpectSameForm(Between(PartText(other, a.path), "<c:title>", "</c:title>"),
                 Between(PartText(other, b.path), "<c:title>", "</c:title>"));

  for (const ChartLegendPosition pos : {ChartLegendPosition::kTop, ChartLegendPosition::kLeft,
                                        ChartLegendPosition::kTopRight, ChartLegendPosition::kBottom}) {
    ChartLegendPatch lp;
    lp.position = pos;
    lp.overlay = pos == ChartLegendPosition::kLeft;
    lp.hidden_series_idx = std::vector<std::uint32_t>{1};
    MustOk(set_chart_legend(wb, 0, id, lp));
    m = Must(get_chart(wb, 0, id));
    EXPECT_EQ(m.legend.position, pos);
    EXPECT_EQ(m.legend.overlay, lp.overlay);
    EXPECT_EQ(m.legend.hidden_series_idx, std::vector<std::uint32_t>{1});
  }
  ChartLegendPatch none;
  none.position = ChartLegendPosition::kNone;
  MustOk(set_chart_legend(wb, 0, id, none));
  EXPECT_FALSE(Must(get_chart(wb, 0, id)).legend.present);
  ChartLegendPatch back;
  back.position = ChartLegendPosition::kRight;
  MustOk(set_chart_legend(wb, 0, id, back));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).legend.position, ChartLegendPosition::kRight);

  for (const ChartBlanksAs b2 : {ChartBlanksAs::kZero, ChartBlanksAs::kSpan, ChartBlanksAs::kAbsent}) {
    for (const Tri vis : {Tri::kFalse, Tri::kAbsent, Tri::kTrue}) {
      MustOk(set_chart_display(wb, 0, id, b2, vis));
      m = Must(get_chart(wb, 0, id));
      EXPECT_EQ(m.disp_blanks_as, b2);
      EXPECT_EQ(m.plot_vis_only, vis);
    }
  }

  ChartPlotPatch pp;
  pp.gap_width = 80;
  pp.overlap = -20;
  MustOk(set_chart_plot(wb, 0, id, 0, pp));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.plots[0].gap_width, 80U);
  EXPECT_EQ(m.plots[0].overlap, -20);
  pp = ChartPlotPatch{};
  pp.vary_colors = true;
  MustOk(set_chart_plot(wb, 0, id, 0, pp));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.plots[0].vary_colors, Tri::kTrue);
  EXPECT_EQ(m.plots[0].series[0].points.size(), 4U);
  EXPECT_FALSE(m.plots[0].series[0].fill.color.has_value());
  pp.vary_colors = false;
  MustOk(set_chart_plot(wb, 0, id, 0, pp));
  m = Must(get_chart(wb, 0, id));
  EXPECT_TRUE(m.plots[0].series[0].points.empty());
  EXPECT_TRUE(m.plots[0].series[0].fill.color.has_value());

  const std::uint32_t pie = Insert(wb, K::kPie);
  pp = ChartPlotPatch{};
  pp.first_slice_angle = 45;
  MustOk(set_chart_plot(wb, 0, pie, 0, pp));
  EXPECT_EQ(Must(get_chart(wb, 0, pie)).plots[0].first_slice_angle, 45U);
  pp = ChartPlotPatch{};
  pp.gap_width = 10;
  ExpectCode(set_chart_plot(wb, 0, pie, 0, pp), Code::kInvalidArgument);
}

TEST(ChartEdit, AxisRoundTrip) {
  Workbook wb = GridBook();
  const std::uint32_t id = Insert(wb, K::kColumn);
  ChartModel m = Must(get_chart(wb, 0, id));
  const std::uint32_t cat = m.axes[0].id;
  const std::uint32_t val = m.axes[1].id;
  ChartAxisPatch p;
  p.visible = false;
  p.title = ChartPatchField<std::string>::Set("Amount");
  p.min = ChartPatchField<double>::Set(-2.5);
  p.max = ChartPatchField<double>::Set(40.0);
  p.major_unit = ChartPatchField<double>::Set(0.5);
  p.number_format = ChartPatchField<std::string>::Set("0.0");
  p.major_gridlines = false;
  p.reversed = true;
  p.log_base = ChartPatchField<double>::Set(10.0);
  p.crosses = ChartAxisCrosses::kAt;
  p.crosses_at = 3.0;
  p.tick_label_pos = ChartPatchField<std::string>::Set("low");
  MustOk(set_chart_axis(wb, 0, id, val, p));
  m = Must(get_chart(wb, 0, id));
  const ChartAxis& a = m.axes[1];
  EXPECT_TRUE(a.deleted);
  EXPECT_EQ(a.title.text, "Amount");
  EXPECT_EQ(a.min, -2.5);
  EXPECT_EQ(a.max, 40.0);
  EXPECT_EQ(a.major_unit, 0.5);
  ASSERT_TRUE(a.num_fmt.has_value());
  EXPECT_EQ(a.num_fmt->code, "0.0");
  EXPECT_FALSE(a.num_fmt->source_linked);
  EXPECT_FALSE(a.major_gridlines);
  EXPECT_TRUE(a.reversed);
  EXPECT_EQ(a.log_base, 10.0);
  EXPECT_EQ(a.crosses, ChartAxisCrosses::kAt);
  EXPECT_EQ(a.crosses_at, 3.0);
  EXPECT_EQ(a.tick_label_pos, "low");
  EXPECT_EQ(Must(get_chart_data(wb, 0, id)).axis_title_texts.size(), 1U);

  ChartAxisPatch undo;
  undo.visible = true;
  undo.title = ChartPatchField<std::string>::Clear();
  undo.min = ChartPatchField<double>::Clear();
  undo.max = ChartPatchField<double>::Clear();
  undo.major_unit = ChartPatchField<double>::Clear();
  undo.number_format = ChartPatchField<std::string>::Clear();
  undo.major_gridlines = true;
  undo.reversed = false;
  undo.log_base = ChartPatchField<double>::Clear();
  undo.crosses = ChartAxisCrosses::kMax;
  undo.tick_label_pos = ChartPatchField<std::string>::Clear();
  MustOk(set_chart_axis(wb, 0, id, val, undo));
  m = Must(get_chart(wb, 0, id));
  EXPECT_FALSE(m.axes[1].deleted);
  EXPECT_EQ(m.axes[1].title.kind, ChartTitleKind::kNone);
  EXPECT_TRUE(std::isnan(m.axes[1].min));
  EXPECT_TRUE(std::isnan(m.axes[1].major_unit));
  EXPECT_FALSE(m.axes[1].num_fmt.has_value());
  EXPECT_TRUE(m.axes[1].major_gridlines);
  EXPECT_FALSE(m.axes[1].reversed);
  EXPECT_EQ(m.axes[1].crosses, ChartAxisCrosses::kMax);
  EXPECT_EQ(m.axes[1].tick_label_pos, "nextTo");

  ChartAxisPatch cp;
  cp.title = ChartPatchField<std::string>::Set("Item");
  cp.major_gridlines = true;
  MustOk(set_chart_axis(wb, 0, id, cat, cp));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.axes[0].title.text, "Item");
  EXPECT_TRUE(m.axes[0].major_gridlines);
}

// Excel re-saves an axis on the far side when it crosses at the maximum, swapped while the crossing axis is
// reversed; undoing either moves it back.
TEST(ChartEdit, AxisSideFollowsReversalAndCrossing) {
  using P = ChartAxisPosition;
  const auto reverse = [](bool on) {
    ChartAxisPatch p;
    p.reversed = on;
    return p;
  };
  const auto crosses = [](ChartAxisCrosses cr) {
    ChartAxisPatch p;
    p.crosses = cr;
    return p;
  };
  const auto sides = [](const Workbook& wb, std::uint32_t id) {
    const ChartModel m = Must(get_chart(wb, 0, id));
    return std::make_pair(m.axes[0].position, m.axes[1].position);
  };
  Workbook wb = GridBook();
  const std::uint32_t col = Insert(wb, K::kColumn);
  ChartModel m = Must(get_chart(wb, 0, col));
  const std::uint32_t cat = m.axes[0].id;
  const std::uint32_t val = m.axes[1].id;
  MustOk(set_chart_axis(wb, 0, col, val, reverse(true)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kTop, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, cat, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, cat, crosses(ChartAxisCrosses::kMin)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kTop, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, val, reverse(false)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, cat, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kTop, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, cat, crosses(ChartAxisCrosses::kAutoZero)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kLeft));
  MustOk(set_chart_axis(wb, 0, col, cat, reverse(true)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kRight));
  MustOk(set_chart_axis(wb, 0, col, val, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kLeft));

  const std::uint32_t bar = Insert(wb, K::kBar, ChartGrouping::kStacked);
  m = Must(get_chart(wb, 0, bar));
  MustOk(set_chart_axis(wb, 0, bar, m.axes[1].id, reverse(true)));
  EXPECT_EQ(sides(wb, bar), std::make_pair(P::kRight, P::kBottom));
  MustOk(set_chart_axis(wb, 0, bar, m.axes[0].id, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, bar), std::make_pair(P::kLeft, P::kBottom));
  MustOk(set_chart_axis(wb, 0, bar, m.axes[1].id, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, bar), std::make_pair(P::kLeft, P::kTop));

  const std::uint32_t xy = Insert(wb, K::kScatter);
  m = Must(get_chart(wb, 0, xy));
  MustOk(set_chart_axis(wb, 0, xy, m.axes[0].id, reverse(true)));
  EXPECT_EQ(sides(wb, xy), std::make_pair(P::kBottom, P::kRight));
  MustOk(set_chart_axis(wb, 0, xy, m.axes[0].id, crosses(ChartAxisCrosses::kMax)));
  EXPECT_EQ(sides(wb, xy), std::make_pair(P::kTop, P::kRight));

  // A type change keeps the derived sides.
  MustOk(set_chart_axis(wb, 0, col, val, crosses(ChartAxisCrosses::kAutoZero)));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kBottom, P::kRight));
  MustOk(set_chart_type(wb, 0, col, 0, K::kBar, ChartGrouping::kClustered));
  EXPECT_EQ(sides(wb, col), std::make_pair(P::kLeft, P::kTop));
}

// Points take the forms Excel re-saves: a bar or area point repeats the series line, a line point its line
// and marker, and a line or scatter point drops a fill of its own.
TEST(ChartEdit, PointFormsFollowSeries) {
  Workbook wb = LoadFixture("chart_bar.xlsx");
  const Located bar = Find(wb, "xl/charts/chart2.xml");
  ChartPointPatch fill;
  fill.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF7030A0U);
  MustOk(set_chart_point(wb, bar.sheet, bar.id, 0, 0, 1, fill));
  EXPECT_EQ(Between(PartText(wb, bar.path), "<c:dPt>", "</c:dPt>"),
            "<c:dPt><c:idx val=\"1\"/><c:invertIfNegative val=\"0\"/><c:bubble3D val=\"0\"/><c:spPr><a:solidFill>"
            "<a:srgbClr val=\"7030A0\"/></a:solidFill><a:ln><a:noFill/></a:ln><a:effectLst/></c:spPr></c:dPt>");
  ChartPointPatch outline;
  outline.line.color = ChartPatchField<std::uint32_t>::Set(0xFFFF0000U);
  MustOk(set_chart_point(wb, bar.sheet, bar.id, 0, 0, 2, outline));
  const ChartModel bar_model = Must(get_chart(wb, bar.sheet, bar.id));
  const ChartSeries& b = bar_model.plots[0].series[0];
  ASSERT_EQ(b.points.size(), 2U);
  EXPECT_EQ(b.points[1].fill.color, b.fill.color);
  EXPECT_EQ(b.points[1].line.color, std::optional<std::uint32_t>(0xFFFF0000U));

  Workbook sw = LoadFixture("chart_scatter.xlsx");
  const Located xy = Find(sw, "xl/charts/chart2.xml");
  MustOk(set_chart_point(sw, xy.sheet, xy.id, 0, 0, 1, fill));
  EXPECT_EQ(Between(PartText(sw, xy.path), "<c:dPt>", "</c:dPt>"),
            "<c:dPt><c:idx val=\"1\"/><c:bubble3D val=\"0\"/><c:spPr><a:effectLst/></c:spPr></c:dPt>");
  MustOk(set_chart_point(sw, xy.sheet, xy.id, 0, 0, 2, outline));
  const ChartModel xy_model = Must(get_chart(sw, xy.sheet, xy.id));
  const ChartSeries& x = xy_model.plots[0].series[0];
  ASSERT_EQ(x.points.size(), 2U);
  EXPECT_EQ(x.points[1].line.color, std::optional<std::uint32_t>(0xFFFF0000U));
  EXPECT_EQ(x.points[1].line.width_emu, x.line.width_emu);
  ASSERT_TRUE(x.points[1].marker.has_value() && x.marker.has_value());
  EXPECT_EQ(x.points[1].marker->fill.color, x.marker->fill.color);

  Workbook lw = LoadFixture("chart_line.xlsx");
  const Located line = Find(lw, "xl/charts/chart4.xml");
  MustOk(set_chart_point(lw, line.sheet, line.id, 0, 0, 1, fill));
  const ChartModel m = Must(get_chart(lw, line.sheet, line.id));
  const ChartSeries& s = m.plots[0].series[0];
  ASSERT_EQ(s.points.size(), 1U);
  EXPECT_FALSE(s.points[0].fill.color.has_value());
  EXPECT_EQ(s.points[0].line.color, s.line.color);
  EXPECT_EQ(s.points[0].line.width_emu, s.line.width_emu);
  ASSERT_TRUE(s.points[0].marker.has_value() && s.marker.has_value());
  EXPECT_EQ(s.points[0].marker->symbol, s.marker->symbol);
  EXPECT_EQ(s.points[0].marker->fill.color, s.marker->fill.color);
}

// Excel drops a line or scatter series' own fill on re-save, transparent or not; line and marker
// transparency stay.
TEST(ChartEdit, LineAndScatterSeriesTakeNoFill) {
  ChartSeriesPatch p;
  p.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF1F77B4U);
  p.fill.transparency_pct = ChartPatchField<double>::Set(20.0);
  p.line.color = ChartPatchField<std::uint32_t>::Set(0xFF333333U);
  p.line.transparency_pct = ChartPatchField<double>::Set(20.0);
  ChartMarkerPatch mk;
  mk.fill.color = ChartPatchField<std::uint32_t>::Set(0xFFAA0000U);
  mk.fill.transparency_pct = ChartPatchField<double>::Set(20.0);
  p.marker = ChartPatchField<ChartMarkerPatch>::Set(mk);
  const std::pair<const char*, const char*> charts[] = {{"chart_line.xlsx", "xl/charts/chart4.xml"},
                                                        {"chart_scatter.xlsx", "xl/charts/chart2.xml"}};
  for (const auto& [file, part] : charts) {
    SCOPED_TRACE(file);
    Workbook wb = LoadFixture(file);
    const Located c = Find(wb, part);
    MustOk(set_chart_series(wb, c.sheet, c.id, 0, 0, p));
    const ChartSeries s = Must(get_chart(wb, c.sheet, c.id)).plots[0].series[0];
    EXPECT_FALSE(s.fill.color.has_value());
    EXPECT_FALSE(s.fill.alpha.has_value());
    EXPECT_EQ(s.line.color, std::optional<std::uint32_t>(0xFF333333U));
    EXPECT_EQ(s.line.alpha, std::optional<std::uint32_t>(80000U));
    ASSERT_TRUE(s.marker.has_value());
    EXPECT_EQ(s.marker->fill.color, std::optional<std::uint32_t>(0xFFAA0000U));
    EXPECT_EQ(s.marker->fill.alpha, std::optional<std::uint32_t>(80000U));
  }

  Workbook wb = LoadFixture("chart_column.xlsx");
  const Located col = Find(wb, "xl/charts/chart1.xml");
  ChartSeriesPatch fill;
  fill.fill = p.fill;
  MustOk(set_chart_series(wb, col.sheet, col.id, 0, 0, fill));
  const ChartSeries s = Must(get_chart(wb, col.sheet, col.id)).plots[0].series[0];
  EXPECT_EQ(s.fill.color, std::optional<std::uint32_t>(0xFF1F77B4U));
  EXPECT_EQ(s.fill.alpha, std::optional<std::uint32_t>(80000U));
}

// A patched series marker with a named symbol takes the form Excel re-saves unchanged: size 5 when it has none;
// with a fill and an outline, the outline gets 9525 EMU and the series accent.
TEST(ChartEdit, SeriesMarkerSettlesAsExcel) {
  const auto marker = [](const Workbook& wb, const Located& c) {
    const std::string text = PartText(wb, c.path);
    return Between(text, "<c:marker>", "</c:marker>");
  };
  const auto patch = [](ChartMarkerSymbol symbol, double size, bool fill, bool line_color, bool dash) {
    ChartMarkerPatch mk;
    mk.symbol = symbol;
    if (size > 0.0) {
      mk.size = ChartPatchField<double>::Set(size);
    }
    if (fill) {
      mk.fill.color = ChartPatchField<std::uint32_t>::Set(0xFFAA0000U);
    }
    if (line_color) {
      mk.line.color = ChartPatchField<std::uint32_t>::Set(0xFF000000U);
    }
    if (dash) {
      mk.line.dash = ChartPatchField<ChartLineDash>::Set(ChartLineDash::kDash);
    }
    ChartSeriesPatch p;
    p.marker = ChartPatchField<ChartMarkerPatch>::Set(mk);
    return p;
  };
  using S = ChartMarkerSymbol;
  struct Row {
    S symbol;
    double size;
    bool fill;
    bool line_color;
    bool dash;
    const char* want;
  };
  const Row rows[] = {
      {S::kDiamond, 0, true, true, false,
       "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"5\"/><c:spPr><a:solidFill><a:srgbClr val=\"AA0000\"/>"
       "</a:solidFill><a:ln w=\"9525\"><a:solidFill><a:srgbClr "
       "val=\"000000\"/></a:solidFill></a:ln><a:effectLst/></c:spPr>"
       "</c:marker>"},
      {S::kDiamond, 7, true, true, false,
       "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"7\"/><c:spPr><a:solidFill><a:srgbClr val=\"AA0000\"/>"
       "</a:solidFill><a:ln w=\"9525\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln><a:effectLst/>"
       "</c:spPr></c:marker>"},
      {S::kDiamond, 7, true, false, true,
       "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"7\"/><c:spPr><a:solidFill><a:srgbClr val=\"AA0000\"/>"
       "</a:solidFill><a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill>"
       "<a:prstDash val=\"dash\"/></a:ln><a:effectLst/></c:spPr></c:marker>"},
      {S::kDiamond, 7, false, true, false,
       "<c:marker><c:symbol val=\"diamond\"/><c:size val=\"7\"/><c:spPr><a:ln><a:solidFill><a:srgbClr "
       "val=\"000000\"/></a:solidFill></a:ln><a:effectLst/></c:spPr></c:marker>"},
      {S::kAuto, 7, true, true, false,
       "<c:marker><c:symbol val=\"auto\"/><c:size val=\"7\"/><c:spPr><a:solidFill><a:srgbClr val=\"AA0000\"/>"
       "</a:solidFill><a:ln><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln><a:effectLst/></c:spPr>"
       "</c:marker>"},
  };
  for (const Row& r : rows) {
    SCOPED_TRACE(r.want);
    Workbook wb = LoadFixture("chart_line.xlsx");
    const Located c = Find(wb, "xl/charts/chart1.xml");
    MustOk(set_chart_series(wb, c.sheet, c.id, 0, 0, patch(r.symbol, r.size, r.fill, r.line_color, r.dash)));
    EXPECT_EQ(marker(wb, c), r.want);
  }
  // A point marker is left as written.
  Workbook wb = LoadFixture("chart_line.xlsx");
  const Located c = Find(wb, "xl/charts/chart1.xml");
  ChartPointPatch pt;
  ChartMarkerPatch mk;
  mk.symbol = S::kSquare;
  mk.fill.color = ChartPatchField<std::uint32_t>::Set(0xFFAA0000U);
  mk.line.color = ChartPatchField<std::uint32_t>::Set(0xFF000000U);
  pt.marker = ChartPatchField<ChartMarkerPatch>::Set(mk);
  MustOk(set_chart_point(wb, c.sheet, c.id, 0, 0, 1, pt));
  const ChartSeries s = Must(get_chart(wb, c.sheet, c.id)).plots[0].series[0];
  ASSERT_TRUE(s.points[0].marker.has_value());
  EXPECT_FALSE(s.points[0].marker->size.has_value());
  EXPECT_FALSE(s.points[0].marker->line.width_emu.has_value());
}

// A percent-stacked change re-derives a linked value-axis format as Excel does, and a change to pie gives an
// undeleted absent title the automatic one.
TEST(ChartEdit, TypeChangeDerivesAxisFormatAndPieTitle) {
  Workbook wb = LoadFixture("chart_line.xlsx");
  const Located line = Find(wb, "xl/charts/chart4.xml");
  const auto val_fmt = [&] {
    const ChartModel m = Must(get_chart(wb, line.sheet, line.id));
    return m.axes[1].num_fmt.has_value() ? m.axes[1].num_fmt->code : std::string();
  };
  MustOk(set_chart_type(wb, line.sheet, line.id, 0, K::kArea, ChartGrouping::kPercentStacked));
  EXPECT_EQ(val_fmt(), "0%");
  MustOk(set_chart_type(wb, line.sheet, line.id, 0, K::kColumn, ChartGrouping::kPercentStacked));
  EXPECT_EQ(val_fmt(), "0%");
  MustOk(set_chart_type(wb, line.sheet, line.id, 0, K::kColumn, ChartGrouping::kStacked));
  EXPECT_EQ(val_fmt(), "General");
  ChartAxisPatch fixed;
  fixed.number_format = ChartPatchField<std::string>::Set("0.0");
  MustOk(set_chart_axis(wb, line.sheet, line.id, Must(get_chart(wb, line.sheet, line.id)).axes[1].id, fixed));
  MustOk(set_chart_type(wb, line.sheet, line.id, 0, K::kLine, ChartGrouping::kPercentStacked));
  EXPECT_EQ(val_fmt(), "0.0");

  Workbook cw = LoadFixture("chart_column.xlsx");
  const Located col = Find(cw, "xl/charts/chart1.xml");
  ASSERT_EQ(Must(get_chart(cw, col.sheet, col.id)).title.kind, ChartTitleKind::kNone);
  MustOk(set_chart_type(cw, col.sheet, col.id, 0, K::kPie, ChartGrouping::kClustered));
  EXPECT_EQ(Must(get_chart(cw, col.sheet, col.id)).title.kind, ChartTitleKind::kAuto);
  const Located other = Find(cw, "xl/charts/chart2.xml");
  MustOk(set_chart_type(cw, other.sheet, other.id, 0, K::kLine, ChartGrouping::kStacked));
  EXPECT_EQ(Must(get_chart(cw, other.sheet, other.id)).title.kind, ChartTitleKind::kNone);
  ChartTitle none;
  none.kind = ChartTitleKind::kNone;
  MustOk(set_chart_title(cw, other.sheet, other.id, none));
  MustOk(set_chart_type(cw, other.sheet, other.id, 0, K::kPie, ChartGrouping::kClustered));
  EXPECT_EQ(Must(get_chart(cw, other.sheet, other.id)).title.kind, ChartTitleKind::kNone);
}

// Pairwise: plot kind x fill x transparency x line x width x dash x labels x marker.
TEST(ChartEdit, SeriesFormatPairwise) {
  struct Row {
    int kind;
    const char* fill;
    const char* alpha;
    const char* line;
    const char* width;
    const char* dash;
    const char* labels;
    const char* marker;
  };
  static const Row kRows[] = {
      {0, "clear", "clear", "keep", "set", "dot", "keep", "clear"},
      {5, "set", "clear", "color", "clear", "keep", "set", "keep"},
      {2, "set", "set", "color", "keep", "dashDot", "keep", "set"},
      {4, "keep", "clear", "keep", "clear", "dashDot", "clear", "set"},
      {2, "set", "keep", "none", "set", "keep", "clear", "clear"},
      {2, "keep", "keep", "keep", "keep", "dash", "keep", "keep"},
      {4, "clear", "keep", "color", "set", "solid", "set", "set"},
      {0, "keep", "keep", "color", "keep", "dot", "set", "clear"},
      {4, "set", "set", "none", "clear", "keep", "keep", "keep"},
      {5, "keep", "keep", "keep", "clear", "solid", "keep", "clear"},
      {0, "set", "set", "keep", "clear", "dash", "clear", "keep"},
      {5, "set", "set", "color", "set", "dash", "set", "clear"},
      {2, "set", "clear", "color", "keep", "solid", "clear", "keep"},
      {0, "keep", "keep", "keep", "keep", "keep", "set", "set"},
      {2, "clear", "clear", "none", "clear", "keep", "set", "set"},
      {5, "set", "set", "keep", "keep", "dot", "clear", "set"},
      {0, "clear", "keep", "none", "keep", "keep", "clear", "set"},
      {4, "clear", "keep", "keep", "keep", "dashDot", "keep", "clear"},
      {4, "set", "clear", "keep", "set", "dash", "set", "set"},
      {0, "keep", "clear", "keep", "keep", "solid", "keep", "clear"},
      {0, "set", "set", "color", "keep", "dashDot", "set", "clear"},
      {5, "keep", "clear", "color", "set", "dashDot", "set", "keep"},
      {5, "keep", "clear", "none", "keep", "keep", "set", "clear"},
      {5, "clear", "clear", "color", "keep", "dash", "clear", "set"},
      {2, "set", "clear", "color", "set", "dot", "keep", "clear"},
      {0, "clear", "clear", "keep", "set", "keep", "keep", "keep"},
      {4, "clear", "clear", "color", "set", "dot", "set", "keep"},
      {2, "clear", "clear", "color", "clear", "dot", "clear", "set"},
      {5, "set", "set", "color", "set", "solid", "set", "clear"},
  };
  const auto is = [](const char* a, const char* b) { return std::string(a) == b; };
  for (const Row& r : kRows) {
    SCOPED_TRACE(testing::Message() << r.kind << r.fill << r.alpha << r.line << r.width << r.dash << r.labels
                                    << r.marker);
    Workbook wb = GridBook();
    const K kind = static_cast<K>(r.kind);
    const std::uint32_t id = Insert(wb, kind);
    const ChartSeries before = Must(get_chart(wb, 0, id)).plots[0].series[0];
    ChartSeriesPatch p;
    if (is(r.fill, "set"))
      p.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF112233U);
    if (is(r.fill, "clear"))
      p.fill.color = ChartPatchField<std::uint32_t>::Clear();
    if (is(r.alpha, "set"))
      p.fill.transparency_pct = ChartPatchField<double>::Set(30.0);
    if (is(r.alpha, "clear"))
      p.fill.transparency_pct = ChartPatchField<double>::Clear();
    if (is(r.line, "color"))
      p.line.color = ChartPatchField<std::uint32_t>::Set(0xFF445566U);
    if (is(r.line, "none"))
      p.line.none = true;
    if (is(r.width, "set"))
      p.line.width_pt = ChartPatchField<double>::Set(2.0);
    if (is(r.width, "clear"))
      p.line.width_pt = ChartPatchField<double>::Clear();
    static const char* const kDash[] = {"solid", "dash", "dot", "dashDot"};
    for (int d = 0; d < 4; ++d) {
      if (is(r.dash, kDash[d]))
        p.line.dash = ChartPatchField<ChartLineDash>::Set(static_cast<ChartLineDash>(d));
    }
    if (is(r.labels, "set")) {
      ChartLabelsPatch l;
      l.show_val = true;
      l.show_cat = true;
      l.separator = ChartPatchField<std::string>::Set("; ");
      l.number_format = ChartPatchField<std::string>::Set("0.0");
      l.font_color = ChartPatchField<std::uint32_t>::Set(0xFF0000FFU);
      l.font_size_pt = ChartPatchField<double>::Set(12.0);
      p.labels = ChartPatchField<ChartLabelsPatch>::Set(l);
    }
    if (is(r.labels, "clear"))
      p.labels = ChartPatchField<ChartLabelsPatch>::Clear();
    if (is(r.marker, "set")) {
      ChartMarkerPatch mk;
      mk.symbol = ChartMarkerSymbol::kDiamond;
      mk.size = ChartPatchField<double>::Set(7.0);
      mk.fill.color = ChartPatchField<std::uint32_t>::Set(0xFFAABBCCU);
      p.marker = ChartPatchField<ChartMarkerPatch>::Set(mk);
      p.smooth = true;
    }
    if (is(r.marker, "clear"))
      p.marker = ChartPatchField<ChartMarkerPatch>::Clear();
    const bool markers = kind == K::kLine || kind == K::kScatter;
    const auto result = set_chart_series(wb, 0, id, 0, 0, p);
    if (!is(r.marker, "keep") && !markers) {
      ExpectCode(result, Code::kInvalidArgument);
      continue;
    }
    MustOk(result);
    const ChartSeries s = Must(get_chart(wb, 0, id)).plots[0].series[0];
    // A line or scatter series takes no fill of its own.
    EXPECT_EQ(s.fill.color, markers               ? before.fill.color
                            : is(r.fill, "set")   ? std::optional<std::uint32_t>(0xFF112233U)
                            : is(r.fill, "clear") ? std::nullopt
                                                  : before.fill.color);
    if (markers) {
      EXPECT_EQ(s.fill.alpha, before.fill.alpha);
    } else if (is(r.fill, "set") || is(r.alpha, "clear")) {
      EXPECT_EQ(s.fill.alpha, is(r.alpha, "set") ? std::optional<std::uint32_t>(70000U) : std::nullopt);
    }
    if (is(r.line, "color")) {
      EXPECT_EQ(s.line.color, std::optional<std::uint32_t>(0xFF445566U));
      EXPECT_FALSE(s.line.none);
    } else if (is(r.line, "none")) {
      EXPECT_TRUE(s.line.none);
      EXPECT_FALSE(s.line.color.has_value());
    } else {
      EXPECT_EQ(s.line.color, before.line.color);
    }
    EXPECT_EQ(s.line.width_emu, is(r.width, "set")     ? std::optional<std::int64_t>(25400)
                                : is(r.width, "clear") ? std::nullopt
                                                       : before.line.width_emu);
    if (!is(r.dash, "keep")) {
      EXPECT_EQ(s.line.dash_name, r.dash);
    }
    if (is(r.labels, "set")) {
      ASSERT_TRUE(s.labels.has_value());
      EXPECT_TRUE(s.labels->show_val);
      EXPECT_TRUE(s.labels->show_cat);
      EXPECT_FALSE(s.labels->show_pct);
      EXPECT_EQ(s.labels->separator, std::optional<std::string>("; "));
      ASSERT_TRUE(s.labels->num_fmt.has_value());
      EXPECT_EQ(s.labels->num_fmt->code, "0.0");
      EXPECT_EQ(s.labels->font_color, std::optional<std::uint32_t>(0xFF0000FFU));
      EXPECT_EQ(s.labels->font_size_hundredths_pt, std::optional<std::uint32_t>(1200U));
    } else if (is(r.labels, "clear")) {
      EXPECT_FALSE(s.labels.has_value());
    }
    if (is(r.marker, "set")) {
      ASSERT_TRUE(s.marker.has_value());
      EXPECT_EQ(s.marker->symbol, ChartMarkerSymbol::kDiamond);
      EXPECT_EQ(s.marker->size, std::optional<std::uint32_t>(7U));
      EXPECT_EQ(s.marker->fill.color, std::optional<std::uint32_t>(0xFFAABBCCU));
      EXPECT_EQ(s.smooth, Tri::kTrue);
    } else if (is(r.marker, "clear")) {
      EXPECT_FALSE(s.marker.has_value());
    }
  }
}

TEST(ChartEdit, PointsSeriesAndSource) {
  Workbook wb = GridBook();
  const std::uint32_t id = Insert(wb, K::kLine);
  ChartPointPatch pt;
  pt.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF0000FFU);
  pt.line.width_pt = ChartPatchField<double>::Set(1.0);
  ChartMarkerPatch mk;
  mk.symbol = ChartMarkerSymbol::kTriangle;
  pt.marker = ChartPatchField<ChartMarkerPatch>::Set(mk);
  MustOk(set_chart_point(wb, 0, id, 0, 1, 2, pt));
  MustOk(set_chart_point(wb, 0, id, 0, 1, 0, pt));
  ChartModel m = Must(get_chart(wb, 0, id));
  ASSERT_EQ(m.plots[0].series[1].points.size(), 2U);
  EXPECT_EQ(m.plots[0].series[1].points[0].idx, 0U);
  EXPECT_EQ(m.plots[0].series[1].points[1].idx, 2U);
  EXPECT_FALSE(m.plots[0].series[1].points[1].fill.color.has_value());
  EXPECT_EQ(m.plots[0].series[1].points[1].line.width_emu, std::optional<std::int64_t>(12700));
  EXPECT_EQ(m.plots[0].series[1].points[1].marker->symbol, ChartMarkerSymbol::kTriangle);
  MustOk(set_chart_point(wb, 0, id, 0, 1, 2, ChartPointPatch{}));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).plots[0].series[1].points.size(), 1U);

  ChartSeriesPatch add;
  add.name = std::string("Sheet1!$D$1");
  add.cat = std::string("Sheet1!$A$2:$A$5");
  add.val = std::string("Sheet1!$D$2:$D$5");
  add.line.color = ChartPatchField<std::uint32_t>::Set(0xFF123456U);
  MustOk(remove_chart_series(wb, 0, id, 0, 2));
  EXPECT_EQ(Must(add_chart_series(wb, 0, id, 0, add)), 2U);
  EXPECT_TRUE(NoRefresh(wb));
  m = Must(get_chart(wb, 0, id));
  ASSERT_EQ(m.plots[0].series.size(), 3U);
  EXPECT_EQ(m.plots[0].series[2].val.formula, "Sheet1!$D$2:$D$5");
  EXPECT_EQ(m.plots[0].series[2].line.color, std::optional<std::uint32_t>(0xFF123456U));
  EXPECT_EQ(Must(get_chart_data(wb, 0, id)).series[2].name, std::optional<std::string>("S3"));

  ChartSeriesPatch refs;
  refs.name = std::string();
  refs.val = std::string("Sheet1!$C$2:$C$5");
  MustOk(set_chart_series(wb, 0, id, 0, 2, refs));
  EXPECT_TRUE(NoRefresh(wb));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.plots[0].series[2].name.kind, ChartDataSource::Kind::kNone);
  EXPECT_EQ(m.plots[0].series[2].val.formula, "Sheet1!$C$2:$C$5");
  ExpectCode(remove_chart_series(wb, 0, id, 0, 7), Code::kInvalidArgument);

  // Source rebuild keeps per-series formatting by idx and drops series past the new count.
  ChartSeriesPatch red;
  red.line.color = ChartPatchField<std::uint32_t>::Set(0xFFFF0000U);
  MustOk(set_chart_series(wb, 0, id, 0, 0, red));
  MustOk(set_chart_source(wb, 0, id, "Sheet1!$A$1:$C$5", D::kRows, H::kAuto, H::kAuto));
  EXPECT_TRUE(NoRefresh(wb));
  m = Must(get_chart(wb, 0, id));
  ASSERT_EQ(m.plots[0].series.size(), 4U);
  EXPECT_EQ(m.plots[0].series[0].line.color, std::optional<std::uint32_t>(0xFFFF0000U));
  ASSERT_TRUE(m.source.has_value());
  EXPECT_EQ(m.source->range, "Sheet1!$A$1:$C$5");
  EXPECT_EQ(m.source->direction, D::kRows);
  MustOk(set_chart_source(wb, 0, id, "Sheet1!$A$1:$B$5", D::kColumns, H::kAuto, H::kAuto));
  m = Must(get_chart(wb, 0, id));
  ASSERT_EQ(m.plots[0].series.size(), 1U);
  EXPECT_EQ(m.plots[0].series[0].line.color, std::optional<std::uint32_t>(0xFFFF0000U));
  ExpectCode(set_chart_source(wb, 0, id, "Nope!$A$1:$B$5", D::kColumns, H::kAuto, H::kAuto),
             Code::kIoChartInvalidReference);
}

TEST(ChartEdit, FilterRoundTrip) {
  Workbook wb = GridBook();
  const std::uint32_t id = Insert(wb, K::kColumn);
  const ChartData full = Must(get_chart_data(wb, 0, id));
  MustOk(set_chart_filter(wb, 0, id, {1}, {0, 2}));
  EXPECT_TRUE(NoRefresh(wb));
  ChartModel m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.category_filter, (std::vector<std::uint32_t>{0, 2}));
  int filtered = 0;
  for (const ChartSeries& s : m.plots[0].series) {
    filtered += s.filtered ? 1 : 0;
    EXPECT_EQ(s.filtered, s.idx == 1);
  }
  EXPECT_EQ(filtered, 1);
  const ChartData data = Must(get_chart_data(wb, 0, id));
  // Filtered categories stay in the data; the model's indices say which to hide.
  EXPECT_EQ(data.categories.labels, (std::vector<std::string>{"a", "b", "c", "d"}));
  MustOk(set_chart_filter(wb, 0, id, {0, 2}, {3}));
  m = Must(get_chart(wb, 0, id));
  EXPECT_EQ(m.category_filter, std::vector<std::uint32_t>{3});
  for (const ChartSeries& s : m.plots[0].series) {
    EXPECT_EQ(s.filtered, s.idx != 1);
  }
  // A type change carries the filtered series along.
  MustOk(set_chart_type(wb, 0, id, 0, K::kLine, ChartGrouping::kStacked));
  m = Must(get_chart(wb, 0, id));
  for (const ChartSeries& s : m.plots[0].series) {
    EXPECT_EQ(s.filtered, s.idx != 1);
  }
  MustOk(set_chart_filter(wb, 0, id, {}, {}));
  m = Must(get_chart(wb, 0, id));
  EXPECT_TRUE(m.category_filter.empty());
  EXPECT_EQ(m.plots[0].series[0].val.formula, "Sheet1!$B$2:$B$5");
  ExpectSameData(Must(get_chart_data(wb, 0, id)), full);
  EXPECT_TRUE(NoRefresh(wb));
}

// ---------------------------------------------------------------------------
// Measured XML forms
// ---------------------------------------------------------------------------

// Series and category filters written the way Excel saves them.
TEST(ChartEdit, FilterFormsMatchExcel) {
  Workbook wb = LoadFixture("chart_filter.xlsx");
  const std::string series_form = PartText(wb, "xl/charts/chart1.xml");
  const std::string category_form = PartText(wb, "xl/charts/chart2.xml");
  const Located c = Find(wb, "xl/charts/chart3.xml");
  const std::string original = PartText(wb, c.path);

  MustOk(set_chart_filter(wb, c.sheet, c.id, {1}, {}));
  ExpectSameForm(NoUid(Between(PartText(wb, c.path), "<c:extLst><c:ext uri=\"{02D57815", "</c:extLst></c:barChart>")),
                 NoUid(Between(series_form, "<c:extLst><c:ext uri=\"{02D57815", "</c:extLst></c:barChart>")));
  MustOk(set_chart_filter(wb, c.sheet, c.id, {}, {1}));
  const std::string cat_text = PartText(wb, c.path);
  for (const char* tag : {"c:cat", "c:val"}) {
    for (int n = 0; n < 2; ++n) {
      SCOPED_TRACE(testing::Message() << tag << n);
      ExpectSameForm(Between(cat_text, std::string("<") + tag + ">", std::string("</") + tag + ">", n),
                     Between(category_form, std::string("<") + tag + ">", std::string("</") + tag + ">", n));
    }
  }
  MustOk(set_chart_filter(wb, c.sheet, c.id, {}, {}));
  ExpectSameForm(Between(PartText(wb, c.path), "<c:plotArea>", "</c:plotArea>"),
                 Between(original, "<c:plotArea>", "</c:plotArea>"));
}

// Series fill / line / labels, point and vary-colours forms.
TEST(ChartEdit, FormatFormsMatchExcel) {
  const Workbook fmt = LoadFixture("chart_format.xlsx");
  const std::string col_form = PartText(fmt, "xl/charts/chart2.xml");
  const std::string vary_form = PartText(fmt, "xl/charts/chart4.xml");
  Workbook wb = LoadFixture("chart_filter.xlsx");
  const Located c = Find(wb, "xl/charts/chart3.xml");

  ChartSeriesPatch p;
  p.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF008000U);
  p.fill.transparency_pct = ChartPatchField<double>::Set(50.0);
  p.line.color = ChartPatchField<std::uint32_t>::Set(0xFF000000U);
  p.line.width_pt = ChartPatchField<double>::Set(2.0);
  ChartLabelsPatch l;
  l.show_val = true;
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(l);
  MustOk(set_chart_series(wb, c.sheet, c.id, 0, 0, p));
  const std::string text = PartText(wb, c.path);
  EXPECT_EQ(Between(text, "<c:spPr><a:solidFill><a:srgbClr", "</c:spPr>"),
            "<c:spPr><a:solidFill><a:srgbClr val=\"008000\"><a:alpha val=\"50000\"/></a:srgbClr></a:solidFill>"
            "<a:ln w=\"25400\"><a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill></a:ln><a:effectLst/>"
            "</c:spPr>");
  ExpectSameForm(Between(text, "<c:dLbls>", "</c:dLbls>"), Between(col_form, "<c:dLbls>", "</c:dLbls>"));

  ChartPointPatch pt;
  pt.fill.color = ChartPatchField<std::uint32_t>::Set(0xFF0000FFU);
  pt.line.none = true;
  MustOk(set_chart_point(wb, c.sheet, c.id, 0, 1, 2, pt));
  std::string want = Between(col_form, "<c:dPt>", "</c:dPt>");
  want = std::regex_replace(want, std::regex("<c:extLst>.*</c:extLst>"), "");
  ExpectSameForm(Between(PartText(wb, c.path), "<c:dPt>", "</c:dPt>"), want);

  const Located single = Find(wb, "xl/charts/chart4.xml");
  ChartPlotPatch vary;
  vary.vary_colors = true;
  MustOk(set_chart_plot(wb, single.sheet, single.id, 0, vary));
  const std::string varied = PartText(wb, single.path);
  const auto points = [](const std::string& t) {
    return t.substr(t.find("<c:dPt>"), t.rfind("</c:dPt>") + 8 - t.find("<c:dPt>"));
  };
  ExpectSameForm(points(varied), points(vary_form));
  ExpectSameForm(Between(varied, "<c:ser>", "<c:invertIfNegative"),
                 Between(vary_form, "<c:ser>", "<c:invertIfNegative"));
}

// ---------------------------------------------------------------------------
// Preview (pairwise: direction x header row x header column x blanks x visibility x source)
// ---------------------------------------------------------------------------

TEST(ChartEdit, PreviewMatchesAppliedEditPairwise) {
  static const int kRows[][6] = {{1, 2, 2, 1, 2, 1}, {0, 2, 0, 0, 1, 0}, {1, 0, 1, 3, 0, 0}, {0, 1, 2, 2, 0, 1},
                                 {2, 1, 2, 1, 1, 0}, {0, 0, 0, 1, 2, 0}, {2, 2, 0, 3, 0, 1}, {1, 2, 1, 1, 0, 0},
                                 {1, 1, 1, 0, 2, 1}, {0, 2, 1, 2, 1, 1}, {1, 2, 0, 2, 2, 0}, {1, 0, 2, 0, 0, 1},
                                 {1, 1, 0, 3, 1, 0}, {2, 0, 1, 2, 1, 0}, {2, 0, 1, 2, 2, 0}, {2, 1, 0, 0, 0, 1},
                                 {0, 0, 2, 3, 2, 0}};
  for (const auto& row : kRows) {
    SCOPED_TRACE(testing::Message() << row[0] << row[1] << row[2] << row[3] << row[4] << row[5]);
    Workbook wb = GridBook();
    const std::uint32_t id = Insert(wb, K::kLine);
    ChartPreviewOverrides o;
    if (row[0] != 0)
      o.direction = static_cast<D>(row[0]);
    if (row[1] != 0)
      o.header_row = row[1] == 1;
    if (row[2] != 0)
      o.header_column = row[2] == 1;
    if (row[3] != 0)
      o.blanks_as = static_cast<ChartBlanksAs>(row[3]);
    if (row[4] != 0)
      o.plot_vis_only = row[4] == 1;
    if (row[5] != 0)
      o.source = "Sheet1!$A$1:$C$4";
    const ChartModel m = Must(get_chart(wb, 0, id));
    const auto before = Parts(wb);
    const ChartData preview = Must(preview_chart_data(wb, 0, id, o));
    EXPECT_EQ(Parts(wb), before);

    const bool resplit = o.source || o.direction || o.header_row || o.header_column;
    const auto mode = [&](const std::optional<bool>& v, bool derived) {
      return v ? (*v ? H::kYes : H::kNo) : o.source ? H::kAuto : (derived ? H::kYes : H::kNo);
    };
    if (resplit) {
      MustOk(set_chart_source(wb, 0, id, o.source ? *o.source : m.source->range,
                              o.direction ? *o.direction
                              : o.source  ? D::kAuto
                                          : m.source->direction,
                              mode(o.header_row, m.source->header_row),
                              mode(o.header_column, m.source->header_column)));
    }
    MustOk(set_chart_display(wb, 0, id, o.blanks_as ? *o.blanks_as : m.disp_blanks_as,
                             o.plot_vis_only ? (*o.plot_vis_only ? Tri::kTrue : Tri::kFalse) : m.plot_vis_only));
    ExpectSameData(preview, Must(get_chart_data(wb, 0, id)));
  }
}

TEST(ChartEdit, DoughnutHoleSizeAndRefusals) {
  Workbook wb = LoadFixture("chart_combo.xlsx");
  Located d;
  auto all = all_chart_parts(wb);
  ASSERT_TRUE(static_cast<bool>(all));
  for (const SheetChartPartRef& e : all.value()) {
    if (PartText(wb, e.ref.part_path).find("<c:doughnutChart>") != std::string::npos) {
      d = Located{e.sheet, e.ref.object_id, e.ref.part_path};
    }
  }
  ASSERT_FALSE(d.path.empty());
  ChartPlotPatch p;
  p.hole_size = 35;
  MustOk(set_chart_plot(wb, d.sheet, d.id, 0, p));
  EXPECT_EQ(Must(get_chart(wb, d.sheet, d.id)).plots[0].hole_size, 35U);
  const auto before = Parts(wb);
  MustOk(set_chart_plot(wb, d.sheet, d.id, 0, p));
  EXPECT_EQ(Parts(wb), before);
  for (const std::uint32_t bad : {9U, 91U}) {
    p.hole_size = bad;
    ExpectCode(set_chart_plot(wb, d.sheet, d.id, 0, p), Code::kInvalidArgument);
  }

  Workbook grid = GridBook();
  const std::uint32_t id = Insert(grid, K::kColumn);
  p.hole_size = 50;
  ExpectCode(set_chart_plot(grid, 0, id, 0, p), Code::kInvalidArgument);
}

/// The plot area with caches, unique ids and axis ids replaced by placeholders.
std::string PlotForm(const std::string& part) {
  std::string out = Between(part, "<c:plotArea>", "</c:plotArea>");
  out = std::regex_replace(out, std::regex("<c:(str|num|multiLvlStr)Cache>[\\s\\S]*?</c:\\1Cache>"), "");
  out = NoUid(out);
  std::vector<std::string> ids;
  std::string mapped;
  const std::regex axis("<c:(axId|crossAx) val=\"(\\d+)\"/>");
  auto last = out.cbegin();
  for (std::sregex_iterator it(out.cbegin(), out.cend(), axis), end; it != end; ++it) {
    const std::string id = (*it)[2];
    if (std::find(ids.begin(), ids.end(), id) == ids.end()) {
      ids.push_back(id);
    }
    mapped.append(last, (*it)[0].first);
    mapped += "<c:" + (*it)[1].str() + " val=\"" + std::to_string(std::find(ids.begin(), ids.end(), id) - ids.begin()) +
              "\"/>";
    last = (*it)[0].second;
  }
  mapped.append(last, out.cend());
  return mapped;
}

// Doughnut insert and the type changes to and from it, against the Excel-saved results.
TEST(ChartEdit, DoughnutFormsMatchExcel) {
  Workbook wb = LoadFixture("chart_doughnut.xlsx");
  EXPECT_TRUE(NoRefresh(wb));  // Excel's caches already match the cells.
  const auto excel = [&](int n) { return PlotForm(PartText(wb, "xl/charts/chart" + std::to_string(n) + ".xml")); };
  const auto made = [&](std::uint32_t id) {
    for (const SheetChartPartRef& e : Must(all_chart_parts(wb))) {
      if (e.ref.object_id == id) {
        return PlotForm(PartText(wb, e.ref.part_path));
      }
    }
    return std::string();
  };
  const auto insert = [&](K kind, const char* source) {
    ChartInsertOptions o;
    o.kind = kind;
    o.source = source;
    return Must(insert_chart(wb, 0, o));
  };
  ExpectSameForm(made(insert(K::kDoughnut, "Sheet1!$A$1:$B$5")), excel(1));
  const std::uint32_t two = insert(K::kDoughnut, "Sheet1!$A$1:$C$5");
  ExpectSameForm(made(two), excel(2));
  EXPECT_TRUE(NoRefresh(wb));

  const std::uint32_t pie = insert(K::kPie, "Sheet1!$A$1:$C$5");
  MustOk(set_chart_type(wb, 0, pie, 0, K::kDoughnut, ChartGrouping::kClustered));
  ExpectSameForm(made(pie), excel(4));
  MustOk(set_chart_type(wb, 0, two, 0, K::kColumn, ChartGrouping::kClustered));
  ExpectSameForm(made(two), excel(5));
  const std::uint32_t column = insert(K::kColumn, "Sheet1!$A$1:$C$5");
  MustOk(set_chart_type(wb, 0, column, 0, K::kDoughnut, ChartGrouping::kClustered));
  ExpectSameForm(made(column), excel(6));
  EXPECT_TRUE(NoRefresh(wb));

  const ChartModel m = Must(get_chart(wb, 0, column));
  ExpectKind(m.plots[0], K::kDoughnut, ChartGrouping::kClustered);
  EXPECT_TRUE(m.axes.empty());
  EXPECT_EQ(m.plots[0].hole_size, 50U);

  ChartSeriesPatch add;
  add.val = std::string("Sheet1!$B$2:$B$5");
  EXPECT_EQ(Must(add_chart_series(wb, 0, column, 0, add)), 2U);
  EXPECT_EQ(Must(get_chart(wb, 0, column)).plots[0].series[2].points.size(), 4U);
  MustOk(set_chart_source(wb, 0, column, "Sheet1!$A$1:$B$5", D::kAuto, H::kAuto, H::kAuto));
  EXPECT_EQ(Must(get_chart(wb, 0, column)).plots[0].series.size(), 1U);
  EXPECT_TRUE(NoRefresh(wb));
}

TEST(ChartEdit, RectRefQuotesSheetNames) {
  EXPECT_EQ(chart_rect_ref("Sheet1", 1, 0, 4, 0), "Sheet1!$A$2:$A$5");
  EXPECT_EQ(chart_rect_ref("My Data", 0, 1, 0, 1), "'My Data'!$B$1");
  EXPECT_EQ(chart_rect_ref("It's", 0, 27, 0, 27), "'It''s'!$AB$1");
}

}  // namespace
}  // namespace formulon
