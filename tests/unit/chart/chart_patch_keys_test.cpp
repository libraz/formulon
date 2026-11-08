#include <cmath>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "chart/chart_edit.h"
#include "chart/chart_model.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "passthrough_part.h"
#include "sheet.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

using Code = FormulonErrorCode;

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

/// Sheet1: B1:C1 names, A2:A4 text categories, B2:C4 numbers.
Workbook GridBook() {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(mac_365_ja_jp_profile());
  Sheet& s = wb.sheet(0);
  s.set_cell_text(0, 1, "S1");
  s.set_cell_text(0, 2, "S2");
  const char* cats[] = {"a", "b", "c"};
  for (std::uint32_t r = 0; r < 3; ++r) {
    s.set_cell_text(r + 1, 0, cats[r]);
    s.set_cell_value(r + 1, 1, Value::number(static_cast<double>(r + 1)));
    s.set_cell_value(r + 1, 2, Value::number(static_cast<double>(r * 2)));
  }
  return wb;
}

std::uint32_t InsertColumn(Workbook& wb) {
  ChartInsertOptions o;
  o.kind = ChartInsertKind::kColumn;
  o.source = "Sheet1!$A$1:$C$4";
  return Must(insert_chart(wb, 0, o));
}

std::string PartText(const Workbook& wb) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, "xl/charts/chart1.xml");
  EXPECT_NE(part, nullptr);
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

Workbook OoxmlRoundTrip(const Workbook& wb) {
  auto saved = io::write_ooxml(wb);
  EXPECT_TRUE(static_cast<bool>(saved));
  if (!saved) {
    return Workbook::create_empty();
  }
  auto read = io::read_ooxml(io::ByteSpan{saved.value().data(), saved.value().size()});
  EXPECT_TRUE(static_cast<bool>(read));
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

const ChartAxis& AxisOf(const ChartModel& m, ChartAxisKind kind) {
  for (const ChartAxis& a : m.axes) {
    if (a.kind == kind) {
      return a;
    }
  }
  ADD_FAILURE() << "no axis of that kind";
  return m.axes.front();
}

ChartValue Num(std::uint32_t idx, double v) {
  ChartValue out;
  out.idx = idx;
  out.kind = ChartValue::Kind::kNumber;
  out.number = v;
  return out;
}

ChartValue Txt(std::uint32_t idx, std::string text) {
  ChartValue out;
  out.idx = idx;
  out.kind = ChartValue::Kind::kText;
  out.text = std::move(text);
  return out;
}

TEST(ChartPatchKeys, AxisMinorUnitGridlinesPositionAndSourceLink) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  const ChartModel before = Must(get_chart(wb, 0, id));
  const ChartAxis& value_axis = AxisOf(before, ChartAxisKind::kValue);

  ChartAxisPatch p;
  p.minor_unit = ChartPatchField<double>::Set(0.5);
  p.minor_gridlines = true;
  p.position = ChartAxisPosition::kRight;
  p.source_linked = false;
  MustOk(set_chart_axis(wb, 0, id, value_axis.id, p));
  const ChartModel after = Must(get_chart(wb, 0, id));
  const ChartAxis& a = AxisOf(after, ChartAxisKind::kValue);
  EXPECT_EQ(a.minor_unit, 0.5);
  EXPECT_TRUE(a.minor_gridlines);
  EXPECT_EQ(a.position, ChartAxisPosition::kRight);
  ASSERT_TRUE(a.num_fmt.has_value());
  EXPECT_FALSE(a.num_fmt->source_linked);
  EXPECT_NE(PartText(wb).find("<c:minorGridlines><c:spPr><a:ln w=\"9525\""), std::string::npos);

  // Feeding the read values back changes nothing.
  const std::string written = PartText(wb);
  ChartAxisPatch same;
  same.minor_unit = ChartPatchField<double>::Set(a.minor_unit);
  same.minor_gridlines = a.minor_gridlines;
  same.position = a.position;
  same.source_linked = a.num_fmt->source_linked;
  MustOk(set_chart_axis(wb, 0, id, a.id, same));
  EXPECT_EQ(PartText(wb), written);

  ChartAxisPatch automatic;
  automatic.minor_unit = ChartPatchField<double>::Clear();
  automatic.minor_gridlines = false;
  MustOk(set_chart_axis(wb, 0, id, a.id, automatic));
  const ChartModel reread = Must(get_chart(wb, 0, id));
  const ChartAxis& cleared = AxisOf(reread, ChartAxisKind::kValue);
  EXPECT_TRUE(std::isnan(cleared.minor_unit));
  EXPECT_FALSE(cleared.minor_gridlines);

  ChartAxisPatch on_category;
  on_category.minor_unit = ChartPatchField<double>::Set(1.0);
  ExpectCode(set_chart_axis(wb, 0, id, AxisOf(after, ChartAxisKind::kCategory).id, on_category),
             Code::kInvalidArgument);
  // A category axis has no units to set, so clearing them (a read value passed back) changes nothing.
  ChartAxisPatch clear_category;
  clear_category.major_unit = ChartPatchField<double>::Clear();
  clear_category.minor_unit = ChartPatchField<double>::Clear();
  MustOk(set_chart_axis(wb, 0, id, AxisOf(after, ChartAxisKind::kCategory).id, clear_category));
  ChartAxisPatch bad_position;
  bad_position.position = static_cast<ChartAxisPosition>(9);
  ExpectCode(set_chart_axis(wb, 0, id, a.id, bad_position), Code::kInvalidArgument);
}

TEST(ChartPatchKeys, LabelsSourceLinkAndDeletedBlock) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  ChartLabelsPatch on;
  on.show_val = true;
  on.number_format = ChartPatchField<std::string>::Set("0.0");
  on.source_linked = true;
  ChartSeriesPatch p;
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(on);
  MustOk(set_chart_series(wb, 0, id, 0, 0, p));
  const ChartSeries s = Must(get_chart(wb, 0, id)).plots[0].series[0];
  ASSERT_TRUE(s.labels.has_value());
  ASSERT_TRUE(s.labels->num_fmt.has_value());
  EXPECT_EQ(s.labels->num_fmt->code, "0.0");
  EXPECT_TRUE(s.labels->num_fmt->source_linked);
  EXPECT_FALSE(s.labels->deleted);

  ChartLabelsPatch gone;
  gone.deleted = true;
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(gone);
  MustOk(set_chart_series(wb, 0, id, 0, 0, p));
  EXPECT_NE(PartText(wb).find("<c:dLbls><c:delete val=\"1\"/></c:dLbls>"), std::string::npos);
  const ChartSeries d = Must(get_chart(wb, 0, id)).plots[0].series[0];
  ASSERT_TRUE(d.labels.has_value());
  EXPECT_TRUE(d.labels->deleted);
  EXPECT_FALSE(d.labels->show_val);

  // While deleted, a patch that leaves `deleted` unset keeps them deleted and refuses any other key.
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(ChartLabelsPatch{});
  MustOk(set_chart_series(wb, 0, id, 0, 0, p));
  EXPECT_TRUE(Must(get_chart(wb, 0, id)).plots[0].series[0].labels->deleted);
  ChartLabelsPatch stray;
  stray.show_cat = true;
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(stray);
  ExpectCode(set_chart_series(wb, 0, id, 0, 0, p), Code::kInvalidArgument);
  EXPECT_TRUE(Must(get_chart(wb, 0, id)).plots[0].series[0].labels->deleted);

  ChartLabelsPatch back;
  back.deleted = false;
  back.show_val = true;
  p.labels = ChartPatchField<ChartLabelsPatch>::Set(back);
  MustOk(set_chart_series(wb, 0, id, 0, 0, p));
  const ChartSeries shown = Must(get_chart(wb, 0, id)).plots[0].series[0];
  EXPECT_FALSE(shown.labels->deleted);
  EXPECT_TRUE(shown.labels->show_val);
}

TEST(ChartPatchKeys, DisplayStyleSetsClearsAndRefusesOutOfRange) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  ChartDisplayPatch p;
  p.style = ChartPatchField<std::int32_t>::Set(26);
  MustOk(set_chart_display(wb, 0, id, p));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).style, std::optional<std::int32_t>(26));
  const std::string text = PartText(wb);
  // The template's 2010 style pair is kept: the Choice reads 100 + n, the Fallback n.
  EXPECT_NE(text.find("<c14:style val=\"126\"/>"), std::string::npos);
  EXPECT_NE(text.find("<c:style val=\"26\"/>"), std::string::npos);
  EXPECT_EQ(Must(get_chart(OoxmlRoundTrip(wb), 0, id)).style, std::optional<std::int32_t>(26));

  p.style = ChartPatchField<std::int32_t>::Clear();
  MustOk(set_chart_display(wb, 0, id, p));
  EXPECT_FALSE(Must(get_chart(wb, 0, id)).style.has_value());

  for (const std::int32_t bad : {0, 49}) {
    p.style = ChartPatchField<std::int32_t>::Set(bad);
    ExpectCode(set_chart_display(wb, 0, id, p), Code::kInvalidArgument);
  }
}

TEST(ChartPatchKeys, SeriesLiteralSourcesWriteAndRoundTrip) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  ChartSeriesPatch p;
  p.name = ChartSeriesSource::Literal({Txt(0, "Typed")});
  p.cat = ChartSeriesSource::Literal({Txt(0, "x"), Txt(1, "y"), Txt(2, "z")});
  p.val = ChartSeriesSource::Literal({Num(0, 1.5), Num(2, 4.0), Num(1, 2.0)});
  MustOk(set_chart_series(wb, 0, id, 0, 0, p));
  const std::string text = PartText(wb);
  EXPECT_NE(text.find("<c:tx><c:v>Typed</c:v></c:tx>"), std::string::npos);
  EXPECT_NE(text.find("<c:cat><c:strLit><c:ptCount val=\"3\"/><c:pt idx=\"0\"><c:v>x</c:v></c:pt>"), std::string::npos);
  EXPECT_NE(text.find("<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val=\"3\"/>"
                      "<c:pt idx=\"0\"><c:v>1.5</c:v></c:pt><c:pt idx=\"1\"><c:v>2</c:v></c:pt>"),
            std::string::npos);

  const Workbook reloaded = OoxmlRoundTrip(wb);
  const Workbook* books[] = {&wb, &reloaded};
  for (const Workbook* book : books) {
    const ChartSeries s = Must(get_chart(*book, 0, id)).plots[0].series[0];
    EXPECT_EQ(s.name.kind, ChartDataSource::Kind::kLiteral);
    ASSERT_EQ(s.name.literal.size(), 1U);
    EXPECT_EQ(s.name.literal[0].text, "Typed");
    EXPECT_EQ(s.cat.kind, ChartDataSource::Kind::kLiteral);
    ASSERT_EQ(s.cat.literal.size(), 3U);
    EXPECT_EQ(s.cat.literal[2].text, "z");
    EXPECT_EQ(s.val.kind, ChartDataSource::Kind::kLiteral);
    ASSERT_EQ(s.val.literal.size(), 3U);
    EXPECT_EQ(s.val.literal[2].number, 4.0);
  }

  // The read literal fed back writes the same part.
  const ChartSeries s = Must(get_chart(wb, 0, id)).plots[0].series[0];
  ChartSeriesPatch same;
  same.val = ChartSeriesSource::Literal(s.val.literal);
  MustOk(set_chart_series(wb, 0, id, 0, 0, same));
  EXPECT_EQ(PartText(wb), text);

  // A reference replaces the literal again.
  ChartSeriesPatch ref;
  ref.val = ChartSeriesSource::Ref("Sheet1!$B$2:$B$4");
  MustOk(set_chart_series(wb, 0, id, 0, 0, ref));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).plots[0].series[0].val.kind, ChartDataSource::Kind::kRef);
}

TEST(ChartPatchKeys, SeriesSourceRefusals) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  const std::string before = PartText(wb);
  const auto refuse = [&](ChartSeriesPatch p) {
    ExpectCode(set_chart_series(wb, 0, id, 0, 0, p), Code::kInvalidArgument);
    EXPECT_EQ(PartText(wb), before);
  };
  ChartSeriesPatch p;
  p.val = ChartSeriesSource::Literal({});
  refuse(p);
  p.val = ChartSeriesSource::Literal({Num(0, 1.0), Txt(1, "a")});
  refuse(p);
  p.val = ChartSeriesSource::Literal({Num(0, 1.0), Num(0, 2.0)});
  refuse(p);
  p.val = ChartSeriesSource::Ref("");
  refuse(p);
  p.val = ChartSeriesSource::None();
  refuse(p);
  ChartSeriesPatch name;
  name.name = ChartSeriesSource::Literal({Txt(0, "a"), Txt(1, "b")});
  refuse(name);

  // A literal value input is enough to add a series.
  ChartSeriesPatch add;
  add.val = ChartSeriesSource::Literal({Num(0, 7.0), Num(1, 8.0), Num(2, 9.0)});
  const std::uint32_t idx = Must(add_chart_series(wb, 0, id, 0, add));
  const ChartModel m = Must(get_chart(wb, 0, id));
  ASSERT_EQ(m.plots[0].series.size(), 3U);
  EXPECT_EQ(m.plots[0].series[2].idx, idx);
  EXPECT_EQ(m.plots[0].series[2].val.kind, ChartDataSource::Kind::kLiteral);
}

TEST(ChartPatchKeys, PointEmptyPatchKeepsAndClearAllRemovesOverride) {
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  const std::string before = PartText(wb);
  ChartPointPatch cleared;
  cleared.fill.color = ChartPatchField<ColorSpec>::Clear();
  cleared.fill.transparency_pct = ChartPatchField<double>::Clear();
  cleared.line.color = ChartPatchField<ColorSpec>::Clear();
  cleared.line.transparency_pct = ChartPatchField<double>::Clear();
  cleared.line.width_pt = ChartPatchField<double>::Clear();
  cleared.line.dash = ChartPatchField<ChartLineDash>::Clear();
  cleared.marker = ChartPatchField<ChartMarkerPatch>::Clear();
  MustOk(set_chart_point(wb, 0, id, 0, 0, 1, cleared));
  EXPECT_EQ(PartText(wb), before);
  MustOk(set_chart_point(wb, 0, id, 0, 0, 1, ChartPointPatch{}));
  EXPECT_EQ(PartText(wb), before);

  // An override survives an empty patch and is removed by one that clears every member.
  ChartPointPatch fill;
  ColorSpec red;
  red.kind = ColorSpec::Kind::kRgb;
  red.rgb = 0xFFFF0000U;
  fill.fill.color = ChartPatchField<ColorSpec>::Set(red);
  MustOk(set_chart_point(wb, 0, id, 0, 0, 1, fill));
  ASSERT_EQ(Must(get_chart(wb, 0, id)).plots[0].series[0].points.size(), 1U);
  MustOk(set_chart_point(wb, 0, id, 0, 0, 1, ChartPointPatch{}));
  EXPECT_EQ(Must(get_chart(wb, 0, id)).plots[0].series[0].points.size(), 1U);
  MustOk(set_chart_point(wb, 0, id, 0, 0, 1, cleared));
  EXPECT_TRUE(Must(get_chart(wb, 0, id)).plots[0].series[0].points.empty());
  EXPECT_EQ(PartText(wb), before);
}

TEST(ChartPatchKeys, InsertKindFromPlotTypeAndBarDirection) {
  using T = ChartPlotType;
  using B = ChartBarDirection;
  EXPECT_EQ(Must(chart_insert_kind(T::kBar, B::kAbsent)), ChartInsertKind::kColumn);
  EXPECT_EQ(Must(chart_insert_kind(T::kBar, B::kCol)), ChartInsertKind::kColumn);
  EXPECT_EQ(Must(chart_insert_kind(T::kBar, B::kBar)), ChartInsertKind::kBar);
  EXPECT_EQ(Must(chart_insert_kind(T::kLine, B::kAbsent)), ChartInsertKind::kLine);
  EXPECT_EQ(Must(chart_insert_kind(T::kArea, B::kAbsent)), ChartInsertKind::kArea);
  EXPECT_EQ(Must(chart_insert_kind(T::kPie, B::kAbsent)), ChartInsertKind::kPie);
  EXPECT_EQ(Must(chart_insert_kind(T::kDoughnut, B::kAbsent)), ChartInsertKind::kDoughnut);
  EXPECT_EQ(Must(chart_insert_kind(T::kScatter, B::kAbsent)), ChartInsertKind::kScatter);
  ExpectCode(chart_insert_kind(T::kLine, B::kBar), Code::kInvalidArgument);
  ExpectCode(chart_insert_kind(T::kBar3D, B::kAbsent), Code::kInvalidArgument);
  ExpectCode(chart_insert_kind(T::kStock, B::kAbsent), Code::kInvalidArgument);

  // Every writable read type passes back through the mapping to the plot it came from.
  Workbook wb = GridBook();
  const std::uint32_t id = InsertColumn(wb);
  ChartInsertOptions bar;
  bar.kind = Must(chart_insert_kind(T::kBar, B::kBar));
  bar.source = "Sheet1!$A$1:$C$4";
  const std::uint32_t bar_id = Must(insert_chart(wb, 0, bar));
  const ChartPlot plot = Must(get_chart(wb, 0, bar_id)).plots[0];
  EXPECT_EQ(plot.type, T::kBar);
  EXPECT_EQ(plot.bar_dir, B::kBar);
  const std::string before = PartText(wb);
  const ChartPlot column = Must(get_chart(wb, 0, id)).plots[0];
  MustOk(set_chart_type(wb, 0, id, 0, Must(chart_insert_kind(column.type, column.bar_dir)), *column.grouping));
  EXPECT_EQ(PartText(wb), before);
}

}  // namespace
}  // namespace formulon
