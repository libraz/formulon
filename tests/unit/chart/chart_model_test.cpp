#include "chart/chart_model.h"

#include <cmath>

#include "gtest/gtest.h"
#include "parser/ref_transforms.h"
#include "workbook_ref_rewrite.h"

namespace formulon {
namespace {

TEST(ChartModel, DefaultModelHasNothingPresent) {
  const ChartModel m;
  EXPECT_TRUE(m.plots.empty());
  EXPECT_TRUE(m.axes.empty());
  EXPECT_EQ(m.title.kind, ChartTitleKind::kNone);
  EXPECT_EQ(m.auto_title_deleted, Tri::kAbsent);
  EXPECT_EQ(m.disp_blanks_as, ChartBlanksAs::kAbsent);
  EXPECT_EQ(m.plot_vis_only, Tri::kAbsent);
  EXPECT_FALSE(m.legend.present);
  EXPECT_FALSE(m.style.has_value());
  EXPECT_FALSE(m.source.has_value());
}

TEST(ChartModel, SeriesAndAxisDefaultsAreAbsent) {
  const ChartSeries s;
  EXPECT_EQ(s.name.kind, ChartDataSource::Kind::kNone);
  EXPECT_FALSE(s.fill.color.has_value());
  EXPECT_FALSE(s.line.width_emu.has_value());
  EXPECT_FALSE(s.marker.has_value());
  EXPECT_EQ(s.smooth, Tri::kAbsent);
  EXPECT_FALSE(s.labels.has_value());

  const ChartAxis a;
  EXPECT_TRUE(std::isnan(a.min));
  EXPECT_TRUE(std::isnan(a.max));
  EXPECT_TRUE(std::isnan(a.major_unit));
  EXPECT_TRUE(std::isnan(a.minor_unit));
  EXPECT_FALSE(a.num_fmt.has_value());
}

TEST(ChartModel, InsertOptionsDefaultToTwoCellAnchorAndAutoTitle) {
  const ChartInsertOptions o;
  EXPECT_EQ(o.placement.anchor_kind, AnchorKind::kTwoCell);
  EXPECT_FALSE(o.title.has_value());
  EXPECT_EQ(o.direction, ChartSeriesDirection::kAuto);
  EXPECT_EQ(o.header_row, ChartHeaderMode::kAuto);
}

TEST(ChartModel, PatchFieldDistinguishesKeepClearSet) {
  ChartPatchField<double> f;
  EXPECT_TRUE(f.is_keep());
  EXPECT_TRUE(ChartPatchField<double>::Clear().is_clear());
  const auto set = ChartPatchField<double>::Set(2.5);
  EXPECT_TRUE(set.is_set());
  EXPECT_EQ(set.value, 2.5);
}

TEST(ChartModel, RewriteFormulaFollowsRowInsert) {
  const parser::RowColShiftTransform t("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kInsert,
                                       /*index=*/1, /*count=*/1);
  const FormulaRewriteResult r = rewrite_formula("Sheet1!$A$2:$A$4", t);
  EXPECT_FALSE(r.error);
  EXPECT_TRUE(r.changed);
  EXPECT_EQ(r.text, "Sheet1!$A$3:$A$5");
}

TEST(ChartModel, RewriteFormulaLeavesOtherSheetAlone) {
  const parser::RowColShiftTransform t("Sheet1", parser::RowColAxis::kRow, parser::RowColEdit::kInsert, 1, 1);
  const FormulaRewriteResult r = rewrite_formula("Sheet2!$A$2:$A$4", t);
  EXPECT_FALSE(r.changed);
  EXPECT_EQ(r.text, "Sheet2!$A$2:$A$4");
}

}  // namespace
}  // namespace formulon
