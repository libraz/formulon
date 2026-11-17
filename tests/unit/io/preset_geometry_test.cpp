#include "io/ooxml/preset_geometry.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string_view>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml/preset_geometry_table.h"

namespace formulon {
namespace preset_geometry {
namespace table_check {
#define FORMULON_PRESET_GEOMETRY_TABLE_NAMES
#include "io/ooxml/preset_geometry_table.inc"
}  // namespace table_check
namespace {

TEST(PresetGeometry, TableFollowsTheNameOrder) {
  ASSERT_EQ(std::size(table_check::kPresetTableNames), preset_count());
  for (std::size_t i = 0; i < preset_count(); ++i) {
    EXPECT_EQ(table_check::kPresetTableNames[i], name_at(i)) << i;
  }
  EXPECT_EQ(index_of("rect"), std::optional<std::size_t>(*index_of("rect")));
  EXPECT_EQ(name_at(*index_of("wedgeRectCallout")), "wedgeRectCallout");
  EXPECT_FALSE(index_of("Rect").has_value());
  EXPECT_FALSE(index_of("").has_value());
  EXPECT_FALSE(index_of("zzz").has_value());
}

std::vector<AdjustGuide> Guides(std::string_view prst) {
  return adjust_guides(*index_of(prst));
}

TEST(PresetGeometry, AdjustGuidesAndDefaults) {
  EXPECT_TRUE(Guides("rect").empty());
  ASSERT_EQ(Guides("roundRect").size(), 1U);
  EXPECT_EQ(Guides("roundRect")[0].name, "adj");
  EXPECT_EQ(Guides("roundRect")[0].value, 16667);
  const std::vector<AdjustGuide> arrow = Guides("rightArrow");
  ASSERT_EQ(arrow.size(), 2U);
  EXPECT_EQ(arrow[0].name, "adj1");
  EXPECT_EQ(arrow[0].value, 50000);
  EXPECT_EQ(arrow[1].name, "adj2");
  const std::vector<AdjustGuide> callout = Guides("wedgeRectCallout");
  ASSERT_EQ(callout.size(), 2U);
  EXPECT_EQ(callout[0].value, -20833);
  EXPECT_EQ(callout[1].value, 62500);
  EXPECT_EQ(Guides("parallelogram")[0].value, 25000);
}

TEST(PresetGeometry, PublishedSiteCounts) {
  const std::pair<std::string_view, std::size_t> counts[] = {
      {"rect", 4},      {"ellipse", 8},    {"triangle", 6},           {"diamond", 4},
      {"roundRect", 4}, {"rightArrow", 4}, {"wedgeRectCallout", 5},   {"parallelogram", 6},
      {"trapezoid", 4}, {"upArrow", 4},    {"straightConnector1", 0},
  };
  for (const auto& [prst, count] : counts) {
    EXPECT_EQ(connection_sites(*index_of(prst), 1000.0, 1000.0, {}).size(), count) << prst;
  }
}

struct Row {
  const char* prst;
  double w, h;
  std::vector<AdjustGuide> adjusts;
  std::size_t idx;
  double x, y;
};

// Mac Excel 16.115: where the end of a connector glued to site `idx` of the shape lands, in EMU from the shape's
// top-left corner, for the shape's size and stated adjust values.
TEST(PresetGeometry, SitesMatchExcel) {
  const Row rows[] = {
      {"diamond", 1524000, 1016000, {}, 0, 762000, 0},
      {"diamond", 1524000, 1016000, {}, 1, 0, 508000},
      {"diamond", 1524000, 1016000, {}, 2, 762000, 1016000},
      {"diamond", 1524000, 1016000, {}, 3, 1524000, 508000},
      {"diamond", 2032000, 1270000, {}, 0, 1016000, 0},
      {"diamond", 2032000, 1270000, {}, 1, 0, 635000},
      {"diamond", 2032000, 1270000, {}, 2, 1016000, 1270000},
      {"diamond", 2032000, 1270000, {}, 3, 2032000, 635000},
      {"ellipse", 1524000, 1016000, {}, 0, 762000, 0},
      {"ellipse", 1524000, 1016000, {}, 1, 223185, 148790},
      {"ellipse", 1524000, 1016000, {}, 2, 0, 508000},
      {"ellipse", 1524000, 1016000, {}, 3, 223185, 867210},
      {"ellipse", 1524000, 1016000, {}, 4, 762000, 1016000},
      {"ellipse", 1524000, 1016000, {}, 5, 1300815, 867210},
      {"ellipse", 1524000, 1016000, {}, 6, 1524000, 508000},
      {"ellipse", 1524000, 1016000, {}, 7, 1300815, 148790},
      {"ellipse", 2032000, 1270000, {}, 0, 1016000, 0},
      {"ellipse", 2032000, 1270000, {}, 1, 297580, 185987},
      {"ellipse", 2032000, 1270000, {}, 2, 0, 635000},
      {"ellipse", 2032000, 1270000, {}, 3, 297580, 1084013},
      {"ellipse", 2032000, 1270000, {}, 4, 1016000, 1270000},
      {"ellipse", 2032000, 1270000, {}, 5, 1734420, 1084013},
      {"ellipse", 2032000, 1270000, {}, 6, 2032000, 635000},
      {"ellipse", 2032000, 1270000, {}, 7, 1734420, 185987},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 0, 762000, 0},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 1, 889000, 0},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 2, 1397000, 508000},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 3, 635000, 1016000},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 4, 762000, 1016000},
      {"parallelogram", 1524000, 1016000, {{"adj", 25000}}, 5, 127000, 508000},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 0, 762000, 0},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 1, 1016000, 0},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 2, 1270000, 508000},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 3, 508000, 1016000},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 4, 762000, 1016000},
      {"parallelogram", 1524000, 1016000, {{"adj", 50000}}, 5, 254000, 508000},
      {"rect", 1016000, 508000, {}, 1, 0, 254000},
      {"rect", 1016000, 508000, {}, 3, 1016000, 254000},
      {"rightArrow", 1524000, 1016000, {{"adj1", 50000}, {"adj2", 50000}}, 0, 1016000, 0},
      {"rightArrow", 1524000, 1016000, {{"adj1", 50000}, {"adj2", 50000}}, 1, 0, 508000},
      {"rightArrow", 1524000, 1016000, {{"adj1", 50000}, {"adj2", 50000}}, 2, 1016000, 1016000},
      {"rightArrow", 1524000, 1016000, {{"adj1", 50000}, {"adj2", 50000}}, 3, 1524000, 508000},
      {"rightArrow", 2032000, 1270000, {{"adj1", 50000}, {"adj2", 50000}}, 0, 1397000, 0},
      {"rightArrow", 2032000, 1270000, {{"adj1", 50000}, {"adj2", 50000}}, 1, 0, 635000},
      {"rightArrow", 2032000, 1270000, {{"adj1", 50000}, {"adj2", 50000}}, 2, 1397000, 1270000},
      {"rightArrow", 2032000, 1270000, {{"adj1", 50000}, {"adj2", 50000}}, 3, 2032000, 635000},
      {"roundRect", 1524000, 1016000, {{"adj", 16667}}, 0, 762000, 0},
      {"roundRect", 1524000, 1016000, {{"adj", 16667}}, 1, 0, 508000},
      {"roundRect", 1524000, 1016000, {{"adj", 16667}}, 2, 762000, 1016000},
      {"roundRect", 1524000, 1016000, {{"adj", 16667}}, 3, 1524000, 508000},
      {"roundRect", 2032000, 1270000, {{"adj", 16667}}, 0, 1016000, 0},
      {"roundRect", 2032000, 1270000, {{"adj", 16667}}, 1, 0, 635000},
      {"roundRect", 2032000, 1270000, {{"adj", 16667}}, 2, 1016000, 1270000},
      {"roundRect", 2032000, 1270000, {{"adj", 16667}}, 3, 2032000, 635000},
      {"trapezoid", 1524000, 1016000, {{"adj", 25000}}, 0, 762000, 0},
      {"trapezoid", 1524000, 1016000, {{"adj", 25000}}, 1, 127000, 508000},
      {"trapezoid", 1524000, 1016000, {{"adj", 25000}}, 2, 762000, 1016000},
      {"trapezoid", 1524000, 1016000, {{"adj", 25000}}, 3, 1397000, 508000},
      {"trapezoid", 1524000, 1016000, {{"adj", 50000}}, 0, 762000, 0},
      {"trapezoid", 1524000, 1016000, {{"adj", 50000}}, 1, 254000, 508000},
      {"trapezoid", 1524000, 1016000, {{"adj", 50000}}, 2, 762000, 1016000},
      {"trapezoid", 1524000, 1016000, {{"adj", 50000}}, 3, 1270000, 508000},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 0, 762000, 0},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 1, 381000, 508000},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 2, 0, 1016000},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 3, 762000, 1016000},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 4, 1524000, 1016000},
      {"triangle", 1524000, 1016000, {{"adj", 50000}}, 5, 1143000, 508000},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 0, 1016000, 0},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 1, 508000, 635000},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 2, 0, 1270000},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 3, 1016000, 1270000},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 4, 2032000, 1270000},
      {"triangle", 2032000, 1270000, {{"adj", 50000}}, 5, 1524000, 635000},
      {"wedgeRectCallout", 1524000, 1016000, {{"adj1", -20833}, {"adj2", 62500}}, 0, 762000, 0},
      {"wedgeRectCallout", 1524000, 1016000, {{"adj1", -20833}, {"adj2", 62500}}, 1, 0, 508000},
      {"wedgeRectCallout", 1524000, 1016000, {{"adj1", -20833}, {"adj2", 62500}}, 2, 762000, 1016000},
      {"wedgeRectCallout", 1524000, 1016000, {{"adj1", -20833}, {"adj2", 62500}}, 3, 1524000, 508000},
      {"wedgeRectCallout", 1524000, 1016000, {{"adj1", -20833}, {"adj2", 62500}}, 4, 444505, 1143000},
      {"wedgeRectCallout", 2032000, 1270000, {{"adj1", -20833}, {"adj2", 62500}}, 0, 1016000, 0},
      {"wedgeRectCallout", 2032000, 1270000, {{"adj1", -20833}, {"adj2", 62500}}, 1, 0, 635000},
      {"wedgeRectCallout", 2032000, 1270000, {{"adj1", -20833}, {"adj2", 62500}}, 2, 1016000, 1270000},
      {"wedgeRectCallout", 2032000, 1270000, {{"adj1", -20833}, {"adj2", 62500}}, 3, 2032000, 635000},
      {"wedgeRectCallout", 2032000, 1270000, {{"adj1", -20833}, {"adj2", 62500}}, 4, 592673, 1428750},
      // upArrow, which the published definitions omit, as Excel places its sites (default and adj2 70000).
      {"upArrow", 1524000, 1016000, {}, 0, 762000, 0},
      {"upArrow", 1524000, 1016000, {}, 1, 0, 508000},
      {"upArrow", 1524000, 1016000, {}, 2, 762000, 1016000},
      {"upArrow", 1524000, 1016000, {}, 3, 1524000, 508000},
      {"upArrow", 2032000, 1270000, {{"adj2", 70000}}, 1, 0, 889000},
      {"upArrow", 2032000, 1270000, {{"adj2", 70000}}, 3, 2032000, 889000},
  };
  for (const Row& row : rows) {
    const std::vector<Point> sites = connection_sites(*index_of(row.prst), row.w, row.h, row.adjusts);
    ASSERT_LT(row.idx, sites.size()) << row.prst;
    EXPECT_NEAR(sites[row.idx].x, row.x, 0.5) << row.prst << " " << row.w << "x" << row.h << " idx " << row.idx;
    EXPECT_NEAR(sites[row.idx].y, row.y, 0.5) << row.prst << " " << row.w << "x" << row.h << " idx " << row.idx;
  }
}

TEST(PresetGeometry, StatedAdjustValuesMoveTheSites) {
  const std::size_t p = *index_of("parallelogram");
  const double plain = connection_sites(p, 1524000.0, 1016000.0, {})[2].x;
  const double wide = connection_sites(p, 1524000.0, 1016000.0, {{"adj", 50000}})[2].x;
  EXPECT_NEAR(plain, 1524000.0 * (1.0 - 25000.0 / 100000.0 * 1016000.0 / 2.0 / 1524000.0), 0.5);
  EXPECT_LT(wide, plain);
  // A guide the preset does not define is ignored.
  EXPECT_EQ(connection_sites(p, 1524000.0, 1016000.0, {{"adj9", 1}})[2].x, plain);
}

}  // namespace
}  // namespace preset_geometry
}  // namespace formulon
