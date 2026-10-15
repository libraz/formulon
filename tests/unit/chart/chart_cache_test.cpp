#include "chart/chart_cache.h"

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "../small_stack.h"
#include "chart/chart_model.h"
#include "drawing/drawing_xml.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml/part_dom.h"
#include "io/ooxml_reader.h"
#include "io/passthrough_part_index.h"
#include "io/xlsb/reader.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

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

bool Contains(const std::string& text, const std::string& needle) {
  return text.find(needle) != std::string::npos;
}

std::string DeepXml(std::size_t depth, const std::string& body) {
  std::string xml;
  xml.reserve(depth * 20U + body.size());
  for (std::size_t i = 0; i < depth; ++i) {
    xml.append("<x:deep>");
  }
  xml.append(body);
  for (std::size_t i = 0; i < depth; ++i) {
    xml.append("</x:deep>");
  }
  return xml;
}

std::string PartText(const Workbook& wb, const std::string& path) {
  const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, path);
  EXPECT_NE(part, nullptr) << path;
  return part == nullptr ? std::string() : std::string(part->bytes.begin(), part->bytes.end());
}

std::vector<PassthroughPart> Refreshed(const Workbook& wb) {
  auto got = refreshed_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(got)) << (got ? "" : got.error().message);
  return got ? std::move(got.value()) : std::vector<PassthroughPart>{};
}

std::string Text(const PassthroughPart& part) {
  return std::string(part.bytes.begin(), part.bytes.end());
}

TEST(ChartCache, PassthroughPathIndexKeepsFirstDuplicate) {
  std::vector<PassthroughPart> parts;
  parts.emplace_back("xl/charts/chart1.xml", "application/xml", std::vector<std::uint8_t>{'f', 'i', 'r', 's', 't'});
  parts.emplace_back("xl/charts/chart1.xml", "application/xml",
                     std::vector<std::uint8_t>{'s', 'e', 'c', 'o', 'n', 'd'});
  parts.emplace_back("xl/charts/chart2.xml", "application/xml", std::vector<std::uint8_t>{'t', 'w', 'o'});

  const io::PassthroughPartIndex index(parts);
  const PassthroughPart* first = index.find("xl/charts/chart1.xml");
  ASSERT_NE(first, nullptr);
  EXPECT_EQ(Text(*first), "first");
  EXPECT_EQ(index.find("xl/charts/missing.xml"), nullptr);
}

/// Path of the first chart part whose text contains `needle`.
std::string ChartPathWith(const Workbook& wb, const std::string& needle) {
  auto refs = all_chart_parts(wb);
  EXPECT_TRUE(static_cast<bool>(refs));
  if (refs) {
    for (const SheetChartPartRef& entry : refs.value()) {
      const PassthroughPart* part = io::ooxml::find_passthrough_part(wb, entry.ref.part_path);
      if (part != nullptr && Contains(std::string(part->bytes.begin(), part->bytes.end()), needle)) {
        return entry.ref.part_path;
      }
    }
  }
  ADD_FAILURE() << "no chart part with " << needle;
  return {};
}

TEST(ChartCache, RefreshUsesFirstDuplicateChartPart) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  std::vector<PassthroughPart> parts = wb.passthrough_parts();
  std::size_t position = parts.size();
  for (std::size_t i = 0; i < parts.size(); ++i) {
    if (parts[i].path == path) {
      position = i;
      break;
    }
  }
  ASSERT_LT(position, parts.size());
  wb.sheet(0).set_cell_value(1, 1, Value::number(99));
  ASSERT_FALSE(Refreshed(wb).empty());
  PassthroughPart duplicate = parts[position];
  duplicate.bytes = {'n', 'o', 't', ' ', 'x', 'm', 'l'};
  parts.insert(parts.begin() + static_cast<std::ptrdiff_t>(position), std::move(duplicate));
  wb.set_passthrough_parts(std::move(parts));
  auto charts = all_chart_parts(wb);
  ASSERT_TRUE(static_cast<bool>(charts));
  ASSERT_FALSE(charts.value().empty());

  // Refresh reads the first duplicate, so its malformed bytes suppress
  // replacement for this path; other charts still refresh.
  bool duplicate_was_refreshed = false;
  for (const PassthroughPart& part : Refreshed(wb)) {
    duplicate_was_refreshed = duplicate_was_refreshed || part.path == path;
  }
  EXPECT_FALSE(duplicate_was_refreshed);
}

/// The refreshed part at `path`, or an empty string when it was not rewritten.
std::string RefreshedText(const Workbook& wb, const std::string& path) {
  for (const PassthroughPart& part : Refreshed(wb)) {
    if (part.path == path) {
      return Text(part);
    }
  }
  return {};
}

void Replace(std::string& text, const std::string& from, const std::string& to) {
  const std::size_t at = text.find(from);
  ASSERT_NE(at, std::string::npos) << from;
  text.replace(at, from.size(), to);
}

void SetPart(Workbook& wb, const std::string& path, const std::string& text) {
  auto stored = wb.replace_passthrough_part(path, std::vector<std::uint8_t>(text.begin(), text.end()));
  ASSERT_TRUE(static_cast<bool>(stored));
}

// Excel's own caches already match its cells: nothing is rewritten.
TEST(ChartCache, ExcelSavedFixturesNeedNoRefresh) {
  const char* const names[] = {"chart_area.xlsx",       "chart_bar.xlsx",    "chart_cache.xlsx",  "chart_chartex.xlsx",
                               "chart_chartsheet.xlsx", "chart_column.xlsx", "chart_combo.xlsx",  "chart_doughnut.xlsx",
                               "chart_data.xlsx",       "chart_filter.xlsx", "chart_format.xlsx", "chart_line.xlsx",
                               "chart_pie.xlsx",        "chart_pivot.xlsx",  "chart_refs.xlsx",   "chart_scatter.xlsx",
                               "chart_xlsb.xlsb"};
  for (const char* name : names) {
    const Workbook wb = LoadFixture(name);
    auto got = refreshed_chart_parts(wb);
    ASSERT_TRUE(static_cast<bool>(got)) << name;
    for (const PassthroughPart& part : got.value()) {
      ADD_FAILURE() << name << " refreshed " << part.path << "\n" << Text(part);
    }
  }
}

// A changed cell is cached in Excel's notation: 15 digits when they round-trip, else 17.
TEST(ChartCache, ChangedCellsAreCachedInExcelNotation) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  wb.sheet(0).set_cell_value(1, 1, Value::number(0.1 + 0.2));
  wb.sheet(0).set_cell_value(2, 1, Value::number(1.0 / 3.0));
  wb.sheet(0).set_cell_value(3, 1, Value::number(1e-20));
  wb.sheet(0).set_cell_value(4, 1, Value::number(0.1));
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"0\"><c:v>0.30000000000000004</c:v></c:pt>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>0.33333333333333331</c:v></c:pt>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"2\"><c:v>9.9999999999999995E-21</c:v></c:pt>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"3\"><c:v>0.1</c:v></c:pt>"));
  // The rewrite is idempotent: stored back, nothing is left to refresh for that part.
  SetPart(wb, path, text);
  EXPECT_EQ(RefreshedText(wb, path), "");
}

TEST(ChartCache, DeepExtensionWalkKeepsMatchedReferenceSubtreesPruned) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  std::string chart = PartText(wb, path);

  const std::size_t root_close = chart.rfind("</c:chartSpace>");
  ASSERT_NE(root_close, std::string::npos);
  chart.insert(root_close, DeepXml(10000U, "<x:leaf/>"));

  const std::size_t num_ref = chart.find("<c:numRef>");
  ASSERT_NE(num_ref, std::string::npos);
  const std::size_t num_ref_close = chart.find("</c:numRef>", num_ref);
  ASSERT_NE(num_ref_close, std::string::npos);
  chart.insert(num_ref_close, "<c:ser><c:val><c:numRef><c:f>Sheet1!$Z$1:$Z$1</c:f></c:numRef></c:val></c:ser>");
  SetPart(wb, path, chart);
  wb.sheet(0).set_cell_value(1, 1, Value::number(50));
  wb.sheet(0).set_cell_value(0, 25, Value::number(7777777.25));

  std::string refreshed;
  testing_support::RunOnSmallStack([&] { refreshed = RefreshedText(wb, path); });
  ASSERT_FALSE(refreshed.empty());
  EXPECT_TRUE(Contains(refreshed, "<c:pt idx=\"0\"><c:v>50</c:v></c:pt>"));
  EXPECT_FALSE(Contains(refreshed, "7777777.25"));
}

// A cache gone stale is rewritten to what Excel wrote for the same cells.
TEST(ChartCache, StaleCacheIsRefreshedToTheOriginal) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  const std::string original = PartText(wb, path);
  std::string stale = original;
  Replace(stale, "<c:v>S1</c:v>", "<c:v>OLD</c:v>");
  Replace(stale, "<c:v>7</c:v>", "<c:v>111</c:v>");
  Replace(stale, "<c:ptCount val=\"4\"/><c:pt idx=\"0\"><c:v>a</c:v>",
          "<c:ptCount val=\"9\"/><c:pt idx=\"0\"><c:v>a</c:v>");
  SetPart(wb, path, stale);
  const std::string refreshed = RefreshedText(wb, path);
  ASSERT_FALSE(refreshed.empty());
  EXPECT_TRUE(Contains(refreshed, "<c:v>S1</c:v>"));
  EXPECT_FALSE(Contains(refreshed, "OLD"));
  EXPECT_TRUE(Contains(refreshed, "<c:v>7</c:v>"));
  EXPECT_FALSE(Contains(refreshed, "111"));
  EXPECT_FALSE(Contains(refreshed, "<c:ptCount val=\"9\"/>"));
  // The same refresh applied to the original yields the same bytes as Excel's part.
  SetPart(wb, path, refreshed);
  EXPECT_EQ(RefreshedText(wb, path), "");
}

TEST(ChartCache, RefAndExternalReferencesAreLeftAlone) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  std::string text = PartText(wb, path);
  Replace(text, "<c:f>Sheet1!$B$2:$B$5</c:f>", "<c:f>#REF!</c:f>");
  Replace(text, "<c:f>Sheet1!$A$2:$A$5</c:f>", "<c:f>[1]Sheet1!$A$2:$A$5</c:f>");
  Replace(text, "<c:v>S1</c:v>", "<c:v>OLD</c:v>");  // the one reference still resolvable
  SetPart(wb, path, text);
  wb.sheet(0).set_cell_value(1, 1, Value::number(50));
  const std::string refreshed = RefreshedText(wb, path);
  ASSERT_FALSE(refreshed.empty());
  EXPECT_TRUE(Contains(refreshed, "<c:v>S1</c:v>"));
  EXPECT_FALSE(Contains(refreshed, "<c:v>50</c:v>"));
  EXPECT_TRUE(Contains(refreshed, "<c:f>#REF!</c:f>"));
  EXPECT_TRUE(Contains(refreshed, "<c:v>d</c:v>"));
}

TEST(ChartCache, PivotAndChartexPartsAreNotRefreshed) {
  for (const char* name : {"chart_pivot.xlsx", "chart_chartex.xlsx"}) {
    Workbook wb = LoadFixture(name);
    auto refs = all_chart_parts(wb);
    ASSERT_TRUE(static_cast<bool>(refs)) << name;
    for (const SheetChartPartRef& entry : refs.value()) {
      std::string text = PartText(wb, entry.ref.part_path);
      const std::size_t at = text.find("<c:v>");
      if (at != std::string::npos) {
        text.replace(at, 5, "<c:v>x");  // stale the first cached value
        SetPart(wb, entry.ref.part_path, text);
      }
    }
    EXPECT_TRUE(Refreshed(wb).empty()) << name;
  }
}

// A whole-column reference keeps Excel's full count; only non-blank cells carry a point.
TEST(ChartCache, WholeColumnUsesTheFullPointCount) {
  Workbook wb = LoadFixture("chart_refs.xlsx");
  const std::string path = ChartPathWith(wb, "<c:f>Sheet1!$B:$B</c:f>");
  wb.sheet(0).set_cell_value(1, 1, Value::number(100));
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:ptCount val=\"1048576\"/>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>100</c:v></c:pt>"));
  EXPECT_FALSE(Contains(text, "<c:pt idx=\"1048575\">"));
}

TEST(ChartCache, MultiLevelCategoryRefreshesBothLevels) {
  Workbook wb = LoadFixture("chart_data.xlsx");
  const std::string path = ChartPathWith(wb, "<c:f>Sheet1!$F$2:$G$9</c:f>");
  wb.sheet(0).set_cell_value(1, 5, Value::text("Z1"));
  wb.sheet(0).set_cell_value(2, 6, Value::text("new"));
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>new</c:v></c:pt>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"0\"><c:v>Z1</c:v></c:pt>"));
  EXPECT_FALSE(Contains(text, "<c:v>G1</c:v>"));
}

// A series hidden by a chart filter keeps its cache current.
TEST(ChartCache, FilteredSeriesIsRefreshed) {
  Workbook wb = LoadFixture("chart_filter.xlsx");
  const std::string path = ChartPathWith(wb, "<c15:sqref>Sheet1!$C$2:$C$5</c15:sqref>");
  wb.sheet(0).set_cell_value(2, 2, Value::number(99));
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>99</c:v></c:pt>"));
  EXPECT_FALSE(Contains(text, "<c:v>5</c:v></c:pt><c:pt idx=\"2\"><c:v>6"));
}

// A numeric category that gains text becomes a string reference.
TEST(ChartCache, NumericCategoryWithTextBecomesStringReference) {
  Workbook wb = LoadFixture("chart_cache.xlsx");
  const std::string path = ChartPathWith(wb, "<c:f>Sheet1!$J$2:$J$5</c:f>");
  wb.sheet(0).set_cell_value(2, 9, Value::text("x"));
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:cat><c:strRef><c:f>Sheet1!$J$2:$J$5</c:f><c:strCache><c:ptCount val=\"4\"/>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>x</c:v></c:pt>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"0\"><c:v>10</c:v></c:pt>"));
}

TEST(ChartCache, LegacyStringCacheEscapesOoxmlNotationOnceAcrossMutation) {
  Workbook wb = LoadFixture("chart_cache.xlsx");
  const std::string path = ChartPathWith(wb, "<c:f>Sheet1!$J$2:$J$5</c:f>");
  const std::string first_value = "literal _x0041_ _x005F_x0041_ & <猫>\r\x01";
  wb.sheet(0).set_cell_value(2, 9, Value::text(first_value));

  const std::string first = RefreshedText(wb, path);
  ASSERT_FALSE(first.empty());
  EXPECT_TRUE(Contains(first,
                       "<c:v>literal _x005F_x0041_ _x005F_x005F_x005F_x0041_ &amp; &lt;猫&gt;"
                       "_x000D__x0001_</c:v>"));
  EXPECT_FALSE(Contains(first, "&amp;amp;"));
  SetPart(wb, path, first);
  EXPECT_TRUE(RefreshedText(wb, path).empty());

  const std::string changed_value = "changed _x0041_ _x005F_x0041_ & <魚>\r\x02";
  wb.sheet(0).set_cell_value(2, 9, Value::text(changed_value));
  const std::string second = RefreshedText(wb, path);
  ASSERT_FALSE(second.empty());
  EXPECT_TRUE(Contains(second,
                       "<c:v>changed _x005F_x0041_ _x005F_x005F_x005F_x0041_ &amp; &lt;魚&gt;"
                       "_x000D__x0002_</c:v>"));
  EXPECT_FALSE(Contains(second, "&amp;amp;"));
  SetPart(wb, path, second);
  EXPECT_TRUE(RefreshedText(wb, path).empty());
}

// A formula's "" in a value range is a 0 point; in a category it is no point.
TEST(ChartCache, EmptyStringIsZeroInValuesAndAbsentInCategories) {
  Workbook wb = LoadFixture("chart_column.xlsx");
  const std::string path = ChartPathWith(wb, "Sheet1!$B$2:$B$5");
  Sheet& s = wb.sheet(0);
  for (std::uint32_t col : {0U, 1U}) {
    s.set_cell_formula(2, col, "=\"\"");
    s.set_cell_cached_value(2, col, Value::text(""));
  }
  const std::string text = RefreshedText(wb, path);
  ASSERT_FALSE(text.empty());
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"1\"><c:v>0</c:v></c:pt>"));
  EXPECT_FALSE(Contains(text, "<c:v>b</c:v>"));
  EXPECT_TRUE(Contains(text, "<c:pt idx=\"2\"><c:v>c</c:v></c:pt>"));
}

// With every point hidden, a reference keeps `c:f` and loses its cache.
TEST(ChartCache, FullyHiddenReferencesLoseTheirCache) {
  Workbook wb = LoadFixture("chart_data.xlsx");
  const std::string multi = ChartPathWith(wb, "<c:f>Sheet1!$F$2:$G$9</c:f>");
  const std::string flat = ChartPathWith(wb, "<c:f>Sheet1!$A$2:$A$9</c:f>");
  for (std::uint32_t row = 1; row <= 8; ++row) {
    RowLayout entry;
    entry.row = row;
    entry.hidden = true;
    insert_row_override(wb.sheet(0).mutable_layout(), entry);
  }
  for (const std::string& path : {multi, flat}) {
    const std::string text = RefreshedText(wb, path);
    ASSERT_FALSE(text.empty()) << path;
    EXPECT_FALSE(Contains(text, "multiLvlStrCache")) << path;
    EXPECT_FALSE(Contains(text, "<c:v>i1</c:v>")) << path;
  }
  EXPECT_TRUE(Contains(RefreshedText(wb, multi), "<c:f>Sheet1!$F$2:$G$9</c:f></c:multiLvlStrRef>"));
}

}  // namespace
}  // namespace formulon
