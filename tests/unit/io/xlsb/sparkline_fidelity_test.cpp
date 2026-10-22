// Sparkline groups in XLSB: the SPARKLINEGROUPS future records are read into
// the sheet's `x14:sparklineGroups` XML, edited through
// `io/sparkline_groups.h` like an .xlsx-loaded sheet's, and written back.
// The Excel fixture is the Windows Excel 365 save under
// tests/fixtures/excel/win/x14_structural/ (a line sparkline at G1 over
// Sheet1!E1:E6).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/sparkline_groups.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "value.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/win/x14_structural/" + file;
  FILE* f = std::fopen(path.c_str(), "rb");
  if (f == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(f, 0, SEEK_END);
  const long size = std::ftell(f);
  std::fseek(f, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), f) != out.size()) {
    out.clear();
  }
  std::fclose(f);
  return out;
}

std::vector<std::uint8_t> SaveXlsb(const Workbook& wb) {
  auto bytes = write_xlsb(wb);
  EXPECT_TRUE(static_cast<bool>(bytes)) << (bytes ? "" : bytes.error().message);
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

Workbook LoadXlsb(const std::vector<std::uint8_t>& bytes) {
  auto r = read_xlsb(SpanOf(bytes));
  EXPECT_TRUE(static_cast<bool>(r)) << (r ? "" : r.error().message);
  return r ? std::move(r.value().workbook) : Workbook::create_empty();
}

/// The framed bytes of the sheet's SPARKLINEGROUPS container, BrtFRTBegin
/// through BrtFRTEnd.
std::vector<std::uint8_t> SparklineContainer(const std::vector<std::uint8_t>& package) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto entry = zip.read_entry("xl/worksheets/sheet1.bin");
  EXPECT_TRUE(static_cast<bool>(entry));
  std::vector<std::uint8_t> out;
  if (!entry) {
    return out;
  }
  const std::vector<std::uint8_t>& bytes = entry.value();
  const std::vector<FramedRecord> recs = split_records(bytes);
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type == 35U && recs[i + 1U].type == 1058U) {
      std::size_t k = i;
      while (k < recs.size() && recs[k].type != 36U) {
        ++k;
      }
      if (k < recs.size()) {
        out.assign(bytes.begin() + static_cast<std::ptrdiff_t>(recs[i].begin),
                   bytes.begin() + static_cast<std::ptrdiff_t>(recs[k].end));
      }
      break;
    }
  }
  return out;
}

TEST(XlsbSparklineFidelity, ExcelGroupsAreListed) {
  const Workbook wb = LoadXlsb(ReadFixture("union_databar_0_base.xlsb"));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(wb.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  const SparklineGroup& g = groups[0];
  EXPECT_EQ(g.type, SparklineType::kLine);
  EXPECT_EQ(g.empty_cells_as, SparklineEmptyCells::kGap);
  EXPECT_EQ(g.color_series.kind, ColorSpec::Kind::kRgb);
  EXPECT_EQ(g.color_series.rgb, 0xFF376092U);
  EXPECT_EQ(g.color_negative.rgb, 0xFFD00000U);
  EXPECT_EQ(g.min_axis_type, SparklineAxisType::kIndividual);
  EXPECT_DOUBLE_EQ(g.line_weight, 0.75);
  EXPECT_EQ(g.uid, "{85D01F1B-F726-4C98-8685-D838BDDBA70F}");
  ASSERT_EQ(g.items.size(), 1U);
  EXPECT_EQ(g.items[0].formula, "Sheet1!E1:E6");
  EXPECT_EQ(g.items[0].location, "G1");
}

TEST(XlsbSparklineFidelity, ExcelGroupsAreRewrittenByteForByte) {
  const std::vector<std::uint8_t> excel = ReadFixture("union_databar_0_base.xlsb");
  const std::vector<std::uint8_t> container = SparklineContainer(excel);
  ASSERT_FALSE(container.empty());
  EXPECT_EQ(SparklineContainer(SaveXlsb(LoadXlsb(excel))), container);
}

TEST(XlsbSparklineFidelity, EditsOfAnXlsbLoadedGroupAreWritten) {
  Workbook wb = LoadXlsb(ReadFixture("union_databar_0_base.xlsb"));
  std::string xml = wb.sheet(0).ext_lst_xml();
  SparklineGroup g = read_sparkline_groups(xml).at(0);
  g.type = SparklineType::kColumn;
  g.markers = true;
  g.high = true;
  g.max_axis_type = SparklineAxisType::kCustom;
  g.has_manual_max = true;
  g.manual_max = 9.5;
  g.color_high.kind = ColorSpec::Kind::kTheme;
  g.color_high.theme = 5;
  g.color_high.tint = -0.25;
  ASSERT_TRUE(static_cast<bool>(update_sparkline_group(xml, 0, g)));
  ASSERT_TRUE(static_cast<bool>(set_group_sparklines(xml, 0, {{"Sheet1!E1:E6", "G1"}, {"Sheet1!C1:C6", "H2"}})));
  wb.sheet(0).set_ext_lst_xml(xml);

  const Workbook back = LoadXlsb(SaveXlsb(wb));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(back.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].type, SparklineType::kColumn);
  EXPECT_TRUE(groups[0].markers);
  EXPECT_TRUE(groups[0].high);
  EXPECT_EQ(groups[0].max_axis_type, SparklineAxisType::kCustom);
  EXPECT_DOUBLE_EQ(groups[0].manual_max, 9.5);
  EXPECT_EQ(groups[0].color_high.kind, ColorSpec::Kind::kTheme);
  EXPECT_EQ(groups[0].color_high.theme, 5U);
  EXPECT_NEAR(groups[0].color_high.tint, -0.25, 1e-4);
  ASSERT_EQ(groups[0].items.size(), 2U);
  EXPECT_EQ(groups[0].items[1].formula, "Sheet1!C1:C6");
  EXPECT_EQ(groups[0].items[1].location, "H2");

  // Removing the last group removes the records.
  Workbook cleared = LoadXlsb(SaveXlsb(wb));
  std::string ext = cleared.sheet(0).ext_lst_xml();
  ASSERT_TRUE(static_cast<bool>(remove_sparkline_group(ext, 0)));
  cleared.sheet(0).set_ext_lst_xml(ext);
  EXPECT_TRUE(SparklineContainer(SaveXlsb(cleared)).empty());
}

SparklineGroup AuthoredGroup() {
  SparklineGroup g;
  g.type = SparklineType::kStacked;
  g.color_series.kind = ColorSpec::Kind::kRgb;
  g.color_series.rgb = 0xFF112233U;
  g.color_negative.kind = ColorSpec::Kind::kIndexed;
  g.color_negative.indexed = 10;
  g.negative = true;
  g.display_x_axis = true;
  g.display_hidden = true;
  g.min_axis_type = SparklineAxisType::kGroup;
  g.max_axis_type = SparklineAxisType::kCustom;
  g.has_manual_max = true;
  g.manual_max = 42.0;
  g.date_axis = true;
  g.date_axis_ref = "Data!A1:D1";
  g.line_weight = 2.25;
  g.empty_cells_as = SparklineEmptyCells::kSpan;
  g.items = {{"Data!A2:D2", "F2"}, {"Data!A3:D3", "F3"}};
  return g;
}

void ExpectAuthored(const SparklineGroup& g) {
  EXPECT_EQ(g.type, SparklineType::kStacked);
  EXPECT_EQ(g.color_series.rgb, 0xFF112233U);
  EXPECT_EQ(g.color_negative.kind, ColorSpec::Kind::kIndexed);
  EXPECT_EQ(g.color_negative.indexed, 10U);
  EXPECT_TRUE(g.negative);
  EXPECT_TRUE(g.display_x_axis);
  EXPECT_TRUE(g.display_hidden);
  EXPECT_EQ(g.min_axis_type, SparklineAxisType::kGroup);
  EXPECT_EQ(g.max_axis_type, SparklineAxisType::kCustom);
  EXPECT_DOUBLE_EQ(g.manual_max, 42.0);
  EXPECT_TRUE(g.date_axis);
  EXPECT_EQ(g.date_axis_ref, "Data!A1:D1");
  EXPECT_DOUBLE_EQ(g.line_weight, 2.25);
  EXPECT_EQ(g.empty_cells_as, SparklineEmptyCells::kSpan);
  ASSERT_EQ(g.items.size(), 2U);
  EXPECT_EQ(g.items[0].formula, "Data!A2:D2");
  EXPECT_EQ(g.items[1].location, "F3");
}

Workbook AuthoredBook() {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Data");
  wb.sheet(0).set_cell_value(1U, 0U, Value::number(1.0));
  std::string xml;
  EXPECT_TRUE(static_cast<bool>(add_sparkline_group(xml, AuthoredGroup())));
  wb.sheet(0).set_ext_lst_xml(xml);
  return wb;
}

TEST(XlsbSparklineFidelity, ModelGroupsAreWrittenAndReadBack) {
  const Workbook back = LoadXlsb(SaveXlsb(AuthoredBook()));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(back.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  ExpectAuthored(groups[0]);
  // xlsb -> xlsb keeps them as well.
  const std::vector<SparklineGroup> again = read_sparkline_groups(LoadXlsb(SaveXlsb(back)).sheet(0).ext_lst_xml());
  ASSERT_EQ(again.size(), 1U);
  ExpectAuthored(again[0]);
}

TEST(XlsbSparklineFidelity, UnencodableGroupIsSkippedAndCounted) {
  Workbook wb = AuthoredBook();
  SparklineGroup bad;
  bad.items = {{"#REF!", "H1"}};
  std::string xml = wb.sheet(0).ext_lst_xml();
  ASSERT_TRUE(static_cast<bool>(add_sparkline_group(xml, bad)));
  wb.sheet(0).set_ext_lst_xml(xml);
  ASSERT_EQ(read_sparkline_groups(xml).size(), 2U);

  auto result = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(result)) << result.error().message;
  EXPECT_EQ(result.value().diagnostics.deferred_feature_count, 1U);
  const std::vector<SparklineGroup> groups =
      read_sparkline_groups(LoadXlsb(result.value().bytes).sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  ExpectAuthored(groups[0]);
}

TEST(XlsbSparklineFidelity, XlsxGroupsSurviveConversionToXlsb) {
  auto xlsx = write_ooxml(AuthoredBook());
  ASSERT_TRUE(static_cast<bool>(xlsx));
  auto loaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  const Workbook back = LoadXlsb(SaveXlsb(loaded.value().workbook));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(back.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  ExpectAuthored(groups[0]);
}

TEST(XlsbSparklineFidelity, XlsbGroupsReachAnXlsxSave) {
  const Workbook wb = LoadXlsb(ReadFixture("union_databar_0_base.xlsb"));
  auto xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(xlsx));
  auto loaded = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(loaded));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(loaded.value().workbook.sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].items.at(0).formula, "Sheet1!E1:E6");
  EXPECT_EQ(groups[0].uid, "{85D01F1B-F726-4C98-8685-D838BDDBA70F}");
}

TEST(XlsbSparklineFidelity, RowInsertMovesAndCopiesDecodedSparklines) {
  Workbook wb = LoadXlsb(ReadFixture("union_databar_0_base.xlsb"));
  ASSERT_TRUE(static_cast<bool>(wb.insert_rows(0, 1, 1)));
  const std::vector<SparklineGroup> groups = read_sparkline_groups(LoadXlsb(SaveXlsb(wb)).sheet(0).ext_lst_xml());
  ASSERT_EQ(groups.size(), 1U);
  ASSERT_EQ(groups[0].items.size(), 2U);
  EXPECT_EQ(groups[0].items[0].formula, "Sheet1!E1:E7");
  EXPECT_EQ(groups[0].items[0].location, "G1");
  EXPECT_EQ(groups[0].items[1].location, "G2");
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
