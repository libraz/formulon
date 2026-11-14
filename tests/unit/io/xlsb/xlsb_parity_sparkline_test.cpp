// Sparkline groups the xlsb writer writes from the model, held against Excel 16.115's own xlsb and xlsx saves of the
// same sources (tests/fixtures/excel/xlsb_parity/sparklines): absent colours, a deleted source, sources on another
// sheet, a 3-D span, another workbook and a name, and a multi-cell location.

#include <cstddef>
#include <cstdint>
#include <map>
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
#include "utils/error.h"
#include "workbook.h"
#include "writer_test_helpers.h"
#include "zip_repack_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using Bytes = std::vector<std::uint8_t>;

constexpr std::uint16_t kBrtFrtBegin = 35;
constexpr std::uint16_t kBrtFrtEnd = 36;
constexpr std::uint16_t kBrtBeginSparklineGroup = 1041;
constexpr std::uint16_t kBrtSparkline = 1043;
constexpr std::uint16_t kBrtBeginSparklineGroups = 1058;
constexpr const char* kSheet = "xl/worksheets/sheet1.bin";

// The cases Excel opened and saved.
const char* const kOpened[] = {"no_color",      "ref_error",       "other_sheet",       "name_source",
                               "threed_source", "external_source", "multicell_location"};
// The cases Excel refused to open.
const char* const kRefused[] = {"ref_error_rows", "union_source", "structured_source", "expr_source"};

std::string Fixture(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/sparklines/" + name;
}

Bytes Part(const Bytes& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : Bytes{};
}

std::string Hex(const std::uint8_t* data, std::size_t size) {
  static const char kHex[] = "0123456789abcdef";
  std::string out;
  for (std::size_t i = 0; i < size; ++i) {
    out.push_back(kHex[data[i] >> 4U]);
    out.push_back(kHex[data[i] & 0x0FU]);
  }
  return out;
}

/// The SPARKLINEGROUPS container of the first sheet, BrtFRTBegin through BrtFRTEnd, as `type:hex` lines. With
/// `group_records_only`, only BrtBeginSparklineGroup and BrtSparkline.
std::vector<std::string> ContainerLines(const Bytes& package, bool group_records_only = false) {
  const Bytes part = Part(package, kSheet);
  const std::vector<FramedRecord> recs = split_records(part);
  std::vector<std::string> out;
  for (std::size_t i = 0; i + 1U < recs.size(); ++i) {
    if (recs[i].type != kBrtFrtBegin || recs[i + 1U].type != kBrtBeginSparklineGroups) {
      continue;
    }
    for (std::size_t k = i; k < recs.size(); ++k) {
      const bool group_record = recs[k].type == kBrtBeginSparklineGroup || recs[k].type == kBrtSparkline;
      if (!group_records_only || group_record) {
        out.push_back(std::to_string(recs[k].type) + ":" + Hex(recs[k].payload.data, recs[k].payload.size));
      }
      if (recs[k].type == kBrtFrtEnd) {
        break;
      }
    }
    break;
  }
  return out;
}

Workbook LoadXlsx(const std::string& file) {
  auto loaded = read_ooxml(SpanOf(ReadFixtureFile(Fixture(file))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << file << ": " << (loaded ? "" : loaded.error().message);
  return loaded ? std::move(loaded.value().workbook) : Workbook::create_empty();
}

Bytes SaveXlsb(const Workbook& wb) {
  auto saved = write_xlsb_with_result(wb);
  EXPECT_TRUE(static_cast<bool>(saved)) << (saved ? "" : saved.error().message);
  return saved ? std::move(saved.value().bytes) : Bytes{};
}

std::vector<SparklineGroup> Groups(const Workbook& wb) {
  return read_sparkline_groups(wb.sheet(0).ext_lst_xml());
}

void ExpectSameColor(const ColorSpec& a, const ColorSpec& b) {
  EXPECT_EQ(a.kind, b.kind);
  if (a.kind == b.kind && a.kind == ColorSpec::Kind::kRgb) {
    EXPECT_EQ(a.rgb, b.rgb);
  }
}

void ExpectSameGroups(const std::vector<SparklineGroup>& ours, const std::vector<SparklineGroup>& excel) {
  ASSERT_EQ(ours.size(), excel.size());
  for (std::size_t i = 0; i < ours.size(); ++i) {
    const SparklineGroup& a = ours[i];
    const SparklineGroup& b = excel[i];
    EXPECT_EQ(a.type, b.type);
    EXPECT_EQ(a.empty_cells_as, b.empty_cells_as);
    EXPECT_EQ(a.min_axis_type, b.min_axis_type);
    EXPECT_EQ(a.max_axis_type, b.max_axis_type);
    EXPECT_DOUBLE_EQ(a.line_weight, b.line_weight);
    EXPECT_EQ(a.date_axis_ref, b.date_axis_ref);
    EXPECT_EQ(a.uid, b.uid);
    ExpectSameColor(a.color_series, b.color_series);
    ExpectSameColor(a.color_negative, b.color_negative);
    ExpectSameColor(a.color_axis, b.color_axis);
    ExpectSameColor(a.color_markers, b.color_markers);
    ExpectSameColor(a.color_first, b.color_first);
    ExpectSameColor(a.color_last, b.color_last);
    ExpectSameColor(a.color_high, b.color_high);
    ExpectSameColor(a.color_low, b.color_low);
    ASSERT_EQ(a.items.size(), b.items.size());
    for (std::size_t k = 0; k < a.items.size(); ++k) {
      EXPECT_EQ(a.items[k].formula, b.items[k].formula);
      EXPECT_EQ(a.items[k].location, b.items[k].location);
    }
  }
}

TEST(XlsbParitySparklineTwins, ExcelSavesAgree) {
  for (const char* c : kOpened) {
    SCOPED_TRACE(c);
    const std::vector<std::string> first = ContainerLines(ReadFixtureFile(Fixture(std::string(c) + ".xlsb")));
    ASSERT_FALSE(first.empty());
    EXPECT_EQ(ContainerLines(ReadFixtureFile(Fixture(std::string(c) + ".2.xlsb"))), first);
  }
}

TEST(XlsbParitySparklineTwins, ConvertedMatchesExcel) {
  // The authored sources carry no group uid, which Excel adds on save, so the group records are compared.
  for (const char* c : kOpened) {
    SCOPED_TRACE(c);
    const Bytes ours = SaveXlsb(LoadXlsx(std::string(c) + "_src.xlsx"));
    EXPECT_EQ(ContainerLines(ours, true), ContainerLines(ReadFixtureFile(Fixture(std::string(c) + ".xlsb")), true));
  }
}

TEST(XlsbParitySparklineTwins, ExcelXlsxConvertsToExcelXlsb) {
  for (const char* c : kOpened) {
    SCOPED_TRACE(c);
    const Bytes ours = SaveXlsb(LoadXlsx(std::string(c) + ".xlsx"));
    EXPECT_EQ(ContainerLines(ours), ContainerLines(ReadFixtureFile(Fixture(std::string(c) + ".xlsb"))));
  }
}

TEST(XlsbParitySparklineTwins, ReadMatchesXlsxTwin) {
  for (const char* c : kOpened) {
    SCOPED_TRACE(c);
    auto read = read_xlsb(SpanOf(ReadFixtureFile(Fixture(std::string(c) + ".xlsb"))));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    EXPECT_EQ(read.value().skipped_feature_count, 0U);
    ExpectSameGroups(Groups(read.value().workbook), Groups(LoadXlsx(std::string(c) + ".xlsx")));
  }
}

TEST(XlsbParitySparklineTwins, ResaveMatchesExcel) {
  for (const char* c : kOpened) {
    SCOPED_TRACE(c);
    const Bytes excel = ReadFixtureFile(Fixture(std::string(c) + ".xlsb"));
    auto read = read_xlsb(SpanOf(excel));
    ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
    EXPECT_EQ(ContainerLines(SaveXlsb(read.value().workbook)), ContainerLines(excel));
  }
}

TEST(XlsbParitySparklineTwins, RefError) {
  // A deleted source is stored without a formula, in the xlsx as in the xlsb.
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");
  SparklineGroup g;
  g.items = {{"#REF!", "G1"}};
  ASSERT_TRUE(static_cast<bool>(add_sheet_sparkline_group(wb, 0, g)));
  EXPECT_EQ(wb.sheet(0).ext_lst_xml().find("xm:f"), std::string::npos) << wb.sheet(0).ext_lst_xml();
  ASSERT_EQ(Groups(wb).at(0).items.at(0).formula, "");

  const Bytes excel = ReadFixtureFile(Fixture("ref_error.xlsb"));
  std::vector<std::string> theirs;
  for (const std::string& line : ContainerLines(excel, true)) {
    if (line.rfind(std::to_string(kBrtSparkline) + ":", 0) == 0) {
      theirs.push_back(line);
    }
  }
  const Bytes ours = SaveXlsb(wb);
  std::vector<std::string> mine;
  for (const std::string& line : ContainerLines(ours, true)) {
    if (line.rfind(std::to_string(kBrtSparkline) + ":", 0) == 0) {
      mine.push_back(line);
    }
  }
  EXPECT_EQ(mine, theirs);
  auto back = read_xlsb(SpanOf(ours));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  const std::vector<SparklineGroup> groups = Groups(back.value().workbook);
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].items.at(0).formula, "");
  EXPECT_EQ(groups[0].items.at(0).location, "G1");
}

TEST(XlsbParitySparklineTwins, AbsentColour) {
  // Excel's absent-colour slots read back absent and stay absent through an xlsx save.
  auto read = read_xlsb(SpanOf(ReadFixtureFile(Fixture("no_color.xlsb"))));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  auto xlsx = write_ooxml(read.value().workbook);
  ASSERT_TRUE(static_cast<bool>(xlsx)) << xlsx.error().message;
  auto back = read_ooxml(SpanOf(xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(back)) << back.error().message;
  const std::string& xml = back.value().workbook.sheet(0).ext_lst_xml();
  EXPECT_EQ(xml.find("colorSeries"), std::string::npos) << xml;
  const std::vector<SparklineGroup> groups = Groups(back.value().workbook);
  ASSERT_EQ(groups.size(), 1U);
  EXPECT_EQ(groups[0].color_series.kind, ColorSpec::Kind::kNone);
  EXPECT_EQ(groups[0].color_low.kind, ColorSpec::Kind::kNone);
}

TEST(XlsbParitySparkline, RefusedSourceIsNotWritable) {
  // Excel refuses these sources at open; the xlsx reader keeps them, and an xlsb save refuses the state.
  for (const char* c : kRefused) {
    SCOPED_TRACE(c);
    const Workbook wb = LoadXlsx(std::string(c) + "_src.xlsx");
    ASSERT_EQ(Groups(wb).size(), 1U);
    auto saved = write_xlsb_with_result(wb);
    ASSERT_FALSE(static_cast<bool>(saved));
    EXPECT_EQ(saved.error().code, FormulonErrorCode::kIoFeatureNotWritable);
    EXPECT_EQ(saved.error().context, "context=write_xlsb kind=sparklines sheet_index=0");
  }
}

TEST(XlsbParitySparkline, UndecodableGroupIsSkippedAndCounted) {
  Bytes excel = ReadFixtureFile(Fixture("no_color.xlsb"));
  Bytes sheet = Part(excel, kSheet);
  bool patched = false;
  for (const FramedRecord& rec : split_records(sheet)) {
    if (rec.type == kBrtSparkline) {
      // An FRTHeader flag outside the measured set.
      sheet[static_cast<std::size_t>(rec.payload.data - sheet.data())] = 0x07U;
      patched = true;
    }
  }
  ASSERT_TRUE(patched);
  auto read = read_xlsb(SpanOf(RepackWithEntries(excel, {{kSheet, sheet}})));
  ASSERT_TRUE(static_cast<bool>(read)) << read.error().message;
  EXPECT_EQ(read.value().skipped_feature_count, 1U);
  EXPECT_TRUE(Groups(read.value().workbook).empty());
  EXPECT_TRUE(ContainerLines(SaveXlsb(read.value().workbook)).empty());
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
