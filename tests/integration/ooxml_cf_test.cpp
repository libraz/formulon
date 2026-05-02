// Copyright 2026 libraz. Licensed under the MIT License.
//
// Integration test for the OOXML conditional-formatting wiring.
// Constructs a minimal in-memory `.xlsx` package whose Sheet1 carries
// two `<conditionalFormatting>` blocks (a `cellIs` rule and a 3-stop
// `colorScale` rule); drives the bytes through `read_ooxml`; asserts
// the parsed CF blocks land on the loaded `Sheet`.

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "miniz.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace {

io::ByteSpan SpanOf(const std::vector<std::uint8_t>& bytes) {
  return io::ByteSpan{bytes.data(), bytes.size()};
}

struct PartFile {
  const char* path;
  std::string_view body;
};

std::vector<std::uint8_t> BuildZip(const std::vector<PartFile>& parts) {
  mz_zip_archive writer{};
  EXPECT_NE(mz_zip_writer_init_heap(&writer, 0, 4096), MZ_FALSE);
  for (const auto& p : parts) {
    EXPECT_NE(mz_zip_writer_add_mem(&writer, p.path, p.body.data(), p.body.size(),
                                    static_cast<mz_uint>(MZ_DEFAULT_COMPRESSION)),
              MZ_FALSE)
        << "miniz add failed for " << p.path;
  }
  void* archive_ptr = nullptr;
  std::size_t archive_size = 0;
  EXPECT_NE(mz_zip_writer_finalize_heap_archive(&writer, &archive_ptr, &archive_size), MZ_FALSE);
  EXPECT_NE(mz_zip_writer_end(&writer), MZ_FALSE);
  std::vector<std::uint8_t> out(static_cast<const std::uint8_t*>(archive_ptr),
                                static_cast<const std::uint8_t*>(archive_ptr) + archive_size);
  mz_free(archive_ptr);
  return out;
}

TEST(OoxmlCF, PackageWithConditionalFormattingLoads) {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "</Types>\n";

  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";

  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";

  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";

  // Sheet1 carries two CF blocks: one cellIs > 50 over A1:A10, and one
  // 3-stop colorScale over B1:B10. The order on the wire (sqref + rule
  // attributes) is what the reader must round-trip into the Sheet.
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "  <conditionalFormatting sqref=\"A1:A10\">\n"
      "    <cfRule type=\"cellIs\" priority=\"1\" operator=\"greaterThan\" dxfId=\"0\" stopIfTrue=\"1\">\n"
      "      <formula>50</formula>\n"
      "    </cfRule>\n"
      "  </conditionalFormatting>\n"
      "  <conditionalFormatting sqref=\"B1:B10\">\n"
      "    <cfRule type=\"colorScale\" priority=\"2\">\n"
      "      <colorScale>\n"
      "        <cfvo type=\"min\"/>\n"
      "        <cfvo type=\"percentile\" val=\"50\"/>\n"
      "        <cfvo type=\"max\"/>\n"
      "        <color rgb=\"FFFF0000\"/>\n"
      "        <color rgb=\"FFFFFF00\"/>\n"
      "        <color rgb=\"FF00FF00\"/>\n"
      "      </colorScale>\n"
      "    </cfRule>\n"
      "  </conditionalFormatting>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  const Workbook& wb = result_or.value().workbook;
  ASSERT_GT(wb.sheet_count(), 0U);
  const Sheet& s1 = wb.sheet(0);
  const auto& cfs = s1.conditional_formats();
  ASSERT_EQ(cfs.size(), 2U);

  // Block 0: cellIs > 50 over A1:A10.
  const auto& b0 = cfs[0];
  ASSERT_EQ(b0.sqref.size(), 1U);
  EXPECT_EQ(b0.sqref[0].first.row, 0U);
  EXPECT_EQ(b0.sqref[0].first.col, 0U);
  EXPECT_EQ(b0.sqref[0].last.row, 9U);
  EXPECT_EQ(b0.sqref[0].last.col, 0U);
  ASSERT_EQ(b0.rules.size(), 1U);
  EXPECT_EQ(b0.rules[0].type, cf::RuleType::CellIs);
  EXPECT_EQ(b0.rules[0].priority, 1);
  EXPECT_TRUE(b0.rules[0].stop_if_true);
  ASSERT_TRUE(b0.rules[0].dxf_id.has_value());
  EXPECT_EQ(b0.rules[0].dxf_id.value(), 0U);
  ASSERT_TRUE(b0.rules[0].op.has_value());
  EXPECT_EQ(b0.rules[0].op.value(), cf::CellIsOperator::GreaterThan);
  ASSERT_TRUE(b0.rules[0].formula1.has_value());
  EXPECT_EQ(b0.rules[0].formula1.value(), "50");

  // Block 1: 3-stop colorScale over B1:B10.
  const auto& b1 = cfs[1];
  ASSERT_EQ(b1.sqref.size(), 1U);
  EXPECT_EQ(b1.sqref[0].first.col, 1U);
  EXPECT_EQ(b1.sqref[0].last.row, 9U);
  ASSERT_EQ(b1.rules.size(), 1U);
  EXPECT_EQ(b1.rules[0].type, cf::RuleType::ColorScale);
  EXPECT_EQ(b1.rules[0].priority, 2);
  ASSERT_TRUE(b1.rules[0].color_scale.has_value());
  EXPECT_EQ(b1.rules[0].color_scale->thresholds.size(), 3U);
  EXPECT_EQ(b1.rules[0].color_scale->thresholds[0].type, cf::CfvoType::Min);
  EXPECT_EQ(b1.rules[0].color_scale->thresholds[1].type, cf::CfvoType::Percentile);
  EXPECT_EQ(b1.rules[0].color_scale->thresholds[1].value, "50");
  EXPECT_EQ(b1.rules[0].color_scale->colors.size(), 3U);
  EXPECT_EQ(b1.rules[0].color_scale->colors[0], cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(b1.rules[0].color_scale->colors[1], cf::Color({255, 255, 0, 255}));
  EXPECT_EQ(b1.rules[0].color_scale->colors[2], cf::Color({0, 255, 0, 255}));
}

TEST(OoxmlCF, RoundTripThroughWriterPreservesConditionalFormats) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block1{};
  block1.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r1;
  r1.type = cf::RuleType::CellIs;
  r1.priority = 1;
  r1.stop_if_true = true;
  r1.dxf_id = 0u;
  r1.op = cf::CellIsOperator::GreaterThan;
  r1.formula1 = "50";
  block1.rules.push_back(std::move(r1));

  cf::ConditionalFormat block2{};
  block2.sqref.push_back({{0, 1}, {9, 1}});
  cf::CFRule r2;
  r2.type = cf::RuleType::ColorScale;
  cf::ColorScaleSpec spec;
  spec.thresholds.push_back({cf::CfvoType::Min, "", true});
  spec.thresholds.push_back({cf::CfvoType::Percentile, "50", true});
  spec.thresholds.push_back({cf::CfvoType::Max, "", true});
  spec.colors.push_back({255, 0, 0, 255});
  spec.colors.push_back({255, 255, 0, 255});
  spec.colors.push_back({0, 255, 0, 255});
  r2.color_scale = std::move(spec);
  r2.priority = 2;
  block2.rules.push_back(std::move(r2));

  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block1));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block2));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const Workbook& reloaded = read_or.value().workbook;
  ASSERT_EQ(reloaded.sheet_count(), 1U);
  const auto& cfs = reloaded.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 2U);

  ASSERT_EQ(cfs[0].rules.size(), 1U);
  EXPECT_EQ(cfs[0].rules[0].type, cf::RuleType::CellIs);
  EXPECT_EQ(cfs[0].rules[0].priority, 1);
  EXPECT_TRUE(cfs[0].rules[0].stop_if_true);
  ASSERT_TRUE(cfs[0].rules[0].op.has_value());
  EXPECT_EQ(cfs[0].rules[0].op.value(), cf::CellIsOperator::GreaterThan);
  EXPECT_EQ(cfs[0].rules[0].formula1.value(), "50");

  ASSERT_EQ(cfs[1].rules.size(), 1U);
  EXPECT_EQ(cfs[1].rules[0].type, cf::RuleType::ColorScale);
  ASSERT_TRUE(cfs[1].rules[0].color_scale.has_value());
  EXPECT_EQ(cfs[1].rules[0].color_scale->thresholds.size(), 3U);
  EXPECT_EQ(cfs[1].rules[0].color_scale->colors[0], cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(cfs[1].rules[0].color_scale->colors[2], cf::Color({0, 255, 0, 255}));
}

TEST(OoxmlCF, EmptyWorkbookHasNoConditionalFormats) {
  const std::string_view content_types =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
      "  <Default Extension=\"rels\" "
      "ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
      "  <Default Extension=\"xml\" ContentType=\"application/xml\"/>\n"
      "  <Override PartName=\"/xl/workbook.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>\n"
      "  <Override PartName=\"/xl/worksheets/sheet1.xml\" "
      "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>\n"
      "</Types>\n";
  const std::string_view package_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" "
      "Target=\"xl/workbook.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view workbook_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" "
      "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
      "  <sheets>\n"
      "    <sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>\n"
      "  </sheets>\n"
      "</workbook>\n";
  const std::string_view workbook_rels =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
      "  <Relationship Id=\"rId1\" "
      "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" "
      "Target=\"worksheets/sheet1.xml\"/>\n"
      "</Relationships>\n";
  const std::string_view sheet1_xml =
      "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"
      "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"
      "  <sheetData/>\n"
      "</worksheet>\n";

  const std::vector<std::uint8_t> bytes = BuildZip({
      {"[Content_Types].xml", content_types},
      {"_rels/.rels", package_rels},
      {"xl/workbook.xml", workbook_xml},
      {"xl/_rels/workbook.xml.rels", workbook_rels},
      {"xl/worksheets/sheet1.xml", sheet1_xml},
  });

  auto result_or = io::read_ooxml(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(result_or)) << "read_ooxml: " << result_or.error().message;
  EXPECT_TRUE(result_or.value().workbook.sheet(0).conditional_formats().empty());
}

TEST(OoxmlCF, RoundTripPreservesDataBarRule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::DataBar;
  r.priority = 1;
  cf::DataBarSpec bar;
  bar.min = {cf::CfvoType::Min, "", true};
  bar.max = {cf::CfvoType::Max, "", true};
  bar.fill = {255, 0, 0, 255};
  bar.min_length_pct = 20;
  bar.max_length_pct = 80;
  bar.show_value = false;
  r.data_bar = std::move(bar);
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::DataBar);
  EXPECT_EQ(rule.priority, 1);
  ASSERT_TRUE(rule.data_bar.has_value());
  EXPECT_EQ(rule.data_bar->min.type, cf::CfvoType::Min);
  EXPECT_EQ(rule.data_bar->min.value, "");
  EXPECT_TRUE(rule.data_bar->min.gte);
  EXPECT_EQ(rule.data_bar->max.type, cf::CfvoType::Max);
  EXPECT_EQ(rule.data_bar->max.value, "");
  EXPECT_TRUE(rule.data_bar->max.gte);
  EXPECT_EQ(rule.data_bar->fill, cf::Color({255, 0, 0, 255}));
  EXPECT_EQ(rule.data_bar->min_length_pct, 20U);
  EXPECT_EQ(rule.data_bar->max_length_pct, 80U);
  EXPECT_FALSE(rule.data_bar->show_value);
}

TEST(OoxmlCF, RoundTripPreservesIconSetRule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::IconSet;
  r.priority = 1;
  cf::IconSetSpec iset;
  iset.name = cf::IconSetName::Five_Arrows;
  iset.reverse = true;
  iset.show_value = false;
  iset.percent = false;
  iset.thresholds.push_back({cf::CfvoType::Number, "20", true});
  iset.thresholds.push_back({cf::CfvoType::Number, "40", false});
  iset.thresholds.push_back({cf::CfvoType::Percentile, "75", true});
  iset.thresholds.push_back({cf::CfvoType::Number, "85", false});
  r.icon_set = std::move(iset);
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::IconSet);
  ASSERT_TRUE(rule.icon_set.has_value());
  EXPECT_EQ(rule.icon_set->name, cf::IconSetName::Five_Arrows);
  EXPECT_TRUE(rule.icon_set->reverse);
  EXPECT_FALSE(rule.icon_set->show_value);
  EXPECT_FALSE(rule.icon_set->percent);
  ASSERT_EQ(rule.icon_set->thresholds.size(), 4U);
  EXPECT_EQ(rule.icon_set->thresholds[0].type, cf::CfvoType::Number);
  EXPECT_EQ(rule.icon_set->thresholds[0].value, "20");
  EXPECT_TRUE(rule.icon_set->thresholds[0].gte);
  EXPECT_EQ(rule.icon_set->thresholds[1].type, cf::CfvoType::Number);
  EXPECT_EQ(rule.icon_set->thresholds[1].value, "40");
  EXPECT_FALSE(rule.icon_set->thresholds[1].gte);
  EXPECT_EQ(rule.icon_set->thresholds[2].type, cf::CfvoType::Percentile);
  EXPECT_EQ(rule.icon_set->thresholds[2].value, "75");
  EXPECT_TRUE(rule.icon_set->thresholds[2].gte);
  EXPECT_EQ(rule.icon_set->thresholds[3].type, cf::CfvoType::Number);
  EXPECT_EQ(rule.icon_set->thresholds[3].value, "85");
  EXPECT_FALSE(rule.icon_set->thresholds[3].gte);
}

TEST(OoxmlCF, RoundTripPreservesTop10Rule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::Top10;
  r.priority = 1;
  r.rank = 5;
  r.percent = true;
  r.bottom = true;
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::Top10);
  ASSERT_TRUE(rule.rank.has_value());
  EXPECT_EQ(rule.rank.value(), 5);
  EXPECT_TRUE(rule.percent);
  EXPECT_TRUE(rule.bottom);
}

TEST(OoxmlCF, RoundTripPreservesAboveAverageRule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::AboveAverage;
  r.priority = 1;
  r.above_average = false;
  r.equal_average = true;
  r.std_dev = 1.5;
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::AboveAverage);
  EXPECT_FALSE(rule.above_average);
  EXPECT_TRUE(rule.equal_average);
  ASSERT_TRUE(rule.std_dev.has_value());
  EXPECT_EQ(rule.std_dev.value(), 1.5);
}

TEST(OoxmlCF, RoundTripPreservesTextRules) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r1;
  r1.type = cf::RuleType::ContainsText;
  r1.priority = 1;
  r1.text = "ERROR";
  block.rules.push_back(std::move(r1));
  cf::CFRule r2;
  r2.type = cf::RuleType::NotContainsText;
  r2.priority = 2;
  r2.text = "OK";
  block.rules.push_back(std::move(r2));
  cf::CFRule r3;
  r3.type = cf::RuleType::BeginsWith;
  r3.priority = 3;
  r3.text = "A_";
  block.rules.push_back(std::move(r3));
  cf::CFRule r4;
  r4.type = cf::RuleType::EndsWith;
  r4.priority = 4;
  r4.text = "_z";
  block.rules.push_back(std::move(r4));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 4U);
  EXPECT_EQ(cfs[0].rules[0].type, cf::RuleType::ContainsText);
  ASSERT_TRUE(cfs[0].rules[0].text.has_value());
  EXPECT_EQ(cfs[0].rules[0].text.value(), "ERROR");
  EXPECT_EQ(cfs[0].rules[1].type, cf::RuleType::NotContainsText);
  ASSERT_TRUE(cfs[0].rules[1].text.has_value());
  EXPECT_EQ(cfs[0].rules[1].text.value(), "OK");
  EXPECT_EQ(cfs[0].rules[2].type, cf::RuleType::BeginsWith);
  ASSERT_TRUE(cfs[0].rules[2].text.has_value());
  EXPECT_EQ(cfs[0].rules[2].text.value(), "A_");
  EXPECT_EQ(cfs[0].rules[3].type, cf::RuleType::EndsWith);
  ASSERT_TRUE(cfs[0].rules[3].text.has_value());
  EXPECT_EQ(cfs[0].rules[3].text.value(), "_z");
}

TEST(OoxmlCF, RoundTripPreservesBlanksAndErrorsRules) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r1;
  r1.type = cf::RuleType::ContainsBlanks;
  r1.priority = 1;
  block.rules.push_back(std::move(r1));
  cf::CFRule r2;
  r2.type = cf::RuleType::NotContainsBlanks;
  r2.priority = 2;
  block.rules.push_back(std::move(r2));
  cf::CFRule r3;
  r3.type = cf::RuleType::ContainsErrors;
  r3.priority = 3;
  block.rules.push_back(std::move(r3));
  cf::CFRule r4;
  r4.type = cf::RuleType::NotContainsErrors;
  r4.priority = 4;
  block.rules.push_back(std::move(r4));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 4U);
  EXPECT_EQ(cfs[0].rules[0].type, cf::RuleType::ContainsBlanks);
  EXPECT_EQ(cfs[0].rules[1].type, cf::RuleType::NotContainsBlanks);
  EXPECT_EQ(cfs[0].rules[2].type, cf::RuleType::ContainsErrors);
  EXPECT_EQ(cfs[0].rules[3].type, cf::RuleType::NotContainsErrors);
}

TEST(OoxmlCF, RoundTripPreservesTimePeriodRule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::TimePeriod;
  r.priority = 1;
  r.time_period = cf::TimePeriod::Last7Days;
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::TimePeriod);
  ASSERT_TRUE(rule.time_period.has_value());
  EXPECT_EQ(rule.time_period.value(), cf::TimePeriod::Last7Days);
}

TEST(OoxmlCF, RoundTripPreservesDuplicateAndUniqueRules) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r1;
  r1.type = cf::RuleType::DuplicateValues;
  r1.priority = 1;
  block.rules.push_back(std::move(r1));
  cf::CFRule r2;
  r2.type = cf::RuleType::UniqueValues;
  r2.priority = 2;
  block.rules.push_back(std::move(r2));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 2U);
  EXPECT_EQ(cfs[0].rules[0].type, cf::RuleType::DuplicateValues);
  EXPECT_EQ(cfs[0].rules[0].priority, 1);
  EXPECT_EQ(cfs[0].rules[1].type, cf::RuleType::UniqueValues);
  EXPECT_EQ(cfs[0].rules[1].priority, 2);
}

TEST(OoxmlCF, RoundTripPreservesExpressionRule) {
  Workbook wb = Workbook::create_empty();
  wb.add_sheet("Sheet1");

  cf::ConditionalFormat block{};
  block.sqref.push_back({{0, 0}, {9, 0}});
  cf::CFRule r;
  r.type = cf::RuleType::Expression;
  r.priority = 1;
  r.formula1 = "A1>10";
  block.rules.push_back(std::move(r));
  wb.sheet(0).mutable_conditional_formats().push_back(std::move(block));

  auto bytes_or = io::write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(bytes_or)) << "write_ooxml: " << bytes_or.error().message;
  auto read_or = io::read_ooxml(SpanOf(bytes_or.value()));
  ASSERT_TRUE(static_cast<bool>(read_or)) << "read_ooxml: " << read_or.error().message;

  const auto& cfs = read_or.value().workbook.sheet(0).conditional_formats();
  ASSERT_EQ(cfs.size(), 1U);
  ASSERT_EQ(cfs[0].rules.size(), 1U);
  const auto& rule = cfs[0].rules[0];
  EXPECT_EQ(rule.type, cf::RuleType::Expression);
  ASSERT_TRUE(rule.formula1.has_value());
  EXPECT_EQ(rule.formula1.value(), "A1>10");
}

}  // namespace
}  // namespace formulon
