// Writer bytes that earlier rested on an unchecked belief about what Excel writes, each now held against an xlsb
// Excel 16.115 saved from the same source (tests/fixtures/excel/xlsb_parity/claims).

#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "io/ooxml_writer.h"
#include "io/xlsb/reader.h"
#include "io/xlsb/record.h"
#include "io/xlsb/writer.h"
#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"
#include "writer_test_helpers.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace io {
namespace xlsb {
namespace {

constexpr std::uint16_t kBrtBeginList = 343;
// rfxList (16), lt, idList, crwHeader, crwTotals, flags, five dxf ids, dwConnID: the fields before stName.
constexpr std::size_t kListNamesAt = 64;

std::vector<std::uint8_t> ReadFile(const std::string& path) {
  std::vector<std::uint8_t> out;
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
    ADD_FAILURE() << "short read on fixture: " << path;
  }
  std::fclose(f);
  return out;
}

std::string Claim(const std::string& name) {
  return std::string(FORMULON_FIXTURES_DIR) + "/excel/xlsb_parity/claims/" + name;
}

std::vector<std::uint8_t> Converted(const std::string& stem) {
  auto loaded = read_ooxml(SpanOf(ReadFile(Claim(stem + "_src.xlsx"))));
  EXPECT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  if (!loaded) {
    return {};
  }
  auto written = write_xlsb_with_result(loaded.value().workbook);
  EXPECT_TRUE(static_cast<bool>(written)) << written.error().message;
  return written ? std::move(written.value().bytes) : std::vector<std::uint8_t>{};
}

std::vector<std::uint8_t> Part(const std::vector<std::uint8_t>& package, const std::string& path) {
  ZipReader zip;
  EXPECT_TRUE(static_cast<bool>(zip.open(SpanOf(package))));
  auto bytes = zip.read_entry(path);
  EXPECT_TRUE(static_cast<bool>(bytes)) << path;
  return bytes ? std::move(bytes.value()) : std::vector<std::uint8_t>{};
}

/// BrtBeginList from stName through stDisplayName, which follows it.
std::vector<std::uint8_t> ListNames(const std::vector<std::uint8_t>& part) {
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == kBrtBeginList && rec.payload.size > kListNamesAt) {
      return std::vector<std::uint8_t>(rec.payload.data + kListNamesAt, rec.payload.data + rec.payload.size);
    }
  }
  ADD_FAILURE() << "no BrtBeginList";
  return {};
}

// Excel writes stName in BrtBeginList whether or not the table's name differs from its display name.
TEST(XlsbParityClaims, TableNameIsWrittenLikeExcel) {
  const std::vector<std::uint8_t> ours = Converted("stname_diff");
  const std::vector<std::uint8_t> excel = ReadFile(Claim("stname_diff.xlsb"));
  for (const char* part : {"xl/tables/table1.bin", "xl/tables/table2.bin"}) {
    SCOPED_TRACE(part);
    EXPECT_EQ(ListNames(Part(ours, part)), ListNames(Part(excel, part)));
  }
}

// The legacy conditional-format rule records (cbFmla included: cce plus 2 per reference token, absolute or relative)
// equal Excel's.
TEST(XlsbParityClaims, LegacyCfRuleRecordsMatchExcel) {
  constexpr std::uint16_t kBrtBeginCFRule = 463;
  const auto rules = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::vector<std::uint8_t>> out;
    for (const FramedRecord& rec : split_records(part)) {
      if (rec.type == kBrtBeginCFRule) {
        out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
    return out;
  };
  const auto ours = rules(Part(Converted("cbfmla_absolute"), "xl/worksheets/sheet1.bin"));
  const auto excel = rules(Part(ReadFile(Claim("cbfmla_absolute.xlsb")), "xl/worksheets/sheet1.bin"));
  ASSERT_FALSE(excel.empty());
  EXPECT_EQ(ours, excel);
}

// A legacy cell formula whose `@` stands where Excel's own intersection does is written without it, so the workbook
// carries no hidden SINGLE name for it; Excel's names part has the defined name alone.
TEST(XlsbParityClaims, ImpliedIntersectionDeclaresNoHiddenName) {
  constexpr std::uint16_t kBrtName = 39;
  const auto names = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::vector<std::uint8_t>> out;
    for (const FramedRecord& rec : split_records(part)) {
      if (rec.type == kBrtName) {
        out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
    return out;
  };
  const auto ours = names(Part(Converted("name_scope_zero"), "xl/workbook.bin"));
  const auto excel = names(Part(ReadFile(Claim("name_scope_zero.xlsb")), "xl/workbook.bin"));
  ASSERT_EQ(excel.size(), 1U);
  EXPECT_EQ(ours, excel);
}

std::size_t RuleCount(const Workbook& wb) {
  std::size_t n = 0;
  for (const auto& format : wb.sheet(0).conditional_formats()) {
    n += format.rules.size();
  }
  return n;
}

// An xlsb may carry a legacy companion block beside an x14 rule, depending on how the book was produced (a fresh
// Excel save writes none). It loads as one rule, and stays one rule through an xlsb and an xlsx save.
TEST(XlsbParityClaims, LegacyCompanionLoadsAsOneRule) {
  const std::vector<std::uint8_t> bytes =
      ReadFile(std::string(FORMULON_FIXTURES_DIR) + "/excel/cf_custom_icon_set.xlsb");
  auto loaded = read_xlsb(SpanOf(bytes));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  const Workbook& wb = loaded.value().workbook;
  ASSERT_EQ(RuleCount(wb), 1U);
  EXPECT_EQ(wb.sheet(0).conditional_formats()[0].rules[0].priority, 1);

  auto as_xlsb = write_xlsb_with_result(wb);
  ASSERT_TRUE(static_cast<bool>(as_xlsb));
  auto back_xlsb = read_xlsb(SpanOf(as_xlsb.value().bytes));
  ASSERT_TRUE(static_cast<bool>(back_xlsb));
  EXPECT_EQ(RuleCount(back_xlsb.value().workbook), 1U);

  auto as_xlsx = write_ooxml(wb);
  ASSERT_TRUE(static_cast<bool>(as_xlsx));
  auto back_xlsx = read_ooxml(SpanOf(as_xlsx.value()));
  ASSERT_TRUE(static_cast<bool>(back_xlsx));
  EXPECT_EQ(RuleCount(back_xlsx.value().workbook), 1U);
}

// Conditional-format records of a data bar, a two-colour scale, icon sets with showValue=0 and with reverse, and a
// custom icon set, converted from xlsx, equal Excel's fresh xlsb save of the same source.
TEST(XlsbParityClaims, CfKindsMatchExcelsFreshSave) {
  const auto cf_records = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::pair<std::uint16_t, std::vector<std::uint8_t>>> out;
    for (const FramedRecord& rec : split_records(part)) {
      const bool legacy = rec.type >= 460U && rec.type <= 471U;
      const bool x14 = (rec.type >= 1046U && rec.type <= 1052U) || rec.type == 1112U || rec.type == 1135U ||
                       rec.type == 1136U || rec.type == 1155U;
      if (legacy || x14) {
        out.emplace_back(rec.type, std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size));
      }
    }
    return out;
  };
  for (const char* kind : {"databar", "colorscale", "iconset_novalue", "iconset_reverse", "custom_iconset"}) {
    SCOPED_TRACE(kind);
    const std::string stem = std::string("../cfcompanion/") + kind;
    const auto ours = cf_records(Part(Converted(stem), "xl/worksheets/sheet1.bin"));
    const auto excel = cf_records(Part(ReadFile(Claim(stem + ".xlsb")), "xl/worksheets/sheet1.bin"));
    ASSERT_FALSE(excel.empty());
    EXPECT_EQ(ours, excel);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
