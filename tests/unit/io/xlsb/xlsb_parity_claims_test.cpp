// Writer bytes that earlier rested on an unchecked belief about what Excel writes, each now held against an xlsb
// Excel 16.115 saved from the same source (tests/fixtures/excel/xlsb_parity/claims and sibling directories).

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

// Conditional-format records of a data bar, a two-colour scale, icon sets with showValue=0 and with reverse, a custom
// icon set, and an x14 icon set beside a colour scale, converted from xlsx, equal Excel's fresh xlsb save of the same
// source.
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
  for (const char* stem_text : {"../cfcompanion/databar", "../cfcompanion/colorscale", "../cfcompanion/iconset_novalue",
                                "../cfcompanion/iconset_reverse", "../cfcompanion/custom_iconset", "icon_set_x14"}) {
    SCOPED_TRACE(stem_text);
    const std::string stem = stem_text;
    const auto ours = cf_records(Part(Converted(stem), "xl/worksheets/sheet1.bin"));
    const auto excel = cf_records(Part(ReadFile(Claim(stem + ".xlsb")), "xl/worksheets/sheet1.bin"));
    ASSERT_FALSE(excel.empty());
    EXPECT_EQ(ours, excel);
  }
}

// The ExternSheet table equals Excel's: defined names first, then each sheet in order with its cell formulas'
// sheet-qualified references, its conditional-format rules' and validations' references (an unqualified one names the
// sheet itself), and the sheet's own range where a rule or validation reads it.
TEST(XlsbParityClaims, ExternSheetMatchesExcel) {
  constexpr std::uint16_t kBrtExternSheet = 362;
  const auto table = [](const std::vector<std::uint8_t>& part) {
    for (const FramedRecord& rec : split_records(part)) {
      if (rec.type == kBrtExternSheet) {
        return std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size);
      }
    }
    return std::vector<std::uint8_t>{};
  };
  for (const char* stem : {"above_average",
                           "cell_and_x14",
                           "cell_only",
                           "cell_self_and_x14",
                           "cell_then_name",
                           "cellis_const",
                           "cellis_ref",
                           "cellis_sheet_with_cell",
                           "cf_cross_only",
                           "colorscale_cfvo_cross",
                           "colorscale_num",
                           "databar_cfvo_cross",
                           "databar_first_cell_last",
                           "databar_num",
                           "databar_then_later_cell",
                           "databar_two_sheets",
                           "duplicates",
                           "dv_and_cf",
                           "dv_const_between",
                           "dv_const_list",
                           "dv_cross_sheet",
                           "dv_same_sheet",
                           "dv_then_databar_next_sheet",
                           "dv_third_to_second",
                           "dv_whole_cross",
                           "expr_const",
                           "iconset_cfvo_cross",
                           "iconset_cfvo_self",
                           "iconset_formula_consts",
                           "iconset_num",
                           "iconset_percent_plain",
                           "name_only",
                           "top10",
                           "x14_on_second_sheet",
                           "x14_same_sheet",
                           "x14_third_to_second"}) {
    SCOPED_TRACE(stem);
    const std::string path = std::string("../externsheet/") + stem;
    EXPECT_EQ(table(Part(Converted(path), "xl/workbook.bin")),
              table(Part(ReadFile(Claim(path + ".xlsb")), "xl/workbook.bin")));
  }
  // Cell formulas naming the third, first and second sheets in row-major order.
  EXPECT_EQ(table(Part(Converted("externsheet_order"), "xl/workbook.bin")),
            table(Part(ReadFile(Claim("externsheet_order.xlsb")), "xl/workbook.bin")));
}

/// The payloads of `type` records in `part`.
std::vector<std::vector<std::uint8_t>> RecordsOf(const std::vector<std::uint8_t>& part, std::uint16_t type) {
  std::vector<std::vector<std::uint8_t>> out;
  for (const FramedRecord& rec : split_records(part)) {
    if (rec.type == type) {
      out.emplace_back(rec.payload.data, rec.payload.data + rec.payload.size);
    }
  }
  return out;
}

std::uint32_t U32At(const std::vector<std::uint8_t>& b, std::size_t at) {
  return static_cast<std::uint32_t>(b[at]) | (static_cast<std::uint32_t>(b[at + 1U]) << 8U) |
         (static_cast<std::uint32_t>(b[at + 2U]) << 16U) | (static_cast<std::uint32_t>(b[at + 3U]) << 24U);
}

// Formulas Excel evaluates as arrays (tests/fixtures/excel/xlsb_parity/arraymode): array (CSE) formulas, a defined
// name's body carrying `@`, and conditional-format formulas. Inside an array formula an argument under a parameter
// other than a value one is evaluated as an array; the formula's root is not.
constexpr std::uint16_t kBrtArrFmla = 426;
// BrtArrFmla: rfx (16 bytes) and flags (1) precede the formula.
constexpr std::size_t kArrFmlaFormulaAt = 17;
constexpr std::uint8_t kPtgName = 0x23;

constexpr std::uint16_t kNameRecord = 39;
constexpr std::uint16_t kCfRuleRecord = 463;

TEST(XlsbParityClaims, ArrayFormulasMatchExcel) {
  for (const char* stem :
       {"cse_abs_area", "cse_abs_at",    "cse_area_op", "cse_at_cell",    "cse_at_op_ctrl", "cse_at_root",
        "cse_sum_at",   "g_abs_op_root", "g_abs_root",  "g_area_op_root", "g_area_root",    "g_cell_op_sum",
        "g_if_root",    "g_index_op",    "g_len_root",  "g_max_abs",      "g_rows_op",      "g_sum_abs_cell",
        "g_sum_area",   "g_sum_if",      "g_sumproduct"}) {
    SCOPED_TRACE(stem);
    const auto ours =
        RecordsOf(Part(Converted(std::string("../arraymode/") + stem), "xl/worksheets/sheet1.bin"), kBrtArrFmla);
    const auto excel = RecordsOf(
        Part(ReadFile(Claim(std::string("../arraymode/") + stem + ".xlsb")), "xl/worksheets/sheet1.bin"), kBrtArrFmla);
    ASSERT_EQ(excel.size(), 1U);
    ASSERT_EQ(ours.size(), 1U);
    EXPECT_EQ(std::vector<std::uint8_t>(ours[0].begin() + kArrFmlaFormulaAt, ours[0].end()),
              std::vector<std::uint8_t>(excel[0].begin() + kArrFmlaFormulaAt, excel[0].end()));
  }
}

// Bodies of names carrying `@` as an operand, a reference argument and a value argument, read from Excel's xlsb and
// written back. The leading PtgName's index into the name table is masked: the two tables order the hidden SINGLE
// name differently.
TEST(XlsbParityClaims, NameBodiesWithAtMatchExcel) {
  const auto bodies = [](const std::vector<std::uint8_t>& part) {
    std::vector<std::vector<std::uint8_t>> out;
    for (std::vector<std::uint8_t> payload : RecordsOf(part, kNameRecord)) {
      // grbit (4), chKey (1), itab (4), name (XLWideString), then cce and rgce.
      const std::size_t cce_at = 4U + 1U + 4U + 4U + 2U * U32At(payload, 9U);
      const std::size_t rgce_at = cce_at + 4U;
      std::vector<std::uint8_t> rgce(payload.begin() + static_cast<std::ptrdiff_t>(rgce_at),
                                     payload.begin() + static_cast<std::ptrdiff_t>(rgce_at + U32At(payload, cce_at)));
      if (rgce.size() > 5U && rgce[0] == kPtgName) {
        std::fill(rgce.begin() + 1, rgce.begin() + 5, 0U);
        out.push_back(std::move(rgce));
      }
    }
    return out;
  };
  const std::vector<std::uint8_t> excel = ReadFile(Claim("../arraymode/name_at.xlsb"));
  auto loaded = read_xlsb(SpanOf(excel));
  ASSERT_TRUE(static_cast<bool>(loaded)) << loaded.error().message;
  auto written = write_xlsb_with_result(loaded.value().workbook);
  ASSERT_TRUE(static_cast<bool>(written));
  const auto theirs = bodies(Part(excel, "xl/workbook.bin"));
  ASSERT_EQ(theirs.size(), 4U);
  EXPECT_EQ(bodies(Part(written.value().bytes, "xl/workbook.bin")), theirs);
}

TEST(XlsbParityClaims, ConditionalFormatFormulasMatchExcel) {
  for (const char* stem : {"cf_area_op", "cf_root_op"}) {
    SCOPED_TRACE(stem);
    const auto excel =
        RecordsOf(Part(ReadFile(Claim(std::string("../arraymode/") + stem + ".xlsb")), "xl/worksheets/sheet1.bin"),
                  kCfRuleRecord);
    ASSERT_EQ(excel.size(), 1U);
    EXPECT_EQ(
        RecordsOf(Part(Converted(std::string("../arraymode/") + stem), "xl/worksheets/sheet1.bin"), kCfRuleRecord),
        excel);
  }
}

// A validation formula is a legacy one: an area under an operator inside a function is intersected, so Excel stores it
// value class.
TEST(XlsbParityClaims, ValidationFormulasMatchExcel) {
  constexpr std::uint16_t kDValRecord = 64;
  for (const char* stem : {"dv_area_op", "dv_root_op"}) {
    SCOPED_TRACE(stem);
    const std::string path = std::string("../arraymode/") + stem;
    const auto excel = RecordsOf(Part(ReadFile(Claim(path + ".xlsb")), "xl/worksheets/sheet1.bin"), kDValRecord);
    ASSERT_EQ(excel.size(), 1U);
    EXPECT_EQ(RecordsOf(Part(Converted(path), "xl/worksheets/sheet1.bin"), kDValRecord), excel);
  }
}

}  // namespace
}  // namespace xlsb
}  // namespace io
}  // namespace formulon
