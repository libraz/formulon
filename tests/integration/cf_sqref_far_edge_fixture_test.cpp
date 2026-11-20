//
// A cell delete whose cells lie on a rectangle's far line, short of the whole line, leaves a conditional
// format's rectangle whole unless a rule ranks the range or has a min, max or percent threshold. The books in
// `tests/fixtures/excel/cf_sqref_far_edge/` are authored with one rule kind per sheet; the expected text is
// Excel's save after the same edit.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <iterator>
#include <sstream>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "cf/cf_rule_traits.h"
#include "cf/cf_types.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFixture(const std::string& name) {
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + name;
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

Workbook Load(const std::vector<std::uint8_t>& bytes) {
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  EXPECT_TRUE(static_cast<bool>(read)) << (read ? "" : read.error().message);
  return read ? std::move(read.value().workbook) : Workbook::create_empty();
}

MergeRange Rect(std::string_view a1) {
  const auto cell = [](std::string_view s) {
    std::uint32_t col = 0;
    std::size_t i = 0;
    for (; i < s.size() && s[i] >= 'A' && s[i] <= 'Z'; ++i) {
      col = col * 26U + static_cast<std::uint32_t>(s[i] - 'A' + 1);
    }
    std::uint32_t row = 0;
    for (; i < s.size(); ++i) {
      row = row * 10U + static_cast<std::uint32_t>(s[i] - '0');
    }
    return std::pair<std::uint32_t, std::uint32_t>{row - 1U, col - 1U};
  };
  const std::size_t colon = a1.find(':');
  const auto first = cell(a1.substr(0, colon));
  const auto last = colon == std::string_view::npos ? first : cell(a1.substr(colon + 1U));
  return MergeRange{first.first, first.second, last.first, last.second};
}

std::string CellText(const CellAddress& cell) {
  std::string letters;
  for (std::uint32_t n = cell.col + 1U; n > 0U; n = (n - 1U) / 26U) {
    letters.insert(letters.begin(), static_cast<char>('A' + (n - 1U) % 26U));
  }
  return letters + std::to_string(cell.row + 1U);
}

// Every block's sqref pieces in stored order, joined by " || " across blocks.
std::string SqrefOf(const Sheet& sheet) {
  std::string out;
  for (const cf::ConditionalFormat& block : sheet.conditional_formats()) {
    std::string text;
    for (const cf::CFCellRange& range : block.sqref) {
      text += text.empty() ? "" : " ";
      text += CellText(range.first);
      if (range.first != range.last) {
        text += ":" + CellText(range.last);
      }
    }
    out += out.empty() ? text : " || " + text;
  }
  return out;
}

// The pieces of `text`, ordered, so that two spellings of the same pieces compare equal.
std::vector<std::string> SortedPieces(const std::string& text) {
  std::istringstream in(text);
  std::vector<std::string> pieces{std::istream_iterator<std::string>(in), std::istream_iterator<std::string>()};
  std::sort(pieces.begin(), pieces.end());
  return pieces;
}

struct SheetCase {
  const char* sheet;
  const char* excel;  // Excel's sqref after the edit
};

void CheckEdit(const char* fixture, std::string_view cells, CellShift shift, const std::vector<SheetCase>& cases,
               bool exact_order) {
  Workbook wb = Load(ReadFixture(std::string("cf_sqref_far_edge/") + fixture));
  for (const SheetCase& c : cases) {
    SCOPED_TRACE(c.sheet);
    std::size_t index = wb.sheet_count();
    for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
      if (wb.sheet(i).name() == c.sheet) {
        index = i;
      }
    }
    ASSERT_LT(index, wb.sheet_count());
    ASSERT_TRUE(static_cast<bool>(wb.delete_cells(index, Rect(cells), shift)));
    const std::string engine = SqrefOf(wb.sheet(index));
    if (exact_order) {
      EXPECT_EQ(engine, c.excel);
    } else {
      EXPECT_EQ(SortedPieces(engine), SortedPieces(c.excel));
    }
  }
}

// Sheets whose rule keeps its rectangle whole: absolute thresholds, no thresholds, and the formula kinds.
const std::vector<SheetCase> kRightLineWhole = {
    {"E01", "D4:F8"}, {"E02", "D4:F8"}, {"E03", "D4:F8"}, {"E04", "D4:F8"}, {"E05", "D4:F8"},
    {"E06", "D4:F8"}, {"E07", "D4:F8"}, {"E08", "D4:F8"}, {"E09", "D4:F8"}, {"E12", "D4:F8"},
    {"E13", "D4:F8"}, {"E14", "D4:F8"}, {"E17", "D4:F8"},
};
const std::vector<SheetCase> kRightLineSplit = {
    {"E10", "D4:F5 D7:F8 D6:E6"},
    {"E11", "D4:F5 D7:F8 D6:E6"},
    {"E15", "D4:F5 D7:F8 D6:E6"},
    {"E16", "D4:F5 D7:F8 D6:E6"},
};
const std::vector<SheetCase> kBottomLineWhole = {
    {"F01", "B3:E5"}, {"F02", "B3:E5"}, {"F03", "B3:E5"}, {"F04", "B3:E5"}, {"F05", "B3:E5"},
    {"F06", "B3:E5"}, {"F07", "B3:E5"}, {"F08", "B3:E5"}, {"F09", "B3:E5"}, {"F12", "B3:E5"},
    {"F13", "B3:E5"}, {"F14", "B3:E5"}, {"F17", "B3:E5"},
};
const std::vector<SheetCase> kBottomLineSplit = {
    {"F10", "B3:E4 D5:E5"},
    {"F11", "B3:E4 D5:E5"},
    {"F15", "B3:E4 D5:E5"},
    {"F16", "B3:E4 D5:E5"},
};

TEST(CfSqrefFarEdgeFixture, RightLineMiddleCellDeleteKeepsRectangleWholeForAbsoluteThresholds) {
  CheckEdit("right_line_middle_left_authored.xlsx", "F6", CellShift::kLeft, kRightLineWhole, true);
}

TEST(CfSqrefFarEdgeFixture, RightLineMiddleCellDeleteSplitsRectangleForRelativeThresholds) {
  CheckEdit("right_line_middle_left_authored.xlsx", "F6", CellShift::kLeft, kRightLineSplit, false);
}

TEST(CfSqrefFarEdgeFixture, BottomLineCellsDeleteKeepsRectangleWholeForAbsoluteThresholds) {
  CheckEdit("bottom_line_two_cells_up_authored.xlsx", "B5:C5", CellShift::kUp, kBottomLineWhole, true);
}

TEST(CfSqrefFarEdgeFixture, BottomLineCellsDeleteSplitsRectangleForRelativeThresholds) {
  CheckEdit("bottom_line_two_cells_up_authored.xlsx", "B5:C5", CellShift::kUp, kBottomLineSplit, false);
}

cf::CfValueObject Cfvo(cf::CfvoType type) {
  return cf::CfValueObject{type, "10", true};
}

cf::CFRule ColorScaleOf(std::vector<cf::CfValueObject> thresholds) {
  cf::CFRule rule;
  rule.type = cf::RuleType::ColorScale;
  rule.color_scale = cf::ColorScaleSpec{std::move(thresholds), {}};
  return rule;
}

cf::CFRule DataBarOf(cf::CfvoType min, cf::CfvoType max) {
  cf::CFRule rule;
  rule.type = cf::RuleType::DataBar;
  rule.data_bar = cf::DataBarSpec{};
  rule.data_bar->min = Cfvo(min);
  rule.data_bar->max = Cfvo(max);
  return rule;
}

cf::CFRule IconSetOf(cf::CfvoType floor, std::vector<cf::CfValueObject> thresholds) {
  cf::CFRule rule;
  rule.type = cf::RuleType::IconSet;
  rule.icon_set = cf::IconSetSpec{};
  rule.icon_set->floor = Cfvo(floor);
  rule.icon_set->thresholds = std::move(thresholds);
  return rule;
}

TEST(CfSqrefFarEdgeFixture, RuleSplitsOnlyWhenRankingOrRelativeThreshold) {
  using cf::CfvoType;
  for (const cf::RuleType type :
       {cf::RuleType::Top10, cf::RuleType::AboveAverage, cf::RuleType::DuplicateValues, cf::RuleType::UniqueValues}) {
    cf::CFRule rule;
    rule.type = type;
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(rule));
  }
  for (const CfvoType type : {CfvoType::Min, CfvoType::Max, CfvoType::Percent, CfvoType::AutoMin, CfvoType::AutoMax}) {
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(ColorScaleOf({Cfvo(CfvoType::Number), Cfvo(type)})));
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(DataBarOf(CfvoType::Number, type)));
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(DataBarOf(type, CfvoType::Number)));
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(IconSetOf(CfvoType::Percentile, {Cfvo(CfvoType::Number), Cfvo(type)})));
    EXPECT_TRUE(cf::far_edge_edit_splits_rule(IconSetOf(type, {Cfvo(CfvoType::Number)})));
  }
  for (const CfvoType type : {CfvoType::Percentile, CfvoType::Number, CfvoType::Formula}) {
    EXPECT_FALSE(cf::far_edge_edit_splits_rule(ColorScaleOf({Cfvo(type), Cfvo(CfvoType::Percentile)})));
    EXPECT_FALSE(cf::far_edge_edit_splits_rule(DataBarOf(type, CfvoType::Number)));
    EXPECT_FALSE(cf::far_edge_edit_splits_rule(IconSetOf(type, {Cfvo(CfvoType::Percentile), Cfvo(CfvoType::Number)})));
  }
  for (const cf::RuleType type : {cf::RuleType::Expression, cf::RuleType::CellIs, cf::RuleType::ContainsText,
                                  cf::RuleType::ContainsBlanks, cf::RuleType::TimePeriod}) {
    cf::CFRule rule;
    rule.type = type;
    EXPECT_FALSE(cf::far_edge_edit_splits_rule(rule));
  }
}

}  // namespace
}  // namespace formulon
