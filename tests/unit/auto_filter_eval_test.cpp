#include "cf/auto_filter_eval.h"

#include <cstdint>
#include <cstdio>
#include <iterator>
#include <memory>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "cf/cf_types.h"
#include "excel_profile.h"
#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
#include "sheet_layout_edit.h"
#include "util/test_eval_helpers.h"
#include "utils/arena.h"
#include "utils/date_time.h"
#include "value.h"
#include "workbook.h"

#ifndef FORMULON_FIXTURES_DIR
#error "FORMULON_FIXTURES_DIR must be defined by the build"
#endif

namespace formulon {
namespace {

std::vector<std::uint8_t> ReadFile(const char* file) {
  std::vector<std::uint8_t> out;
  const std::string path = std::string(FORMULON_FIXTURES_DIR) + "/excel/" + file;
  FILE* fp = std::fopen(path.c_str(), "rb");
  if (fp == nullptr) {
    ADD_FAILURE() << "could not open fixture: " << path;
    return out;
  }
  std::fseek(fp, 0, SEEK_END);
  const long size = std::ftell(fp);
  std::fseek(fp, 0, SEEK_SET);
  out.resize(size > 0 ? static_cast<std::size_t>(size) : 0U);
  if (std::fread(out.data(), 1, out.size(), fp) != out.size()) {
    ADD_FAILURE() << "short read on fixture: " << path;
    out.clear();
  }
  std::fclose(fp);
  return out;
}

bool RowHidden(const Sheet& sheet, std::uint32_t row) {
  for (const RowLayout& r : sheet.layout().row_overrides) {
    if (r.row == row) {
      return r.hidden;
    }
  }
  return false;
}

void SetRowHidden(Sheet& sheet, std::uint32_t row, bool hidden) {
  upsert_row_override(sheet.mutable_layout(), row, sheet.format_defaults().zero_height).hidden = hidden;
}

// Every sheet of this workbook was filtered and then saved by Excel, so its
// saved row-hidden state is the visible set Mac Excel (ja-JP) computed. Relative-date
// cases were filtered on Monday 2026-10-05.
std::unique_ptr<Workbook> LoadKinds() {
  const std::vector<std::uint8_t> bytes = ReadFile("auto_filter_kinds.xlsx");
  auto read = io::read_ooxml(io::ByteSpan{bytes.data(), bytes.size()});
  if (!read) {
    ADD_FAILURE() << "fixture did not load";
    return nullptr;
  }
  auto wb = std::make_unique<Workbook>(std::move(read.value().workbook));
  wb->set_pinned_now(date_time::CivilTime{{2026, 10, 5U}, {12U, 0U, 0U}});
  wb->set_excel_profile(mac_365_ja_jp_profile());
  return wb;
}

TEST(AutoFilterEval, ProbeTable) {
  std::unique_ptr<Workbook> wb = LoadKinds();
  ASSERT_NE(wb, nullptr);
  ASSERT_EQ(wb->sheet_count(), 60U);
  std::size_t evaluated = 0;
  for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
    const Sheet& sheet = wb->sheet(i);
    const AutoFilter* filter = sheet.auto_filter();
    ASSERT_NE(filter, nullptr) << sheet.name();
    auto visible = evaluate_auto_filter(*wb, sheet, *filter);
    ASSERT_TRUE(static_cast<bool>(visible)) << sheet.name();
    ASSERT_EQ(visible.value().size(), filter->range.last_row - filter->range.first_row) << sheet.name();
    std::string mismatched;
    for (std::uint32_t k = 0; k < visible.value().size(); ++k) {
      const std::uint32_t row = filter->range.first_row + 1U + k;
      if (visible.value()[k] == RowHidden(sheet, row)) {
        mismatched += " " + std::to_string(row + 1U);
      }
    }
    EXPECT_TRUE(mismatched.empty()) << sheet.name() << " differs on rows" << mismatched;
    ++evaluated;
  }
  EXPECT_EQ(evaluated, 60U);
}

TEST(AutoFilterEval, IconFilterMatchesConditionalFormatIconsAcrossVisibleRows) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("header"));
  const std::vector<double> values{10.0, 30.0, 50.0, 70.0, 90.0};
  for (std::size_t i = 0; i < values.size(); ++i) {
    sheet.set_cell_value(static_cast<std::uint32_t>(i + 1U), 0, Value::number(values[i]));
  }

  cf::CFRule icon_rule;
  icon_rule.type = cf::RuleType::IconSet;
  icon_rule.priority = 1;
  icon_rule.id = "arrows";
  cf::IconSetSpec icon_spec;
  icon_spec.name = cf::IconSetName::Three_Arrows;
  icon_spec.floor = cf::CfValueObject{cf::CfvoType::Number, "0", true};
  icon_spec.thresholds = {cf::CfValueObject{cf::CfvoType::Number, "33", true},
                          cf::CfValueObject{cf::CfvoType::Number, "67", true}};
  icon_rule.icon_set = std::move(icon_spec);
  cf::ConditionalFormat conditional_format;
  conditional_format.sqref.push_back(cf::CFCellRange{CellAddress{1, 0}, CellAddress{5, 0}});
  conditional_format.rules.push_back(std::move(icon_rule));
  sheet.mutable_conditional_formats().push_back(std::move(conditional_format));

  AutoFilter filter;
  filter.range = MergeRange{0, 0, 5, 0};
  FilterColumn column;
  column.kind = FilterKind::kIcon;
  column.icon.icon_set = cf::IconSetName::Three_Arrows;
  column.icon.icon_id = 2U;
  filter.columns.push_back(std::move(column));

  const auto visible = evaluate_auto_filter(wb, sheet, filter);
  ASSERT_TRUE(static_cast<bool>(visible));
  ASSERT_EQ(visible.value().size(), values.size());
  EXPECT_EQ(visible.value(), (std::vector<bool>{false, false, false, true, true}));
}

TEST(AutoFilterEval, ConditionalFormatResourceFailureLeavesRowVisibilityUnchanged) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0U, 0U, Value::text("header"));
  sheet.set_cell_value(1U, 0U, Value::number(1.0));
  sheet.set_cell_value(2U, 0U, Value::number(2.0));
  SetRowHidden(sheet, 1U, true);
  SetRowHidden(sheet, 2U, false);

  cf::CFRule expression;
  expression.type = cf::RuleType::Expression;
  expression.priority = 1;
  expression.id = "resource-limited-expression";
  expression.formula1 = "A1>0";
  cf::ConditionalFormat block;
  block.sqref.push_back(cf::CFCellRange{CellAddress{1U, 0U}, CellAddress{2U, 0U}});
  block.rules.push_back(std::move(expression));
  sheet.mutable_conditional_formats().push_back(std::move(block));

  AutoFilter filter;
  filter.range = MergeRange{0U, 0U, 2U, 0U};
  FilterColumn column;
  column.kind = FilterKind::kColor;
  filter.columns.push_back(std::move(column));

  Arena arena(512U, 0U);
  AutoFilterEvalDeps deps;
  deps.arena = &arena;
  const auto result = apply_auto_filter(wb, 0U, filter, deps);

  ASSERT_FALSE(static_cast<bool>(result));
  EXPECT_EQ(result.error().code, FormulonErrorCode::kSecResourceLimit);
  EXPECT_NE(result.error().context.find("cell="), std::string::npos);
  EXPECT_NE(result.error().context.find("ceiling="), std::string::npos);
  EXPECT_TRUE(RowHidden(sheet, 1U));
  EXPECT_FALSE(RowHidden(sheet, 2U));
}

TEST(AutoFilterEval, Top10UsesNthElementAndKeepsCutoffTies) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("header"));
  std::uint32_t row = 1U;
  for (const double value : {5.0, 5.0, 4.0, 1.0}) {
    sheet.set_cell_value(row++, 0, Value::number(value));
  }

  AutoFilter filter;
  filter.range = MergeRange{0, 0, 4, 0};
  FilterColumn column;
  column.kind = FilterKind::kTop10;
  column.top10.top = true;
  column.top10.val = 2.0;
  filter.columns.push_back(column);
  auto visible = evaluate_auto_filter(wb, sheet, filter);
  ASSERT_TRUE(static_cast<bool>(visible));
  EXPECT_EQ(visible.value(), (std::vector<bool>{true, true, false, false}));

  filter.columns[0].top10.top = false;
  visible = evaluate_auto_filter(wb, sheet, filter);
  ASSERT_TRUE(static_cast<bool>(visible));
  EXPECT_EQ(visible.value(), (std::vector<bool>{false, false, true, true}));

  filter.columns[0].top10.top = true;
  filter.columns[0].top10.percent = true;
  filter.columns[0].top10.val = 50.0;
  visible = evaluate_auto_filter(wb, sheet, filter);
  ASSERT_TRUE(static_cast<bool>(visible));
  EXPECT_EQ(visible.value(), (std::vector<bool>{true, true, false, false}));
}

TEST(AutoFilterEval, ApplyMergesRowsAndPreservesMetadata) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("header"));
  std::uint32_t row = 1U;
  for (const char* value : {"keep", "drop", "keep", "drop"}) {
    sheet.set_cell_value(row++, 0, Value::text(value));
  }

  RowLayout before;
  before.row = 0U;
  before.outline_level = 2U;
  before.has_style = true;
  before.style_xf = 4U;
  sheet.mutable_layout().row_overrides.push_back(before);
  RowLayout first_body;
  first_body.row = 1U;
  first_body.hidden = true;
  first_body.height = 22.0;
  first_body.has_height = true;
  first_body.custom_height = true;
  first_body.has_style = true;
  first_body.style_xf = 7U;
  sheet.mutable_layout().row_overrides.push_back(first_body);
  RowLayout third_body;
  third_body.row = 3U;
  third_body.outline_level = 1U;
  third_body.hidden = false;
  sheet.mutable_layout().row_overrides.push_back(third_body);
  RowLayout after;
  after.row = 5U;
  after.hidden = true;
  sheet.mutable_layout().row_overrides.push_back(after);

  AutoFilter filter;
  filter.range = MergeRange{0, 0, 4, 0};
  FilterColumn column;
  column.kind = FilterKind::kValues;
  column.values.values.push_back("keep");
  filter.columns.push_back(column);
  ASSERT_TRUE(static_cast<bool>(apply_auto_filter(wb, 0, filter)));

  ASSERT_EQ(sheet.layout().row_overrides.size(), 6U);
  EXPECT_EQ(sheet.layout().row_overrides[0].row, 0U);
  EXPECT_EQ(sheet.layout().row_overrides[0].style_xf, 4U);
  const RowLayout* row1 = find_row_override(sheet.layout(), 1U);
  ASSERT_NE(row1, nullptr);
  EXPECT_FALSE(row1->hidden);
  EXPECT_DOUBLE_EQ(row1->height, 22.0);
  EXPECT_TRUE(row1->custom_height);
  EXPECT_EQ(row1->style_xf, 7U);
  const RowLayout* row2 = find_row_override(sheet.layout(), 2U);
  ASSERT_NE(row2, nullptr);
  EXPECT_TRUE(row2->hidden);
  const RowLayout* row3 = find_row_override(sheet.layout(), 3U);
  ASSERT_NE(row3, nullptr);
  EXPECT_FALSE(row3->hidden);
  EXPECT_EQ(row3->outline_level, 1U);
  const RowLayout* row4 = find_row_override(sheet.layout(), 4U);
  ASSERT_NE(row4, nullptr);
  EXPECT_TRUE(row4->hidden);
  const RowLayout* row5 = find_row_override(sheet.layout(), 5U);
  ASSERT_NE(row5, nullptr);
  EXPECT_TRUE(row5->hidden);
}

TEST(AutoFilterEval, ApplyFilterMakesMatchingRowsExplicitlyVisibleWithZeroHeightDefault) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.mutable_format_defaults().zero_height = true;
  sheet.set_cell_value(0, 0, Value::text("header"));
  sheet.set_cell_value(1, 0, Value::text("keep"));
  sheet.set_cell_value(2, 0, Value::text("drop"));
  AutoFilter filter;
  filter.range = MergeRange{0, 0, 2, 0};
  FilterColumn column;
  column.kind = FilterKind::kValues;
  column.values.values.push_back("keep");
  filter.columns.push_back(column);

  ASSERT_TRUE(static_cast<bool>(apply_auto_filter(wb, 0, filter)));
  const RowLayout* matching = find_row_override(sheet.layout(), 1U);
  ASSERT_NE(matching, nullptr);
  EXPECT_FALSE(matching->hidden);
  const RowLayout* non_matching = find_row_override(sheet.layout(), 2U);
  ASSERT_NE(non_matching, nullptr);
  EXPECT_TRUE(non_matching->hidden);

  // The explicit hidden state must survive if the sheet later stops using a
  // hidden default; otherwise clearing zeroHeight would reveal filtered rows.
  sheet.mutable_format_defaults().zero_height = false;
  EXPECT_TRUE(RowHidden(sheet, 2U));
}

// Reapplying re-derives every body row: a matching row hidden by hand is shown
// again, non-matching rows are hidden, and rows outside the range keep their state.
TEST(AutoFilterEval, ApplyUnhidesMatchingRows) {
  std::unique_ptr<Workbook> wb = LoadKinds();
  ASSERT_NE(wb, nullptr);
  const std::size_t index = wb->sheet_index_by_name("val_case_apple");
  ASSERT_LT(index, wb->sheet_count());
  Sheet& sheet = wb->sheet(index);
  const AutoFilter filter = *sheet.auto_filter();

  std::set<std::uint32_t> excel_hidden;
  for (std::uint32_t row = 1; row <= filter.range.last_row; ++row) {
    if (RowHidden(sheet, row)) {
      excel_hidden.insert(row);
    }
    SetRowHidden(sheet, row, false);
  }
  SetRowHidden(sheet, 6, true);   // "apple" (row 7) hidden by hand.
  SetRowHidden(sheet, 19, true);  // Row 20, outside the range.

  ASSERT_TRUE(static_cast<bool>(apply_auto_filter(*wb, index, filter)));
  for (std::uint32_t row = 1; row <= filter.range.last_row; ++row) {
    EXPECT_EQ(RowHidden(sheet, row), excel_hidden.count(row) == 1U) << "row " << row + 1U;
  }
  EXPECT_FALSE(RowHidden(sheet, 6));
  EXPECT_TRUE(RowHidden(sheet, 19));

  // The sheet now carries a criterion, so SUBTOTAL(3) skips the filtered rows
  // and the outside row alike: rows 6..8 are the only visible values.
  const Value count = test::EvalSourceIn("=SUBTOTAL(3,A2:A20)", *wb, sheet);
  ASSERT_TRUE(count.is_number());
  EXPECT_DOUBLE_EQ(count.as_number(), 3.0);
}

// Custom-filter conditions x cell types x and/or.
//
// Pairwise model (all-pairs by Latin-square construction, join = (cond + cell) % 3):
//   cond (7): a* | ?b | *b* | <>x | <>a* | >b | <>" " (non-blank)
//   cell (10): "a" "ab" "*b" "APPLE" "x" ="" blank '5 "ba" "~"   (fixture rows 2,4,5,8,9,10,11,12,13,14)
//   join (3): single | and *b* | or *b*
// Every cond x cell pair occurs once and each join meets every cond and every
// cell. Expected values combine single-condition results Excel produced on
// the same rows of the fixture's custom-filter sheets.
struct Cond {
  const char* name;
  FilterOperator op;
  const char* val;
  std::set<std::uint32_t> excel_visible;  // 1-based rows of the fixture's text cases.
};

const std::vector<Cond>& Conds() {
  static const std::vector<Cond> conds = {
      {"a*", FilterOperator::kEqual, "a*", {2, 4, 6, 7, 8, 15}},
      {"?b", FilterOperator::kEqual, "?b", {4, 5}},
      {"*b*", FilterOperator::kEqual, "*b*", {3, 4, 5, 13, 15, 16}},
      {"<>x", FilterOperator::kNotEqual, "x", {2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 13, 14, 15, 16}},
      {"<>a*", FilterOperator::kNotEqual, "a*", {3, 5, 9, 10, 11, 12, 13, 14, 16}},
      {">b", FilterOperator::kGreaterThan, "b", {9, 13, 16}},
      {"nonblank", FilterOperator::kNotEqual, " ", {2, 3, 4, 5, 6, 7, 8, 9, 12, 13, 14, 15, 16}},
  };
  return conds;
}

TEST(AutoFilterEval, CustomConditionPairwise) {
  std::unique_ptr<Workbook> wb = LoadKinds();
  ASSERT_NE(wb, nullptr);
  const std::size_t index = wb->sheet_index_by_name("cus_ne_x");
  ASSERT_LT(index, wb->sheet_count());
  const Sheet& sheet = wb->sheet(index);
  const std::vector<std::uint32_t> cells = {2, 4, 5, 8, 9, 10, 11, 12, 13, 14};
  const Cond& second = Conds()[2];

  std::size_t cases = 0;
  for (std::size_t c = 0; c < Conds().size(); ++c) {
    for (std::size_t k = 0; k < cells.size(); ++k) {
      const Cond& cond = Conds()[c];
      const std::uint32_t row = cells[k];
      const int join = static_cast<int>((c + k) % 3U);

      AutoFilter filter;
      filter.range = MergeRange{0, 0, 15, 0};
      FilterColumn column;
      column.kind = FilterKind::kCustom;
      column.custom.filters.push_back(CustomFilter{cond.op, cond.val});
      bool expected = cond.excel_visible.count(row) == 1U;
      if (join != 0) {
        column.custom.and_join = join == 1;
        column.custom.filters.push_back(CustomFilter{second.op, second.val});
        const bool other = second.excel_visible.count(row) == 1U;
        expected = join == 1 ? (expected && other) : (expected || other);
      }
      filter.columns.push_back(column);

      auto visible = evaluate_auto_filter(*wb, sheet, filter);
      ASSERT_TRUE(static_cast<bool>(visible));
      EXPECT_EQ(visible.value()[row - 2U], expected)
          << cond.name << " row " << row << " join " << (join == 0 ? "single" : (join == 1 ? "and" : "or"));
      ++cases;
    }
  }
  EXPECT_EQ(cases, 70U);
}

TEST(AutoFilterEval, OpaqueFilterIsRejected) {
  Workbook wb = Workbook::create();
  AutoFilter filter;
  filter.opaque_xml = "<autoFilter ref=\"A1:$B$2\"/>";
  auto visible = evaluate_auto_filter(wb, wb.sheet(0), filter);
  ASSERT_FALSE(static_cast<bool>(visible));
  EXPECT_EQ(visible.error().code, FormulonErrorCode::kAutoFilterInvalid);
}

TEST(AutoFilterEval, TableCriterionMakesSheetFiltered) {
  Workbook wb = Workbook::create();
  EXPECT_FALSE(sheet_has_filter_criteria(&wb, wb.sheet(0)));
  TableMetadata table;
  table.id = 1;
  table.name = "T1";
  table.display_name = "T1";
  table.ref = "A1:B3";
  wb.mutable_tables().push_back(table);
  AutoFilter filter;
  filter.range = MergeRange{0, 0, 2, 1};
  ASSERT_TRUE(static_cast<bool>(wb.set_table_auto_filter(0, filter)));
  EXPECT_FALSE(sheet_has_filter_criteria(&wb, wb.sheet(0)));
  FilterColumn column;
  column.kind = FilterKind::kValues;
  column.values.values.push_back("x");
  filter.columns.push_back(column);
  ASSERT_TRUE(static_cast<bool>(wb.set_table_auto_filter(0, filter)));
  EXPECT_TRUE(sheet_has_filter_criteria(&wb, wb.sheet(0)));
}

// Locale collation: value-list and custom filters follow the workbook profile.
// Expectations are the measured equality and ordering fixtures
// (tools/collation/data/fixtures).
struct FilterEqualityCase {
  const char* label;
  const char* a;
  const char* b;
  bool ja;
  bool en;
  bool zh;
  bool win_ja;
};

constexpr FilterEqualityCase kFilterEqualityCases[] = {
    {"circled one vs 1", "\xE2\x91\xA0", "1", true, true, false, false},
    {"fullwidth A vs A", "\xEF\xBC\xA1", "A", true, true, false, false},
    {"katakana a vs hiragana a", "\xE3\x82\xA2", "\xE3\x81\x82", true, true, false, false},
    {"iteration marks", "\xE3\x82\x9D", "\xE3\x83\xBD", true, false, true, true},
    {"sharp s vs ss", "\xC3\x9F", "ss", true, true, true, true},
    {"a-umlaut vs a", "\xC3\xA4", "a", false, false, false, false},
    {"A vs a", "A", "a", true, true, true, true},
};

struct FilterProfile {
  const char* name;
  ExcelProfile profile;
};

const FilterProfile kFilterProfiles[] = {
    {"mac-ja_JP", {ExcelHost::kMac365, ExcelLocale::kJaJP}},
    {"mac-en_US", {ExcelHost::kMac365, ExcelLocale::kEnUS}},
    {"mac-zh_CN", {ExcelHost::kMac365, ExcelLocale::kZhCN}},
    {"win-ja_JP", {ExcelHost::kWin365, ExcelLocale::kJaJP}},
};

bool ExpectedFilterEqual(const FilterEqualityCase& c, std::size_t p) {
  switch (p) {
    case 0:
      return c.ja;
    case 1:
      return c.en;
    case 2:
      return c.zh;
    default:
      return c.win_ja;
  }
}

// One header row and one body row holding `cell`; returns the body row's flag.
bool BodyRowVisible(ExcelProfile profile, const char* cell, FilterColumn column) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(profile);
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("h"));
  sheet.set_cell_value(1, 0, Value::text(cell));
  AutoFilter filter;
  filter.range = MergeRange{0, 0, 1, 0};
  filter.columns.push_back(std::move(column));
  auto visible = evaluate_auto_filter(wb, sheet, filter);
  EXPECT_TRUE(static_cast<bool>(visible));
  return visible && visible.value().size() == 1U && visible.value()[0];
}

TEST(AutoFilterEvalCollation, ValueListEqualityFollowsProfile) {
  for (std::size_t p = 0; p < std::size(kFilterProfiles); ++p) {
    for (const FilterEqualityCase& c : kFilterEqualityCases) {
      FilterColumn column;
      column.kind = FilterKind::kValues;
      column.values.values.push_back(c.b);
      EXPECT_EQ(BodyRowVisible(kFilterProfiles[p].profile, c.a, column), ExpectedFilterEqual(c, p))
          << kFilterProfiles[p].name << " " << c.label;
    }
  }
}

TEST(AutoFilterEvalCollation, CustomEqualityFollowsProfile) {
  for (std::size_t p = 0; p < std::size(kFilterProfiles); ++p) {
    for (const FilterEqualityCase& c : kFilterEqualityCases) {
      for (FilterOperator op : {FilterOperator::kEqual, FilterOperator::kNotEqual}) {
        FilterColumn column;
        column.kind = FilterKind::kCustom;
        column.custom.filters.push_back(CustomFilter{op, c.b});
        const bool equal = ExpectedFilterEqual(c, p);
        EXPECT_EQ(BodyRowVisible(kFilterProfiles[p].profile, c.a, column),
                  op == FilterOperator::kEqual ? equal : !equal)
            << kFilterProfiles[p].name << " " << c.label;
      }
    }
  }
}

// Leading and trailing U+0020 are trimmed from the entry and the cell text before comparing.
TEST(AutoFilterEvalCollation, ValueListTrimsAsciiSpaces) {
  const ExcelProfile profile{ExcelHost::kMac365, ExcelLocale::kJaJP};
  auto visible = [&](const char* cell, const char* entry) {
    FilterColumn column;
    column.kind = FilterKind::kValues;
    column.values.values.push_back(entry);
    return BodyRowVisible(profile, cell, column);
  };
  for (const char* cell : {"abc", "abc ", "abc  ", " abc"}) {
    EXPECT_TRUE(visible(cell, "abc")) << cell;
  }
  for (const char* cell : {"a bc", "abc\xC2\xA0", "abc\xE3\x80\x80", "abc\t", "abc\xC2\xAD"}) {
    EXPECT_FALSE(visible(cell, "abc")) << cell;
  }
  EXPECT_TRUE(visible("abc", " abc"));
  EXPECT_TRUE(visible("abc", "abc "));
}

// A wildcard criterion keeps the engine's pattern matching on every profile.
TEST(AutoFilterEvalCollation, WildcardCriterionIsUnchanged) {
  for (const FilterProfile& fp : kFilterProfiles) {
    FilterColumn column;
    column.kind = FilterKind::kCustom;
    column.custom.filters.push_back(CustomFilter{FilterOperator::kEqual, "a*"});
    EXPECT_TRUE(BodyRowVisible(fp.profile, "ab", column)) << fp.name;
    EXPECT_FALSE(BodyRowVisible(fp.profile, "ba", column)) << fp.name;
  }
}

// An escaped metacharacter is a literal: "~*" matches the text "*" only.
TEST(AutoFilterEvalCollation, EscapedCriterionIsLiteral) {
  for (const FilterProfile& fp : kFilterProfiles) {
    FilterColumn column;
    column.kind = FilterKind::kCustom;
    column.custom.filters.push_back(CustomFilter{FilterOperator::kEqual, "~*"});
    EXPECT_TRUE(BodyRowVisible(fp.profile, "*", column)) << fp.name;
    EXPECT_FALSE(BodyRowVisible(fp.profile, "ab", column)) << fp.name;
  }
}

// "$" sorts after "%" on ja_JP and en_US and before it on zh_CN.
TEST(AutoFilterEvalCollation, CustomOrderingFollowsProfile) {
  const bool expected_greater[] = {true, true, false};
  for (std::size_t p = 0; p < 3U; ++p) {
    FilterColumn column;
    column.kind = FilterKind::kCustom;
    column.custom.filters.push_back(CustomFilter{FilterOperator::kGreaterThan, "%"});
    EXPECT_EQ(BodyRowVisible(kFilterProfiles[p].profile, "$", column), expected_greater[p]) << kFilterProfiles[p].name;
  }
}

}  // namespace
}  // namespace formulon
