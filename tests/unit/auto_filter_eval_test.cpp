#include "cf/auto_filter_eval.h"

#include <cstdint>
#include <cstdio>
#include <memory>
#include <set>
#include <string>
#include <utility>
#include <vector>

#include "gtest/gtest.h"
#include "io/ooxml_reader.h"
#include "sheet.h"
#include "util/test_eval_helpers.h"
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
  for (RowLayout& r : sheet.mutable_layout().row_overrides) {
    if (r.row == row) {
      r.hidden = hidden;
      return;
    }
  }
  RowLayout fresh;
  fresh.row = row;
  fresh.hidden = hidden;
  sheet.mutable_layout().row_overrides.push_back(fresh);
}

// Every sheet of this workbook was filtered and then saved by Excel, so its
// saved row-hidden state is the visible set Excel computed. Relative-date
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

}  // namespace
}  // namespace formulon
