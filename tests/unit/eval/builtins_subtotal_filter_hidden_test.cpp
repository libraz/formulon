// SUBTOTAL(1..11) excludes filter-hidden rows and 101..111 every hidden row.
// A hidden row counts as filtered exactly when an AutoFilter on its sheet (the
// sheet's own or a table's) carries a criterion, inside the filter range or
// not. Expected values were measured in Excel after save, close and reopen.

#include <cstdint>
#include <set>
#include <string>
#include <vector>

#include "auto_filter.h"
#include "gtest/gtest.h"
#include "sheet.h"
#include "table.h"
#include "util/test_eval_helpers.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace {

enum class Filter : std::uint8_t {
  kNone,            // No AutoFilter.
  kNoCriterion,     // Sheet AutoFilter without a criterion.
  kCriterion,       // Sheet AutoFilter with a column-A criterion.
  kTwoCriteria,     // Sheet AutoFilter with criteria on columns A and B.
  kTableCriterion,  // Table over the range whose AutoFilter has a criterion.
};

struct Case {
  const char* name;
  Filter filter;
  std::uint32_t last_ref_row;  // 1-based last row of the filter range.
  std::uint32_t last_row;      // 1-based last data row probed.
  std::vector<std::uint32_t> hidden;
  std::vector<std::uint32_t> excluded_by_9;
  std::vector<std::uint32_t> excluded_by_109;
};

const std::vector<std::uint32_t> kOddPlus6And25 = {3, 5, 6, 7, 9, 11, 13, 15, 17, 19, 21, 25};
const std::vector<std::uint32_t> kOddPlus6And25And40 = {3, 5, 6, 7, 9, 11, 13, 15, 17, 19, 21, 25, 40};

const std::vector<Case>& Cases() {
  static const std::vector<Case> cases = {
      {"Cond", Filter::kCriterion, 21, 30, kOddPlus6And25, kOddPlus6And25, kOddPlus6And25},
      {"NoCrit", Filter::kNoCriterion, 21, 30, {4, 9, 25}, {}, {4, 9, 25}},
      {"Tbl", Filter::kTableCriterion, 21, 30, kOddPlus6And25, kOddPlus6And25, kOddPlus6And25},
      {"C1", Filter::kCriterion, 21, 40, kOddPlus6And25And40, kOddPlus6And25And40, kOddPlus6And25And40},
      {"C2", Filter::kCriterion, 10, 40, {3, 5, 7, 9, 15, 25}, {3, 5, 7, 9, 15, 25}, {3, 5, 7, 9, 15, 25}},
      {"N1", Filter::kNoCriterion, 21, 40, {4, 9, 25}, {}, {4, 9, 25}},
      {"X", Filter::kNone, 21, 40, {4, 25}, {}, {4, 25}},
      {"C3", Filter::kCriterion, 21, 40, {}, {}, {}},
      {"C4", Filter::kTwoCriteria, 21, 40, kOddPlus6And25, kOddPlus6And25, kOddPlus6And25},
  };
  return cases;
}

FilterColumn ValuesColumn(std::uint32_t col_id, std::vector<std::string> values) {
  FilterColumn column;
  column.col_id = col_id;
  column.kind = FilterKind::kValues;
  column.values.values = std::move(values);
  return column;
}

// Column A holds "x" on even and "y" on odd rows 2..21 and "o" from row 23;
// column B holds the row number. Row 22 is empty.
Workbook Build(const Case& c) {
  Workbook wb = Workbook::create();
  Sheet& sheet = wb.sheet(0);
  sheet.set_cell_value(0, 0, Value::text("key"));
  sheet.set_cell_value(0, 1, Value::text("val"));
  for (std::uint32_t r = 2; r <= c.last_row; ++r) {
    if (r == 22) {
      continue;
    }
    const char* key = r > 22 ? "o" : (r % 2U == 0U ? "x" : "y");
    sheet.set_cell_value(r - 1U, 0, Value::text(key));
    sheet.set_cell_value(r - 1U, 1, Value::number(static_cast<double>(r)));
  }
  for (std::uint32_t r : c.hidden) {
    RowLayout layout;
    layout.row = r - 1U;
    layout.hidden = true;
    sheet.mutable_layout().row_overrides.push_back(layout);
  }

  AutoFilter filter;
  filter.range = MergeRange{0, 0, c.last_ref_row - 1U, 1};
  if (c.filter == Filter::kCriterion || c.filter == Filter::kTwoCriteria || c.filter == Filter::kTableCriterion) {
    filter.columns.push_back(ValuesColumn(0, {"x"}));
  }
  if (c.filter == Filter::kTwoCriteria) {
    filter.columns.push_back(ValuesColumn(1, {"2", "4"}));
  }
  if (c.filter == Filter::kTableCriterion) {
    TableMetadata table;
    table.id = 1;
    table.name = "T1";
    table.display_name = "T1";
    table.ref = "A1:B21";
    table.columns.emplace_back(1U, "key", "", "", "");
    table.columns.emplace_back(2U, "val", "", "", "");
    wb.mutable_tables().push_back(table);
    EXPECT_TRUE(static_cast<bool>(wb.set_table_auto_filter(0, filter)));
  } else if (c.filter != Filter::kNone) {
    EXPECT_TRUE(static_cast<bool>(wb.set_sheet_auto_filter(0, filter)));
  }
  return wb;
}

double Eval(const Workbook& wb, const std::string& formula) {
  const Value v = test::EvalSourceIn(formula, wb, wb.sheet(0));
  EXPECT_TRUE(v.is_number()) << formula;
  return v.is_number() ? v.as_number() : -1.0;
}

TEST(SubtotalFilterHidden, ProbeTable) {
  for (const Case& c : Cases()) {
    const Workbook wb = Build(c);
    const std::set<std::uint32_t> by9(c.excluded_by_9.begin(), c.excluded_by_9.end());
    const std::set<std::uint32_t> by109(c.excluded_by_109.begin(), c.excluded_by_109.end());
    for (std::uint32_t r = 2; r <= c.last_row; ++r) {
      if (r == 22) {
        continue;
      }
      const std::string ref = "B" + std::to_string(r);
      EXPECT_EQ(Eval(wb, "=SUBTOTAL(9," + ref + ")") == 0.0, by9.count(r) == 1U) << c.name << " row " << r;
      EXPECT_EQ(Eval(wb, "=SUBTOTAL(109," + ref + ")") == 0.0, by109.count(r) == 1U) << c.name << " row " << r;
    }
  }
}

// Region totals measured on the Excel-built sheets.
TEST(SubtotalFilterHidden, RegionSums) {
  struct Sums {
    const char* name;
    double in9, in109, out9, out109, all9, all109;
  };
  const Sums sums[] = {
      {"Cond", 104, 104, 187, 187, 291, 291},
      {"NoCrit", 230, 217, 212, 187, 442, 404},
      {"Tbl", 104, 104, 187, 187, 291, 291},
  };
  for (const Sums& s : sums) {
    const Case* found = nullptr;
    for (const Case& c : Cases()) {
      if (std::string(c.name) == s.name) {
        found = &c;
      }
    }
    ASSERT_NE(found, nullptr);
    const Workbook wb = Build(*found);
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(9,B2:B21)"), s.in9) << s.name;
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(109,B2:B21)"), s.in109) << s.name;
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(9,B23:B30)"), s.out9) << s.name;
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(109,B23:B30)"), s.out109) << s.name;
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(9,B1:B30)"), s.all9) << s.name;
    EXPECT_DOUBLE_EQ(Eval(wb, "=SUBTOTAL(109,B1:B30)"), s.all109) << s.name;
  }
}

}  // namespace
}  // namespace formulon
