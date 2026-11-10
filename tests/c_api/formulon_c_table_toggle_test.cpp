// fm_sheet_update_table: the order of header / totals toggles, range and column patches, and refusals that change
// nothing.

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_pivot_test_helpers.h"
#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr std::uint32_t kLastRow = 1048575U;

/// Sheet1 with table `Sales` over `ref` (two columns, B:C by default): Product / Amount headers on its first row, then
/// R1 / 10, R2 / 20, ... on the data rows, and a totals row on the last row when `totals`.
void Create(WorkbookGuard& wb, fm_cell_rect_t ref = {2, 1, 4, 2}, std::int32_t totals = 0) {
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, ref.first_row, ref.first_col, "Product"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, ref.first_row, ref.first_col + 1U, "Amount"), 0);
  const std::uint32_t data_end = ref.last_row - (totals != 0 ? 1U : 0U);
  for (std::uint32_t row = ref.first_row + 1U; row <= data_end; ++row) {
    const std::uint32_t n = row - ref.first_row;
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, row, ref.first_col, ("R" + std::to_string(n)).c_str()), 0);
    ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, row, ref.first_col + 1U, 10.0 * n), 0);
  }
  const char* columns[] = {"Product", "Amount"};
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, "Sales", columns, 2, "TableStyleMedium2", 1, totals, &index), 0)
      << fm_get_last_error_message();
}

/// A second table `name` over the two-column `ref`, headers K / L.
void AddOther(WorkbookGuard& wb, const char* name, fm_cell_rect_t ref) {
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, ref.first_row, ref.first_col, "K"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, ref.first_row, ref.first_col + 1U, "L"), 0);
  const char* columns[] = {"K", "L"};
  std::size_t index = 99U;
  ASSERT_EQ(AddTable(wb.handle, 0, &ref, name, columns, 2, "TableStyleMedium2", 1, 0, &index), 0)
      << fm_get_last_error_message();
}

fm_table_info_t Info(fm_workbook_t* wb, std::size_t index = 0U) {
  fm_table_info_t info = FM_STRUCT_INIT(fm_table_info_t);
  EXPECT_EQ(fm_sheet_get_table_at(wb, 0, index, &info), 0);
  return info;
}

/// A patch setting the toggles given as 0 / 1 (-1 keeps one) and, with `range`, the range.
fm_table_patch_t Patch(std::int32_t header, std::int32_t totals, const fm_cell_rect_t* range = nullptr) {
  fm_table_patch_t patch = FM_STRUCT_INIT(fm_table_patch_t);
  if (header >= 0) {
    patch.header_row_op = FM_PATCH_OP_SET;
    patch.header_row = header;
  }
  if (totals >= 0) {
    patch.totals_row_op = FM_PATCH_OP_SET;
    patch.totals_row = totals;
  }
  if (range != nullptr) {
    patch.range_op = FM_PATCH_OP_SET;
    patch.range = *range;
  }
  return patch;
}

fm_status_t Update(fm_workbook_t* wb, const fm_table_patch_t& patch) {
  return fm_sheet_update_table(wb, 0, 0, &patch);
}

std::string Text(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  return v.kind == FM_VALUE_KIND_TEXT ? std::string(v.u.text) : std::string();
}

std::string Formula(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  const char* text = nullptr;
  std::int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_formula(wb, 0, row, col, &present, &text), 0);
  return present != 0 && text != nullptr ? text : "";
}

std::string Describe(fm_workbook_t* wb, std::uint32_t row, std::uint32_t col) {
  fm_value_t v{};
  EXPECT_EQ(fm_sheet_get_value(wb, 0, row, col, &v), 0);
  std::string out = std::to_string(row) + "," + std::to_string(col) + " k" + std::to_string(v.kind);
  if (v.kind == FM_VALUE_KIND_TEXT) {
    out += " " + std::string(v.u.text);
  } else if (v.kind == FM_VALUE_KIND_NUMBER) {
    out += " " + std::to_string(v.u.number);
  }
  return out + " " + Formula(wb, row, col);
}

/// Everything a refused patch must leave as it was: each table's range and rows, the cells of A1:F14 and of the
/// grid's last two rows, and the merge count.
std::vector<std::string> Snapshot(fm_workbook_t* wb) {
  std::vector<std::string> out;
  std::size_t tables = 0;
  EXPECT_EQ(fm_sheet_get_table_count(wb, 0, &tables), 0);
  for (std::size_t i = 0; i < tables; ++i) {
    const fm_table_info_t info = Info(wb, i);
    out.push_back("table " + std::to_string(info.range.first_row) + "," + std::to_string(info.range.first_col) + ":" +
                  std::to_string(info.range.last_row) + "," + std::to_string(info.range.last_col) + " h" +
                  std::to_string(info.header_row_count) + " t" + std::to_string(info.totals_row_count));
  }
  for (std::uint32_t row = 0; row < 14U; ++row) {
    for (std::uint32_t col = 0; col < 6U; ++col) {
      out.push_back(Describe(wb, row, col));
    }
  }
  for (std::uint32_t col = 0; col < 6U; ++col) {
    out.push_back(Describe(wb, kLastRow - 1U, col));
    out.push_back(Describe(wb, kLastRow, col));
  }
  std::size_t merges = 0;
  EXPECT_EQ(fm_sheet_get_merge_count(wb, 0, &merges), 0);
  out.push_back("merges " + std::to_string(merges));
  return out;
}

/// Applies `patch` and expects `code` with nothing changed.
void ExpectRefusedUnchanged(WorkbookGuard& wb, const fm_table_patch_t& patch, fm_status_t code) {
  const std::vector<std::string> before = Snapshot(wb.handle);
  EXPECT_EQ(Update(wb.handle, patch), code) << fm_get_last_error_message();
  EXPECT_EQ(Snapshot(wb.handle), before);
}

void ExpectRange(fm_workbook_t* wb, fm_cell_rect_t want, std::size_t index = 0U) {
  const fm_table_info_t info = Info(wb, index);
  EXPECT_EQ(info.range.first_row, want.first_row);
  EXPECT_EQ(info.range.first_col, want.first_col);
  EXPECT_EQ(info.range.last_row, want.last_row);
  EXPECT_EQ(info.range.last_col, want.last_col);
}

// Range x toggle combinations, generated pairwise over this model:
//   toggle: header_off | header_on | totals_on | totals_off
//   above:  empty | occupied        (the row above the table when the patch runs)
//   range:  absent | present        (present = the post-toggle range grown by one row)
//   totals: absent | present        (the table's totals row before the patch)
//   IF toggle = totals_on THEN totals = absent; IF toggle = totals_off THEN totals = present
// A header_on case first hides the header of Sales; its row above is then the vacated header row.
struct Combo {
  const char* toggle;
  bool above_occupied;
  bool range;
  bool totals;
  fm_cell_rect_t want;
};

TEST(FormulonCApiTableToggleOrder, RangeAndTogglesComposeInOneSequence) {
  const Combo combos[] = {
      {"header_on", false, false, false, {2, 1, 4, 2}}, {"header_off", true, false, true, {3, 1, 5, 2}},
      {"totals_off", false, false, true, {2, 1, 4, 2}}, {"totals_off", true, true, true, {2, 1, 5, 2}},
      {"totals_off", false, true, true, {2, 1, 5, 2}},  {"totals_on", true, false, false, {2, 1, 5, 2}},
      {"totals_on", false, true, false, {2, 1, 6, 2}},  {"header_off", false, false, false, {3, 1, 4, 2}},
      {"header_on", true, true, true, {3, 1, 7, 2}},    {"header_off", true, true, true, {3, 1, 6, 2}},
  };
  for (const Combo& c : combos) {
    const std::string toggle = c.toggle;
    SCOPED_TRACE(toggle + (c.above_occupied ? " occupied" : " empty") + (c.range ? " range" : "") +
                 (c.totals ? " totals" : ""));
    WorkbookGuard wb;
    Create(wb, fm_cell_rect_t{2, 1, c.totals ? 5U : 4U, 2}, c.totals ? 1 : 0);
    if (toggle == "header_on") {
      ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0) << fm_get_last_error_message();
    }
    const std::uint32_t above = Info(wb.handle).range.first_row - 1U;
    if (c.above_occupied) {
      ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, above, 2, "X"), 0);
    }
    const std::int32_t header = toggle == "header_off" ? 0 : toggle == "header_on" ? 1 : -1;
    const std::int32_t totals = toggle == "totals_off" ? 0 : toggle == "totals_on" ? 1 : -1;
    ASSERT_EQ(Update(wb.handle, Patch(header, totals, c.range ? &c.want : nullptr)), 0) << fm_get_last_error_message();
    ExpectRange(wb.handle, c.want);
    const fm_table_info_t info = Info(wb.handle);
    const bool header_after = toggle != "header_off";
    const bool totals_after = toggle == "totals_on" || (c.totals && toggle != "totals_off");
    EXPECT_EQ(info.header_row_count, header_after ? 1U : 0U);
    EXPECT_EQ(info.totals_row_count, totals_after ? 1U : 0U);
    if (header_after) {
      EXPECT_EQ(Text(wb.handle, c.want.first_row, 1), "Product");
      EXPECT_EQ(Text(wb.handle, c.want.first_row, 2), "Amount");
    }
    EXPECT_EQ(Text(wb.handle, c.want.first_row + (header_after ? 1U : 0U), 1), "R1");
    if (c.above_occupied) {
      EXPECT_EQ(Text(wb.handle, above, 2), "X") << "the occupant stays put";
    }
  }
}

TEST(FormulonCApiTableToggleOrder, HeaderOffThenTotalsOnShiftsOnlyForTheTotalsRow) {
  WorkbookGuard wb;
  Create(wb);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 1, "x"), 0);
  ASSERT_EQ(Update(wb.handle, Patch(0, 1)), 0) << fm_get_last_error_message();
  ExpectRange(wb.handle, fm_cell_rect_t{3, 1, 5, 2});
  EXPECT_EQ(Text(wb.handle, 2, 1), "") << "the header cells are deleted";
  EXPECT_EQ(Text(wb.handle, 6, 1), "x") << "the occupied row below moved down for the totals row";
  EXPECT_EQ(Formula(wb.handle, 5, 2), "=SUBTOTAL(109,[Amount])");
}

TEST(FormulonCApiTableToggleOrder, HeaderOnByInsertThenTotalsOnShiftTwice) {
  WorkbookGuard wb;
  Create(wb);
  ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "X"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 6, 1, "below"), 0);
  ASSERT_EQ(Update(wb.handle, Patch(1, 1)), 0) << fm_get_last_error_message();
  ExpectRange(wb.handle, fm_cell_rect_t{3, 1, 6, 2});
  EXPECT_EQ(Text(wb.handle, 3, 1), "Product");
  EXPECT_EQ(Text(wb.handle, 4, 1), "R1");
  EXPECT_EQ(Formula(wb.handle, 6, 2), "=SUBTOTAL(109,[Amount])");
  EXPECT_EQ(Text(wb.handle, 8, 1), "below") << "one row for the header, one for the totals row";
}

TEST(FormulonCApiTableToggleOrder, TotalsDefaultsComeBeforeTheColumnPatch) {
  WorkbookGuard wb;
  Create(wb);
  const std::vector<TableColumnElement> columns = TableColumnsAt(1, [] {
    TableColumnElement element{};
    element.totals_row_function_op = FM_PATCH_OP_SET;
    element.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_MAX;
    return element;
  }());
  ASSERT_EQ(UpdateTableColumns(wb.handle, 0, 0, columns, Patch(-1, 1)), 0) << fm_get_last_error_message();
  EXPECT_EQ(Text(wb.handle, 5, 1), "Total") << "the default label of the first column stays";
  EXPECT_EQ(Formula(wb.handle, 5, 2), "=SUBTOTAL(104,[Amount])") << "the patch overrides the default function";
}

TEST(FormulonCApiTableToggleOrder, ACustomFunctionNeedsAFormulaOverTheDefaults) {
  WorkbookGuard wb;
  Create(wb);
  // Settings given while the totals row has never been shown give way to the defaults when it is shown.
  ASSERT_EQ(SetTableColumnTotals(wb.handle, 0, 0, 1, FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, "=1"), 0)
      << fm_get_last_error_message();
  TableColumnElement custom{};
  custom.totals_row_function_op = FM_PATCH_OP_SET;
  custom.totals_row_function = FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM;
  const std::vector<std::string> before = Snapshot(wb.handle);
  EXPECT_EQ(UpdateTableColumns(wb.handle, 0, 0, TableColumnsAt(1, custom), Patch(-1, 1)), kInvalidArgument);
  EXPECT_EQ(Snapshot(wb.handle), before);
  // Once shown, the stored formula carries a custom function through a hide and show.
  custom.totals_row_formula_op = FM_PATCH_OP_SET;
  custom.totals_row_formula = "=2";
  ASSERT_EQ(UpdateTableColumns(wb.handle, 0, 0, TableColumnsAt(1, custom), Patch(-1, 1)), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Formula(wb.handle, 5, 2), "=2");
  ASSERT_EQ(Update(wb.handle, Patch(-1, 0)), 0);
  custom.totals_row_formula_op = FM_PATCH_OP_KEEP;
  custom.totals_row_formula = nullptr;
  ASSERT_EQ(UpdateTableColumns(wb.handle, 0, 0, TableColumnsAt(1, custom), Patch(-1, 1)), 0)
      << fm_get_last_error_message();
  EXPECT_EQ(Formula(wb.handle, 5, 2), "=2");
}

TEST(FormulonCApiTableToggleOrder, TheRangeIsReadAgainstThePostToggleFirstRow) {
  {
    WorkbookGuard wb;
    Create(wb);
    const fm_cell_rect_t old_first{2, 1, 5, 2};
    ExpectRefusedUnchanged(wb, Patch(0, -1, &old_first), kInvalidArgument);
    const fm_cell_rect_t moved{3, 1, 5, 2};
    ASSERT_EQ(Update(wb.handle, Patch(0, -1, &moved)), 0) << fm_get_last_error_message();
    ExpectRange(wb.handle, moved);
  }
  {
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
    // Growing up: the range starts on the old first row minus one.
    const fm_cell_rect_t old_first{3, 1, 5, 2};
    ExpectRefusedUnchanged(wb, Patch(1, -1, &old_first), kInvalidArgument);
    const fm_cell_rect_t grown{2, 1, 5, 2};
    ASSERT_EQ(Update(wb.handle, Patch(1, -1, &grown)), 0) << fm_get_last_error_message();
    ExpectRange(wb.handle, grown);
  }
  {
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 1, "X"), 0);
    // Inserting: the range keeps the old first row.
    const fm_cell_rect_t up{2, 1, 6, 2};
    ExpectRefusedUnchanged(wb, Patch(1, -1, &up), kInvalidArgument);
    const fm_cell_rect_t kept{3, 1, 6, 2};
    ASSERT_EQ(Update(wb.handle, Patch(1, -1, &kept)), 0) << fm_get_last_error_message();
    ExpectRange(wb.handle, kept);
  }
}

TEST(FormulonCApiTableToggleOrder, TheRangeCheckSeesStructuresWhereTheToggleShiftsThem) {
  WorkbookGuard wb;
  Create(wb);
  ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "X"), 0);
  AddOther(wb, "Other", fm_cell_rect_t{8, 1, 9, 2});
  // The header insert moves Other from B9:C10 down to B10:C11 along with the table's columns.
  const fm_cell_rect_t overlapping{3, 1, 9, 2};
  ExpectRefusedUnchanged(wb, Patch(1, -1, &overlapping), FM_STATUS_PRECONDITION_FAILED);
  const fm_cell_rect_t clear{3, 1, 8, 2};
  ASSERT_EQ(Update(wb.handle, Patch(1, -1, &clear)), 0) << fm_get_last_error_message();
  ExpectRange(wb.handle, clear);
  ExpectRange(wb.handle, fm_cell_rect_t{9, 1, 10, 2}, 1U);
}

// Each refusal site the pre-check mirrors, reached through a patch that also toggles, leaves everything as it was.

TEST(FormulonCApiTableToggleOrder, HeaderToggleRefusalsChangeNothing) {
  {
    // Hiding the header of a table with no data row.
    WorkbookGuard wb;
    Create(wb, fm_cell_rect_t{2, 1, 3, 2}, 1);
    ExpectRefusedUnchanged(wb, Patch(0, -1), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // Showing it by an insert that pushes a value off the grid.
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "X"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, kLastRow, 2, "edge"), 0);
    ExpectRefusedUnchanged(wb, Patch(1, -1), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // Showing it by an insert whose band cuts another table.
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "X"), 0);
    AddOther(wb, "Other", fm_cell_rect_t{8, 0, 9, 1});
    const fm_cell_rect_t range{3, 1, 7, 2};
    ExpectRefusedUnchanged(wb, Patch(1, 1, &range), FM_STATUS_PRECONDITION_FAILED);
  }
}

TEST(FormulonCApiTableToggleOrder, TotalsToggleRefusalsChangeNothing) {
  {
    // Showing the totals row of a table on the grid's last row.
    WorkbookGuard wb;
    Create(wb, fm_cell_rect_t{kLastRow - 1U, 1, kLastRow, 2});
    ExpectRefusedUnchanged(wb, Patch(0, 1), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // Hiding the totals row of a table with no data row.
    WorkbookGuard wb;
    Create(wb, fm_cell_rect_t{2, 1, 3, 2}, 1);
    ExpectRefusedUnchanged(wb, Patch(-1, 0), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // A straddling merge below keeps the cells from shifting, and the row below is busy.
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 5, 1, "x"), 0);
    const fm_cell_rect_t merge{6, 2, 6, 3};
    ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0, &merge, nullptr), 0);
    ExpectRefusedUnchanged(wb, Patch(0, 1), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // The header grows without a shift, then the totals insert pushes a value off the grid.
    WorkbookGuard wb;
    Create(wb);
    ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, kLastRow, 2, "edge"), 0);
    ExpectRefusedUnchanged(wb, Patch(1, 1), FM_STATUS_PRECONDITION_FAILED);
  }
}

TEST(FormulonCApiTableToggleOrder, TwoInsertsTogetherPushOffWhatEachAloneKeeps) {
  WorkbookGuard wb;
  Create(wb);
  ASSERT_EQ(Update(wb.handle, Patch(0, -1)), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "X"), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, kLastRow - 1U, 2, "edge"), 0);
  ExpectRefusedUnchanged(wb, Patch(1, 1), FM_STATUS_PRECONDITION_FAILED);
  ASSERT_EQ(Update(wb.handle, Patch(1, -1)), 0) << fm_get_last_error_message();
  EXPECT_EQ(Text(wb.handle, kLastRow, 2), "edge");
}

TEST(FormulonCApiTableToggleOrder, RangeAndColumnPatchRefusalsAfterTogglesChangeNothing) {
  {
    // A range over another table, beside a header toggle.
    WorkbookGuard wb;
    Create(wb);
    AddOther(wb, "Other", fm_cell_rect_t{8, 1, 9, 2});
    const fm_cell_rect_t range{3, 1, 9, 2};
    ExpectRefusedUnchanged(wb, Patch(0, -1, &range), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // A range change while the table's creation shift is marked.
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 1, "Product"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 2, 2, "Amount"), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 3, 2, "under"), 0);
    const char* columns[] = {"Product", "Amount"};
    std::size_t index = 99U;
    const fm_cell_rect_t header_only{2, 1, 2, 2};
    ASSERT_EQ(AddTable(wb.handle, 0, &header_only, "Sales", columns, 2, "TableStyleMedium2", 1, 0, &index), 0)
        << fm_get_last_error_message();
    const fm_cell_rect_t range{2, 1, 5, 2};
    ExpectRefusedUnchanged(wb, Patch(-1, -1, &range), FM_STATUS_PRECONDITION_FAILED);
  }
  {
    // A custom totals formula that does not parse, beside both toggles.
    WorkbookGuard wb;
    Create(wb);
    const std::vector<std::string> before = Snapshot(wb.handle);
    EXPECT_NE(
        UpdateTableColumns(wb.handle, 0, 0,
                           TableColumnsAt(1, TableColumnTotals(FM_TABLE_TOTALS_ROW_FUNCTION_CUSTOM, nullptr, "=SUM((")),
                           Patch(0, 1)),
        0);
    EXPECT_EQ(Snapshot(wb.handle), before);
  }
  {
    // A calculated formula the column refuses, beside a totals toggle and a range.
    WorkbookGuard wb;
    Create(wb);
    const fm_cell_rect_t range{2, 1, 6, 2};
    const std::vector<std::string> before = Snapshot(wb.handle);
    EXPECT_EQ(UpdateTableColumns(wb.handle, 0, 0, TableColumnsAt(1, TableColumnCalculated("=SEQUENCE(3)")),
                                 Patch(-1, 1, &range)),
              kInvalidArgument);
    EXPECT_EQ(Snapshot(wb.handle), before);
  }
}

}  // namespace
