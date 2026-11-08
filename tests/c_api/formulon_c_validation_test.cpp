//
// Stable C ABI tests for data-validation evaluation: checking a proposed
// value against the covering rule, and paging through the invalid cells.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "formulon_c_out_index.h"
#include "gtest/gtest.h"
#include "utils/error.h"

namespace {

static_assert(sizeof(fm_validation_outcome_t) == 20U, "fm_validation_outcome_t ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr std::uint64_t kCursorEnd = std::numeric_limits<std::uint64_t>::max();
constexpr std::uint64_t kCursorStride = 16384U;

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

struct RangeGuard {
  fm_cell_range_t* handle = nullptr;
  ~RangeGuard() { fm_cell_range_destroy(handle); }
  RangeGuard() = default;
  RangeGuard(const RangeGuard&) = delete;
  RangeGuard& operator=(const RangeGuard&) = delete;
};

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0) << fm_get_last_error_message();
  const fm_status_t rc = fm_workbook_load(bytes, len, &out->handle);
  fm_buffer_destroy(bytes);
  ASSERT_EQ(rc, 0) << fm_get_last_error_message();
}

fm_value_t Number(double v) {
  fm_value_t value{};
  value.kind = FM_VALUE_KIND_NUMBER;
  value.u.number = v;
  return value;
}

fm_value_t Text(const char* text) {
  fm_value_t value{};
  value.kind = FM_VALUE_KIND_TEXT;
  value.u.text = text;
  return value;
}

/// Adds a whole-number rule over `range` (between `low` and `high`).
void AddWholeBetween(fm_workbook_t* wb, fm_cell_rect_t range, const char* low, const char* high,
                     std::uint8_t error_style) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = &range;
  v.range_count = 1;
  v.type = 1;
  v.op = 0;
  v.error_style = error_style;
  v.show_error_message = FM_TRISTATE_TRUE;
  v.formula1 = low;
  v.formula2 = high;
  ASSERT_EQ(fm_sheet_add_validation(wb, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
}

fm_validation_outcome_t Validate(const fm_workbook_t* wb, std::uint32_t row, std::uint32_t col, fm_value_t value) {
  fm_validation_outcome_t out = FM_STRUCT_INIT(fm_validation_outcome_t);
  EXPECT_EQ(fm_sheet_validate_value(wb, 0, row, col, &value, &out), 0) << fm_get_last_error_message();
  return out;
}

/// Lists every invalid cell, `limit` per page, as `(row, col)` pairs.
std::vector<std::pair<std::uint32_t, std::uint32_t>> ListInvalid(const fm_workbook_t* wb, std::uint32_t limit,
                                                                 std::size_t* pages) {
  std::vector<std::pair<std::uint32_t, std::uint32_t>> cells;
  std::uint64_t cursor = 0;
  *pages = 0;
  while (cursor != kCursorEnd && *pages < 16U) {
    RangeGuard page;
    EXPECT_EQ(fm_sheet_get_invalid_cells(wb, 0, cursor, limit, &page.handle), 0) << fm_get_last_error_message();
    if (page.handle == nullptr) {
      break;
    }
    ++*pages;
    std::size_t count = 0;
    EXPECT_EQ(fm_cell_range_get_cell_count(page.handle, &count), 0);
    EXPECT_LE(count, limit == 0U ? 65536U : limit);
    for (std::size_t i = 0; i < count; ++i) {
      std::uint32_t row = 0;
      std::uint32_t col = 0;
      fm_value_t value{};
      EXPECT_EQ(fm_cell_range_get_cell_at(page.handle, i, &row, &col, nullptr, &value), 0);
      cells.emplace_back(row, col);
    }
    EXPECT_EQ(fm_cell_range_next_cursor(page.handle, &cursor), 0);
  }
  return cells;
}

/// A2:A10 whole 1..10 (warning) over B2:B10 whole 5..6 (stop); A1 is free.
void BuildRules(fm_workbook_t* wb) {
  AddWholeBetween(wb, fm_cell_rect_t{1, 0, 9, 0}, "1", "10", 1);
  AddWholeBetween(wb, fm_cell_rect_t{1, 1, 9, 1}, "5", "6", 0);
}

void ExpectVerdicts(const fm_workbook_t* wb) {
  const fm_validation_outcome_t free_cell = Validate(wb, 0, 0, Number(99));
  EXPECT_EQ(free_cell.has_rule, 0);
  EXPECT_EQ(free_cell.valid, 1);

  const fm_validation_outcome_t ok = Validate(wb, 1, 0, Number(5));
  EXPECT_EQ(ok.has_rule, 1);
  EXPECT_EQ(ok.valid, 1);
  EXPECT_EQ(ok.rule_index, 0U);
  EXPECT_EQ(ok.error_style, 1);

  const fm_validation_outcome_t too_big = Validate(wb, 1, 0, Number(20));
  EXPECT_EQ(too_big.has_rule, 1);
  EXPECT_EQ(too_big.valid, 0);

  // Text is never a whole number, even when it spells one.
  EXPECT_EQ(Validate(wb, 1, 0, Text("5")).valid, 0);
  EXPECT_EQ(Validate(wb, 1, 0, Number(2.5)).valid, 0);

  const fm_validation_outcome_t second = Validate(wb, 4, 1, Number(7));
  EXPECT_EQ(second.has_rule, 1);
  EXPECT_EQ(second.valid, 0);
  EXPECT_EQ(second.rule_index, 1U);
  EXPECT_EQ(second.error_style, 0);
}

TEST(FormulonCApiValidation, ValidateValueUsesTheCoveringRule) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  BuildRules(wb.handle);
  ExpectVerdicts(wb.handle);
}

// The shared table is in `ST_DataValidationOperator` order; zero is `between`.
TEST(FormulonCApiValidation, ComparisonOperatorOrdinalsFollowTheSchemaOrder) {
  EXPECT_EQ(FM_COMPARISON_OPERATOR_BETWEEN, 0);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_NOT_BETWEEN, 1);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_EQUAL, 2);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_NOT_EQUAL, 3);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_LESS_THAN, 4);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL, 5);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_GREATER_THAN, 6);
  EXPECT_EQ(FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, 7);
}

// Each member of the shared table compares as it is named.
TEST(FormulonCApiValidation, EachComparisonOperatorComparesAsNamed) {
  const struct {
    int32_t op;
    bool four;
    bool five;
    bool six;
  } kCases[] = {
      {FM_COMPARISON_OPERATOR_BETWEEN, false, true, true},
      {FM_COMPARISON_OPERATOR_NOT_BETWEEN, true, false, false},
      {FM_COMPARISON_OPERATOR_EQUAL, false, true, false},
      {FM_COMPARISON_OPERATOR_NOT_EQUAL, true, false, true},
      {FM_COMPARISON_OPERATOR_LESS_THAN, true, false, false},
      {FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL, true, true, false},
      {FM_COMPARISON_OPERATOR_GREATER_THAN, false, false, true},
      {FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL, false, true, true},
  };
  for (const auto& c : kCases) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    const fm_cell_rect_t range{0, 0, 0, 0};
    fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
    v.ranges = &range;
    v.range_count = 1;
    v.type = FM_VALIDATION_TYPE_WHOLE;
    v.op = static_cast<std::uint8_t>(c.op);
    v.formula1 = "5";
    v.formula2 = "6";
    ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
    EXPECT_EQ(Validate(wb.handle, 0, 0, Number(4)).valid, c.four ? 1 : 0) << c.op;
    EXPECT_EQ(Validate(wb.handle, 0, 0, Number(5)).valid, c.five ? 1 : 0) << c.op;
    EXPECT_EQ(Validate(wb.handle, 0, 0, Number(6)).valid, c.six ? 1 : 0) << c.op;
  }
}

TEST(FormulonCApiValidation, OverlappingRulesApplyTheFirst) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  AddWholeBetween(wb.handle, fm_cell_rect_t{0, 0, 4, 0}, "1", "3", 0);
  AddWholeBetween(wb.handle, fm_cell_rect_t{0, 0, 9, 0}, "1", "100", 2);
  const fm_validation_outcome_t overlap = Validate(wb.handle, 2, 0, Number(50));
  EXPECT_EQ(overlap.rule_index, 0U);
  EXPECT_EQ(overlap.valid, 0);
  const fm_validation_outcome_t tail = Validate(wb.handle, 7, 0, Number(50));
  EXPECT_EQ(tail.rule_index, 1U);
  EXPECT_EQ(tail.valid, 1);
  EXPECT_EQ(tail.error_style, 2);
}

TEST(FormulonCApiValidation, InvalidCellsPageWithoutGaps) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  BuildRules(wb.handle);
  // A1 is outside every rule; A2, B3, A5 and B9 break theirs.
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 500), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 0, 50), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 1, 1, 5), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 2, 1, 9), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 3, 0, 3), 0);
  ASSERT_EQ(fm_sheet_set_text(wb.handle, 0, 4, 0, "x"), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 8, 1, 1), 0);

  const std::vector<std::pair<std::uint32_t, std::uint32_t>> expected = {{1, 0}, {2, 1}, {4, 0}, {8, 1}};
  std::size_t pages = 0;
  EXPECT_EQ(ListInvalid(wb.handle, 0, &pages), expected);
  EXPECT_EQ(pages, 1U);
  EXPECT_EQ(ListInvalid(wb.handle, 2, &pages), expected);
  EXPECT_EQ(pages, 2U);
  EXPECT_EQ(ListInvalid(wb.handle, 1, &pages), expected);
  EXPECT_EQ(pages, 4U);

  // A page that ends exactly at the last invalid cell reports no further one.
  RangeGuard last;
  ASSERT_EQ(fm_sheet_get_invalid_cells(wb.handle, 0, 4 * kCursorStride + 1, 1, &last.handle), 0);
  std::uint64_t next = 0;
  ASSERT_EQ(fm_cell_range_next_cursor(last.handle, &next), 0);
  EXPECT_EQ(next, kCursorEnd);
  std::size_t count = 0;
  ASSERT_EQ(fm_cell_range_get_cell_count(last.handle, &count), 0);
  EXPECT_EQ(count, 1U);

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  ExpectVerdicts(reloaded.handle);
  EXPECT_EQ(ListInvalid(reloaded.handle, 3, &pages), expected);
}

TEST(FormulonCApiValidation, NoRulesListsNothing) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_number(wb.handle, 0, 0, 0, 1), 0);
  RangeGuard page;
  ASSERT_EQ(fm_sheet_get_invalid_cells(wb.handle, 0, 0, 0, &page.handle), 0);
  std::size_t count = 9;
  ASSERT_EQ(fm_cell_range_get_cell_count(page.handle, &count), 0);
  EXPECT_EQ(count, 0U);
  std::uint64_t next = 0;
  ASSERT_EQ(fm_cell_range_next_cursor(page.handle, &next), 0);
  EXPECT_EQ(next, kCursorEnd);
}

/// A one-cell list rule over A1 with every tristate left at DEFAULT.
fm_data_validation_t DefaultListRule(const fm_cell_rect_t* range) {
  fm_data_validation_t v = FM_STRUCT_INIT(fm_data_validation_t);
  v.ranges = range;
  v.range_count = 1;
  v.type = 3;
  v.formula1 = "\"a,b\"";
  return v;
}

TEST(FormulonCApiValidation, UnsetTristatesReadBackAsExcelDefaults) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 0, 0};
  const fm_data_validation_t v = DefaultListRule(&range);
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();

  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, 0, &got), 0);
  EXPECT_EQ(got.allow_blank, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_in_cell_dropdown, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_input_message, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_error_message, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiValidation, ExplicitFalseTristatesRoundTripThroughSave) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 0, 0};
  fm_data_validation_t v = DefaultListRule(&range);
  v.allow_blank = FM_TRISTATE_FALSE;
  v.show_in_cell_dropdown = FM_TRISTATE_FALSE;
  v.show_input_message = FM_TRISTATE_FALSE;
  v.show_error_message = FM_TRISTATE_FALSE;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation(reloaded.handle, 0, 0, &got), 0);
  EXPECT_EQ(got.allow_blank, FM_TRISTATE_FALSE);
  EXPECT_EQ(got.show_in_cell_dropdown, FM_TRISTATE_FALSE);
  EXPECT_EQ(got.show_input_message, FM_TRISTATE_FALSE);
  EXPECT_EQ(got.show_error_message, FM_TRISTATE_FALSE);
}

TEST(FormulonCApiValidation, ExplicitTrueTristatesReadBackAsTrue) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 0, 0};
  fm_data_validation_t v = DefaultListRule(&range);
  v.allow_blank = FM_TRISTATE_TRUE;
  v.show_in_cell_dropdown = FM_TRISTATE_TRUE;
  ASSERT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), 0) << fm_get_last_error_message();
  fm_data_validation_t got = FM_STRUCT_INIT(fm_data_validation_t);
  ASSERT_EQ(fm_sheet_get_validation(wb.handle, 0, 0, &got), 0);
  EXPECT_EQ(got.allow_blank, FM_TRISTATE_TRUE);
  EXPECT_EQ(got.show_in_cell_dropdown, FM_TRISTATE_TRUE);
}

TEST(FormulonCApiValidation, TristateOutsideDomainIsRejectedWithoutStoringTheRule) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  const fm_cell_rect_t range{0, 0, 0, 0};
  for (int field = 0; field < 4; ++field) {
    fm_data_validation_t v = DefaultListRule(&range);
    int32_t* slots[] = {&v.allow_blank, &v.show_in_cell_dropdown, &v.show_input_message, &v.show_error_message};
    *slots[field] = 3;
    EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), kInvalidArgument) << "field " << field;
    *slots[field] = -1;
    EXPECT_EQ(fm_sheet_add_validation(wb.handle, 0, &v, ignored_index()), kInvalidArgument) << "field " << field;
  }
  std::size_t count = 9;
  ASSERT_EQ(fm_sheet_get_validation_count(wb.handle, 0, &count), 0);
  EXPECT_EQ(count, 0U);
}

TEST(FormulonCApiValidation, RejectsBadArguments) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  fm_validation_outcome_t out = FM_STRUCT_INIT(fm_validation_outcome_t);
  fm_value_t value = Number(1);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 3, 0, 0, &value, &out), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 0, 0, nullptr, &out), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 0, 0, &value, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 2000000, 0, &value, &out), kInvalidArgument);
  value.kind = FM_VALUE_KIND_ARRAY;
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 0, 0, &value, &out), kInvalidArgument);
  value = Text(nullptr);
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 0, 0, &value, &out), kBindingNullPointer);
  value = Number(std::numeric_limits<double>::infinity());
  EXPECT_EQ(fm_sheet_validate_value(wb.handle, 0, 0, 0, &value, &out), kInvalidArgument);

  RangeGuard page;
  EXPECT_EQ(fm_sheet_get_invalid_cells(wb.handle, 0, 0, 0, nullptr), kBindingNullPointer);
  EXPECT_EQ(fm_sheet_get_invalid_cells(wb.handle, 0, 1048576ULL * kCursorStride, 0, &page.handle), kInvalidArgument);
  EXPECT_EQ(page.handle, nullptr);
  EXPECT_EQ(fm_sheet_get_invalid_cells(wb.handle, 2, 0, 0, &page.handle), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
}

}  // namespace
