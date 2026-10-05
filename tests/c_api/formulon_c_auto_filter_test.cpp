//
// Stable C ABI tests for the typed sheet and table AutoFilter: the record
// shapes, get/set/remove, apply/clear/evaluate over the filtered rows, and
// the state surviving a save and reload.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <set>
#include <string>
#include <vector>

#include "auto_filter.h"
#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "gtest/gtest.h"
#include "io/auto_filter_xml.h"
#include "utils/error.h"

namespace {

constexpr bool kWasm32 = sizeof(void*) == 4U;
static_assert(sizeof(fm_date_group_item) == 8U, "fm_date_group_item ABI layout changed");
static_assert(sizeof(fm_filter_column) == (kWasm32 ? 152U : 192U), "fm_filter_column ABI layout changed");
static_assert(sizeof(fm_sort_condition) == (kWasm32 ? 48U : 56U), "fm_sort_condition ABI layout changed");
static_assert(sizeof(fm_auto_filter) == (kWasm32 ? 64U : 80U), "fm_auto_filter ABI layout changed");

constexpr fm_status_t kInvalidArgument = static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
constexpr fm_status_t kNotFound = static_cast<fm_status_t>(formulon::FormulonErrorCode::kNotFound);
constexpr fm_status_t kAutoFilterInvalid = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAutoFilterInvalid);
constexpr fm_status_t kBindingNullPointer = static_cast<fm_status_t>(formulon::FormulonErrorCode::kBindingNullPointer);
constexpr std::uint32_t kAbsentDxf = std::numeric_limits<std::uint32_t>::max();

struct WorkbookGuard {
  fm_workbook_t* handle = nullptr;
  ~WorkbookGuard() { fm_workbook_destroy(handle); }
  WorkbookGuard() = default;
  WorkbookGuard(const WorkbookGuard&) = delete;
  WorkbookGuard& operator=(const WorkbookGuard&) = delete;
};

void SaveAndReload(fm_workbook_t* wb, WorkbookGuard* out) {
  std::uint8_t* bytes = nullptr;
  std::size_t len = 0;
  ASSERT_EQ(fm_workbook_save(wb, &bytes, &len), 0) << fm_last_error_message();
  const fm_status_t rc = fm_workbook_load(bytes, len, &out->handle);
  fm_buffer_free(bytes);
  ASSERT_EQ(rc, 0) << fm_last_error_message();
}

std::set<std::uint32_t> HiddenRows(const fm_workbook_t* wb) {
  std::set<std::uint32_t> rows;
  std::size_t count = 0;
  EXPECT_EQ(fm_sheet_get_row_override_count(wb, 0, &count), 0);
  for (std::size_t i = 0; i < count; ++i) {
    fm_row_layout_t layout{};
    EXPECT_EQ(fm_sheet_get_row_override(wb, 0, i, &layout), 0);
    if (layout.hidden != 0) {
      rows.insert(layout.row);
    }
  }
  return rows;
}

bool HasDefinedName(const fm_workbook_t* wb, const std::string& wanted) {
  const std::size_t count = fm_workbook_defined_name_count(wb);
  for (std::size_t i = 0; i < count; ++i) {
    const char* name = nullptr;
    const char* formula = nullptr;
    int32_t local = 0;
    EXPECT_EQ(fm_workbook_defined_name_at(wb, i, &name, &formula, &local), 0);
    if (wanted == name) {
      return true;
    }
  }
  return false;
}

/// Header "N" in A1 and the numbers 1..5 in A2:A6.
void FillNumbers(fm_workbook_t* wb) {
  ASSERT_EQ(fm_workbook_set_text(wb, 0, 0, 0, "N"), 0);
  for (std::uint32_t i = 1; i <= 5; ++i) {
    ASSERT_EQ(fm_workbook_set_number(wb, 0, i, 0, static_cast<double>(i)), 0);
  }
}

fm_filter_column ValuesColumn(std::uint32_t col_id, const char* const* values, std::uint32_t count) {
  fm_filter_column c{};
  c.col_id = col_id;
  c.show_button = 1;
  c.kind = 1;
  c.values = values;
  c.value_count = count;
  return c;
}

/// Copy of an AutoFilter record whose views may die with the next read.
struct ColumnCopy {
  fm_filter_column record{};
  std::vector<std::string> values;
  std::vector<fm_date_group_item> date_groups;
  std::string val1, val2, val_iso, max_val_iso;
};

struct FilterCopy {
  fm_auto_filter record{};
  std::vector<ColumnCopy> columns;
  std::vector<fm_sort_condition> conditions;
  std::vector<std::string> custom_lists;
};

std::string Str(const char* text) {
  return text != nullptr ? std::string(text) : std::string();
}

FilterCopy Copy(const fm_auto_filter& in) {
  FilterCopy out;
  out.record = in;
  for (std::uint32_t i = 0; i < in.column_count; ++i) {
    const fm_filter_column& c = in.columns[i];
    ColumnCopy col;
    col.record = c;
    for (std::uint32_t v = 0; v < c.value_count; ++v) {
      col.values.push_back(Str(c.values[v]));
    }
    col.date_groups.assign(c.date_groups, c.date_groups + c.date_group_count);
    col.val1 = Str(c.val1);
    col.val2 = Str(c.val2);
    col.val_iso = Str(c.val_iso);
    col.max_val_iso = Str(c.max_val_iso);
    out.columns.push_back(col);
  }
  for (std::uint32_t i = 0; i < in.condition_count; ++i) {
    out.conditions.push_back(in.conditions[i]);
    out.custom_lists.push_back(Str(in.conditions[i].custom_list));
  }
  return out;
}

FilterCopy GetSheetFilter(const fm_workbook_t* wb) {
  fm_auto_filter out{};
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_auto_filter(wb, 0, &out, &present), 0) << fm_last_error_message();
  EXPECT_EQ(present, 1);
  return Copy(out);
}

/// One column of every criterion kind plus a two-key sort state.
class EveryKind {
 public:
  EveryKind() {
    columns_[0] = ValuesColumn(0, values_, 2);
    columns_[0].filter_blank = 1;
    groups_[0] = fm_date_group_item{2024, 3, 0, 0, 0, 0, 1};
    columns_[0].date_groups = groups_;
    columns_[0].date_group_count = 1;

    columns_[1].col_id = 1;
    columns_[1].show_button = 1;
    columns_[1].kind = 2;
    columns_[1].custom_and = 1;
    columns_[1].custom_count = 2;
    columns_[1].op1 = 4;
    columns_[1].val1 = "10";
    columns_[1].op2 = 3;
    columns_[1].val2 = "a*";

    columns_[2].col_id = 2;
    columns_[2].show_button = 1;
    columns_[2].kind = 3;
    columns_[2].top = 0;
    columns_[2].percent = 1;
    columns_[2].top_val = 25.0;
    columns_[2].has_filter_val = 1;
    columns_[2].filter_val = 3.5;

    columns_[3].col_id = 3;
    columns_[3].show_button = 1;
    columns_[3].kind = 4;
    columns_[3].dynamic_type = 1;
    columns_[3].has_dyn_val = 1;
    columns_[3].dyn_val = 12.5;

    columns_[4].col_id = 4;
    columns_[4].show_button = 1;
    columns_[4].kind = 5;
    columns_[4].dxf_id = kAbsentDxf;
    columns_[4].cell_color = 0;

    columns_[5].col_id = 5;
    columns_[5].hidden_button = 1;
    columns_[5].show_button = 1;
    columns_[5].kind = 6;
    columns_[5].icon_set = 13;
    columns_[5].icon_id = 2;
    columns_[5].has_icon_id = 1;

    conditions_[0].ref = fm_merge_range{1, 0, 19, 0};
    conditions_[0].descending = 1;
    conditions_[1].ref = fm_merge_range{1, 1, 19, 1};
    conditions_[1].sort_by = 3;
    conditions_[1].icon_set = 4;
    conditions_[1].icon_id = 1;
    conditions_[1].has_icon_id = 1;
    conditions_[1].custom_list = "";

    filter_.range = fm_merge_range{0, 0, 19, 6};
    filter_.columns = columns_;
    filter_.column_count = 6;
    filter_.has_sort = 1;
    filter_.sort_ref = fm_merge_range{1, 0, 19, 6};
    filter_.case_sensitive = 1;
    filter_.conditions = conditions_;
    filter_.condition_count = 2;
  }

  const fm_auto_filter& record() const { return filter_; }

 private:
  const char* values_[2] = {"x", "10"};
  fm_date_group_item groups_[1]{};
  fm_filter_column columns_[6]{};
  fm_sort_condition conditions_[2]{};
  fm_auto_filter filter_{};
};

void ExpectEveryKind(const FilterCopy& got) {
  EXPECT_EQ(got.record.range.first_row, 0U);
  EXPECT_EQ(got.record.range.last_row, 19U);
  EXPECT_EQ(got.record.range.last_col, 6U);
  ASSERT_EQ(got.columns.size(), 6U);

  const ColumnCopy& values = got.columns[0];
  EXPECT_EQ(values.record.kind, 1);
  EXPECT_EQ(values.record.filter_blank, 1);
  EXPECT_EQ(values.values, (std::vector<std::string>{"x", "10"}));
  ASSERT_EQ(values.date_groups.size(), 1U);
  EXPECT_EQ(values.date_groups[0].year, 2024);
  EXPECT_EQ(values.date_groups[0].month, 3);
  EXPECT_EQ(values.date_groups[0].grouping, 1);

  const ColumnCopy& custom = got.columns[1];
  EXPECT_EQ(custom.record.kind, 2);
  EXPECT_EQ(custom.record.col_id, 1U);
  EXPECT_EQ(custom.record.custom_and, 1);
  EXPECT_EQ(custom.record.custom_count, 2);
  EXPECT_EQ(custom.record.op1, 4);
  EXPECT_EQ(custom.val1, "10");
  EXPECT_EQ(custom.record.op2, 3);
  EXPECT_EQ(custom.val2, "a*");

  const ColumnCopy& top = got.columns[2];
  EXPECT_EQ(top.record.kind, 3);
  EXPECT_EQ(top.record.top, 0);
  EXPECT_EQ(top.record.percent, 1);
  EXPECT_DOUBLE_EQ(top.record.top_val, 25.0);
  EXPECT_EQ(top.record.has_filter_val, 1);
  EXPECT_DOUBLE_EQ(top.record.filter_val, 3.5);

  const ColumnCopy& dynamic = got.columns[3];
  EXPECT_EQ(dynamic.record.kind, 4);
  EXPECT_EQ(dynamic.record.dynamic_type, 1);
  EXPECT_EQ(dynamic.record.has_dyn_val, 1);
  EXPECT_DOUBLE_EQ(dynamic.record.dyn_val, 12.5);
  EXPECT_EQ(dynamic.record.has_dyn_max_val, 0);

  const ColumnCopy& color = got.columns[4];
  EXPECT_EQ(color.record.kind, 5);
  EXPECT_EQ(color.record.dxf_id, kAbsentDxf);
  EXPECT_EQ(color.record.cell_color, 0);

  const ColumnCopy& icon = got.columns[5];
  EXPECT_EQ(icon.record.kind, 6);
  EXPECT_EQ(icon.record.hidden_button, 1);
  EXPECT_EQ(icon.record.icon_set, 13);
  EXPECT_EQ(icon.record.has_icon_id, 1);
  EXPECT_EQ(icon.record.icon_id, 2);

  EXPECT_EQ(got.record.has_sort, 1);
  EXPECT_EQ(got.record.sort_ref.first_row, 1U);
  EXPECT_EQ(got.record.case_sensitive, 1);
  ASSERT_EQ(got.conditions.size(), 2U);
  EXPECT_EQ(got.conditions[0].descending, 1);
  EXPECT_EQ(got.conditions[0].ref.last_row, 19U);
  EXPECT_EQ(got.conditions[1].sort_by, 3);
  EXPECT_EQ(got.conditions[1].icon_set, 4);
  EXPECT_EQ(got.conditions[1].has_icon_id, 1);
  EXPECT_EQ(got.conditions[1].icon_id, 1);
}

TEST(FormulonCApiAutoFilter, AbsentFilterReadsAsNotPresent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  fm_auto_filter out{};
  out.column_count = 7U;
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(out.column_count, 0U);
  EXPECT_EQ(out.columns, nullptr);

  std::size_t len = 0;
  std::uint32_t first_row = 0;
  EXPECT_EQ(fm_sheet_apply_auto_filter(wb.handle, 0), kNotFound);
  EXPECT_EQ(fm_sheet_clear_auto_filter(wb.handle, 0), kNotFound);
  EXPECT_EQ(fm_sheet_evaluate_auto_filter(wb.handle, 0, nullptr, 0, &len, &first_row), kNotFound);
  EXPECT_EQ(fm_sheet_remove_auto_filter(wb.handle, 0), 0);
  EXPECT_EQ(fm_sheet_get_auto_filter(wb.handle, 5, &out, &present), kInvalidArgument);
  EXPECT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, nullptr, &present), kBindingNullPointer);
}

TEST(FormulonCApiAutoFilter, EveryCriterionKindRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const EveryKind every;
  ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &every.record()), 0) << fm_last_error_message();
  ExpectEveryKind(GetSheetFilter(wb.handle));
  EXPECT_TRUE(HasDefinedName(wb.handle, "_xlnm._FilterDatabase"));

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  ExpectEveryKind(GetSheetFilter(reloaded.handle));

  ASSERT_EQ(fm_sheet_remove_auto_filter(reloaded.handle, 0), 0);
  fm_auto_filter out{};
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_auto_filter(reloaded.handle, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);
  EXPECT_FALSE(HasDefinedName(reloaded.handle, "_xlnm._FilterDatabase"));
}

TEST(FormulonCApiAutoFilter, ApplyHidesRowsAndClearShowsThem) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillNumbers(wb.handle);
  const char* values[] = {"2", "4"};
  const fm_filter_column column = ValuesColumn(0, values, 2);
  fm_auto_filter filter{};
  filter.range = fm_merge_range{0, 0, 5, 0};
  filter.columns = &column;
  filter.column_count = 1;
  ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), 0) << fm_last_error_message();
  EXPECT_TRUE(HiddenRows(wb.handle).empty());

  std::size_t len = 0;
  std::uint32_t first_row = 0;
  ASSERT_EQ(fm_sheet_evaluate_auto_filter(wb.handle, 0, nullptr, 0, &len, &first_row), 0);
  EXPECT_EQ(len, 5U);
  EXPECT_EQ(first_row, 1U);
  std::vector<std::uint8_t> match(len, 9U);
  ASSERT_EQ(fm_sheet_evaluate_auto_filter(wb.handle, 0, match.data(), match.size(), &len, &first_row), 0);
  EXPECT_EQ(match, (std::vector<std::uint8_t>{0, 1, 0, 1, 0}));
  EXPECT_EQ(fm_sheet_evaluate_auto_filter(wb.handle, 0, nullptr, 1, &len, &first_row), kBindingNullPointer);

  ASSERT_EQ(fm_sheet_apply_auto_filter(wb.handle, 0), 0) << fm_last_error_message();
  EXPECT_EQ(HiddenRows(wb.handle), (std::set<std::uint32_t>{1, 3, 5}));

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  EXPECT_EQ(HiddenRows(reloaded.handle), (std::set<std::uint32_t>{1, 3, 5}));
  const FilterCopy kept = GetSheetFilter(reloaded.handle);
  ASSERT_EQ(kept.columns.size(), 1U);
  EXPECT_EQ(kept.columns[0].values, (std::vector<std::string>{"2", "4"}));

  ASSERT_EQ(fm_sheet_clear_auto_filter(reloaded.handle, 0), 0) << fm_last_error_message();
  EXPECT_TRUE(HiddenRows(reloaded.handle).empty());
  const FilterCopy cleared = GetSheetFilter(reloaded.handle);
  EXPECT_EQ(cleared.columns.size(), 0U);
  EXPECT_EQ(cleared.record.range.last_row, 5U);
}

TEST(FormulonCApiAutoFilter, ClearKeepsColumnsThatHideTheirButton) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillNumbers(wb.handle);
  const char* values[] = {"3"};
  fm_filter_column column = ValuesColumn(0, values, 1);
  column.hidden_button = 1;
  fm_auto_filter filter{};
  filter.range = fm_merge_range{0, 0, 5, 0};
  filter.columns = &column;
  filter.column_count = 1;
  ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), 0);
  ASSERT_EQ(fm_sheet_apply_auto_filter(wb.handle, 0), 0);
  EXPECT_EQ(HiddenRows(wb.handle).size(), 4U);
  ASSERT_EQ(fm_sheet_clear_auto_filter(wb.handle, 0), 0);
  EXPECT_TRUE(HiddenRows(wb.handle).empty());
  const FilterCopy cleared = GetSheetFilter(wb.handle);
  ASSERT_EQ(cleared.columns.size(), 1U);
  EXPECT_EQ(cleared.columns[0].record.kind, 0);
  EXPECT_EQ(cleared.columns[0].record.hidden_button, 1);
}

TEST(FormulonCApiAutoFilter, RejectsMalformedRecords) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* values[] = {"1"};
  fm_filter_column column = ValuesColumn(3, values, 1);
  fm_auto_filter filter{};
  filter.range = fm_merge_range{0, 0, 5, 1};
  filter.columns = &column;
  filter.column_count = 1;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kAutoFilterInvalid);

  column.col_id = 0;
  column.kind = 7;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kInvalidArgument);

  column.kind = 1;
  column.values = nullptr;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kInvalidArgument);

  column.kind = 2;
  column.custom_count = 3;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kAutoFilterInvalid);
  column.custom_count = 1;
  column.op1 = 6;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kInvalidArgument);

  column.op1 = 0;
  filter.has_sort = 1;
  filter.sort_method = 3;
  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, &filter), kInvalidArgument);

  EXPECT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, nullptr), kBindingNullPointer);
  fm_auto_filter out{};
  int32_t present = 1;
  ASSERT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);
}

TEST(FormulonCApiAutoFilter, OpaqueFilterStaysOnTheXmlView) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  const char* xml =
      "<autoFilter ref=\"A1:B5\"><filterColumn colId=\"0\"><customFilters>"
      "<customFilter operator=\"bogus\" val=\"1\"/></customFilters></filterColumn></autoFilter>";
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, xml), 0) << fm_last_error_message();
  fm_auto_filter out{};
  int32_t present = 0;
  EXPECT_EQ(fm_sheet_get_auto_filter(wb.handle, 0, &out, &present), kAutoFilterInvalid);
  EXPECT_EQ(fm_sheet_apply_auto_filter(wb.handle, 0), kAutoFilterInvalid);
  const char* view = nullptr;
  ASSERT_EQ(fm_sheet_get_auto_filter_xml(wb.handle, 0, &view), 0);
  EXPECT_STREQ(view, xml);
}

TEST(FormulonCApiAutoFilter, TableFilterAppliesAndRoundTrips) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  FillNumbers(wb.handle);
  const char* names[] = {"N"};
  std::size_t table = 0;
  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "A1:A6", "Nums", "Nums", names, 1, nullptr, 1, 0, &table), 0)
      << fm_last_error_message();

  fm_auto_filter out{};
  int32_t present = 0;
  ASSERT_EQ(fm_table_get_auto_filter(wb.handle, table, &out, &present), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(out.range.last_row, 5U);
  EXPECT_EQ(out.column_count, 0U);

  fm_filter_column column{};
  column.kind = 2;
  column.show_button = 1;
  column.custom_count = 1;
  column.op1 = 5;
  column.val1 = "3";
  fm_auto_filter filter{};
  filter.range = fm_merge_range{0, 0, 5, 0};
  filter.columns = &column;
  filter.column_count = 1;
  ASSERT_EQ(fm_table_set_auto_filter(wb.handle, table, &filter), 0) << fm_last_error_message();
  EXPECT_FALSE(HasDefinedName(wb.handle, "_xlnm._FilterDatabase"));

  std::vector<std::uint8_t> match(5U);
  std::size_t len = 0;
  std::uint32_t first_row = 0;
  ASSERT_EQ(fm_table_evaluate_auto_filter(wb.handle, table, match.data(), match.size(), &len, &first_row), 0);
  EXPECT_EQ(match, (std::vector<std::uint8_t>{0, 0, 0, 1, 1}));
  ASSERT_EQ(fm_table_apply_auto_filter(wb.handle, table), 0) << fm_last_error_message();
  EXPECT_EQ(HiddenRows(wb.handle), (std::set<std::uint32_t>{1, 2, 3}));

  WorkbookGuard reloaded;
  SaveAndReload(wb.handle, &reloaded);
  ASSERT_EQ(fm_table_get_auto_filter(reloaded.handle, 0, &out, &present), 0) << fm_last_error_message();
  EXPECT_EQ(present, 1);
  ASSERT_EQ(out.column_count, 1U);
  EXPECT_EQ(out.columns[0].kind, 2);
  EXPECT_EQ(out.columns[0].op1, 5);
  EXPECT_STREQ(out.columns[0].val1, "3");
  EXPECT_EQ(HiddenRows(reloaded.handle), (std::set<std::uint32_t>{1, 2, 3}));

  ASSERT_EQ(fm_table_clear_auto_filter(reloaded.handle, 0), 0);
  EXPECT_TRUE(HiddenRows(reloaded.handle).empty());
  ASSERT_EQ(fm_table_remove_auto_filter(reloaded.handle, 0), 0);
  ASSERT_EQ(fm_table_get_auto_filter(reloaded.handle, 0, &out, &present), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(fm_table_apply_auto_filter(reloaded.handle, 0), kNotFound);
  EXPECT_EQ(fm_table_get_auto_filter(reloaded.handle, 4, &out, &present), kInvalidArgument);
  EXPECT_EQ(fm_table_set_auto_filter(reloaded.handle, 4, &filter), kInvalidArgument);
}

/// Unmodelled content on every level: an `xr:uid` and an unknown attribute,
/// a `calendarType`, an unknown child, column and filter extensions, and a
/// sort state with an attribute and extension of its own.
constexpr const char* kUnmodelledXml =
    "<autoFilter ref=\"A1:C9\" xr:uid=\"{00000000-0000-0000-0000-000000000001}\" x:future=\"1\">"
    "<filterColumn colId=\"0\" x:col=\"2\"><filters calendarType=\"japan\"><filter val=\"x\"/></filters>"
    "<futureChild a=\"1\"/><extLst><ext uri=\"{00000000-0000-0000-0000-00000000FFFF}\"><y:other/></ext></extLst>"
    "</filterColumn>"
    "<filterColumn colId=\"2\"><customFilters><customFilter operator=\"greaterThan\" val=\"3\"/>"
    "</customFilters></filterColumn>"
    "<sortState ref=\"A2:C9\" x:sort=\"3\"><sortCondition ref=\"A2:A9\"/>"
    "<extLst><ext uri=\"{00000000-0000-0000-0000-00000000DDDD}\"/></extLst></sortState>"
    "<extLst><ext uri=\"{00000000-0000-0000-0000-00000000EEEE}\"/></extLst></autoFilter>";

/// A settable record rebuilt over the owned strings and arrays of `copy`,
/// which must outlive it.
class OwnedRecord {
 public:
  explicit OwnedRecord(const FilterCopy& copy) : conditions_(copy.conditions), record_(copy.record) {
    values_.resize(copy.columns.size());
    for (std::size_t i = 0; i < copy.columns.size(); ++i) {
      const ColumnCopy& c = copy.columns[i];
      for (const std::string& v : c.values) {
        values_[i].push_back(v.c_str());
      }
      fm_filter_column column = c.record;
      column.values = values_[i].empty() ? nullptr : values_[i].data();
      column.date_groups = c.date_groups.empty() ? nullptr : c.date_groups.data();
      column.val1 = c.val1.c_str();
      column.val2 = c.val2.c_str();
      column.val_iso = c.val_iso.c_str();
      column.max_val_iso = c.max_val_iso.c_str();
      columns_.push_back(column);
    }
    for (std::size_t i = 0; i < conditions_.size(); ++i) {
      conditions_[i].custom_list = copy.custom_lists[i].c_str();
    }
    record_.columns = columns_.empty() ? nullptr : columns_.data();
    record_.conditions = conditions_.empty() ? nullptr : conditions_.data();
  }

  const fm_auto_filter* get() const { return &record_; }

 private:
  std::vector<std::vector<const char*>> values_;
  std::vector<fm_filter_column> columns_;
  std::vector<fm_sort_condition> conditions_;
  fm_auto_filter record_{};
};

TEST(FormulonCApiAutoFilter, GetThenSetKeepsUnmodelledContent) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  ASSERT_EQ(fm_sheet_set_auto_filter_xml(wb.handle, 0, kUnmodelledXml), 0) << fm_last_error_message();
  const std::string before = formulon::io::auto_filter_xml(wb.handle->workbook().sheet(0).auto_filter());
  ASSERT_NE(before.find("calendarType"), std::string::npos);
  ASSERT_NE(before.find("x:future"), std::string::npos);
  const FilterCopy copy = GetSheetFilter(wb.handle);
  const OwnedRecord record(copy);
  ASSERT_EQ(fm_sheet_set_auto_filter(wb.handle, 0, record.get()), 0) << fm_last_error_message();
  EXPECT_EQ(formulon::io::auto_filter_xml(wb.handle->workbook().sheet(0).auto_filter()), before);

  const char* names[] = {"A", "B", "C"};
  std::size_t table = 0;
  ASSERT_EQ(fm_workbook_table_create(wb.handle, 0, "E1:G9", "T", "T", names, 3, nullptr, 1, 0, &table), 0);
  std::string table_xml = kUnmodelledXml;
  table_xml.replace(table_xml.find("A1:C9"), 5, "E1:G9");
  table_xml.replace(table_xml.find("A2:C9"), 5, "E2:G9");
  table_xml.replace(table_xml.find("A2:A9"), 5, "E2:E9");
  wb.handle->workbook().mutable_tables()[table].auto_filter_xml.set(formulon::io::auto_filter_from_xml(table_xml));
  const std::string table_before =
      formulon::io::auto_filter_xml(wb.handle->workbook().tables()[table].auto_filter_xml.get());
  fm_auto_filter out{};
  int32_t present = 0;
  ASSERT_EQ(fm_table_get_auto_filter(wb.handle, table, &out, &present), 0) << fm_last_error_message();
  ASSERT_EQ(present, 1);
  const FilterCopy table_copy = Copy(out);
  const OwnedRecord table_record(table_copy);
  ASSERT_EQ(fm_table_set_auto_filter(wb.handle, table, table_record.get()), 0) << fm_last_error_message();
  EXPECT_EQ(formulon::io::auto_filter_xml(wb.handle->workbook().tables()[table].auto_filter_xml.get()), table_before);
}

}  // namespace
