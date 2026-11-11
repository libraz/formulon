// Stable C ABI PivotTable grouping and item-binding tests.

#include <cstddef>
#include <cstdint>
#include <limits>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"
#include "value.h"

namespace {

fm_status_t BuildDaysGroupedPivot(fm_workbook_t* wb, double start_serial, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  fm_status_t st = fm_workbook_add_pivot_cache(wb, 0, &cache_id);
  if (st != 0) {
    return st;
  }
  st = SetPivotCacheWorksheetSource(wb, cache_id, /*present=*/1, "A1:B71", "Sheet1", nullptr);
  if (st != 0) {
    return st;
  }
  std::size_t date_index = 99;
  st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Date", &date_index);
  if (st != 0) {
    return st;
  }
  std::size_t amount_index = 99;
  st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Amount", &amount_index);
  if (st != 0) {
    return st;
  }
  for (double serial = 0.0; serial <= 70.0; serial += 1.0) {
    std::size_t rec_index = 99;
    st = fm_workbook_add_pivot_cache_record(wb, cache_id, &rec_index);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_set_pivot_cache_record_date(wb, cache_id, rec_index, date_index, serial);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, amount_index, 1.0);
    if (st != 0) {
      return st;
    }
  }
  std::size_t pivot_index = 99;
  st = AddPivot(wb, 0, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot_index);
  if (st != 0) {
    return st;
  }
  // The pivot carries one field per cache field, in cache order.
  const std::size_t date_field = date_index;
  st = fm_sheet_set_pivot_field_axis(wb, 0, pivot_index, date_field, FM_PIVOT_AXIS_ROW);
  if (st != 0) {
    return st;
  }
  fm_pivot_field_patch_t date_patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  date_patch.subtotal_top_op = FM_PATCH_OP_SET;
  date_patch.subtotal_top = 0;
  st = fm_sheet_update_pivot_field(wb, 0, pivot_index, date_field, &date_patch);
  if (st != 0) {
    return st;
  }
  if (st != 0) {
    return st;
  }
  // The pivot carries one field per cache field, in cache order.
  const std::size_t amount_field = amount_index;
  st = fm_sheet_set_pivot_field_axis(wb, 0, pivot_index, amount_field, FM_PIVOT_AXIS_VALUE);
  if (st != 0) {
    return st;
  }
  fm_pivot_field_patch_t amount_patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  amount_patch.subtotal_top_op = FM_PATCH_OP_SET;
  amount_patch.subtotal_top = 0;
  st = fm_sheet_update_pivot_field(wb, 0, pivot_index, amount_field, &amount_patch);
  if (st != 0) {
    return st;
  }
  if (st != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(date_field)};
  st = fm_sheet_set_pivot_row_field_order(wb, 0, pivot_index, row_order, 1U);
  if (st != 0) {
    return st;
  }
  fm_pivot_data_field_t df_spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  df_spec.has_show_as_base_field = 0;
  df_spec.has_show_as_base_item = 0;
  std::size_t df_index = 99;
  st = fm_sheet_add_pivot_data_field(wb, 0, pivot_index, &df_spec, &df_index);
  if (st != 0) {
    return st;
  }
  // Grouped last: the group field is appended after every source field.
  st = SetPivotFieldDateGroup(wb, 0, pivot_index, date_field, FM_PIVOT_PERIOD_DAYS, FM_PIVOT_CALENDAR_GREGORIAN,
                              /*interval_days=*/7, start_serial,
                              /*end_or_nan=*/std::numeric_limits<double>::quiet_NaN());
  if (st != 0) {
    return st;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_index;
  }
  return 0;
}

// Collects (row label, Sum(Amount)) pairs from a one-row-field,
// one-data-field pivot's projected grid, in row order.
std::vector<std::pair<std::string, double>> CollectDaysGroupRows(fm_pivot_report_t* handle) {
  std::vector<fm_pivot_cell_info_t> cells = CollectCells(handle);
  std::vector<std::pair<std::string, double>> labels_by_row;
  std::vector<double> sums_by_row;
  for (const fm_pivot_cell_info_t& c : cells) {
    if (c.kind == FM_PIVOT_CELL_KIND_ROW_LABEL && c.value.kind == FM_VALUE_KIND_TEXT) {
      labels_by_row.emplace_back(std::string(c.value.u.text), 0.0);
    } else if (c.kind == FM_PIVOT_CELL_KIND_DATA && c.value.kind == FM_VALUE_KIND_NUMBER) {
      sums_by_row.push_back(c.value.u.number);
    }
  }
  EXPECT_EQ(labels_by_row.size(), sums_by_row.size());
  for (std::size_t i = 0; i < labels_by_row.size() && i < sums_by_row.size(); ++i) {
    labels_by_row[i].second = sums_by_row[i];
  }
  return labels_by_row;
}

// Region shared items are `[0] = "North"`, `[1] = blank`; the two records
// carry one of each. The blank shared item deliberately does NOT sit at index
// 0, which is what a name-addressed item silently binds to.
fm_status_t BuildBlankItemPivot(fm_workbook_t* wb, std::uint32_t* out_cache_id, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  if (fm_status_t st = fm_workbook_add_pivot_cache(wb, 0U, &cache_id); st != 0) {
    return st;
  }
  std::size_t region_index = 99;
  std::size_t amount_index = 99;
  if (fm_status_t st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Region", &region_index); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_add_pivot_cache_field(wb, cache_id, "Amount", &amount_index); st != 0) {
    return st;
  }
  if (fm_status_t st =
          fm_workbook_add_pivot_cache_field_shared_item_text(wb, cache_id, region_index, "North", ignored_index());
      st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_add_pivot_cache_field_shared_item_blank(wb, cache_id, region_index, ignored_index());
      st != 0) {
    return st;
  }

  std::size_t rec_index = 99;
  if (fm_status_t st = fm_workbook_add_pivot_cache_record(wb, cache_id, &rec_index); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, region_index, 0.0); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, amount_index, 100.0);
      st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_add_pivot_cache_record(wb, cache_id, &rec_index); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_set_pivot_cache_record_blank(wb, cache_id, rec_index, region_index); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, amount_index, 200.0);
      st != 0) {
    return st;
  }

  std::size_t pivot_index = 99;
  if (fm_status_t st = AddPivot(wb, 0, "PT", cache_id, 0U, 0U, &pivot_index); st != 0) {
    return st;
  }
  // The pivot carries one field per cache field, in cache order.
  const std::size_t region_field = region_index;
  if (fm_status_t st = fm_sheet_set_pivot_field_axis(wb, 0, pivot_index, region_field, FM_PIVOT_AXIS_ROW); st != 0) {
    return st;
  }
  fm_pivot_field_patch_t region_patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  region_patch.subtotal_top_op = FM_PATCH_OP_SET;
  region_patch.subtotal_top = 0;
  if (fm_status_t st = fm_sheet_update_pivot_field(wb, 0, pivot_index, region_field, &region_patch); st != 0) {
    return st;
  }
  // The pivot carries one field per cache field, in cache order.
  const std::size_t amount_field = amount_index;
  if (fm_status_t st = fm_sheet_set_pivot_field_axis(wb, 0, pivot_index, amount_field, FM_PIVOT_AXIS_VALUE); st != 0) {
    return st;
  }
  fm_pivot_field_patch_t amount_patch = FM_STRUCT_INIT(fm_pivot_field_patch_t);
  amount_patch.subtotal_top_op = FM_PATCH_OP_SET;
  amount_patch.subtotal_top = 0;
  if (fm_status_t st = fm_sheet_update_pivot_field(wb, 0, pivot_index, amount_field, &amount_patch); st != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(region_field)};
  if (fm_status_t st = fm_sheet_set_pivot_row_field_order(wb, 0, pivot_index, row_order, 1U); st != 0) {
    return st;
  }
  fm_pivot_data_field_t df_spec = FM_STRUCT_INIT(fm_pivot_data_field_t);
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGGREGATION_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  df_spec.has_show_as_base_field = 0;
  df_spec.has_show_as_base_item = 0;
  std::size_t df_index = 99;
  if (fm_status_t st = fm_sheet_add_pivot_data_field(wb, 0, pivot_index, &df_spec, &df_index); st != 0) {
    return st;
  }
  if (out_cache_id != nullptr) {
    *out_cache_id = cache_id;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_index;
  }
  return 0;
}

// Sum of every `FM_PIVOT_CELL_KIND_DATA` cell in the projected layout.
double SumDataCells(fm_pivot_report_t* handle) {
  double total = 0.0;
  const std::size_t n = count_of(fm_pivot_report_get_cell_count, handle);
  for (std::size_t i = 0; i < n; ++i) {
    fm_pivot_cell_info_t cell = FM_STRUCT_INIT(fm_pivot_cell_info_t);
    EXPECT_EQ(fm_pivot_report_get_cell_at(handle, i, &cell), 0);
    if (cell.kind == FM_PIVOT_CELL_KIND_DATA && cell.value.kind == FM_VALUE_KIND_NUMBER) {
      total += cell.value.u.number;
    }
  }
  return total;
}

}  // namespace

TEST(FormulonCApiPivot, DaysGroupMatchesPivotWeek1900FixtureThroughTheAbi) {
  // Ground truth transcribed from
  // tests/fixtures/excel/win/pivot_week_1900/results.json (Windows Excel
  // 365 ja-JP, build 16.0.20228): four Start variants of the same
  // By-7-days grouping, read from the real Excel-rendered pivot. `auto`
  // (Start=True) and `start0` (Start=0) coincide because 0 is already
  // the data minimum. The labels are the ja-JP short dates.
  const double kAuto = std::numeric_limits<double>::quiet_NaN();
  const std::vector<std::pair<std::string, double>> kAutoAndStart0Rows = {
      {"1900/1/0 - 1900/1/6", 7.0},   {"1900/1/7 - 1900/1/13", 7.0},  {"1900/1/14 - 1900/1/20", 7.0},
      {"1900/1/21 - 1900/1/27", 7.0}, {"1900/1/28 - 1900/2/3", 7.0},  {"1900/2/4 - 1900/2/10", 7.0},
      {"1900/2/11 - 1900/2/17", 7.0}, {"1900/2/18 - 1900/2/24", 7.0}, {"1900/2/25 - 1900/3/2", 7.0},
      {"1900/3/3 - 1900/3/9", 7.0},   {"1900/3/10 - 1900/3/11", 1.0},
  };
  const std::vector<std::pair<std::string, double>> kStart1Rows = {
      {"<1900/1/1", 1.0},
      {"1900/1/1 - 1900/1/7", 7.0},
      {"1900/1/8 - 1900/1/14", 7.0},
      {"1900/1/15 - 1900/1/21", 7.0},
      {"1900/1/22 - 1900/1/28", 7.0},
      {"1900/1/29 - 1900/2/4", 7.0},
      {"1900/2/5 - 1900/2/11", 7.0},
      {"1900/2/12 - 1900/2/18", 7.0},
      {"1900/2/19 - 1900/2/25", 7.0},
      {"1900/2/26 - 1900/3/3", 7.0},
      {"1900/3/4 - 1900/3/10", 7.0},
  };
  const std::vector<std::pair<std::string, double>> kStart61Rows = {
      {"<1900/3/1", 61.0},
      {"1900/3/1 - 1900/3/7", 7.0},
      {"1900/3/8 - 1900/3/11", 3.0},
  };

  const struct {
    const char* name;
    double start_serial;
    const std::vector<std::pair<std::string, double>>* expected;
  } variants[] = {
      {"auto", kAuto, &kAutoAndStart0Rows},
      {"start0", 0.0, &kAutoAndStart0Rows},
      {"start1", 1.0, &kStart1Rows},
      {"start61", 61.0, &kStart61Rows},
  };

  for (const auto& variant : variants) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0) << variant.name;
    ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
    std::size_t pivot_index = 0;
    ASSERT_EQ(BuildDaysGroupedPivot(wb.handle, variant.start_serial, &pivot_index), 0)
        << variant.name << ": " << fm_get_last_error_message();

    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0)
        << variant.name << ": " << fm_get_last_error_message();

    const std::vector<std::pair<std::string, double>> rows = CollectDaysGroupRows(projected.handle);
    ASSERT_EQ(rows.size(), variant.expected->size()) << variant.name;
    for (std::size_t i = 0; i < rows.size(); ++i) {
      EXPECT_EQ(rows[i].first, (*variant.expected)[i].first) << variant.name << " row " << i;
      EXPECT_DOUBLE_EQ(rows[i].second, (*variant.expected)[i].second) << variant.name << " row " << i;
    }
  }
}

TEST(FormulonCApiPivot, DateGroupRegroupsByPeriodsAndClearsThroughAGroupField) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildDaysGroupedPivot(wb.handle, std::numeric_limits<double>::quiet_NaN(), &pivot_index), 0)
      << fm_get_last_error_message();
  const double kAuto = std::numeric_limits<double>::quiet_NaN();

  // Regrouping through the days group field (2) replaces it on the base.
  ASSERT_EQ(SetPivotFieldDateGroup(wb.handle, 0, pivot_index, 2, FM_PIVOT_PERIOD_MONTHS | FM_PIVOT_PERIOD_YEARS,
                                   FM_PIVOT_CALENDAR_GREGORIAN, 1, kAuto, kAuto),
            0)
      << fm_get_last_error_message();
  std::size_t field_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot_index, &field_count), 0);
  EXPECT_EQ(field_count, 4U);
  fm_pivot_field_info_t date = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, &date), 0);
  fm_pivot_date_group_t state = FM_STRUCT_INIT(fm_pivot_date_group_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_date_group(wb.handle, 0, pivot_index, 0, &present, &state), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(state.periods, FM_PIVOT_PERIOD_MONTHS | FM_PIVOT_PERIOD_YEARS);
  EXPECT_EQ(state.calendar, FM_PIVOT_CALENDAR_GREGORIAN);
  EXPECT_EQ(state.interval_days, 1U);
  EXPECT_EQ(state.auto_start, FM_TRISTATE_TRUE);
  EXPECT_EQ(state.auto_end, FM_TRISTATE_TRUE);
  EXPECT_EQ(date.has_group_base_index, 0);
  EXPECT_EQ(date.axis, FM_PIVOT_AXIS_ABSENT);
  fm_pivot_field_info_t months = FM_STRUCT_INIT(fm_pivot_field_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 2, &months), 0);
  EXPECT_STREQ(months.source_name, "\xE6\x9C\x88 (Date)");  // 月 (Date)
  EXPECT_EQ(months.has_group_base_index, 1);
  EXPECT_EQ(months.group_base_index, 0U);
  EXPECT_EQ(months.axis, FM_PIVOT_AXIS_ROW);
  std::uint32_t order[8] = {};
  std::size_t order_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, order, 8U, &order_count), 0);
  ASSERT_EQ(order_count, 2U);
  EXPECT_EQ(order[0], 3U);
  EXPECT_EQ(order[1], 2U);
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_index, "1900\xE5\xB9\xB4"));  // 1900年
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_index, "3\xE6\x9C\x88"));     // 3月

  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0, pivot_index, 3, nullptr), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_get_pivot_field_count(wb.handle, 0, pivot_index, &field_count), 0);
  EXPECT_EQ(field_count, 2U);
  ASSERT_EQ(fm_sheet_get_pivot_field_at(wb.handle, 0, pivot_index, 0, &date), 0);
  ASSERT_EQ(fm_sheet_get_pivot_field_date_group(wb.handle, 0, pivot_index, 0, &present, &state), 0);
  EXPECT_EQ(present, 0);
  EXPECT_EQ(date.axis, FM_PIVOT_AXIS_ROW);
  ASSERT_EQ(fm_sheet_get_pivot_row_field_order(wb.handle, 0, pivot_index, order, 8U, &order_count), 0);
  ASSERT_EQ(order_count, 1U);
  EXPECT_EQ(order[0], 0U);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0, pivot_index, &cache_id), 0);
  std::size_t cache_fields = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_count(wb.handle, cache_id, &cache_fields), 0);
  EXPECT_EQ(cache_fields, 2U);
}

TEST(FormulonCApiPivot, AddItemAtBindsTheBlankItemByCacheIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // Both source rows contribute before any manual filter exists.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0)
        << fm_get_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 300.0);
  }

  // Joining the row axis listed North and then the blank shared item (cache
  // index 1); hiding the blank one leaves North.
  ASSERT_EQ(SetPivotFieldItemVisible(wb.handle, 0, pivot_index, 0, /*item_index=*/1U, /*visible=*/0), 0);
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0)
        << fm_get_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 100.0);
  }
}

TEST(FormulonCApiPivot, AddItemWithEmptyNameCannotHideTheBlankRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();

  // A name-addressed item defaults to cache index 0, which here holds
  // "North". The item is unlabelled, so it takes its label from that binding
  // and hides North, never the blank row. This is the gap
  // `fm_sheet_add_pivot_field_item_by_cache_index` closes.
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(wb.handle, 0, pivot_index, 0), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0, pivot_index, 0, "North", 1, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0, pivot_index, 0, "", 0, ignored_index()), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 200.0);
}

TEST(FormulonCApiPivot, AddItemAtHidesANonBlankItemBeforeAnyReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(wb.handle, 0, pivot_index, 0), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, pivot_index, 0, /*cache_index=*/1U,
                                                         /*visible=*/0, ignored_index()),
            0);
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0)
        << fm_get_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 400.0);  // North only: 100 + 300.
  }
  EXPECT_FALSE(LayoutHasText(wb.handle, pivot_index, "South"));

  // The label is derived at evaluation time, so repopulating the shared
  // items after the item was added changes nothing.
  ASSERT_EQ(fm_workbook_clear_pivot_cache_field_shared_items(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0, "North", ignored_index()), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(wb.handle, cache_id, 0, "South", ignored_index()), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_sheet_get_pivot_layout(wb.handle, 0, pivot_index, &projected.handle), 0) << fm_get_last_error_message();
  EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 400.0);
}

TEST(FormulonCApiPivot, SoleVisiblePageItemAddedByCacheIndexNamesItsValue) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  ASSERT_EQ(fm_sheet_set_pivot_row_field_order(wb.handle, 0, pivot_index, nullptr, 0U), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0, pivot_index, 0, FM_PIVOT_AXIS_PAGE), 0);
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(wb.handle, 0, pivot_index, 0), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, pivot_index, 0, /*cache_index=*/0U,
                                                         /*visible=*/0, ignored_index()),
            0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, pivot_index, 0, /*cache_index=*/1U,
                                                         /*visible=*/1, ignored_index()),
            0);
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_index, "South"));
}

namespace {

// Whether column `col` of sheet 0 holds the text `needle` in its first 100 rows.
bool SheetColumnHasText(const fm_workbook_t* wb, std::uint32_t col, std::string_view needle) {
  for (std::uint32_t r = 0; r < 100U; ++r) {
    fm_value_t v{};
    if (fm_sheet_get_value(wb, 0U, r, col, &v) == 0 && v.kind == FM_VALUE_KIND_TEXT && needle == v.u.text) {
      return true;
    }
  }
  return false;
}

bool ProjectionPending(const fm_workbook_t* wb, std::size_t pivot_index) {
  std::int32_t pending = -1;
  EXPECT_EQ(fm_sheet_is_pivot_projection_pending(wb, 0U, pivot_index, &pending), 0);
  return pending == 1;
}

double LayoutDataSum(fm_workbook_t* wb, std::size_t pivot_index) {
  PivotCellsGuard projected;
  EXPECT_EQ(fm_sheet_get_pivot_layout(wb, 0U, pivot_index, &projected.handle), 0) << fm_get_last_error_message();
  return projected.handle != nullptr ? SumDataCells(projected.handle) : 0.0;
}

// A two-column source on Sheet1 from A1, a NULL first-column entry being a
// blank cell, refreshed into a new cache; the pivot at F1 sums the second
// column with no field on an axis.
struct RefreshedPivot {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;

  void Build(const char* label_header, const std::vector<const char*>& labels) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 0U, label_header), 0);
    ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, 0U, 1U, "Sales"), 0);
    const double sales[] = {10.0, 30.0, 5.0, 1.0, 2.0, 3.0, 4.0, 6.0, 7.0, 8.0, 9.0, 11.0};
    ASSERT_LE(labels.size(), sizeof(sales) / sizeof(sales[0]));
    for (std::uint32_t i = 0; i < labels.size(); ++i) {
      if (labels[i] != nullptr) {
        ASSERT_EQ(fm_sheet_set_text(wb.handle, 0U, i + 1U, 0U, labels[i]), 0);
      }
      ASSERT_EQ(fm_sheet_set_number(wb.handle, 0U, i + 1U, 1U, sales[i]), 0);
    }
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &cache_id), 0);
    const std::string ref = "A1:B" + std::to_string(labels.size() + 1U);
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, /*present=*/1, ref.c_str(), "Sheet1", nullptr), 0);
    ASSERT_EQ(fm_workbook_refresh_pivot_cache(wb.handle, cache_id), 0) << fm_get_last_error_message();
    ASSERT_EQ(AddPivot(wb.handle, 0U, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/5U, &pivot), 0)
        << fm_get_last_error_message();
    ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot, 1U, FM_PIVOT_AXIS_VALUE), 0);
    fm_pivot_data_field_t sum = FM_STRUCT_INIT(fm_pivot_data_field_t);
    sum.name = "Sum of Sales";
    sum.field_index = 1U;
    sum.aggregation = FM_PIVOT_AGGREGATION_SUM;
    sum.number_format = "";
    sum.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
    ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, pivot, &sum, ignored_index()), 0)
        << fm_get_last_error_message();
  }

  std::size_t SharedCount(const fm_workbook_t* book, std::uint32_t id) const {
    std::size_t count = 99;
    EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(book, id, 0U, &count), 0);
    return count;
  }
};

}  // namespace

TEST(FormulonCApiPivot, AFieldJoiningRowsAfterARefreshGetsSharedItems) {
  RefreshedPivot f;
  f.Build("Region", {"East", "West", nullptr});
  // Refreshed with no pivot placing it, Region holds its values inline.
  ASSERT_EQ(f.SharedCount(f.wb.handle, f.cache_id), 0U);

  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(f.SharedCount(f.wb.handle, f.cache_id), 3U);
  const char* const expected[] = {"East", "West"};
  for (std::size_t i = 0; i < 2U; ++i) {
    fm_value_t v{};
    ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(f.wb.handle, f.cache_id, 0U, i, &v), 0);
    ASSERT_EQ(v.kind, FM_VALUE_KIND_TEXT);
    EXPECT_STREQ(v.u.text, expected[i]);
  }
  fm_value_t blank{};
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_value(f.wb.handle, f.cache_id, 0U, 2U, &blank), 0);
  EXPECT_EQ(blank.kind, FM_VALUE_KIND_BLANK);
  EXPECT_DOUBLE_EQ(LayoutDataSum(f.wb.handle, f.pivot), 45.0);

  // Hiding West (listed second) leaves East and the blank.
  ASSERT_EQ(SetPivotFieldItemVisible(f.wb.handle, 0U, f.pivot, 0U, /*item_index=*/1U, /*visible=*/0), 0);
  EXPECT_DOUBLE_EQ(LayoutDataSum(f.wb.handle, f.pivot), 15.0);
  // Hiding the blank instead leaves East and West.
  ASSERT_EQ(fm_sheet_clear_pivot_field_items(f.wb.handle, 0U, f.pivot, 0U), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(f.wb.handle, 0U, f.pivot, 0U, /*cache_index=*/2U,
                                                         /*visible=*/0, ignored_index()),
            0);
  EXPECT_DOUBLE_EQ(LayoutDataSum(f.wb.handle, f.pivot), 40.0);

  // The saved items name the shared items, so a reload hides the same member.
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  WorkbookGuard reloaded;
  ASSERT_EQ(fm_workbook_load(saved.data, saved.len, &reloaded.handle), 0) << fm_get_last_error_message();
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(reloaded.handle, 0U, f.pivot, &cache_id), 0);
  ASSERT_EQ(f.SharedCount(reloaded.handle, cache_id), 3U);
  std::size_t item_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_item_count(reloaded.handle, 0U, f.pivot, 0U, &item_count), 0);
  EXPECT_EQ(item_count, 3U);
  for (std::size_t i = 0; i < item_count; ++i) {
    fm_pivot_item_info_t item = FM_STRUCT_INIT(fm_pivot_item_info_t);
    ASSERT_EQ(fm_sheet_get_pivot_field_item_at(reloaded.handle, 0U, f.pivot, 0U, i, &item), 0);
    EXPECT_EQ(item.has_cache_index, 1) << i;
    EXPECT_LT(item.cache_index, 3U) << i;
    EXPECT_EQ(item.visible, item.cache_index == 2U ? 0 : 1) << i;
  }
  EXPECT_DOUBLE_EQ(LayoutDataSum(reloaded.handle, f.pivot), 40.0);
}

TEST(FormulonCApiPivot, AFieldJoiningRowsSavesItsItemsInDisplayOrder) {
  // Excel 365 measured: the joined field's shared items keep the first
  // spelling in source order with the blank in place, and its <items> list
  // them sorted ascending, blank last, then the default subtotal.
  RefreshedPivot f;
  f.Build("Product",
          {"pear", "Apple", nullptr, "Pear", "apple", "Fig", nullptr, "Apple", "Fig", "PEAR", "fig", "Apple"});
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0)
      << fm_get_last_error_message();
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string cache = ExtractZipEntry(package, "xl/pivotCache/pivotCacheDefinition1.xml");
  EXPECT_NE(cache.find("<sharedItems containsBlank=\"1\" count=\"4\"><s v=\"pear\"/><s v=\"Apple\"/><m/>"
                       "<s v=\"Fig\"/></sharedItems>"),
            std::string::npos)
      << cache;
  const std::string table = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  EXPECT_NE(table.find("<items count=\"5\"><item x=\"1\"/><item x=\"3\"/><item x=\"0\"/><item x=\"2\"/>"
                       "<item t=\"default\"/></items>"),
            std::string::npos)
      << table;
}

TEST(FormulonCApiPivot, AnItemPastTheSharedItemsIsLeftOutOfASave) {
  RefreshedPivot f;
  f.Build("Region", {"East", "West", nullptr});
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(f.wb.handle, 0U, f.pivot, 0U, /*cache_index=*/7U,
                                                         /*visible=*/0, ignored_index()),
            0);
  EXPECT_DOUBLE_EQ(LayoutDataSum(f.wb.handle, f.pivot), 45.0);
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(f.wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  const std::vector<std::uint8_t> package(saved.data, saved.data + saved.len);
  const std::string table = ExtractZipEntry(package, "xl/pivotTables/pivotTable1.xml");
  EXPECT_EQ(table.find("x=\"7\""), std::string::npos) << table;
  EXPECT_NE(table.find("<items count=\"4\">"), std::string::npos) << table;
}

TEST(FormulonCApiPivot, AFieldLeavingTheAxisKeepsItsSharedItems) {
  RefreshedPivot f;
  f.Build("Region", {"East", "West", nullptr});
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ABSENT), 0);
  EXPECT_EQ(f.SharedCount(f.wb.handle, f.cache_id), 3U);
}

TEST(FormulonCApiPivot, DateGroupingProjectsTheGroupedPivotAndLeavesTheOthersPending) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::size_t pivot_index = 0;
  // The fixture's last step groups the Date field (D) by 7 days.
  ASSERT_EQ(BuildDaysGroupedPivot(wb.handle, std::numeric_limits<double>::quiet_NaN(), &pivot_index), 0)
      << fm_get_last_error_message();
  EXPECT_FALSE(ProjectionPending(wb.handle, pivot_index));
  EXPECT_TRUE(SheetColumnHasText(wb.handle, 3U, "1900/1/0 - 1900/1/6"));

  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0U, pivot_index, &cache_id), 0);
  std::size_t other = 0;
  ASSERT_EQ(AddPivot(wb.handle, 0U, "PT2", cache_id, /*anchor_row=*/0U, /*anchor_col=*/10U, &other), 0)
      << fm_get_last_error_message();
  ASSERT_FALSE(ProjectionPending(wb.handle, other));

  // Clearing through the group field (2) puts the dates back on the rows.
  ASSERT_EQ(fm_sheet_set_pivot_field_date_group(wb.handle, 0U, pivot_index, 2U, nullptr), 0)
      << fm_get_last_error_message();
  EXPECT_FALSE(ProjectionPending(wb.handle, pivot_index));
  EXPECT_FALSE(SheetColumnHasText(wb.handle, 3U, "1900/1/0 - 1900/1/6"));
  EXPECT_TRUE(ProjectionPending(wb.handle, other));
}

TEST(FormulonCApiPivot, NumericGroupingProjectsTheGroupedPivot) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_index), 0) << fm_get_last_error_message();
  // Amount (100, 200, 300, 400) joins Region on the rows, then groups by 100.
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot_index, 1U, FM_PIVOT_AXIS_ROW), 0)
      << fm_get_last_error_message();
  ASSERT_FALSE(SheetColumnHasText(wb.handle, 3U, "100-199"));
  fm_pivot_numeric_group_t spec = FM_STRUCT_INIT(fm_pivot_numeric_group_t);
  spec.interval = 100.0;
  ASSERT_EQ(fm_sheet_set_pivot_field_numeric_group(wb.handle, 0U, pivot_index, 1U, &spec), 0)
      << fm_get_last_error_message();
  EXPECT_FALSE(ProjectionPending(wb.handle, pivot_index));
  EXPECT_TRUE(SheetColumnHasText(wb.handle, 3U, "100-199"));
  // The last value, 400, is the end on a range edge, so the last range reads up to it.
  EXPECT_TRUE(SheetColumnHasText(wb.handle, 3U, "300-400"));
}

namespace {

constexpr fm_status_t kAlreadyExists = static_cast<fm_status_t>(formulon::FormulonErrorCode::kAlreadyExists);

// A cache filled through the record setters: Product (text) and Sales, one
// record per entry; the pivot at D1 sums Sales with no field on an axis.
struct HandFilledPivot {
  WorkbookGuard wb;
  std::uint32_t cache_id = 0;
  std::size_t pivot = 0;

  void Build(const std::vector<const char*>& products) {
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &cache_id), 0);
    ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, /*present=*/1, "A1:B5", "Sheet1", nullptr), 0);
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "Product", ignored_index()), 0);
    ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "Sales", ignored_index()), 0);
    for (std::size_t i = 0; i < products.size(); ++i) {
      AddRecord(products[i], static_cast<double>(i + 1U));
    }
    ASSERT_EQ(AddPivot(wb.handle, 0U, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot), 0)
        << fm_get_last_error_message();
  }

  void AddRecord(const char* product, double sales) {
    std::size_t record = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_text(wb.handle, cache_id, record, 0U, product), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 1U, sales), 0);
  }

  std::size_t SharedCount(std::size_t field) const {
    std::size_t count = 99;
    EXPECT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, field, &count), 0);
    return count;
  }

  fm_value_t RecordValue(std::size_t record, std::size_t field) const {
    fm_value_t v{};
    EXPECT_EQ(fm_workbook_get_pivot_cache_record_value(wb.handle, cache_id, record, field, &v), 0);
    return v;
  }
};

}  // namespace

TEST(FormulonCApiPivot, AFieldLeavingEveryAxisSavesItsItems) {
  // Excel 365 measured: a field taken off the axes keeps its <items>, only
  // the axis attribute goes.
  RefreshedPivot f;
  f.Build("Product",
          {"pear", "Apple", nullptr, "Pear", "apple", "Fig", nullptr, "Apple", "Fig", "PEAR", "fig", "Apple"});
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ABSENT), 0)
      << fm_get_last_error_message();
  std::size_t item_count = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_item_count(f.wb.handle, 0U, f.pivot, 0U, &item_count), 0);
  EXPECT_EQ(item_count, 4U);
  const std::string table = SavedPivotXml(f.wb.handle);
  EXPECT_NE(table.find("<pivotField showAll=\"0\"><items count=\"5\"><item x=\"1\"/><item x=\"3\"/><item x=\"0\"/>"
                       "<item x=\"2\"/><item t=\"default\"/></items></pivotField>"),
            std::string::npos)
      << table;
}

TEST(FormulonCApiPivot, AHandFilledFieldJoiningRowsSharesItsValuesAndRecordWritesStayValues) {
  HandFilledPivot f;
  f.Build({"pear", "Apple", "PEAR"});
  ASSERT_EQ(f.SharedCount(0U), 0U);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0)
      << fm_get_last_error_message();
  ASSERT_EQ(f.SharedCount(0U), 2U);

  // A later text write reuses the item of its case variant.
  f.AddRecord("apple", 4.0);
  EXPECT_EQ(f.SharedCount(0U), 2U);
  fm_value_t apple = f.RecordValue(3U, 0U);
  ASSERT_EQ(apple.kind, FM_VALUE_KIND_TEXT);
  EXPECT_STREQ(apple.u.text, "Apple");
  // A number is stored as a number, not read as an item index.
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(f.wb.handle, f.cache_id, 3U, 0U, 7.0), 0);
  EXPECT_EQ(f.SharedCount(0U), 3U);
  const fm_value_t seven = f.RecordValue(3U, 0U);
  ASSERT_EQ(seven.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(seven.u.number, 7.0);
  ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(f.wb.handle, f.cache_id, 0U, 0U, 1.0), 0);
  const fm_value_t one = f.RecordValue(0U, 0U);
  ASSERT_EQ(one.kind, FM_VALUE_KIND_NUMBER);
  EXPECT_EQ(one.u.number, 1.0);
  EXPECT_EQ(f.SharedCount(0U), 4U);
}

TEST(FormulonCApiPivot, ADateGroupBaseSharedFromDatesSavesItsItemsAsDates) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ASSERT_EQ(fm_workbook_set_profile_id(wb.handle, "win-365-ja_JP"), 0);
  std::size_t pivot_index = 0;
  ASSERT_EQ(BuildDaysGroupedPivot(wb.handle, std::numeric_limits<double>::quiet_NaN(), &pivot_index), 0)
      << fm_get_last_error_message();
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_sheet_get_pivot_cache_id(wb.handle, 0U, pivot_index, &cache_id), 0);
  std::size_t count = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_shared_item_count(wb.handle, cache_id, 0U, &count), 0);
  ASSERT_EQ(count, 71U);
  for (std::size_t i = 0; i < count; ++i) {
    std::int32_t is_date = 0;
    ASSERT_EQ(fm_workbook_is_pivot_cache_field_shared_item_date(wb.handle, cache_id, 0U, i, &is_date), 0);
    EXPECT_EQ(is_date, 1) << i;
  }
  BufferGuard saved;
  ASSERT_EQ(fm_workbook_save(wb.handle, FM_WORKBOOK_FORMAT_XLSX, &saved.data, &saved.len), 0)
      << fm_get_last_error_message();
  const std::string cache = ExtractZipEntry(std::vector<std::uint8_t>(saved.data, saved.data + saved.len),
                                            "xl/pivotCache/pivotCacheDefinition1.xml");
  const std::size_t first_field = cache.find("<cacheField name=\"Date\"");
  ASSERT_NE(first_field, std::string::npos) << cache;
  const std::string date_field = cache.substr(first_field, cache.find("</cacheField>", first_field) - first_field);
  EXPECT_NE(date_field.find("<d v="), std::string::npos) << date_field;
  EXPECT_EQ(date_field.find("<n v="), std::string::npos) << date_field;
}

TEST(FormulonCApiPivot, AGroupedReportThatCannotBePlacedIsBlocked) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache(wb.handle, 0U, &cache_id), 0);
  ASSERT_EQ(SetPivotCacheWorksheetSource(wb.handle, cache_id, /*present=*/1, "A1:B4", "Sheet1", nullptr), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "Date", ignored_index()), 0);
  ASSERT_EQ(fm_workbook_add_pivot_cache_field(wb.handle, cache_id, "Amount", ignored_index()), 0);
  // One date in each of 1901, 1902 and 1903.
  const double serials[] = {400.0, 800.0, 1200.0};
  for (std::size_t i = 0; i < 3U; ++i) {
    std::size_t record = 0;
    ASSERT_EQ(fm_workbook_add_pivot_cache_record(wb.handle, cache_id, &record), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_date(wb.handle, cache_id, record, 0U, serials[i]), 0);
    ASSERT_EQ(fm_workbook_set_pivot_cache_record_number(wb.handle, cache_id, record, 1U, 1.0), 0);
  }
  std::size_t pivot_index = 0;
  ASSERT_EQ(AddPivot(wb.handle, 0U, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot_index), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot_index, 0U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(wb.handle, 0U, pivot_index, 1U, FM_PIVOT_AXIS_VALUE), 0);
  fm_pivot_data_field_t sum = FM_STRUCT_INIT(fm_pivot_data_field_t);
  sum.name = "Sum of Amount";
  sum.field_index = 1U;
  sum.aggregation = FM_PIVOT_AGGREGATION_SUM;
  sum.number_format = "";
  sum.show_as = FM_PIVOT_SHOW_VALUES_AS_NORMAL;
  ASSERT_EQ(fm_sheet_add_pivot_data_field(wb.handle, 0U, pivot_index, &sum, ignored_index()), 0)
      << fm_get_last_error_message();
  // Three dates fill D1:E5; a year and a month row each need D1:E8, which D7:D8 blocks.
  const fm_cell_rect_t merge{6U, 3U, 7U, 3U};
  ASSERT_EQ(fm_sheet_add_merge(wb.handle, 0U, &merge, nullptr), 0) << fm_get_last_error_message();

  EXPECT_EQ(SetPivotFieldDateGroup(wb.handle, 0U, pivot_index, 0U, FM_PIVOT_PERIOD_MONTHS | FM_PIVOT_PERIOD_YEARS,
                                   FM_PIVOT_CALENDAR_GREGORIAN, 1U, std::numeric_limits<double>::quiet_NaN(),
                                   std::numeric_limits<double>::quiet_NaN()),
            0)
      << fm_get_last_error_message();
  // The grouping is applied and the report is blocked, not pending; the merge stays.
  std::size_t fields = 0;
  ASSERT_EQ(fm_workbook_get_pivot_cache_field_count(wb.handle, cache_id, &fields), 0);
  EXPECT_EQ(fields, 4U);
  EXPECT_FALSE(ProjectionPending(wb.handle, pivot_index));
  std::int32_t present = 0;
  fm_spill_info_t spill = FM_STRUCT_INIT(fm_spill_info_t);
  ASSERT_EQ(fm_sheet_get_pivot_spill(wb.handle, 0U, pivot_index, &present, &spill), 0) << fm_get_last_error_message();
  ASSERT_EQ(present, 1);
  EXPECT_EQ(spill.blocked, 1);
  EXPECT_EQ(spill.rect.first_row, 0U);
  EXPECT_EQ(spill.rect.first_col, 3U);
  EXPECT_EQ(spill.rect.last_row, 7U);
  EXPECT_EQ(spill.rect.last_col, 4U);
  fm_value_t anchor{};
  ASSERT_EQ(fm_sheet_get_value(wb.handle, 0U, 0U, 3U, &anchor), 0);
  ASSERT_EQ(anchor.kind, FM_VALUE_KIND_ERROR);
  EXPECT_EQ(anchor.u.error_code, static_cast<int32_t>(formulon::ErrorCode::Spill));
  std::size_t merges = 0;
  ASSERT_EQ(fm_sheet_get_merge_count(wb.handle, 0U, &merges), 0);
  EXPECT_EQ(merges, 1U);
}

TEST(FormulonCApiPivot, AddingASharedItemTheFieldHoldsIsRefused) {
  RefreshedPivot f;
  f.Build("Region", {"East", "West", nullptr});
  ASSERT_EQ(fm_sheet_set_pivot_field_axis(f.wb.handle, 0U, f.pivot, 0U, FM_PIVOT_AXIS_ROW), 0);
  ASSERT_EQ(f.SharedCount(f.wb.handle, f.cache_id), 3U);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(f.wb.handle, f.cache_id, 0U, "EAST", ignored_index()),
            kAlreadyExists);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_blank(f.wb.handle, f.cache_id, 0U, ignored_index()),
            kAlreadyExists);
  EXPECT_EQ(f.SharedCount(f.wb.handle, f.cache_id), 3U);
  std::size_t added = 0;
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_text(f.wb.handle, f.cache_id, 0U, "North", &added), 0);
  EXPECT_EQ(added, 3U);
  // A date and the number of its serial are two items.
  ASSERT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(f.wb.handle, f.cache_id, 0U, 5.0, ignored_index()), 0);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_date(f.wb.handle, f.cache_id, 0U, 5.0, ignored_index()), 0);
  EXPECT_EQ(fm_workbook_add_pivot_cache_field_shared_item_number(f.wb.handle, f.cache_id, 0U, 5.0, ignored_index()),
            kAlreadyExists);
}
