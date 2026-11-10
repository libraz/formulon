// Stable C ABI PivotTable grouping and item-binding tests.

#include <limits>
#include <utility>

#include "formulon_c_out_index.h"
#include "formulon_c_pivot_test_helpers.h"

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
    st = fm_workbook_set_pivot_cache_record_number(wb, cache_id, rec_index, date_index, serial);
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
  fm_pivot_date_group_spec_t state = FM_STRUCT_INIT(fm_pivot_date_group_spec_t);
  std::int32_t present = 0;
  ASSERT_EQ(fm_sheet_get_pivot_field_date_group(wb.handle, 0, pivot_index, 0, &present, &state), 0);
  EXPECT_EQ(present, 1);
  EXPECT_EQ(state.periods, FM_PIVOT_PERIOD_MONTHS | FM_PIVOT_PERIOD_YEARS);
  EXPECT_EQ(state.calendar, FM_PIVOT_CALENDAR_GREGORIAN);
  EXPECT_EQ(state.interval_days, 1U);
  EXPECT_EQ(state.has_start, 0);
  EXPECT_EQ(state.has_end, 0);
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

  // The blank shared item is at cache index 1, so the index-addressed adder is
  // the only way to name it: an empty label carries no binding of its own.
  ASSERT_EQ(fm_sheet_add_pivot_field_item(wb.handle, 0, pivot_index, 0, "North", 1, ignored_index()), 0);
  ASSERT_EQ(fm_sheet_add_pivot_field_item_by_cache_index(wb.handle, 0, pivot_index, 0, /*cache_index=*/1U,
                                                         /*visible=*/0, ignored_index()),
            0);
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
