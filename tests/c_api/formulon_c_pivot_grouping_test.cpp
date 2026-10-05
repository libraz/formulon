// Stable C ABI PivotTable grouping and item-binding tests.

#include <utility>

#include "formulon_c_pivot_test_helpers.h"

namespace {

fm_status_t BuildDaysGroupedPivot(fm_workbook_t* wb, double start_serial, std::size_t* out_pivot_index) {
  std::uint32_t cache_id = 0;
  fm_status_t st = fm_workbook_pivot_cache_create(wb, 0, &cache_id);
  if (st != 0) {
    return st;
  }
  st = fm_workbook_pivot_cache_set_worksheet_source(wb, cache_id, /*present=*/1, "A1:B71", "Sheet1", nullptr);
  if (st != 0) {
    return st;
  }
  std::size_t date_idx = 99;
  st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Date", &date_idx);
  if (st != 0) {
    return st;
  }
  std::size_t amount_idx = 99;
  st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Amount", &amount_idx);
  if (st != 0) {
    return st;
  }
  for (double serial = 0.0; serial <= 70.0; serial += 1.0) {
    std::size_t rec_idx = 99;
    st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, date_idx, serial);
    if (st != 0) {
      return st;
    }
    st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, 1.0);
    if (st != 0) {
      return st;
    }
  }
  std::size_t pivot_idx = 99;
  st = fm_workbook_pivot_create(wb, 0, "PT", cache_id, /*anchor_row=*/0U, /*anchor_col=*/3U, &pivot_idx);
  if (st != 0) {
    return st;
  }
  fm_pivot_field_spec_t date_spec{};
  date_spec.source_name = "Date";
  date_spec.custom_name = "";
  date_spec.axis = FM_PIVOT_AXIS_ROW;
  date_spec.subtotal_top = 0;
  date_spec.number_format = "";
  std::size_t date_field = 99;
  st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &date_spec, &date_field);
  if (st != 0) {
    return st;
  }
  st = fm_workbook_pivot_field_set_date_group(wb, 0, pivot_idx, date_field, FM_PIVOT_DATE_DAYS,
                                              FM_PIVOT_CALENDAR_GREGORIAN, -1, -1, /*interval_days=*/7, start_serial,
                                              /*end_serial_or_neg1=*/-1.0);
  if (st != 0) {
    return st;
  }
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 0;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &amount_spec, &amount_field);
  if (st != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(date_field)};
  st = fm_workbook_pivot_set_row_field_order(wb, 0, pivot_idx, row_order, 1U);
  if (st != 0) {
    return st;
  }
  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  std::size_t df_idx = 99;
  st = fm_workbook_pivot_data_field_add(wb, 0, pivot_idx, &df_spec, &df_idx);
  if (st != 0) {
    return st;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_idx;
  }
  return 0;
}

// Collects (row label, Sum(Amount)) pairs from a one-row-field,
// one-data-field pivot's projected grid, in row order.
std::vector<std::pair<std::string, double>> CollectDaysGroupRows(fm_pivot_cells_t* handle) {
  std::vector<fm_pivot_cell_t> cells = CollectCells(handle);
  std::vector<std::pair<std::string, double>> labels_by_row;
  std::vector<double> sums_by_row;
  for (const fm_pivot_cell_t& c : cells) {
    if (c.kind == FM_PIVOT_CELL_ROW_LABEL && c.value.kind == FM_VAL_TEXT) {
      labels_by_row.emplace_back(std::string(c.value.u.text), 0.0);
    } else if (c.kind == FM_PIVOT_CELL_DATA && c.value.kind == FM_VAL_NUMBER) {
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
  if (fm_status_t st = fm_workbook_pivot_cache_create(wb, 0U, &cache_id); st != 0) {
    return st;
  }
  std::size_t region_idx = 99;
  std::size_t amount_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Region", &region_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add(wb, cache_id, "Amount", &amount_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add_shared_item_text(wb, cache_id, region_idx, "North"); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_field_add_shared_item_blank(wb, cache_id, region_idx); st != 0) {
    return st;
  }

  std::size_t rec_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, region_idx, 0.0); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, 100.0); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_add(wb, cache_id, &rec_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_blank(wb, cache_id, rec_idx, region_idx); st != 0) {
    return st;
  }
  if (fm_status_t st = fm_workbook_pivot_cache_record_set_number(wb, cache_id, rec_idx, amount_idx, 200.0); st != 0) {
    return st;
  }

  std::size_t pivot_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_create(wb, 0, "PT", cache_id, 0U, 0U, &pivot_idx); st != 0) {
    return st;
  }
  fm_pivot_field_spec_t region_spec{};
  region_spec.source_name = "Region";
  region_spec.custom_name = "";
  region_spec.axis = FM_PIVOT_AXIS_ROW;
  region_spec.subtotal_top = 0;
  region_spec.number_format = "";
  std::size_t region_field = 99;
  if (fm_status_t st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &region_spec, &region_field); st != 0) {
    return st;
  }
  fm_pivot_field_spec_t amount_spec{};
  amount_spec.source_name = "Amount";
  amount_spec.custom_name = "";
  amount_spec.axis = FM_PIVOT_AXIS_VALUE;
  amount_spec.subtotal_top = 0;
  amount_spec.number_format = "";
  std::size_t amount_field = 99;
  if (fm_status_t st = fm_workbook_pivot_field_add(wb, 0, pivot_idx, &amount_spec, &amount_field); st != 0) {
    return st;
  }
  const std::uint32_t row_order[] = {static_cast<std::uint32_t>(region_field)};
  if (fm_status_t st = fm_workbook_pivot_set_row_field_order(wb, 0, pivot_idx, row_order, 1U); st != 0) {
    return st;
  }
  fm_pivot_data_field_spec_t df_spec{};
  df_spec.name = "Sum of Amount";
  df_spec.field_index = static_cast<std::uint32_t>(amount_field);
  df_spec.aggregation = FM_PIVOT_AGG_SUM;
  df_spec.number_format = "";
  df_spec.show_as = FM_PIVOT_SHOW_AS_NORMAL;
  df_spec.show_as_base_field = -1;
  df_spec.show_as_base_item = -1;
  std::size_t df_idx = 99;
  if (fm_status_t st = fm_workbook_pivot_data_field_add(wb, 0, pivot_idx, &df_spec, &df_idx); st != 0) {
    return st;
  }
  if (out_cache_id != nullptr) {
    *out_cache_id = cache_id;
  }
  if (out_pivot_index != nullptr) {
    *out_pivot_index = pivot_idx;
  }
  return 0;
}

// Sum of every `FM_PIVOT_CELL_DATA` cell in the projected layout.
double SumDataCells(fm_pivot_cells_t* handle) {
  double total = 0.0;
  const std::size_t n = fm_pivot_cells_count(handle);
  for (std::size_t i = 0; i < n; ++i) {
    fm_pivot_cell_t cell{};
    EXPECT_EQ(fm_pivot_cells_at(handle, i, &cell), 0);
    if (cell.kind == FM_PIVOT_CELL_DATA && cell.value.kind == FM_VAL_NUMBER) {
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
  // the data minimum.
  const double kAuto = -1.0;
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
    ASSERT_EQ(fm_workbook_create(&wb.handle), 0) << variant.name;
    std::size_t pivot_idx = 0;
    ASSERT_EQ(BuildDaysGroupedPivot(wb.handle, variant.start_serial, &pivot_idx), 0)
        << variant.name << ": " << fm_last_error_message();

    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0)
        << variant.name << ": " << fm_last_error_message();

    const std::vector<std::pair<std::string, double>> rows = CollectDaysGroupRows(projected.handle);
    ASSERT_EQ(rows.size(), variant.expected->size()) << variant.name;
    for (std::size_t i = 0; i < rows.size(); ++i) {
      EXPECT_EQ(rows[i].first, (*variant.expected)[i].first) << variant.name << " row " << i;
      EXPECT_DOUBLE_EQ(rows[i].second, (*variant.expected)[i].second) << variant.name << " row " << i;
    }
  }
}

TEST(FormulonCApiPivot, AddItemAtBindsTheBlankItemByCacheIndex) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // Both source rows contribute before any manual filter exists.
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 300.0);
  }

  // The blank shared item is at cache index 1, so the index-addressed adder is
  // the only way to name it: an empty label carries no binding of its own.
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "North", 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, /*cache_index=*/1U, /*visible=*/0), 0);
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 100.0);
  }
}

TEST(FormulonCApiPivot, AddItemWithEmptyNameCannotHideTheBlankRow) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildBlankItemPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();

  // A name-addressed item defaults to cache index 0, which here holds
  // "North". The item is unlabelled, so it takes its label from that binding
  // and hides North, never the blank row. This is the gap
  // `fm_workbook_pivot_field_add_item_at` closes.
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "North", 1), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item(wb.handle, 0, pivot_idx, 0, "", 0), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 200.0);
}

TEST(FormulonCApiPivot, AddItemAtHidesANonBlankItemBeforeAnyReload) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_field_clear_items(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, /*cache_index=*/1U, /*visible=*/0), 0);
  {
    PivotCellsGuard projected;
    ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
    EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 400.0);  // North only: 100 + 300.
  }
  EXPECT_FALSE(LayoutHasText(wb.handle, pivot_idx, "South"));

  // The label is derived at evaluation time, so repopulating the shared
  // items after the item was added changes nothing.
  ASSERT_EQ(fm_workbook_pivot_cache_field_clear_shared_items(wb.handle, cache_id, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb.handle, cache_id, 0, "North"), 0);
  ASSERT_EQ(fm_workbook_pivot_cache_field_add_shared_item_text(wb.handle, cache_id, 0, "South"), 0);
  PivotCellsGuard projected;
  ASSERT_EQ(fm_workbook_pivot_layout(wb.handle, 0, pivot_idx, &projected.handle), 0) << fm_last_error_message();
  EXPECT_DOUBLE_EQ(SumDataCells(projected.handle), 400.0);
}

TEST(FormulonCApiPivot, SoleVisiblePageItemAddedByCacheIndexNamesItsValue) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create(&wb.handle), 0);
  std::uint32_t cache_id = 0;
  std::size_t pivot_idx = 0;
  ASSERT_EQ(BuildScratchPivot(wb.handle, &cache_id, &pivot_idx), 0) << fm_last_error_message();
  ASSERT_EQ(fm_workbook_pivot_set_row_field_order(wb.handle, 0, pivot_idx, nullptr, 0U), 0);
  ASSERT_EQ(fm_workbook_pivot_field_set_axis(wb.handle, 0, pivot_idx, 0, FM_PIVOT_AXIS_PAGE), 0);
  ASSERT_EQ(fm_workbook_pivot_field_clear_items(wb.handle, 0, pivot_idx, 0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, /*cache_index=*/0U, /*visible=*/0), 0);
  ASSERT_EQ(fm_workbook_pivot_field_add_item_at(wb.handle, 0, pivot_idx, 0, /*cache_index=*/1U, /*visible=*/1), 0);
  EXPECT_TRUE(LayoutHasText(wb.handle, pivot_idx, "South"));
}
