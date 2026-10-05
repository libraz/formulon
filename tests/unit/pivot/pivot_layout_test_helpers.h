#pragma once

#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::pivot::layout_test {

inline Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

inline PivotCache build_basic_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, product));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "Widget", 100.0);
  add("North", "Gadget", 50.0);
  add("South", "Widget", 200.0);
  add("South", "Gadget", 300.0);
  add("North", "Widget", 25.0);
  return cache;
}

inline PivotTable build_table(std::vector<std::uint32_t> row_fields, std::vector<std::uint32_t> col_fields) {
  PivotTable table;
  table.set_name("Pivot1");
  table.set_pivot_cache_id(1);
  table.set_anchor(2, 3, 1, 1);  // D3.

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField product;
  product.source_name = "Product";
  product.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(product));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = std::move(row_fields);
  table.mutable_col_field_order() = std::move(col_fields);

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

inline PivotCache build_region_year_quarter_cache() {
  PivotCache cache;
  cache.set_cache_id(2);
  cache.mutable_fields().push_back(PivotCacheField{"Region", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Year", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Quarter", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});

  auto add = [&](const char* region, const char* year, const char* quarter, double amount) {
    PivotCacheRecord rec;
    rec.cells.push_back(owned_text(cache, region));
    rec.cells.push_back(owned_text(cache, year));
    rec.cells.push_back(owned_text(cache, quarter));
    rec.cells.push_back(Value::number(amount));
    cache.mutable_records().push_back(std::move(rec));
  };

  add("North", "2025", "Q1", 10.0);
  add("North", "2025", "Q2", 20.0);
  add("North", "2026", "Q1", 30.0);
  add("South", "2025", "Q1", 40.0);
  add("South", "2025", "Q2", 50.0);
  add("South", "2026", "Q1", 60.0);
  return cache;
}

inline PivotTable build_region_year_quarter_table() {
  PivotTable table;
  table.set_name("PivotMixed");
  table.set_pivot_cache_id(2);
  table.set_anchor(2, 3, 1, 1);

  PivotField region;
  region.source_name = "Region";
  region.axis = PivotAxis::Row;
  PivotField year;
  year.source_name = "Year";
  year.axis = PivotAxis::Col;
  year.subtotal_top = true;
  PivotField quarter;
  quarter.source_name = "Quarter";
  quarter.axis = PivotAxis::Col;
  PivotField amount;
  amount.source_name = "Amount";
  amount.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region));
  table.mutable_fields().push_back(std::move(year));
  table.mutable_fields().push_back(std::move(quarter));
  table.mutable_fields().push_back(std::move(amount));
  table.mutable_row_field_order() = {0};
  table.mutable_col_field_order() = {1, 2};
  table.set_grand_totals(/*rows=*/false, /*cols=*/false);

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 3;
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));
  return table;
}

inline PivotLayoutOptions ja_jp_layout_options() {
  PivotLayoutOptions options;
  options.row_labels_label = "行ラベル";
  options.column_labels_label = "列ラベル";
  options.grand_total_label = "総計";
  options.subtotal_suffix = " 集計";
  return options;
}

inline const PivotCell* find_cell(const PivotCells& cells, std::uint32_t row, std::uint32_t col) {
  for (const PivotCell& cell : cells.cells) {
    if (cell.row == row && cell.col == col) {
      return &cell;
    }
  }
  return nullptr;
}

}  // namespace formulon::pivot::layout_test
