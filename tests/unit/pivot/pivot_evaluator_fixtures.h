//
// Shared fixtures for the pivot evaluator tests: a small Region / Product /
// Amount cache, a SUM(Amount) table over it, and a row-label lookup.

#ifndef FORMULON_TESTS_UNIT_PIVOT_PIVOT_EVALUATOR_FIXTURES_H_
#define FORMULON_TESTS_UNIT_PIVOT_PIVOT_EVALUATOR_FIXTURES_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "value.h"

namespace formulon::pivot::test {

// Pushes `s` into the cache's text storage so the returned `Value::text`
// holds a stable view for the lifetime of the cache.
inline Value owned_text(PivotCache& cache, std::string s) {
  cache.mutable_text_storage().push_back(std::move(s));
  return Value::text(cache.text_storage().back());
}

// Builds a 3-column cache: Region (text), Product (text), Amount (number).
// Records form a small, easy-to-reason-about dataset.
//
//   Region  Product  Amount
//   ------  -------  ------
//   North   Widget    100
//   North   Gadget     50
//   South   Widget    200
//   South   Gadget    300
//   North   Widget     25
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

// Builds a `PivotTable` matching `build_basic_cache()` with the supplied
// row/col field indices and a single SUM(Amount) data field.
inline PivotTable build_sum_amount_table(std::vector<std::uint32_t> row_fields, std::vector<std::uint32_t> col_fields) {
  PivotTable table;
  table.set_pivot_cache_id(1);

  PivotField region_f;
  region_f.source_name = "Region";
  region_f.axis = PivotAxis::Row;
  PivotField product_f;
  product_f.source_name = "Product";
  product_f.axis = PivotAxis::Row;
  PivotField amount_f;
  amount_f.source_name = "Amount";
  amount_f.axis = PivotAxis::Value;

  table.mutable_fields().push_back(std::move(region_f));
  table.mutable_fields().push_back(std::move(product_f));
  table.mutable_fields().push_back(std::move(amount_f));

  PivotDataField sum_amount;
  sum_amount.name = "Sum of Amount";
  sum_amount.field_index = 2;  // Amount.
  sum_amount.aggregation = Aggregation::Sum;
  table.mutable_data_fields().push_back(std::move(sum_amount));

  table.mutable_row_field_order() = std::move(row_fields);
  table.mutable_col_field_order() = std::move(col_fields);
  return table;
}

// Convenience: locate a top-level row leaf by its label, return the
// index into `result.values`.
inline std::size_t row_index(const PivotResult& r, const std::string& label) {
  for (std::size_t i = 0; i < r.rows.size(); ++i) {
    if (r.rows[i].label == label) {
      return i;
    }
  }
  return static_cast<std::size_t>(-1);
}
}  // namespace formulon::pivot::test

#endif  // FORMULON_TESTS_UNIT_PIVOT_PIVOT_EVALUATOR_FIXTURES_H_
