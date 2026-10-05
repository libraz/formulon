#pragma once

#include <cstdint>
#include <utility>

#include "pivot/pivot_cache.h"
#include "pivot_evaluator_fixtures.h"
#include "value.h"

namespace formulon::pivot::filter_test {

// Shared-item Region cache used by label and page-selection tests. Records
// carry shared-item indices, matching the representation loaded from OOXML.
inline PivotCache build_shared_north_south_cache() {
  PivotCache cache;
  cache.set_cache_id(1);
  PivotCacheField region;
  region.name = "Region";
  region.shared_items.push_back(test::owned_text(cache, "North"));
  region.shared_items.push_back(test::owned_text(cache, "South"));
  cache.mutable_fields().push_back(std::move(region));
  cache.mutable_fields().push_back(PivotCacheField{"Product", {}});
  cache.mutable_fields().push_back(PivotCacheField{"Amount", {}});
  auto add = [&](std::uint32_t region_index, const char* product, double amount) {
    PivotCacheRecord rec;
    rec.cells = {Value::number(region_index), test::owned_text(cache, product), Value::number(amount)};
    rec.cell_is_index = {true, false, false};
    cache.mutable_records().push_back(std::move(rec));
  };
  add(0U, "Widget", 100.0);
  add(0U, "Gadget", 50.0);
  add(1U, "Widget", 200.0);
  add(1U, "Gadget", 300.0);
  add(0U, "Widget", 25.0);
  return cache;
}

}  // namespace formulon::pivot::filter_test
