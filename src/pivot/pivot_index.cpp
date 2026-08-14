//
// Implementation of `find_pivot_at_anchor`. See pivot_index.h for the
// public contract.

#include "pivot/pivot_index.h"

#include <cstddef>
#include <cstdint>
#include <memory>
#include <string_view>
#include <vector>

#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace pivot {

const PivotTable* find_pivot_at_anchor(const Workbook& wb, std::string_view sheet_name, std::uint32_t row,
                                       std::uint32_t col) noexcept {
  // Resolve the sheet via the workbook's Unicode-simple-fold lookup so this
  // helper agrees with `EvalContext::resolve_ref` on sheet identity.
  const Sheet* sheet = wb.sheet_by_name(sheet_name);
  if (sheet == nullptr) {
    return nullptr;
  }
  for (const std::unique_ptr<PivotTable>& pt : sheet->pivot_tables()) {
    if (pt == nullptr) {
      continue;
    }
    if (pt->contains(row, col)) {
      return pt.get();
    }
  }
  return nullptr;
}

void resolve_pivot_names(PivotTable& table, const PivotCache& cache) {
  const std::vector<PivotCacheField>& cache_fields = cache.fields();
  std::vector<PivotField>& fields = table.mutable_fields();
  for (std::size_t fi = 0; fi < fields.size(); ++fi) {
    if (fi >= cache_fields.size()) {
      break;
    }
    PivotField& field = fields[fi];
    const PivotCacheField& cache_field = cache_fields[fi];
    if (field.source_name.empty()) {
      field.source_name = cache_field.name;
    }
    for (PivotItem& item : field.items) {
      if (!item.name.empty()) {
        continue;
      }
      if (item.cache_index < cache_field.shared_items.size()) {
        item.name = display_string(cache_field.shared_items[item.cache_index]);
      }
    }
  }
}

void resolve_all_pivot_names(Workbook& wb) {
  for (std::size_t si = 0; si < wb.sheet_count(); ++si) {
    for (std::unique_ptr<PivotTable>& pt : wb.sheet(si).mutable_pivot_tables()) {
      if (pt == nullptr) {
        continue;
      }
      if (const PivotCache* cache = wb.find_pivot_cache(pt->pivot_cache_id()); cache != nullptr) {
        resolve_pivot_names(*pt, *cache);
      }
    }
  }
}

}  // namespace pivot
}  // namespace formulon
