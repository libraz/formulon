//
// Implementation of `current_retained_part_fingerprint`. See
// `retained_part_fingerprint.h` for the contract.

#include "io/xlsb/retained_part_fingerprint.h"

#include <memory>
#include <vector>

#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/styles_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "utils/fnv1a.h"

namespace formulon {
namespace io {
namespace xlsb {

std::optional<std::uint64_t> current_retained_part_fingerprint(const Workbook& wb, const PassthroughPart& part) {
  switch (part.retained_origin) {
    case PassthroughPart::RetainedOrigin::kNone:
      return std::nullopt;

    case PassthroughPart::RetainedOrigin::kPivotTable: {
      if (part.origin_sheet_index >= wb.sheet_count()) {
        return std::nullopt;
      }
      const std::vector<std::unique_ptr<pivot::PivotTable>>& tables = wb.sheet(part.origin_sheet_index).pivot_tables();
      if (part.origin_pivot_index >= tables.size() || tables[part.origin_pivot_index] == nullptr) {
        return std::nullopt;
      }
      return fnv1a64(write_pivot_table_definition(*tables[part.origin_pivot_index]));
    }

    case PassthroughPart::RetainedOrigin::kPivotCacheDefinition: {
      const pivot::PivotCache* cache = wb.find_pivot_cache(part.origin_cache_id);
      if (cache == nullptr) {
        return std::nullopt;
      }
      return fnv1a64(write_pivot_cache_definition(*cache));
    }

    case PassthroughPart::RetainedOrigin::kPivotCacheRecords: {
      const pivot::PivotCache* cache = wb.find_pivot_cache(part.origin_cache_id);
      if (cache == nullptr) {
        return std::nullopt;
      }
      return fnv1a64(write_pivot_cache_records(*cache));
    }

    case PassthroughPart::RetainedOrigin::kStyles:
      return fnv1a64(write_styles(wb.styles()));
  }
  return std::nullopt;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
