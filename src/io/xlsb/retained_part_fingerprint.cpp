//
// Implementation of `current_retained_part_fingerprint`. See
// `retained_part_fingerprint.h` for the contract.

#include "io/xlsb/retained_part_fingerprint.h"

#include <memory>
#include <string>
#include <vector>

#include "auto_filter.h"
#include "cell_style_save_plan.h"
#include "cf/cf_types.h"
#include "io/auto_filter_xml.h"
#include "io/pivot_cache_writer.h"
#include "io/pivot_table_writer.h"
#include "io/styles_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "table.h"
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

    case PassthroughPart::RetainedOrigin::kStyles: {
      // The style tables and the cell formats in use, so clearing a format's last user also stales the part.
      std::string state = write_styles(wb.styles());
      const UsedCellXfs used = collect_used_cell_xfs(wb);
      state += used.all_used ? 'A' : 'U';
      for (const bool u : used.used) {
        state += u ? '1' : '0';
      }
      for (const std::uint32_t m : used.merged) {
        state += ',' + std::to_string(m);
      }
      return fnv1a64(state);
    }
  }
  return std::nullopt;
}

std::uint64_t current_dxf_reference_fingerprint(const Workbook& wb) {
  std::string state;
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    state += ";sheet";
    for (const cf::ConditionalFormat& block : sheet.conditional_formats()) {
      state += ";block";
      for (const cf::CFRule& rule : block.rules) {
        state += rule.dxf_id.has_value() ? ',' + std::to_string(*rule.dxf_id) : std::string(",-");
        state += ':' + std::to_string(rule.priority);
        if (rule.ext_lst_raw.has_value()) {
          state += *rule.ext_lst_raw;
        }
      }
      if (block.ext_lst_raw.has_value()) {
        state += *block.ext_lst_raw;
      }
    }
    if (const AutoFilter* filter = sheet.auto_filter(); filter != nullptr) {
      state += ";filter" + serialize_auto_filter(*filter);
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    state += ";table" + std::to_string(table.id) + table.root_extra_attrs + table.sort_state_xml;
    for (const TableColumn& column : table.columns) {
      state += ";col" + column.extra_attrs;
    }
    if (const AutoFilter* filter = table.auto_filter_xml.get(); filter != nullptr) {
      state += ";filter" + serialize_auto_filter(*filter);
    }
  }
  return fnv1a64(state);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
