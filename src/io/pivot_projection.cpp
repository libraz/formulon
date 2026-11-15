#include "io/pivot_projection.h"

#include <memory>
#include <utility>

#include "io/pivot_line_items_writer.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"

namespace formulon::io {

ProjectedPivot project_pivot(const Workbook& wb, const pivot::PivotTable& table, const PivotCacheWritePlan* plan) {
  ProjectedPivot out;
  // A table built in memory has no items read with it; an edited one has stale ones.
  out.needs_items = table.layout_dirty() || table.definition_source() == pivot::PivotDefinitionSource::kBuilt;
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(wb.excel_profile());
  const pivot::PivotCache* cache = wb.find_pivot_cache(table.pivot_cache_id());
  // Reuse the result a host evaluation left behind; otherwise evaluate
  // without memoising, since this path has no formula host.
  std::shared_ptr<const pivot::PivotResult> result = table.last_result();
  if (!result) {
    if (cache == nullptr) {
      return out;
    }
    auto eval_or =
        pivot::evaluate(table, *cache, options, pivot::PivotFilterEnv{wb.pinned_now(), wb.date1904(), &wb.styles()});
    if (!eval_or) {
      return out;
    }
    result = std::make_shared<const pivot::PivotResult>(std::move(eval_or.value()));
  }
  const bool generate = out.needs_items && cache != nullptr && !table.projection_pending();
  pivot::PivotLineItems lines;
  auto layout_or = pivot::layout(table, *result, options, generate ? &lines : nullptr);
  if (!layout_or) {
    return out;
  }
  const pivot::PivotCells& cells = layout_or.value();
  // A withheld report is saved as its anchor alone.
  if (!table.blocked().has_value()) {
    out.span = PivotRenderedSpan{cells.body_rows, cells.body_cols, cells.body_top};
  }
  if (generate) {
    out.items = write_pivot_line_items(table, *cache, plan, lines);
  }
  return out;
}

}  // namespace formulon::io
