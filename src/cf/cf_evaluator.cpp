//
// Implementation of the CF evaluator's orchestration: `make_match`,
// per-cell `evaluate_cf_at`, and the viewport `evaluate_cf_for_range`
// driver. Rule-by-rule matching lives in `rule_match.cpp`; visual
// (ColorScale / DataBar / IconSet) resolution lives in
// `scale_evaluator.cpp`; shared primitives (literal comparison,
// population gathering, percentile, weekday helpers) live in
// `cf_helpers.cpp`.
//
// See cf/cf_evaluator.h for the public contract.

#include "cf/cf_evaluator.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "cf/cf_helpers.h"
#include "cf/cf_match.h"
#include "cf/cf_types.h"
#include "cf/rule_match.h"
#include "cf/scale_evaluator.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/index_sort.h"
#include "utils/rect_iterator.h"
#include "value.h"

namespace formulon::cf {

CFMatch make_match(const CFRule& rule) {
  CFMatch match;
  match.rule_id = rule.id;
  match.priority = rule.priority;
  // Value-only overload covers dxf-driven rules. Visual rule kinds
  // need the cell value and `CFEvalContext` to resolve their render
  // payload; callers should use the context-aware overload below.
  match.kind = CFMatchKind::DifferentialFormat;
  match.dxf_id = rule.dxf_id;
  return match;
}

namespace {

bool sqref_contains(const std::vector<CFCellRange>& sqref, CellAddress target) {
  for (const CFCellRange& range : sqref) {
    if (target.row >= range.first.row && target.row <= range.last.row && target.col >= range.first.col &&
        target.col <= range.last.col) {
      return true;
    }
  }
  return false;
}

}  // namespace

CFMatch make_match(const CFRule& rule, const Value& cell_value, const CFEvalContext& ctx) {
  CFMatch match;
  match.rule_id = rule.id;
  match.priority = rule.priority;
  if (rule.type == RuleType::ColorScale) {
    match.kind = CFMatchKind::ColorScale;
    match.resolved_fill_color = scales::resolve_color_scale(rule, cell_value, ctx);
    return match;
  }
  if (rule.type == RuleType::DataBar) {
    match.kind = CFMatchKind::DataBar;
    match.data_bar_render = scales::resolve_data_bar(rule, cell_value, ctx);
    return match;
  }
  if (rule.type == RuleType::IconSet) {
    match.kind = CFMatchKind::IconSet;
    match.icon_render = scales::resolve_icon_set(rule, cell_value, ctx);
    return match;
  }
  // Every other kind is dxf-driven; the context-aware overload behaves
  // identically to the value-only one for those.
  match.kind = CFMatchKind::DifferentialFormat;
  match.dxf_id = rule.dxf_id;
  return match;
}

namespace {
// Whether `type` is a range-aware rule kind that benefits from a
// cached population. CellIs / Expression / TimePeriod / text-family /
// blanks-family / errors-family rules don't need it.
bool rule_uses_population(RuleType type) {
  switch (type) {
    case RuleType::ColorScale:
    case RuleType::DataBar:
    case RuleType::IconSet:
    case RuleType::AboveAverage:
    case RuleType::Top10:
      return true;
    default:
      return false;
  }
}

bool rule_uses_duplicates(RuleType type) {
  return type == RuleType::DuplicateValues || type == RuleType::UniqueValues;
}

/// Overlapping part of `a` and `b`, or `nullopt` when they are disjoint
/// (which includes either input being an inverted rectangle).
std::optional<CFCellRange> intersect_ranges(CFCellRange a, CFCellRange b) {
  CFCellRange hit{};
  hit.first.row = std::max(a.first.row, b.first.row);
  hit.first.col = std::max(a.first.col, b.first.col);
  hit.last.row = std::min(a.last.row, b.last.row);
  hit.last.col = std::min(a.last.col, b.last.col);
  if (hit.first.row > hit.last.row || hit.first.col > hit.last.col) {
    return std::nullopt;
  }
  return hit;
}

/// Inclusive `[first, last]` interval along one axis.
struct Span {
  std::uint32_t first;
  std::uint32_t last;
};

/// Sorts `spans` ascending and coalesces the overlapping ones in place,
/// leaving a disjoint ascending sequence. Adjacent-but-disjoint spans
/// stay separate; the sweep visits them in order either way.
void merge_spans(std::vector<Span>& spans) {
  sort_by_index(spans, [](const Span& lhs, const Span& rhs) {
    return lhs.first < rhs.first || (lhs.first == rhs.first && lhs.last < rhs.last);
  });
  std::size_t out = 0;
  for (std::size_t i = 0; i < spans.size(); ++i) {
    if (out > 0 && spans[i].first <= spans[out - 1].last) {
      spans[out - 1].last = std::max(spans[out - 1].last, spans[i].last);
      continue;
    }
    spans[out] = spans[i];
    ++out;
  }
  spans.resize(out);
}

}  // namespace

struct PreparedCfEvaluator::Impl {
  struct RulePlanEntry {
    const ConditionalFormat* block = nullptr;
    const CFRule* rule = nullptr;
    CellAddress anchor{};
    std::size_t block_index = 0;
  };

  const Sheet& sheet;
  CFHost host;
  std::vector<RulePlanEntry> plan;
  std::vector<std::optional<ColorScalePopulation>> populations;
  std::vector<std::optional<DuplicateValuePopulation>> duplicates;
  PreparedCfFormulaCache formula_cache;

  Impl(const Sheet& source, const CFHost& source_host) : sheet(source), host(source_host) {
    if (host.arena == nullptr || host.registry == nullptr || host.eval_ctx == nullptr) {
      return;
    }
    const std::vector<ConditionalFormat>& blocks = sheet.conditional_formats();
    populations.resize(blocks.size());
    duplicates.resize(blocks.size());
    for (std::size_t block_index = 0; block_index < blocks.size(); ++block_index) {
      const ConditionalFormat& block = blocks[block_index];
      const CellAddress anchor = sqref_anchor(block.sqref);
      for (const CFRule& rule : block.rules) {
        plan.push_back(RulePlanEntry{&block, &rule, anchor, block_index});
      }
    }
    // sort_by_index is stable, so equal priorities keep document order.
    sort_by_index(plan, [](const RulePlanEntry& lhs, const RulePlanEntry& rhs) {
      return lhs.rule->priority < rhs.rule->priority;
    });
  }

  std::vector<CFMatch> evaluate(CellAddress target) {
    std::vector<CFMatch> matches;
    if (host.arena == nullptr || host.registry == nullptr || host.eval_ctx == nullptr) {
      return matches;
    }
    const eval::ScopedEvalProfile profile_scope(host.eval_ctx->excel_profile());
    const Value cell_value = sheet.resolve_cell_value(target.row, target.col);
    for (const RulePlanEntry& entry : plan) {
      const ConditionalFormat& block = *entry.block;
      const CFRule& rule = *entry.rule;
      if (!sqref_contains(block.sqref, target)) {
        continue;
      }

      CFEvalContext ctx;
      ctx.anchor = entry.anchor;
      ctx.target = target;
      ctx.arena = host.arena;
      ctx.registry = host.registry;
      ctx.eval_ctx = host.eval_ctx;
      ctx.today_serial = host.today_serial;
      ctx.sqref = &block.sqref;
      ctx.formula_cache = &formula_cache;

      if (rule_uses_population(rule.type)) {
        std::optional<ColorScalePopulation>& slot = populations[entry.block_index];
        if (!slot.has_value()) {
          slot = helpers::gather_population(block.sqref, sheet);
        }
        ctx.cached_population = &*slot;
      }
      if (rule_uses_duplicates(rule.type)) {
        std::optional<DuplicateValuePopulation>& slot = duplicates[entry.block_index];
        if (!slot.has_value()) {
          slot = helpers::gather_duplicate_population(block.sqref, sheet, host.eval_ctx->excel_profile());
        }
        ctx.cached_duplicates = &*slot;
      }

      if (!match_rule(rule, cell_value, ctx)) {
        continue;
      }
      matches.push_back(make_match(rule, cell_value, ctx));
      if (rule.stop_if_true) {
        break;
      }
    }
    return matches;
  }
};

PreparedCfEvaluator::PreparedCfEvaluator(const Sheet& sheet, const CFHost& host)
    : impl_(std::make_unique<Impl>(sheet, host)) {}

PreparedCfEvaluator::~PreparedCfEvaluator() = default;

PreparedCfEvaluator::PreparedCfEvaluator(PreparedCfEvaluator&&) noexcept = default;

PreparedCfEvaluator& PreparedCfEvaluator::operator=(PreparedCfEvaluator&&) noexcept = default;

std::vector<CFMatch> PreparedCfEvaluator::evaluate(CellAddress target) {
  return impl_->evaluate(target);
}

std::vector<CFMatch> evaluate_cf_at(const Sheet& sheet, CellAddress target, const CFHost& host) {
  PreparedCfEvaluator prepared(sheet, host);
  return prepared.evaluate(target);
}

Expected<std::vector<CFRangeCellMatches>, Error> evaluate_cf_for_range(const Sheet& sheet, CFCellRange range,
                                                                       const CFHost& host) {
  const utils::RectRange request(range.first.row, range.first.col, range.last.row, range.last.col);
  if (request.size() > kCfMaxViewportCells) {
    return make_error(FormulonErrorCode::kSecResourceLimit, "conditional-format range exceeds the viewport ceiling",
                      "cells=" + std::to_string(request.size()) + " ceiling=" + std::to_string(kCfMaxViewportCells));
  }

  std::vector<CFRangeCellMatches> results;
  if (host.arena == nullptr || host.registry == nullptr || host.eval_ctx == nullptr) {
    return results;
  }

  // Only cells a `<conditionalFormatting>` block covers can produce a
  // match, so the walk is driven by the blocks rather than by the
  // request: each sqref range is clipped to the request up front and
  // the sweep below visits the union of those clips. A viewport that
  // spans far more than the sheet's formatted area therefore costs the
  // formatted area, not the viewport.
  std::vector<CFCellRange> covered;
  for (const ConditionalFormat& block : sheet.conditional_formats()) {
    for (const CFCellRange& sqref_range : block.sqref) {
      if (std::optional<CFCellRange> clipped = intersect_ranges(sqref_range, range); clipped.has_value()) {
        covered.push_back(*clipped);
      }
    }
  }
  if (covered.empty()) {
    return results;
  }

  // One priority plan and its per-block caches serve the whole viewport.
  PreparedCfEvaluator prepared(sheet, host);

  // Row-major sweep over the union of the clipped rectangles. Rows come
  // from their merged row spans (ascending, disjoint) and each row's
  // columns from the merged column spans of the rectangles covering
  // that row, so the emitted order is exactly what a dense walk of the
  // request would produce — overlapping blocks visit a shared cell
  // once. A single-cell request (first == last) yields one visit.
  std::vector<Span> row_spans;
  row_spans.reserve(covered.size());
  for (const CFCellRange& rect : covered) {
    row_spans.push_back({rect.first.row, rect.last.row});
  }
  merge_spans(row_spans);

  std::vector<Span> col_spans;
  for (const Span& rows : row_spans) {
    for (std::uint32_t row = rows.first; row <= rows.last; ++row) {
      col_spans.clear();
      for (const CFCellRange& rect : covered) {
        if (row >= rect.first.row && row <= rect.last.row) {
          col_spans.push_back({rect.first.col, rect.last.col});
        }
      }
      merge_spans(col_spans);
      for (const Span& cols : col_spans) {
        for (std::uint32_t col = cols.first; col <= cols.last; ++col) {
          CellAddress cell{};
          cell.row = row;
          cell.col = col;
          std::vector<CFMatch> matches = prepared.evaluate(cell);
          if (matches.empty()) {
            continue;
          }
          CFRangeCellMatches entry;
          entry.cell = cell;
          entry.matches = std::move(matches);
          results.push_back(std::move(entry));
        }
      }
    }
  }
  return results;
}

}  // namespace formulon::cf
