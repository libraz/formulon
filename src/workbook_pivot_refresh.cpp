// @size-budget: 8 KB
//
// `Workbook::refresh_pivot_cache`: Excel's Refresh of a worksheet-source
// pivot cache, and `Workbook::check_pivot_growth`, the rule every edit that
// grows a pivot report answers to.
//
// The refresh runs in two phases. The first writes nothing: it resolves and
// reads the source, builds the new cache apart from the old one, and lays out
// a remapped copy of every dependent pivot against it to check where each
// report would land. Only when every report may be written does the second
// phase swap the cache and the definitions in and project each report.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <utility>
#include <vector>

#include "chart/chart_template.h"
#include "eval/eval_profile_scope.h"
#include "eval/pivot_formula.h"
#include "eval/text_format/display_text.h"
#include "merge_range.h"
#include "parser/ast_format.h"
#include "parser/token.h"
#include "parser/tokenizer.h"
#include "pivot/field_grouping.h"
#include "pivot/field_remap.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_evaluator.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_refresh.h"
#include "pivot/pivot_source.h"
#include "pivot/pivot_table.h"
#include "sheet.h"
#include "style_resolve.h"
#include "styles.h"
#include "utils/a1_ref.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/resource_budget.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_internal.h"

namespace formulon {
namespace {

// `Sheet!A1` for a message.
std::string cell_name(const Sheet& sheet, std::uint32_t row, std::uint32_t col) {
  std::string out;
  parser::append_sheet_name(sheet.name(), /*force_quote=*/false, out);
  out.push_back('!');
  a1::append_a1(out, row, col);
  return out;
}

bool inside(const MergeRange& range, std::uint32_t row, std::uint32_t col) {
  return row >= range.first_row && row <= range.last_row && col >= range.first_col && col <= range.last_col;
}

// The source rectangle's header text, values, date marks and number formats.
// A merged cell other than its top-left reads as blank.
pivot::RefreshSource read_source(const Workbook& wb, const Sheet& sheet, const MergeRange& rect) {
  pivot::RefreshSource source;
  source.columns = rect.last_col - rect.first_col + 1U;
  source.rows = rect.last_row - rect.first_row;
  std::vector<MergeRange> merges;
  for (const MergeRange& merge : sheet.merges()) {
    if (intersects(normalized(merge), rect)) {
      merges.push_back(normalized(merge));
    }
  }
  const auto covered = [&merges](std::uint32_t row, std::uint32_t col) {
    for (const MergeRange& merge : merges) {
      if (inside(merge, row, col) && (row != merge.first_row || col != merge.first_col)) {
        return true;
      }
    }
    return false;
  };
  const StylesTable& styles = wb.styles();
  const ExcelProfile profile = wb.excel_profile();
  std::unordered_map<std::uint32_t, bool> date_format_of;
  const auto xf_of = [&](std::uint32_t row, std::uint32_t col) -> const CellXf* {
    if (styles.cell_xfs.empty()) {
      return nullptr;
    }
    const std::uint32_t index = select_effective_xf(sheet, row, col).xf_index;
    return &styles.cell_xfs[index < styles.cell_xfs.size() ? index : 0U];
  };

  source.headers.reserve(source.columns);
  for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
    source.headers.push_back(covered(rect.first_row, col)
                                 ? std::string()
                                 : text_format::format_cell_for_display(wb, sheet, rect.first_row, col).text);
  }
  source.num_fmt_ids.assign(source.columns, 0U);
  std::vector<bool> fmt_seen(source.columns, false);
  std::vector<bool> fmt_mixed(source.columns, false);
  source.cells.reserve(static_cast<std::size_t>(source.rows) * source.columns);
  for (std::uint32_t row = rect.first_row + 1U; row <= rect.last_row && rect.last_row > rect.first_row; ++row) {
    for (std::uint32_t col = rect.first_col; col <= rect.last_col; ++col) {
      const std::size_t c = col - rect.first_col;
      Value value = covered(row, col) ? Value::blank() : sheet.resolve_cell_value(row, col);
      if (value.is_text()) {
        source.text.emplace_back(value.as_text());
        value = Value::text(source.text.back());
      } else if (!value.is_number() && !value.is_boolean() && !value.is_error()) {
        value = Value::blank();
      }
      const CellXf* xf = xf_of(row, col);
      const std::uint32_t fmt = xf == nullptr ? 0U : static_cast<std::uint32_t>(xf->num_fmt_id);
      if (!fmt_seen[c]) {
        fmt_seen[c] = true;
        source.num_fmt_ids[c] = fmt;
      } else if (source.num_fmt_ids[c] != fmt) {
        fmt_mixed[c] = true;
      }
      bool is_date = false;
      if (value.is_number() && xf != nullptr) {
        auto it = date_format_of.find(fmt);
        if (it == date_format_of.end()) {
          it = date_format_of
                   .emplace(fmt, chart_is_date_format(text_format::number_format_code_for_xf(styles, xf, profile)))
                   .first;
        }
        is_date = it->second;
      }
      source.cells.push_back(value);
      source.cell_is_date.push_back(is_date);
    }
  }
  for (std::size_t c = 0; c < source.columns; ++c) {
    if (fmt_mixed[c]) {
      source.num_fmt_ids[c] = 0U;
    }
  }
  return source;
}

// `formula` with every reference to a field named in `dropped` spelled `#NAME?`.
std::string name_dropped_fields(const std::string& formula, const std::vector<std::string>& dropped) {
  parser::Tokenizer tokenizer(formula);
  std::string out;
  std::size_t done = 0;
  for (const parser::Token& token : tokenizer.tokens()) {
    if (token.kind != parser::TokenKind::Ident && token.kind != parser::TokenKind::SheetName) {
      continue;
    }
    const std::string_view name = token.kind == parser::TokenKind::SheetName ? token.text : token.lexeme;
    bool gone = false;
    for (const std::string& field : dropped) {
      gone = gone || pivot::same_header(name, field);
    }
    if (!gone) {
      continue;
    }
    const auto start = static_cast<std::size_t>(token.lexeme.data() - formula.data());
    out.append(formula, done, start - done);
    out.append("#NAME?");
    done = start + token.lexeme.size();
  }
  out.append(formula, done, std::string::npos);
  return out;
}

bool on_axis(pivot::PivotAxis axis) {
  return axis == pivot::PivotAxis::Row || axis == pivot::PivotAxis::Col || axis == pivot::PivotAxis::Page;
}

}  // namespace

Expected<void, Error> Workbook::check_pivot_growth(const std::vector<PivotGrowth>& growth) const {
  for (std::size_t i = 0; i < growth.size(); ++i) {
    const PivotGrowth& g = growth[i];
    if (g.sheet >= sheets_.size() || g.pivot >= sheets_[g.sheet].pivot_tables().size() ||
        sheets_[g.sheet].pivot_tables()[g.pivot] == nullptr) {
      return make_error(FormulonErrorCode::kInvalidArgument, "pivot_index is out of range",
                        "sheet=" + std::to_string(g.sheet) + " pivot=" + std::to_string(g.pivot));
    }
    if (!g.extent.has_value()) {
      continue;
    }
    const Sheet& sheet = sheets_[g.sheet];
    const MergeRange target = normalized(*g.extent);
    const MergeRange own = detail::pivot_occupied_extent(*sheet.pivot_tables()[g.pivot]);
    RETURN_IF_ERROR(detail::check_projection_target(*this, g.sheet, g.pivot, target, own));

    // The first non-empty cell, in row-major order, the report would newly cover.
    std::optional<std::pair<std::uint32_t, std::uint32_t>> blocking;
    for (const auto& [row, cells] : sheet.rows()) {
      if (row < target.first_row || row > target.last_row) {
        continue;
      }
      for (const auto& [col, cell] : cells) {
        if (col < target.first_col || col > target.last_col || inside(own, row, col) ||
            (cell.formula_text.empty() && cell.cached_value.is_blank())) {
          continue;
        }
        const std::pair<std::uint32_t, std::uint32_t> at{row, col};
        if (!blocking.has_value() || at < *blocking) {
          blocking = at;
        }
      }
    }
    if (blocking.has_value()) {
      return make_error(
          FormulonErrorCode::kPreconditionFailed,
          "The pivot report would overwrite non-empty cell " + cell_name(sheet, blocking->first, blocking->second));
    }

    for (std::size_t j = i + 1U; j < growth.size(); ++j) {
      const PivotGrowth& other = growth[j];
      if (other.sheet == g.sheet && other.extent.has_value() && intersects(target, normalized(*other.extent))) {
        return make_error(FormulonErrorCode::kPreconditionFailed,
                          "The pivot reports of " + sheet.pivot_tables()[g.pivot]->name() + " and " +
                              sheet.pivot_tables()[other.pivot]->name() + " would overlap");
      }
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::refresh_pivot_cache(std::uint32_t cache_id, const eval::FunctionRegistry& registry) {
  std::size_t slot = pivot_caches_.size();
  for (std::size_t i = 0; i < pivot_caches_.size(); ++i) {
    if (pivot_caches_[i] != nullptr && pivot_caches_[i]->cache_id() == cache_id) {
      slot = i;
      break;
    }
  }
  if (slot == pivot_caches_.size()) {
    return make_error(FormulonErrorCode::kNotFound, "cache_id does not name a pivot cache",
                      "cache_id=" + std::to_string(cache_id));
  }
  RETURN_IF_ERROR(ensure_calculated(registry));
  const pivot::PivotCache& old = *pivot_caches_[slot];
  ASSIGN_OR_RETURN(const pivot::WorksheetSourceRange range, pivot::resolve_worksheet_source(*this, old));
  const MergeRange rect = range.rect;
  const std::uint64_t record_cells = static_cast<std::uint64_t>(rect.last_row - rect.first_row) *
                                     static_cast<std::uint64_t>(rect.last_col - rect.first_col + 1U);
  if (record_cells > kMaxPivotCacheRecordCells) {
    return make_error(FormulonErrorCode::kSecurityResourceLimit, "The pivot cache source exceeds the record cell limit",
                      "cells=" + std::to_string(record_cells));
  }

  const eval::ScopedEvalProfile profile_scope(excel_profile_);
  const pivot::RefreshSource source = read_source(*this, sheets_[range.sheet], rect);
  pivot::RefreshContext context;
  context.axis_fields.assign(old.fields().size(), false);
  for (const pivot::PivotTable* table : pivots_of_cache(cache_id)) {
    for (std::size_t f = 0; f < table->fields().size() && f < context.axis_fields.size(); ++f) {
      context.axis_fields[f] = context.axis_fields[f] || on_axis(table->fields()[f].axis);
    }
  }
  context.blank_header_word = pivot::blank_header_field_word(excel_profile_);
  context.profile = excel_profile_;
  context.date1904 = date1904_;
  ASSIGN_OR_RETURN(pivot::RefreshPlan plan, pivot::build_refreshed_cache(old, source, context));

  // A calculated field naming a field the source lost names `#NAME?` instead.
  std::vector<std::string> dropped;
  for (std::size_t i = 0; i < old.fields().size(); ++i) {
    if (!plan.remap.fields[i].has_value() && old.fields()[i].is_database_field) {
      dropped.push_back(old.fields()[i].name);
    }
  }
  if (!dropped.empty()) {
    for (pivot::PivotCacheField& field : plan.cache.mutable_fields()) {
      if (field.formula.has_value()) {
        field.formula = name_dropped_fields(*field.formula, dropped);
      }
    }
  }
  pivot::materialize_group_fields(plan.cache, date1904_);

  // Phase 1: lay every dependent pivot out against the new cache, writing nothing.
  struct Planned {
    std::size_t sheet = 0;
    std::size_t pivot = 0;
    std::unique_ptr<pivot::PivotTable> table;
  };
  std::vector<Planned> planned;
  for (std::size_t s = 0; s < sheets_.size(); ++s) {
    const std::vector<std::unique_ptr<pivot::PivotTable>>& tables = sheets_[s].pivot_tables();
    for (std::size_t p = 0; p < tables.size(); ++p) {
      if (tables[p] != nullptr && tables[p]->pivot_cache_id() == cache_id) {
        std::unique_ptr<pivot::PivotTable> copy = tables[p]->clone();
        pivot::remap_table(*copy, plan.cache, plan.remap);
        pivot::clear_derived_state(*copy);
        planned.push_back(Planned{s, p, std::move(copy)});
      }
    }
  }
  // The first pivot placing a calculated item's field fixes its area position.
  for (auto it = planned.rbegin(); it != planned.rend(); ++it) {
    pivot::place_calculated_item_areas(plan.cache, *it->table);
  }
  const pivot::PivotLayoutOptions options = pivot::pivot_layout_options_for(excel_profile_);
  eval::PivotFormulaEvaluator formulas;
  std::vector<PivotGrowth> growth;
  for (const Planned& entry : planned) {
    auto result = pivot::evaluate(*entry.table, plan.cache, options,
                                  pivot::PivotFilterEnv{pinned_now(), date1904_, &styles_, &formulas});
    if (!result) {
      return std::move(result.error());
    }
    auto cells = pivot::layout(*entry.table, result.value(), options);
    if (!cells) {
      return std::move(cells.error());
    }
    PivotGrowth planned_growth;
    planned_growth.sheet = entry.sheet;
    planned_growth.pivot = entry.pivot;
    const pivot::PivotCells& laid = cells.value();
    if (laid.rows > 0U && laid.cols > 0U) {
      planned_growth.extent = MergeRange{laid.top, laid.left, laid.top + laid.rows - 1U, laid.left + laid.cols - 1U};
    }
    growth.push_back(std::move(planned_growth));
  }
  RETURN_IF_ERROR(check_pivot_growth(growth));

  // Phase 2: swap the cache and the definitions in, then project every report.
  pivot_caches_[slot] = std::make_unique<pivot::PivotCache>(std::move(plan.cache));
  pivot_caches_[slot]->mark_records_from_source();
  for (Planned& entry : planned) {
    pivot::PivotTable& table = *sheets_[entry.sheet].mutable_pivot_tables()[entry.pivot];
    // The copy is a clone, whose runtime id is zero; the table stays the same pivot.
    const std::uint32_t runtime_id = table.runtime_id().value();
    table = std::move(*entry.table);
    table.runtime_id().assign(runtime_id);
    pivot::clear_derived_state(table);
  }
  for (const Planned& entry : planned) {
    pivot::PivotTable& table = *sheets_[entry.sheet].mutable_pivot_tables()[entry.pivot];
    if (project_pivot(entry.sheet, entry.pivot)) {
      pivot::drop_derived_location(table);
    } else {
      table.set_projection_pending(true);
    }
  }
  mark_all_formulas_dirty();
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
