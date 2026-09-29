//
// Implementation of `DynamicReadLog`. See `dynamic_read_log.h`.

#include "eval/dynamic_read_log.h"

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>

#include "eval/cell_evaluator.h"
#include "eval/volatile_tracker.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace eval {

bool DynamicReadLog::observes(CellNodeId cell) const {
  return volatiles_.contains_dynamic_reference(cell) || graph_.has_dynamic_dependencies(cell);
}

void DynamicReadLog::observe(CellNodeId cell, std::uint64_t ordinal, EvaluateCellOptions& opts, const Sheet* sheet) {
  if (!observes(cell)) {
    return;
  }
  // The reader is registered even when it ends up reading nothing, so its
  // learned edges from an earlier recalc are dropped.
  readers_[cell] = ordinal;
  current_ = cell;
  current_ordinal_ = ordinal;
  active_ = true;
  opts.dynamic_read_callback = &DynamicReadLog::record_trampoline;
  opts.dynamic_read_user_data = this;
  if (sheet != nullptr) {
    Sheet::CellRead read;
    sheet->read_formula_cell(cell.row, cell.col, read);
    keep_prior_value(cell, read.value());
  }
}

void DynamicReadLog::record(std::string_view sheet, const DeclaredRect& rect) {
  if (!active_) {
    return;
  }
  const std::size_t sheet_id = sheet.empty() ? current_.sheet_id : workbook_.sheet_index_by_name(sheet);
  reads_.push_back(DynamicRead{current_, current_ordinal_, sheet_id, rect});
}

void DynamicReadLog::record_trampoline(void* user_data, std::string_view sheet, const DeclaredRect& rect) {
  static_cast<DynamicReadLog*>(user_data)->record(sheet, rect);
}

void DynamicReadLog::note_commit(CellNodeId cell, std::uint64_t ordinal) {
  commits_[cell] = ordinal;
}

void DynamicReadLog::note_iterative_member(CellNodeId cell, std::uint64_t component) {
  iterative_members_[cell] = component;
}

void DynamicReadLog::keep_prior_value(CellNodeId cell, const Value& prior) {
  if (prior_values_.count(cell) != 0U) {
    return;
  }
  switch (prior.kind()) {
    case ValueKind::Blank:
    case ValueKind::Number:
    case ValueKind::Bool:
    case ValueKind::Error:
      prior_values_.emplace(cell, PriorValue{prior, {}});
      return;
    case ValueKind::Text: {
      PriorValue& kept =
          prior_values_.emplace(cell, PriorValue{Value::blank(), std::string(prior.as_text())}).first->second;
      kept.value = Value::text(kept.text);
      return;
    }
    default:
      return;
  }
}

void DynamicReadLog::restore_prior_values(Workbook& workbook, const std::vector<CellNodeId>& cells) const {
  for (const CellNodeId c : cells) {
    if (c.sheet_id >= workbook.sheet_count()) {
      continue;
    }
    const auto kept = prior_values_.find(c);
    if (kept != prior_values_.end()) {
      workbook.sheet(c.sheet_id).set_cell_cached_value(c.row, c.col, kept->second.value);
    }
  }
}

void DynamicReadLog::clear_wave() {
  readers_.clear();
  reads_.clear();
  commits_.clear();
  iterative_members_.clear();
  active_ = false;
}

std::optional<std::uint64_t> DynamicReadLog::commit_ordinal(CellNodeId cell) const {
  const auto pos = commits_.find(cell);
  if (pos == commits_.end()) {
    return std::nullopt;
  }
  return pos->second;
}

std::optional<std::uint64_t> DynamicReadLog::iterative_component(CellNodeId cell) const {
  const auto pos = iterative_members_.find(cell);
  if (pos == iterative_members_.end()) {
    return std::nullopt;
  }
  return pos->second;
}

}  // namespace eval
}  // namespace formulon
