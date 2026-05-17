// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// C ABI - conditional-format mutation surface.
//
// Each rule lives inside a `<conditionalFormatting>` block. The mutation
// API exposes a flattened, per-rule view: `fm_sheet_cf_count` totals
// rules across all blocks, `fm_sheet_cf_get_at` reaches into the
// containing block to surface its sqref. Adds always create a fresh
// single-rule block to keep insertion order deterministic and avoid
// merging with semantically distinct sqref unions; removes prune the
// block too when its rule list goes empty.

#include <cstddef>
#include <cstdint>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cf/cf_types.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_range_count;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

// `fm_cf_cell_range_t` mirrors `cf::CFCellRange` so the OOXML reader's
// pre-allocated vector buffer can be handed back as a borrowed
// `const fm_cf_cell_range_t*` view without per-call repacking.
static_assert(sizeof(fm_cf_cell_range_t) == sizeof(formulon::cf::CFCellRange),
              "fm_cf_cell_range_t / cf::CFCellRange size mismatch");
static_assert(alignof(fm_cf_cell_range_t) == alignof(formulon::cf::CFCellRange),
              "fm_cf_cell_range_t / cf::CFCellRange align mismatch");
static_assert(offsetof(fm_cf_cell_range_t, first_row) == offsetof(formulon::cf::CFCellRange, first),
              "fm_cf_cell_range_t::first_row layout mismatch");
static_assert(offsetof(fm_cf_cell_range_t, last_row) == offsetof(formulon::cf::CFCellRange, last),
              "fm_cf_cell_range_t::last_row layout mismatch");

namespace {

// Returns `true` for the three visual rule types whose payloads
// (color_scale / data_bar / icon_set sub-specs) are not yet creatable
// through the C ABI. The OOXML reader / writer still round-trip them
// verbatim - this gate only fires on the mutation entry point.
bool is_visual_rule_type(std::uint8_t type) {
  switch (static_cast<formulon::cf::RuleType>(type)) {
    case formulon::cf::RuleType::ColorScale:
    case formulon::cf::RuleType::DataBar:
    case formulon::cf::RuleType::IconSet:
      return true;
    default:
      return false;
  }
}

// Walks the sheet's `conditional_formats` vector and resolves the
// `flat_idx`-th rule into the (block_idx, rule_idx) pair. Returns
// `true` on success; `false` when `flat_idx` is past the end.
bool resolve_flat_index(const std::vector<formulon::cf::ConditionalFormat>& blocks, std::size_t flat_idx,
                        std::size_t* out_block, std::size_t* out_rule) {
  std::size_t cursor = 0;
  for (std::size_t b = 0; b < blocks.size(); ++b) {
    const auto& block = blocks[b];
    if (flat_idx < cursor + block.rules.size()) {
      *out_block = b;
      *out_rule = flat_idx - cursor;
      return true;
    }
    cursor += block.rules.size();
  }
  return false;
}

// Materialises a `cf::CFRule` view onto the wire-format `fm_cf_rule_t`.
// All non-engaged variant fields are zero-initialised first so callers
// observe deterministic defaults. String views borrow the engine's
// storage; the contract documented in the header is "valid until the
// next CF mutation".
void fill_rule(const formulon::cf::ConditionalFormat& block, const formulon::cf::CFRule& rule, fm_cf_rule_t* out) {
  *out = fm_cf_rule_t{};
  out->id = rule.id.c_str();
  out->type = static_cast<std::uint8_t>(rule.type);
  out->priority = rule.priority;
  out->stop_if_true = rule.stop_if_true ? 1 : 0;
  if (rule.dxf_id.has_value()) {
    out->dxf_id_engaged = 1;
    out->dxf_id = *rule.dxf_id;
  }
  out->sqref = block.sqref.empty() ? nullptr : reinterpret_cast<const fm_cf_cell_range_t*>(block.sqref.data());
  out->sqref_count = static_cast<std::uint32_t>(block.sqref.size());
  out->formula1 = rule.formula1.has_value() ? rule.formula1->c_str() : nullptr;
  out->formula2 = rule.formula2.has_value() ? rule.formula2->c_str() : nullptr;
  if (rule.op.has_value()) {
    out->op_engaged = 1;
    out->op = static_cast<std::uint8_t>(*rule.op);
  }
  if (rule.rank.has_value()) {
    out->rank_engaged = 1;
    out->rank = *rule.rank;
  }
  out->percent = rule.percent ? 1 : 0;
  out->bottom = rule.bottom ? 1 : 0;
  out->above_average = rule.above_average ? 1 : 0;
  out->equal_average = rule.equal_average ? 1 : 0;
  if (rule.std_dev.has_value()) {
    out->std_dev_engaged = 1;
    out->std_dev = *rule.std_dev;
  }
  out->text = rule.text.has_value() ? rule.text->c_str() : nullptr;
  if (rule.time_period.has_value()) {
    out->time_period_engaged = 1;
    out->time_period = static_cast<std::uint8_t>(*rule.time_period);
  }
}

}  // namespace

extern "C" fm_status_t fm_sheet_cf_count(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t* out_count) {
  clear_last_error();
  if (wb == nullptr || out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_count: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_count: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  std::size_t total = 0;
  for (const auto& block : wb->workbook().sheet(sheet_index).conditional_formats()) {
    total += block.rules.size();
  }
  *out_count = total;
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_get_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx,
                                          fm_cf_rule_t* out) {
  clear_last_error();
  if (wb == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_get_at: NULL argument");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_get_at: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  const auto& blocks = wb->workbook().sheet(sheet_index).conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (!resolve_flat_index(blocks, idx, &b, &r)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_get_at: idx out of range",
                             "idx=" + std::to_string(idx));
  }
  fill_rule(blocks[b], blocks[b].rules[r], out);
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_add_rule(fm_workbook_t* wb, std::size_t sheet_index, fm_cf_rule_t rule) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_add_rule: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  if (rule.sqref_count == 0) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: sqref_count must be >= 1");
  }
  if (rule.sqref == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_cf_add_rule: sqref is NULL while sqref_count > 0");
  }
  if (!check_range_count(rule.sqref_count, "fm_sheet_cf_add_rule")) {
    return static_cast<fm_status_t>(formulon::FormulonErrorCode::kInvalidArgument);
  }
  if (is_visual_rule_type(rule.type)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: visual rule types (ColorScale/DataBar/IconSet) "
                             "are not creatable through this API",
                             "type=" + std::to_string(rule.type));
  }
  if (rule.type > static_cast<std::uint8_t>(formulon::cf::RuleType::UniqueValues)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_add_rule: unknown rule type",
                             "type=" + std::to_string(rule.type));
  }

  formulon::cf::ConditionalFormat new_block;
  new_block.sqref.reserve(rule.sqref_count);
  for (std::uint32_t i = 0; i < rule.sqref_count; ++i) {
    formulon::cf::CFCellRange r;
    r.first = formulon::CellAddress{rule.sqref[i].first_row, rule.sqref[i].first_col};
    r.last = formulon::CellAddress{rule.sqref[i].last_row, rule.sqref[i].last_col};
    new_block.sqref.push_back(r);
  }

  formulon::cf::CFRule out_rule;
  out_rule.type = static_cast<formulon::cf::RuleType>(rule.type);

  // Auto-assign priority to (max_existing + 1) when caller passed <= 0.
  std::int32_t max_priority = 0;
  for (const auto& block : wb->workbook().sheet(sheet_index).conditional_formats()) {
    for (const auto& existing : block.rules) {
      if (existing.priority > max_priority) {
        max_priority = existing.priority;
      }
    }
  }
  out_rule.priority = (rule.priority > 0) ? rule.priority : (max_priority + 1);
  out_rule.stop_if_true = rule.stop_if_true != 0;
  if (rule.dxf_id_engaged != 0) {
    out_rule.dxf_id = rule.dxf_id;
  }
  if (rule.id != nullptr && rule.id[0] != '\0') {
    out_rule.id = rule.id;
  } else {
    // Synthesize a stable id from priority. The format mirrors the
    // x14:cfRule guid-like string Excel emits, but uses a priority
    // suffix so add-then-list is deterministic.
    out_rule.id = "{cf-" + std::to_string(out_rule.priority) + "}";
  }
  if (rule.formula1 != nullptr) {
    out_rule.formula1 = std::string(rule.formula1);
  }
  if (rule.formula2 != nullptr) {
    out_rule.formula2 = std::string(rule.formula2);
  }
  if (rule.op_engaged != 0) {
    out_rule.op = static_cast<formulon::cf::CellIsOperator>(rule.op);
  }
  if (rule.rank_engaged != 0) {
    out_rule.rank = rule.rank;
  }
  out_rule.percent = (rule.percent != 0);
  out_rule.bottom = (rule.bottom != 0);
  out_rule.above_average = (rule.above_average != 0);
  out_rule.equal_average = (rule.equal_average != 0);
  if (rule.std_dev_engaged != 0) {
    out_rule.std_dev = rule.std_dev;
  }
  if (rule.text != nullptr) {
    out_rule.text = std::string(rule.text);
  }
  if (rule.time_period_engaged != 0) {
    out_rule.time_period = static_cast<formulon::cf::TimePeriod>(rule.time_period);
  }

  new_block.rules.push_back(std::move(out_rule));
  wb->workbook().sheet(sheet_index).mutable_conditional_formats().push_back(std::move(new_block));
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_remove_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_remove_at: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_remove_at: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  auto& blocks = wb->workbook().sheet(sheet_index).mutable_conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (!resolve_flat_index(blocks, idx, &b, &r)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_remove_at: idx out of range",
                             "idx=" + std::to_string(idx));
  }
  blocks[b].rules.erase(blocks[b].rules.begin() + static_cast<std::ptrdiff_t>(r));
  if (blocks[b].rules.empty()) {
    blocks.erase(blocks.begin() + static_cast<std::ptrdiff_t>(b));
  }
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_clear(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (wb == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_clear: wb is NULL");
  }
  if (sheet_index >= wb->workbook().sheet_count()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_clear: sheet_index out of range",
                             "sheet_index=" + std::to_string(sheet_index));
  }
  wb->workbook().sheet(sheet_index).mutable_conditional_formats().clear();
  return 0;
}
