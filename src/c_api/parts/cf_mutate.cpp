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

#include <algorithm>
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
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::TextStore;

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
fm_cf_color_t from_cf_color(formulon::cf::Color color) {
  return fm_cf_color_t{color.r, color.g, color.b, color.a};
}

fm_cfvo_t from_cfvo(const formulon::cf::CfValueObject& src, TextStore& text_store) {
  fm_cfvo_t out{};
  out.type = static_cast<std::uint8_t>(src.type);
  out.gte = src.gte ? 1 : 0;
  if (!src.value.empty()) {
    text_store.push_back(src.value);
    out.value = text_store.back().c_str();
  }
  return out;
}

void fill_rule(const formulon::cf::ConditionalFormat& block, const formulon::cf::CFRule& rule, TextStore& text_store,
               std::vector<fm_cfvo_t>& cfvo_scratch, std::vector<fm_cf_color_t>& color_scratch, fm_cf_rule_t* out) {
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
  if (rule.color_scale.has_value()) {
    const auto& spec = *rule.color_scale;
    const std::size_t count = std::min(spec.thresholds.size(), spec.colors.size());
    cfvo_scratch.reserve(cfvo_scratch.size() + count);
    color_scratch.reserve(color_scratch.size() + count);
    const std::size_t threshold_start = cfvo_scratch.size();
    const std::size_t color_start = color_scratch.size();
    for (std::size_t i = 0; i < count; ++i) {
      cfvo_scratch.push_back(from_cfvo(spec.thresholds[i], text_store));
      color_scratch.push_back(from_cf_color(spec.colors[i]));
    }
    out->color_scale_thresholds = count == 0 ? nullptr : cfvo_scratch.data() + threshold_start;
    out->color_scale_colors = count == 0 ? nullptr : color_scratch.data() + color_start;
    out->color_scale_count = static_cast<std::uint32_t>(count);
  }
  if (rule.data_bar.has_value()) {
    const auto& spec = *rule.data_bar;
    out->data_bar_engaged = 1;
    out->data_bar_min = from_cfvo(spec.min, text_store);
    out->data_bar_max = from_cfvo(spec.max, text_store);
    out->data_bar_fill = from_cf_color(spec.fill);
    out->data_bar_show_value = spec.show_value ? 1 : 0;
    out->data_bar_min_length_pct = spec.min_length_pct;
    out->data_bar_max_length_pct = spec.max_length_pct;
  }
  if (rule.icon_set.has_value()) {
    const auto& spec = *rule.icon_set;
    cfvo_scratch.reserve(cfvo_scratch.size() + spec.thresholds.size());
    const std::size_t threshold_start = cfvo_scratch.size();
    for (const auto& threshold : spec.thresholds) {
      cfvo_scratch.push_back(from_cfvo(threshold, text_store));
    }
    out->icon_set_engaged = 1;
    out->icon_set_name = static_cast<std::uint8_t>(spec.name);
    out->icon_set_thresholds = spec.thresholds.empty() ? nullptr : cfvo_scratch.data() + threshold_start;
    out->icon_set_threshold_count = static_cast<std::uint32_t>(spec.thresholds.size());
    out->icon_set_reverse = spec.reverse ? 1 : 0;
    out->icon_set_show_value = spec.show_value ? 1 : 0;
    out->icon_set_percent = spec.percent ? 1 : 0;
  }
}

formulon::cf::Color to_cf_color(fm_cf_color_t color) {
  return formulon::cf::Color{color.r, color.g, color.b, color.a};
}

formulon::cf::CfValueObject to_cfvo(const fm_cfvo_t& src) {
  formulon::cf::CfValueObject out;
  out.type = static_cast<formulon::cf::CfvoType>(src.type);
  out.value = src.value != nullptr ? std::string(src.value) : std::string();
  out.gte = src.gte != 0;
  return out;
}

bool is_valid_cfvo_type(std::uint8_t type) {
  return type <= static_cast<std::uint8_t>(formulon::cf::CfvoType::AutoMax);
}

bool copy_color_scale_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.color_scale_count < 2 || rule.color_scale_count > 3 || rule.color_scale_thresholds == nullptr ||
      rule.color_scale_colors == nullptr) {
    return false;
  }
  formulon::cf::ColorScaleSpec spec;
  spec.thresholds.reserve(rule.color_scale_count);
  spec.colors.reserve(rule.color_scale_count);
  for (std::uint32_t i = 0; i < rule.color_scale_count; ++i) {
    if (!is_valid_cfvo_type(rule.color_scale_thresholds[i].type)) {
      return false;
    }
    spec.thresholds.push_back(to_cfvo(rule.color_scale_thresholds[i]));
    spec.colors.push_back(to_cf_color(rule.color_scale_colors[i]));
  }
  out_rule->color_scale = std::move(spec);
  return true;
}

bool copy_data_bar_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.data_bar_engaged == 0 || !is_valid_cfvo_type(rule.data_bar_min.type) ||
      !is_valid_cfvo_type(rule.data_bar_max.type) || rule.data_bar_min_length_pct > 100 ||
      rule.data_bar_max_length_pct > 100) {
    return false;
  }
  formulon::cf::DataBarSpec spec;
  spec.min = to_cfvo(rule.data_bar_min);
  spec.max = to_cfvo(rule.data_bar_max);
  spec.fill = to_cf_color(rule.data_bar_fill);
  spec.negative_fill = spec.fill;
  spec.show_value = rule.data_bar_show_value != 0;
  spec.min_length_pct = rule.data_bar_min_length_pct;
  spec.max_length_pct = rule.data_bar_max_length_pct;
  out_rule->data_bar = spec;
  return true;
}

bool copy_icon_set_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.icon_set_engaged == 0 ||
      rule.icon_set_name > static_cast<std::uint8_t>(formulon::cf::IconSetName::Five_Quarters) ||
      (rule.icon_set_threshold_count > 0 && rule.icon_set_thresholds == nullptr)) {
    return false;
  }
  formulon::cf::IconSetSpec spec;
  spec.name = static_cast<formulon::cf::IconSetName>(rule.icon_set_name);
  spec.thresholds.reserve(rule.icon_set_threshold_count);
  for (std::uint32_t i = 0; i < rule.icon_set_threshold_count; ++i) {
    if (!is_valid_cfvo_type(rule.icon_set_thresholds[i].type)) {
      return false;
    }
    spec.thresholds.push_back(to_cfvo(rule.icon_set_thresholds[i]));
  }
  spec.reverse = rule.icon_set_reverse != 0;
  spec.show_value = rule.icon_set_show_value != 0;
  spec.percent = rule.icon_set_percent != 0;
  out_rule->icon_set = std::move(spec);
  return true;
}

}  // namespace

extern "C" fm_status_t fm_sheet_cf_count(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_count: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_count"); rc != 0) {
    return rc;
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
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_cf_get_at: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_get_at"); rc != 0) {
    return rc;
  }
  const auto& blocks = wb->workbook().sheet(sheet_index).conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (!resolve_flat_index(blocks, idx, &b, &r)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_sheet_cf_get_at: idx out of range",
                             "idx=" + std::to_string(idx));
  }
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->read_scratch.clear();
  mutable_wb->cfvo_scratch.clear();
  mutable_wb->cf_color_scratch.clear();
  fill_rule(blocks[b], blocks[b].rules[r], mutable_wb->read_scratch, mutable_wb->cfvo_scratch,
            mutable_wb->cf_color_scratch, out);
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_add_rule(fm_workbook_t* wb, std::size_t sheet_index, fm_cf_rule_t rule,
                                            std::size_t* out_index) {
  clear_last_error();
  if (out_index == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_cf_add_rule: out_index is NULL");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_add_rule"); rc != 0) {
    return rc;
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
  if (out_rule.type == formulon::cf::RuleType::ColorScale && !copy_color_scale_payload(rule, &out_rule)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: invalid colorScale payload");
  }
  if (out_rule.type == formulon::cf::RuleType::DataBar && !copy_data_bar_payload(rule, &out_rule)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: invalid dataBar payload");
  }
  if (out_rule.type == formulon::cf::RuleType::IconSet && !copy_icon_set_payload(rule, &out_rule)) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                             "fm_sheet_cf_add_rule: invalid iconSet payload");
  }

  new_block.rules.push_back(std::move(out_rule));
  auto& blocks_mut = wb->workbook().sheet(sheet_index).mutable_conditional_formats();
  // The new block is appended after every existing block, so its single
  // rule lands at the flattened index equal to the total rule count
  // observed just before the append.
  std::size_t new_index = 0;
  for (const auto& block : blocks_mut) {
    new_index += block.rules.size();
  }
  blocks_mut.push_back(std::move(new_block));
  *out_index = new_index;
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_remove_at(fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_remove_at"); rc != 0) {
    return rc;
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
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_clear"); rc != 0) {
    return rc;
  }
  wb->workbook().sheet(sheet_index).mutable_conditional_formats().clear();
  return 0;
}
