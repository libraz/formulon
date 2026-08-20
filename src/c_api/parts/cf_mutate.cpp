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
//
// Removals and clears also reconcile the sheet's raw worksheet-level
// `<extLst>` x14 overlay (see `io/cf_overlay.h`): the save path re-emits
// that overlay verbatim, so any `<x14:cfRule id="{GUID}">` whose legacy
// counterpart was just removed must be pruned here or the deleted rule
// resurfaces on reopen. Adds never remove a model rule and therefore
// cannot orphan an overlay entry.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "cf/cf_types.h"
#include "io/cf_overlay.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::BorrowedArrayArena;
using formulon::c_api::BorrowedStringArena;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::validate;

// `fm_cf_cell_range_t` mirrors `cf::CFCellRange`, so the model's range
// vector copies into the getter's arena as one memcpy and the wire type
// needs no per-field repacking.
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

// Materialises a `cf::CFRule` onto the wire-format `fm_cf_rule_t`.
// All non-engaged variant fields are zero-initialised first so callers
// observe deterministic defaults.
//
// Every pointer field points into the handle's CF arenas, never into the
// model. Pointing at the model would tie the record's lifetime to the
// block, and the ABI's only edit path removes that block while the caller
// still holds the record. Copying costs one pass over a rule's strings and
// ranges; it buys a record that only the next `fm_sheet_cf_get_at`
// invalidates. Each array is adopted as one finished block whose address
// does not move as later payloads are adopted, so a rule engaging more
// than one visual payload cannot invalidate a pointer already written
// into `*out`.
fm_cf_color_t from_cf_color(formulon::cf::Color color) {
  return fm_cf_color_t{color.r, color.g, color.b, color.a};
}

fm_cfvo_t from_cfvo(const formulon::cf::CfValueObject& src, BorrowedStringArena& text_arena) {
  fm_cfvo_t out{};
  out.type = static_cast<std::uint8_t>(src.type);
  out.gte = src.gte ? 1 : 0;
  if (!src.value.empty()) {
    out.value = text_arena.emplace(src.value);
  }
  return out;
}

void fill_rule(const formulon::cf::ConditionalFormat& block, const formulon::cf::CFRule& rule, std::size_t dxf_count,
               BorrowedStringArena& text_arena, BorrowedArrayArena<fm_cf_cell_range_t>& range_arena,
               BorrowedArrayArena<fm_cfvo_t>& cfvo_arena, BorrowedArrayArena<fm_cf_color_t>& color_arena,
               fm_cf_rule_t* out) {
  *out = fm_cf_rule_t{};
  out->id = text_arena.emplace(rule.id);
  out->type = static_cast<std::uint8_t>(rule.type);
  out->priority = rule.priority;
  out->stop_if_true = rule.stop_if_true ? 1 : 0;
  // A `dxf_id` the styles table cannot resolve is reported as absent
  // rather than handed out. The header calls this field an index into
  // `styles.dxfs[]`, so surfacing an unresolvable one would give the
  // caller a value `fm_styles_get_dxf` rejects and break the only edit
  // path the ABI offers - get a rule, remove it, add the amended record
  // back - since `fm_sheet_cf_add_rule` refuses the same index.
  if (rule.dxf_id.has_value() && static_cast<std::size_t>(*rule.dxf_id) < dxf_count) {
    out->dxf_id_engaged = 1;
    out->dxf_id = *rule.dxf_id;
  }
  // `cf::CFCellRange` and `fm_cf_cell_range_t` are layout-identical (see the
  // static_asserts above), so the block's ranges copy across as one memcpy.
  if (!block.sqref.empty()) {
    const auto* first = reinterpret_cast<const fm_cf_cell_range_t*>(block.sqref.data());
    out->sqref = range_arena.adopt(std::vector<fm_cf_cell_range_t>(first, first + block.sqref.size()));
  }
  out->sqref_count = static_cast<std::uint32_t>(block.sqref.size());
  out->formula1 = rule.formula1.has_value() ? text_arena.emplace(*rule.formula1) : nullptr;
  out->formula2 = rule.formula2.has_value() ? text_arena.emplace(*rule.formula2) : nullptr;
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
  out->text = rule.text.has_value() ? text_arena.emplace(*rule.text) : nullptr;
  if (rule.time_period.has_value()) {
    out->time_period_engaged = 1;
    out->time_period = static_cast<std::uint8_t>(*rule.time_period);
  }
  if (rule.color_scale.has_value()) {
    const auto& spec = *rule.color_scale;
    const std::size_t count = std::min(spec.thresholds.size(), spec.colors.size());
    std::vector<fm_cfvo_t> thresholds;
    std::vector<fm_cf_color_t> colors;
    thresholds.reserve(count);
    colors.reserve(count);
    for (std::size_t i = 0; i < count; ++i) {
      thresholds.push_back(from_cfvo(spec.thresholds[i], text_arena));
      colors.push_back(from_cf_color(spec.colors[i]));
    }
    out->color_scale_thresholds = cfvo_arena.adopt(std::move(thresholds));
    out->color_scale_colors = color_arena.adopt(std::move(colors));
    out->color_scale_count = static_cast<std::uint32_t>(count);
  }
  if (rule.data_bar.has_value()) {
    const auto& spec = *rule.data_bar;
    out->data_bar_engaged = 1;
    out->data_bar_min = from_cfvo(spec.min, text_arena);
    out->data_bar_max = from_cfvo(spec.max, text_arena);
    out->data_bar_fill = from_cf_color(spec.fill);
    out->data_bar_show_value = spec.show_value ? 1 : 0;
    out->data_bar_min_length_pct = spec.min_length_pct;
    out->data_bar_max_length_pct = spec.max_length_pct;
    // Engage every extension field unconditionally so this record, handed
    // straight back to `fm_sheet_cf_add_rule`, rebuilds the same spec. The
    // model has no "unset" state for gradient / axis / negative fill, so
    // reporting them as engaged is what makes the round trip exact.
    out->data_bar_gradient_engaged = 1;
    out->data_bar_gradient = spec.gradient ? 1 : 0;
    out->data_bar_axis_position_engaged = 1;
    out->data_bar_axis_position = static_cast<std::uint8_t>(spec.axis_position);
    out->data_bar_negative_fill_engaged = 1;
    out->data_bar_negative_fill = from_cf_color(spec.negative_fill);
    out->data_bar_border_engaged = spec.border.has_value() ? 1 : 0;
    if (spec.border.has_value()) {
      out->data_bar_border = from_cf_color(*spec.border);
    }
    out->data_bar_negative_border_engaged = spec.negative_border.has_value() ? 1 : 0;
    if (spec.negative_border.has_value()) {
      out->data_bar_negative_border = from_cf_color(*spec.negative_border);
    }
    out->data_bar_axis_color_engaged = 1;
    out->data_bar_axis_color = from_cf_color(spec.axis_color);
  }
  if (rule.icon_set.has_value()) {
    const auto& spec = *rule.icon_set;
    std::vector<fm_cfvo_t> thresholds;
    thresholds.reserve(spec.thresholds.size());
    for (const auto& threshold : spec.thresholds) {
      thresholds.push_back(from_cfvo(threshold, text_arena));
    }
    out->icon_set_engaged = 1;
    out->icon_set_name = static_cast<std::uint8_t>(spec.name);
    out->icon_set_thresholds = cfvo_arena.adopt(std::move(thresholds));
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

// The three payload copiers below check payload *shape* only - counts,
// pointers, and the length ordering the bar geometry needs. Every enum
// domain a payload carries (`fm_cfvo_t::type`, `icon_set_name`,
// `data_bar_axis_position`) is checked by `validate(rule, ...)`, which
// `fm_sheet_cf_add_rule` runs before it reaches here.
bool copy_color_scale_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.color_scale_count < 2 || rule.color_scale_count > 3 || rule.color_scale_thresholds == nullptr ||
      rule.color_scale_colors == nullptr) {
    return false;
  }
  formulon::cf::ColorScaleSpec spec;
  spec.thresholds.reserve(rule.color_scale_count);
  spec.colors.reserve(rule.color_scale_count);
  for (std::uint32_t i = 0; i < rule.color_scale_count; ++i) {
    spec.thresholds.push_back(to_cfvo(rule.color_scale_thresholds[i]));
    spec.colors.push_back(to_cf_color(rule.color_scale_colors[i]));
  }
  out_rule->color_scale = std::move(spec);
  return true;
}

bool copy_data_bar_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.data_bar_engaged == 0 || rule.data_bar_min_length_pct > 100 || rule.data_bar_max_length_pct > 100 ||
      rule.data_bar_min_length_pct > rule.data_bar_max_length_pct) {
    return false;
  }
  // Start from the model defaults and override only what the caller
  // engaged, so a zero-initialized record produces the same spec it did
  // before the extension fields existed.
  formulon::cf::DataBarSpec spec;
  spec.min = to_cfvo(rule.data_bar_min);
  spec.max = to_cfvo(rule.data_bar_max);
  spec.fill = to_cf_color(rule.data_bar_fill);
  spec.negative_fill = rule.data_bar_negative_fill_engaged != 0 ? to_cf_color(rule.data_bar_negative_fill) : spec.fill;
  if (rule.data_bar_border_engaged != 0) {
    spec.border = to_cf_color(rule.data_bar_border);
  }
  if (rule.data_bar_negative_border_engaged != 0) {
    spec.negative_border = to_cf_color(rule.data_bar_negative_border);
  }
  if (rule.data_bar_axis_position_engaged != 0) {
    spec.axis_position = static_cast<formulon::cf::DataBarAxisPosition>(rule.data_bar_axis_position);
  }
  if (rule.data_bar_axis_color_engaged != 0) {
    spec.axis_color = to_cf_color(rule.data_bar_axis_color);
  }
  if (rule.data_bar_gradient_engaged != 0) {
    spec.gradient = rule.data_bar_gradient != 0;
  }
  spec.show_value = rule.data_bar_show_value != 0;
  spec.min_length_pct = rule.data_bar_min_length_pct;
  spec.max_length_pct = rule.data_bar_max_length_pct;
  out_rule->data_bar = spec;
  return true;
}

bool copy_icon_set_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  if (rule.icon_set_engaged == 0 || (rule.icon_set_threshold_count > 0 && rule.icon_set_thresholds == nullptr)) {
    return false;
  }
  formulon::cf::IconSetSpec spec;
  spec.name = static_cast<formulon::cf::IconSetName>(rule.icon_set_name);
  spec.thresholds.reserve(rule.icon_set_threshold_count);
  for (std::uint32_t i = 0; i < rule.icon_set_threshold_count; ++i) {
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
  // The CF arenas are refreshed here and nowhere else, so the record stays
  // readable across CF mutations and across unrelated reads that recycle
  // `read_scratch`. Only a successful get reaches this point, so a rejected
  // call leaves the previous rule's storage intact.
  fm_workbook_t* mutable_wb = const_cast<fm_workbook_t*>(wb);
  mutable_wb->cf_text_scratch.clear();
  mutable_wb->cf_range_scratch.clear();
  mutable_wb->cfvo_scratch.clear();
  mutable_wb->cf_color_scratch.clear();
  fill_rule(blocks[b], blocks[b].rules[r], wb->workbook().styles().dxfs.size(), mutable_wb->cf_text_scratch,
            mutable_wb->cf_range_scratch, mutable_wb->cfvo_scratch, mutable_wb->cf_color_scratch, out);
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
  if (auto rc = validate(rule, "fm_sheet_cf_add_rule"); rc != 0) {
    return rc;
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
    // A `dxfId` past the end of `<dxfs>` makes Excel treat the whole
    // `xl/styles.xml` reference as broken and strip every conditional format
    // from the workbook on open. The writer already refuses to emit an
    // unresolvable `dxfId`, but rejecting it here tells the caller at set
    // time which index was wrong instead of silently losing the format.
    const std::size_t dxf_count = wb->workbook().styles().dxfs.size();
    if (static_cast<std::size_t>(rule.dxf_id) >= dxf_count) {
      return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument,
                               "fm_sheet_cf_add_rule: dxf_id out of range",
                               "dxf_id=" + std::to_string(rule.dxf_id) + " dxfs_count=" + std::to_string(dxf_count));
    }
    out_rule.dxf_id = rule.dxf_id;
  }
  if (rule.id != nullptr && rule.id[0] != '\0') {
    // Already checked against the `ST_Guid` shape by `validate` above.
    out_rule.id = rule.id;
  } else {
    // The id reaches the file as `<x14:cfRule id="...">`, whose schema
    // type is a GUID, so a synthesized one has to be GUID-shaped or
    // Excel rejects the extension block. It is derived from the priority
    // rather than drawn at random so add-then-list is deterministic and
    // two saves of the same model produce the same bytes. The version
    // nibble is `0`, which no random (version 4) GUID Excel generates
    // can carry, so a synthesized id can never collide with a loaded
    // one.
    char buf[40];
    std::snprintf(buf, sizeof(buf), "{FC000000-0000-0000-0000-%012X}", static_cast<unsigned>(out_rule.priority));
    out_rule.id = buf;
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
  auto& sheet = wb->workbook().sheet(sheet_index);
  auto& blocks = sheet.mutable_conditional_formats();
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
  // Keep the raw x14 overlay in step with the model so the writer's
  // verbatim re-emission cannot resurrect the removed rule.
  sheet.set_ext_lst_xml(formulon::io::reconcile_x14_cf_overlay(sheet.ext_lst_xml(), sheet.conditional_formats()));
  return 0;
}

extern "C" fm_status_t fm_sheet_cf_clear(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index, "fm_sheet_cf_clear"); rc != 0) {
    return rc;
  }
  auto& sheet = wb->workbook().sheet(sheet_index);
  sheet.mutable_conditional_formats().clear();
  // With no model rules left, every id-bearing x14 overlay entry is a
  // dangling reference; reconciliation prunes them all.
  sheet.set_ext_lst_xml(formulon::io::reconcile_x14_cf_overlay(sheet.ext_lst_xml(), sheet.conditional_formats()));
  return 0;
}
