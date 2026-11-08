//
// C ABI - conditional-format mutation surface.
//
// Each rule lives inside a `<conditionalFormatting>` block. The mutation
// API exposes a flattened, per-rule view: `fm_sheet_get_cf_rule_count` totals
// rules across all blocks, `fm_sheet_get_cf_rule` reaches into the
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
#include <limits>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/enum_map.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "cf/cf_authoring.h"
#include "cf/cf_types.h"
#include "cf/cf_x14_id.h"
#include "io/cf_overlay.h"
#include "parser/formula_prefix.h"
#include "utils/error.h"
#include "utils/strings.h"
#include "workbook.h"

using formulon::c_api::BorrowedArrayArena;
using formulon::c_api::BorrowedStringArena;
using formulon::c_api::parts::cf_operator_from_c;
using formulon::c_api::parts::check_enum_domain;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::check_stored_formula;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::color_spec_from_c;
using formulon::c_api::parts::color_spec_to_c;
using formulon::c_api::parts::data_bar_axis_position_from_c;
using formulon::c_api::parts::data_bar_direction_from_c;
using formulon::c_api::parts::date_period_from_c;
using formulon::c_api::parts::fresh_arena;
using formulon::c_api::parts::ingest_entered_formula;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::kIconSetNameTokens;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::ordinal_of;
using formulon::c_api::parts::text_or_empty;
using formulon::c_api::parts::to_c;
using formulon::c_api::parts::token_of;
using formulon::c_api::parts::validate;
using formulon::c_api::parts::with_formula_prefix;

// `fm_cell_rect_t` mirrors `cf::CFCellRange`, so the model's range
// vector copies into the getter's arena as one memcpy and the wire type
// needs no per-field repacking.
static_assert(sizeof(fm_cell_rect_t) == sizeof(formulon::cf::CFCellRange),
              "fm_cell_rect_t / cf::CFCellRange size mismatch");
static_assert(alignof(fm_cell_rect_t) == alignof(formulon::cf::CFCellRange),
              "fm_cell_rect_t / cf::CFCellRange align mismatch");
static_assert(offsetof(fm_cell_rect_t, first_row) == offsetof(formulon::cf::CFCellRange, first),
              "cell rect field first_row layout mismatch");
static_assert(offsetof(fm_cell_rect_t, last_row) == offsetof(formulon::cf::CFCellRange, last),
              "cell rect field last_row layout mismatch");

namespace {

// Walks the sheet's `conditional_formats` vector and resolves the
// `flat_index`-th rule into the (block_index, rule_index) pair. Returns
// 0 on success; the `check_index` error when `flat_index` is past the end.
fm_status_t resolve_flat_index(const std::vector<formulon::cf::ConditionalFormat>& blocks, std::size_t flat_index,
                               std::size_t* out_block, std::size_t* out_rule) {
  std::size_t cursor = 0;
  for (std::size_t b = 0; b < blocks.size(); ++b) {
    const auto& block = blocks[b];
    if (flat_index < cursor + block.rules.size()) {
      *out_block = b;
      *out_rule = flat_index - cursor;
      return 0;
    }
    cursor += block.rules.size();
  }
  return check_index(flat_index, cursor, "index");
}

// Materialises a `cf::CFRule` onto the wire-format `fm_cf_rule_t`.
// All non-engaged variant fields are zero-initialised first so callers
// observe deterministic defaults.
//
// Every pointer field points into the handle's CF arenas, never into the
// model. Pointing at the model would tie the record's lifetime to the
// block, and the ABI's only edit path removes that block while the caller
// still holds the record. Copying costs one pass over a rule's strings and
// ranges; it buys a record that only the next `fm_sheet_get_cf_rule`
// invalidates. Each array is adopted as one finished block whose address
// does not move as later payloads are adopted, so a rule engaging more
// than one visual payload cannot invalidate a pointer already written
// into `*out`.
// A theme / indexed / auto colour keeps its kind; a literal one reads as
// `FM_COLOR_KIND_RGB`, or `FM_COLOR_KIND_NONE` when all four channels are zero.
fm_color_spec_t cf_color_to_c(const formulon::cf::Color& color) {
  if (color.is_symbolic()) {
    return color_spec_to_c(color.spec);
  }
  const std::uint32_t argb = (static_cast<std::uint32_t>(color.a) << 24U) |
                             (static_cast<std::uint32_t>(color.r) << 16U) |
                             (static_cast<std::uint32_t>(color.g) << 8U) | static_cast<std::uint32_t>(color.b);
  return color_spec_to_c(formulon::ColorSpec{}, argb);
}

fm_cfvo_t from_cfvo(const formulon::Workbook& wb, const formulon::cf::CfValueObject& src,
                    BorrowedStringArena& text_arena) {
  fm_cfvo_t out{};
  out.type = static_cast<std::uint8_t>(src.type);
  out.gte = src.gte ? 1 : 0;
  if (!src.value.empty()) {
    out.value = text_arena.emplace(
        src.type == formulon::cf::CfvoType::Formula
            ? with_formula_prefix(formulon::parser::display_cell_shaped_names(wb.spell_self_book_names(src.value)))
            : src.value);
  }
  return out;
}

void fill_rule(const formulon::Workbook& wb, const formulon::cf::ConditionalFormat& block,
               const formulon::cf::CFRule& rule, std::size_t dxf_count, BorrowedStringArena& text_arena,
               BorrowedArrayArena<fm_cell_rect_t>& range_arena, BorrowedArrayArena<fm_cfvo_t>& cfvo_arena,
               BorrowedArrayArena<fm_color_spec_t>& color_arena, BorrowedArrayArena<fm_cf_icon_t>& icon_arena,
               fm_cf_rule_t* out) {
  const std::uint32_t struct_size = out->struct_size;
  *out = fm_cf_rule_t{};
  out->struct_size = struct_size;
  out->id = text_arena.emplace(rule.id);
  out->type = static_cast<std::uint8_t>(rule.type);
  out->priority = rule.priority;
  out->stop_if_true = rule.stop_if_true ? 1 : 0;
  // A `dxf_id` the styles table cannot resolve is reported as absent
  // rather than handed out. The header calls this field an index into
  // `styles.dxfs[]`, so surfacing an unresolvable one would give the
  // caller a value `fm_workbook_get_dxf` rejects and break the only edit
  // path the ABI offers - get a rule, remove it, add the amended record
  // back - since `fm_sheet_add_cf_rule` refuses the same index.
  if (rule.dxf_id.has_value() && static_cast<std::size_t>(*rule.dxf_id) < dxf_count) {
    out->dxf_index_engaged = 1;
    out->dxf_index = *rule.dxf_id;
  }
  // `cf::CFCellRange` and `fm_cell_rect_t` are layout-identical (see the
  // static_asserts above), so the block's ranges copy across as one memcpy.
  if (!block.sqref.empty()) {
    const auto* first = reinterpret_cast<const fm_cell_rect_t*>(block.sqref.data());
    out->ranges = range_arena.adopt(std::vector<fm_cell_rect_t>(first, first + block.sqref.size()));
  }
  out->range_count = static_cast<std::uint32_t>(block.sqref.size());
  out->formula1 = rule.formula1.has_value()
                      ? text_arena.emplace(with_formula_prefix(
                            formulon::parser::display_cell_shaped_names(wb.spell_self_book_names(*rule.formula1))))
                      : nullptr;
  out->formula2 = rule.formula2.has_value()
                      ? text_arena.emplace(with_formula_prefix(
                            formulon::parser::display_cell_shaped_names(wb.spell_self_book_names(*rule.formula2))))
                      : nullptr;
  if (rule.op.has_value()) {
    out->op_engaged = 1;
    out->op = static_cast<std::uint8_t>(to_c(*rule.op));
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
    out->time_period = static_cast<std::uint8_t>(to_c(*rule.time_period));
  }
  if (rule.color_scale.has_value()) {
    const auto& spec = *rule.color_scale;
    const std::size_t count = std::min(spec.thresholds.size(), spec.colors.size());
    std::vector<fm_cfvo_t> thresholds;
    std::vector<fm_color_spec_t> colors;
    thresholds.reserve(count);
    colors.reserve(count);
    for (std::size_t i = 0; i < count; ++i) {
      thresholds.push_back(from_cfvo(wb, spec.thresholds[i], text_arena));
      colors.push_back(cf_color_to_c(spec.colors[i]));
    }
    out->color_scale_thresholds = cfvo_arena.adopt(std::move(thresholds));
    out->color_scale_colors = color_arena.adopt(std::move(colors));
    out->color_scale_count = static_cast<std::uint32_t>(count);
  }
  if (rule.data_bar.has_value()) {
    const auto& spec = *rule.data_bar;
    out->data_bar_engaged = 1;
    out->data_bar_min = from_cfvo(wb, spec.min, text_arena);
    out->data_bar_max = from_cfvo(wb, spec.max, text_arena);
    out->data_bar_fill = cf_color_to_c(spec.fill);
    out->data_bar_show_value = spec.show_value ? 1 : 0;
    out->data_bar_min_length_pct = spec.min_length_pct;
    out->data_bar_max_length_pct = spec.max_length_pct;
    // Engage every extension field unconditionally so this record, handed
    // straight back to `fm_sheet_add_cf_rule`, rebuilds the same spec. The
    // model has no "unset" state for gradient / axis / negative fill, so
    // reporting them as engaged is what makes the round trip exact.
    out->data_bar_gradient_engaged = 1;
    out->data_bar_gradient = spec.gradient ? 1 : 0;
    out->data_bar_axis_position_engaged = 1;
    out->data_bar_axis_position = to_c(spec.axis_position);
    out->data_bar_negative_fill_engaged = 1;
    out->data_bar_negative_fill = cf_color_to_c(spec.negative_fill);
    out->data_bar_border_engaged = spec.border.has_value() ? 1 : 0;
    if (spec.border.has_value()) {
      out->data_bar_border = cf_color_to_c(*spec.border);
    }
    out->data_bar_negative_border_engaged = spec.negative_border.has_value() ? 1 : 0;
    if (spec.negative_border.has_value()) {
      out->data_bar_negative_border = cf_color_to_c(*spec.negative_border);
    }
    out->data_bar_axis_color_engaged = 1;
    out->data_bar_axis_color = cf_color_to_c(spec.axis_color);
    out->data_bar_direction = to_c(spec.direction);
  }
  if (rule.icon_set.has_value()) {
    const auto& spec = *rule.icon_set;
    std::vector<fm_cfvo_t> thresholds;
    thresholds.reserve(spec.thresholds.size());
    for (const auto& threshold : spec.thresholds) {
      thresholds.push_back(from_cfvo(wb, threshold, text_arena));
    }
    out->icon_set_engaged = 1;
    out->icon_set_name = static_cast<std::uint8_t>(spec.name);
    out->icon_set_thresholds = cfvo_arena.adopt(std::move(thresholds));
    out->icon_set_threshold_count = static_cast<std::uint32_t>(spec.thresholds.size());
    out->icon_set_reverse = spec.reverse ? 1 : 0;
    out->icon_set_show_value = spec.show_value ? 1 : 0;
    out->icon_set_percent = spec.percent ? 1 : 0;
    out->icon_set_floor_engaged = 1;
    out->icon_set_floor = from_cfvo(wb, spec.floor, text_arena);
    std::vector<fm_cf_icon_t> icons;
    icons.reserve(spec.custom_icons.size());
    for (const auto& icon : spec.custom_icons) {
      icons.push_back(fm_cf_icon_t{ordinal_of(kIconSetNameTokens, icon.icon_set, -1), icon.icon_id});
    }
    out->icon_set_icon_count = static_cast<std::uint32_t>(icons.size());
    out->icon_set_icons = icon_arena.adopt(std::move(icons));
  }
}

// Inverse of `cf_color_to_c`: `FM_COLOR_KIND_NONE` is the literal colour with all
// four channels zero. The kind is checked by `check_payload_colors`.
formulon::cf::Color cf_color_from_c(const fm_color_spec_t& src) {
  std::uint32_t argb = 0;
  const formulon::ColorSpec spec = color_spec_from_c(src, &argb);
  formulon::cf::Color out;
  out.spec = spec;
  if (out.is_symbolic()) {
    return out;
  }
  out.spec = formulon::ColorSpec{};
  out.a = static_cast<std::uint8_t>((argb >> 24U) & 0xFFU);
  out.r = static_cast<std::uint8_t>((argb >> 16U) & 0xFFU);
  out.g = static_cast<std::uint8_t>((argb >> 8U) & 0xFFU);
  out.b = static_cast<std::uint8_t>(argb & 0xFFU);
  return out;
}

fm_status_t check_color_kind(const fm_color_spec_t& color, const char* field) {
  return check_enum_domain(color.kind, FM_COLOR_KIND_NONE, FM_COLOR_KIND_AUTO, field);
}

// Rejects a colour of the engaged payloads whose `kind` names no
// `fm_color_kind_t`. The colour-scale array is scanned only in the shape
// `copy_color_scale_payload` accepts.
fm_status_t check_payload_colors(const fm_cf_rule_t& rule) {
  if (rule.color_scale_colors != nullptr && rule.color_scale_count >= 2U && rule.color_scale_count <= 3U) {
    for (std::uint32_t i = 0; i < rule.color_scale_count; ++i) {
      if (auto rc = check_color_kind(rule.color_scale_colors[i], "color_scale_colors[].kind"); rc != 0) {
        return rc;
      }
    }
  }
  if (rule.data_bar_engaged == 0) {
    return 0;
  }
  if (auto rc = check_color_kind(rule.data_bar_fill, "data_bar_fill.kind"); rc != 0) {
    return rc;
  }
  if (rule.data_bar_negative_fill_engaged != 0) {
    if (auto rc = check_color_kind(rule.data_bar_negative_fill, "data_bar_negative_fill.kind"); rc != 0) {
      return rc;
    }
  }
  if (rule.data_bar_border_engaged != 0) {
    if (auto rc = check_color_kind(rule.data_bar_border, "data_bar_border.kind"); rc != 0) {
      return rc;
    }
  }
  if (rule.data_bar_negative_border_engaged != 0) {
    if (auto rc = check_color_kind(rule.data_bar_negative_border, "data_bar_negative_border.kind"); rc != 0) {
      return rc;
    }
  }
  if (rule.data_bar_axis_color_engaged != 0) {
    return check_color_kind(rule.data_bar_axis_color, "data_bar_axis_color.kind");
  }
  return 0;
}

formulon::cf::CfValueObject to_cfvo(const fm_cfvo_t& src) {
  formulon::cf::CfValueObject out;
  out.type = static_cast<formulon::cf::CfvoType>(src.type);
  out.value = text_or_empty(src.value);
  out.gte = src.gte != 0;
  return out;
}

// The three payload copiers below check payload *shape* only - counts,
// pointers, and the length ordering the bar geometry needs. Every enum
// domain a payload carries (`fm_cfvo_t::type`, `icon_set_name`,
// `data_bar_axis_position`, `data_bar_direction`) is checked by
// `validate(rule, ...)`, which
// `fm_sheet_add_cf_rule` runs before it reaches here.
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
    spec.colors.push_back(cf_color_from_c(rule.color_scale_colors[i]));
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
  spec.fill = cf_color_from_c(rule.data_bar_fill);
  spec.negative_fill =
      rule.data_bar_negative_fill_engaged != 0 ? cf_color_from_c(rule.data_bar_negative_fill) : spec.fill;
  if (rule.data_bar_border_engaged != 0) {
    spec.border = cf_color_from_c(rule.data_bar_border);
  }
  if (rule.data_bar_negative_border_engaged != 0) {
    spec.negative_border = cf_color_from_c(rule.data_bar_negative_border);
  }
  if (rule.data_bar_axis_position_engaged != 0) {
    spec.axis_position = data_bar_axis_position_from_c(rule.data_bar_axis_position);
  }
  if (rule.data_bar_axis_color_engaged != 0) {
    spec.axis_color = cf_color_from_c(rule.data_bar_axis_color);
  }
  if (rule.data_bar_gradient_engaged != 0) {
    spec.gradient = rule.data_bar_gradient != 0;
  }
  spec.direction = data_bar_direction_from_c(rule.data_bar_direction);
  spec.show_value = rule.data_bar_show_value != 0;
  spec.min_length_pct = rule.data_bar_min_length_pct;
  spec.max_length_pct = rule.data_bar_max_length_pct;
  out_rule->data_bar = spec;
  return true;
}

bool copy_icon_set_payload(const fm_cf_rule_t& rule, formulon::cf::CFRule* out_rule) {
  // An N-icon set carries exactly its N-1 boundaries; the floor is separate.
  if (rule.icon_set_engaged == 0 || rule.icon_set_thresholds == nullptr ||
      rule.icon_set_threshold_count + 1U !=
          formulon::cf::icon_count(static_cast<formulon::cf::IconSetName>(rule.icon_set_name))) {
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
  if (rule.icon_set_floor_engaged != 0) {
    spec.floor = to_cfvo(rule.icon_set_floor);
  }
  // Shape and domain of the icons are checked by `validate(rule, ...)`.
  if (rule.icon_set_icons != nullptr) {
    spec.custom_icons.reserve(rule.icon_set_icon_count);
    for (std::uint32_t i = 0; i < rule.icon_set_icon_count; ++i) {
      spec.custom_icons.push_back(
          formulon::cf::CfIcon{std::string(token_of(kIconSetNameTokens, rule.icon_set_icons[i].icon_set)),
                               rule.icon_set_icons[i].icon_index});
    }
  }
  out_rule->icon_set = std::move(spec);
  return true;
}

// Checks every formula-typed threshold of `out_rule`'s visual payload, then
// stores each the way files carry it (see `ingest_entered_formula`).
fm_status_t store_cfvo_formulas(fm_workbook_t* wb, formulon::cf::CFRule* out_rule) {
  // Each threshold with the `field=<name>` its parse failure reports.
  std::vector<std::pair<formulon::cf::CfValueObject*, const char*>> cfvos;
  if (out_rule->color_scale.has_value()) {
    for (auto& v : out_rule->color_scale->thresholds) {
      cfvos.emplace_back(&v, "field=color_scale_threshold");
    }
  }
  if (out_rule->data_bar.has_value()) {
    cfvos.emplace_back(&out_rule->data_bar->min, "field=data_bar_min");
    cfvos.emplace_back(&out_rule->data_bar->max, "field=data_bar_max");
  }
  if (out_rule->icon_set.has_value()) {
    cfvos.emplace_back(&out_rule->icon_set->floor, "field=icon_floor");
    for (auto& v : out_rule->icon_set->thresholds) {
      cfvos.emplace_back(&v, "field=icon_threshold");
    }
  }
  for (const auto& [v, context] : cfvos) {
    if (v->type != formulon::cf::CfvoType::Formula) {
      continue;
    }
    std::string stored;
    if (auto rc = check_stored_formula(v->value, context, &stored); rc != 0) {
      return rc;
    }
  }
  for (const auto& entry : cfvos) {
    if (entry.first->type == formulon::cf::CfvoType::Formula) {
      entry.first->value = ingest_entered_formula(wb, entry.first->value);
    }
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_cf_rule_count(const fm_workbook_t* wb, std::size_t sheet_index,
                                                  std::size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  std::size_t total = 0;
  for (const auto& block : wb->workbook().sheet(sheet_index).conditional_formats()) {
    total += block.rules.size();
  }
  *out_count = total;
  return 0;
}

extern "C" fm_status_t fm_sheet_get_cf_rule(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx,
                                            fm_cf_rule_t* out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  const auto& blocks = wb->workbook().sheet(sheet_index).conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (auto rc = resolve_flat_index(blocks, idx, &b, &r); rc != 0) {
    return rc;
  }
  // The CF arenas are refreshed here and nowhere else, so the record stays
  // readable across CF mutations and across unrelated reads that recycle
  // `read_scratch`. Only a successful get reaches this point, so a rejected
  // call leaves the previous rule's storage intact.
  auto& text_arena = fresh_arena(wb, &fm_workbook_t::cf_text_scratch);
  auto& range_arena = fresh_arena(wb, &fm_workbook_t::cf_range_scratch);
  auto& cfvo_arena = fresh_arena(wb, &fm_workbook_t::cfvo_scratch);
  auto& color_arena = fresh_arena(wb, &fm_workbook_t::cf_color_scratch);
  auto& icon_arena = fresh_arena(wb, &fm_workbook_t::cf_icon_scratch);
  fill_rule(wb->workbook(), blocks[b], blocks[b].rules[r], wb->workbook().styles().dxfs.size(), text_arena, range_arena,
            cfvo_arena, color_arena, icon_arena, out);
  return 0;
}

namespace {

// The rule's priority: the caller's when positive, else one past the sheet's
// highest.
fm_status_t assign_priority(const fm_cf_rule_t& rule, const formulon::Sheet& sheet, formulon::cf::CFRule* out) {
  if (rule.priority > 0) {
    out->priority = rule.priority;
    return 0;
  }
  std::int32_t max_priority = 0;
  for (const auto& block : sheet.conditional_formats()) {
    for (const auto& existing : block.rules) {
      if (existing.priority > max_priority) {
        max_priority = existing.priority;
      }
    }
  }
  if (max_priority == std::numeric_limits<std::int32_t>::max()) {
    return invalid_argument("auto-assignment would overflow", "max_existing_priority=" + std::to_string(max_priority));
  }
  out->priority = max_priority + 1;
  return 0;
}

// The rule's id: the caller's when given and not on the sheet yet, else one
// derived from the priority so add-then-list is deterministic. A colliding id
// is replaced rather than refused, so a read rule re-adds to its own sheet.
void assign_id(const fm_cf_rule_t& rule, const formulon::Sheet& sheet, std::size_t sheet_index,
               formulon::cf::CFRule* out) {
  if (rule.id != nullptr && rule.id[0] != '\0') {
    // Already checked against the `ST_Guid` shape by `validate`; a GUID's hex digits match in either case.
    bool collides = false;
    for (const auto& block : sheet.conditional_formats()) {
      for (const auto& existing : block.rules) {
        collides = collides || formulon::strings::case_insensitive_eq(existing.id, rule.id);
      }
    }
    if (!collides) {
      out->id = rule.id;
      return;
    }
  }
  std::unordered_set<std::string> used_ids;
  for (const auto& block : sheet.conditional_formats()) {
    for (const auto& existing : block.rules) {
      used_ids.insert(existing.id);
    }
  }
  out->id = formulon::cf::x14_rule_id(sheet_index, static_cast<std::uint32_t>(out->priority), used_ids);
}

// The colour-scale, data-bar or icon-set payload the rule's type carries.
fm_status_t copy_payloads(const fm_cf_rule_t& rule, formulon::cf::CFRule* out) {
  if (out->type == formulon::cf::RuleType::ColorScale && !copy_color_scale_payload(rule, out)) {
    return invalid_argument("invalid colorScale payload");
  }
  if (out->type == formulon::cf::RuleType::DataBar && !copy_data_bar_payload(rule, out)) {
    return invalid_argument("invalid dataBar payload");
  }
  if (out->type == formulon::cf::RuleType::IconSet && !copy_icon_set_payload(rule, out)) {
    return invalid_argument("invalid iconSet payload");
  }
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_add_cf_rule(fm_workbook_t* wb, std::size_t sheet_index, const fm_cf_rule_t* rule,
                                            std::size_t* out_index) {
  clear_last_error();
  if (rule == nullptr) {
    return null_argument("rule");
  }
  if (out_index == nullptr) {
    return null_argument("out_index");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*rule); rc != 0) {
    return rc;
  }
  if (auto rc = validate(*rule); rc != 0) {
    return rc;
  }
  if (auto rc = check_payload_colors(*rule); rc != 0) {
    return rc;
  }

  formulon::cf::ConditionalFormat new_block;
  new_block.sqref.reserve(rule->range_count);
  for (std::size_t i = 0; i < rule->range_count; ++i) {
    formulon::cf::CFCellRange r;
    r.first = formulon::CellAddress{rule->ranges[i].first_row, rule->ranges[i].first_col};
    r.last = formulon::CellAddress{rule->ranges[i].last_row, rule->ranges[i].last_col};
    new_block.sqref.push_back(r);
  }

  formulon::cf::CFRule out_rule;
  out_rule.type = static_cast<formulon::cf::RuleType>(rule->type);

  if (auto rc = assign_priority(*rule, wb->workbook().sheet(sheet_index), &out_rule); rc != 0) {
    return rc;
  }
  out_rule.stop_if_true = rule->stop_if_true != 0;
  if (rule->dxf_index_engaged != 0) {
    // A `dxfId` past the end of `<dxfs>` makes Excel treat the whole
    // `xl/styles.xml` reference as broken and strip every conditional format
    // from the workbook on open. The writer already refuses to emit an
    // unresolvable `dxfId`, but rejecting it here tells the caller at set
    // time which index was wrong instead of silently losing the format.
    const std::size_t dxf_count = wb->workbook().styles().dxfs.size();
    if (static_cast<std::size_t>(rule->dxf_index) >= dxf_count) {
      return invalid_argument("dxf_index out of range", "dxf_index=" + std::to_string(rule->dxf_index) +
                                                            " dxfs_count=" + std::to_string(dxf_count));
    }
    out_rule.dxf_id = rule->dxf_index;
  }
  assign_id(*rule, wb->workbook().sheet(sheet_index), sheet_index, &out_rule);
  if (rule->formula1 != nullptr) {
    if (auto rc = check_stored_formula(rule->formula1, "field=formula1", &out_rule.formula1.emplace()); rc != 0) {
      return rc;
    }
  }
  if (rule->formula2 != nullptr) {
    if (auto rc = check_stored_formula(rule->formula2, "field=formula2", &out_rule.formula2.emplace()); rc != 0) {
      return rc;
    }
  }
  if (rule->op_engaged != 0) {
    out_rule.op = cf_operator_from_c(rule->op);
  }
  if (rule->rank_engaged != 0) {
    out_rule.rank = rule->rank;
  }
  out_rule.percent = (rule->percent != 0);
  out_rule.bottom = (rule->bottom != 0);
  out_rule.above_average = (rule->above_average != 0);
  out_rule.equal_average = (rule->equal_average != 0);
  if (rule->std_dev_engaged != 0) {
    out_rule.std_dev = rule->std_dev;
  }
  if (rule->text != nullptr) {
    out_rule.text = std::string(rule->text);
  }
  if (rule->time_period_engaged != 0) {
    out_rule.time_period = date_period_from_c(rule->time_period);
  }
  if (!out_rule.formula1.has_value() || out_rule.formula1->empty()) {
    // Excel stores a derived formula for these types and readers surface it.
    if (auto derived = formulon::cf::derived_rule_formula(out_rule, formulon::cf::sqref_anchor(new_block.sqref))) {
      out_rule.formula1 = std::move(*derived);
    }
  }
  if (auto rc = copy_payloads(*rule, &out_rule); rc != 0) {
    return rc;
  }
  if (auto rc = store_cfvo_formulas(wb, &out_rule); rc != 0) {
    return rc;
  }
  // Every check passed, so binding the books the formulas name commits nothing early.
  if (rule->formula1 != nullptr && rule->formula1[0] != '\0') {
    out_rule.formula1 = ingest_entered_formula(wb, rule->formula1);
  }
  if (rule->formula2 != nullptr && rule->formula2[0] != '\0') {
    out_rule.formula2 = ingest_entered_formula(wb, rule->formula2);
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

extern "C" fm_status_t fm_sheet_remove_cf_rule(fm_workbook_t* wb, std::size_t sheet_index, std::size_t idx) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto& sheet = wb->workbook().sheet(sheet_index);
  auto& blocks = sheet.mutable_conditional_formats();
  std::size_t b = 0;
  std::size_t r = 0;
  if (auto rc = resolve_flat_index(blocks, idx, &b, &r); rc != 0) {
    return rc;
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

extern "C" fm_status_t fm_sheet_clear_cf_rules(fm_workbook_t* wb, std::size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  auto& sheet = wb->workbook().sheet(sheet_index);
  sheet.mutable_conditional_formats().clear();
  // With no model rules left, every id-bearing x14 overlay entry is a
  // dangling reference; reconciliation prunes them all.
  sheet.set_ext_lst_xml(formulon::io::reconcile_x14_cf_overlay(sheet.ext_lst_xml(), sheet.conditional_formats()));
  return 0;
}
