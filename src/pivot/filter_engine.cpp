
#include "pivot/filter_engine.h"

#include <algorithm>
#include <cassert>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_set>
#include <utility>
#include <variant>
#include <vector>

#include "eval/collated_wildcard.h"
#include "eval/collation.h"
#include "eval/eval_profile_scope.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/field_lookup.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_index.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/record_access.h"
#include "pivot/value_order.h"
#include "utils/checked_index.h"
#include "utils/date_time.h"
#include "utils/double_parse.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon::pivot {
namespace {

// The text payload of `slot`; empty for any other alternative.
std::string_view payload_text(const std::variant<std::monostate, int, double, std::string>& slot) {
  if (const auto* s = std::get_if<std::string>(&slot)) {
    return *s;
  }
  return {};
}

// The numeric alternative of `slot` (an int widens to double), as stored.
std::optional<double> raw_payload_number(const std::variant<std::monostate, int, double, std::string>& slot) {
  if (const auto* d = std::get_if<double>(&slot)) {
    return *d;
  }
  if (const auto* i = std::get_if<int>(&slot)) {
    return static_cast<double>(*i);
  }
  return std::nullopt;
}

// The value a label comparison reads for `cell` of field `field_index`: a group
// field's materialised number stands for its group item's text.
Value label_value(const PivotCache& cache, std::size_t field_index, const Value& cell) {
  if (!cell.is_number() || field_index >= cache.fields().size() || !is_group_field(cache.fields()[field_index])) {
    return cell;
  }
  const std::vector<Value>& items = cache.fields()[field_index].field_group->group_items;
  const std::optional<std::size_t> item = index_from_double(cell.as_number(), items.size());
  return item.has_value() ? items[*item] : cell;
}

// Whole-label wildcard match under the profile's `=` character equivalence.
// The matcher is built on first use, keeping the pattern parse and the
// collation classes across the labels of one filter.
bool caption_wildcard_matches(std::optional<eval::CollatedWildcard>& matcher, std::string_view pattern,
                              std::string_view label) {
  assert(eval::eval_profile_scope_active());
  if (!matcher) {
    matcher.emplace(pattern, eval::current_eval_profile());
  }
  return matcher->match(label);
}

// Equality and ordering of a caption filter use the strength that merged the
// items, so which spelling of a merged item survives cannot change the result.
int caption_collate(std::string_view lhs, std::string_view rhs) {
  assert(eval::eval_profile_scope_active());
  return eval::collate(lhs, rhs, eval::current_eval_profile(), eval::CollationStrength::kDynamic);
}

// An operand carrying a wildcard or an escape matches as a pattern.
bool is_caption_pattern(std::string_view operand) {
  return operand.find_first_of("*?~") != std::string_view::npos;
}

// The caption operand as a number, for comparison with a numeric label.
std::optional<double> caption_number(std::string_view operand) {
  double parsed = 0.0;
  if (!parse_excel_decimal(operand, &parsed) || !std::isfinite(parsed)) {
    return std::nullopt;
  }
  return parsed;
}

// Orders a label against one caption operand: numerically when both are
// numbers; a pattern operand counts a matching label as equal and any other
// label as greater; otherwise by collation.
int caption_order(std::string_view label, std::optional<double> label_number, std::string_view operand,
                  std::optional<double> operand_number, bool pattern, std::optional<eval::CollatedWildcard>& matcher) {
  if (label_number.has_value() && operand_number.has_value()) {
    return *label_number < *operand_number ? -1 : (*label_number > *operand_number ? 1 : 0);
  }
  if (pattern) {
    return caption_wildcard_matches(matcher, operand, label) ? 0 : 1;
  }
  return caption_collate(label, operand);
}

// The inclusive calendar-month range of a `DateQ*` / `DateM*` type.
std::pair<unsigned, unsigned> calendar_month_range(FilterType type) {
  if (type >= FilterType::DateM1) {
    const unsigned month = static_cast<unsigned>(type) - static_cast<unsigned>(FilterType::DateM1) + 1U;
    return {month, month};
  }
  const unsigned first = (static_cast<unsigned>(type) - static_cast<unsigned>(FilterType::DateQ1)) * 3U + 1U;
  return {first, first + 2U};
}

// Rewrites each leaf set in `leaf_sets` from the pre-filter leaf index space
// into the surviving-leaf one, dropping pruned leaves. Returns a per-set mask
// marking the sets that still cover at least one surviving leaf; a set that
// empties out belongs to a subtotal whose whole group was filtered away.
std::vector<bool> remap_leaf_sets(std::vector<std::vector<std::size_t>>& leaf_sets, const std::vector<bool>& keep) {
  std::vector<std::size_t> new_index(keep.size(), 0);
  std::size_t surviving = 0;
  for (std::size_t i = 0; i < keep.size(); ++i) {
    new_index[i] = surviving;
    if (keep[i]) {
      ++surviving;
    }
  }
  std::vector<bool> keep_set(leaf_sets.size(), false);
  for (std::size_t s = 0; s < leaf_sets.size(); ++s) {
    std::vector<std::size_t>& set = leaf_sets[s];
    std::vector<std::size_t> remapped;
    remapped.reserve(set.size());
    for (const std::size_t leaf : set) {
      if (leaf < keep.size() && keep[leaf]) {
        remapped.push_back(new_index[leaf]);
      }
    }
    keep_set[s] = !remapped.empty();
    set = std::move(remapped);
  }
  return keep_set;
}

// True when table field `field`, or a field sharing its group base, sits on
// the row, column or page axis.
bool placed_on_axis(const PivotTable& table, const PivotCache& cache, std::size_t field) {
  const std::size_t base = field < cache.fields().size() ? group_base_index(cache, field) : field;
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    const PivotAxis axis = table.fields()[fi].axis;
    if (axis != PivotAxis::Row && axis != PivotAxis::Col && axis != PivotAxis::Page) {
      continue;
    }
    if (fi == field || (fi < cache.fields().size() && group_base_index(cache, fi) == base)) {
      return true;
    }
  }
  return false;
}

}  // namespace

void drop_field_aggregates(PivotTable& table, std::size_t field_index) {
  std::vector<PivotDataField>& data_fields = table.mutable_data_fields();
  std::vector<std::size_t> new_index(data_fields.size());
  std::size_t kept = 0;
  for (std::size_t i = 0; i < data_fields.size(); ++i) {
    new_index[i] = data_fields[i].field_index == field_index ? data_fields.size() : kept++;
  }
  const PivotField& field = table.fields()[field_index];
  table.erase_active_filters_if([&](PivotFilter& filter) {
    if (pivot_field_has_name(field, filter.field_name)) {
      return true;
    }
    if (!is_value_filter_type(filter.type) || filter.data_field_index >= new_index.size()) {
      return false;
    }
    const std::size_t moved = new_index[filter.data_field_index];
    filter.data_field_index = static_cast<decltype(filter.data_field_index)>(moved);
    return moved == data_fields.size();
  });
  data_fields.erase(std::remove_if(data_fields.begin(), data_fields.end(),
                                   [field_index](const PivotDataField& df) { return df.field_index == field_index; }),
                    data_fields.end());
}

// Pulls the numeric payload of a filter when one is expected. The
// variant carries either `int` or `double`; both coerce to double here.
// Returns 0 for non-numeric variants (caller treats as no-op).
double filter_number_value(const PivotFilter& f) {
  return raw_payload_number(f.value).value_or(0.0);
}

// Pulls the numeric upper-bound payload of a range filter. Returns
// `nullopt` when `value_high` is `monostate` (i.e. caller didn't set
// it), signalling the range is unbounded above; the calling filter
// then no-ops.
std::optional<double> filter_number_value_high(const PivotFilter& f) {
  return raw_payload_number(f.value_high);
}

PreparedRecordFilter::PreparedRecordFilter(const PivotTable& table, const PivotCache& cache, const PivotFilterEnv& env)
    : cache_(&cache), date1904_(env.date1904) {
  // A field with no hidden item filters nothing, so it never reaches the
  // record loop at all. Excel writes an `items[]` list for every field it
  // places on an axis, which is why presence of the list cannot stand in for
  // presence of a filter.
  for (std::size_t fi = 0; fi < table.fields().size(); ++fi) {
    HiddenItems hidden;
    hidden.field_index = fi;
    // A page field with a single selected item hides every other item. With
    // subtotalHiddenItems its hidden items are not a filter at all.
    const PivotPageField* page = table.find_page_field(static_cast<std::uint32_t>(fi));
    const bool is_page = page != nullptr;
    const std::optional<std::uint32_t> selected = is_page ? page->item_index : std::nullopt;
    if (is_page && table.subtotal_hidden_items()) {
      continue;
    }
    const std::vector<PivotItem>& field_items = table.fields()[fi].items;
    const bool group_field = fi < cache.fields().size() && is_group_field(cache.fields()[fi]);
    for (std::size_t item_pos = 0; item_pos < field_items.size(); ++item_pos) {
      const PivotItem& item = field_items[item_pos];
      const bool selected_out = selected.has_value() && *selected < field_items.size() && *selected != item_pos;
      if (item.visible && !selected_out) {
        continue;
      }
      if (group_field) {
        if (item.has_cache_index && pivot_item_cache_value(cache, fi, item) != nullptr) {
          hidden.group_items.insert(item.cache_index);
        }
        continue;
      }
      if (item.name.empty()) {
        // The one axis item with no label of its own is the blank: it binds to
        // a cache value that renders to nothing. Identify it by that binding
        // rather than by the label it is drawn with — the placeholder is an
        // ordinary string a genuine text value is free to spell identically,
        // and matching on it would hide that value too.
        const Value* bound = pivot_item_cache_value(cache, fi, item);
        if (bound != nullptr && bound->is_blank()) {
          hidden.hides_blank = true;
          continue;
        }
      }
      // An item whose binding does not resolve has no label and is merely
      // malformed; it must not filter anything.
      std::string label = pivot_item_label(cache, fi, item);
      if (label.empty()) {
        continue;
      }
      hidden.labels.insert(std::string_view{hidden_labels_.emplace_back(std::move(label))});
    }
    if (hidden.labels.empty() && !hidden.hides_blank && hidden.group_items.empty()) {
      continue;
    }
    hidden_items_.push_back(std::move(hidden));
  }
  // A relative period resolves against one clock reading, taken once: an
  // evaluation spanning midnight would otherwise filter its early records
  // against one day and its later ones against the next.
  std::optional<date_time::CivilTime> now;
  for (const PivotFilter& f : table.active_filters()) {
    if (is_value_filter_type(f.type)) {
      continue;  // Ranked after aggregation by `apply_value_filters`.
    }
    const std::optional<std::size_t> fi = filter_field_index(table, f);
    if (!fi.has_value()) {
      continue;
    }
    // A caption or date filter acts only on a field placed on the row, column
    // or page axis, directly or through a group field over the same base; on
    // any other field it stays inert and is kept as written.
    if (!placed_on_axis(table, cache, *fi)) {
      continue;
    }
    ResolvedFilter resolved;
    resolved.field_index = *fi;
    resolved.type = f.type;
    if (is_caption_filter_type(f.type)) {
      if (!std::holds_alternative<std::string>(f.value)) {
        continue;
      }
      const std::string_view low = payload_text(f.value);
      if (std::optional<std::string> pattern = caption_wildcard_pattern(f.type, low)) {
        resolved.low = std::move(*pattern);
      } else {
        resolved.low.assign(low);
        resolved.low_pattern = is_caption_pattern(low);
        resolved.low_number = caption_number(low);
      }
      if (is_range_filter_type(f.type)) {
        if (!std::holds_alternative<std::string>(f.value_high)) {
          continue;  // Unbounded above: no-op.
        }
        const std::string_view high = payload_text(f.value_high);
        resolved.high.assign(high);
        resolved.high_pattern = is_caption_pattern(high);
        resolved.high_number = caption_number(high);
      }
      filters_.push_back(std::move(resolved));
      continue;
    }
    if (is_operandless_filter_type(f.type)) {
      if (f.type >= FilterType::DateQ1) {
        const auto [low, high] = calendar_month_range(f.type);
        resolved.date_low = static_cast<double>(low);
        resolved.date_high = static_cast<double>(high);
      } else {
        if (!now.has_value()) {
          now = env.pinned_now.has_value() ? *env.pinned_now : date_time::host_civil_time();
        }
        const DateWindow window = resolve_relative_period(f.type, *now, env.date1904);
        resolved.date_low = window.low;
        resolved.date_high = window.high;
      }
      filters_.push_back(std::move(resolved));
      continue;
    }
    // A date comparison or range: a missing, non-finite or out-of-domain bound
    // leaves the filter inert.
    const std::optional<double> low = payload_number(f.value);
    if (!low.has_value() || !is_valid_date_serial(*low, env.date1904)) {
      continue;
    }
    resolved.date_low = *low;
    if (is_range_filter_type(f.type)) {
      const std::optional<double> high = payload_number(f.value_high);
      if (!high.has_value() || !is_valid_date_serial(*high, env.date1904)) {
        continue;
      }
      resolved.date_high = *high;
    }
    filters_.push_back(std::move(resolved));
  }
}

bool PreparedRecordFilter::caption_passes(const ResolvedFilter& f, const Value& v) const {
  const std::optional<double> number = v.is_number() ? std::optional<double>(v.as_number()) : std::nullopt;
  std::string rendered;
  std::string_view label;
  if (v.is_text()) {
    label = v.as_text();
  } else {
    rendered = display_string(v);
    label = rendered;
  }
  // A numeric label orders before a plain text operand that does not read as a number.
  const auto unordered = [&](bool operand_numeric, bool pattern) {
    return number.has_value() && !operand_numeric && !pattern;
  };
  const bool low_unordered = unordered(f.low_number.has_value(), f.low_pattern);
  const bool high_unordered = unordered(f.high_number.has_value(), f.high_pattern);
  const auto order_low = [&]() {
    return low_unordered ? -1 : caption_order(label, number, f.low, f.low_number, f.low_pattern, f.matcher);
  };
  const auto order_high = [&]() {
    return high_unordered ? -1 : caption_order(label, number, f.high, f.high_number, f.high_pattern, f.matcher_high);
  };
  // "Between" with a bound the label cannot order against matches nothing.
  const auto in_between = [&]() { return !low_unordered && !high_unordered && order_low() >= 0 && order_high() <= 0; };
  const auto equal = [&]() {
    return f.low_pattern ? caption_wildcard_matches(f.matcher, f.low, label) : caption_collate(label, f.low) == 0;
  };
  switch (f.type) {
    case FilterType::LabelEqual:
      return equal();
    case FilterType::LabelNotEqual:
      return !equal();
    case FilterType::LabelBeginsWith:
    case FilterType::LabelEndsWith:
    case FilterType::LabelContains:
      return caption_wildcard_matches(f.matcher, f.low, label);
    case FilterType::LabelNotBeginsWith:
    case FilterType::LabelNotEndsWith:
    case FilterType::LabelNotContains:
      return !caption_wildcard_matches(f.matcher, f.low, label);
    case FilterType::LabelGreaterThan:
      return order_low() > 0;
    case FilterType::LabelGreaterThanOrEqual:
      return order_low() >= 0;
    case FilterType::LabelLessThan:
      return order_low() < 0;
    case FilterType::LabelLessThanOrEqual:
      return order_low() <= 0;
    case FilterType::LabelBetween:
      return in_between();
    case FilterType::LabelNotBetween:
      return order_low() < 0 || order_high() > 0;
    default:
      return true;
  }
}

bool PreparedRecordFilter::date_passes(const ResolvedFilter& f, const Value& v, bool is_date) const {
  // Cells that are no date (text, blanks, errors, plain numbers) skip a date filter
  // rather than being dropped.
  if (!is_date || !v.is_number() || !is_valid_date_serial(v.as_number(), date1904_)) {
    return true;
  }
  const double serial = v.as_number();
  switch (f.type) {
    case FilterType::DateEqual:
      return serial == f.date_low;
    case FilterType::DateNotEqual:
      return serial != f.date_low;
    case FilterType::DateOlderThan:
      return serial < f.date_low;
    case FilterType::DateOlderThanOrEqual:
      return serial <= f.date_low;
    case FilterType::DateNewerThan:
      return serial > f.date_low;
    case FilterType::DateNewerThanOrEqual:
      return serial >= f.date_low;
    case FilterType::DateNotBetween:
      return serial < f.date_low || serial > f.date_high;
    default:
      break;
  }
  if (f.type >= FilterType::DateQ1) {
    const double month = static_cast<double>(date_time::ymd_from_serial(std::floor(serial), date1904_).m);
    return month >= f.date_low && month <= f.date_high;
  }
  // `LabelDate` and the relative periods: an inclusive window.
  return serial >= f.date_low && serial <= f.date_high;
}

bool PreparedRecordFilter::passes(const PivotCacheRecord& record) const {
  const PivotCache& cache = *cache_;
  for (const HiddenItems& hidden : hidden_items_) {
    const Value v = cell_value(cache, record, hidden.field_index);
    if (!hidden.group_items.empty()) {
      const std::optional<std::size_t> number =
          v.is_number() ? index_from_double(v.as_number(), std::size_t{0xFFFFFFFFU}) : std::nullopt;
      if (number.has_value() && hidden.group_items.count(static_cast<std::uint32_t>(*number)) != 0) {
        return false;
      }
      continue;
    }
    if (v.is_blank()) {
      // A blank renders to the empty label, which no named item can spell:
      // the unlabelled item's rule is the only one that can hide it.
      if (hidden.hides_blank) {
        return false;
      }
      continue;
    }
    if (hidden.labels.empty()) {
      continue;  // Hides only the blank, and this record is not blank.
    }
    // A hidden item matches the label the grid draws for the record's value.
    // Text is its own label and is looked up without building a string; every
    // other kind renders once per field here rather than once per item.
    if (v.is_text()) {
      if (hidden.labels.count(v.as_text()) != 0) {
        return false;
      }
      continue;
    }
    const std::string rendered = display_string(v);
    if (hidden.labels.count(std::string_view{rendered}) != 0) {
      return false;
    }
  }
  for (const ResolvedFilter& f : filters_) {
    const Value v = cell_value(cache, record, f.field_index);
    const bool kept = is_caption_filter_type(f.type)
                          ? caption_passes(f, label_value(cache, f.field_index, v))
                          : date_passes(f, v, cell_is_date_value(cache, record, f.field_index));
    if (!kept) {
      return false;
    }
  }
  return true;
}

std::optional<double> payload_number(const std::variant<std::monostate, int, double, std::string>& slot) {
  const std::optional<double> number = raw_payload_number(slot);
  return number.has_value() && !std::isfinite(*number) ? std::nullopt : number;
}

std::optional<std::string> caption_wildcard_pattern(FilterType type, std::string_view value) {
  const bool leading = type == FilterType::LabelEndsWith || type == FilterType::LabelNotEndsWith ||
                       type == FilterType::LabelContains || type == FilterType::LabelNotContains;
  const bool trailing = type == FilterType::LabelBeginsWith || type == FilterType::LabelNotBeginsWith ||
                        type == FilterType::LabelContains || type == FilterType::LabelNotContains;
  if (!leading && !trailing) {
    return std::nullopt;
  }
  std::string pattern;
  pattern.reserve(value.size() + 2U);
  if (leading) {
    pattern.push_back('*');
  }
  pattern.append(value);
  if (trailing) {
    pattern.push_back('*');
  }
  return pattern;
}

std::optional<std::size_t> filter_field_index(const PivotTable& table, const PivotFilter& f) {
  if (f.field_name.empty()) {
    if (f.source.has_value() && f.source->field_index < table.fields().size()) {
      return static_cast<std::size_t>(f.source->field_index);
    }
    return std::nullopt;
  }
  return resolve_field_by_any_name(table, f.field_name);
}

DateWindow resolve_relative_period(FilterType type, const date_time::CivilTime& now, bool date1904) {
  const auto serial = [date1904](int year, unsigned month, unsigned day) {
    return date_time::serial_from_ymd(year, month, day, date1904);
  };
  // Normalised (year, month) arithmetic. `serial_from_ymd` normalises an
  // out-of-range *day*, but a month of 0 would fall outside the era math it
  // is built on, so the rollover is done here instead.
  const auto add_months = [](int year, unsigned month, int delta) {
    date_time::shift_year_month(year, month, delta);
    return std::pair<int, unsigned>{year, month};
  };
  // A whole-month window, given the month it starts in and its length.
  const auto span = [&](int year, unsigned month, int months) {
    const auto [end_year, end_month] = add_months(year, month, months);
    return DateWindow{serial(year, month, 1U), serial(end_year, end_month, 1U) - 1.0};
  };

  const int year = now.date.y;
  const unsigned month = now.date.m;
  const double today = serial(year, month, now.date.d);
  // The calendar quarter's first month: 1, 4, 7 or 10.
  const unsigned quarter_start = ((month - 1U) / 3U) * 3U + 1U;
  // A whole calendar week, `weeks` away from the one holding today. Excel
  // anchors these on the calendar week running Sunday through Saturday,
  // not on a rolling seven days, so `ThisWeek` read on a Friday still
  // starts the preceding Sunday and the three windows tile exactly.
  const auto week_span = [&](int weeks) {
    const double start = date_time::sunday_of_week(today, date1904) + static_cast<double>(weeks) * 7.0;
    return DateWindow{start, start + 6.0};
  };

  switch (type) {
    case FilterType::DateToday:
      return DateWindow{today, today};
    case FilterType::DateYesterday:
      return DateWindow{today - 1.0, today - 1.0};
    case FilterType::DateTomorrow:
      return DateWindow{today + 1.0, today + 1.0};
    case FilterType::DateThisMonth:
      return span(year, month, 1);
    case FilterType::DateLastMonth: {
      const auto [prev_year, prev_month] = add_months(year, month, -1);
      return span(prev_year, prev_month, 1);
    }
    case FilterType::DateNextMonth: {
      const auto [next_year, next_month] = add_months(year, month, 1);
      return span(next_year, next_month, 1);
    }
    case FilterType::DateThisQuarter:
      return span(year, quarter_start, 3);
    case FilterType::DateLastQuarter: {
      const auto [prev_year, prev_month] = add_months(year, quarter_start, -3);
      return span(prev_year, prev_month, 3);
    }
    case FilterType::DateNextQuarter: {
      const auto [next_year, next_month] = add_months(year, quarter_start, 3);
      return span(next_year, next_month, 3);
    }
    case FilterType::DateThisYear:
      return span(year, 1U, 12);
    case FilterType::DateLastYear:
      return span(year - 1, 1U, 12);
    case FilterType::DateNextYear:
      return span(year + 1, 1U, 12);
    case FilterType::DateYearToDate:
      // Excel's "year to date" runs from January 1 through today inclusive,
      // not through the end of the year.
      return DateWindow{serial(year, 1U, 1U), today};
    case FilterType::DateThisWeek:
      return week_span(0);
    case FilterType::DateLastWeek:
      return week_span(-1);
    case FilterType::DateNextWeek:
      return week_span(1);
    default:
      break;
  }
  return DateWindow{today, today};
}

std::optional<std::vector<bool>> build_value_filter_keep(const PivotFilter& f, const AxisScores& axis) {
  const std::size_t n = axis.scores.size();
  std::vector<bool> keep(n, false);
  const bool top =
      f.type == FilterType::ValueTop10 || f.type == FilterType::ValueTopPercent || f.type == FilterType::ValueTopSum;
  const auto by_score = [&](std::uint32_t lhs, std::uint32_t rhs) {
    return top ? axis.scores[lhs] > axis.scores[rhs] : axis.scores[lhs] < axis.scores[rhs];
  };
  if (f.type == FilterType::ValueTop10 || f.type == FilterType::ValueBottomCount) {
    // The filter value is embedder-supplied, so saturate it instead of
    // narrowing it blind: a request larger than the axis keeps every leaf,
    // while a negative one — and NaN, which loses every comparison — keeps
    // none.
    const double requested = filter_number_value(f);
    const std::size_t top_n =
        index_from_double(requested, n + 1U).value_or(requested > static_cast<double>(n) ? n : 0U);
    // All-blank leaves sink to the bottom of the ranking regardless of N or
    // direction, so a Bottom-N filter never picks a blank as one of its
    // "lowest" items.
    const auto ranked = [&](std::uint32_t lhs, std::uint32_t rhs) {
      if (axis.all_blank[lhs] != axis.all_blank[rhs]) {
        return !axis.all_blank[lhs];
      }
      return by_score(lhs, rhs);
    };
    std::vector<std::uint32_t> order;
    sorted_index_order(order, static_cast<std::uint32_t>(n), make_index_less(ranked));
    const std::size_t k = std::min(top_n, n);
    for (std::size_t i = 0; i < k; ++i) {
      if (!axis.all_blank[order[i]]) {
        keep[order[i]] = true;
      }
    }
    return keep;
  }
  if (f.type == FilterType::ValueTopPercent || f.type == FilterType::ValueBottomPercent ||
      f.type == FilterType::ValueTopSum || f.type == FilterType::ValueBottomSum) {
    // All-blank leaves never contribute and never survive, matching the
    // count types.
    std::vector<std::uint32_t> order;
    order.reserve(n);
    double total = 0.0;
    for (std::size_t i = 0; i < n; ++i) {
      if (!axis.all_blank[i]) {
        order.push_back(static_cast<std::uint32_t>(i));
        total += axis.scores[i];
      }
    }
    sort_index_order(order, make_index_less(by_score));
    const double target = filter_number_value(f);
    const bool percent = f.type == FilterType::ValueTopPercent || f.type == FilterType::ValueBottomPercent;
    const double threshold = percent ? total * target / 100.0 : target;
    double running = 0.0;
    for (const std::uint32_t idx : order) {
      if (running >= threshold) {
        break;
      }
      keep[idx] = true;
      running += axis.scores[idx];
    }
    return keep;
  }
  const double lo = filter_number_value(f);
  const auto keep_where = [&](auto&& test) {
    for (std::size_t i = 0; i < n; ++i) {
      keep[i] = !axis.all_blank[i] && test(axis.scores[i]);
    }
    return keep;
  };
  switch (f.type) {
    case FilterType::ValueGreaterThan:
      return keep_where([&](double s) { return s > lo; });
    case FilterType::ValueGreaterThanOrEqual:
      return keep_where([&](double s) { return s >= lo; });
    case FilterType::ValueLessThan:
      return keep_where([&](double s) { return s < lo; });
    case FilterType::ValueLessThanOrEqual:
      return keep_where([&](double s) { return s <= lo; });
    case FilterType::ValueEqual:
      return keep_where([&](double s) { return s == lo; });
    case FilterType::ValueNotEqual:
      return keep_where([&](double s) { return s != lo; });
    case FilterType::ValueBetween:
    case FilterType::ValueNotBetween: {
      const auto hi_or = filter_number_value_high(f);
      if (!hi_or) {
        return std::nullopt;  // Unbounded above -> no-op for this filter.
      }
      const bool inside = f.type == FilterType::ValueBetween;
      return keep_where([&](double s) { return (s >= lo && s <= *hi_or) == inside; });
    }
    default:
      return std::nullopt;
  }
}

std::optional<std::vector<bool>> build_grouped_value_filter_keep(const PivotFilter& f,
                                                                 const std::vector<AxisLeafGroup>& groups,
                                                                 const AxisScores& group_scores,
                                                                 std::size_t leaf_count) {
  std::vector<bool> keep(leaf_count, false);
  std::size_t run_start = 0;
  while (run_start < groups.size()) {
    std::size_t run_end = run_start + 1U;
    while (run_end < groups.size() && groups[run_end].parent == groups[run_start].parent) {
      ++run_end;
    }
    AxisScores run;
    run.scores.assign(group_scores.scores.begin() + static_cast<std::ptrdiff_t>(run_start),
                      group_scores.scores.begin() + static_cast<std::ptrdiff_t>(run_end));
    run.all_blank.assign(group_scores.all_blank.begin() + static_cast<std::ptrdiff_t>(run_start),
                         group_scores.all_blank.begin() + static_cast<std::ptrdiff_t>(run_end));
    const auto run_keep = build_value_filter_keep(f, run);
    if (!run_keep) {
      return std::nullopt;
    }
    for (std::size_t g = run_start; g < run_end; ++g) {
      if (!(*run_keep)[g - run_start]) {
        continue;
      }
      const AxisLeafGroup& group = groups[g];
      for (std::size_t leaf = group.first_leaf; leaf < group.first_leaf + group.leaf_count && leaf < leaf_count;
           ++leaf) {
        keep[leaf] = true;
      }
    }
    run_start = run_end;
  }
  return keep;
}

void compact_leaf_axis(PivotResult& result, const std::vector<bool>& keep, LeafAxis axis,
                       std::vector<std::vector<std::size_t>>& row_subtotal_leaf_sets,
                       std::vector<std::vector<std::size_t>>& col_subtotal_leaf_sets) {
  if (axis == LeafAxis::Row) {
    keep_marked(result.values, keep);
    keep_marked(result.row_leaf_totals, keep);
    // A column subtotal renders one value per row leaf, so its matrix is
    // indexed by the axis being pruned.
    for (ColSubtotal& col_subtotal : result.col_subtotals) {
      keep_marked(col_subtotal.values, keep);
    }
    const std::vector<bool> keep_subtotal = remap_leaf_sets(row_subtotal_leaf_sets, keep);
    keep_marked(result.row_subtotals, keep_subtotal);
    keep_marked(result.subtotals, keep_subtotal);
    keep_marked(row_subtotal_leaf_sets, keep_subtotal);
    return;
  }
  for (auto& row_slot : result.values) {
    keep_marked(row_slot, keep);
  }
  keep_marked(result.col_leaf_totals, keep);
  // A row subtotal renders one value per column leaf.
  for (RowSubtotal& row_subtotal : result.row_subtotals) {
    keep_marked(row_subtotal.col_values, keep);
  }
  const std::vector<bool> keep_subtotal = remap_leaf_sets(col_subtotal_leaf_sets, keep);
  keep_marked(result.col_subtotals, keep_subtotal);
  keep_marked(col_subtotal_leaf_sets, keep_subtotal);
  // The row x column subtotal intersections are indexed by column subtotal.
  for (RowSubtotal& row_subtotal : result.row_subtotals) {
    keep_marked(row_subtotal.col_subtotal_values, keep_subtotal);
  }
}

}  // namespace formulon::pivot
