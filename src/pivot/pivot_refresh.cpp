// @size-budget: 9 KB
//
// Pivot cache rebuild from source values. See the header for the rules.

#include "pivot/pivot_refresh.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <initializer_list>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "eval/text_format/display_text.h"
#include "excel_profile.h"
#include "num_fmt_locale.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/field_grouping_internal.h"
#include "pivot/field_remap.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_locale.h"
#include "pivot/record_access.h"
#include "pivot/shared_items_hints.h"
#include "pivot/value_order.h"
#include "sheet_name.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::pivot {
namespace {

using IndexMap = std::vector<std::optional<std::uint32_t>>;

bool flag_at(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

bool any_set(const std::vector<bool>& flags) {
  return std::find(flags.begin(), flags.end(), true) != flags.end();
}

std::optional<std::uint32_t> mapped(const IndexMap& map, std::size_t index) {
  return index < map.size() ? map[index] : std::nullopt;
}

// Hash key of a field name under the header fold.
std::string header_key(std::string_view name) {
  if (std::optional<std::string> key = sheet_names::wide_key(name)) {
    return "w" + *key;
  }
  return "b" + std::string(name);
}

// Hash key of an item under the item fold: case-only for text, and no two
// kinds (a date and a number included) ever equal.
std::string item_key(const Value& v, bool is_date) {
  switch (v.kind()) {
    case ValueKind::Text:
      return "s" + collection_names::fold_key(v.as_text());
    case ValueKind::Number: {
      const double d = v.as_number() == 0.0 ? 0.0 : v.as_number();
      char bytes[sizeof d];
      std::memcpy(bytes, &d, sizeof d);
      return (is_date ? "d" : "n") + std::string(bytes, sizeof bytes);
    }
    case ValueKind::Bool:
      return v.as_boolean() ? "b1" : "b0";
    case ValueKind::Error:
      return "e" + std::to_string(static_cast<int>(v.as_error()));
    default:
      return "m";
  }
}

// `v` with any text payload copied into `cache`.
Value owned(PivotCache& cache, const Value& v) {
  return v.is_text() ? intern_text(cache.mutable_text_storage(), v.as_text()) : v;
}

void append_item(PivotCacheField& field, Value value, bool is_date, bool missing, bool calculated) {
  field.shared_items.push_back(value);
  field.shared_item_is_date.push_back(is_date);
  field.shared_item_is_missing.push_back(missing);
  field.shared_item_is_calculated.push_back(calculated);
}

bool is_date_range(const FieldGroup& group) {
  return group.kind == GroupKind::Range && (group.range.is_date || group.range.by != GroupBy::Range);
}

std::string short_date_text(double serial, const RefreshContext& context) {
  return text_format::format_value_for_display(Value::number(std::floor(serial)),
                                               localized_builtin_num_fmt(14U, context.profile), context.date1904,
                                               context.profile)
      .text;
}

// Moves a date range group's bounds and boundary labels to the base's current dates.
Expected<void, Error> refresh_date_bounds(PivotCache& cache, FieldGroup& group, const std::vector<Value>& base,
                                          const std::vector<bool>& base_is_date, const RefreshContext& context) {
  RangeGroup& range = group.range;
  if (!range.auto_start && !range.auto_end) {
    return Expected<void, Error>::Ok();
  }
  std::optional<double> lo;
  std::optional<double> hi;
  for (std::size_t r = 0; r < base.size(); ++r) {
    if (!base[r].is_number() || !flag_at(base_is_date, r) ||
        !is_valid_date_serial(base[r].as_number(), context.date1904)) {
      continue;
    }
    const double v = base[r].as_number();
    lo = lo.has_value() ? std::min(*lo, v) : v;
    hi = hi.has_value() ? std::max(*hi, v) : v;
  }
  if (!lo.has_value()) {
    return Expected<void, Error>::Ok();
  }
  const GroupBounds data = auto_date_group_bounds(*lo, *hi, context.date1904);
  const double start =
      range.auto_start ? data.start : group_date_serial(range.start_date, context.date1904).value_or(data.start);
  const double end = range.auto_end ? data.end : group_date_serial(range.end_date, context.date1904).value_or(data.end);
  const std::string start_iso = group_date_iso(start, context.date1904);
  const std::string end_iso = group_date_iso(end, context.date1904);
  if (start_iso == range.start_date && end_iso == range.end_date) {
    return Expected<void, Error>::Ok();
  }
  if (range.by == GroupBy::Years || (range.by == GroupBy::Days && range.interval > 1.0)) {
    return make_error(FormulonErrorCode::kNotImplemented,
                      "Refreshing a years or multi-day date group whose bounds move is not implemented");
  }
  range.start_date = start_iso;
  range.end_date = end_iso;
  if (group.group_items.size() >= 2U) {
    group.group_items.front() =
        intern_text(cache.mutable_text_storage(), date_group_before_label(short_date_text(start, context)));
    group.group_items.back() =
        intern_text(cache.mutable_text_storage(), date_group_after_label(short_date_text(end, context)));
  }
  return Expected<void, Error>::Ok();
}

// Regenerates a numeric range group's automatic bounds and items from the
// base's current numbers; with both bounds fixed the items stay verbatim.
// `items` maps each old group item to the new one of the same label while the
// two lists agree from the start.
Expected<void, Error> refresh_numeric_range(PivotCache& cache, FieldGroup& group, const std::vector<Value>& base,
                                            IndexMap& items) {
  const RangeGroup& range = group.range;
  NumericGroupState state;
  state.start = range.start_num;
  state.end = range.end_num;
  state.interval = range.interval;
  state.auto_start = range.auto_start;
  state.auto_end = range.auto_end;
  std::optional<double> lo;
  std::optional<double> hi;
  for (const Value& v : base) {
    if (v.is_number()) {
      lo = lo.has_value() ? std::min(*lo, v.as_number()) : v.as_number();
      hi = hi.has_value() ? std::max(*hi, v.as_number()) : v.as_number();
    }
  }
  const std::size_t old_count = group.group_items.size();
  items.clear();
  for (std::size_t k = 0; k < old_count; ++k) {
    items.push_back(static_cast<std::uint32_t>(k));
  }
  if ((!state.auto_start && !state.auto_end) || !lo.has_value()) {
    return detail::check_numeric_labels(state);
  }
  state.start = state.auto_start ? *lo : state.start;
  state.end = state.auto_end ? *hi : state.end;
  if (state.start > state.end) {
    return make_error(FormulonErrorCode::kNotImplemented,
                      "Refreshing a numeric group whose fixed bound passes the data's other end is not implemented");
  }
  const std::vector<Value> old_items = group.group_items;
  RETURN_IF_ERROR(detail::set_numeric_range(cache, group, state));
  std::size_t same = 0;
  while (same < old_count && same < group.group_items.size() && old_items[same] == group.group_items[same]) {
    ++same;
  }
  for (std::size_t k = same; k < old_count; ++k) {
    items[k] = std::nullopt;
  }
  return Expected<void, Error>::Ok();
}

// Gives the base items a discrete group has not seen a group item of their own.
void refresh_discrete_map(PivotCache& cache, FieldGroup& group, const PivotCacheField& base,
                          const IndexMap& base_items) {
  if (base_items.empty()) {
    return;
  }
  std::vector<std::optional<std::uint32_t>> map(base.shared_items.size());
  for (std::size_t k = 0; k < group.discrete_map.size(); ++k) {
    if (const std::optional<std::uint32_t> item = mapped(base_items, k); item.has_value() && *item < map.size()) {
      map[*item] = group.discrete_map[k];
    }
  }
  group.discrete_map.clear();
  for (std::size_t n = 0; n < map.size(); ++n) {
    if (!map[n].has_value()) {
      map[n] = static_cast<std::uint32_t>(group.group_items.size());
      group.group_items.push_back(intern_text(cache.mutable_text_storage(), display_string(base.shared_items[n])));
    }
    group.discrete_map.push_back(*map[n]);
  }
}

// Fixes every cell encoding `record` leaves to inference, so a field turning
// shared does not reread another field's numbers as indices.
void pin_cell_encodings(const PivotCache& cache, PivotCacheRecord& record) {
  if (record.cell_is_index.size() >= record.cells.size()) {
    return;
  }
  std::vector<bool> pinned(record.cells.size(), false);
  for (std::size_t fi = 0; fi < pinned.size() && fi < cache.fields().size(); ++fi) {
    pinned[fi] = cell_is_shared_index(record, cache.fields()[fi], fi);
  }
  record.cell_is_index = std::move(pinned);
}

// Points cell `field` of `record` at shared item `index`.
void store_item_index(const PivotCache& cache, PivotCacheRecord& record, std::size_t field, std::uint32_t index) {
  if (record.cells.size() <= field) {
    record.cells.resize(field + 1U, Value::blank());
  }
  pin_cell_encodings(cache, record);
  record.cells[field] = Value::number(static_cast<double>(index));
  record.cell_is_index[field] = true;
  if (field < record.cell_is_date.size()) {
    record.cell_is_date[field] = false;
  }
  // A shared item may view this cell's text, so a later write must not reuse its storage.
  if (field < record.cell_text_slot.size()) {
    record.cell_text_slot[field] = kNoCacheTextSlot;
  }
}

}  // namespace

bool same_header(std::string_view a, std::string_view b) {
  return header_key(a) == header_key(b);
}

std::vector<std::string> refreshed_field_names(const std::vector<std::string>& headers, std::string_view blank_word) {
  std::vector<std::string> names(headers.size());
  std::unordered_set<std::string> taken;
  const auto claim = [&](std::size_t column, const std::string& base) {
    std::string candidate = base;
    for (std::size_t k = 2; taken.count(header_key(candidate)) != 0U; ++k) {
      candidate = base + std::to_string(k);
    }
    taken.insert(header_key(candidate));
    names[column] = std::move(candidate);
  };
  for (std::size_t c = 0; c < headers.size(); ++c) {
    if (!headers[c].empty()) {
      claim(c, headers[c]);
    }
  }
  for (std::size_t c = 0; c < headers.size(); ++c) {
    if (!headers[c].empty()) {
      continue;
    }
    claim(c, std::string(blank_word) + std::to_string(c + 1U));
  }
  return names;
}

Expected<RefreshPlan, Error> build_refreshed_cache(const PivotCache& old, const RefreshSource& source,
                                                   const RefreshContext& context) {
  const std::vector<PivotCacheField>& old_fields = old.fields();
  const std::size_t old_count = old_fields.size();
  std::vector<std::string> names = refreshed_field_names(source.headers, context.blank_header_word);
  const std::uint32_t columns = source.columns;

  // Old database fields meet the new columns by name.
  const auto is_source_field = [&](std::size_t i) {
    return old_fields[i].is_database_field && !old_fields[i].formula.has_value();
  };
  std::unordered_map<std::string, std::size_t> old_by_name;
  for (std::size_t i = 0; i < old_count; ++i) {
    if (is_source_field(i)) {
      old_by_name.emplace(header_key(old_fields[i].name), i);
    }
  }
  IndexMap field_map(old_count);
  std::vector<std::optional<std::size_t>> column_source(columns);
  for (std::uint32_t c = 0; c < columns; ++c) {
    const auto it = old_by_name.find(header_key(names[c]));
    if (it != old_by_name.end() && !field_map[it->second].has_value()) {
      field_map[it->second] = c;
      column_source[c] = it->second;
    }
  }
  // Calculated fields stay; a group field stays while its base does.
  std::vector<bool> keep(old_count, false);
  for (std::size_t i = 0; i < old_count; ++i) {
    keep[i] = is_source_field(i) ? field_map[i].has_value() : old_fields[i].formula.has_value();
  }
  for (bool grew = true; grew;) {
    grew = false;
    for (std::size_t i = 0; i < old_count; ++i) {
      const std::optional<FieldGroup>& group = old_fields[i].field_group;
      if (!keep[i] && !is_source_field(i) && group.has_value() && group->base.has_value() && *group->base < old_count &&
          *group->base != i && keep[*group->base]) {
        keep[i] = true;
        grew = true;
      }
    }
  }
  std::vector<std::size_t> derived_source;
  std::uint32_t next = columns;
  for (std::size_t i = 0; i < old_count; ++i) {
    if (!is_source_field(i) && keep[i]) {
      field_map[i] = next++;
      derived_source.push_back(i);
    }
  }

  // Shared items stay on axis fields, discrete and numeric group bases and fields with calculated items.
  std::vector<bool> shared(old_count, false);
  for (std::size_t i = 0; i < old_count; ++i) {
    shared[i] = shared[i] || flag_at(context.axis_fields, i) || any_set(old_fields[i].shared_item_is_calculated);
    const std::optional<FieldGroup>& group = old_fields[i].field_group;
    const bool keeps_base_items = group.has_value() && (group->kind == GroupKind::Discrete ||
                                                        (group->kind == GroupKind::Range && !group->range.is_date));
    if (keeps_base_items && group->base.has_value() && *group->base < old_count) {
      shared[*group->base] = true;
    }
  }
  const bool purge_missing = old.missing_items_limit().has_value() && *old.missing_items_limit() == 0U;

  RefreshPlan plan;
  PivotCache& cache = plan.cache;
  cache.copy_definition_from(old);
  std::vector<PivotCacheField>& fields = cache.mutable_fields();
  fields.resize(next);
  plan.remap.fields = field_map;
  plan.remap.items.resize(old_count);
  std::vector<PivotCacheRecord>& records = cache.mutable_records();
  records.resize(source.rows);
  for (PivotCacheRecord& record : records) {
    record.cells.assign(columns, Value::blank());
    record.cell_is_index.assign(columns, false);
    record.cell_is_date.assign(columns, false);
  }

  std::vector<std::vector<Value>> column_values(columns);
  std::vector<std::vector<bool>> column_dates(columns);
  for (std::uint32_t c = 0; c < columns; ++c) {
    std::vector<Value>& values = column_values[c];
    std::vector<bool>& dates = column_dates[c];
    values.reserve(source.rows);
    for (std::size_t r = 0; r < source.rows; ++r) {
      const std::size_t at = r * columns + c;
      const Value v = at < source.cells.size() ? source.cells[at] : Value::blank();
      values.push_back(v);
      dates.push_back(v.is_number() && flag_at(source.cell_is_date, at));
    }

    PivotCacheField& field = fields[c];
    field.name = names[c];
    field.num_fmt_id = c < source.num_fmt_ids.size() ? source.num_fmt_ids[c] : 0U;
    field.shared_items_hints = derive_shared_items_hints(values, dates, context.date1904);
    apply_mixed_date_slots(values, dates, context.date1904, field.shared_items_hints);
    const std::optional<std::size_t> from = column_source[c];
    if (from.has_value()) {
      const PivotCacheField& previous = old_fields[*from];
      field.extra_attrs = previous.extra_attrs;
      field.trailing_children_xml = previous.trailing_children_xml;
      field.field_group = previous.field_group;
    }
    if (!from.has_value() || !shared[*from]) {
      // Inline text still folds to the first spelling of its item.
      std::unordered_map<std::string, Value> first_text;
      for (std::size_t r = 0; r < source.rows; ++r) {
        if (values[r].is_text()) {
          const auto first = first_text.emplace(item_key(values[r], false), Value::blank());
          if (first.second) {
            first.first->second = owned(cache, values[r]);
          }
          records[r].cells[c] = first.first->second;
        } else {
          records[r].cells[c] = owned(cache, values[r]);
        }
        records[r].cell_is_date[c] = dates[r];
      }
      if (from.has_value() && !old_fields[*from].shared_items.empty()) {
        plan.remap.items[*from].assign(old_fields[*from].shared_items.size(), std::nullopt);
      }
      continue;
    }

    std::unordered_map<std::string, std::uint32_t> index_of;
    for (std::size_t r = 0; r < source.rows; ++r) {
      const auto inserted =
          index_of.emplace(item_key(values[r], dates[r]), static_cast<std::uint32_t>(field.shared_items.size()));
      if (inserted.second) {
        append_item(field, owned(cache, values[r]), dates[r], false, false);
      }
      records[r].cells[c] = Value::number(static_cast<double>(inserted.first->second));
      records[r].cell_is_index[c] = true;
    }
    field.items_follow_values = true;
    const PivotCacheField& previous = old_fields[*from];
    const bool keeps_missing = !purge_missing || any_set(previous.shared_item_is_calculated);
    IndexMap& items = plan.remap.items[*from];
    items.assign(previous.shared_items.size(), std::nullopt);
    for (const bool calculated : {false, true}) {
      for (std::size_t k = 0; k < previous.shared_items.size(); ++k) {
        if (flag_at(previous.shared_item_is_calculated, k) != calculated) {
          continue;
        }
        const Value& value = previous.shared_items[k];
        const bool is_date = flag_at(previous.shared_item_is_date, k);
        const auto present = calculated ? index_of.end() : index_of.find(item_key(value, is_date));
        if (present != index_of.end()) {
          items[k] = present->second;
        } else if (calculated || keeps_missing) {
          items[k] = static_cast<std::uint32_t>(field.shared_items.size());
          append_item(field, owned(cache, value), is_date, !calculated, calculated);
        }
      }
    }
  }

  for (const std::size_t i : derived_source) {
    const PivotCacheField& previous = old_fields[i];
    PivotCacheField& field = fields[*field_map[i]];
    field.name = previous.name;
    field.num_fmt_id = previous.num_fmt_id;
    field.formula = previous.formula;
    field.is_database_field = false;
    field.shared_items_hints = previous.shared_items_hints;
    field.extra_attrs = previous.extra_attrs;
    field.trailing_children_xml = previous.trailing_children_xml;
    field.field_group = previous.field_group;
    for (std::size_t k = 0; k < previous.shared_items.size(); ++k) {
      append_item(field, owned(cache, previous.shared_items[k]), flag_at(previous.shared_item_is_date, k),
                  flag_at(previous.shared_item_is_missing, k), flag_at(previous.shared_item_is_calculated, k));
    }
  }

  // Groups: links follow the field map, text moves into the new cache, and
  // bounds and maps follow the base's current values.
  std::vector<std::size_t> old_of(next, old_count);
  for (std::size_t i = 0; i < old_count; ++i) {
    if (field_map[i].has_value()) {
      old_of[*field_map[i]] = i;
    }
  }
  for (std::size_t j = 0; j < next; ++j) {
    if (!fields[j].field_group.has_value()) {
      continue;
    }
    FieldGroup& group = *fields[j].field_group;
    for (std::optional<std::uint32_t>* link : {&group.par, &group.base}) {
      if (link->has_value()) {
        *link = mapped(field_map, **link);
      }
    }
    for (Value& item : group.group_items) {
      item = owned(cache, item);
    }
    if (!group.base.has_value() || *group.base >= columns || old_of[j] == old_count) {
      continue;
    }
    const std::uint32_t base = *group.base;
    if (is_date_range(group)) {
      RETURN_IF_ERROR(refresh_date_bounds(cache, group, column_values[base], column_dates[base], context));
    } else if (group.kind == GroupKind::Range) {
      RETURN_IF_ERROR(refresh_numeric_range(cache, group, column_values[base], plan.remap.items[old_of[j]]));
    } else if (group.kind == GroupKind::Discrete && old_of[base] < old_count) {
      const std::size_t old_group_items = group.group_items.size();
      refresh_discrete_map(cache, group, fields[base], plan.remap.items[old_of[base]]);
      IndexMap& group_items = plan.remap.items[old_of[j]];
      group_items.clear();
      for (std::size_t k = 0; k < old_group_items; ++k) {
        group_items.push_back(static_cast<std::uint32_t>(k));
      }
    }
  }

  for (const PivotCalculatedItem& entry : old.calculated_items()) {
    const std::optional<std::uint32_t> field = mapped(field_map, entry.cache_field_index);
    if (!field.has_value()) {
      continue;
    }
    PivotCalculatedItem moved = entry;
    moved.cache_field_index = *field;
    std::vector<PivotAreaReference> references;
    for (PivotAreaReference reference : moved.area.references) {
      if (reference.field.has_value()) {
        const std::size_t old_field = *reference.field;
        reference.field = mapped(field_map, old_field);
        if (!reference.field.has_value()) {
          continue;
        }
        if (old_field < plan.remap.items.size() && !plan.remap.items[old_field].empty()) {
          std::vector<std::uint32_t> items;
          for (const std::uint32_t x : reference.items) {
            if (const std::optional<std::uint32_t> item = mapped(plan.remap.items[old_field], x)) {
              items.push_back(*item);
            }
          }
          reference.items = std::move(items);
        }
      }
      references.push_back(std::move(reference));
    }
    moved.area.references = std::move(references);
    cache.mutable_calculated_items().push_back(std::move(moved));
  }

  plan.remap.names.reserve(next);
  for (const PivotCacheField& field : fields) {
    plan.remap.names.push_back(field.name);
  }
  return plan;
}

bool share_field_items(PivotCache& cache, std::size_t field) {
  if (field >= cache.fields().size() || cache.records().empty()) {
    return false;
  }
  const PivotCacheField& current = cache.fields()[field];
  if (!current.is_database_field || current.formula.has_value() || !current.shared_items.empty()) {
    return false;
  }
  // The values stay as the source gave them, so a refreshed cache stays refreshed.
  const bool from_source = cache.records_from_source();
  std::vector<PivotCacheRecord>& records = cache.mutable_records();
  PivotCacheField& target = cache.mutable_fields()[field];
  // Every encoding is pinned while the field is still inline, before any item exists.
  for (PivotCacheRecord& record : records) {
    pin_cell_encodings(cache, record);
  }
  std::unordered_map<std::string, std::uint32_t> index_of;
  for (PivotCacheRecord& record : records) {
    const bool indexed = field < record.cell_is_index.size() && record.cell_is_index[field];
    const Value value = indexed || field >= record.cells.size() ? Value::blank() : record.cells[field];
    const bool is_date = value.is_number() && flag_at(record.cell_is_date, field);
    const auto inserted =
        index_of.emplace(item_key(value, is_date), static_cast<std::uint32_t>(target.shared_items.size()));
    if (inserted.second) {
      append_item(target, value, is_date, false, false);
    }
    store_item_index(cache, record, field, inserted.first->second);
  }
  target.items_follow_values = true;
  if (from_source) {
    cache.mark_records_from_source();
  }
  return true;
}

void share_axis_fields(PivotCache& cache, const PivotTable& table) {
  for (std::size_t f = 0; f < table.fields().size(); ++f) {
    const PivotAxis axis = table.fields()[f].axis;
    if (axis == PivotAxis::Row || axis == PivotAxis::Col || axis == PivotAxis::Page) {
      share_field_items(cache, f);
    }
  }
}

std::optional<std::uint32_t> find_shared_item(const PivotCacheField& field, const Value& value, bool is_date) {
  const std::string key = item_key(value, value.is_number() && is_date);
  for (std::size_t k = 0; k < field.shared_items.size(); ++k) {
    const Value& item = field.shared_items[k];
    if (item_key(item, item.is_number() && flag_at(field.shared_item_is_date, k)) == key) {
      return static_cast<std::uint32_t>(k);
    }
  }
  return std::nullopt;
}

void write_shared_record_value(PivotCache& cache, std::size_t record, std::size_t field, const Value& value,
                               bool is_date) {
  const bool date = value.is_number() && is_date;
  std::optional<std::uint32_t> index = find_shared_item(cache.fields()[field], value, date);
  if (!index.has_value()) {
    const Value stored = owned(cache, value);
    PivotCacheField& target = cache.mutable_fields()[field];
    const std::size_t count = target.shared_items.size();
    // Items listed by hand may carry shorter flag vectors.
    target.shared_item_is_date.resize(count, false);
    target.shared_item_is_missing.resize(count, false);
    target.shared_item_is_calculated.resize(count, false);
    index = static_cast<std::uint32_t>(count);
    append_item(target, stored, date, false, false);
  }
  store_item_index(cache, cache.mutable_records()[record], field, *index);
}

void keep_field_items(PivotField& field, const PivotCacheField& cache_field) {
  const std::vector<Value>& shared = cache_field.shared_items;
  const auto listed = [&cache_field](std::size_t k) {
    return !flag_at(cache_field.shared_item_is_missing, k) && !flag_at(cache_field.shared_item_is_calculated, k);
  };
  const auto add = [&field](std::size_t k) {
    PivotItem item;
    item.has_cache_index = true;
    item.cache_index = static_cast<std::uint32_t>(k);
    field.items.push_back(std::move(item));
  };
  if (shared.empty()) {
    return;
  }
  if (field.items.empty()) {
    if (!field.marker_items.empty()) {
      return;
    }
    std::vector<std::uint32_t> order;
    if (field.sort.manual || !field.sort.by_field.empty()) {
      for (std::size_t k = 0; k < shared.size(); ++k) {
        order.push_back(static_cast<std::uint32_t>(k));
      }
    } else {
      const auto less = [&shared](std::uint32_t a, std::uint32_t b) { return value_less(shared[a], shared[b]); };
      sorted_index_order(order, static_cast<std::uint32_t>(shared.size()), make_index_less(less));
      if (!field.sort.ascending) {
        std::reverse(order.begin(), order.end());
      }
    }
    for (const std::uint32_t k : order) {
      if (listed(k)) {
        add(k);
      }
    }
    return;
  }
  // An item without a cache index names a shared item by label, else by position.
  std::vector<bool> used(shared.size(), false);
  for (std::size_t pos = 0; pos < field.items.size(); ++pos) {
    const PivotItem& item = field.items[pos];
    std::size_t target = item.has_cache_index ? item.cache_index : pos;
    if (!item.has_cache_index && !item.name.empty()) {
      for (std::size_t k = 0; k < shared.size(); ++k) {
        if (display_string(shared[k]) == item.name) {
          target = k;
          break;
        }
      }
    }
    if (target < used.size()) {
      used[target] = true;
    }
  }
  for (std::size_t k = 0; k < shared.size(); ++k) {
    if (!used[k] && listed(k)) {
      add(k);
    }
  }
}

void list_axis_field_items(PivotTable& table, const PivotCache& cache) {
  const std::size_t count = std::min(table.fields().size(), cache.fields().size());
  for (std::size_t f = 0; f < count; ++f) {
    PivotField& field = table.mutable_fields()[f];
    const bool on_axis = field.axis == PivotAxis::Row || field.axis == PivotAxis::Col || field.axis == PivotAxis::Page;
    if (on_axis && field.items.empty()) {
      keep_field_items(field, cache.fields()[f]);
    }
  }
}

}  // namespace formulon::pivot
