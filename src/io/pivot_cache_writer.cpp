//
// Implementation of the pivot-cache writer pair. See
// pivot_cache_writer.h for the public contract; see
// pivot_cache_reader.cpp for the symmetric grammar definition.

#include "io/pivot_cache_writer.h"

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <map>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/ooxml_defs.h"
#include "io/pivot_attr_names.h"
#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pivot/date_serial.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_cache.h"
#include "pivot/pivot_table.h"
#include "pivot/record_access.h"
#include "pivot/shared_items_hints.h"
#include "pivot/value_order.h"
#include "utils/index_sort.h"
#include "value.h"

namespace formulon::io {
namespace {

/// Emits one of `<s>`, `<n>`, `<b>`, `<m/>`, `<e>` for the given value, or
/// `<d>` for a number `as_date` marks as a valid date serial of the workbook's
/// date system.
/// Mirrors `DecodeTypedValue` in `pivot_cache_reader.cpp`. Anything other
/// than the five recognised kinds falls through to a `<m/>` so a
/// malformed cache cannot escape as malformed XML; this is a defensive
/// fallback that should not fire on a well-formed `PivotCache`.
void AppendInlineTypedValue(std::string& out, const Value& v, bool as_date, bool date1904) {
  if (as_date && v.is_number() && pivot::is_valid_date_serial(v.as_number(), date1904)) {
    out.append("<d");
    append_xml_attr(out, "v", pivot::group_date_iso(v.as_number(), date1904));
    out.append("/>");
    return;
  }
  if (v.is_text()) {
    out.append("<s");
    append_xml_attr(out, "v", v.as_text());
    out.append("/>");
    return;
  }
  if (v.is_number()) {
    out.append("<n");
    // Shared with the sheet writer so a value is spelled the same way
    // wherever it lands in the package. NaN / infinities never reach
    // here: the reader rejects non-numeric `<n v=...>` payloads, and a
    // caller wanting to keep such a payload has to encode it as an
    // Error value upstream.
    append_xml_attr_number(out, "v", v.as_number());
    out.append("/>");
    return;
  }
  if (v.is_boolean()) {
    out.append("<b");
    emit_xsd_bool_attr(out, "v", v.as_boolean());
    out.append("/>");
    return;
  }
  if (v.is_blank()) {
    out.append("<m/>");
    return;
  }
  if (v.is_error()) {
    out.append("<e");
    // Error display names are static literals (see `display_name` in
    // value.h); they contain no XML-critical characters but we still
    // route through the escaper to keep the encoding rule uniform.
    append_xml_attr(out, "v", display_name(v.as_error()));
    out.append("/>");
    return;
  }
  // Array / Ref / Lambda variants should never appear in a cache; emit
  // a blank placeholder rather than failing the writer (which has no
  // error channel). A well-formed `PivotCache` never trips this path.
  out.append("<m/>");
}

/// `AppendInlineTypedValue` for a shared item, adding `u="1"` to one a
/// refresh kept although the source no longer holds it and `f="1"` to a
/// calculated one.
void AppendSharedItem(std::string& out, const Value& v, bool as_date, bool missing, bool calculated, bool date1904) {
  AppendInlineTypedValue(out, v, as_date, date1904);
  if (out.size() < 2U || out.compare(out.size() - 2U, 2U, "/>") != 0) {
    return;
  }
  if (missing) {
    out.insert(out.size() - 2U, " u=\"1\"");
  }
  if (calculated) {
    out.insert(out.size() - 2U, " f=\"1\"");
  }
}

/// Emits `<x v="N"/>` for an integer-valued shared-items index. The
/// caller has already verified that the matching field is shared and
/// that the record cell is `Value::number(...)`. Negative or fractional
/// values would be reader-rejected on round-trip; we floor to the
/// nearest non-negative integer to keep the bytes well-formed even on
/// malformed input.
void AppendSharedIndex(std::string& out, double raw) {
  long long idx = (raw < 0.0) ? 0 : static_cast<long long>(raw);
  out.append("<x");
  append_xml_attr_int(out, "v", idx);
  out.append("/>");
}

/// Emits the captured `<sharedItems>` content-hint / range attributes
/// verbatim (present stays present with its value, absent stays absent).
/// Emits nothing when no hints were captured; callers pick the fallback
/// placeholder for the range-typed / empty case.
void AppendSharedItemsHints(std::string& out, const pivot::SharedItemsHints& h) {
  if (h.has_contains_semi_mixed) {
    emit_xsd_bool_attr(out, "containsSemiMixedTypes", h.contains_semi_mixed);
  }
  if (h.has_contains_non_date) {
    emit_xsd_bool_attr(out, "containsNonDate", h.contains_non_date);
  }
  if (h.has_contains_date) {
    emit_xsd_bool_attr(out, "containsDate", h.contains_date);
  }
  if (h.has_contains_string) {
    emit_xsd_bool_attr(out, "containsString", h.contains_string);
  }
  if (h.has_contains_blank) {
    emit_xsd_bool_attr(out, "containsBlank", h.contains_blank);
  }
  if (h.has_contains_mixed_types) {
    emit_xsd_bool_attr(out, "containsMixedTypes", h.contains_mixed_types);
  }
  if (h.has_contains_number) {
    emit_xsd_bool_attr(out, "containsNumber", h.contains_number);
  }
  if (h.has_contains_integer) {
    emit_xsd_bool_attr(out, "containsInteger", h.contains_integer);
  }
  if (h.has_min_value && !h.min_value.empty()) {
    append_xml_attr(out, "minValue", h.min_value);
  }
  if (h.has_max_value && !h.max_value.empty()) {
    append_xml_attr(out, "maxValue", h.max_value);
  }
  if (h.has_min_date && !h.min_date.empty()) {
    append_xml_attr(out, "minDate", h.min_date);
  }
  if (h.has_max_date && !h.max_date.empty()) {
    append_xml_attr(out, "maxDate", h.max_date);
  }
  append_raw_attrs(out, h.extra_attrs);
}

/// Sets a boolean content hint and its presence flag.
void SetHint(bool& flag, bool& has, bool value) {
  flag = value;
  has = true;
}

/// True when bit `i` of `flags` is set; a shorter vector reads as unset.
bool FlagAt(const std::vector<bool>& flags, std::size_t i) {
  return i < flags.size() && flags[i];
}

bool OnDiscreteAxis(pivot::PivotAxis axis) {
  return axis == pivot::PivotAxis::Row || axis == pivot::PivotAxis::Col || axis == pivot::PivotAxis::Page;
}

/// True for a range group that buckets dates: one with ISO bounds, or one
/// grouping by a calendar unit.
bool IsDateRange(const pivot::FieldGroup& group) {
  return group.kind == pivot::GroupKind::Range && (group.range.is_date || group.range.by != pivot::GroupBy::Range);
}

/// True when some field of `cache` is a date range group over field `fi`.
bool IsDateGroupBase(const pivot::PivotCache& cache, std::size_t fi) {
  for (const pivot::PivotCacheField& f : cache.fields()) {
    if (f.field_group.has_value() && f.field_group->base == fi && IsDateRange(*f.field_group)) {
      return true;
    }
  }
  return false;
}

/// The hints Excel writes for an all-date field; not `present` when a record
/// holds anything but a non-negative number.
pivot::SharedItemsHints DateHints(const pivot::PivotCache& cache, std::size_t fi, bool date1904) {
  pivot::SharedItemsHints h;
  double lo = 0.0;
  double hi = 0.0;
  std::size_t count = 0;
  for (const pivot::PivotCacheRecord& record : cache.records()) {
    const Value v = pivot::cell_value(cache, record, fi);
    if (!v.is_number() || !std::isfinite(v.as_number()) || v.as_number() < 0.0) {
      return h;
    }
    lo = count == 0 ? v.as_number() : std::min(lo, v.as_number());
    hi = count == 0 ? v.as_number() : std::max(hi, v.as_number());
    ++count;
  }
  if (count == 0) {
    return h;
  }
  SetHint(h.contains_semi_mixed, h.has_contains_semi_mixed, false);
  SetHint(h.contains_non_date, h.has_contains_non_date, false);
  SetHint(h.contains_date, h.has_contains_date, true);
  SetHint(h.contains_string, h.has_contains_string, false);
  const pivot::GroupBounds bounds = pivot::auto_date_group_bounds(lo, hi, date1904);
  h.has_min_date = true;
  h.min_date = pivot::group_date_iso(bounds.start, date1904);
  h.has_max_date = true;
  h.max_date = pivot::group_date_iso(bounds.end, date1904);
  h.present = true;
  return h;
}

/// Emits `<fieldGroup>`. `base_hints` are the hints written for the base
/// field; a date range with no bounds of its own takes theirs, as Excel's
/// `startDate` / `endDate` repeat the base's `minDate` / `maxDate`.
void AppendFieldGroup(std::string& out, const pivot::FieldGroup& group, const pivot::SharedItemsHints* base_hints,
                      bool date1904) {
  out.append("<fieldGroup");
  if (group.par.has_value()) {
    append_xml_attr_uint(out, "par", *group.par);
  }
  if (group.base.has_value()) {
    append_xml_attr_uint(out, "base", *group.base);
  }
  append_raw_attrs(out, group.extra_attrs);
  if (group.kind == pivot::GroupKind::None && group.group_items.empty()) {
    out.append("/>");
    return;
  }
  out.push_back('>');
  if (group.kind == pivot::GroupKind::Range) {
    const pivot::RangeGroup& r = group.range;
    out.append("<rangePr");
    if (!r.auto_start) {
      emit_xsd_bool_attr(out, "autoStart", false);
    }
    if (!r.auto_end) {
      emit_xsd_bool_attr(out, "autoEnd", false);
    }
    if (r.by != pivot::GroupBy::Range) {
      append_xml_attr(out, "groupBy", enum_name(kPivotGroupByNames, r.by));
    }
    if (IsDateRange(group)) {
      const bool from_base = base_hints != nullptr && base_hints->has_min_date && base_hints->has_max_date;
      const std::string& start = r.start_date.empty() && from_base ? base_hints->min_date : r.start_date;
      const std::string& end = r.end_date.empty() && from_base ? base_hints->max_date : r.end_date;
      if (!start.empty()) {
        append_xml_attr(out, "startDate", start);
      }
      if (!end.empty()) {
        append_xml_attr(out, "endDate", end);
      }
    } else if (r.by == pivot::GroupBy::Range) {
      append_xml_attr_number(out, "startNum", r.start_num);
      append_xml_attr_number(out, "endNum", r.end_num);
    }
    if (r.interval != 1.0) {
      append_xml_attr_number(out, "groupInterval", r.interval);
    }
    out.append("/>");
  } else if (group.kind == pivot::GroupKind::Discrete) {
    out.append("<discretePr");
    append_xml_attr_uint(out, "count", group.discrete_map.size());
    out.push_back('>');
    for (const std::uint32_t x : group.discrete_map) {
      out.append("<x");
      append_xml_attr_uint(out, "v", x);
      out.append("/>");
    }
    out.append("</discretePr>");
  }
  if (!group.group_items.empty()) {
    out.append("<groupItems");
    append_xml_attr_uint(out, "count", group.group_items.size());
    out.push_back('>');
    for (std::size_t i = 0; i < group.group_items.size(); ++i) {
      AppendInlineTypedValue(out, group.group_items[i], FlagAt(group.group_item_is_date, i), date1904);
    }
    out.append("</groupItems>");
  }
  out.append("</fieldGroup>");
}

/// Emits `<calculatedItems>`; nothing when the cache has none.
void AppendCalculatedItems(std::string& out, const std::vector<pivot::PivotCalculatedItem>& items) {
  if (items.empty()) {
    return;
  }
  out.append("<calculatedItems");
  append_xml_attr_uint(out, "count", items.size());
  out.push_back('>');
  for (const pivot::PivotCalculatedItem& item : items) {
    out.append("<calculatedItem");
    if (item.has_field_attr) {
      append_xml_attr_uint(out, "field", item.cache_field_index);
    }
    append_xml_attr(out, "formula", item.formula);
    if (!item.has_area && item.extra_children_xml.empty()) {
      out.append("/>");
      continue;
    }
    out.push_back('>');
    const pivot::PivotArea& area = item.area;
    if (item.has_area) {
      out.append("<pivotArea");
      if (area.cache_index.has_value()) {
        emit_xsd_bool_attr(out, "cacheIndex", *area.cache_index);
      }
      if (area.outline.has_value()) {
        emit_xsd_bool_attr(out, "outline", *area.outline);
      }
      if (area.field_position.has_value()) {
        append_xml_attr_uint(out, "fieldPosition", *area.field_position);
      }
      append_raw_attrs(out, area.extra_attrs);
      if (area.references.empty() && area.extra_children_xml.empty()) {
        out.append("/>");
      } else {
        out.push_back('>');
        if (!area.references.empty()) {
          out.append("<references");
          append_xml_attr_uint(out, "count", area.references.size());
          out.push_back('>');
          for (const pivot::PivotAreaReference& ref : area.references) {
            out.append("<reference");
            if (ref.field.has_value()) {
              append_xml_attr_uint(out, "field", *ref.field);
            }
            append_xml_attr_uint(out, "count", ref.items.size());
            append_raw_attrs(out, ref.extra_attrs);
            out.push_back('>');
            for (const std::uint32_t x : ref.items) {
              out.append("<x");
              append_xml_attr_uint(out, "v", x);
              out.append("/>");
            }
            out.append(ref.extra_children_xml);
            out.append("</reference>");
          }
          out.append("</references>");
        }
        out.append(area.extra_children_xml);
        out.append("</pivotArea>");
      }
    }
    out.append(item.extra_children_xml);
    out.append("</calculatedItem>");
  }
  out.append("</calculatedItems>");
}

}  // namespace

PivotCacheWritePlan build_pivot_cache_write_plan(const pivot::PivotCache& cache,
                                                 const std::vector<const pivot::PivotTable*>& tables, bool date1904) {
  PivotCacheWritePlan plan;
  plan.date1904 = date1904;
  plan.refresh_on_load = cache.passthrough_attrs().empty();
  for (const pivot::PivotTable* table : tables) {
    plan.refresh_on_load = plan.refresh_on_load || (table != nullptr && table->layout_dirty());
  }
  plan.fields.resize(cache.fields().size());
  for (std::size_t fi = 0; fi < cache.fields().size(); ++fi) {
    const pivot::PivotCacheField& field = cache.fields()[fi];
    if (!field.is_database_field || !field.shared_items.empty() || field.shared_items_hints.present) {
      continue;
    }
    if (field.field_group.has_value()) {
      if (IsDateGroupBase(cache, fi)) {
        plan.fields[fi].hints = DateHints(cache, fi, date1904);
        plan.fields[fi].dates = plan.fields[fi].hints.present;
      }
      continue;
    }
    bool wanted = false;
    for (const pivot::PivotTable* table : tables) {
      if (table != nullptr && fi < table->fields().size()) {
        const pivot::PivotField& pf = table->fields()[fi];
        wanted = wanted || OnDiscreteAxis(pf.axis) || !pf.items.empty();
      }
    }
    for (std::size_t r = 0; r < cache.records().size() && !wanted; ++r) {
      const Value v = pivot::cell_value(cache, cache.records()[r], fi);
      wanted = v.is_text() || v.is_boolean() || v.is_error() || v.is_blank();
    }
    PivotCacheFieldWritePlan& fp = plan.fields[fi];
    if (!wanted) {
      std::vector<Value> values;
      std::vector<bool> marks;
      values.reserve(cache.records().size());
      for (const pivot::PivotCacheRecord& record : cache.records()) {
        values.push_back(pivot::cell_value(cache, record, fi));
        marks.push_back(FlagAt(record.cell_is_date, fi));
      }
      fp.hints = pivot::derive_shared_items_hints(values, marks, date1904);
      continue;
    }
    std::map<Value, std::uint32_t, pivot::ValueLess> seen;
    fp.record_index.reserve(cache.records().size());
    for (const pivot::PivotCacheRecord& record : cache.records()) {
      const Value v = pivot::cell_value(cache, record, fi);
      const auto inserted = seen.emplace(v, static_cast<std::uint32_t>(fp.items.size()));
      if (inserted.second) {
        fp.items.push_back(v);
        fp.item_is_date.push_back(FlagAt(record.cell_is_date, fi));
      }
      fp.record_index.push_back(inserted.first->second);
    }
    fp.derived = true;
    fp.hints = pivot::derive_shared_items_hints(fp.items, fp.item_is_date, date1904);
    const auto item_less = [&fp](std::uint32_t a, std::uint32_t b) {
      return pivot::value_less(fp.items[a], fp.items[b]);
    };
    sorted_index_order(fp.asc_order, static_cast<std::uint32_t>(fp.items.size()), make_index_less(item_less));
  }
  return plan;
}

std::string write_pivot_cache_definition(const pivot::PivotCache& cache, const PivotCacheWritePlan* plan) {
  const bool date1904 = plan != nullptr && plan->date1904;
  std::string out;
  // Pre-reserve a rough lower bound: declaration + root attrs + per-field
  // cost. Empirically ~80B per shared item, ~120B per field; this trims
  // a couple of growth-pass reallocations on large caches without
  // bloating the small-cache path.
  std::size_t shared_items_total = 0;
  for (const pivot::PivotCacheField& f : cache.fields()) {
    shared_items_total += f.shared_items.size();
  }
  out.reserve(256 + cache.fields().size() * 96 + shared_items_total * 48);

  out.append(kXmlDecl);
  out.append("<pivotCacheDefinition");
  append_xml_attr(out, "xmlns", kNsSpreadsheetMl);
  append_xml_attr(out, "xmlns:r", kNsOfficeRelationships);
  if (cache.save_data()) {
    append_xml_attr(out, "r:id", "rId1");
  }
  append_xml_attr_uint(out, "recordCount", cache.records().size());
  // A cache built in memory has no rendered pivot cells behind it, and an
  // edited table drops the ones it was read with, so either asks Excel to
  // rebuild the report on open. Otherwise the captured attributes stand.
  const bool refresh_on_load =
      (plan != nullptr ? plan->refresh_on_load : cache.passthrough_attrs().empty()) || cache.refresh_on_load();
  if (refresh_on_load) {
    append_xml_attr(out, "refreshOnLoad", "1");
  } else if (cache.refresh_on_load_attr().has_value()) {
    append_xml_attr(out, "refreshOnLoad", "0");
  }
  if (!cache.save_data()) {
    emit_xsd_bool_attr(out, "saveData", false);
  }
  if (cache.missing_items_limit().has_value()) {
    append_xml_attr_uint(out, "missingItemsLimit", *cache.missing_items_limit());
  }
  // Re-emit any unmodelled root attributes captured on read (refreshedBy,
  // refreshedDate, createdVersion, ...).
  for (const auto& attr : cache.passthrough_attrs()) {
    if (!refresh_on_load || attr.first != "refreshOnLoad") {
      append_xml_attr(out, attr.first, attr.second);
    }
  }
  out.append(">");

  // Re-emit the `<cacheSource>` with its `<worksheetSource>` child when
  // the reader captured one, so Excel's Refresh can locate the source
  // range / defined name. Falls back to the minimal self-closing form
  // for caches built from scratch (no source captured).
  const pivot::WorksheetSource& wsrc = cache.worksheet_source();
  out.append("<cacheSource");
  if (cache.cache_source_attrs().empty()) {
    append_xml_attr(out, "type", "worksheet");
  } else {
    append_raw_attrs(out, cache.cache_source_attrs());
  }
  if (wsrc.present) {
    out.append("><worksheetSource");
    if (!wsrc.ref.empty()) {
      append_xml_attr(out, "ref", wsrc.ref);
    }
    if (!wsrc.sheet.empty()) {
      append_xml_attr(out, "sheet", wsrc.sheet);
    }
    if (!wsrc.name.empty()) {
      append_xml_attr(out, "name", wsrc.name);
    }
    out.append("/>");
    out.append(cache.cache_source_children_xml());
    out.append("</cacheSource>");
  } else if (!cache.cache_source_children_xml().empty()) {
    out.append(">");
    out.append(cache.cache_source_children_xml());
    out.append("</cacheSource>");
  } else {
    out.append("/>");
  }

  out.append("<cacheFields");
  append_xml_attr_uint(out, "count", cache.fields().size());
  out.append(">");

  std::size_t field_index = 0;
  for (const pivot::PivotCacheField& field : cache.fields()) {
    const PivotCacheFieldWritePlan* field_plan =
        plan != nullptr && field_index < plan->fields.size() ? &plan->fields[field_index] : nullptr;
    const PivotCacheFieldWritePlan* derived = field_plan != nullptr && field_plan->derived ? field_plan : nullptr;
    ++field_index;
    out.append("<cacheField");
    append_xml_attr(out, "name", field.name);
    if (field.num_fmt_id.has_value()) {
      append_xml_attr_uint(out, "numFmtId", *field.num_fmt_id);
    }
    if (field.formula.has_value()) {
      append_xml_attr(out, "formula", *field.formula);
    }
    // `databaseField` defaults to true; emit `="0"` only for a
    // grouping-derived field so it is excluded from record output on read.
    if (!field.is_database_field) {
      emit_xsd_bool_attr(out, "databaseField", false);
    }
    append_raw_attrs(out, field.extra_attrs);
    if (field.formula.has_value()) {
      // A calculated field carries no shared items or group.
      if (field.trailing_children_xml.empty()) {
        out.append("/>");
      } else {
        out.append(">");
        out.append(field.trailing_children_xml);
        out.append("</cacheField>");
      }
      continue;
    }
    out.append(">");

    const bool group_field = !field.is_database_field && field.field_group.has_value() && field.shared_items.empty() &&
                             !field.shared_items_hints.present;
    if (group_field) {
      // A group field takes its items from `<groupItems>`; Excel writes no
      // `<sharedItems>` for it.
    } else if (derived != nullptr) {
      out.append("<sharedItems");
      AppendSharedItemsHints(out, derived->hints);
      append_xml_attr_uint(out, "count", derived->items.size());
      out.append(">");
      for (std::size_t i = 0; i < derived->items.size(); ++i) {
        AppendInlineTypedValue(out, derived->items[i], FlagAt(derived->item_is_date, i), date1904);
      }
      out.append("</sharedItems>");
    } else if (field.shared_items.empty()) {
      // Range-typed field: re-emit the captured numeric / date range +
      // grouping hints so Excel's Refresh keeps its grouping boundaries.
      // Falls back to a minimal `containsNumber="1"` placeholder when the
      // field was built from scratch with no records to derive hints from.
      out.append("<sharedItems");
      if (field.shared_items_hints.present) {
        AppendSharedItemsHints(out, field.shared_items_hints);
      } else if (field_plan != nullptr && field_plan->hints.present) {
        AppendSharedItemsHints(out, field_plan->hints);
      } else {
        emit_xsd_bool_attr(out, "containsNumber", true);
      }
      out.append("/>");
    } else {
      // Discrete field: emit any captured content hints (containsBlank /
      // containsString / ...) alongside the item count so a discrete
      // field's hints survive the round trip instead of being dropped.
      out.append("<sharedItems");
      const auto any_set = [](const std::vector<bool>& flags) {
        return std::find(flags.begin(), flags.end(), true) != flags.end();
      };
      if (field.shared_items_hints.present || any_set(field.shared_item_is_calculated)) {
        AppendSharedItemsHints(out, field.shared_items_hints);
      } else {
        // A field built without captured hints: Excel asks to repair shared
        // items that carry no content hints.
        AppendSharedItemsHints(
            out, pivot::derive_shared_items_hints(field.shared_items, field.shared_item_is_date, date1904));
      }
      append_xml_attr_uint(out, "count", field.shared_items.size());
      out.append(">");
      for (std::size_t i = 0; i < field.shared_items.size(); ++i) {
        AppendSharedItem(out, field.shared_items[i], FlagAt(field.shared_item_is_date, i),
                         FlagAt(field.shared_item_is_missing, i), FlagAt(field.shared_item_is_calculated, i), date1904);
      }
      out.append("</sharedItems>");
    }
    if (field.field_group.has_value()) {
      const std::optional<std::uint32_t> base = field.field_group->base;
      const pivot::SharedItemsHints* base_hints = nullptr;
      if (base.has_value() && *base < cache.fields().size()) {
        base_hints = &cache.fields()[*base].shared_items_hints;
        if (!base_hints->present && plan != nullptr && *base < plan->fields.size()) {
          base_hints = &plan->fields[*base].hints;
        }
      }
      AppendFieldGroup(out, *field.field_group, base_hints, date1904);
    }
    out.append(field.trailing_children_xml);
    out.append("</cacheField>");
  }

  out.append("</cacheFields>");
  out.append(cache.trailing_children_xml());
  AppendCalculatedItems(out, cache.calculated_items());
  out.append(cache.trailing_children_after_calculated_xml());
  out.append("</pivotCacheDefinition>");
  return out;
}

std::string write_pivot_cache_records(const pivot::PivotCache& cache, const PivotCacheWritePlan* plan) {
  const bool date1904 = plan != nullptr && plan->date1904;
  std::string out;
  // Per-record cost is roughly 7B (`<r></r>`) + per-cell cost. `<x v="N"/>`
  // is ~10B for small N, inline-typed cells run ~12B for blanks up to
  // ~60B for short text payloads. Reserve a conservative lower bound to
  // skip the first few reallocations.
  const std::size_t field_count = cache.fields().size();
  out.reserve(128 + cache.records().size() * (16 + field_count * 12));

  out.append(kXmlDecl);
  out.append("<pivotCacheRecords");
  append_xml_attr(out, "xmlns", kNsSpreadsheetMl);
  append_xml_attr_uint(out, "count", cache.records().size());
  out.append(">");

  std::size_t record_ordinal = 0;
  for (const pivot::PivotCacheRecord& record : cache.records()) {
    const std::size_t r = record_ordinal++;
    out.append("<r>");
    for (std::size_t i = 0; i < field_count; ++i) {
      // Only database fields carry a per-record cell; grouping-derived
      // fields (`databaseField="0"`) are excluded so the record arity
      // matches Excel's (one cell per database field).
      if (!cache.fields()[i].is_database_field) {
        continue;
      }
      if (plan != nullptr && i < plan->fields.size() && plan->fields[i].derived &&
          r < plan->fields[i].record_index.size()) {
        out.append("<x");
        append_xml_attr_uint(out, "v", plan->fields[i].record_index[r]);
        out.append("/>");
        continue;
      }
      // Missing trailing cells become `<m/>` so every `<r>` has exactly one
      // child per database field; the reader pre-fills symmetrically.
      if (i >= record.cells.size()) {
        out.append("<m/>");
        continue;
      }
      const Value& cell = record.cells[i];
      const bool emit_index = pivot::cell_is_shared_index(record, cache.fields()[i], i) && cell.is_number();
      const bool as_date =
          FlagAt(record.cell_is_date, i) || (plan != nullptr && i < plan->fields.size() && plan->fields[i].dates);
      if (emit_index) {
        AppendSharedIndex(out, cell.as_number());
      } else {
        AppendInlineTypedValue(out, cell, as_date, date1904);
      }
    }
    out.append("</r>");
  }

  out.append("</pivotCacheRecords>");
  return out;
}

}  // namespace formulon::io
