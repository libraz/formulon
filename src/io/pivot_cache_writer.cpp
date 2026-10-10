//
// Implementation of the pivot-cache writer pair. See
// pivot_cache_writer.h for the public contract; see
// pivot_cache_reader.cpp for the symmetric grammar definition.

#include "io/pivot_cache_writer.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>

#include "io/xml_utils.h"
#include "io/xsd_bool.h"
#include "pivot/pivot_cache.h"
#include "value.h"

namespace formulon::io {
namespace {

constexpr std::string_view kPivotNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
constexpr std::string_view kRelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Emits one of `<s>`, `<n>`, `<b>`, `<m/>`, `<e>` for the given value.
/// Mirrors `DecodeTypedValue` in `pivot_cache_reader.cpp`. Anything other
/// than the five recognised kinds falls through to a `<m/>` so a
/// malformed cache cannot escape as malformed XML; this is a defensive
/// fallback that should not fire on a well-formed `PivotCache`.
void AppendInlineTypedValue(std::string& out, const Value& v) {
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
}

}  // namespace

std::string write_pivot_cache_definition(const pivot::PivotCache& cache) {
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
  append_xml_attr(out, "xmlns", kPivotNs);
  append_xml_attr(out, "xmlns:r", kRelsNs);
  append_xml_attr(out, "r:id", "rId1");
  append_xml_attr_uint(out, "recordCount", cache.records().size());
  // Re-emit any unmodelled root attributes captured on read (refreshedBy,
  // refreshedDate, createdVersion, ...).
  append_raw_attrs(out, cache.passthrough_attrs());
  out.append(">");

  // Re-emit the `<cacheSource>` with its `<worksheetSource>` child when
  // the reader captured one, so Excel's Refresh can locate the source
  // range / defined name. Falls back to the minimal self-closing form
  // for caches built from scratch (no source captured).
  const pivot::WorksheetSource& wsrc = cache.worksheet_source();
  if (wsrc.present) {
    out.append("<cacheSource type=\"worksheet\"><worksheetSource");
    if (!wsrc.ref.empty()) {
      append_xml_attr(out, "ref", wsrc.ref);
    }
    if (!wsrc.sheet.empty()) {
      append_xml_attr(out, "sheet", wsrc.sheet);
    }
    if (!wsrc.name.empty()) {
      append_xml_attr(out, "name", wsrc.name);
    }
    out.append("/></cacheSource>");
  } else {
    out.append("<cacheSource type=\"worksheet\"/>");
  }

  out.append("<cacheFields");
  append_xml_attr_uint(out, "count", cache.fields().size());
  out.append(">");

  for (const pivot::PivotCacheField& field : cache.fields()) {
    out.append("<cacheField");
    append_xml_attr(out, "name", field.name);
    // `databaseField` defaults to true; emit `="0"` only for a
    // grouping-derived field so it is excluded from record output on read.
    if (!field.is_database_field) {
      emit_xsd_bool_attr(out, "databaseField", false);
    }
    out.append(">");

    if (field.shared_items.empty()) {
      // Range-typed field: re-emit the captured numeric / date range +
      // grouping hints so Excel's Refresh keeps its grouping boundaries.
      // Falls back to a minimal `containsNumber="1"` placeholder when the
      // field was built from scratch (no hints captured).
      out.append("<sharedItems");
      if (field.shared_items_hints.present) {
        AppendSharedItemsHints(out, field.shared_items_hints);
      } else {
        emit_xsd_bool_attr(out, "containsNumber", true);
      }
      out.append("/>");
    } else {
      // Discrete field: emit any captured content hints (containsBlank /
      // containsString / ...) alongside the item count so a discrete
      // field's hints survive the round trip instead of being dropped.
      out.append("<sharedItems");
      AppendSharedItemsHints(out, field.shared_items_hints);
      append_xml_attr_uint(out, "count", field.shared_items.size());
      out.append(">");
      for (const Value& item : field.shared_items) {
        AppendInlineTypedValue(out, item);
      }
      out.append("</sharedItems>");
    }
    // Re-emit the captured `<fieldGroup>` verbatim (grouping definition).
    out.append(field.field_group_xml);
    out.append("</cacheField>");
  }

  out.append("</cacheFields>");
  out.append("</pivotCacheDefinition>");
  return out;
}

std::string write_pivot_cache_records(const pivot::PivotCache& cache) {
  std::string out;
  // Per-record cost is roughly 7B (`<r></r>`) + per-cell cost. `<x v="N"/>`
  // is ~10B for small N, inline-typed cells run ~12B for blanks up to
  // ~60B for short text payloads. Reserve a conservative lower bound to
  // skip the first few reallocations.
  const std::size_t field_count = cache.fields().size();
  out.reserve(128 + cache.records().size() * (16 + field_count * 12));

  out.append(kXmlDecl);
  out.append("<pivotCacheRecords");
  append_xml_attr(out, "xmlns", kPivotNs);
  append_xml_attr_uint(out, "count", cache.records().size());
  out.append(">");

  for (const pivot::PivotCacheRecord& record : cache.records()) {
    out.append("<r>");
    for (std::size_t i = 0; i < field_count; ++i) {
      // Only database fields carry a per-record cell; grouping-derived
      // fields (`databaseField="0"`) are excluded so the record arity
      // matches Excel's (one cell per database field).
      if (!cache.fields()[i].is_database_field) {
        continue;
      }
      // Missing trailing cells become `<m/>` so every `<r>` has exactly one
      // child per database field; the reader pre-fills symmetrically.
      if (i >= record.cells.size()) {
        out.append("<m/>");
        continue;
      }
      const Value& cell = record.cells[i];
      // Prefer the per-cell encoding flag captured on read; fall back to
      // inferring "shared field + numeric cell = index" for hand-built
      // caches that leave `cell_is_index` empty.
      bool emit_index;
      if (i < record.cell_is_index.size()) {
        emit_index = record.cell_is_index[i] && cell.is_number();
      } else {
        emit_index = !cache.fields()[i].shared_items.empty() && cell.is_number();
      }
      if (emit_index) {
        AppendSharedIndex(out, cell.as_number());
      } else {
        AppendInlineTypedValue(out, cell);
      }
    }
    out.append("</r>");
  }

  out.append("</pivotCacheRecords>");
  return out;
}

}  // namespace formulon::io
