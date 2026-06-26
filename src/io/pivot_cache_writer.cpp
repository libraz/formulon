// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of the pivot-cache writer pair. See
// pivot_cache_writer.h for the public contract; see
// pivot_cache_reader.cpp for the symmetric grammar definition.

#include "io/pivot_cache_writer.h"

#include <cmath>
#include <cstddef>
#include <cstdint>
#include <cstdio>
#include <string>
#include <string_view>

#include "io/xml_escape.h"
#include "pivot/pivot_cache.h"
#include "value.h"

namespace formulon::io {
namespace {

constexpr std::string_view kXmlDecl = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
constexpr std::string_view kPivotNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
constexpr std::string_view kRelsNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Appends a double using `%.17g` (round-trip safe under IEEE 754) with a
/// fast-path for zero so positive and negative zero render identically.
/// NaN / infinities should never reach this writer (the reader rejects
/// non-numeric `<n v=...>` payloads via strtod) and are deliberately not
/// special-cased here -- they would emit "nan"/"inf" which would fail to
/// round-trip back through the reader's `ParseDouble`. Callers that wish
/// to preserve such payloads must encode them as Error values upstream.
/// This mirrors `AppendNumberValue` in `ooxml_writer_cell.cpp`.
void AppendNumber(std::string& out, double v) {
  if (v == 0.0) {
    out.push_back('0');
    return;
  }
  char buf[32];
  std::snprintf(buf, sizeof(buf), "%.17g", v);
  out.append(buf);
}

/// Emits one of `<s>`, `<n>`, `<b>`, `<m/>`, `<e>` for the given value.
/// Mirrors `DecodeTypedValue` in `pivot_cache_reader.cpp`. Anything other
/// than the five recognised kinds falls through to a `<m/>` so a
/// malformed cache cannot escape as malformed XML; this is a defensive
/// fallback that should not fire on a well-formed `PivotCache`.
void AppendInlineTypedValue(std::string& out, const Value& v) {
  if (v.is_text()) {
    out.append("<s v=\"");
    AppendXmlEscaped(out, v.as_text());
    out.append("\"/>");
    return;
  }
  if (v.is_number()) {
    out.append("<n v=\"");
    AppendNumber(out, v.as_number());
    out.append("\"/>");
    return;
  }
  if (v.is_boolean()) {
    out.append("<b v=\"");
    out.push_back(v.as_boolean() ? '1' : '0');
    out.append("\"/>");
    return;
  }
  if (v.is_blank()) {
    out.append("<m/>");
    return;
  }
  if (v.is_error()) {
    out.append("<e v=\"");
    // Error display names are static literals (see `display_name` in
    // value.h); they contain no XML-critical characters but we still
    // route through the escaper to keep the encoding rule uniform.
    AppendXmlEscaped(out, display_name(v.as_error()));
    out.append("\"/>");
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
  out.append("<x v=\"");
  out.append(std::to_string(idx));
  out.append("\"/>");
}

/// Appends ` name="value"` when `value` is non-empty, escaping the body.
/// Used for the `<worksheetSource>` ref / sheet / name attributes.
void AppendOptionalAttr(std::string& out, std::string_view name, std::string_view value) {
  if (value.empty()) {
    return;
  }
  out.push_back(' ');
  out.append(name);
  out.append("=\"");
  AppendXmlEscaped(out, value);
  out.push_back('"');
}

/// Appends ` name="1"` when `flag` is set. The OOXML `<sharedItems>`
/// content hints are boolean attributes that default to false; only
/// emit the set ones to keep the output lean.
void AppendBoolHint(std::string& out, std::string_view name, bool flag) {
  if (!flag) {
    return;
  }
  out.push_back(' ');
  out.append(name);
  out.append("=\"1\"");
}

/// Emits the `<sharedItems>` opening attributes for a range-typed field
/// from the captured hints. When no hints were captured, falls back to
/// the legacy minimal `containsNumber="1"` placeholder so freshly built
/// caches still produce schema-valid output.
void AppendSharedItemsHints(std::string& out, const pivot::SharedItemsHints& h) {
  if (!h.present) {
    out.append(" containsNumber=\"1\"");
    return;
  }
  AppendBoolHint(out, "containsSemiMixedTypes", h.contains_semi_mixed);
  AppendBoolHint(out, "containsNonDate", h.contains_non_date);
  AppendBoolHint(out, "containsDate", h.contains_date);
  AppendBoolHint(out, "containsString", h.contains_string);
  AppendBoolHint(out, "containsBlank", h.contains_blank);
  AppendBoolHint(out, "containsMixedTypes", h.contains_mixed_types);
  AppendBoolHint(out, "containsNumber", h.contains_number);
  AppendBoolHint(out, "containsInteger", h.contains_integer);
  if (h.has_min_value) {
    AppendOptionalAttr(out, "minValue", h.min_value);
  }
  if (h.has_max_value) {
    AppendOptionalAttr(out, "maxValue", h.max_value);
  }
  if (h.has_min_date) {
    AppendOptionalAttr(out, "minDate", h.min_date);
  }
  if (h.has_max_date) {
    AppendOptionalAttr(out, "maxDate", h.max_date);
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
  out.append("<pivotCacheDefinition xmlns=\"");
  out.append(kPivotNs);
  out.append("\" xmlns:r=\"");
  out.append(kRelsNs);
  out.append("\" r:id=\"rId1\" recordCount=\"");
  out.append(std::to_string(cache.records().size()));
  out.append("\">");

  // Re-emit the `<cacheSource>` with its `<worksheetSource>` child when
  // the reader captured one, so Excel's Refresh can locate the source
  // range / defined name. Falls back to the minimal self-closing form
  // for caches built from scratch (no source captured).
  const pivot::WorksheetSource& wsrc = cache.worksheet_source();
  if (wsrc.present) {
    out.append("<cacheSource type=\"worksheet\"><worksheetSource");
    AppendOptionalAttr(out, "ref", wsrc.ref);
    AppendOptionalAttr(out, "sheet", wsrc.sheet);
    AppendOptionalAttr(out, "name", wsrc.name);
    out.append("/></cacheSource>");
  } else {
    out.append("<cacheSource type=\"worksheet\"/>");
  }

  out.append("<cacheFields count=\"");
  out.append(std::to_string(cache.fields().size()));
  out.append("\">");

  for (const pivot::PivotCacheField& field : cache.fields()) {
    out.append("<cacheField name=\"");
    AppendXmlEscaped(out, field.name);
    out.append("\">");

    if (field.shared_items.empty()) {
      // Range-typed field: re-emit the captured numeric / date range +
      // grouping hints so Excel's Refresh keeps its grouping boundaries.
      // Falls back to a minimal `containsNumber="1"` placeholder when the
      // field was built from scratch (no hints captured).
      out.append("<sharedItems");
      AppendSharedItemsHints(out, field.shared_items_hints);
      out.append("/>");
    } else {
      out.append("<sharedItems count=\"");
      out.append(std::to_string(field.shared_items.size()));
      out.append("\">");
      for (const Value& item : field.shared_items) {
        AppendInlineTypedValue(out, item);
      }
      out.append("</sharedItems>");
    }
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
  out.append("<pivotCacheRecords xmlns=\"");
  out.append(kPivotNs);
  out.append("\" count=\"");
  out.append(std::to_string(cache.records().size()));
  out.append("\">");

  for (const pivot::PivotCacheRecord& record : cache.records()) {
    out.append("<r>");
    for (std::size_t i = 0; i < field_count; ++i) {
      // Pad missing trailing cells with `<m/>` so every `<r>` has exactly
      // `field_count` children. The reader pads symmetrically on read,
      // but emitting them explicitly keeps the bytes self-describing.
      if (i >= record.cells.size()) {
        out.append("<m/>");
        continue;
      }
      const Value& cell = record.cells[i];
      const bool field_is_shared = !cache.fields()[i].shared_items.empty();
      if (field_is_shared && cell.is_number()) {
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
