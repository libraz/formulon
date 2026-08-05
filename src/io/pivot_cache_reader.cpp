//
// Implementation of the pivot-cache reader pair. See
// pivot_cache_reader.h for the public contract.

#include "io/pivot_cache_reader.h"

#include <cstddef>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <deque>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/iso_date.h"
#include "io/xml_utils.h"
#include "pivot/pivot_cache.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "value.h"

namespace formulon::io {
namespace {

/// Maps an Excel error display name (e.g. `"#DIV/0!"`) to its `ErrorCode`.
/// Unknown spellings fall back to `ErrorCode::Value`, matching how Excel
/// itself reports unrecognised cache error payloads to users.
ErrorCode ParseErrorDisplay(std::string_view text) {
  if (text == "#NULL!")
    return ErrorCode::Null;
  if (text == "#DIV/0!")
    return ErrorCode::Div0;
  if (text == "#VALUE!")
    return ErrorCode::Value;
  if (text == "#REF!")
    return ErrorCode::Ref;
  if (text == "#NAME?")
    return ErrorCode::Name;
  if (text == "#NUM!")
    return ErrorCode::Num;
  if (text == "#N/A")
    return ErrorCode::NA;
  if (text == "#GETTING_DATA")
    return ErrorCode::GettingData;
  if (text == "#SPILL!")
    return ErrorCode::Spill;
  if (text == "#CALC!")
    return ErrorCode::Calc;
  if (text == "#FIELD!")
    return ErrorCode::Field;
  if (text == "#BLOCKED!")
    return ErrorCode::Blocked;
  if (text == "#CONNECT!")
    return ErrorCode::Connect;
  if (text == "#EXTERNAL!")
    return ErrorCode::External;
  if (text == "#BUSY!")
    return ErrorCode::Busy;
  if (text == "#PYTHON!")
    return ErrorCode::Python;
  if (text == "#UNKNOWN!")
    return ErrorCode::Unknown;
  return ErrorCode::Value;
}

/// Parses a numeric attribute body as a locale-independent double.
/// Returns false on empty input or trailing non-whitespace garbage; the
/// out-parameter is unchanged in that case. Mirrors the helper in
/// `cell_parser.cpp` so the pivot path produces identical numeric
/// behaviour (Excel writes round-trip-friendly decimal strings, never
/// localised).
bool ParseDouble(std::string_view text, double* out) {
  if (text.empty()) {
    return false;
  }
  char small_buf[64];
  const char* nstr = nullptr;
  std::string heap;
  if (text.size() < sizeof(small_buf)) {
    std::memcpy(small_buf, text.data(), text.size());
    small_buf[text.size()] = '\0';
    nstr = small_buf;
  } else {
    heap.assign(text.data(), text.size());
    nstr = heap.c_str();
  }
  char* end = nullptr;
  const double v = std::strtod(nstr, &end);
  if (end == nstr) {
    return false;
  }
  while (end != nullptr && *end != '\0') {
    if (*end != ' ' && *end != '\t' && *end != '\r' && *end != '\n') {
      return false;
    }
    ++end;
  }
  *out = v;
  return true;
}

/// Returns true iff the boolean attribute body is the OOXML literal
/// `"1"` (true) or `"0"` (false). On unrecognised input, defaults to
/// `false` and reports failure via the `*ok` flag.
bool ParseBoolFlag(std::string_view text, bool* ok) {
  if (text == "1" || text == "true") {
    *ok = true;
    return true;
  }
  if (text == "0" || text == "false" || text.empty()) {
    *ok = true;
    return false;
  }
  *ok = false;
  return false;
}

/// Decodes one value-bearing element (`<s>`, `<n>`, `<b>`, `<m>`, `<e>`)
/// into a `Value`. Used both for `<sharedItems>` children and for
/// inline-typed `<r>` cells in the records part. Text payloads are
/// appended into the caller-supplied `text_storage` (a deque for
/// pointer-stability across appends) and `Value::text` aliases that
/// entry. Returns `kIoSheetCorrupt` for malformed numeric / boolean
/// bodies.
Expected<Value, Error> DecodeTypedValue(const pugi::xml_node& node, std::deque<std::string>& text_storage) {
  const std::string_view name = node.name();
  if (name == "s") {
    text_storage.emplace_back(node.attribute("v").as_string());
    return Value::text(text_storage.back());
  }
  if (name == "n") {
    std::string_view raw = node.attribute("v").as_string();
    double num = 0.0;
    if (!ParseDouble(raw, &num)) {
      std::string ctx("context=pivot_cache_reader v=");
      ctx.append(raw);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivot cache: <n v=...> not a number", std::move(ctx));
    }
    return Value::number(num);
  }
  if (name == "b") {
    std::string_view raw = node.attribute("v").as_string();
    bool ok = false;
    const bool flag = ParseBoolFlag(raw, &ok);
    if (!ok) {
      std::string ctx("context=pivot_cache_reader v=");
      ctx.append(raw);
      return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivot cache: <b v=...> must be 0 or 1", std::move(ctx));
    }
    return Value::boolean(flag);
  }
  if (name == "d") {
    // `<d v="YYYY-MM-DDThh:mm:ss">` is a typed date item. Resolve the
    // ISO 8601 body to an Excel serial so the value participates in date
    // arithmetic like any other number. A non-conforming body that does
    // not parse degrades to text rather than failing the whole cache;
    // the important invariant is that a `<d>` always consumes a field
    // slot (see `IsTypedValueElement`) so record/field columns stay
    // aligned.
    std::string_view raw = node.attribute("v").as_string();
    double serial = 0.0;
    if (parse_iso_date_serial(raw, &serial)) {
      return Value::number(serial);
    }
    text_storage.emplace_back(raw);
    return Value::text(text_storage.back());
  }
  if (name == "m") {
    return Value::blank();
  }
  if (name == "e") {
    return Value::error(ParseErrorDisplay(node.attribute("v").as_string()));
  }
  // Sentinel: caller must check `is_typed_value_node` before calling.
  return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivot cache: unexpected child element",
                    std::string("context=pivot_cache_reader element=").append(name));
}

/// True if `name` is one of the inline-typed value elements that
/// `DecodeTypedValue` understands (`<s>`, `<n>`, `<b>`, `<d>`, `<m>`,
/// `<e>`). Used to gate decoding and, critically, to decide which
/// children consume a field slot in the records part: a typed `<d>`
/// date item must advance the field index exactly like `<n>`, otherwise
/// every subsequent column shifts by one.
bool IsTypedValueElement(std::string_view name) {
  return name == "s" || name == "n" || name == "b" || name == "d" || name == "m" || name == "e";
}

/// Reads the numeric / date range + content hint attributes from a
/// `<sharedItems>` element into `out`. Boolean hints default to false
/// (absent); the bound attributes (`minValue`/`maxValue`/`minDate`/
/// `maxDate`) record both presence and the raw string body so the writer
/// re-emits exactly what was read. `out->present` is set when any hint
/// attribute was found, gating the writer's legacy placeholder fallback.
void ReadSharedItemsHints(const pugi::xml_node& items, pivot::SharedItemsHints* out) {
  bool any = false;
  const auto read_flag = [&](const char* attr, bool* dst, bool* has) {
    if (pugi::xml_attribute a = items.attribute(attr); a) {
      bool ok = false;
      *dst = ParseBoolFlag(a.as_string(), &ok);
      *has = true;
      any = true;
    }
  };
  read_flag("containsNumber", &out->contains_number, &out->has_contains_number);
  read_flag("containsInteger", &out->contains_integer, &out->has_contains_integer);
  read_flag("containsDate", &out->contains_date, &out->has_contains_date);
  read_flag("containsString", &out->contains_string, &out->has_contains_string);
  read_flag("containsBlank", &out->contains_blank, &out->has_contains_blank);
  read_flag("containsSemiMixedTypes", &out->contains_semi_mixed, &out->has_contains_semi_mixed);
  read_flag("containsNonDate", &out->contains_non_date, &out->has_contains_non_date);
  read_flag("containsMixedTypes", &out->contains_mixed_types, &out->has_contains_mixed_types);
  const auto read_str = [&](const char* attr, bool* present, std::string* dst) {
    if (pugi::xml_attribute a = items.attribute(attr); a) {
      *present = true;
      dst->assign(a.as_string());
      any = true;
    }
  };
  read_str("minValue", &out->has_min_value, &out->min_value);
  read_str("maxValue", &out->has_max_value, &out->max_value);
  read_str("minDate", &out->has_min_date, &out->min_date);
  read_str("maxDate", &out->has_max_date, &out->max_date);
  out->present = any;
}

}  // namespace

Expected<pivot::PivotCache, Error> read_pivot_cache_definition(const std::vector<std::uint8_t>& definition_bytes) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, definition_bytes, "pivot_cache_reader", "pivotCacheDefinition*.xml"));
  pugi::xml_node root = doc.child("pivotCacheDefinition");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "pivotCacheDefinition*.xml: missing <pivotCacheDefinition> root", "context=pivot_cache_reader");
  }

  pivot::PivotCache cache;

  // Preserve unmodelled root attributes (`refreshedBy`, `refreshedDate`,
  // `refreshOnLoad`, `createdVersion`, ...) for verbatim round-trip. `r:id`
  // and `recordCount` are written from the structured state; namespace
  // declarations are skipped by the capture helper.
  capture_unknown_attrs(root, {"r:id", "recordCount"}, cache.mutable_passthrough_attrs());

  // External cache sources require live data-connection plumbing we do
  // not implement; fail explicitly so the workbook reader can surface
  // a useful message rather than silently producing an empty cache.
  if (pugi::xml_node src = root.child("cacheSource"); src) {
    const std::string_view type = src.attribute("type").as_string();
    if (type == "external") {
      return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                        "pivotCacheDefinition*.xml: external cache source not supported",
                        "context=pivot_cache_reader cacheSource.type=external");
    }
    // type defaults to "worksheet"; any other value (e.g. "consolidation",
    // "scenario") is accepted but treated as worksheet-equivalent here.
    // Capture the `<worksheetSource>` child so Excel's Refresh can locate
    // the source range / defined name after a round trip; dropping it
    // makes Refresh fail or repoint to the wrong range.
    if (pugi::xml_node ws = src.child("worksheetSource"); ws) {
      pivot::WorksheetSource& wsrc = cache.mutable_worksheet_source();
      wsrc.present = true;
      if (pugi::xml_attribute a = ws.attribute("ref"); a) {
        wsrc.ref = a.value();
      }
      if (pugi::xml_attribute a = ws.attribute("sheet"); a) {
        wsrc.sheet = a.value();
      }
      if (pugi::xml_attribute a = ws.attribute("name"); a) {
        wsrc.name = a.value();
      }
    }
  }

  pugi::xml_node fields = root.child("cacheFields");
  if (fields) {
    for (pugi::xml_node f = fields.child("cacheField"); f; f = f.next_sibling("cacheField")) {
      pivot::PivotCacheField field;
      if (pugi::xml_attribute name_attr = f.attribute("name"); name_attr) {
        field.name = name_attr.value();
      }
      // `databaseField` defaults to true; `databaseField="0"` marks a
      // grouping-derived field that carries no per-record cell.
      field.is_database_field = attr_bool(f, "databaseField", true);
      // Capture any `<fieldGroup>` verbatim so a grouped field round-trips
      // even though the grouping structure is not modelled.
      if (pugi::xml_node group = f.child("fieldGroup"); group) {
        append_raw_xml(field.field_group_xml, group);
      }

      // Walk `<sharedItems>` children. Any typed value child marks this
      // as a discrete (shared) field; absence of children leaves
      // `shared_items` empty so the records part is read as inline values.
      if (pugi::xml_node items = f.child("sharedItems"); items) {
        // Capture the numeric / date range + grouping hint attributes so
        // Excel's Refresh keeps its grouping boundaries; dropping them
        // loses grouping hints and can mis-bucket grouped fields.
        ReadSharedItemsHints(items, &field.shared_items_hints);
        for (pugi::xml_node child = items.first_child(); child; child = child.next_sibling()) {
          const std::string_view child_name = child.name();
          if (!IsTypedValueElement(child_name)) {
            continue;
          }
          auto val_or = DecodeTypedValue(child, cache.mutable_text_storage());
          if (!val_or) {
            return val_or.error();
          }
          field.shared_items.push_back(val_or.value());
        }
      }
      cache.mutable_fields().push_back(std::move(field));
    }
  }

  return cache;
}

Expected<void, Error> read_pivot_cache_records(const std::vector<std::uint8_t>& records_bytes,
                                               pivot::PivotCache& cache) {
  pugi::xml_document doc;
  RETURN_IF_ERROR(load_xml_buffer(doc, records_bytes, "pivot_cache_reader", "pivotCacheRecords*.xml"));
  pugi::xml_node root = doc.child("pivotCacheRecords");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "pivotCacheRecords*.xml: missing <pivotCacheRecords> root", "context=pivot_cache_reader");
  }

  const std::size_t field_count = cache.fields().size();
  // Only database fields carry a per-record cell (`<x>` / typed value);
  // grouping-derived fields (`databaseField="0"`) are computed from a base
  // field and have no cell. Map the j-th record child onto the position of
  // the j-th database field so record cells stay index-aligned with
  // `cache.fields()` (group-field slots remain blank).
  std::vector<std::size_t> db_positions;
  for (std::size_t i = 0; i < field_count; ++i) {
    if (cache.fields()[i].is_database_field) {
      db_positions.push_back(i);
    }
  }
  std::vector<pivot::PivotCacheRecord>& records = cache.mutable_records();

  for (pugi::xml_node r = root.child("r"); r; r = r.next_sibling("r")) {
    pivot::PivotCacheRecord record;
    // Pre-fill every field slot with an inline blank; database-field slots
    // are overwritten as the record's children are consumed.
    record.cells.assign(field_count, Value::blank());
    record.cell_is_index.assign(field_count, false);

    std::size_t db_idx = 0;
    for (pugi::xml_node child = r.first_child(); child; child = child.next_sibling()) {
      // Forward compatibility: silently ignore any extra children once we
      // have already populated one cell per database field.
      if (db_idx >= db_positions.size()) {
        break;
      }
      const std::string_view child_name = child.name();
      const std::size_t field_pos = db_positions[db_idx];
      if (child_name == "x") {
        // `<x v="N">` indexes the matching field's shared_items. Missing
        // `v` defaults to 0 (matches OOXML schema default). The cell stores
        // the raw index (as a Number), not the resolved shared value;
        // `pivot::cell_value` performs the shared_items lookup on read.
        const int raw = child.attribute("v").as_int(0);
        if (raw < 0 || static_cast<std::size_t>(raw) >= cache.fields()[field_pos].shared_items.size()) {
          std::string ctx("context=pivot_cache_reader field=");
          ctx.append(std::to_string(field_pos)).append(" index=").append(std::to_string(raw));
          return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivot cache: <x v=...> index out of range",
                            std::move(ctx));
        }
        record.cells[field_pos] = Value::number(static_cast<double>(raw));
        record.cell_is_index[field_pos] = true;
        ++db_idx;
        continue;
      }
      if (IsTypedValueElement(child_name)) {
        auto val_or = DecodeTypedValue(child, cache.mutable_text_storage());
        if (!val_or) {
          return val_or.error();
        }
        record.cells[field_pos] = val_or.value();
        record.cell_is_index[field_pos] = false;
        ++db_idx;
        continue;
      }
      // Unknown element: skip without consuming a field slot. This keeps
      // the record-vs-field alignment honest in the face of forward-
      // compat additions Excel might emit.
    }
    records.push_back(std::move(record));
  }

  return Expected<void, Error>::Ok();
}

}  // namespace formulon::io
