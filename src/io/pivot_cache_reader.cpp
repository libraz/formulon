// Copyright 2026 libraz. Licensed under the MIT License.
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

#include "pivot/pivot_cache.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
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
/// `DecodeTypedValue` understands. Used to ignore unknown children
/// (e.g. `<d/>` date placeholders that Excel does not always emit) and
/// to gate the strtod path.
bool IsTypedValueElement(std::string_view name) {
  return name == "s" || name == "n" || name == "b" || name == "m" || name == "e";
}

}  // namespace

Expected<pivot::PivotCache, Error> read_pivot_cache_definition(const std::vector<std::uint8_t>& definition_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(definition_bytes.data(), definition_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=pivot_cache_reader desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "pivotCacheDefinition*.xml: pugixml parse failed",
                      std::move(ctx));
  }
  pugi::xml_node root = doc.child("pivotCacheDefinition");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "pivotCacheDefinition*.xml: missing <pivotCacheDefinition> root", "context=pivot_cache_reader");
  }

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
    // The bytes-preserve path will round-trip the original attribute.
  }

  pivot::PivotCache cache;
  pugi::xml_node fields = root.child("cacheFields");
  if (fields) {
    for (pugi::xml_node f = fields.child("cacheField"); f; f = f.next_sibling("cacheField")) {
      pivot::PivotCacheField field;
      if (pugi::xml_attribute name_attr = f.attribute("name"); name_attr) {
        field.name = name_attr.value();
      }

      // Walk `<sharedItems>` children. Any typed value child marks this
      // as a discrete (shared) field; absence of children leaves
      // `shared_items` empty so the records part is read as inline
      // values. `<fieldGroup>` (date grouping / numeric grouping) is
      // intentionally not parsed here — a follow-up PR will add real
      // semantics; for now group fields land with empty shared_items
      // and their records carry the raw `<x>` indices unchanged.
      if (pugi::xml_node items = f.child("sharedItems"); items) {
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
  pugi::xml_parse_result parse =
      doc.load_buffer(records_bytes.data(), records_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=pivot_cache_reader desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "pivotCacheRecords*.xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("pivotCacheRecords");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid,
                      "pivotCacheRecords*.xml: missing <pivotCacheRecords> root", "context=pivot_cache_reader");
  }

  const std::size_t field_count = cache.fields().size();
  std::vector<pivot::PivotCacheRecord>& records = cache.mutable_records();

  for (pugi::xml_node r = root.child("r"); r; r = r.next_sibling("r")) {
    pivot::PivotCacheRecord record;
    record.cells.reserve(field_count);

    std::size_t field_idx = 0;
    for (pugi::xml_node child = r.first_child(); child; child = child.next_sibling()) {
      // Forward compatibility: silently ignore any extra children once we
      // have already populated one cell per cache field.
      if (field_idx >= field_count) {
        break;
      }
      const std::string_view child_name = child.name();
      if (child_name == "x") {
        // `<x v="N">` indexes the matching field's shared_items. Missing
        // `v` defaults to 0 (matches OOXML schema default).
        const int raw = child.attribute("v").as_int(0);
        if (raw < 0 || static_cast<std::size_t>(raw) >= cache.fields()[field_idx].shared_items.size()) {
          std::string ctx("context=pivot_cache_reader field=");
          ctx.append(std::to_string(field_idx)).append(" index=").append(std::to_string(raw));
          return make_error(FormulonErrorCode::kIoSheetCorrupt, "pivot cache: <x v=...> index out of range",
                            std::move(ctx));
        }
        record.cells.push_back(cache.fields()[field_idx].shared_items[static_cast<std::size_t>(raw)]);
        ++field_idx;
        continue;
      }
      if (IsTypedValueElement(child_name)) {
        auto val_or = DecodeTypedValue(child, cache.mutable_text_storage());
        if (!val_or) {
          return val_or.error();
        }
        record.cells.push_back(val_or.value());
        ++field_idx;
        continue;
      }
      // Unknown element: skip without consuming a field slot. This keeps
      // the record-vs-field alignment honest in the face of forward-
      // compat additions Excel might emit.
    }
    // Trailing blanks: Excel sometimes elides them. Pad up to the
    // declared field count so consumers can index by field index.
    while (record.cells.size() < field_count) {
      record.cells.push_back(Value::blank());
    }
    records.push_back(std::move(record));
  }

  return Expected<void, Error>::Ok();
}

}  // namespace formulon::io
