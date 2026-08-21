#include "io/xlsb/pivot_reader.h"

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "io/xlsb/record.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"
#include "value.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

/// Pivot record ids.
///
/// Established by differential decode against the same workbook re-saved
/// as `.xlsx` (see this module's header), not read off a specification.
/// The ids are what the decoders match on; the names describe the role
/// each record was observed to play.
enum : std::uint16_t {
  // Cache definition part.
  kBeginPCDefinition = 179,
  kEndPCDefinition = 180,
  kBeginPCDFields = 181,
  kEndPCDFields = 182,
  kBeginPCDField = 183,
  kEndPCDField = 184,
  kBeginPCDFAtbl = 189,
  kEndPCDFAtbl = 190,
  kPCDIString = 24,

  // Cache records part.
  kBeginPCRecords = 193,
  kEndPCRecords = 194,
  kPCRecordRow = 33,

  // Pivot table part.
  kBeginPivotTableDef = 280,
  kEndPivotTableDef = 315,
  kBeginPivotFields = 287,
  kEndPivotFields = 288,
  kBeginPivotField = 285,
  kEndPivotField = 286,
  kBeginPivotFieldItems = 283,
  kEndPivotFieldItems = 284,
  kBeginPivotFieldItem = 282,
  kEndPivotFieldItem = 281,
  kPivotRowFields = 309,
  kPivotColFields = 311,
  kBeginPivotDataFields = 295,
  kEndPivotDataFields = 296,
  kBeginPivotDataField = 293,
  kEndPivotDataField = 294,
  kPivotTableLocation = 314,
};

/// `BrtPivotTableLocation`: an `RfX` (first row, last row, first column,
/// last column) followed by the `<location>` offsets. Only the rectangle
/// and the two row offsets are decoded -- the third offset was observed
/// holding an absolute column where the XML form carries one relative to
/// the pivot, and nothing here needs it, so it is left unset rather than
/// converted on an unverified reading.
constexpr std::size_t kLocationMinU32s = 6;

/// Bytes of `BrtBeginPCDField` that precede the field's name.
constexpr std::size_t kPCDFieldNameOffset = 20;
/// Bytes of `BrtBeginPivotTableDef` that precede the table's name.
constexpr std::size_t kPivotTableDefNameOffset = 32;
/// Bytes of `BrtBeginPivotDataField` between the aggregation selector and
/// the data field's display name.
constexpr std::size_t kDataFieldNameGap = 17;

/// A `BrtBeginPivotFieldItem` payload: item kind, two flag bytes, and the
/// cache-item index. The flag bytes carry per-item state (hidden, missing,
/// expanded) whose bit assignments have not been measured, so a non-zero
/// value is refused rather than ignored -- silently dropping a hidden-item
/// filter would over-count every aggregate on that field.
constexpr std::size_t kPivotFieldItemSize = 7;
/// Item kind for an entry that names a cache item; other kinds are the
/// automatic default / subtotal rows, which carry no cache index.
constexpr std::uint8_t kPivotFieldItemKindData = 0;

Error CorruptError(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_reader");
}

Expected<double, Error> ReadDouble(ByteSpan& cursor) {
  if (cursor.size < sizeof(double)) {
    return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb pivot record double truncated",
                      "context=xlsb_pivot_reader");
  }
  double value = 0.0;
  std::memcpy(&value, cursor.data, sizeof(value));
  cursor.data += sizeof(double);
  cursor.size -= sizeof(double);
  return value;
}

/// How a record cell encodes the value for one cache field.
///
/// The record stream is not self-describing: the width and meaning of a
/// cell come from what the definition said about its field. A field with
/// shared items stores a 4-byte index into them; a range-typed numeric
/// field stores the 8-byte value inline.
enum class CellEncoding : std::uint8_t { SharedIndex, Double };

/// Maps an XLSB aggregation selector to the model enum.
///
/// Spelled out rather than cast, even though the two numberings happen to
/// agree today: the model enum is ours to reorder and a silent cast would
/// turn a reordering into wrong pivot totals rather than a build error.
/// Every selector below was measured by re-saving one pivot per
/// `<dataField subtotal>` value and diffing the record.
Expected<pivot::Aggregation, Error> AggregationFromSelector(std::uint32_t selector) {
  switch (selector) {
    case 0U:
      return pivot::Aggregation::Sum;
    case 1U:
      return pivot::Aggregation::Count;
    case 2U:
      return pivot::Aggregation::Average;
    case 3U:
      return pivot::Aggregation::Max;
    case 4U:
      return pivot::Aggregation::Min;
    case 5U:
      return pivot::Aggregation::Product;
    case 6U:
      return pivot::Aggregation::CountNumbers;
    case 7U:
      return pivot::Aggregation::StdDev;
    case 8U:
      return pivot::Aggregation::StdDevP;
    case 9U:
      return pivot::Aggregation::Var;
    case 10U:
      return pivot::Aggregation::VarP;
    default:
      return CorruptError("xlsb pivot data field uses an unknown aggregation selector");
  }
}

/// Reads `count` little-endian u32s that follow a leading count in an
/// axis-field-order record, rejecting a payload that is not exactly the
/// size the count implies.
Expected<std::vector<std::uint32_t>, Error> ReadFieldOrder(ByteSpan payload) {
  auto count_or = read_u32(payload);
  if (!count_or) {
    return count_or.error();
  }
  const std::uint32_t count = count_or.value();
  if (payload.size != static_cast<std::size_t>(count) * sizeof(std::uint32_t)) {
    return CorruptError("xlsb pivot axis field order length does not match its count");
  }
  std::vector<std::uint32_t> order;
  order.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    auto index_or = read_u32(payload);
    if (!index_or) {
      return index_or.error();
    }
    order.push_back(index_or.value());
  }
  return order;
}

/// Decodes the definition part into fields plus the per-field record
/// encoding the records pass needs.
Expected<void, Error> DecodeCacheDefinition(ByteSpan cursor, pivot::PivotCache* cache,
                                            std::vector<CellEncoding>* encodings) {
  bool in_definition = false;
  bool in_field = false;
  bool in_shared_items = false;
  bool field_has_shared_items = false;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    switch (rec.type) {
      case kBeginPCDefinition:
        in_definition = true;
        break;
      case kEndPCDefinition:
        in_definition = false;
        break;
      case kBeginPCDField: {
        if (!in_definition) {
          return CorruptError("xlsb pivot cache field outside a cache definition");
        }
        ByteSpan payload = rec.payload;
        if (payload.size < kPCDFieldNameOffset) {
          return CorruptError("xlsb pivot cache field header truncated");
        }
        payload.data += kPCDFieldNameOffset;
        payload.size -= kPCDFieldNameOffset;
        auto name_or = read_xlwidestring(payload);
        if (!name_or) {
          return name_or.error();
        }
        pivot::PivotCacheField field;
        field.name = std::move(name_or).value();
        cache->mutable_fields().push_back(std::move(field));
        in_field = true;
        field_has_shared_items = false;
        break;
      }
      case kEndPCDField: {
        if (!in_field) {
          return CorruptError("xlsb pivot cache field end without a start");
        }
        // A field with shared items indexes into them; one without stores
        // its values inline, and the only inline width this module has
        // measured is an 8-byte double. The record-width check in
        // `DecodeCacheRecords` is what catches a field whose real inline
        // encoding is something else.
        encodings->push_back(field_has_shared_items ? CellEncoding::SharedIndex : CellEncoding::Double);
        in_field = false;
        break;
      }
      case kBeginPCDFAtbl:
        in_shared_items = true;
        break;
      case kEndPCDFAtbl:
        in_shared_items = false;
        break;
      case kPCDIString: {
        if (!in_shared_items || cache->mutable_fields().empty()) {
          return CorruptError("xlsb pivot cache shared item outside a field's item table");
        }
        ByteSpan payload = rec.payload;
        auto text_or = read_xlwidestring(payload);
        if (!text_or) {
          return text_or.error();
        }
        // `Value::text` aliases its argument rather than owning it, so the
        // bytes have to outlive the record: the cache's `text_storage` is a
        // deque precisely so appends keep earlier entries pointer-stable.
        cache->mutable_text_storage().emplace_back(std::move(text_or.value()));
        cache->mutable_fields().back().shared_items.push_back(Value::text(cache->text_storage().back()));
        field_has_shared_items = true;
        break;
      }
      default:
        // Records outside this module's remit (source range, styling,
        // future-record wrappers) are skipped. A shared item of a type
        // other than string is the one skip that would be unsafe, and it
        // cannot pass unnoticed: it leaves `shared_items` short of what
        // the records index into, which the bounds check in
        // `DecodeCacheRecords` rejects.
        break;
    }
  }
  if (in_definition || in_field || in_shared_items) {
    return CorruptError("xlsb pivot cache definition ended inside an open block");
  }
  if (cache->fields().empty()) {
    return CorruptError("xlsb pivot cache definition declares no fields");
  }
  return Expected<void, Error>::Ok();
}

/// Decodes the records part against the encodings the definition implied.
Expected<void, Error> DecodeCacheRecords(ByteSpan cursor, const std::vector<CellEncoding>& encodings,
                                         pivot::PivotCache* cache) {
  std::size_t expected_width = 0;
  for (const CellEncoding encoding : encodings) {
    expected_width += (encoding == CellEncoding::SharedIndex) ? sizeof(std::uint32_t) : sizeof(double);
  }
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    if (rec.type != kPCRecordRow) {
      // Only the row records carry model state; the surrounding begin /
      // end markers and any future records are skipped.
      continue;
    }
    // The width check is the safety net for the whole encoding
    // inference: a field whose real cell width differs from what the
    // definition implied shifts every later field, and all but a
    // compensating pair of errors changes the total.
    if (rec.payload.size != expected_width) {
      return CorruptError("xlsb pivot cache record width does not match the field encodings");
    }
    ByteSpan payload = rec.payload;
    pivot::PivotCacheRecord record;
    record.cells.reserve(encodings.size());
    record.cell_is_index.reserve(encodings.size());
    for (std::size_t i = 0; i < encodings.size(); ++i) {
      if (encodings[i] == CellEncoding::SharedIndex) {
        auto index_or = read_u32(payload);
        if (!index_or) {
          return index_or.error();
        }
        if (index_or.value() >= cache->fields()[i].shared_items.size()) {
          return CorruptError("xlsb pivot cache record indexes a shared item the definition did not carry");
        }
        record.cells.push_back(Value::number(static_cast<double>(index_or.value())));
        record.cell_is_index.push_back(true);
      } else {
        auto value_or = ReadDouble(payload);
        if (!value_or) {
          return value_or.error();
        }
        record.cells.push_back(Value::number(value_or.value()));
        record.cell_is_index.push_back(false);
      }
    }
    cache->mutable_records().push_back(std::move(record));
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<pivot::PivotCache, Error> read_pivot_cache_bin(ByteSpan definition, ByteSpan records) {
  pivot::PivotCache cache;
  std::vector<CellEncoding> encodings;
  if (auto status = DecodeCacheDefinition(definition, &cache, &encodings); !status) {
    return status.error();
  }
  if (encodings.size() != cache.fields().size()) {
    return CorruptError("xlsb pivot cache definition closed fewer fields than it opened");
  }
  if (auto status = DecodeCacheRecords(records, encodings, &cache); !status) {
    return status.error();
  }
  return cache;
}

Expected<pivot::PivotTable, Error> read_pivot_table_bin(ByteSpan cursor) {
  pivot::PivotTable table;
  std::vector<std::uint32_t> row_order;
  std::vector<std::uint32_t> col_order;
  std::size_t declared_field_count = 0;
  bool in_definition = false;
  bool in_field_items = false;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return rec_or.error();
    }
    const XlsbRecord& rec = rec_or.value();
    switch (rec.type) {
      case kBeginPivotTableDef: {
        ByteSpan payload = rec.payload;
        if (payload.size < kPivotTableDefNameOffset) {
          return CorruptError("xlsb pivot table definition header truncated");
        }
        payload.data += kPivotTableDefNameOffset;
        payload.size -= kPivotTableDefNameOffset;
        auto name_or = read_xlwidestring(payload);
        if (!name_or) {
          return name_or.error();
        }
        table.set_name(std::move(name_or).value());
        auto caption_or = read_xlwidestring(payload);
        if (!caption_or) {
          return caption_or.error();
        }
        table.set_data_caption(std::move(caption_or).value());
        in_definition = true;
        break;
      }
      case kEndPivotTableDef:
        in_definition = false;
        break;
      case kPivotTableLocation: {
        ByteSpan payload = rec.payload;
        if (payload.size < kLocationMinU32s * sizeof(std::uint32_t)) {
          return CorruptError("xlsb pivot table location payload truncated");
        }
        std::uint32_t bounds[kLocationMinU32s] = {};
        for (std::uint32_t& slot : bounds) {
          auto value_or = read_u32(payload);
          if (!value_or) {
            return value_or.error();
          }
          slot = value_or.value();
        }
        if (bounds[1] < bounds[0] || bounds[3] < bounds[2]) {
          return CorruptError("xlsb pivot table location rectangle is inverted");
        }
        table.set_anchor(bounds[0], bounds[2], bounds[1] - bounds[0] + 1U, bounds[3] - bounds[2] + 1U);
        // Excel wrote this extent for the report as it stood, so it is
        // preserved rather than re-projected on write.
        table.mark_span_authored();
        table.set_location_attributes(bounds[4], bounds[5], std::nullopt, std::nullopt, std::nullopt);
        break;
      }
      case kBeginPivotFields: {
        ByteSpan payload = rec.payload;
        auto count_or = read_u32(payload);
        if (!count_or) {
          return count_or.error();
        }
        declared_field_count = count_or.value();
        break;
      }
      case kBeginPivotField: {
        // The field's own header carries axis and data-field flags whose
        // bit assignments have not been measured. They are not read: the
        // row / column / data lists below state the same assignment
        // unambiguously, so there is nothing to gain by inferring it.
        pivot::PivotField field;
        field.axis = pivot::PivotAxis::None;
        table.mutable_fields().push_back(std::move(field));
        break;
      }
      case kBeginPivotFieldItems:
        in_field_items = true;
        break;
      case kEndPivotFieldItems:
        in_field_items = false;
        break;
      case kBeginPivotFieldItem: {
        if (!in_field_items || table.fields().empty()) {
          return CorruptError("xlsb pivot field item outside a field's item list");
        }
        if (rec.payload.size != kPivotFieldItemSize) {
          return CorruptError("xlsb pivot field item payload has an unexpected size");
        }
        ByteSpan payload = rec.payload;
        auto kind_or = read_u8(payload);
        if (!kind_or) {
          return kind_or.error();
        }
        auto flags_or = read_u16(payload);
        if (!flags_or) {
          return flags_or.error();
        }
        if (flags_or.value() != 0U) {
          // Hidden / missing / expanded state lives here. Dropping a
          // hidden-item filter would over-count every aggregate on the
          // field, so an unmeasured flag refuses the table instead.
          return CorruptError("xlsb pivot field item carries flags this reader has not characterised");
        }
        auto index_or = read_u32(payload);
        if (!index_or) {
          return index_or.error();
        }
        if (kind_or.value() != kPivotFieldItemKindData) {
          // The automatic default / subtotal entry; the evaluator
          // synthesises it, so it contributes no modelled item.
          break;
        }
        pivot::PivotItem item;
        item.has_cache_index = true;
        item.cache_index = index_or.value();
        table.mutable_fields().back().items.push_back(std::move(item));
        break;
      }
      case kPivotRowFields: {
        auto order_or = ReadFieldOrder(rec.payload);
        if (!order_or) {
          return order_or.error();
        }
        row_order = std::move(order_or).value();
        break;
      }
      case kPivotColFields: {
        auto order_or = ReadFieldOrder(rec.payload);
        if (!order_or) {
          return order_or.error();
        }
        col_order = std::move(order_or).value();
        break;
      }
      case kBeginPivotDataField: {
        ByteSpan payload = rec.payload;
        auto field_index_or = read_u32(payload);
        if (!field_index_or) {
          return field_index_or.error();
        }
        auto selector_or = read_u32(payload);
        if (!selector_or) {
          return selector_or.error();
        }
        auto aggregation_or = AggregationFromSelector(selector_or.value());
        if (!aggregation_or) {
          return aggregation_or.error();
        }
        if (payload.size < kDataFieldNameGap) {
          return CorruptError("xlsb pivot data field header truncated");
        }
        payload.data += kDataFieldNameGap;
        payload.size -= kDataFieldNameGap;
        auto name_or = read_xlwidestring(payload);
        if (!name_or) {
          return name_or.error();
        }
        pivot::PivotDataField data_field;
        data_field.field_index = field_index_or.value();
        data_field.aggregation = aggregation_or.value();
        data_field.name = std::move(name_or).value();
        table.mutable_data_fields().push_back(std::move(data_field));
        break;
      }
      default:
        break;
    }
  }
  if (in_definition || in_field_items) {
    return CorruptError("xlsb pivot table definition ended inside an open block");
  }
  if (table.fields().size() != declared_field_count) {
    return CorruptError("xlsb pivot table field count does not match the fields decoded");
  }
  if (table.data_fields().empty()) {
    return CorruptError("xlsb pivot table declares no data field");
  }

  const std::size_t field_count = table.fields().size();
  for (const std::uint32_t index : row_order) {
    if (index >= field_count) {
      return CorruptError("xlsb pivot row axis names a field the table does not have");
    }
    table.mutable_fields()[index].axis = pivot::PivotAxis::Row;
  }
  for (const std::uint32_t index : col_order) {
    if (index >= field_count) {
      return CorruptError("xlsb pivot column axis names a field the table does not have");
    }
    table.mutable_fields()[index].axis = pivot::PivotAxis::Col;
  }
  for (const pivot::PivotDataField& data_field : table.data_fields()) {
    if (data_field.field_index >= field_count) {
      return CorruptError("xlsb pivot data field names a field the table does not have");
    }
    table.mutable_fields()[data_field.field_index].axis = pivot::PivotAxis::Value;
  }
  table.mutable_row_field_order() = std::move(row_order);
  table.mutable_col_field_order() = std::move(col_order);
  return table;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
