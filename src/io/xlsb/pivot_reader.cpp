#include "io/xlsb/pivot_reader.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "io/auto_filter_xml.h"
#include "io/pivot_table_reader.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xml_utils.h"
#include "pivot/field_grouping.h"
#include "pivot/pivot_types.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/status_macros.h"
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
  /// A shared string item that carries an attribute block (observed only as
  /// `u="1"`, an item a refresh kept though the source no longer holds it).
  kPCDIStringAttr = 31,
  /// A shared-item array of one type (doubles or strings) and its end marker.
  kPCDIArray = 191,
  kEndPCDIArray = 192,

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

/// Pivot style and filter records, named and laid out per [MS-XLSB]
/// 2.1.7.40 (PIVOTTABLE: ... BrtTableStyleClient [SXFILTERS] ...).
enum : std::uint16_t {
  kBeginSXFilters = 599,  // BrtBeginSXFilters
  kEndSXFilters = 600,    // BrtEndSXFilters
  kBeginSXFilter = 601,   // BrtBeginSXFILTER
  kEndSXFilter = 602,     // BrtEndSXFilter
  kSxFilter15 = 2079,     // BrtSxFilter15
};

/// `ST_PivotFilterType` spellings indexed by `PivotFilterType` (2.5.104),
/// which numbers the same list in the same order.
constexpr const char* kPivotFilterTypes[] = {
    "unknown",
    "count",
    "percent",
    "sum",
    "captionEqual",
    "captionNotEqual",
    "captionBeginsWith",
    "captionNotBeginsWith",
    "captionEndsWith",
    "captionNotEndsWith",
    "captionContains",
    "captionNotContains",
    "captionGreaterThan",
    "captionGreaterThanOrEqual",
    "captionLessThan",
    "captionLessThanOrEqual",
    "captionBetween",
    "captionNotBetween",
    "valueEqual",
    "valueNotEqual",
    "valueGreaterThan",
    "valueGreaterThanOrEqual",
    "valueLessThan",
    "valueLessThanOrEqual",
    "valueBetween",
    "valueNotBetween",
    "dateEqual",
    "dateOlderThan",
    "dateNewerThan",
    "dateBetween",
    "tomorrow",
    "today",
    "yesterday",
    "nextWeek",
    "thisWeek",
    "lastWeek",
    "nextMonth",
    "thisMonth",
    "lastMonth",
    "nextQuarter",
    "thisQuarter",
    "lastQuarter",
    "nextYear",
    "thisYear",
    "lastYear",
    "yearToDate",
    "Q1",
    "Q2",
    "Q3",
    "Q4",
    "M1",
    "M2",
    "M3",
    "M4",
    "M5",
    "M6",
    "M7",
    "M8",
    "M9",
    "M10",
    "M11",
    "M12",
    "dateNotEqual",
    "dateOlderThanOrEqual",
    "dateNewerThanOrEqual",
    "dateNotBetween",
};

constexpr std::uint32_t kNoIndex = 0xFFFFFFFFU;

/// `BrtPivotTableLocation`: an `RfX` (first row, last row, first column,
/// last column) followed by the `<location>` offsets. Only the rectangle
/// and the three first-row / first-data offsets are decoded. The binary form
/// holds them as absolute sheet coordinates where the XML form carries them
/// relative to the pivot's own first row and column (a report anchored at A3
/// stores 2, 2, 0 for the twin's `firstHeaderRow="0" firstDataRow="0"
/// firstDataCol="0"`), so each is taken relative to the rectangle's origin. A
/// payload too short to carry the column offset leaves it unset.
constexpr std::size_t kLocationMinU32s = 6;

/// Bytes of `BrtBeginPCDField` that precede the field's name.
constexpr std::size_t kPCDFieldNameOffset = 20;
/// Bytes of `BrtBeginPivotTableDef` that precede the table's name.
constexpr std::size_t kPivotTableDefNameOffset = 32;
/// Payload byte of `BrtBeginPivotTableDef` whose bit 0 is `useAutoFormatting`.
constexpr std::size_t kPivotTableDefFlagsByte = 5;
/// Bytes of `BrtBeginPivotDataField` between the aggregation selector and
/// the data field's display name.
constexpr std::size_t kDataFieldNameGap = 17;

/// A `BrtBeginPivotFieldItem` payload: item kind, two flag bytes, and the
/// cache-item index. The flag bytes carry per-item state (hidden, missing,
/// expanded) whose bit assignments are not modelled, so a non-zero
/// value is refused rather than ignored -- silently dropping a hidden-item
/// filter would over-count every aggregate on that field.
constexpr std::size_t kPivotFieldItemSize = 7;
/// Item flag bit of a `<item m="1">`: the entry names a cache item the source no longer holds.
constexpr std::uint16_t kPivotFieldItemMissing = 0x0008;
/// Shared-item array element types, and the attribute block of a `kPCDIStringAttr` item
/// (`01 00 00 00 00 00`) the twin `.xlsx` shows as `u="1"`.
constexpr std::uint16_t kArrayOfDoubles = 1;
constexpr std::uint16_t kArrayOfStrings = 2;
constexpr std::uint8_t kUnusedItemAttrs[6] = {1, 0, 0, 0, 0, 0};
/// Item kind for an entry that names a cache item; other kinds are the
/// automatic default / subtotal rows, which carry no cache index.
constexpr std::uint8_t kPivotFieldItemKindData = 0;

Error CorruptError(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, "context=xlsb_pivot_reader");
}

/// How a record cell encodes the value for one cache field.
///
/// The record stream is not self-describing: the width and meaning of a
/// cell come from what the definition said about its field. A field with
/// shared items stores a 4-byte index into them; a range-typed field
/// stores its value inline, as an 8-byte double or an 8-byte date.
enum class CellEncoding : std::uint8_t { SharedIndex, Double, Date };

// BrtBeginPCDFAtbl flags: fNonDates (bit 1) and fDateInField (bit 2). A field with no shared items whose only
// values are dates stores each record cell as an 8-byte date: u16 year, u16 month, u8 day, u8 hour, u8 minute,
// u8 second.
constexpr std::uint16_t kPCDFAtblNonDates = 0x0002U;
constexpr std::uint16_t kPCDFAtblDateInField = 0x0004U;
constexpr std::size_t kInlineDateBytes = 8;

/// The serial of an inline record date, through the same conversion the xlsx reader applies to `<d v>`. Only
/// midnight values are characterised; a time of day is refused rather than guessed.
Expected<double, Error> InlineDateSerial(ByteSpan& payload, bool date1904) {
  ASSIGN_OR_RETURN(const std::uint16_t year, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint16_t month, read_u16(payload));
  ASSIGN_OR_RETURN(const std::uint8_t day, read_u8(payload));
  ASSIGN_OR_RETURN(const std::uint8_t hour, read_u8(payload));
  ASSIGN_OR_RETURN(const std::uint8_t minute, read_u8(payload));
  ASSIGN_OR_RETURN(const std::uint8_t second, read_u8(payload));
  if (hour != 0U || minute != 0U || second != 0U) {
    return CorruptError("xlsb pivot cache record date carries a time of day this reader has not characterised");
  }
  std::string iso;
  const auto append_padded = [&iso](unsigned value, std::size_t width) {
    std::string digits = std::to_string(value);
    iso.append(digits.size() < width ? width - digits.size() : 0U, '0').append(digits);
  };
  append_padded(year, 4U);
  iso.push_back('-');
  append_padded(month, 2U);
  iso.push_back('-');
  append_padded(day, 2U);
  iso.append("T00:00:00");
  const std::optional<double> serial = pivot::group_date_serial(iso, date1904);
  if (!serial.has_value()) {
    return CorruptError("xlsb pivot cache record date is not a calendar date");
  }
  return *serial;
}

/// Maps an XLSB aggregation selector to the model enum.
///
/// Spelled out rather than cast, even though the two numberings happen to
/// agree today: the model enum is ours to reorder and a silent cast would
/// turn a reordering into wrong pivot totals rather than a build error.
/// Every selector below is the record Excel writes for one
/// `<dataField subtotal>` value.
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
/// size the count implies. The Values pseudo-field (`x="-2"` in the `.xlsx`
/// twin, stored as 0xFFFFFFFE) is kept out of the order; its position among
/// the entries goes to `*values_position`, as the `.xlsx` reader records it.
Expected<std::vector<std::uint32_t>, Error> ReadFieldOrder(ByteSpan payload,
                                                           std::optional<std::size_t>* values_position) {
  ASSIGN_OR_RETURN(const std::uint32_t count, read_u32(payload));
  // Divide first: `count * 4` wraps on wasm32.
  if (count > payload.size / sizeof(std::uint32_t) || payload.size != count * sizeof(std::uint32_t)) {
    return CorruptError("xlsb pivot axis field order length does not match its count");
  }
  std::vector<std::uint32_t> order;
  order.reserve(count);
  for (std::uint32_t i = 0; i < count; ++i) {
    ASSIGN_OR_RETURN(auto index, read_u32(payload));
    if (index == static_cast<std::uint32_t>(pivot::kValuesFieldPosition)) {
      if (!values_position->has_value()) {
        *values_position = i;
      }
      continue;
    }
    order.push_back(index);
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
  bool field_is_date_only = false;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    // Inside a cache field only the records the supported layout carries are
    // accepted. Anything else (a grouping block, a non-string shared item)
    // has an unsupported layout and would otherwise be skipped silently.
    if (in_field && rec.type != kEndPCDField && rec.type != kBeginPCDFAtbl && rec.type != kEndPCDFAtbl &&
        rec.type != kPCDIString && rec.type != kPCDIStringAttr && rec.type != kPCDIArray && rec.type != kEndPCDIArray) {
      return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt,
                        "xlsb pivot cache field carries an uncharacterised record (grouping?)",
                        "context=xlsb_pivot_reader record_id=" + std::to_string(rec.type));
    }
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
        ASSIGN_OR_RETURN(auto name, read_xlwidestring(payload));
        pivot::PivotCacheField field;
        field.name = std::move(name);
        cache->mutable_fields().push_back(std::move(field));
        in_field = true;
        field_has_shared_items = false;
        field_is_date_only = false;
        break;
      }
      case kEndPCDField: {
        if (!in_field) {
          return CorruptError("xlsb pivot cache field end without a start");
        }
        // A field with shared items indexes into them; one without stores its values inline: an 8-byte date
        // when the field holds dates only, else an 8-byte double. The record-width check in
        // `DecodeCacheRecords` catches a field whose real inline encoding is something else.
        encodings->push_back(field_has_shared_items ? CellEncoding::SharedIndex
                             : field_is_date_only   ? CellEncoding::Date
                                                    : CellEncoding::Double);
        in_field = false;
        break;
      }
      case kBeginPCDFAtbl: {
        in_shared_items = true;
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(payload));
        field_is_date_only = (flags & (kPCDFAtblNonDates | kPCDFAtblDateInField)) == kPCDFAtblDateInField;
        break;
      }
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
          return std::move(text_or.error());
        }
        // `Value::text` aliases its argument rather than owning it, so the
        // bytes have to outlive the record: the cache's `text_storage` is a
        // deque precisely so appends keep earlier entries pointer-stable.
        cache->mutable_text_storage().emplace_back(std::move(text_or.value()));
        cache->mutable_fields().back().shared_items.push_back(Value::text(cache->text_storage().back()));
        field_has_shared_items = true;
        break;
      }
      case kPCDIStringAttr: {
        if (!in_shared_items || cache->mutable_fields().empty()) {
          return CorruptError("xlsb pivot cache shared item outside a field's item table");
        }
        ByteSpan payload = rec.payload;
        auto text_or = read_xlwidestring(payload);
        if (!text_or) {
          return std::move(text_or.error());
        }
        // Only the attribute block the twin pins is read; any other is refused, not guessed.
        if (payload.size != sizeof(kUnusedItemAttrs) ||
            !std::equal(payload.data, payload.data + payload.size, std::begin(kUnusedItemAttrs))) {
          return CorruptError("xlsb pivot cache shared item carries attributes this reader has not characterised");
        }
        cache->mutable_text_storage().emplace_back(std::move(text_or.value()));
        pivot::PivotCacheField& field = cache->mutable_fields().back();
        field.shared_items.push_back(Value::text(cache->text_storage().back()));
        field.shared_item_is_missing.resize(field.shared_items.size(), false);
        field.shared_item_is_missing.back() = true;
        field_has_shared_items = true;
        break;
      }
      case kPCDIArray: {
        if (!in_shared_items || cache->mutable_fields().empty()) {
          return CorruptError("xlsb pivot cache shared item outside a field's item table");
        }
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(const std::uint16_t element_type, read_u16(payload));
        ASSIGN_OR_RETURN(const std::uint32_t count, read_u32(payload));
        if (element_type != kArrayOfDoubles && element_type != kArrayOfStrings) {
          return CorruptError(
              "xlsb pivot cache shared item array has an element type this reader has not characterised");
        }
        pivot::PivotCacheField& field = cache->mutable_fields().back();
        for (std::uint32_t i = 0; i < count; ++i) {
          if (element_type == kArrayOfDoubles) {
            ASSIGN_OR_RETURN(auto value, read_double(payload));
            field.shared_items.push_back(Value::number(value));
          } else {
            ASSIGN_OR_RETURN(auto text, read_xlwidestring(payload));
            cache->mutable_text_storage().emplace_back(std::move(text));
            field.shared_items.push_back(Value::text(cache->text_storage().back()));
          }
        }
        if (payload.size != 0U) {
          return CorruptError("xlsb pivot cache shared item array length does not match its count");
        }
        field_has_shared_items = true;
        break;
      }
      case kEndPCDIArray:
        break;
      default:
        // Records outside a cache field (source range, styling,
        // future-record wrappers) are outside this module's remit and are
        // skipped; inside a field they were rejected above.
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
Expected<void, Error> DecodeCacheRecords(ByteSpan cursor, const std::vector<CellEncoding>& encodings, bool date1904,
                                         pivot::PivotCache* cache) {
  std::size_t expected_width = 0;
  for (const CellEncoding encoding : encodings) {
    expected_width += encoding == CellEncoding::SharedIndex ? sizeof(std::uint32_t)
                      : encoding == CellEncoding::Date      ? kInlineDateBytes
                                                            : sizeof(double);
  }
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
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
        ASSIGN_OR_RETURN(auto index, read_u32(payload));
        if (index >= cache->fields()[i].shared_items.size()) {
          return CorruptError("xlsb pivot cache record indexes a shared item the definition did not carry");
        }
        record.cells.push_back(Value::number(static_cast<double>(index)));
        record.cell_is_index.push_back(true);
      } else if (encodings[i] == CellEncoding::Date) {
        ASSIGN_OR_RETURN(const double serial, InlineDateSerial(payload, date1904));
        record.cells.push_back(Value::number(serial));
        record.cell_is_index.push_back(false);
        record.cell_is_date.resize(encodings.size(), false);
        record.cell_is_date[i] = true;
      } else {
        ASSIGN_OR_RETURN(auto value, read_double(payload));
        record.cells.push_back(Value::number(value));
        record.cell_is_index.push_back(false);
      }
    }
    cache->mutable_records().push_back(std::move(record));
  }
  return Expected<void, Error>::Ok();
}

/// BrtTableStyleClient as `<pivotTableStyleInfo>`; fFirstColumn does not
/// apply to a pivot.
pivot::PivotStyleInfo DecodeStyleClient(ByteSpan payload) {
  pivot::PivotStyleInfo info;
  auto flags = read_u16(payload);
  auto name = read_xlnullablewidestring(payload);
  if (flags) {
    info.show_last_column = (flags.value() & kStyleClientLastColumn) != 0U;
    info.show_row_stripes = (flags.value() & kStyleClientRowStripes) != 0U;
    info.show_col_stripes = (flags.value() & kStyleClientColumnStripes) != 0U;
    info.show_row_headers = (flags.value() & kStyleClientRowHeaders) != 0U;
    info.show_col_headers = (flags.value() & kStyleClientColumnHeaders) != 0U;
  }
  if (name) {
    info.name = std::move(name.value());
  }
  return info;
}

/// Spells one SXFILTER (BrtBeginSXFILTER, payload `begin`, just read) as its
/// OOXML `<filter>` element, consuming through BrtEndSXFilter.
Expected<void, Error> AppendFilterXml(ByteSpan& cursor, ByteSpan begin, std::string& out) {
  ASSIGN_OR_RETURN(const std::uint32_t fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t mp_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t type, read_u32(begin));
  // Documented as unused; Excel stores the filter's evalOrder there.
  ASSIGN_OR_RETURN(const std::uint32_t eval_order, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t id, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_hier, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(begin));
  if (type == 0U || type >= std::size(kPivotFilterTypes) || fld >= 0xFFFFFFFEU) {
    return CorruptError("xlsb pivot filter has an unknown type or field");
  }
  out.append("<filter");
  append_xml_attr_uint(out, "fld", fld);
  append_xml_attr_uint_if(out, "mpFld", mp_fld, kNoIndex);
  append_xml_attr(out, "type", kPivotFilterTypes[type]);
  append_xml_attr_int(out, "evalOrder", static_cast<std::int32_t>(eval_order));
  append_xml_attr_uint(out, "id", id);
  append_xml_attr_uint_if(out, "iMeasureHier", measure_hier, kNoIndex);
  append_xml_attr_uint_if(out, "iMeasureFld", measure_fld, kNoIndex);
  const char* const string_attrs[] = {"name", "description", "stringValue1", "stringValue2"};
  for (std::uint32_t bit = 0; bit < 4U; ++bit) {
    if ((flags & (1U << bit)) != 0U) {
      ASSIGN_OR_RETURN(auto text, read_xlwidestring(begin));
      append_xml_attr(out, string_attrs[bit], text);
    }
  }
  out.push_back('>');
  ASSIGN_OR_RETURN(const XlsbRecord afilter, read_record(cursor));
  if (afilter.type != kBeginAFilter) {
    return CorruptError("xlsb pivot filter carries no AutoFilter");
  }
  ASSIGN_OR_RETURN(const AutoFilter criteria, read_auto_filter_block(cursor, afilter.payload));
  out.append(serialize_auto_filter(criteria));
  for (;;) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == kEndSXFilter) {
      break;
    }
    // FRTSXFILTER: only a BrtSxFilter15 with fWholeDay clear says nothing
    // the OOXML element would not.
    if (rec.type != kFrtBegin) {
      return CorruptError("xlsb pivot filter carries an uncharacterised record");
    }
    ASSIGN_OR_RETURN(const XlsbRecord inner, read_record(cursor));
    ASSIGN_OR_RETURN(const XlsbRecord end, read_record(cursor));
    const bool whole_day = std::any_of(inner.payload.data, inner.payload.data + inner.payload.size,
                                       [](std::uint8_t b) { return b != 0U; });
    if (inner.type != kSxFilter15 || end.type != kFrtEnd || whole_day) {
      return CorruptError("xlsb pivot filter carries a future record this reader does not model");
    }
  }
  out.append("</filter>");
  return Expected<void, Error>::Ok();
}

/// Reads SXFILTERS after its BrtBeginSXFilters (payload `begin`) into
/// `out`'s `<filter>` list.
Expected<void, Error> AppendFiltersXml(ByteSpan& cursor, ByteSpan begin, std::string& out) {
  ASSIGN_OR_RETURN(const std::uint32_t declared, read_u32(begin));
  for (std::uint32_t seen = 0;; ++seen) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == kEndSXFilters) {
      if (seen != declared) {
        return CorruptError("xlsb pivot filter count does not match the filters decoded");
      }
      return Expected<void, Error>::Ok();
    }
    if (rec.type != kBeginSXFilter) {
      return CorruptError("xlsb pivot filter list carries an uncharacterised record");
    }
    RETURN_IF_ERROR(AppendFilterXml(cursor, rec.payload, out));
  }
}

}  // namespace

Expected<pivot::PivotCache, Error> read_pivot_cache_bin(ByteSpan definition, ByteSpan records, bool date1904) {
  pivot::PivotCache cache;
  std::vector<CellEncoding> encodings;
  if (auto status = DecodeCacheDefinition(definition, &cache, &encodings); !status) {
    return std::move(status.error());
  }
  if (encodings.size() != cache.fields().size()) {
    return CorruptError("xlsb pivot cache definition closed fewer fields than it opened");
  }
  if (auto status = DecodeCacheRecords(records, encodings, date1904, &cache); !status) {
    return std::move(status.error());
  }
  return cache;
}

Expected<pivot::PivotTable, Error> read_pivot_table_bin(ByteSpan cursor) {
  pivot::PivotTable table;
  table.set_definition_source(pivot::PivotDefinitionSource::kXlsb);
  std::vector<std::uint32_t> row_order;
  std::vector<std::uint32_t> col_order;
  std::size_t declared_field_count = 0;
  bool in_definition = false;
  bool in_field_items = false;
  std::string filters_xml;
  while (cursor.size > 0) {
    auto rec_or = read_record(cursor);
    if (!rec_or) {
      return std::move(rec_or.error());
    }
    const XlsbRecord& rec = rec_or.value();
    switch (rec.type) {
      case kBeginPivotTableDef: {
        ByteSpan payload = rec.payload;
        if (payload.size < kPivotTableDefNameOffset) {
          return CorruptError("xlsb pivot table definition header truncated");
        }
        table.set_auto_formatting((payload.data[kPivotTableDefFlagsByte] & 0x01U) != 0U);
        payload.data += kPivotTableDefNameOffset;
        payload.size -= kPivotTableDefNameOffset;
        ASSIGN_OR_RETURN(auto name, read_xlwidestring(payload));
        table.set_name(std::move(name));
        ASSIGN_OR_RETURN(auto caption, read_xlwidestring(payload));
        table.set_data_caption(std::move(caption));
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
          ASSIGN_OR_RETURN(auto value, read_u32(payload));
          slot = value;
        }
        if (bounds[1] < bounds[0] || bounds[3] < bounds[2]) {
          return CorruptError("xlsb pivot table location rectangle is inverted");
        }
        if (bounds[1] >= Sheet::kMaxRows || bounds[3] >= Sheet::kMaxCols) {
          return CorruptError("xlsb pivot table location rectangle exceeds sheet grid bounds");
        }
        table.set_anchor(bounds[0], bounds[2], bounds[1] - bounds[0] + 1U, bounds[3] - bounds[2] + 1U);
        // Excel wrote this extent for the report as it stood, so it is
        // preserved rather than re-projected on write.
        table.mark_span_authored();
        const auto relative = [](std::uint32_t absolute, std::uint32_t origin) {
          return absolute >= origin ? absolute - origin : absolute;
        };
        std::optional<std::uint32_t> first_data_col;
        if (payload.size >= sizeof(std::uint32_t)) {
          ASSIGN_OR_RETURN(auto value, read_u32(payload));
          first_data_col = relative(value, bounds[2]);
        }
        table.set_location_attributes(relative(bounds[4], bounds[0]), relative(bounds[5], bounds[0]), first_data_col,
                                      std::nullopt, std::nullopt);
        break;
      }
      case kBeginPivotFields: {
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(auto count, read_u32(payload));
        declared_field_count = count;
        break;
      }
      case kBeginPivotField: {
        // The field's own header carries axis and data-field flags whose
        // bit assignments are not modelled. They are not read: the
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
        ASSIGN_OR_RETURN(auto kind, read_u8(payload));
        auto flags_or = read_u16(payload);
        if (!flags_or) {
          return std::move(flags_or.error());
        }
        if ((flags_or.value() & ~kPivotFieldItemMissing) != 0U) {
          // Hidden / expanded state lives here. Dropping a hidden-item
          // filter would over-count every aggregate on the field, so an
          // unknown flag refuses the table instead.
          return CorruptError("xlsb pivot field item carries flags this reader has not characterised");
        }
        ASSIGN_OR_RETURN(auto index, read_u32(payload));
        if (kind != kPivotFieldItemKindData) {
          // The automatic default / subtotal entry; the evaluator
          // synthesises it, so it contributes no modelled item.
          break;
        }
        pivot::PivotItem item;
        item.has_cache_index = true;
        item.cache_index = index;
        if ((flags_or.value() & kPivotFieldItemMissing) != 0U) {
          item.passthrough_attrs.emplace_back("m", "1");
        }
        table.mutable_fields().back().items.push_back(std::move(item));
        break;
      }
      case kPivotRowFields: {
        std::optional<std::size_t> values_position;
        ASSIGN_OR_RETURN(auto order, ReadFieldOrder(rec.payload, &values_position));
        row_order = std::move(order);
        table.set_row_values_position(values_position);
        break;
      }
      case kPivotColFields: {
        std::optional<std::size_t> values_position;
        ASSIGN_OR_RETURN(auto order, ReadFieldOrder(rec.payload, &values_position));
        col_order = std::move(order);
        table.set_col_values_position(values_position);
        break;
      }
      case kBeginPivotDataField: {
        ByteSpan payload = rec.payload;
        ASSIGN_OR_RETURN(auto field_index, read_u32(payload));
        ASSIGN_OR_RETURN(auto selector, read_u32(payload));
        ASSIGN_OR_RETURN(auto aggregation, AggregationFromSelector(selector));
        if (payload.size < kDataFieldNameGap) {
          return CorruptError("xlsb pivot data field header truncated");
        }
        payload.data += kDataFieldNameGap;
        payload.size -= kDataFieldNameGap;
        ASSIGN_OR_RETURN(auto name, read_xlwidestring(payload));
        pivot::PivotDataField data_field;
        data_field.field_index = field_index;
        data_field.aggregation = aggregation;
        data_field.name = std::move(name);
        table.mutable_data_fields().push_back(std::move(data_field));
        break;
      }
      case kTableStyleClient:
        table.set_style_info(DecodeStyleClient(rec.payload));
        break;
      case kBeginSXFilters:
        if (auto status = AppendFiltersXml(cursor, rec.payload, filters_xml); !status) {
          return std::move(status.error());
        }
        break;
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
  // The filters decode against the field axes, exactly as `<filters>` does.
  if (!filters_xml.empty() && !read_pivot_filters_xml("<filters>" + filters_xml + "</filters>", &table)) {
    return CorruptError("xlsb pivot filters do not spell as OOXML");
  }
  return table;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
