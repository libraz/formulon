#include "io/xlsb/pivot_cache_definition_records.h"

#include <algorithm>
#include <string>
#include <utility>

#include "io/xlsb/frt_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {

namespace pivot_cache_definition {

std::size_t ItemRun::size() const {
  switch (type) {
    case RunType::kNumbers:
      return numbers.size();
    case RunType::kStrings:
      return strings.size();
    case RunType::kDates:
      return dates.size();
  }
  return 0U;
}

std::size_t ItemTable::item_count() const {
  std::size_t count = 0;
  for (const ItemEntry& entry : entries) {
    count += entry.is_run ? entry.run.size() : 1U;
  }
  return count;
}

}  // namespace pivot_cache_definition

namespace {

namespace pcd = pivot_cache_definition;
using pivot_records::kBeginPCD14;
using pivot_records::kBeginPCDCalcItems;
using pivot_records::kBeginPCDFAtbl;
using pivot_records::kBeginPCDFGroup;
using pivot_records::kBeginPCDField;
using pivot_records::kBeginPCDFields;
using pivot_records::kBeginPCDIRun;
using pivot_records::kBeginPCDSource;
using pivot_records::kBeginPCDSRange;
using pivot_records::kBeginPivotCacheDef;
using pivot_records::kBeginPRFItem;
using pivot_records::kDataSourceClean;
using pivot_records::kEndPCD14;
using pivot_records::kEndPCDCalcItems;
using pivot_records::kEndPCDFAtbl;
using pivot_records::kEndPCDFGRange;
using pivot_records::kEndPCDField;
using pivot_records::kEndPCDFields;
using pivot_records::kEndPCDIRun;
using pivot_records::kEndPCDSource;
using pivot_records::kEndPCDSRange;
using pivot_records::kEndPivotCacheDef;
using pivot_records::kEndPNPair;
using pivot_records::kEndPRFItem;
using pivot_records::kPCDIAMissing;
using pivot_records::kPCDIAString;
using pivot_records::kPCDINumber;
using pivot_records::kPCDIString;
using pivot_records::RawRecord;

/// Single shared-item records with no Excel-saved sample: bare missing, boolean, error, lone date.
constexpr std::uint16_t kPCDIMissing = 20;
constexpr std::uint16_t kPCDIBoolean = 22;
constexpr std::uint16_t kPCDIError = 23;
constexpr std::uint16_t kPCDIDatetime = 25;

/// BrtFRTBegin payloads Excel writes around BrtBeginPCD14 and the dataSourceClean record.
constexpr std::uint8_t kPcd14FrtVersion[] = {0x02, 0x0E, 0x00, 0x00};
constexpr std::uint8_t kDataSourceCleanFrtVersion[] = {0x54, 0x00, 0x02, 0x00};
/// BrtBeginPCD14 payload of an `x14:pivotCacheDefinition` with no attributes.
constexpr std::size_t kPcd14Bytes = 9;

/// BrtBeginPivotCacheDef byte 16: refreshedBy and the records relationship id both follow.
constexpr std::uint8_t kCacheStringsPresent = 0x03U;
constexpr std::uint8_t kCacheKnownFlags = pcd::kCacheSaveDataEnableRefresh | pcd::kCacheRefreshOnLoad;

/// BrtBeginPCDSource iSrcType of a worksheet source.
constexpr std::uint32_t kSourceWorksheet = 0;
/// BrtBeginPCDSRange flag bytes 0-2 of a range on a named sheet (fLoadSheet in byte 2).
constexpr std::uint8_t kSRangeSheetFlags[] = {0x00, 0x00, 0x02};

/// BrtBeginPCDField: flag byte, zero byte, numFmtId (u16), 16 zero bytes, then the name.
constexpr std::size_t kFieldReservedBytes = 16;

constexpr std::uint16_t kAtblKnownFlags = pcd::kAtblSemiMixedTypesString | pcd::kAtblNonDate | pcd::kAtblDate |
                                          pcd::kAtblBlank | pcd::kAtblNumber | pcd::kAtblInteger | pcd::kAtblMinMax |
                                          pcd::kAtblAlways;

/// Attribute block of BrtPCDIAString / BrtPCDIAMissing: flag byte then five zero bytes.
constexpr std::size_t kItemAttributeBytes = 6;
constexpr std::uint8_t kItemKnownAttributes = pcd::kItemUnused | pcd::kItemFormula;

constexpr const char* kContext = "context=xlsb_pivot_cache_definition record_id=";

Error Corrupt(const char* message, std::uint16_t id) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, kContext + std::to_string(id));
}

Error Unknown(const char* message, std::uint16_t id) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, message, kContext + std::to_string(id));
}

Error Unmeasured(const char* what, std::uint16_t id) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord,
                    std::string("unmeasured xlsb pivot cache definition layout: ") + what,
                    kContext + std::to_string(id));
}

bool AllZero(const std::uint8_t* p, std::size_t n) {
  return std::all_of(p, p + n, [](std::uint8_t b) { return b == 0U; });
}

bool Equals(ByteSpan span, const std::uint8_t* expected, std::size_t n) {
  return span.size == n && std::equal(expected, expected + n, span.data);
}

RawRecord ToRaw(const XlsbRecord& rec) {
  return RawRecord{rec.type, std::vector<std::uint8_t>(rec.payload.data, rec.payload.data + rec.payload.size)};
}

/// The framed records of the part, walked in order.
class Records {
 public:
  explicit Records(std::vector<XlsbRecord> recs) : recs_(std::move(recs)) {}

  bool done() const { return pos_ == recs_.size(); }
  bool at(std::uint16_t type) const { return !done() && recs_[pos_].type == type; }
  bool at(std::uint16_t type, std::uint16_t next) const {
    return at(type) && pos_ + 1U < recs_.size() && recs_[pos_ + 1U].type == next;
  }
  const XlsbRecord& peek() const { return recs_[pos_]; }
  const XlsbRecord& take() { return recs_[pos_++]; }

  /// The payload of the next record, which must be `type`.
  Expected<ByteSpan, Error> expect(std::uint16_t type) {
    if (done()) {
      return Corrupt("xlsb pivot cache definition ends before an expected record", type);
    }
    if (!at(type)) {
      return Unknown("xlsb pivot cache definition has a record out of place", peek().type);
    }
    return take().payload;
  }

  /// The next record, which must be `type` with an empty payload.
  Expected<void, Error> expect_end(std::uint16_t type) {
    ASSIGN_OR_RETURN(const ByteSpan payload, expect(type));
    if (payload.size != 0U) {
      return Corrupt("xlsb pivot cache definition end record carries a payload", type);
    }
    return Expected<void, Error>::Ok();
  }

 private:
  std::vector<XlsbRecord> recs_;
  std::size_t pos_ = 0;
};

Expected<void, Error> ExpectConsumed(ByteSpan payload, std::uint16_t type) {
  if (payload.size != 0U) {
    return Corrupt("xlsb pivot cache definition record is longer than its layout", type);
  }
  return Expected<void, Error>::Ok();
}

Expected<pcd::Date, Error> ReadDate(ByteSpan& payload) {
  pcd::Date date;
  ASSIGN_OR_RETURN(date.year, read_u16(payload));
  ASSIGN_OR_RETURN(date.month, read_u16(payload));
  ASSIGN_OR_RETURN(date.day, read_u8(payload));
  ASSIGN_OR_RETURN(date.hour, read_u8(payload));
  ASSIGN_OR_RETURN(date.minute, read_u8(payload));
  ASSIGN_OR_RETURN(date.second, read_u8(payload));
  return date;
}

/// Reads the attribute block that ends a BrtPCDIAString / BrtPCDIAMissing.
Expected<std::uint8_t, Error> ReadItemAttributes(ByteSpan payload, std::uint16_t type) {
  if (payload.size != kItemAttributeBytes) {
    return Corrupt("xlsb pivot cache item attribute block has the wrong size", type);
  }
  const std::uint8_t flags = payload.data[0];
  if ((flags & ~kItemKnownAttributes) != 0 || !AllZero(payload.data + 1, 5U)) {
    return Unmeasured("shared item attribute bits", type);
  }
  return flags;
}

// BrtBeginPivotCacheDef: refreshedVersion, minRefreshableVersion, createdVersion (u8 each), flag byte,
// missingItemsLimit (u32), refreshedDate (double), presence byte, recordCount (u32), refreshedBy and the
// records relationship id (XLWideString each).
Expected<void, Error> DecodeHeader(ByteSpan p, pcd::Document* doc) {
  ASSIGN_OR_RETURN(doc->refreshed_version, read_u8(p));
  ASSIGN_OR_RETURN(doc->min_refreshable_version, read_u8(p));
  ASSIGN_OR_RETURN(doc->created_version, read_u8(p));
  ASSIGN_OR_RETURN(doc->flags, read_u8(p));
  if ((doc->flags & ~kCacheKnownFlags) != 0 ||
      (doc->flags & pcd::kCacheSaveDataEnableRefresh) != pcd::kCacheSaveDataEnableRefresh) {
    return Unmeasured("cache definition flag bits", kBeginPivotCacheDef);
  }
  ASSIGN_OR_RETURN(doc->missing_items_limit, read_u32(p));
  ASSIGN_OR_RETURN(doc->refreshed_date, read_double(p));
  ASSIGN_OR_RETURN(const std::uint8_t presence, read_u8(p));
  if (presence != kCacheStringsPresent) {
    return Unmeasured("cache definition without refreshedBy or records relationship", kBeginPivotCacheDef);
  }
  ASSIGN_OR_RETURN(doc->record_count, read_u32(p));
  ASSIGN_OR_RETURN(doc->refreshed_by, read_xlwidestring(p));
  ASSIGN_OR_RETURN(doc->records_rel_id, read_xlwidestring(p));
  return ExpectConsumed(p, kBeginPivotCacheDef);
}

// BrtBeginPCDSource (iSrcType u32, connectionId u32) around BrtBeginPCDSRange (three flag bytes, sheet name,
// RfX range).
Expected<void, Error> DecodeSource(Records& recs, pcd::WorksheetSource* source) {
  ASSIGN_OR_RETURN(ByteSpan p, recs.expect(kBeginPCDSource));
  ASSIGN_OR_RETURN(const std::uint32_t source_type, read_u32(p));
  if (source_type != kSourceWorksheet) {
    return Unmeasured("non-worksheet cache source", kBeginPCDSource);
  }
  ASSIGN_OR_RETURN(source->connection_id, read_u32(p));
  RETURN_IF_ERROR(ExpectConsumed(p, kBeginPCDSource));

  ASSIGN_OR_RETURN(ByteSpan range, recs.expect(kBeginPCDSRange));
  if (range.size < sizeof(kSRangeSheetFlags)) {
    return Corrupt("xlsb pivot cache source range truncated", kBeginPCDSRange);
  }
  if (!std::equal(std::begin(kSRangeSheetFlags), std::end(kSRangeSheetFlags), range.data)) {
    return Unmeasured("cache source range flags (named range, relationship id or no sheet)", kBeginPCDSRange);
  }
  range.data += sizeof(kSRangeSheetFlags);
  range.size -= sizeof(kSRangeSheetFlags);
  ASSIGN_OR_RETURN(source->sheet, read_xlwidestring(range));
  ASSIGN_OR_RETURN(source->ref, read_rfx(range));
  RETURN_IF_ERROR(ExpectConsumed(range, kBeginPCDSRange));
  RETURN_IF_ERROR(recs.expect_end(kEndPCDSRange));
  return recs.expect_end(kEndPCDSource);
}

// BrtBeginPCDIRun: element type (u16), count (u32), then `count` doubles, XLWideStrings or 8-byte dates.
Expected<pcd::ItemRun, Error> DecodeRun(ByteSpan p) {
  pcd::ItemRun run;
  ASSIGN_OR_RETURN(const std::uint16_t type, read_u16(p));
  ASSIGN_OR_RETURN(const std::uint32_t count, read_u32(p));
  // Every element takes at least four bytes; reject a count the payload cannot hold before reserving.
  if (count > p.size / 4U) {
    return Corrupt("xlsb pivot cache item run count exceeds its payload", kBeginPCDIRun);
  }
  switch (type) {
    case static_cast<std::uint16_t>(pcd::RunType::kNumbers):
      run.type = pcd::RunType::kNumbers;
      run.numbers.reserve(count);
      for (std::uint32_t i = 0; i < count; ++i) {
        ASSIGN_OR_RETURN(const double value, read_double(p));
        run.numbers.push_back(value);
      }
      break;
    case static_cast<std::uint16_t>(pcd::RunType::kStrings):
      run.type = pcd::RunType::kStrings;
      run.strings.reserve(count);
      for (std::uint32_t i = 0; i < count; ++i) {
        ASSIGN_OR_RETURN(std::string text, read_xlwidestring(p));
        run.strings.push_back(std::move(text));
      }
      break;
    case static_cast<std::uint16_t>(pcd::RunType::kDates):
      run.type = pcd::RunType::kDates;
      run.dates.reserve(count);
      for (std::uint32_t i = 0; i < count; ++i) {
        ASSIGN_OR_RETURN(const pcd::Date date, ReadDate(p));
        run.dates.push_back(date);
      }
      break;
    default:
      return Unmeasured("shared item run element type", kBeginPCDIRun);
  }
  RETURN_IF_ERROR(ExpectConsumed(p, kBeginPCDIRun));
  return run;
}

Expected<pcd::Item, Error> DecodeItem(const XlsbRecord& rec) {
  pcd::Item item;
  ByteSpan p = rec.payload;
  switch (rec.type) {
    case kPCDIString: {
      item.type = pcd::ItemType::kString;
      ASSIGN_OR_RETURN(item.text, read_xlwidestring(p));
      RETURN_IF_ERROR(ExpectConsumed(p, rec.type));
      break;
    }
    case kPCDINumber: {
      item.type = pcd::ItemType::kNumber;
      ASSIGN_OR_RETURN(item.number, read_double(p));
      RETURN_IF_ERROR(ExpectConsumed(p, rec.type));
      break;
    }
    case kPCDIAString: {
      item.type = pcd::ItemType::kAttributedString;
      ASSIGN_OR_RETURN(item.text, read_xlwidestring(p));
      ASSIGN_OR_RETURN(item.attributes, ReadItemAttributes(p, rec.type));
      break;
    }
    case kPCDIAMissing: {
      item.type = pcd::ItemType::kAttributedMissing;
      ASSIGN_OR_RETURN(item.attributes, ReadItemAttributes(p, rec.type));
      break;
    }
    case kPCDIMissing:
    case kPCDIBoolean:
    case kPCDIError:
    case kPCDIDatetime:
      return Unmeasured("lone shared item record", rec.type);
    default:
      return Unknown("xlsb pivot cache item table carries an unknown record", rec.type);
  }
  return item;
}

// BrtBeginPCDFAtbl: flags (u16), item count (u32), then min and max (double each) when kAtblMinMax is set.
Expected<pcd::ItemTable, Error> DecodeItemTable(Records& recs) {
  pcd::ItemTable table;
  ASSIGN_OR_RETURN(ByteSpan p, recs.expect(kBeginPCDFAtbl));
  ASSIGN_OR_RETURN(table.flags, read_u16(p));
  if ((table.flags & ~kAtblKnownFlags) != 0 || (table.flags & pcd::kAtblAlways) == 0) {
    return Unmeasured("shared items flag bits", kBeginPCDFAtbl);
  }
  ASSIGN_OR_RETURN(const std::uint32_t count, read_u32(p));
  if ((table.flags & pcd::kAtblMinMax) != 0) {
    ASSIGN_OR_RETURN(table.min, read_double(p));
    ASSIGN_OR_RETURN(table.max, read_double(p));
  }
  RETURN_IF_ERROR(ExpectConsumed(p, kBeginPCDFAtbl));

  while (!recs.done() && !recs.at(kEndPCDFAtbl)) {
    const XlsbRecord& rec = recs.take();
    pcd::ItemEntry entry;
    if (rec.type == kBeginPCDIRun) {
      entry.is_run = true;
      ASSIGN_OR_RETURN(entry.run, DecodeRun(rec.payload));
      RETURN_IF_ERROR(recs.expect_end(kEndPCDIRun));
    } else {
      ASSIGN_OR_RETURN(entry.item, DecodeItem(rec));
    }
    table.entries.push_back(std::move(entry));
  }
  RETURN_IF_ERROR(recs.expect_end(kEndPCDFAtbl));
  if (table.item_count() != count) {
    return Corrupt("xlsb pivot cache item count does not match its items", kBeginPCDFAtbl);
  }
  return table;
}

bool IsGroupingRecord(std::uint16_t type) {
  return (type >= kBeginPCDFGroup && type <= kEndPCDFGRange) || type == kBeginPCDIRun || type == kEndPCDIRun;
}

bool IsCalcItemRecord(std::uint16_t type) {
  return (type >= kBeginPCDCalcItems && type <= kEndPNPair) || type == kBeginPRFItem || type == kEndPRFItem;
}

// BrtBeginPCDField: flag byte, zero byte, numFmtId (u16), 16 zero bytes, name (XLWideString).
Expected<pcd::Field, Error> DecodeField(Records& recs) {
  pcd::Field field;
  ASSIGN_OR_RETURN(ByteSpan p, recs.expect(kBeginPCDField));
  ASSIGN_OR_RETURN(field.flags, read_u8(p));
  ASSIGN_OR_RETURN(const std::uint8_t zero, read_u8(p));
  ASSIGN_OR_RETURN(field.num_fmt_id, read_u16(p));
  if (p.size < kFieldReservedBytes) {
    return Corrupt("xlsb pivot cache field header truncated", kBeginPCDField);
  }
  if ((field.flags & ~pcd::kFieldDatabase) != 0 || zero != 0U || !AllZero(p.data, kFieldReservedBytes)) {
    return Unmeasured("cache field header bits (caption, formula, SQL type, server field)", kBeginPCDField);
  }
  p.data += kFieldReservedBytes;
  p.size -= kFieldReservedBytes;
  ASSIGN_OR_RETURN(field.name, read_xlwidestring(p));
  if (p.size != 0U) {
    return Unmeasured("cache field header tail after the name", kBeginPCDField);
  }

  if (recs.at(kBeginPCDFAtbl)) {
    ASSIGN_OR_RETURN(field.items, DecodeItemTable(recs));
  }
  if (recs.at(kBeginPCDFGroup)) {
    while (!recs.done() && !recs.at(kEndPCDField)) {
      if (!IsGroupingRecord(recs.peek().type)) {
        return Unknown("xlsb pivot cache field grouping carries an unknown record", recs.peek().type);
      }
      field.grouping.push_back(ToRaw(recs.take()));
    }
  }
  RETURN_IF_ERROR(recs.expect_end(kEndPCDField));
  return field;
}

// The calculated-item block, from BrtBeginPCDCalcItems through its BrtEndPCDCalcItems.
Expected<void, Error> DecodeCalcItems(Records& recs, std::vector<RawRecord>* out) {
  while (!recs.done()) {
    const std::uint16_t type = recs.peek().type;
    if (!IsCalcItemRecord(type)) {
      return Unknown("xlsb pivot cache calculated items carry an unknown record", type);
    }
    out->push_back(ToRaw(recs.take()));
    if (type == kEndPCDCalcItems) {
      return Expected<void, Error>::Ok();
    }
  }
  return Corrupt("xlsb pivot cache calculated items are not closed", kBeginPCDCalcItems);
}

// BrtFRTBegin (version) + `type` + BrtFRTEnd; returns the wrapped record's payload.
Expected<ByteSpan, Error> DecodeFrt(Records& recs, std::uint16_t type, const std::uint8_t* version,
                                    std::size_t version_size) {
  ASSIGN_OR_RETURN(const ByteSpan begin, recs.expect(kFrtBegin));
  if (!Equals(begin, version, version_size)) {
    return Unmeasured("future-record version", type);
  }
  return recs.expect(type);
}

void EmitDate(std::vector<std::uint8_t>& dst, const pcd::Date& date) {
  emit_u16(dst, date.year);
  emit_u16(dst, date.month);
  emit_u8(dst, date.day);
  emit_u8(dst, date.hour);
  emit_u8(dst, date.minute);
  emit_u8(dst, date.second);
}

void EmitEnd(std::vector<std::uint8_t>& dst, std::uint16_t type) {
  emit_record(dst, type, ByteSpan{});
}

void EmitRaw(std::vector<std::uint8_t>& dst, const std::vector<RawRecord>& recs) {
  for (const RawRecord& rec : recs) {
    emit_record(dst, rec.type, rec.payload);
  }
}

void EmitItemAttributes(std::vector<std::uint8_t>& dst, std::uint8_t attributes) {
  emit_u8(dst, attributes);
  dst.insert(dst.end(), kItemAttributeBytes - 1U, 0U);
}

void EmitItemEntry(std::vector<std::uint8_t>& dst, const pcd::ItemEntry& entry) {
  std::vector<std::uint8_t> p;
  if (entry.is_run) {
    const pcd::ItemRun& run = entry.run;
    emit_u16(p, static_cast<std::uint16_t>(run.type));
    emit_u32(p, static_cast<std::uint32_t>(run.size()));
    switch (run.type) {
      case pcd::RunType::kNumbers:
        for (const double value : run.numbers) {
          emit_double(p, value);
        }
        break;
      case pcd::RunType::kStrings:
        for (const std::string& text : run.strings) {
          emit_xlwidestring(p, text);
        }
        break;
      case pcd::RunType::kDates:
        for (const pcd::Date& date : run.dates) {
          EmitDate(p, date);
        }
        break;
    }
    emit_record(dst, kBeginPCDIRun, p);
    EmitEnd(dst, kEndPCDIRun);
    return;
  }
  const pcd::Item& item = entry.item;
  switch (item.type) {
    case pcd::ItemType::kString:
      emit_xlwidestring(p, item.text);
      emit_record(dst, kPCDIString, p);
      break;
    case pcd::ItemType::kNumber:
      emit_double(p, item.number);
      emit_record(dst, kPCDINumber, p);
      break;
    case pcd::ItemType::kAttributedString:
      emit_xlwidestring(p, item.text);
      EmitItemAttributes(p, item.attributes);
      emit_record(dst, kPCDIAString, p);
      break;
    case pcd::ItemType::kAttributedMissing:
      EmitItemAttributes(p, item.attributes);
      emit_record(dst, kPCDIAMissing, p);
      break;
  }
}

void EmitField(std::vector<std::uint8_t>& dst, const pcd::Field& field) {
  std::vector<std::uint8_t> p;
  emit_u8(p, field.flags);
  emit_u8(p, 0U);
  emit_u16(p, field.num_fmt_id);
  p.insert(p.end(), kFieldReservedBytes, 0U);
  emit_xlwidestring(p, field.name);
  emit_record(dst, kBeginPCDField, p);
  if (field.items.has_value()) {
    const pcd::ItemTable& table = *field.items;
    p.clear();
    emit_u16(p, table.flags);
    emit_u32(p, static_cast<std::uint32_t>(table.item_count()));
    if ((table.flags & pcd::kAtblMinMax) != 0) {
      emit_double(p, table.min);
      emit_double(p, table.max);
    }
    emit_record(dst, kBeginPCDFAtbl, p);
    for (const pcd::ItemEntry& entry : table.entries) {
      EmitItemEntry(dst, entry);
    }
    EmitEnd(dst, kEndPCDFAtbl);
  }
  EmitRaw(dst, field.grouping);
  EmitEnd(dst, kEndPCDField);
}

}  // namespace

Expected<pivot_cache_definition::Document, Error> decode_pivot_cache_definition(ByteSpan bytes) {
  std::vector<XlsbRecord> framed;
  while (bytes.size > 0U) {
    ASSIGN_OR_RETURN(XlsbRecord rec, read_record(bytes));
    framed.push_back(rec);
  }
  Records recs(std::move(framed));
  pcd::Document doc;

  if (recs.at(kAcBegin)) {
    ASSIGN_OR_RETURN(const ByteSpan ac, recs.expect(kAcBegin));
    if (!Equals(ac, kAcUidBeginPayload.data(), kAcUidBeginPayload.size())) {
      return Unmeasured("revision-uid wrapper version", kAcBegin);
    }
    ASSIGN_OR_RETURN(const ByteSpan uid, recs.expect(kUid));
    if (uid.size != kGuidBytes) {
      return Corrupt("xlsb pivot cache uid is not a GUID", kUid);
    }
    std::array<std::uint8_t, kGuidBytes> guid{};
    std::copy(uid.data, uid.data + kGuidBytes, guid.begin());
    doc.uid = guid;
    RETURN_IF_ERROR(recs.expect_end(kAcEnd));
  }

  ASSIGN_OR_RETURN(const ByteSpan header, recs.expect(kBeginPivotCacheDef));
  RETURN_IF_ERROR(DecodeHeader(header, &doc));
  RETURN_IF_ERROR(DecodeSource(recs, &doc.source));

  ASSIGN_OR_RETURN(ByteSpan fields, recs.expect(kBeginPCDFields));
  ASSIGN_OR_RETURN(const std::uint32_t field_count, read_u32(fields));
  RETURN_IF_ERROR(ExpectConsumed(fields, kBeginPCDFields));
  while (!recs.done() && !recs.at(kEndPCDFields)) {
    ASSIGN_OR_RETURN(pcd::Field field, DecodeField(recs));
    doc.fields.push_back(std::move(field));
  }
  RETURN_IF_ERROR(recs.expect_end(kEndPCDFields));
  if (doc.fields.size() != field_count) {
    return Corrupt("xlsb pivot cache field count does not match its fields", kBeginPCDFields);
  }

  if (recs.at(kBeginPCDCalcItems)) {
    RETURN_IF_ERROR(DecodeCalcItems(recs, &doc.calculated_items));
  }
  if (recs.at(kFrtBegin, kBeginPCD14)) {
    ASSIGN_OR_RETURN(const ByteSpan pcd14, DecodeFrt(recs, kBeginPCD14, kPcd14FrtVersion, sizeof(kPcd14FrtVersion)));
    if (pcd14.size != kPcd14Bytes || !AllZero(pcd14.data, kPcd14Bytes)) {
      return Unmeasured("x14 cache definition attributes", kBeginPCD14);
    }
    RETURN_IF_ERROR(recs.expect_end(kEndPCD14));
    RETURN_IF_ERROR(recs.expect_end(kFrtEnd));
    doc.has_pcd14 = true;
  }
  if (recs.at(kFrtBegin, kDataSourceClean)) {
    ASSIGN_OR_RETURN(const ByteSpan clean,
                     DecodeFrt(recs, kDataSourceClean, kDataSourceCleanFrtVersion, sizeof(kDataSourceCleanFrtVersion)));
    if (clean.size != 1U) {
      return Corrupt("xlsb pivot cache dataSourceClean is not one byte", kDataSourceClean);
    }
    doc.data_source_clean = clean.data[0];
    RETURN_IF_ERROR(recs.expect_end(kFrtEnd));
  }
  RETURN_IF_ERROR(recs.expect_end(kEndPivotCacheDef));
  if (!recs.done()) {
    return Unknown("xlsb pivot cache definition has records after its end", recs.peek().type);
  }
  return doc;
}

std::vector<std::uint8_t> encode_pivot_cache_definition(const pivot_cache_definition::Document& doc) {
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> p;
  if (doc.uid.has_value()) {
    emit_record(out, kAcBegin, ByteSpan{kAcUidBeginPayload.data(), kAcUidBeginPayload.size()});
    emit_record(out, kUid, ByteSpan{doc.uid->data(), doc.uid->size()});
    EmitEnd(out, kAcEnd);
  }

  emit_u8(p, doc.refreshed_version);
  emit_u8(p, doc.min_refreshable_version);
  emit_u8(p, doc.created_version);
  emit_u8(p, doc.flags);
  emit_u32(p, doc.missing_items_limit);
  emit_double(p, doc.refreshed_date);
  emit_u8(p, kCacheStringsPresent);
  emit_u32(p, doc.record_count);
  emit_xlwidestring(p, doc.refreshed_by);
  emit_xlwidestring(p, doc.records_rel_id);
  emit_record(out, kBeginPivotCacheDef, p);

  p.clear();
  emit_u32(p, kSourceWorksheet);
  emit_u32(p, doc.source.connection_id);
  emit_record(out, kBeginPCDSource, p);
  p.assign(std::begin(kSRangeSheetFlags), std::end(kSRangeSheetFlags));
  emit_xlwidestring(p, doc.source.sheet);
  emit_rfx(p, doc.source.ref);
  emit_record(out, kBeginPCDSRange, p);
  EmitEnd(out, kEndPCDSRange);
  EmitEnd(out, kEndPCDSource);

  p.clear();
  emit_u32(p, static_cast<std::uint32_t>(doc.fields.size()));
  emit_record(out, kBeginPCDFields, p);
  for (const pcd::Field& field : doc.fields) {
    EmitField(out, field);
  }
  EmitEnd(out, kEndPCDFields);

  EmitRaw(out, doc.calculated_items);
  if (doc.has_pcd14) {
    emit_record(out, kFrtBegin, ByteSpan{kPcd14FrtVersion, sizeof(kPcd14FrtVersion)});
    p.assign(kPcd14Bytes, 0U);
    emit_record(out, kBeginPCD14, p);
    EmitEnd(out, kEndPCD14);
    EmitEnd(out, kFrtEnd);
  }
  if (doc.data_source_clean.has_value()) {
    emit_record(out, kFrtBegin, ByteSpan{kDataSourceCleanFrtVersion, sizeof(kDataSourceCleanFrtVersion)});
    p.assign(1U, *doc.data_source_clean);
    emit_record(out, kDataSourceClean, p);
    EmitEnd(out, kFrtEnd);
  }
  EmitEnd(out, kEndPivotCacheDef);
  return out;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
