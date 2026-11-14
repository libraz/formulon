#include "io/xlsb/pivot_cache_records.h"

#include <cmath>
#include <cstddef>
#include <cstring>
#include <limits>
#include <string>
#include <utility>

#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "utils/date_time.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

using cache_records::CellKind;
using cache_records::PivotCacheRecordsDoc;
using cache_records::RecordCell;
using cache_records::RecordDate;

constexpr const char* kContext = "context=xlsb_pivot_cache_records";

/// BrtBeginPCDField byte 0: set on a field the source range supplies (`databaseField`, default 1); clear on a
/// grouping-only field, which has no record cell.
constexpr std::uint8_t kPCDFieldDatabase = 0x04U;

/// BrtBeginPCDFAtbl flags, by the `sharedItems` attribute each mirrors. With no shared items, a field holding
/// only dates has fDateInField alone, only text fHasText, only numbers fNumField (fIntField too when integral);
/// fNonDates rides on the text and number shapes. Blank and mixed-type inline cells are refused.
constexpr std::uint16_t kAtblNonDates = 0x0002U;
constexpr std::uint16_t kAtblDateInField = 0x0004U;
constexpr std::uint16_t kAtblHasText = 0x0008U;
constexpr std::uint16_t kAtblHasBlank = 0x0010U;
constexpr std::uint16_t kAtblMixedTypes = 0x0020U;
constexpr std::uint16_t kAtblNumField = 0x0040U;
/// u16 flags then u32 citems, the part of BrtBeginPCDFAtbl the kind needs.
constexpr std::size_t kAtblHeadBytes = 6;

constexpr std::size_t kCountBytes = 4;
constexpr std::size_t kDateBytes = 8;

/// 1900-03-01, the first 1900-system serial past the phantom 1900-02-29.
constexpr double kFirstSerialAfterPhantomDay = 61.0;

Error Corrupt(const char* message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, message, kContext);
}

bool IsCalendarDate(const RecordDate& date) {
  return date.month >= 1U && date.month <= 12U && date.day >= 1U &&
         date.day <= date_time::days_in_month(static_cast<int>(date.year), date.month) && date.hour < 24U &&
         date.minute < 60U && date.second < 60U;
}

Expected<CellKind, Error> KindOfAtbl(ByteSpan payload) {
  if (payload.size < kAtblHeadBytes) {
    return Corrupt("xlsb pivot cache field item table header is truncated");
  }
  const std::uint16_t flags = load_u16(payload.data);
  const std::uint32_t item_count = load_u32(payload.data + 2);
  if (item_count > 0U) {
    return CellKind::kSharedIndex;
  }
  const auto shape = static_cast<std::uint16_t>(
      flags & (kAtblDateInField | kAtblHasText | kAtblHasBlank | kAtblMixedTypes | kAtblNumField));
  if (shape == kAtblDateInField && (flags & kAtblNonDates) == 0U) {
    return CellKind::kDate;
  }
  if (shape == kAtblHasText) {
    return CellKind::kText;
  }
  if (shape == kAtblNumField) {
    return CellKind::kNumber;
  }
  return Corrupt("xlsb pivot cache field without shared items has an inline value shape that is not characterised");
}

Expected<RecordCell, Error> DecodeCell(ByteSpan& cursor, CellKind kind) {
  RecordCell cell;
  cell.kind = kind;
  switch (kind) {
    case CellKind::kSharedIndex: {
      ASSIGN_OR_RETURN(cell.index, read_u32(cursor));
      break;
    }
    case CellKind::kNumber: {
      ASSIGN_OR_RETURN(cell.number, read_double(cursor));
      break;
    }
    case CellKind::kText: {
      const std::uint8_t* start = cursor.data;
      ASSIGN_OR_RETURN(cell.text, read_xlwidestring(cursor));
      // A string the encoder would spell differently (a lone surrogate) cannot round trip.
      std::vector<std::uint8_t> respelt;
      emit_xlwidestring(respelt, cell.text);
      if (respelt.size() != static_cast<std::size_t>(cursor.data - start) ||
          std::memcmp(respelt.data(), start, respelt.size()) != 0) {
        return Corrupt("xlsb pivot cache record text does not round trip through UTF-8");
      }
      break;
    }
    case CellKind::kDate: {
      if (cursor.size < kDateBytes) {
        return make_error(FormulonErrorCode::kIoXlsbRecordTruncated, "xlsb pivot cache record date is truncated",
                          kContext);
      }
      cell.date.year = load_u16(cursor.data);
      cell.date.month = load_u16(cursor.data + 2);
      cell.date.day = cursor.data[4];
      cell.date.hour = cursor.data[5];
      cell.date.minute = cursor.data[6];
      cell.date.second = cursor.data[7];
      cursor.data += kDateBytes;
      cursor.size -= kDateBytes;
      if (!IsCalendarDate(cell.date)) {
        return Corrupt("xlsb pivot cache record date is not a calendar date");
      }
      break;
    }
  }
  return cell;
}

Expected<void, Error> EncodeCell(std::vector<std::uint8_t>& dst, const RecordCell& cell) {
  switch (cell.kind) {
    case CellKind::kSharedIndex:
      emit_u32(dst, cell.index);
      break;
    case CellKind::kNumber:
      emit_double(dst, cell.number);
      break;
    case CellKind::kText:
      emit_xlwidestring(dst, cell.text);
      break;
    case CellKind::kDate:
      if (!IsCalendarDate(cell.date)) {
        return refuse_write("XLSB pivot cache record date is not a calendar date");
      }
      emit_u16(dst, cell.date.year);
      emit_u16(dst, cell.date.month);
      emit_u8(dst, cell.date.day);
      emit_u8(dst, cell.date.hour);
      emit_u8(dst, cell.date.minute);
      emit_u8(dst, cell.date.second);
      break;
  }
  return Expected<void, Error>::Ok();
}

}  // namespace

Expected<cache_records::PivotCacheRecordsDoc, Error> decode_pivot_cache_records(
    ByteSpan part, const std::vector<cache_records::CellKind>& kinds) {
  if (kinds.empty()) {
    return Corrupt("xlsb pivot cache records decoded against no database fields");
  }
  ByteSpan cursor = part;
  ASSIGN_OR_RETURN(const XlsbRecord begin, read_record(cursor));
  if (begin.type != pivot_records::kBeginPivotCacheRecords || begin.payload.size != kCountBytes) {
    return Corrupt("xlsb pivot cache records part does not open with its record count");
  }
  const std::uint32_t count = load_u32(begin.payload.data);
  PivotCacheRecordsDoc doc;
  bool ended = false;
  while (cursor.size > 0U) {
    if (ended) {
      return Corrupt("xlsb pivot cache records part carries records after its end");
    }
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == pivot_records::kEndPivotCacheRecords) {
      if (rec.payload.size != 0U) {
        return Corrupt("xlsb pivot cache records end record carries a payload");
      }
      ended = true;
      continue;
    }
    if (rec.type != pivot_records::kPCRRecord) {
      return make_error(FormulonErrorCode::kIoXlsbUnknownRecord,
                        "xlsb pivot cache records part carries an unknown record",
                        std::string(kContext) + " record_id=" + std::to_string(rec.type));
    }
    ByteSpan row_bytes = rec.payload;
    std::vector<RecordCell> row;
    row.reserve(kinds.size());
    for (const CellKind kind : kinds) {
      auto cell = DecodeCell(row_bytes, kind);
      if (!cell) {
        if (cell.error().code == FormulonErrorCode::kIoXlsbRecordTruncated) {
          return Corrupt("xlsb pivot cache record is narrower than its field kinds");
        }
        return std::move(cell.error());
      }
      row.push_back(std::move(cell.value()));
    }
    if (row_bytes.size != 0U) {
      return Corrupt("xlsb pivot cache record is wider than its field kinds");
    }
    doc.rows.push_back(std::move(row));
  }
  if (!ended) {
    return Corrupt("xlsb pivot cache records part has no end record");
  }
  if (doc.rows.size() != count) {
    return Corrupt("xlsb pivot cache record count does not match its rows");
  }
  return doc;
}

Expected<std::vector<std::uint8_t>, Error> encode_pivot_cache_records(
    const cache_records::PivotCacheRecordsDoc& doc, const std::vector<cache_records::CellKind>& kinds) {
  if (static_cast<std::uint64_t>(doc.rows.size()) > std::numeric_limits<std::uint32_t>::max()) {
    return refuse_write("XLSB pivot cache holds more records than its count can name");
  }
  std::vector<std::uint8_t> out;
  std::vector<std::uint8_t> payload;
  emit_u32(payload, static_cast<std::uint32_t>(doc.rows.size()));
  emit_record(out, pivot_records::kBeginPivotCacheRecords, payload);
  for (const std::vector<RecordCell>& row : doc.rows) {
    if (row.size() != kinds.size()) {
      return refuse_write("XLSB pivot cache record does not hold one cell per database field");
    }
    payload.clear();
    for (std::size_t i = 0; i < row.size(); ++i) {
      if (row[i].kind != kinds[i]) {
        return refuse_write("XLSB pivot cache record cell kind does not match its field");
      }
      RETURN_IF_ERROR(EncodeCell(payload, row[i]));
    }
    emit_record(out, pivot_records::kPCRRecord, payload);
  }
  payload.clear();
  emit_record(out, pivot_records::kEndPivotCacheRecords, payload);
  return out;
}

Expected<std::vector<cache_records::CellKind>, Error> cell_kinds_of_cache_definition(ByteSpan definition_part) {
  std::vector<CellKind> kinds;
  ByteSpan cursor = definition_part;
  bool in_field = false;
  bool database_field = false;
  bool atbl_seen = false;
  CellKind field_kind = CellKind::kSharedIndex;
  while (cursor.size > 0U) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    switch (rec.type) {
      case pivot_records::kBeginPCDField:
        if (in_field) {
          return Corrupt("xlsb pivot cache field opens inside another field");
        }
        if (rec.payload.size == 0U) {
          return Corrupt("xlsb pivot cache field header is truncated");
        }
        in_field = true;
        database_field = (rec.payload.data[0] & kPCDFieldDatabase) != 0U;
        atbl_seen = false;
        break;
      case pivot_records::kBeginPCDFAtbl: {
        if (!in_field || atbl_seen) {
          return Corrupt("xlsb pivot cache item table outside a field or repeated within one");
        }
        atbl_seen = true;
        if (database_field) {
          ASSIGN_OR_RETURN(field_kind, KindOfAtbl(rec.payload));
        }
        break;
      }
      case pivot_records::kEndPCDField:
        if (!in_field) {
          return Corrupt("xlsb pivot cache field end without a start");
        }
        if (database_field) {
          if (!atbl_seen) {
            return Corrupt("xlsb pivot cache database field has no item table");
          }
          kinds.push_back(field_kind);
        }
        in_field = false;
        break;
      default:
        // Everything else describes the cache, not the record layout.
        break;
    }
  }
  if (in_field) {
    return Corrupt("xlsb pivot cache definition ends inside a field");
  }
  if (kinds.empty()) {
    return Corrupt("xlsb pivot cache definition declares no database fields");
  }
  return kinds;
}

Expected<double, Error> record_date_serial(const cache_records::RecordDate& date, bool date1904) {
  if (!IsCalendarDate(date)) {
    return Corrupt("xlsb pivot cache record date is not a calendar date");
  }
  const double day = date_time::serial_from_ymd(static_cast<int>(date.year), date.month, date.day, false);
  const double gap = date1904 ? date_time::kDate1904EpochGap : 0.0;
  // Days before 1900-03-01 sit across the phantom day, where the inline form is not characterised.
  if (day < kFirstSerialAfterPhantomDay || day - gap < 0.0 || day > date_time::max_date_serial(false)) {
    return Corrupt("xlsb pivot cache record date lies outside the characterised serial range");
  }
  // The day plus the time as one 1900-system serial, then the 1904 shift: the xlsx `<d v>` arithmetic.
  const double seconds = static_cast<double>(date.hour) * 3600.0 + static_cast<double>(date.minute) * 60.0 +
                         static_cast<double>(date.second);
  return (day + seconds / static_cast<double>(date_time::kSecondsPerDay)) - gap;
}

Expected<cache_records::RecordDate, Error> record_date_from_serial(double serial, bool date1904) {
  if (!std::isfinite(serial)) {
    return refuse_write("XLSB pivot cache record date is not finite");
  }
  const double serial1900 = serial + (date1904 ? date_time::kDate1904EpochGap : 0.0);
  double day = std::floor(serial1900);
  long long seconds = std::llround((serial1900 - day) * static_cast<double>(date_time::kSecondsPerDay));
  if (seconds >= date_time::kSecondsPerDay) {
    day += 1.0;
    seconds -= date_time::kSecondsPerDay;
  }
  if (day < kFirstSerialAfterPhantomDay || day > date_time::max_date_serial(false)) {
    return refuse_write("XLSB pivot cache record date lies outside the characterised serial range");
  }
  const date_time::YMD ymd = date_time::ymd_from_serial(day, false);
  RecordDate date;
  date.year = static_cast<std::uint16_t>(ymd.y);
  date.month = static_cast<std::uint16_t>(ymd.m);
  date.day = static_cast<std::uint8_t>(ymd.d);
  date.hour = static_cast<std::uint8_t>(seconds / 3600);
  date.minute = static_cast<std::uint8_t>(seconds / 60 % 60);
  date.second = static_cast<std::uint8_t>(seconds % 60);
  // A fraction of a second, or a serial the decode side would not reproduce, has no cell spelling.
  auto back = record_date_serial(date, date1904);
  if (!back || back.value() != serial) {
    return refuse_write("XLSB pivot cache record date is not a whole second");
  }
  return date;
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
