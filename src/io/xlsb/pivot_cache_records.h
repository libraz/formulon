//
// The pivot cache records part (xl/pivotCache/pivotCacheRecordsN.bin),
// [MS-XLSB] PIVOTCACHERECORDS: BrtBeginPivotCacheRecords (193, u32 record
// count), one BrtPCRRecord (33) per source row, BrtEndPivotCacheRecords (194).
// A row is not self-describing: each database field of the cache definition
// fixes its cell's layout, so decode and encode take the per-field kinds that
// `cell_kinds_of_cache_definition` derives from the definition part.

#ifndef FORMULON_IO_XLSB_PIVOT_CACHE_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_CACHE_RECORDS_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/xlsb/pivot_records_common.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace cache_records {

/// Layout of one record cell, fixed per database field by the definition.
enum class CellKind : std::uint8_t {
  kSharedIndex,  ///< u32 index into the field's shared items (`<x>`).
  kNumber,       ///< IEEE 754 double (`<n>`).
  kText,         ///< XLWideString (`<s>`).
  kDate,         ///< u16 year, u16 month, u8 day, u8 hour, u8 minute, u8 second (`<d>`).
};

/// The calendar day of a date cell. Only midnight is accepted: a time of day
/// is refused on both decode and encode.
struct RecordDate {
  std::uint16_t year = 0;
  std::uint16_t month = 0;
  std::uint8_t day = 0;
};

/// One record cell; only the member matching `kind` is meaningful.
struct RecordCell {
  CellKind kind = CellKind::kSharedIndex;
  std::uint32_t index = 0;
  double number = 0.0;
  std::string text;
  RecordDate date;
};

/// The whole part. The record count of the begin record is `rows.size()`.
struct PivotCacheRecordsDoc {
  std::vector<std::vector<RecordCell>> rows;
};

}  // namespace cache_records

/// Decodes a records part against `kinds` (one entry per database field).
/// Fails on any record other than 193 / 33 / 194 in that order, a count that
/// disagrees with the rows, a row whose width `kinds` does not explain
/// exactly, a non-calendar date or a time of day.
Expected<cache_records::PivotCacheRecordsDoc, Error> decode_pivot_cache_records(
    ByteSpan part, const std::vector<cache_records::CellKind>& kinds);

/// Encodes `doc`; every row must hold exactly one cell of the matching kind
/// per entry of `kinds`. Inverse of `decode_pivot_cache_records` byte for byte.
Expected<std::vector<std::uint8_t>, Error> encode_pivot_cache_records(
    const cache_records::PivotCacheRecordsDoc& doc, const std::vector<cache_records::CellKind>& kinds);

/// Derives the record cell kinds from a cache definition part: one entry per
/// database field (BrtBeginPCDField byte 0 bit 0x04), in field order. A field
/// whose BrtBeginPCDFAtbl declares shared items indexes them; otherwise its
/// flags must name exactly one of date, text or number. Inline blank or
/// mixed-type cells are refused.
Expected<std::vector<cache_records::CellKind>, Error> cell_kinds_of_cache_definition(ByteSpan definition_part);

/// Serial of a date cell under the workbook's date system. Days before
/// 1900-03-01 in the 1900 system are refused.
Expected<double, Error> record_date_serial(const cache_records::RecordDate& date, bool date1904);

/// The date cell naming `serial`; a fractional serial (time of day) and a
/// serial outside the range `record_date_serial` produces are refused.
Expected<cache_records::RecordDate, Error> record_date_from_serial(double serial, bool date1904);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CACHE_RECORDS_H_
