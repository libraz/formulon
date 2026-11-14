//
// Codec for the XLSB pivot cache definition part
// (`xl/pivotCache/pivotCacheDefinitionN.bin`), as Excel lays it out:
//
//   [ACUID] BrtBeginPivotCacheDef
//     BrtBeginPCDSource BrtBeginPCDSRange BrtEndPCDSRange BrtEndPCDSource
//     BrtBeginPCDFields { BrtBeginPCDField [item table] [grouping] BrtEndPCDField }* BrtEndPCDFields
//     [calculated items] [FRT x14 cache definition] [FRT dataSourceClean]
//   BrtEndPivotCacheDef
//
// Decoding is fail-closed: a record out of place, a flag bit or record kind
// Excel's own parts never show, or a size that does not match its counts is
// an error. The grouping block of a field and the calculated-item block are
// carried verbatim as `RawRecord`s. `encode(decode(x)) == x` byte for byte.

#ifndef FORMULON_IO_XLSB_PIVOT_CACHE_DEFINITION_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_CACHE_DEFINITION_RECORDS_H_

#include <array>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/zip_reader.h"
#include "merge_range.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_cache_definition {

/// BrtBeginPivotCacheDef flag byte: `refreshOnLoad`.
inline constexpr std::uint8_t kCacheRefreshOnLoad = 0x04U;
/// BrtBeginPivotCacheDef flag byte: `saveData` and `enableRefresh`, set together in every Excel-saved cache.
inline constexpr std::uint8_t kCacheSaveDataEnableRefresh = 0x11U;

/// BrtBeginPCDField flag: the field comes from the source (`databaseField`, default on).
inline constexpr std::uint8_t kFieldDatabase = 0x04U;

/// BrtBeginPCDFAtbl flags (`sharedItems@contains*`).
inline constexpr std::uint16_t kAtblSemiMixedTypesString = 0x0009U;  ///< containsSemiMixedTypes + containsString.
inline constexpr std::uint16_t kAtblNonDate = 0x0002U;
inline constexpr std::uint16_t kAtblDate = 0x0004U;
inline constexpr std::uint16_t kAtblBlank = 0x0010U;
inline constexpr std::uint16_t kAtblNumber = 0x0040U;
inline constexpr std::uint16_t kAtblInteger = 0x0080U;
/// The min/max doubles follow (`minValue`/`maxValue`, or `minDate`/`maxDate` as serials).
inline constexpr std::uint16_t kAtblMinMax = 0x0100U;
/// Set in every item table Excel writes.
inline constexpr std::uint16_t kAtblAlways = 0x0400U;

/// Attribute-block flags of an attributed item: `u="1"` (unused) and `f="1"` (calculated).
inline constexpr std::uint8_t kItemUnused = 0x01U;
inline constexpr std::uint8_t kItemFormula = 0x02U;

/// A date as Excel's 1900/1904 serial calendar spells it (serial 0 of 1900 is 1900-01-00).
struct Date {
  std::uint16_t year = 0;
  std::uint16_t month = 0;
  std::uint8_t day = 0;
  std::uint8_t hour = 0;
  std::uint8_t minute = 0;
  std::uint8_t second = 0;
};

/// One single-record shared item.
enum class ItemType : std::uint8_t {
  kString,            ///< BrtPCDIString (24): `text`.
  kNumber,            ///< BrtPCDINumber (21): `number`.
  kAttributedString,  ///< BrtPCDIAString (31): `text` + `attributes`.
  kAttributedMissing  ///< BrtPCDIAMissing (27): `attributes` only.
};

struct Item {
  ItemType type = ItemType::kString;
  std::string text;
  double number = 0.0;
  std::uint8_t attributes = 0;  ///< kItemUnused / kItemFormula.
};

/// BrtBeginPCDIRun element type.
enum class RunType : std::uint16_t { kNumbers = 0x0001U, kStrings = 0x0002U, kDates = 0x0020U };

/// One BrtBeginPCDIRun .. BrtEndPCDIRun: items of one type; only the vector `type` names is used.
struct ItemRun {
  RunType type = RunType::kStrings;
  std::vector<double> numbers;
  std::vector<std::string> strings;
  std::vector<Date> dates;

  std::size_t size() const;
};

/// One record of the item table: a single item or a run.
struct ItemEntry {
  bool is_run = false;
  Item item;
  ItemRun run;
};

/// BrtBeginPCDFAtbl .. BrtEndPCDFAtbl (`sharedItems`); `count` is derived from `entries`.
struct ItemTable {
  std::uint16_t flags = 0;
  double min = 0.0;  ///< Written only when `flags` has kAtblMinMax.
  double max = 0.0;
  std::vector<ItemEntry> entries;

  std::size_t item_count() const;
};

/// BrtBeginPCDField .. BrtEndPCDField (`cacheField`).
struct Field {
  std::uint8_t flags = kFieldDatabase;
  std::uint16_t num_fmt_id = 0;
  std::string name;
  std::optional<ItemTable> items;
  /// Records between the item table and BrtEndPCDField (the fieldGroup block), verbatim.
  std::vector<pivot_records::RawRecord> grouping;
};

/// BrtBeginPCDSource + BrtBeginPCDSRange of a worksheet source (`cacheSource type="worksheet"`).
struct WorksheetSource {
  std::uint32_t connection_id = 0;
  std::string sheet;
  MergeRange ref;
};

/// The whole part. `fields` count is derived.
struct Document {
  std::optional<std::array<std::uint8_t, kGuidBytes>> uid;  ///< `xr:uid`, wire byte order.
  std::uint8_t refreshed_version = 0;
  std::uint8_t min_refreshable_version = 0;
  std::uint8_t created_version = 0;
  std::uint8_t flags = kCacheSaveDataEnableRefresh;
  std::uint32_t missing_items_limit = 0xFFFFFFFFU;  ///< 0xFFFFFFFF when absent.
  double refreshed_date = 0.0;
  std::uint32_t record_count = 0;
  std::string refreshed_by;
  std::string records_rel_id;
  WorksheetSource source;
  std::vector<Field> fields;
  /// BrtBeginPCDCalcItems .. BrtEndPCDCalcItems, verbatim; empty when absent.
  std::vector<pivot_records::RawRecord> calculated_items;
  bool has_pcd14 = false;                         ///< `x14:pivotCacheDefinition` (no attributes).
  std::optional<std::uint8_t> data_source_clean;  ///< `xlpdsc:dataSourceClean`.
};

}  // namespace pivot_cache_definition

/// Decodes a cache definition part. Errors: kIoXlsbRecordTruncated / kIoXlsbRecordCorrupt for
/// malformed framing or sizes; kIoXlsbUnknownRecord for a record or flag outside the layout, with
/// a message starting "unmeasured" when the record kind exists but its layout has no Excel sample.
Expected<pivot_cache_definition::Document, Error> decode_pivot_cache_definition(ByteSpan bytes);

/// Encodes `doc` as a cache definition part.
std::vector<std::uint8_t> encode_pivot_cache_definition(const pivot_cache_definition::Document& doc);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_CACHE_DEFINITION_RECORDS_H_
