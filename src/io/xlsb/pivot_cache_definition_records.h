//
// Codec for the XLSB pivot cache definition part
// (`xl/pivotCache/pivotCacheDefinitionN.bin`), as Excel lays it out:
//
//   [ACUID] BrtBeginPivotCacheDef
//     BrtBeginPCDSource BrtBeginPCDSRange BrtEndPCDSRange BrtEndPCDSource
//     BrtBeginPCDFields { BrtBeginPCDField [item table] [grouping] [names] BrtEndPCDField }* BrtEndPCDFields
//     [calculated items] [FRT x14 cache definition] [FRT dataSourceClean]
//   BrtEndPivotCacheDef
//
// Decoding is fail-closed: a record out of place, a flag bit or record kind
// Excel's own parts never show, or a size that does not match its counts is
// an error. The grouping block of a field, a calculated field's PNames block
// and the calculated-item block are carried verbatim as `RawRecord`s. `encode(decode(x)) == x` byte for byte.

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
/// BrtBeginPivotCacheDef flag byte: `saveData`; with it clear the cache has no records part or relationship.
inline constexpr std::uint8_t kCacheSaveData = 0x01U;
/// BrtBeginPivotCacheDef flag byte: `invalid`, set only on a cache whose records Excel keeps as rich (`RichInfo`).
inline constexpr std::uint8_t kCacheInvalid = 0x02U;
/// BrtBeginPivotCacheDef flag byte: `enableRefresh`, set in every Excel-saved cache.
inline constexpr std::uint8_t kCacheEnableRefresh = 0x10U;
/// `saveData` and `enableRefresh` together, a cache's default.
inline constexpr std::uint8_t kCacheSaveDataEnableRefresh = kCacheSaveData | kCacheEnableRefresh;

/// BrtBeginPCDField flag: the field comes from the source (`databaseField`, default on).
inline constexpr std::uint8_t kFieldDatabase = 0x04U;
/// BrtBeginPCDField second flag byte: a formula follows the name (`cacheField@formula`).
inline constexpr std::uint8_t kFieldFormula = 0x01U;

/// BrtBeginPCDFAtbl flags (`sharedItems@contains*`).
inline constexpr std::uint16_t kAtblSemiMixedTypesString = 0x0009U;  ///< containsSemiMixedTypes + containsString.
inline constexpr std::uint16_t kAtblNonDate = 0x0002U;
inline constexpr std::uint16_t kAtblDate = 0x0004U;
inline constexpr std::uint16_t kAtblBlank = 0x0010U;
inline constexpr std::uint16_t kAtblMixedTypes = 0x0020U;
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
  kString,             ///< BrtPCDIString (24): `text`.
  kNumber,             ///< BrtPCDINumber (21): `number`.
  kAttributedString,   ///< BrtPCDIAString (31): `text` + `attributes`.
  kAttributedMissing,  ///< BrtPCDIAMissing (27): `attributes` only.
  kMissing,            ///< BrtPCDIMissing (20): no payload (`<m/>`).
  kBoolean,            ///< BrtPCDIBoolean (22): `boolean`.
  kError,              ///< BrtPCDIError (23): `error`, the BErr byte.
  kDate                ///< BrtPCDIDatetime (25): `date`.
};

struct Item {
  ItemType type = ItemType::kString;
  std::string text;
  double number = 0.0;
  std::uint8_t attributes = 0;  ///< kItemUnused / kItemFormula.
  bool boolean = false;
  std::uint8_t error = 0;
  Date date;
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
  /// A calculated field's formula (kFieldFormula): PivotParsedFormula rgce and rgcb.
  std::optional<std::vector<std::uint8_t>> formula_rgce;
  std::vector<std::uint8_t> formula_rgcb;
  /// The formula's PNames block (BrtBeginPNames .. BrtEndPNames), verbatim; empty when absent.
  std::vector<pivot_records::RawRecord> names;
};

/// BrtBeginPCDSource + BrtBeginPCDSRange of a worksheet source (`cacheSource type="worksheet"`): a range on `sheet`,
/// or, when `name` is set, a defined name or table (`sheet` set only for a sheet-scoped name).
struct WorksheetSource {
  std::uint32_t connection_id = 0;
  std::string sheet;
  std::string name;
  MergeRange ref;
};

/// The whole part. `fields` count is derived.
/// The rich-data extensions of a cache holding rich values: `xxpvi:cacheVersionInfo` with lastRefreshFeature
/// `RichData`, then `xprd:richInfo` with pivotIgnoreInvalidCache="1". The records part keeps the ordinary
/// relationship type, and the header names no records relationship id: richInfo carries it instead.
struct RichInfo {
  std::array<std::uint8_t, kGuidBytes> cache_guid{};  ///< `pivotCacheGuid`, wire byte order.
  std::string records_rel_id;
};

struct Document {
  std::optional<std::array<std::uint8_t, kGuidBytes>> uid;  ///< `xr:uid`, wire byte order.
  std::uint8_t refreshed_version = 0;
  std::uint8_t min_refreshable_version = 0;
  std::uint8_t created_version = 0;
  std::uint8_t flags = kCacheSaveDataEnableRefresh;
  std::uint32_t missing_items_limit = 0xFFFFFFFFU;  ///< 0xFFFFFFFF when absent.
  double refreshed_date = 0.0;
  std::uint32_t record_count = 0;
  std::optional<std::string> refreshed_by;  ///< absent on a cache no one has refreshed by name.
  std::string records_rel_id;
  WorksheetSource source;
  std::vector<Field> fields;
  /// BrtBeginPCDCalcItems .. BrtEndPCDCalcItems, verbatim; empty when absent.
  std::vector<pivot_records::RawRecord> calculated_items;
  bool has_pcd14 = false;                         ///< `x14:pivotCacheDefinition` (no attributes).
  std::optional<RichInfo> rich;                   ///< with the `invalid` flag; between x14 and dataSourceClean.
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
