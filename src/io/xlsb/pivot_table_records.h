// Byte codec for an XLSB pivot table part (`xl/pivotTables/pivotTableN.bin`).
//
// Record order: [ACUID] SXView( Location, SXVDs( SXVD( [SXVIs] [FRT(SXVD14)] [AutoSortScope] )* ),
// [ISXVDRws], SXLIRws, [ISXVDCols], SXLICols, [SXPIs], SXDIs( SXDI( [FRT(SXDI14)] )* ),
// [SXFormats | SXCrtFormats], TableStyleClient, [SXFilters], [FRT(SXView14)], FRT(SXView16) ).
// Every record is decoded into a
// field of the structs below and re-encoded byte for byte; flag bytes whose
// bits are not all named are carried whole. The autoSortScope and filter
// blocks are carried verbatim as raw records for the codecs that own them.
// Any record, flag or optional string not listed here is refused.

#ifndef FORMULON_IO_XLSB_PIVOT_TABLE_RECORDS_H_
#define FORMULON_IO_XLSB_PIVOT_TABLE_RECORDS_H_

#include <array>
#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_line_item_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/zip_reader.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// BrtBeginSXView flag bits, by header byte.
inline constexpr std::uint8_t kSxViewIndentMask = 0x7F;         ///< byte 3: `indent`.
inline constexpr std::uint8_t kSxViewNoHeaders = 0x80;          ///< byte 3: `showHeaders="0"`.
inline constexpr std::uint8_t kSxViewAutoFormat = 0x01;         ///< byte 5: `useAutoFormatting`.
inline constexpr std::uint8_t kSxViewShowError = 0x02;          ///< byte 5: `showError`.
inline constexpr std::uint8_t kSxViewRowGrand = 0x20;           ///< byte 5: `rowGrandTotals`.
inline constexpr std::uint8_t kSxViewColGrand = 0x40;           ///< byte 5: `colGrandTotals`.
inline constexpr std::uint8_t kSxViewMergeItem = 0x04;          ///< byte 6: `mergeItem`.
inline constexpr std::uint8_t kSxViewHasGrandCaption = 0x10;    ///< byte 6: grand-total caption follows.
inline constexpr std::uint8_t kSxViewNoDropZones = 0x10;        ///< byte 8: clear for `gridDropZones="1"`.
inline constexpr std::uint8_t kSxViewNoErrorCaption = 0x40;     ///< byte 8: clear when an error caption follows.
inline constexpr std::uint8_t kSxViewNoMissingCaption = 0x80;   ///< byte 8: clear when a missing caption follows.
inline constexpr std::uint8_t kSxViewSingleFieldFilter = 0x02;  ///< byte 9: `multipleFieldFilters="0"`.

/// BrtBeginSXView: the table header (32 bytes) and its strings.
struct SxView {
  std::uint8_t created_version = 0;  ///< byte 0
  std::uint8_t flags1 = 0;           ///< byte 1, carried
  std::uint8_t flags2 = 0;           ///< byte 2, carried
  std::uint8_t flags3 = 0;           ///< byte 3: indent and kSxViewNoHeaders
  std::uint8_t flags4 = 0;           ///< byte 4, carried
  std::uint8_t flags5 = 0;           ///< byte 5: kSxViewAutoFormat .. kSxViewColGrand
  std::uint8_t flags6 = 0;           ///< byte 6: kSxViewMergeItem; caption bit set from `grand_total_caption`
  std::uint8_t flags7 = 0;           ///< byte 7, carried
  /// byte 8: 0x06 outline/outlineData, 0x09 compact/compactData, kSxViewNoDropZones;
  /// the caption bits are set from `error_caption` / `missing_caption`.
  std::uint8_t flags8 = 0;
  std::uint8_t flags9 = 0;                    ///< byte 9: kSxViewSingleFieldFilter
  std::uint8_t flags10 = 0;                   ///< byte 10, carried
  std::uint8_t flags11 = 0;                   ///< byte 11, carried
  std::uint8_t data_axis = 0;                 ///< byte 12: 1 rows, 2 columns (`dataOnRows`)
  std::uint8_t flags13 = 0;                   ///< byte 13, carried
  std::uint8_t updated_version = 0;           ///< byte 14
  std::uint8_t min_refreshable_version = 0;   ///< byte 15
  std::uint32_t data_position = 0xFFFFFFFFU;  ///< bytes 16-19 (`dataPosition`; all ones = absent)
  std::uint32_t reserved20 = 0;               ///< bytes 20-23, carried
  std::uint32_t reserved24 = 0;               ///< bytes 24-27, carried
  std::uint32_t cache_id = 0;                 ///< bytes 28-31 (`cacheId`)
  std::string name;
  std::string data_caption;
  /// At most one of the three captions may be set: the order of several is unmeasured.
  std::optional<std::string> grand_total_caption;
  std::optional<std::string> error_caption;
  std::optional<std::string> missing_caption;
};

/// BrtBeginSXLocation: `location`, offsets absolute.
struct SxLocation {
  std::uint32_t first_row = 0;
  std::uint32_t last_row = 0;
  std::uint32_t first_col = 0;
  std::uint32_t last_col = 0;
  std::uint32_t first_header_row = 0;
  std::uint32_t first_data_row = 0;
  std::uint32_t first_data_col = 0;
  std::uint32_t row_page_count = 0;
  std::uint32_t col_page_count = 0;
};

/// BrtBeginSXVI flag bits.
inline constexpr std::uint16_t kSxItemHidden = 0x0001;      ///< `h="1"`
inline constexpr std::uint16_t kSxItemCollapsed = 0x0002;   ///< `sd="0"`
inline constexpr std::uint16_t kSxItemHasCaption = 0x0010;  ///< `n`; set from `caption`
inline constexpr std::uint16_t kSxItemCalculated = 0x0004;  ///< `f="1"`
inline constexpr std::uint16_t kSxItemMissing = 0x0008;     ///< `m="1"`
/// Largest item kind (ST_ItemType order, `varP`).
inline constexpr std::uint8_t kSxItemKindMax = 12;

/// BrtBeginSXVI: one `<item>`.
struct SxItem {
  std::uint8_t kind = 0;  ///< 0 data, 1 default, 2.. subtotal functions
  std::uint16_t flags = 0;
  std::uint32_t index = 0;  ///< cache item index; all ones for subtotal kinds
  std::optional<std::string> caption;
};

/// BrtBeginSXVD header bits.
inline constexpr std::uint8_t kSxFieldAxisRow = 0x01;
inline constexpr std::uint8_t kSxFieldAxisCol = 0x02;
inline constexpr std::uint8_t kSxFieldAxisPage = 0x04;
inline constexpr std::uint8_t kSxFieldAxisData = 0x08;
/// `subtotals` bits: default, sum, countA, avg, max, min, product, count, stdDev, stdDevP, var, varP.
inline constexpr std::uint16_t kSxFieldSubtotalDefault = 0x0001;
inline constexpr std::uint8_t kSxFieldCompact = 0x10;         ///< `name_flags`
inline constexpr std::uint8_t kSxFieldHasDisplayName = 0x20;  ///< `name_flags`; set from `display_name`
inline constexpr std::uint32_t kSxFieldShowAll = 0x00000020;
inline constexpr std::uint32_t kSxFieldOutline = 0x00000040;
inline constexpr std::uint32_t kSxFieldInsertBlankRow = 0x00000080;
inline constexpr std::uint32_t kSxFieldSubtotalTop = 0x00000100;
inline constexpr std::uint32_t kSxFieldAutoSort = 0x00001000;
inline constexpr std::uint32_t kSxFieldAscendSort = 0x00002000;
inline constexpr std::uint32_t kSxFieldMeasureFilter = 0x00020000;

/// BrtBeginSXVD: one `<pivotField>` with its items and autoSortScope.
struct SxField {
  std::uint8_t axis = 0;              ///< byte 0: kSxFieldAxis* bits
  std::uint16_t subtotals = 0;        ///< bytes 1-2
  std::uint8_t name_flags = 0;        ///< byte 3
  std::uint32_t num_fmt = 0;          ///< bytes 4-7 (`numFmtId`)
  std::uint32_t flags = 0;            ///< bytes 8-11; 0x06 dragToCol/dragToPage, unnamed bits carried
  std::uint32_t item_page_count = 0;  ///< bytes 12-15
  std::uint32_t reserved16 = 0;       ///< bytes 16-19, carried
  std::optional<std::string> display_name;
  /// Absent when the field has no BrtBeginSXVIs block.
  std::optional<std::vector<SxItem>> items;
  /// The x14 `fillDownLabels="1"`, held by the one SXVD14 payload Excel writes.
  /// The block is written after the items; a field without items is unmeasured.
  bool fill_down_labels = false;
  /// BrtBeginAutoSortScope through BrtEndAutoSortScope, verbatim; empty when absent.
  std::vector<pivot_records::RawRecord> auto_sort_scope;
};

/// BrtBeginSXPI: one `<pageField>` (name and caption strings are unmeasured and refused).
struct SxPageField {
  std::uint32_t field = 0;
  std::uint32_t item = 0;  ///< 0x001000FE when no single item is selected
  std::uint32_t hierarchy = 0xFFFFFFFFU;
};

/// SXDI14 `pivotShowAs` codes: `percentOfParentRow` 9 through `rankDescending` 14, in ST_PivotShowAs order.
inline constexpr std::uint32_t kSxDataShowAsPercentOfParentRow = 9U;
inline constexpr std::uint32_t kSxDataShowAsRankDescending = 14U;

/// BrtBeginSXDI: one `<dataField>`.
struct SxDataField {
  std::uint32_t field = 0;
  std::uint32_t subtotal = 0;
  std::uint32_t show_data_as = 0;
  std::uint32_t base_field = 0;
  std::uint32_t base_item = 0;
  std::uint32_t num_fmt = 0;
  std::string name;
  /// The x14 `pivotShowAs` (SXDI14 bytes 4-7), beside `show_data_as` 0.
  std::optional<std::uint32_t> pivot_show_as;
};

/// BrtBeginSXFormat: one `<format>`. Bytes 0-1 are `01 00` for `action="formatting"`,
/// the only action measured.
struct SxFormat {
  std::uint32_t dxf_id = 0;  ///< bytes 2-5
  pivot_area::PivotArea area;
};

/// BrtBeginSXCrtFormat: one `<chartFormat>`. Only `chart="0"` (bytes 0-3) and
/// `series="1"` (bytes 8-11) are measured, so they are not members.
struct SxChartFormat {
  std::uint32_t format = 0;  ///< bytes 4-7
  pivot_area::PivotArea area;
};

/// BrtTableStyleClient: `pivotTableStyleInfo`.
struct SxStyleClient {
  std::uint16_t flags = 0;
  std::optional<std::string> name;  ///< nullopt for a null XLNullableWideString
};

/// One FRT block (BrtFRTBegin, a begin/end record pair, BrtFRTEnd), payloads as read.
struct SxFrtBlock {
  std::vector<std::uint8_t> frt_payload;
  std::vector<std::uint8_t> payload;
};

/// A whole pivot table part.
struct PivotTableDoc {
  std::optional<std::array<std::uint8_t, 16>> uid;  ///< BrtUid in the ACUID block
  SxView view;
  SxLocation location;
  std::vector<SxField> fields;
  std::optional<std::vector<std::uint32_t>> row_fields;  ///< `-2` is the data pseudo-field
  SxLineItems row_items;
  std::optional<std::vector<std::uint32_t>> col_fields;
  SxLineItems col_items;
  std::optional<std::vector<SxPageField>> page_fields;
  std::vector<SxDataField> data_fields;
  /// SXFormats; empty when absent. A table with both lists is unmeasured.
  std::vector<SxFormat> formats;
  /// SXCrtFormats; empty when absent.
  std::vector<SxChartFormat> chart_formats;
  SxStyleClient style;
  /// BrtBeginSXFilters through BrtEndSXFilters, verbatim; empty when absent.
  std::vector<pivot_records::RawRecord> filters;
  /// Around BrtBeginSXView14 / BrtEndSXView14; Excel omits it for `hideValuesRow="0"`.
  std::optional<SxFrtBlock> view14;
  SxFrtBlock view16;  ///< around BrtBeginSXView16 / BrtEndSXView16
};

/// Decodes a pivot table part.
///
/// Errors: `kIoXlsbRecordTruncated` for broken framing, `kIoXlsbUnknownRecord`
/// for a record out of place, `kIoXlsbRecordCorrupt` for a malformed or
/// uncharacterised payload, including a format, chart format, pivot area or
/// x14 record state no Excel-saved part shows.
Expected<PivotTableDoc, Error> decode_pivot_table(ByteSpan part);

/// Encodes `doc` as a pivot table part. Counts are derived from the vectors and
/// optional-string presence bits from the optionals.
///
/// Errors: `kInvalidArgument` when `doc` sets more than one view caption, or
/// holds a format, chart format, pivot area or x14 state no Excel-saved part shows.
Expected<std::vector<std::uint8_t>, Error> encode_pivot_table(const PivotTableDoc& doc);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_TABLE_RECORDS_H_
