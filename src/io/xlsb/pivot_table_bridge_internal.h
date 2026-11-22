// Record constants and small helpers shared by the read and write halves of the XLSB pivot table
// bridge (pivot_table_bridge.cpp, pivot_table_bridge_writer.cpp). Internal to those two translation units.

#ifndef FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_INTERNAL_H_
#define FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_INTERNAL_H_

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "io/pivot_attr_names.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_line_item_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/pivot_table_records.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace pivot_table_bridge_detail {

namespace ids = pivot_records;
using AttrList = std::vector<std::pair<std::string, std::string>>;
using RawRecords = std::vector<ids::RawRecord>;

constexpr const char* kContext = "context=xlsb_pivot_table_bridge";
constexpr const char* kNotWritableKind = "pivot_table";

/// BrtBeginSXView bytes no Excel-saved part varies: 1, 2, 4 and 7 whole, the
/// bits of 5, 6, 8 and 9 the codec leaves unnamed, and bytes 10, 11, 13.
constexpr std::uint8_t kViewFlags1 = 0x41;
constexpr std::uint8_t kViewFlags2 = 0x40;
constexpr std::uint8_t kViewFlags4 = 0xF0;
/// Byte 4: `asteriskTotals`, which Excel sets on a table whose cache stops saving its data.
constexpr std::uint8_t kViewAsteriskTotals = 0x02;
/// Byte 4: `showEmptyRow` and `showEmptyCol`.
constexpr std::uint8_t kViewShowEmptyRow = 0x04;
constexpr std::uint8_t kViewShowEmptyCol = 0x08;
constexpr std::uint8_t kViewFlags4Named = kViewAsteriskTotals | kViewShowEmptyRow | kViewShowEmptyCol;
/// Byte 5: `pageOverThenDown` and `subtotalHiddenItems`; byte 13 is `pageWrap`.
constexpr std::uint8_t kViewPageOverThenDown = 0x08;
constexpr std::uint8_t kViewSubtotalHiddenItems = 0x10;
/// Byte 7: `applyWidthHeightFormats`; every other bit is clear.
constexpr std::uint8_t kViewApplyWidthHeight = 0x20;
constexpr std::uint8_t kViewFlags5Named = kSxViewAutoFormat | kSxViewShowError | kSxViewRowGrand | kSxViewColGrand |
                                          kViewPageOverThenDown | kViewSubtotalHiddenItems;
constexpr std::uint8_t kViewFlags5Fixed = 0x04;
/// Byte 6 bit 0: the first row field is in outline form; clear without row fields.
constexpr std::uint8_t kViewRowOutline = 0x01;
constexpr std::uint8_t kViewFlags6Named = kViewRowOutline | kSxViewMergeItem | kSxViewHasGrandCaption;
constexpr std::uint8_t kViewFlags6Fixed = 0x0A;
/// Byte 8: `compact` with `compactData`, `outline` with `outlineData`; Excel sets each pair together.
constexpr std::uint8_t kViewCompactPair = 0x09;
constexpr std::uint8_t kViewOutlinePair = 0x06;
constexpr std::uint8_t kViewFlags8Named =
    kViewCompactPair | kViewOutlinePair | kSxViewNoDropZones | kSxViewNoErrorCaption | kSxViewNoMissingCaption;
constexpr std::uint32_t kViewReserved20 = 1;
constexpr std::uint8_t kDataOnRows = 1;
constexpr std::uint8_t kDataOnCols = 2;
constexpr std::uint32_t kNoDataPosition = 0xFFFFFFFFU;

/// BrtBeginSXVD: `dragToRow` (cleared on a calculated field); `dragToCol` / `dragToPage`, cleared together; a bit
/// cleared by `includeNewItemsInFilter="1"`; the unnamed bits every Excel-saved field sets; the default
/// `itemPageCount`.
constexpr std::uint32_t kFieldDragToRow = 0x00000001U;
constexpr std::uint32_t kFieldDragPair = 0x00000006U;
constexpr std::uint32_t kFieldNewItemsOutsideFilter = 0x00040000U;
constexpr std::uint32_t kFieldFixed = 0x00008018U;
constexpr std::uint32_t kFieldNamed = kFieldDragToRow | kFieldDragPair | kSxFieldShowAll | kSxFieldOutline |
                                      kSxFieldInsertBlankRow | kSxFieldSubtotalTop | kSxFieldAutoSort |
                                      kSxFieldAscendSort | kSxFieldMeasureFilter | kFieldNewItemsOutsideFilter;
constexpr std::uint32_t kItemPageCount = 10;
constexpr std::uint32_t kFieldReserved16 = 0xFFFFFFFFU;
/// `subtotals` bit of `kPivotSubtotalNames[k]` is `1 << (k + 1)`.
constexpr std::uint16_t kSubtotalMask = 0x0FFF;

/// BrtBeginSXVI kinds: a cache item, `default`, then `kPivotSubtotalNames` in order.
constexpr std::uint8_t kItemData = 0;
constexpr std::uint8_t kItemDefault = 1;
constexpr std::uint8_t kItemFirstFn = 2;
constexpr std::uint32_t kNoItemIndex = 0xFFFFFFFFU;

constexpr std::uint32_t kValuesField = static_cast<std::uint32_t>(pivot::kValuesFieldPosition);
constexpr std::uint32_t kNoPageItem = 0x001000FEU;

/// BrtTableStyleClient bits. fFirstColumn has no `<pivotTableStyleInfo>`
/// attribute: Windows Excel sets it on every table, Mac Excel on none.
constexpr std::uint16_t kStyleFirstColumn = 0x0001;
constexpr std::uint16_t kStyleLastColumn = 0x0002;
constexpr std::uint16_t kStyleRowStripes = 0x0004;
constexpr std::uint16_t kStyleColStripes = 0x0008;
constexpr std::uint16_t kStyleRowHeaders = 0x0010;
constexpr std::uint16_t kStyleColHeaders = 0x0020;

/// BrtBeginPRFilter bit set in every Excel-saved reference, `selected` or not.
constexpr std::uint32_t kAreaFilterAlways = 0x000001U;

/// The SXView14 / SXView16 blocks of every Excel-saved part, and the table
/// extLst `<ext>` elements the xlsx twin carries for them. Excel writes SXView14
/// only for `hideValuesRow="1"`, and omits both it and its `<ext>` for `0`.
constexpr std::array<std::uint8_t, 4> kView14Frt = {0x02, 0x0E, 0x00, 0x00};
constexpr std::array<std::uint8_t, 21> kView14Payload = {0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
                                                         0x00, 0x00, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF,
                                                         0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF, 0xFF};
constexpr std::array<std::uint8_t, 4> kView16Frt = {0x03, 0x11, 0x00, 0x00};
constexpr std::array<std::uint8_t, 3> kView16Payload = {0x01, 0x01, 0x00};
constexpr std::string_view kView14Uri = "{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}";
constexpr std::string_view kView16Uri = "{747A6164-185A-40DC-8AA5-F01512510D54}";
constexpr std::string_view kView14Ext =
    "<ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotTableDefinition "
    "hideValuesRow=\"1\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext>";
/// The x14 extension at `hideValuesRow="0"`, which Excel drops: its `.xlsb` holds no SXView14 record.
constexpr std::string_view kView14OffExt =
    "<ext uri=\"{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}\" "
    "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"><x14:pivotTableDefinition "
    "hideValuesRow=\"0\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"/></ext>";
constexpr std::string_view kView16Ext =
    "<ext uri=\"{747A6164-185A-40DC-8AA5-F01512510D54}\" "
    "xmlns:xpdl=\"http://schemas.microsoft.com/office/spreadsheetml/2016/pivotdefaultlayout\">"
    "<xpdl:pivotTableDefinition16/></ext>";
/// The version-info extension of a table over rich values, which its xlsb part holds as a version-info block.
constexpr std::string_view kVersionInfoUri = "{9F748A41-CAEA-4470-BF7A-CE61E8FFA7F9}";
constexpr std::string_view kVersionInfoExt =
    "<ext uri=\"{9F748A41-CAEA-4470-BF7A-CE61E8FFA7F9}\" "
    "xmlns:xxpvi=\"http://schemas.microsoft.com/office/spreadsheetml/2022/pivotVersionInfo\"><xxpvi:pivotVersionInfo>"
    "<xxpvi:lastUpdateFeature>RichData</xxpvi:lastUpdateFeature></xxpvi:pivotVersionInfo></ext>";
/// `pivotTableDefinition` extension URIs Office defines beyond the two above
/// ([MS-XLSX] x15 pivotTableData and pivotTableUISettings). Their records are
/// unmapped, so a table carrying one is refused; any other URI is one Excel's
/// xlsb save drops.
constexpr std::string_view kOfficeTableExtUris[] = {"{E67621CE-5B39-4880-91FE-76760E9C1902}",
                                                    "{44433962-1CF7-4059-B4EE-95C3D5FFCF73}"};

constexpr std::string_view kXrPrefix = "xr";

/// Root attributes Excel writes on every pivot table, `0` but for `applyWidthHeightFormats`
/// (SXView byte 7); an absent one reads as `0`.
constexpr std::string_view kApplyFormats[] = {"applyNumberFormats",    "applyBorderFormats",
                                              "applyFontFormats",      "applyPatternFormats",
                                              "applyAlignmentFormats", "applyWidthHeightFormats"};
constexpr std::size_t kApplyWidthHeight = std::size(kApplyFormats) - 1U;

/// `<i t>` spellings by BrtBeginSXLI `t`.
constexpr std::string_view kLineItemTypes[] = {"data",    "default", "sum",     "countA", "avg",
                                               "max",     "min",     "product", "count",  "stdDev",
                                               "stdDevP", "var",     "varP",    "grand",  "blank"};
static_assert(std::size(kLineItemTypes) == kSxLineItemTypeMax + 1U);

/// BrtBeginSXDI `show_data_as` codes, in ST_ShowDataAs order.
constexpr pivot::ShowValuesAs kShowDataAs[] = {
    pivot::ShowValuesAs::Normal,
    pivot::ShowValuesAs::DifferenceFrom,
    pivot::ShowValuesAs::PercentOf,
    pivot::ShowValuesAs::PercentDifferenceFrom,
    pivot::ShowValuesAs::RunningTotalInRow,
    pivot::ShowValuesAs::PercentOfRow,
    pivot::ShowValuesAs::PercentOfCol,
    pivot::ShowValuesAs::PercentOfTotal,
    pivot::ShowValuesAs::Index,
};

/// SXDI14 `pivotShowAs` codes from kSxDataShowAsPercentOfParentRow, in ST_PivotShowAs order.
constexpr pivot::ShowValuesAs kPivotShowAs[] = {
    pivot::ShowValuesAs::PercentOfParentRow, pivot::ShowValuesAs::PercentOfParentCol,
    pivot::ShowValuesAs::PercentOfParent,    pivot::ShowValuesAs::PercentOfRunningTotal,
    pivot::ShowValuesAs::RankAscending,      pivot::ShowValuesAs::RankDescending,
};
static_assert(std::size(kPivotShowAs) == kSxDataShowAsRankDescending - kSxDataShowAsPercentOfParentRow + 1U);

/// `ST_PivotFilterType` spellings by BrtBeginSXFilter type (PivotFilterType order).
constexpr std::string_view kPivotFilterTypes[] = {
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
/// BrtBeginSXFilter `cch*` presence bits, in the order the strings follow.
constexpr std::string_view kFilterStringAttrs[] = {"name", "description", "stringValue1", "stringValue2"};
constexpr std::uint32_t kNoFilterIndex = 0xFFFFFFFFU;

/// BrtBeginPRule `type` codes and the model area types they spell.
struct AreaTypeCode {
  std::uint8_t code;
  pivot::PivotAreaType type;
};
constexpr AreaTypeCode kAreaTypes[] = {
    {pivot_area::kAreaTypeNormal, pivot::PivotAreaType::Normal},
    {pivot_area::kAreaTypeData, pivot::PivotAreaType::Data},
    {pivot_area::kAreaTypeAll, pivot::PivotAreaType::All},
    {pivot_area::kAreaTypeButton, pivot::PivotAreaType::Button},
};

inline Error Corrupt(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb pivot table: " + message, kContext);
}

inline Error Unmeasured(const std::string& what) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "unmeasured xlsb pivot table state: " + what, kContext);
}

inline const char* Flag(bool on) {
  return on ? "1" : "0";
}

/// Parses a decimal integer no greater than `max`, spelled as `std::to_string` spells it.
inline std::optional<std::uint32_t> ParseUnsigned(std::string_view text, std::uint32_t max) {
  if (text.empty() || text.size() > 10U || (text.size() > 1U && text.front() == '0')) {
    return std::nullopt;
  }
  std::uint64_t value = 0;
  for (const char c : text) {
    if (c < '0' || c > '9') {
      return std::nullopt;
    }
    value = value * 10U + static_cast<std::uint64_t>(c - '0');
  }
  if (value > max) {
    return std::nullopt;
  }
  return static_cast<std::uint32_t>(value);
}

template <std::size_t N>
bool SameBytes(const std::vector<std::uint8_t>& bytes, const std::array<std::uint8_t, N>& want) {
  return bytes.size() == N && std::equal(want.begin(), want.end(), bytes.begin());
}

inline std::string_view SubtotalItemType(std::uint8_t kind) {
  return kind == kItemDefault ? std::string_view("default") : kPivotSubtotalNames[kind - kItemFirstFn].item;
}

/// The BrtBeginSXVI kind of a subtotal `<item t>`, or 0 for any other type.
inline std::uint8_t SubtotalItemKind(std::string_view type) {
  if (type == "default") {
    return kItemDefault;
  }
  for (std::size_t k = 0; k < kPivotSubtotalNames.size(); ++k) {
    if (kPivotSubtotalNames[k].item == type) {
      return static_cast<std::uint8_t>(kItemFirstFn + k);
    }
  }
  return kItemData;
}

/// The report form a compact / outline pair spells, as the xlsx reader resolves it.
inline pivot::PivotLayout FormOf(bool compact, bool outline) {
  return compact ? pivot::PivotLayout::Compact : (outline ? pivot::PivotLayout::Outline : pivot::PivotLayout::Tabular);
}
}  // namespace pivot_table_bridge_detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PIVOT_TABLE_BRIDGE_INTERNAL_H_
