#include "io/xlsb/pivot_table_bridge.h"

#include <algorithm>
#include <array>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <iterator>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "io/auto_filter_xml.h"
#include "io/ooxml_defs.h"
#include "io/pivot_attr_names.h"
#include "io/pivot_built_defaults.h"
#include "io/pivot_sort_scope.h"
#include "io/pivot_table_reader.h"
#include "io/pivot_table_writer_internal.h"
#include "io/xlsb/auto_filter_records.h"
#include "io/xlsb/frt_records.h"
#include "io/xlsb/guid_text.h"
#include "io/xlsb/pivot_area_records.h"
#include "io/xlsb/pivot_line_item_records.h"
#include "io/xlsb/pivot_records_common.h"
#include "io/xlsb/pivot_table_records.h"
#include "io/xlsb/record.h"
#include "io/xlsb/record_writer.h"
#include "io/xml_utils.h"
#include "pivot/filter_engine.h"
#include "pivot/pivot_types.h"
#include "pugixml.hpp"
#include "utils/a1_ref.h"
#include "utils/status_macros.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace {

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

Error Corrupt(const std::string& message) {
  return make_error(FormulonErrorCode::kIoXlsbRecordCorrupt, "xlsb pivot table: " + message, kContext);
}

Error Unmeasured(const std::string& what) {
  return make_error(FormulonErrorCode::kIoXlsbUnknownRecord, "unmeasured xlsb pivot table state: " + what, kContext);
}

const char* Flag(bool on) {
  return on ? "1" : "0";
}

/// Parses a decimal integer no greater than `max`, spelled as `std::to_string` spells it.
std::optional<std::uint32_t> ParseUnsigned(std::string_view text, std::uint32_t max) {
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

std::string_view SubtotalItemType(std::uint8_t kind) {
  return kind == kItemDefault ? std::string_view("default") : kPivotSubtotalNames[kind - kItemFirstFn].item;
}

/// The BrtBeginSXVI kind of a subtotal `<item t>`, or 0 for any other type.
std::uint8_t SubtotalItemKind(std::string_view type) {
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
pivot::PivotLayout FormOf(bool compact, bool outline) {
  return compact ? pivot::PivotLayout::Compact : (outline ? pivot::PivotLayout::Outline : pivot::PivotLayout::Tabular);
}

Expected<void, Error> CheckView(const PivotTableDoc& doc) {
  const SxView& v = doc.view;
  const auto paired = [&v](std::uint8_t pair) {
    const std::uint8_t bits = static_cast<std::uint8_t>(v.flags8 & pair);
    return bits == 0U || bits == pair;
  };
  if (v.flags1 != kViewFlags1 || v.flags2 != kViewFlags2 || (v.flags4 & ~kViewFlags4Named) != kViewFlags4 ||
      (v.flags7 & ~kViewApplyWidthHeight) != 0U || (v.flags5 & ~kViewFlags5Named) != kViewFlags5Fixed ||
      (v.flags6 & ~kViewFlags6Named) != kViewFlags6Fixed || (v.flags8 & ~kViewFlags8Named) != 0U ||
      (v.flags9 & ~kSxViewSingleFieldFilter) != 0U || v.flags10 != 0U || v.flags11 != 0U ||
      v.reserved20 != kViewReserved20 || !paired(kViewCompactPair) || !paired(kViewOutlinePair)) {
    return Unmeasured("view header bits");
  }
  if (v.data_position != kNoDataPosition || (v.data_axis != kDataOnRows && v.data_axis != kDataOnCols)) {
    return Unmeasured("data axis or position");
  }
  if ((doc.view14 &&
       (!SameBytes(doc.view14->frt_payload, kView14Frt) || !SameBytes(doc.view14->payload, kView14Payload))) ||
      !SameBytes(doc.view16.frt_payload, kView16Frt) || !SameBytes(doc.view16.payload, kView16Payload) ||
      (doc.version_info && (!SameBytes(doc.version_info->frt_payload, pivot_records::kVersionInfoFrtVersion) ||
                            !SameBytes(doc.version_info->payload, pivot_records::kRichDataFeature)))) {
    return Unmeasured("table future records");
  }
  return Expected<void, Error>::Ok();
}

/// Root attributes as the xlsx reader captures them from Excel's part: the
/// namespace bookkeeping of `xr:uid`, then document order.
void ReadRootAttrs(const PivotTableDoc& doc, AttrList& attrs) {
  if (doc.uid) {
    attrs.emplace_back("xmlns:mc", std::string(kNsMarkupCompat));
    attrs.emplace_back("xmlns:xr", std::string(kNsXr));
    attrs.emplace_back("mc:Ignorable", std::string(kXrPrefix));
    attrs.emplace_back("xr:uid", FormatGuid(doc.uid->data()));
  }
  for (std::size_t i = 0; i < std::size(kApplyFormats); ++i) {
    const bool on = i == kApplyWidthHeight && (doc.view.flags7 & kViewApplyWidthHeight) != 0U;
    attrs.emplace_back(std::string(kApplyFormats[i]), Flag(on));
  }
  attrs.emplace_back("updatedVersion", std::to_string(doc.view.updated_version));
  // Excel leaves out a version of 0, the schema default.
  if (doc.view.min_refreshable_version != 0U) {
    attrs.emplace_back("minRefreshableVersion", std::to_string(doc.view.min_refreshable_version));
  }
  if ((doc.view.flags4 & kViewAsteriskTotals) != 0U) {
    attrs.emplace_back("asteriskTotals", "1");
  }
  if (doc.view.created_version != 0U) {
    attrs.emplace_back("createdVersion", std::to_string(doc.view.created_version));
  }
}

/// Appends the form attributes that depart from `layout`, as the xlsx reader keeps them.
void CaptureForm(bool compact, bool outline, pivot::PivotLayout layout, bool root, AttrList& out) {
  if (compact == (layout == pivot::PivotLayout::Compact) && outline == (layout != pivot::PivotLayout::Tabular)) {
    return;
  }
  out.emplace_back("compact", Flag(compact));
  if (root) {
    out.emplace_back("compactData", Flag(compact));
  }
  out.emplace_back("outline", Flag(outline));
  if (root) {
    out.emplace_back("outlineData", Flag(outline));
  }
}

Expected<void, Error> ReadLocation(const SxLocation& loc, pivot::PivotTable& table) {
  if (loc.last_row < loc.first_row || loc.last_col < loc.first_col || loc.last_row >= a1::kMaxRows ||
      loc.last_col >= a1::kMaxCols) {
    return Corrupt("location rectangle is inverted or past the grid");
  }
  if (loc.first_header_row < loc.first_row || loc.first_data_row < loc.first_row ||
      loc.first_data_col < loc.first_col || loc.first_header_row > loc.last_row || loc.first_data_row > loc.last_row ||
      loc.first_data_col > loc.last_col) {
    return Corrupt("location offsets fall outside the table");
  }
  table.set_anchor(loc.first_row, loc.first_col, loc.last_row - loc.first_row + 1U, loc.last_col - loc.first_col + 1U);
  table.mark_span_authored();
  const auto count = [](std::uint32_t v) { return v == 0U ? std::nullopt : std::optional<std::uint32_t>(v); };
  table.set_location_attributes(loc.first_header_row - loc.first_row, loc.first_data_row - loc.first_row,
                                loc.first_data_col - loc.first_col, count(loc.row_page_count),
                                count(loc.col_page_count));
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadItems(const SxField& in, pivot::PivotField& field) {
  if (!in.items) {
    return Expected<void, Error>::Ok();
  }
  if (in.items->empty()) {
    return Unmeasured("empty field item list");
  }
  for (const SxItem& item : *in.items) {
    if (item.kind == kItemData) {
      pivot::PivotItem entry;
      entry.visible = (item.flags & kSxItemHidden) == 0U;
      entry.caption = item.caption.value_or(std::string());
      entry.blank_caption = item.caption && item.caption->empty();
      entry.show_detail = (item.flags & kSxItemCollapsed) == 0U;
      entry.calculated = (item.flags & kSxItemCalculated) != 0U;
      if ((item.flags & kSxItemMissing) != 0U) {
        entry.passthrough_attrs.emplace_back("m", "1");
      }
      entry.has_cache_index = true;
      entry.cache_index = item.index;
      field.items.push_back(std::move(entry));
      continue;
    }
    const bool collapsed = (item.flags & kSxItemCollapsed) != 0U;
    if (item.index != kNoItemIndex || item.caption || (item.flags & ~kSxItemCollapsed) != 0U ||
        (collapsed && item.kind != kItemDefault)) {
      return Unmeasured("subtotal item flags");
    }
    const std::string_view type = SubtotalItemType(item.kind);
    std::string xml("<item t=\"");
    xml.append(type);
    xml.append(collapsed ? "\" sd=\"0\"/>" : "\"/>");
    field.marker_items.push_back(
        pivot::PivotMarkerItem{std::move(xml), std::string(type), field.items.size(), !collapsed});
  }
  return Expected<void, Error>::Ok();
}

/// The field's `<autoSortScope>` as the xlsx reader keeps it in `raw_children_xml`.
Expected<void, Error> ReadAutoSortScope(const RawRecords& block, std::string& out) {
  if (block.empty()) {
    return Expected<void, Error>::Ok();
  }
  if (block.size() < 3U || !block.front().payload.empty() || !block.back().payload.empty() ||
      block.back().type != ids::kEndAutoSortScope) {
    return Unmeasured("sort scope framing");
  }
  std::size_t pos = 1;
  ASSIGN_OR_RETURN(const pivot_area::PivotArea area, pivot_area::decode_pivot_area(block, &pos));
  if (pos + 1U != block.size() || area.field != pivot_area::kNoField || area.type != pivot_area::kAreaTypeNormal ||
      (area.flags & ~pivot_area::kAreaDataOnly) != 0U || !area.filters || area.filters->empty()) {
    return Unmeasured("sort scope pivot area shape");
  }
  out.append("<autoSortScope><pivotArea");
  if ((area.flags & pivot_area::kAreaDataOnly) == 0U) {
    out.append(" dataOnly=\"0\"");
  }
  // Excel's one sorted-by-value scope carries outline="0" fieldPosition="0" with no bit for either.
  out.append(" outline=\"0\" fieldPosition=\"0\"><references");
  append_xml_attr_uint(out, "count", area.filters->size());
  out.push_back('>');
  for (const pivot_area::PivotAreaFilter& ref : *area.filters) {
    if ((ref.flags & kAreaFilterAlways) == 0U ||
        (ref.flags & ~(kAreaFilterAlways | pivot_area::kFilterSelected)) != 0U) {
      return Unmeasured("sort scope reference flags");
    }
    out.append("<reference");
    append_xml_attr_uint(out, "field", ref.field);
    append_xml_attr_uint(out, "count", ref.items.size());
    if ((ref.flags & pivot_area::kFilterSelected) == 0U) {
      out.append(" selected=\"0\"");
    }
    out.push_back('>');
    for (const std::uint32_t x : ref.items) {
      out.append("<x");
      append_xml_attr_uint(out, "v", x);
      out.append("/>");
    }
    out.append("</reference>");
  }
  out.append("</references></pivotArea></autoSortScope>");
  return Expected<void, Error>::Ok();
}

/// A format's pivot area as the xlsx reader types Excel's `<pivotArea>`. The
/// codec has already refused every shape no Excel part shows.
pivot::PivotTableArea ReadTableArea(const pivot_area::PivotArea& in) {
  pivot::PivotTableArea area;
  if (in.field != pivot_area::kNoField) {
    area.field = in.field;
  }
  for (const AreaTypeCode& entry : kAreaTypes) {
    if (entry.code == in.type) {
      area.type = entry.type;
    }
  }
  const auto has = [&in](std::uint32_t bit) { return (in.flags & bit) != 0U; };
  area.data_only = has(pivot_area::kAreaDataOnly);
  area.label_only = has(pivot_area::kAreaLabelOnly);
  area.grand_row = has(pivot_area::kAreaGrandRow);
  area.grand_col = has(pivot_area::kAreaGrandCol);
  area.collapsed_levels_are_subtotals = has(pivot_area::kAreaCollapsedLevelsAreSubtotals);
  if ((in.flags & pivot_area::kAreaAxisMask) == pivot_area::kAreaAxisRow) {
    area.axis = pivot::PivotAxis::Row;
  }
  // Every Excel-saved area carries outline="0" fieldPosition="0" with no bit for either.
  area.outline = false;
  area.field_position = 0U;
  if (in.filters) {
    for (const pivot_area::PivotAreaFilter& filter : *in.filters) {
      pivot::PivotTableAreaReference ref;
      ref.field = filter.field;
      ref.items = filter.items;
      ref.selected = (filter.flags & pivot_area::kFilterSelected) != 0U;
      area.references.push_back(std::move(ref));
    }
  }
  return area;
}

void ReadFormats(const PivotTableDoc& doc, pivot::PivotTable& table) {
  for (const SxFormat& in : doc.formats) {
    pivot::PivotFormat format;
    format.dxf_id = in.dxf_id;
    format.area = ReadTableArea(in.area);
    table.mutable_formats().push_back(std::move(format));
  }
  for (const SxChartFormat& in : doc.chart_formats) {
    pivot::PivotChartFormat format;
    format.format = in.format;
    format.series = true;
    format.area = ReadTableArea(in.area);
    table.mutable_chart_formats().push_back(std::move(format));
  }
}

/// The table extLst the xlsx reader keeps for Excel's part.
std::string TableExtLst(bool hide_values_row, bool rich_version_info) {
  std::string out("<extLst>");
  if (hide_values_row) {
    out.append(kView14Ext);
  }
  out.append(kView16Ext);
  if (rich_version_info) {
    out.append(kVersionInfoExt);
  }
  out.append("</extLst>");
  return out;
}

Expected<pivot::PivotField, Error> ReadField(const SxField& in) {
  constexpr std::uint8_t kAxes = kSxFieldAxisRow | kSxFieldAxisCol | kSxFieldAxisPage;
  const std::uint8_t axis = static_cast<std::uint8_t>(in.axis & kAxes);
  if ((in.axis & ~(kAxes | kSxFieldAxisData)) != 0U || (axis & (axis - 1U)) != 0U) {
    return Corrupt("field sits on several axes");
  }
  const std::uint32_t drag = in.flags & kFieldDragPair;
  if ((in.name_flags & ~(kSxFieldCompact | kSxFieldHasDisplayName)) != 0U || (in.flags & ~kFieldNamed) != kFieldFixed ||
      (drag != 0U && drag != kFieldDragPair) || in.item_page_count != kItemPageCount ||
      in.reserved16 != kFieldReserved16 || (in.subtotals & ~kSubtotalMask) != 0U ||
      (in.display_name && in.display_name->empty())) {
    return Unmeasured("field header bits");
  }
  pivot::PivotField field;
  field.axis = axis == kSxFieldAxisRow              ? pivot::PivotAxis::Row
               : axis == kSxFieldAxisCol            ? pivot::PivotAxis::Col
               : axis == kSxFieldAxisPage           ? pivot::PivotAxis::Page
               : (in.axis & kSxFieldAxisData) != 0U ? pivot::PivotAxis::Value
                                                    : pivot::PivotAxis::None;
  field.custom_name = in.display_name.value_or(std::string());
  if (in.num_fmt != 0U) {
    field.number_format = std::to_string(in.num_fmt);
  }
  field.subtotal_top = (in.flags & kSxFieldSubtotalTop) != 0U;
  field.default_subtotal = (in.subtotals & kSxFieldSubtotalDefault) != 0U;
  for (std::size_t k = 0; k < kPivotSubtotalNames.size(); ++k) {
    if ((in.subtotals & (1U << (k + 1U))) != 0U) {
      field.subtotal_fns.push_back(kPivotSubtotalNames[k].fn);
    }
  }
  if (field.default_subtotal && !field.subtotal_fns.empty()) {
    return Unmeasured("default subtotal together with custom functions");
  }
  // Windows Excel sets the ascending bit on manually ordered fields too; it is derived on write.
  if ((in.flags & kSxFieldAutoSort) != 0U) {
    field.sort.stated = true;
    field.sort.ascending = (in.flags & kSxFieldAscendSort) != 0U;
  } else {
    field.sort.manual = true;
  }
  RETURN_IF_ERROR(ReadItems(in, field));
  field.insert_blank_row = (in.flags & kSxFieldInsertBlankRow) != 0U;
  field.show_all = (in.flags & kSxFieldShowAll) != 0U;
  if ((in.flags & kFieldDragToRow) == 0U) {
    field.passthrough_attrs.emplace_back("dragToRow", "0");
  }
  if (drag == 0U) {
    field.passthrough_attrs.emplace_back("dragToCol", "0");
    field.passthrough_attrs.emplace_back("dragToPage", "0");
  }
  if ((in.flags & kSxFieldMeasureFilter) != 0U) {
    field.passthrough_attrs.emplace_back("measureFilter", "1");
  }
  // Measured on manually ordered fields only, where Excel also sets the ascending bit.
  if ((in.flags & kFieldNewItemsOutsideFilter) == 0U) {
    if ((in.flags & kSxFieldAutoSort) != 0U || (in.flags & kSxFieldAscendSort) == 0U) {
      return Unmeasured("new items in a filter on a field that is not manually ordered");
    }
    field.passthrough_attrs.emplace_back("includeNewItemsInFilter", "1");
  }
  RETURN_IF_ERROR(ReadAutoSortScope(in.auto_sort_scope, field.raw_children_xml));
  field.fill_down_labels = in.fill_down_labels;
  return field;
}

Expected<void, Error> ReadFieldOrder(const std::optional<std::vector<std::uint32_t>>& in, std::size_t field_count,
                                     std::vector<std::uint32_t>& order, std::optional<std::size_t>& values_position) {
  if (!in) {
    return Expected<void, Error>::Ok();
  }
  if (in->empty()) {
    return Unmeasured("empty axis field list");
  }
  for (std::size_t i = 0; i < in->size(); ++i) {
    const std::uint32_t x = (*in)[i];
    if (x == kValuesField) {
      if (values_position) {
        return Corrupt("axis field list names the Values field twice");
      }
      values_position = i;
      continue;
    }
    if (x >= field_count) {
      return Corrupt("axis field list names a field the table does not have");
    }
    order.push_back(x);
  }
  return Expected<void, Error>::Ok();
}

/// `<rowItems>` / `<colItems>` spelled as Excel writes them.
std::string LineItemsXml(const SxLineItems& block, std::string_view tag) {
  // Excel writes neither the block nor the element for an axis with no line items.
  if (block.items.empty()) {
    return std::string();
  }
  std::string out("<");
  out.append(tag);
  append_xml_attr_uint(out, "count", block.items.size());
  out.push_back('>');
  for (const SxLineItem& line : block.items) {
    out.append("<i");
    if (line.t != 0U) {
      append_xml_attr(out, "t", kLineItemTypes[line.t]);
    }
    if (line.r != 0U) {
      append_xml_attr_uint(out, "r", line.r);
    }
    if (line.i != 0U) {
      append_xml_attr_uint(out, "i", line.i);
    }
    if (line.x.empty()) {
      out.append("/>");
      continue;
    }
    out.push_back('>');
    for (const std::uint32_t v : line.x) {
      out.append("<x");
      if (v != 0U) {
        append_xml_attr_uint(out, "v", v);
      }
      out.append("/>");
    }
    out.append("</i>");
  }
  out.append("</");
  out.append(tag);
  out.push_back('>');
  return out;
}

Expected<void, Error> ReadDataFields(const PivotTableDoc& doc, pivot::PivotTable& table) {
  for (const SxDataField& in : doc.data_fields) {
    if (in.field >= doc.fields.size()) {
      return Corrupt("data field names a field the table does not have");
    }
    if (in.subtotal >= kPivotAggregationNames.size() || in.show_data_as >= std::size(kShowDataAs)) {
      return Unmeasured("data field function or display mode");
    }
    pivot::PivotDataField entry;
    entry.name = in.name;
    entry.field_index = in.field;
    entry.aggregation = static_cast<pivot::Aggregation>(kPivotAggregationNames[in.subtotal].value);
    if (in.num_fmt != 0U) {
      entry.number_format = std::to_string(in.num_fmt);
    }
    // The codec admits SXDI14 only with a known code, beside a plain SXDI display mode.
    entry.show_as = in.pivot_show_as ? kPivotShowAs[*in.pivot_show_as - kSxDataShowAsPercentOfParentRow]
                                     : kShowDataAs[in.show_data_as];
    entry.show_as_base_field = in.base_field;
    entry.show_as_base_item = in.base_item;
    table.mutable_data_fields().push_back(std::move(entry));
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> ReadStyle(const SxStyleClient& in, pivot::PivotTable& table) {
  constexpr std::uint16_t kNamed =
      kStyleFirstColumn | kStyleLastColumn | kStyleRowStripes | kStyleColStripes | kStyleRowHeaders | kStyleColHeaders;
  if ((in.flags & ~kNamed) != 0U || (in.name && in.name->empty())) {
    return Unmeasured("style client bits");
  }
  pivot::PivotStyleInfo info;
  info.name = in.name.value_or(std::string());
  info.show_last_column = (in.flags & kStyleLastColumn) != 0U;
  info.show_row_stripes = (in.flags & kStyleRowStripes) != 0U;
  info.show_col_stripes = (in.flags & kStyleColStripes) != 0U;
  info.show_row_headers = (in.flags & kStyleRowHeaders) != 0U;
  info.show_col_headers = (in.flags & kStyleColHeaders) != 0U;
  table.set_style_info(std::move(info));
  return Expected<void, Error>::Ok();
}

/// Each listed field sits on its list's axis, and a field carries the data bit
/// exactly when a data field refers to it, as the writer derives both.
Expected<void, Error> CheckAxes(const PivotTableDoc& doc, const pivot::PivotTable& table) {
  const auto listed_on = [&table](const std::vector<std::uint32_t>& order, pivot::PivotAxis axis) {
    return std::all_of(order.begin(), order.end(),
                       [&table, axis](std::uint32_t f) { return table.fields()[f].axis == axis; });
  };
  const bool pages_ok =
      std::all_of(table.page_fields().begin(), table.page_fields().end(), [&table](const pivot::PivotPageField& p) {
        return table.fields()[p.field_index].axis == pivot::PivotAxis::Page;
      });
  if (!listed_on(table.row_field_order(), pivot::PivotAxis::Row) ||
      !listed_on(table.col_field_order(), pivot::PivotAxis::Col) || !pages_ok) {
    return Corrupt("axis field list names a field on another axis");
  }
  for (std::size_t i = 0; i < doc.fields.size(); ++i) {
    const bool referenced =
        std::any_of(doc.data_fields.begin(), doc.data_fields.end(), [i](const SxDataField& d) { return d.field == i; });
    if (referenced != ((doc.fields[i].axis & kSxFieldAxisData) != 0U)) {
      return Corrupt("field data bit disagrees with the data fields");
    }
  }
  return Expected<void, Error>::Ok();
}

/// One SXFILTER (BrtBeginSXFilter payload `begin`, just read) as its `<filter>` element.
Expected<void, Error> AppendFilterXml(ByteSpan& cursor, ByteSpan begin, std::size_t field_count,
                                      std::size_t data_field_count, std::string& out) {
  ASSIGN_OR_RETURN(const std::uint32_t fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t mp_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t type, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t eval_order, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t id, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_fld, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint32_t measure_hier, read_u32(begin));
  ASSIGN_OR_RETURN(const std::uint16_t flags, read_u16(begin));
  if (type == 0U || type >= std::size(kPivotFilterTypes) || (flags & ~0x000FU) != 0U) {
    return Unmeasured("filter type or string flags");
  }
  if (fld >= field_count || (measure_fld != kNoFilterIndex && measure_fld >= data_field_count)) {
    return Corrupt("filter names a field or data field the table does not have");
  }
  out.append("<filter");
  append_xml_attr_uint(out, "fld", fld);
  append_xml_attr_uint_if(out, "mpFld", mp_fld, kNoFilterIndex);
  append_xml_attr(out, "type", kPivotFilterTypes[type]);
  append_xml_attr_int(out, "evalOrder", static_cast<std::int32_t>(eval_order));
  append_xml_attr_uint(out, "id", id);
  append_xml_attr_uint_if(out, "iMeasureHier", measure_hier, kNoFilterIndex);
  append_xml_attr_uint_if(out, "iMeasureFld", measure_fld, kNoFilterIndex);
  for (std::size_t bit = 0; bit < std::size(kFilterStringAttrs); ++bit) {
    if ((flags & (1U << bit)) != 0U) {
      ASSIGN_OR_RETURN(const std::string text, read_xlwidestring(begin));
      append_xml_attr(out, kFilterStringAttrs[bit], text);
    }
  }
  if (begin.size != 0U) {
    return Unmeasured("filter header trailing data");
  }
  out.push_back('>');
  // Excel wraps a null uid ahead of each filter's AutoFilter.
  ASSIGN_OR_RETURN(const XlsbRecord ac, read_record(cursor));
  ASSIGN_OR_RETURN(const XlsbRecord uid, read_record(cursor));
  ASSIGN_OR_RETURN(const XlsbRecord ac_end, read_record(cursor));
  const bool null_uid =
      uid.payload.size == kGuidBytes &&
      std::all_of(uid.payload.data, uid.payload.data + uid.payload.size, [](std::uint8_t b) { return b == 0U; });
  if (ac.type != ids::kACBegin || ac.payload.size != kAcUidBeginPayload.size() ||
      !std::equal(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end(), ac.payload.data) || uid.type != ids::kUid ||
      !null_uid || ac_end.type != ids::kACEnd || ac_end.payload.size != 0U) {
    return Unmeasured("filter uid block");
  }
  ASSIGN_OR_RETURN(const XlsbRecord afilter, read_record(cursor));
  if (afilter.type != kBeginAFilter) {
    return Corrupt("filter carries no AutoFilter");
  }
  ASSIGN_OR_RETURN(const AutoFilter criteria, read_auto_filter_block(cursor, afilter.payload));
  out.append(serialize_auto_filter(criteria));
  ASSIGN_OR_RETURN(const XlsbRecord end, read_record(cursor));
  if (end.type != ids::kEndSXFilter || end.payload.size != 0U) {
    return Unmeasured("filter future record");
  }
  out.append("</filter>");
  return Expected<void, Error>::Ok();
}

/// SXFILTERS spelled as the `<filters>` element the xlsx reader decodes.
Expected<std::string, Error> FiltersXml(const RawRecords& block, std::size_t field_count,
                                        std::size_t data_field_count) {
  std::vector<std::uint8_t> bytes;
  for (const ids::RawRecord& rec : block) {
    emit_record(bytes, rec.type, rec.payload);
  }
  ByteSpan cursor{bytes.data(), bytes.size()};
  ASSIGN_OR_RETURN(const XlsbRecord begin, read_record(cursor));
  if (begin.type != ids::kBeginSXFilters) {
    return Corrupt("filter list does not open with its count");
  }
  ByteSpan count_payload = begin.payload;
  ASSIGN_OR_RETURN(const std::uint32_t declared, read_u32(count_payload));
  std::string out("<filters>");
  for (std::uint32_t seen = 0;; ++seen) {
    ASSIGN_OR_RETURN(const XlsbRecord rec, read_record(cursor));
    if (rec.type == ids::kEndSXFilters) {
      if (seen != declared || cursor.size != 0U) {
        return Corrupt("filter count does not match the filters decoded");
      }
      break;
    }
    if (rec.type != ids::kBeginSXFilter) {
      return Corrupt("filter list carries an uncharacterised record");
    }
    RETURN_IF_ERROR(AppendFilterXml(cursor, rec.payload, field_count, data_field_count, out));
  }
  out.append("</filters>");
  return out;
}

/// The table's own report-form attributes, when the xlsx reader kept an override.
struct FormOverride {
  std::optional<bool> compact;
  std::optional<bool> compact_data;
  std::optional<bool> outline;
  std::optional<bool> outline_data;
};

bool ParseFlag(std::string_view value, std::optional<bool>& out) {
  if (out.has_value() || (value != "0" && value != "1")) {
    return false;
  }
  out = value == "1";
  return true;
}

/// Reads the root passthrough attributes into `doc`; false for any the records cannot hold.
bool WriteRootAttrs(const pivot::PivotTable& table, PivotTableDoc& doc, FormOverride& form) {
  enum : unsigned {
    kCreated = 1U << 6U,
    kUpdated = 1U << 7U,
    kMinRefreshable = 1U << 8U,
    kNsMc = 1U << 9U,
    kNsXrDecl = 1U << 10U,
    kIgnorable = 1U << 11U,
    kUidSeen = 1U << 12U,
    kAsteriskTotals = 1U << 13U,
    kUidBookkeeping = kNsMc | kNsXrDecl | kIgnorable | kUidSeen,
  };
  unsigned seen = 0;
  // Each attribute absent from the part is its schema default, but `updatedVersion`
  // is the xlsx writer's own.
  doc.view.flags4 = 0;
  doc.view.flags7 = 0;
  doc.view.created_version = 0;
  doc.view.min_refreshable_version = 0;
  doc.view.updated_version = kBuiltTableUpdatedVersion;
  const auto take = [&seen](unsigned bit) {
    if ((seen & bit) != 0U) {
      return false;
    }
    seen |= bit;
    return true;
  };
  const auto version = [](std::string_view text, std::uint8_t& out) {
    const std::optional<std::uint32_t> v = ParseUnsigned(text, 0xFFU);
    out = static_cast<std::uint8_t>(v.value_or(0U));
    return v.has_value();
  };
  for (const auto& [name, value] : table.passthrough_attrs()) {
    bool ok = false;
    std::size_t apply = 0;
    while (apply < std::size(kApplyFormats) && name != kApplyFormats[apply]) {
      ++apply;
    }
    if (apply < std::size(kApplyFormats)) {
      ok = take(1U << apply) && (value == "0" || (apply == kApplyWidthHeight && value == "1"));
      if (ok && value == "1") {
        doc.view.flags7 = kViewApplyWidthHeight;
      }
    } else if (name == "xmlns:mc") {
      ok = take(kNsMc) && value == kNsMarkupCompat;
    } else if (name == "xmlns:xr") {
      ok = take(kNsXrDecl) && value == kNsXr;
    } else if (name == "mc:Ignorable") {
      ok = take(kIgnorable) && value == kXrPrefix;
    } else if (name == "xr:uid") {
      std::array<std::uint8_t, kGuidBytes> guid{};
      ok = take(kUidSeen) && ParseGuid(value, guid);
      doc.uid = guid;
    } else if (name == "createdVersion") {
      ok = take(kCreated) && version(value, doc.view.created_version);
    } else if (name == "updatedVersion") {
      ok = take(kUpdated) && version(value, doc.view.updated_version);
    } else if (name == "minRefreshableVersion") {
      ok = take(kMinRefreshable) && version(value, doc.view.min_refreshable_version);
    } else if (name == "asteriskTotals") {
      // Excel omits it at 0.
      ok = take(kAsteriskTotals) && value == "1";
      doc.view.flags4 = kViewAsteriskTotals;
    } else if (name == "compact") {
      ok = ParseFlag(value, form.compact);
    } else if (name == "compactData") {
      ok = ParseFlag(value, form.compact_data);
    } else if (name == "outline") {
      ok = ParseFlag(value, form.outline);
    } else if (name == "outlineData") {
      ok = ParseFlag(value, form.outline_data);
    } else if (name == "chartFormat") {
      // Excel omits `chartFormat` at 0.
      const std::optional<std::uint32_t> id = ParseUnsigned(value, 0xFFFFFFFFU);
      ok = doc.view.reserved24 == 0U && id.has_value() && *id != 0U;
      doc.view.reserved24 = id.value_or(0U);
    } else {
      // `dataOnRows` restates the Values position; no Excel part carries a `dataPosition`.
      ok = name == "dataOnRows" && value == Flag(table.row_values_position().has_value());
    }
    if (!ok) {
      return false;
    }
  }
  const unsigned bookkeeping = seen & kUidBookkeeping;
  return bookkeeping == (doc.uid ? kUidBookkeeping : 0U);
}

/// The Excel defaults the header bytes this bridge does not name stand for.
bool HasExcelDefaults(const pivot::PivotTable& t) {
  return t.enable_drill() && t.show_missing() && t.show_drill() && !t.print_drill() && t.show_data_tips() &&
         !t.field_print_titles() && t.item_print_titles() && t.show_data_drop_down() && t.preserve_formatting() &&
         t.page_wrap() <= 0xFFU && t.custom_list_sort() && t.indent() <= kSxViewIndentMask && t.alt_text().empty() &&
         t.alt_text_summary().empty();
}

/// The field's `compact` / `outline`: its own override, else the table's form.
bool FieldForm(const pivot::PivotField& field, pivot::PivotLayout layout, bool& compact, bool& outline) {
  std::optional<bool> own_compact;
  std::optional<bool> own_outline;
  for (const auto& [name, value] : field.passthrough_attrs) {
    if ((name == "compact" && !ParseFlag(value, own_compact)) ||
        (name == "outline" && !ParseFlag(value, own_outline))) {
      return false;
    }
  }
  if (own_compact.has_value() != own_outline.has_value()) {
    return false;
  }
  compact = own_compact.value_or(layout == pivot::PivotLayout::Compact);
  outline = own_outline.value_or(layout != pivot::PivotLayout::Tabular);
  return true;
}

/// A decimal `numFmtId`, or 0 when absent; false for a spelling the record cannot keep.
bool NumFmt(const std::string& text, std::uint32_t& out) {
  if (text.empty()) {
    out = 0;
    return true;
  }
  const std::optional<std::uint32_t> id = ParseUnsigned(text, 0xFFFFFFFFU);
  out = id.value_or(0U);
  return out != 0U;
}

bool WriteItems(const pivot::PivotField& field, std::optional<std::vector<SxItem>>& out) {
  if (field.items.empty() && field.marker_items.empty()) {
    return true;
  }
  std::vector<SxItem> items;
  for (std::size_t i = 0; i < field.items.size(); ++i) {
    const pivot::PivotItem& item = field.items[i];
    SxItem rec;
    rec.kind = kItemData;
    rec.index = item.has_cache_index ? item.cache_index : static_cast<std::uint32_t>(i);
    rec.flags =
        static_cast<std::uint16_t>((item.visible ? 0U : kSxItemHidden) | (item.show_detail ? 0U : kSxItemCollapsed) |
                                   (item.calculated ? kSxItemCalculated : 0U));
    for (const auto& [name, value] : item.passthrough_attrs) {
      if (name != "m" || value != "1" || (rec.flags & kSxItemMissing) != 0U) {
        return false;
      }
      rec.flags = static_cast<std::uint16_t>(rec.flags | kSxItemMissing);
    }
    if (!item.caption.empty() || item.blank_caption) {
      rec.caption = item.caption;
    }
    items.push_back(std::move(rec));
  }
  // The subtotal entries follow the field's selection as the xlsx writer emits
  // them; `grand` and `blank` have no item kind.
  const std::vector<pivot::SubtotalFn> fns = pivot::ordered_subtotal_fns(field);
  std::vector<std::uint8_t> expected;
  if (field.default_subtotal && fns.empty()) {
    expected.push_back(kItemDefault);
  }
  for (const pivot::SubtotalFn fn : fns) {
    expected.push_back(static_cast<std::uint8_t>(kItemFirstFn + static_cast<std::uint8_t>(fn)));
  }
  std::vector<std::uint8_t> authored;
  bool default_detail = true;
  for (const pivot::PivotMarkerItem& marker : field.marker_items) {
    const std::uint8_t kind = SubtotalItemKind(marker.type);
    if (kind == kItemData) {
      return false;
    }
    authored.push_back(kind);
    if (kind == kItemDefault) {
      default_detail = marker.show_detail;
    }
  }
  const bool keep = authored == expected;
  for (std::size_t k = 0; k < expected.size(); ++k) {
    const std::uint8_t kind = expected[k];
    bool detail = kind != kItemDefault || default_detail;
    if (keep) {
      const pivot::PivotMarkerItem& marker = field.marker_items[k];
      const std::string plain = "<item t=\"" + std::string(marker.type) + "\"/>";
      const bool collapsed_default = kind == kItemDefault && marker.xml == "<item t=\"default\" sd=\"0\"/>";
      if (marker.xml != plain && !collapsed_default) {
        return false;
      }
      detail = kind != kItemDefault || marker.show_detail;
    }
    SxItem rec;
    rec.kind = kind;
    rec.index = kNoItemIndex;
    rec.flags = static_cast<std::uint16_t>(detail ? 0U : kSxItemCollapsed);
    items.push_back(std::move(rec));
  }
  out = std::move(items);
  return true;
}

/// True when `node` carries no attribute outside `allowed`.
bool OnlyAttrs(const pugi::xml_node& node, std::initializer_list<std::string_view> allowed) {
  for (pugi::xml_attribute attr = node.first_attribute(); attr; attr = attr.next_attribute()) {
    if (std::find(allowed.begin(), allowed.end(), std::string_view(attr.name())) == allowed.end()) {
      return false;
    }
  }
  return true;
}

/// The element children of `node`, or nullopt when it holds text or anything else.
std::optional<std::vector<pugi::xml_node>> Elements(const pugi::xml_node& node) {
  std::vector<pugi::xml_node> out;
  for (pugi::xml_node child = node.first_child(); child; child = child.next_sibling()) {
    if (child.type() != pugi::node_element) {
      return std::nullopt;
    }
    out.push_back(child);
  }
  return out;
}

/// Parses `xml` as one element named `name`.
pugi::xml_node LoadOne(pugi::xml_document& doc, std::string_view xml, std::string_view name) {
  if (!doc.load_buffer(xml.data(), xml.size(), pugi::parse_default, pugi::encoding_utf8)) {
    return pugi::xml_node();
  }
  const pugi::xml_node root = doc.first_child();
  if (!root || root.next_sibling() || root.type() != pugi::node_element || std::string_view(root.name()) != name) {
    return pugi::xml_node();
  }
  return root;
}

bool WriteAutoSortScope(std::string_view xml, RawRecords& out) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node scope = LoadOne(doc, xml, "autoSortScope");
  const std::optional<std::vector<pugi::xml_node>> areas = scope ? Elements(scope) : std::nullopt;
  if (!areas || areas->size() != 1U || scope.first_attribute()) {
    return false;
  }
  const pugi::xml_node area_node = areas->front();
  const std::optional<std::vector<pugi::xml_node>> refs_nodes = Elements(area_node);
  if (std::string_view(area_node.name()) != "pivotArea" ||
      !OnlyAttrs(area_node, {"dataOnly", "outline", "fieldPosition"}) || attr_str(area_node, "outline") != "0" ||
      attr_str(area_node, "fieldPosition") != "0" ||
      (area_node.attribute("dataOnly") && attr_str(area_node, "dataOnly") != "0") || !refs_nodes ||
      refs_nodes->size() != 1U || std::string_view(refs_nodes->front().name()) != "references") {
    return false;
  }
  pivot_area::PivotArea area;
  area.flags = area_node.attribute("dataOnly") ? 0U : pivot_area::kAreaDataOnly;
  const pugi::xml_node refs = refs_nodes->front();
  const std::optional<std::vector<pugi::xml_node>> ref_nodes = Elements(refs);
  if (!OnlyAttrs(refs, {"count"}) || !ref_nodes || ref_nodes->empty() ||
      attr_u32(refs, "count", 0U) != ref_nodes->size()) {
    return false;
  }
  std::vector<pivot_area::PivotAreaFilter> filters;
  for (const pugi::xml_node& ref : *ref_nodes) {
    const std::optional<std::vector<pugi::xml_node>> xs = Elements(ref);
    const std::optional<std::uint32_t> field = parse_xml_u32_attr_strict(ref.attribute("field"));
    if (std::string_view(ref.name()) != "reference" || !OnlyAttrs(ref, {"field", "count", "selected"}) || !field ||
        !xs || attr_u32(ref, "count", 0U) != xs->size() ||
        (ref.attribute("selected") && attr_str(ref, "selected") != "0")) {
      return false;
    }
    pivot_area::PivotAreaFilter filter;
    filter.field = *field;
    filter.flags = kAreaFilterAlways | (ref.attribute("selected") ? 0U : pivot_area::kFilterSelected);
    for (const pugi::xml_node& x : *xs) {
      if (std::string_view(x.name()) != "x" || !OnlyAttrs(x, {"v"}) || x.first_child()) {
        return false;
      }
      filter.items.push_back(attr_u32(x, "v", 0U));
    }
    filters.push_back(std::move(filter));
  }
  area.filters = std::move(filters);
  out.push_back(ids::RawRecord{ids::kBeginAutoSortScope, {}});
  pivot_area::encode_pivot_area(area, &out);
  out.push_back(ids::RawRecord{ids::kEndAutoSortScope, {}});
  return true;
}

/// A typed pivot area in the one shape Excel's parts show; false for any other.
bool WriteTableArea(const pivot::PivotTableArea& in, pivot_area::PivotArea& out) {
  if (in.outline || in.field_position != std::optional<std::uint32_t>(0U) ||
      (in.axis && *in.axis != pivot::PivotAxis::Row)) {
    return false;
  }
  const auto* type = std::find_if(std::begin(kAreaTypes), std::end(kAreaTypes),
                                  [&in](const AreaTypeCode& entry) { return entry.type == in.type; });
  if (type == std::end(kAreaTypes)) {
    return false;
  }
  out.field = in.field.value_or(pivot_area::kNoField);
  out.type = type->code;
  out.flags = (in.axis ? pivot_area::kAreaAxisRow : 0U) | (in.data_only ? pivot_area::kAreaDataOnly : 0U) |
              (in.label_only ? pivot_area::kAreaLabelOnly : 0U) | (in.grand_row ? pivot_area::kAreaGrandRow : 0U) |
              (in.grand_col ? pivot_area::kAreaGrandCol : 0U) |
              (in.collapsed_levels_are_subtotals ? pivot_area::kAreaCollapsedLevelsAreSubtotals : 0U);
  if (in.references.empty()) {
    out.filters.reset();
    return true;
  }
  std::vector<pivot_area::PivotAreaFilter> filters;
  for (const pivot::PivotTableAreaReference& ref : in.references) {
    if (!ref.field || ref.items.empty()) {
      return false;
    }
    pivot_area::PivotAreaFilter filter;
    filter.field = *ref.field;
    filter.flags = kAreaFilterAlways | (ref.selected ? pivot_area::kFilterSelected : 0U);
    filter.items = ref.items;
    filters.push_back(std::move(filter));
  }
  out.filters = std::move(filters);
  return true;
}

/// The typed formats and chart formats; a table holding both is unmeasured.
bool WriteFormats(const pivot::PivotTable& table, PivotTableDoc& doc) {
  if (!table.formats().empty() && !table.chart_formats().empty()) {
    return false;
  }
  for (const pivot::PivotFormat& in : table.formats()) {
    SxFormat out;
    if (in.action != pivot::PivotFormatAction::Formatting || !in.dxf_id || !WriteTableArea(in.area, out.area)) {
      return false;
    }
    out.dxf_id = *in.dxf_id;
    doc.formats.push_back(std::move(out));
  }
  for (const pivot::PivotChartFormat& in : table.chart_formats()) {
    SxChartFormat out;
    if (in.chart != 0U || !in.series || !WriteTableArea(in.area, out.area)) {
      return false;
    }
    out.format = in.format;
    doc.chart_formats.push_back(std::move(out));
  }
  return true;
}

/// Reads the table extLst into whether SXView14 is written: Excel's two
/// `<ext>` elements as the reader keeps them, the SXView14 one optional, plus
/// any extension of a URI Office does not define, which Excel's xlsb save drops. The RichData version-info
/// extension after them sets `rich_version_info`.
bool WriteTableExtLst(std::string_view xml, bool& view14, bool& rich_version_info) {
  pugi::xml_document doc;
  const pugi::xml_node ext_lst = LoadOne(doc, xml, "extLst");
  const std::optional<std::vector<pugi::xml_node>> exts = ext_lst ? Elements(ext_lst) : std::nullopt;
  if (!exts || ext_lst.first_attribute()) {
    return false;
  }
  view14 = false;
  rich_version_info = false;
  bool view16 = false;
  bool view14_off_seen = false;
  for (const pugi::xml_node& ext : *exts) {
    const std::string_view uri = ext.attribute("uri").as_string();
    if (std::string_view(ext.name()) != "ext") {
      return false;
    }
    if (uri == kView14Uri || uri == kView16Uri) {
      // The SXView14 extension precedes the SXView16 one, each at most once.
      const bool is14 = uri == kView14Uri;
      const std::string ext_xml = raw_xml(ext);
      const bool view14_off = is14 && ext_xml == kView14OffExt;
      if (view16 || (is14 && (view14 || view14_off_seen)) ||
          (ext_xml != (is14 ? kView14Ext : kView16Ext) && !view14_off)) {
        return false;
      }
      if (view14_off) {
        view14_off_seen = true;
      } else {
        (is14 ? view14 : view16) = true;
      }
      continue;
    }
    if (uri == kVersionInfoUri) {
      if (!view16 || rich_version_info || raw_xml(ext) != kVersionInfoExt) {
        return false;
      }
      rich_version_info = true;
      continue;
    }
    if (uri.empty() || std::find(std::begin(kOfficeTableExtUris), std::end(kOfficeTableExtUris), uri) !=
                           std::end(kOfficeTableExtUris)) {
      return false;
    }
  }
  return view16;
}

bool WriteField(const pivot::PivotTable& table, std::size_t index, ExcelHost host, SxField& out) {
  const pivot::PivotField& field = table.fields()[index];
  // Of the x14 field extension only `fillDownLabels` has a record.
  if (!field.x14_attrs.empty() || !field.x14_children_xml.empty() || !field.raw_ext_xml.empty()) {
    return false;
  }
  bool used_as_data = false;
  for (const pivot::PivotDataField& df : table.data_fields()) {
    used_as_data = used_as_data || df.field_index == index;
  }
  const bool measure_filter =
      std::any_of(table.active_filters().begin(), table.active_filters().end(), [&](const pivot::PivotFilter& f) {
        return pivot::is_value_filter_type(f.type) && !detail::ReEmitsSource(table, f) &&
               detail::FilterFieldIndex(table, f) == index;
      });
  bool compact = false;
  bool outline = false;
  if (!FieldForm(field, table.layout(), compact, outline)) {
    return false;
  }
  bool drag_off = false;
  bool row_drag_off = false;
  bool measure_attr = false;
  bool include_new_items = false;
  unsigned drag_seen = 0;
  for (const auto& [name, value] : field.passthrough_attrs) {
    if (name == "dragToRow") {
      if (value != "0" || row_drag_off) {
        return false;
      }
      row_drag_off = true;
    } else if (name == "dragToCol" || name == "dragToPage") {
      const unsigned bit = name == "dragToCol" ? 1U : 2U;
      if (value != "0" || (drag_seen & bit) != 0U) {
        return false;
      }
      drag_seen |= bit;
    } else if (name == "measureFilter") {
      if (value != "1" || measure_attr) {
        return false;
      }
      measure_attr = true;
    } else if (name == "includeNewItemsInFilter") {
      if (value != "1" || include_new_items) {
        return false;
      }
      include_new_items = true;
    } else if (name != "compact" && name != "outline") {
      return false;
    }
  }
  if (drag_seen != 0U && drag_seen != 3U) {
    return false;
  }
  drag_off = drag_seen == 3U;
  switch (field.axis) {
    case pivot::PivotAxis::Row:
      out.axis = kSxFieldAxisRow;
      break;
    case pivot::PivotAxis::Col:
      out.axis = kSxFieldAxisCol;
      break;
    case pivot::PivotAxis::Page:
      out.axis = kSxFieldAxisPage;
      break;
    default:
      out.axis = 0;
      break;
  }
  if (used_as_data) {
    out.axis = static_cast<std::uint8_t>(out.axis | kSxFieldAxisData);
  }
  const std::vector<pivot::SubtotalFn> fns = pivot::ordered_subtotal_fns(field);
  out.subtotals = static_cast<std::uint16_t>(field.default_subtotal && fns.empty() ? kSxFieldSubtotalDefault : 0U);
  for (const pivot::SubtotalFn fn : fns) {
    out.subtotals = static_cast<std::uint16_t>(out.subtotals | (1U << (static_cast<unsigned>(fn) + 1U)));
  }
  out.name_flags = static_cast<std::uint8_t>(compact ? kSxFieldCompact : 0U);
  if (!NumFmt(field.number_format, out.num_fmt)) {
    return false;
  }
  const bool on_axis = field.axis == pivot::PivotAxis::Row || field.axis == pivot::PivotAxis::Col ||
                       field.axis == pivot::PivotAxis::Page;
  const bool auto_sort = (on_axis || field.sort.stated) && !field.sort.manual;
  if (include_new_items && (auto_sort || !field.sort.ascending)) {
    return false;
  }
  out.flags =
      kFieldFixed | (include_new_items ? 0U : kFieldNewItemsOutsideFilter) | (row_drag_off ? 0U : kFieldDragToRow) |
      (drag_off ? 0U : kFieldDragPair) | (field.show_all ? kSxFieldShowAll : 0U) | (outline ? kSxFieldOutline : 0U) |
      (field.insert_blank_row ? kSxFieldInsertBlankRow : 0U) | (field.subtotal_top ? kSxFieldSubtotalTop : 0U) |
      (auto_sort ? kSxFieldAutoSort : 0U) |
      ((auto_sort || include_new_items || host == ExcelHost::kWin365) && field.sort.ascending ? kSxFieldAscendSort
                                                                                              : 0U) |
      (measure_attr || measure_filter ? kSxFieldMeasureFilter : 0U);
  out.item_page_count = kItemPageCount;
  out.reserved16 = kFieldReserved16;
  if (!field.custom_name.empty()) {
    out.display_name = field.custom_name;
  }
  if (!WriteItems(field, out.items) ||
      !WriteAutoSortScope(pivot_field_trailing_children(table, field), out.auto_sort_scope)) {
    return false;
  }
  // SXVD14 is measured after the items of a field without a sort scope.
  if (field.fill_down_labels && (!out.items || !out.auto_sort_scope.empty())) {
    return false;
  }
  out.fill_down_labels = field.fill_down_labels;
  return true;
}

/// The axis list of `order` with the Values pseudo-field at `values_position`, as the xlsx writer spells it.
bool WriteFieldOrder(const std::vector<std::uint32_t>& order, std::optional<std::size_t> values_position,
                     std::size_t field_count, std::optional<std::vector<std::uint32_t>>& out) {
  if (order.empty() && !values_position) {
    return true;
  }
  if (values_position && *values_position > order.size()) {
    return false;
  }
  std::vector<std::uint32_t> list;
  for (std::size_t i = 0; i <= order.size(); ++i) {
    if (values_position && *values_position == i) {
      list.push_back(kValuesField);
    }
    if (i < order.size()) {
      if (order[i] >= field_count) {
        return false;
      }
      list.push_back(order[i]);
    }
  }
  out = std::move(list);
  return true;
}

bool WriteLineItems(std::string_view xml, std::string_view tag, SxLineItems& out) {
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node root = LoadOne(doc, xml, tag);
  const std::optional<std::vector<pugi::xml_node>> lines = root ? Elements(root) : std::nullopt;
  if (!lines || !OnlyAttrs(root, {"count"}) || attr_u32(root, "count", 0U) != lines->size()) {
    return false;
  }
  for (const pugi::xml_node& node : *lines) {
    const std::optional<std::vector<pugi::xml_node>> xs = Elements(node);
    if (std::string_view(node.name()) != "i" || !OnlyAttrs(node, {"t", "r", "i"}) || !xs) {
      return false;
    }
    SxLineItem line;
    if (const pugi::xml_attribute t = node.attribute("t"); t) {
      const auto* it = std::find(std::begin(kLineItemTypes), std::end(kLineItemTypes), std::string_view(t.value()));
      if (it == std::end(kLineItemTypes)) {
        return false;
      }
      line.t = static_cast<std::uint8_t>(it - std::begin(kLineItemTypes));
    }
    const std::uint32_t r = attr_u32(node, "r", 0U);
    if (r > 0xFFFFU) {
      return false;
    }
    line.r = static_cast<std::uint16_t>(r);
    line.i = attr_u32(node, "i", 0U);
    for (const pugi::xml_node& x : *xs) {
      if (std::string_view(x.name()) != "x" || !OnlyAttrs(x, {"v"}) || x.first_child()) {
        return false;
      }
      line.x.push_back(attr_u32(x, "v", 0U));
    }
    out.items.push_back(std::move(line));
  }
  return true;
}

bool WritePages(const pivot::PivotTable& table, PivotTableDoc& doc) {
  const std::vector<std::uint32_t> order = table.page_field_order();
  if (order.empty()) {
    return true;
  }
  std::vector<SxPageField> pages;
  for (const std::uint32_t index : order) {
    const pivot::PivotPageField* entry = table.find_page_field(index);
    const std::optional<std::int32_t> hierarchy = entry != nullptr ? entry->hierarchy : std::optional<std::int32_t>(-1);
    if (!hierarchy || (entry != nullptr && (!entry->name.empty() || !entry->caption.empty()))) {
      return false;
    }
    SxPageField page;
    page.field = index;
    page.item = entry != nullptr && entry->item_index ? *entry->item_index : kNoPageItem;
    page.hierarchy = static_cast<std::uint32_t>(*hierarchy);
    pages.push_back(page);
  }
  doc.page_fields = std::move(pages);
  return true;
}

bool WriteDataFields(const pivot::PivotTable& table, PivotTableDoc& doc) {
  for (const pivot::PivotDataField& in : table.data_fields()) {
    SxDataField out;
    // An x14 display mode is SXDI14 beside a plain SXDI.
    const auto* x14 = std::find(std::begin(kPivotShowAs), std::end(kPivotShowAs), in.show_as);
    const bool is_x14 = x14 != std::end(kPivotShowAs);
    const pivot::ShowValuesAs show_as =
        is_x14 ? pivot::ShowValuesAs::Normal
               : (in.show_as == pivot::ShowValuesAs::RunningTotalInCol ? pivot::ShowValuesAs::RunningTotalInRow
                                                                       : in.show_as);
    const auto* code = std::find(std::begin(kShowDataAs), std::end(kShowDataAs), show_as);
    if (in.field_index >= table.fields().size() || code == std::end(kShowDataAs) ||
        !NumFmt(in.number_format, out.num_fmt) || !in.raw_ext_xml.empty() || !in.x14_attrs.empty() ||
        !in.x14_children_xml.empty()) {
      return false;
    }
    out.field = in.field_index;
    out.subtotal = static_cast<std::uint32_t>(kPivotAggregationNames.size());
    for (std::size_t k = 0; k < kPivotAggregationNames.size(); ++k) {
      if (kPivotAggregationNames[k].value == static_cast<std::uint8_t>(in.aggregation)) {
        out.subtotal = static_cast<std::uint32_t>(k);
        break;
      }
    }
    if (out.subtotal == kPivotAggregationNames.size()) {
      return false;
    }
    out.show_data_as = static_cast<std::uint32_t>(code - std::begin(kShowDataAs));
    out.base_field = in.show_as_base_field.value_or(0U);
    out.base_item = in.show_as_base_item.value_or(0U);
    out.name = in.name;
    if (is_x14) {
      out.pivot_show_as = kSxDataShowAsPercentOfParentRow + static_cast<std::uint32_t>(x14 - std::begin(kPivotShowAs));
    }
    doc.data_fields.push_back(std::move(out));
  }
  return true;
}

/// Splits framed records into raw records; false when they do not frame whole.
bool SplitRecords(const std::vector<std::uint8_t>& bytes, RawRecords& out) {
  ByteSpan cursor{bytes.data(), bytes.size()};
  while (cursor.size > 0U) {
    auto rec = read_record(cursor);
    if (!rec) {
      return false;
    }
    const ByteSpan payload = rec.value().payload;
    out.push_back(
        ids::RawRecord{rec.value().type, std::vector<std::uint8_t>(payload.data, payload.data + payload.size)});
  }
  return true;
}

/// One `<filter>` as its SXFILTER block.
bool WriteFilter(const pugi::xml_node& node, RawRecords& out) {
  const std::optional<std::vector<pugi::xml_node>> children = Elements(node);
  const std::optional<std::uint32_t> fld = parse_xml_u32_attr_strict(node.attribute("fld"));
  const std::optional<std::uint32_t> id = parse_xml_u32_attr_strict(node.attribute("id"));
  const auto* type = std::find(std::begin(kPivotFilterTypes) + 1, std::end(kPivotFilterTypes),
                               std::string_view(attr_str(node, "type")));
  if (std::string_view(node.name()) != "filter" || !fld || !id || type == std::end(kPivotFilterTypes) ||
      !OnlyAttrs(node, {"fld", "mpFld", "type", "evalOrder", "id", "iMeasureHier", "iMeasureFld", "name", "description",
                        "stringValue1", "stringValue2"}) ||
      !children || children->size() != 1U || std::string_view(children->front().name()) != "autoFilter") {
    return false;
  }
  const auto index_attr = [&node](const char* name) {
    const pugi::xml_attribute attr = node.attribute(name);
    return attr ? parse_xml_u32_attr_strict(attr) : std::optional<std::uint32_t>(kNoFilterIndex);
  };
  const std::optional<std::uint32_t> mp_fld = index_attr("mpFld");
  const std::optional<std::uint32_t> measure_fld = index_attr("iMeasureFld");
  const std::optional<std::uint32_t> measure_hier = index_attr("iMeasureHier");
  if (!mp_fld || !measure_fld || !measure_hier) {
    return false;
  }
  std::vector<std::uint8_t> p;
  emit_u32(p, *fld);
  emit_u32(p, *mp_fld);
  emit_u32(p, static_cast<std::uint32_t>(type - std::begin(kPivotFilterTypes)));
  emit_u32(p, static_cast<std::uint32_t>(attr_i32(node, "evalOrder", 0)));
  emit_u32(p, *id);
  emit_u32(p, *measure_fld);
  emit_u32(p, *measure_hier);
  std::uint16_t flags = 0;
  for (std::size_t bit = 0; bit < std::size(kFilterStringAttrs); ++bit) {
    if (node.attribute(std::string(kFilterStringAttrs[bit]).c_str())) {
      flags = static_cast<std::uint16_t>(flags | (1U << bit));
    }
  }
  emit_u16(p, flags);
  for (const std::string_view name : kFilterStringAttrs) {
    if (const pugi::xml_attribute attr = node.attribute(std::string(name).c_str()); attr) {
      emit_xlwidestring(p, attr.value());
    }
  }
  // The xlsx load rule's comparison state is a caption filter's: = and <> match as a pattern, ordering is literal.
  auto criteria = parse_auto_filter_xml(raw_xml(children->front()));
  if (!criteria) {
    return false;
  }
  std::vector<std::uint8_t> afilter;
  if (emit_auto_filter_block(afilter, criteria.value()) != 0U) {
    return false;
  }
  out.push_back(ids::RawRecord{ids::kBeginSXFilter, std::move(p)});
  out.push_back(
      ids::RawRecord{ids::kACBegin, std::vector<std::uint8_t>(kAcUidBeginPayload.begin(), kAcUidBeginPayload.end())});
  out.push_back(ids::RawRecord{ids::kUid, std::vector<std::uint8_t>(kGuidBytes, 0U)});
  out.push_back(ids::RawRecord{ids::kACEnd, {}});
  if (!SplitRecords(afilter, out)) {
    return false;
  }
  out.push_back(ids::RawRecord{ids::kEndSXFilter, {}});
  return true;
}

/// `<filters>` as the xlsx writer would spell it, encoded as SXFILTERS.
bool WriteFilters(const pivot::PivotTable& table, RawRecords& out) {
  std::string xml;
  detail::AppendFilters(xml, table);
  if (xml.empty()) {
    return true;
  }
  pugi::xml_document doc;
  const pugi::xml_node root = LoadOne(doc, xml, "filters");
  const std::optional<std::vector<pugi::xml_node>> filters = root ? Elements(root) : std::nullopt;
  if (!filters) {
    return false;
  }
  std::vector<std::uint8_t> count;
  emit_u32(count, static_cast<std::uint32_t>(filters->size()));
  out.push_back(ids::RawRecord{ids::kBeginSXFilters, std::move(count)});
  for (const pugi::xml_node& node : *filters) {
    if (!WriteFilter(node, out)) {
      return false;
    }
  }
  out.push_back(ids::RawRecord{ids::kEndSXFilters, {}});
  return true;
}

bool WriteLocation(const pivot::PivotTable& table, std::optional<PivotRenderedSpan> rendered_span, SxLocation& loc) {
  // A withheld report is saved as its anchor cell alone, with zero offsets and no page rows or columns.
  if (table.blocked()) {
    if (table.anchor_row() >= a1::kMaxRows || table.anchor_col() >= a1::kMaxCols) {
      return false;
    }
    loc = SxLocation{};
    loc.first_row = loc.last_row = loc.first_header_row = loc.first_data_row = table.anchor_row();
    loc.first_col = loc.last_col = loc.first_data_col = table.anchor_col();
    return true;
  }
  // An offset the model does not carry is derived from the report shape, as the .xlsx writer derives it.
  const detail::LocationDefaults derived = detail::ComputeLocationDefaults(table);
  const std::uint32_t header = table.location_first_header_row().value_or(derived.first_header_row);
  const std::uint32_t data_row = table.location_first_data_row().value_or(derived.first_data_row);
  const std::uint32_t data_col = table.location_first_data_col().value_or(derived.first_data_col);
  const std::optional<std::uint32_t> row_pages =
      table.location_row_page_count() ? table.location_row_page_count() : derived.row_page_count;
  const std::optional<std::uint32_t> col_pages =
      table.location_col_page_count() ? table.location_col_page_count() : derived.col_page_count;
  // The span rule of `write_pivot_table_definition`: a stale read span is
  // replaced by the projection, a caller's span is a floor under it.
  std::uint32_t span_rows = table.span_rows();
  std::uint32_t span_cols = table.span_cols();
  std::uint32_t top = table.anchor_row();
  if (!table.has_authored_span()) {
    if (!rendered_span) {
      return false;
    }
    const bool stale = table.span_read();
    span_rows = !stale && span_rows > rendered_span->rows ? span_rows : rendered_span->rows;
    span_cols = !stale && span_cols > rendered_span->cols ? span_cols : rendered_span->cols;
    top = top > rendered_span->top ? top : rendered_span->top;
  }
  if (span_rows == 0U || span_cols == 0U || top >= a1::kMaxRows || table.anchor_col() >= a1::kMaxCols ||
      span_rows > a1::kMaxRows - top || span_cols > a1::kMaxCols - table.anchor_col() || header >= a1::kMaxRows ||
      data_row >= a1::kMaxRows || data_col >= a1::kMaxCols || row_pages == 0U || col_pages == 0U) {
    return false;
  }
  loc.first_row = top;
  loc.last_row = top + span_rows - 1U;
  loc.first_col = table.anchor_col();
  loc.last_col = table.anchor_col() + span_cols - 1U;
  loc.first_header_row = top + header;
  loc.first_data_row = top + data_row;
  loc.first_data_col = table.anchor_col() + data_col;
  loc.row_page_count = row_pages.value_or(0U);
  loc.col_page_count = col_pages.value_or(0U);
  return true;
}

/// Fills `doc` from `table`; false for a state the records cannot hold or no Excel part shows.
bool WriteTable(const pivot::PivotTable& table, ExcelHost host, std::optional<PivotRenderedSpan> rendered_span,
                const PivotItemsXml* items, PivotTableDoc& doc) {
  if ((table.layout_dirty() && items == nullptr) || !HasExcelDefaults(table) || !table.style_info() ||
      static_cast<int>(table.grand_total_caption().has_value()) + static_cast<int>(table.error_caption().has_value()) +
              static_cast<int>(table.missing_caption().has_value()) >
          1) {
    return false;
  }
  // Untyped `<formats>` and its siblings have no records here; the table
  // extLst holds SXView14 / SXView16 and extensions Excel's xlsb save drops.
  // A table with no captured extLst takes the xlsx writer's; Excel adds SXView16 to it.
  bool view14 = table_hides_values_row(table);
  bool rich_version_info = false;
  if (!table.raw_passthrough_before_style_info().empty() ||
      (!table.raw_passthrough_xml().empty() &&
       (!WriteTableExtLst(table.raw_passthrough_xml(), view14, rich_version_info) ||
        table.hide_values_row().value_or(view14) != view14)) ||
      !WriteFormats(table, doc)) {
    return false;
  }
  FormOverride form;
  if (!WriteRootAttrs(table, doc, form)) {
    return false;
  }
  const bool form_override = form.compact || form.compact_data || form.outline || form.outline_data;
  if (form_override && (!form.compact || !form.compact_data || !form.outline || !form.outline_data ||
                        *form.compact != *form.compact_data || *form.outline != *form.outline_data)) {
    return false;
  }
  const bool compact = form_override ? *form.compact : table.layout() == pivot::PivotLayout::Compact;
  const bool outline = form_override ? *form.outline : table.layout() != pivot::PivotLayout::Tabular;
  // The view's row outline bit is set when any row field is in outline form.
  bool row_outline = false;
  for (const std::uint32_t index : table.row_field_order()) {
    bool field_compact = false;
    bool field_outline = false;
    if (index >= table.fields().size() ||
        !FieldForm(table.fields()[index], table.layout(), field_compact, field_outline)) {
      return false;
    }
    row_outline = row_outline || field_outline;
  }
  if (table.row_field_order().empty() && table.row_values_position()) {
    return false;
  }

  SxView& v = doc.view;
  v.flags1 = kViewFlags1;
  v.flags2 = kViewFlags2;
  v.flags3 = static_cast<std::uint8_t>(table.indent() | (table.show_headers() ? 0U : kSxViewNoHeaders));
  v.flags4 = static_cast<std::uint8_t>(kViewFlags4 | (v.flags4 & kViewAsteriskTotals) |
                                       (table.show_empty_row() ? kViewShowEmptyRow : 0U) |
                                       (table.show_empty_col() ? kViewShowEmptyCol : 0U));
  v.flags5 = static_cast<std::uint8_t>(
      kViewFlags5Fixed | (table.auto_formatting() ? kSxViewAutoFormat : 0U) |
      (table.show_error() ? kSxViewShowError : 0U) | (table.grand_totals_cols() ? kSxViewRowGrand : 0U) |
      (table.grand_totals_rows() ? kSxViewColGrand : 0U) | (table.page_over_then_down() ? kViewPageOverThenDown : 0U) |
      (table.subtotal_hidden_items() ? kViewSubtotalHiddenItems : 0U));
  v.flags13 = static_cast<std::uint8_t>(table.page_wrap());
  v.flags6 = static_cast<std::uint8_t>(kViewFlags6Fixed | (table.merge_item() ? kSxViewMergeItem : 0U) |
                                       (row_outline ? kViewRowOutline : 0U));
  v.flags8 = static_cast<std::uint8_t>((compact ? kViewCompactPair : 0U) | (outline ? kViewOutlinePair : 0U) |
                                       (table.grid_drop_zones() ? 0U : kSxViewNoDropZones));
  v.flags9 = static_cast<std::uint8_t>(table.multiple_field_filters() ? 0U : kSxViewSingleFieldFilter);
  v.data_axis = table.row_values_position() ? kDataOnRows : kDataOnCols;
  v.data_position = kNoDataPosition;
  v.reserved20 = kViewReserved20;
  v.cache_id = table.pivot_cache_id();
  v.name = table.name();
  v.data_caption = table.data_caption();
  v.grand_total_caption = table.grand_total_caption();
  v.error_caption = table.error_caption();
  v.missing_caption = table.missing_caption();

  if (!WriteLocation(table, rendered_span, doc.location)) {
    return false;
  }
  doc.fields.resize(table.fields().size());
  for (std::size_t i = 0; i < table.fields().size(); ++i) {
    if (!WriteField(table, i, host, doc.fields[i])) {
      return false;
    }
  }
  const std::size_t field_count = table.fields().size();
  if (!WriteFieldOrder(table.row_field_order(), table.row_values_position(), field_count, doc.row_fields) ||
      !WriteFieldOrder(table.col_field_order(), table.col_values_position(), field_count, doc.col_fields) ||
      !WriteLineItems(items != nullptr ? items->row_items : table.raw_passthrough_after_row_fields(), "rowItems",
                      doc.row_items) ||
      !WriteLineItems(items != nullptr ? items->col_items : table.raw_col_items(), "colItems", doc.col_items) ||
      !WritePages(table, doc) || !WriteDataFields(table, doc) || !WriteFilters(table, doc.filters)) {
    return false;
  }
  const pivot::PivotStyleInfo& style = *table.style_info();
  doc.style.flags = static_cast<std::uint16_t>(
      (style.show_last_column ? kStyleLastColumn : 0U) | (style.show_row_stripes ? kStyleRowStripes : 0U) |
      (style.show_col_stripes ? kStyleColStripes : 0U) | (style.show_row_headers ? kStyleRowHeaders : 0U) |
      (style.show_col_headers ? kStyleColHeaders : 0U) | (host == ExcelHost::kWin365 ? kStyleFirstColumn : 0U));
  if (!style.name.empty()) {
    doc.style.name = style.name;
  }
  if (view14) {
    doc.view14 = SxFrtBlock{{kView14Frt.begin(), kView14Frt.end()}, {kView14Payload.begin(), kView14Payload.end()}};
  }
  doc.view16.frt_payload.assign(kView16Frt.begin(), kView16Frt.end());
  doc.view16.payload.assign(kView16Payload.begin(), kView16Payload.end());
  if (rich_version_info) {
    doc.version_info =
        SxFrtBlock{{pivot_records::kVersionInfoFrtVersion.begin(), pivot_records::kVersionInfoFrtVersion.end()},
                   {pivot_records::kRichDataFeature.begin(), pivot_records::kRichDataFeature.end()}};
  }
  return true;
}

}  // namespace

Expected<pivot::PivotTable, Error> pivot_table_from_xlsb(ByteSpan part) {
  ASSIGN_OR_RETURN(const PivotTableDoc doc, decode_pivot_table(part));
  RETURN_IF_ERROR(CheckView(doc));
  const SxView& view = doc.view;
  pivot::PivotTable table;
  table.set_definition_source(pivot::PivotDefinitionSource::kXlsb);
  table.set_name(view.name);
  table.set_data_caption(view.data_caption);
  table.set_pivot_cache_id(view.cache_id);
  // `rowGrandTotals` totals the rows into a bottom line, the model's column totals.
  table.set_grand_totals(/*rows=*/(view.flags5 & kSxViewColGrand) != 0U,
                         /*cols=*/(view.flags5 & kSxViewRowGrand) != 0U);
  const bool root_compact = (view.flags8 & kViewCompactPair) != 0U;
  const bool root_outline = (view.flags8 & kViewOutlinePair) != 0U;
  table.set_layout(FormOf(root_compact, root_outline));
  table.set_grand_total_caption(view.grand_total_caption);
  table.set_auto_formatting((view.flags5 & kSxViewAutoFormat) != 0U);
  table.set_show_error((view.flags5 & kSxViewShowError) != 0U);
  table.set_error_caption(view.error_caption);
  table.set_missing_caption(view.missing_caption);
  table.set_show_headers((view.flags3 & kSxViewNoHeaders) == 0U);
  table.set_grid_drop_zones((view.flags8 & kSxViewNoDropZones) == 0U);
  table.set_item_print_titles(true);
  table.set_multiple_field_filters((view.flags9 & kSxViewSingleFieldFilter) == 0U);
  table.set_merge_item((view.flags6 & kSxViewMergeItem) != 0U);
  table.set_show_empty_row((view.flags4 & kViewShowEmptyRow) != 0U);
  table.set_show_empty_col((view.flags4 & kViewShowEmptyCol) != 0U);
  table.set_page_over_then_down((view.flags5 & kViewPageOverThenDown) != 0U);
  table.set_subtotal_hidden_items((view.flags5 & kViewSubtotalHiddenItems) != 0U);
  table.set_page_wrap(view.flags13);
  table.set_indent(static_cast<std::uint32_t>(view.flags3 & kSxViewIndentMask));
  ReadRootAttrs(doc, table.mutable_passthrough_attrs());
  RETURN_IF_ERROR(ReadLocation(doc.location, table));

  for (const SxField& in : doc.fields) {
    ASSIGN_OR_RETURN(pivot::PivotField field, ReadField(in));
    table.mutable_fields().push_back(std::move(field));
  }
  const std::size_t field_count = doc.fields.size();
  std::optional<std::size_t> row_values;
  RETURN_IF_ERROR(ReadFieldOrder(doc.row_fields, field_count, table.mutable_row_field_order(), row_values));
  table.set_row_values_position(row_values);
  // The first row field's own form decides the table's, as in the xlsx reader; the view's row outline bit is set
  // when any row field is in outline form.
  bool row_outline = false;
  if (!table.row_field_order().empty()) {
    const SxField& first = doc.fields[table.row_field_order().front()];
    table.set_layout(FormOf((first.name_flags & kSxFieldCompact) != 0U, (first.flags & kSxFieldOutline) != 0U));
    for (const std::uint32_t index : table.row_field_order()) {
      row_outline = row_outline || (doc.fields[index].flags & kSxFieldOutline) != 0U;
    }
  }
  if (((view.flags6 & kViewRowOutline) != 0U) != row_outline ||
      (view.data_axis == kDataOnRows) != row_values.has_value()) {
    return Unmeasured("row form or data axis disagreeing with the row fields");
  }
  CaptureForm(root_compact, root_outline, table.layout(), /*root=*/true, table.mutable_passthrough_attrs());
  if (view.reserved24 != 0U) {
    table.mutable_passthrough_attrs().emplace_back("chartFormat", std::to_string(view.reserved24));
  }
  for (std::size_t i = 0; i < field_count; ++i) {
    CaptureForm((doc.fields[i].name_flags & kSxFieldCompact) != 0U, (doc.fields[i].flags & kSxFieldOutline) != 0U,
                table.layout(), /*root=*/false, table.mutable_fields()[i].passthrough_attrs);
  }
  std::optional<std::size_t> col_values;
  RETURN_IF_ERROR(ReadFieldOrder(doc.col_fields, field_count, table.mutable_col_field_order(), col_values));
  table.set_col_values_position(col_values);
  table.mutable_raw_passthrough_after_row_fields() = LineItemsXml(doc.row_items, "rowItems");
  table.mutable_raw_col_items() = LineItemsXml(doc.col_items, "colItems");

  if (doc.page_fields) {
    for (const SxPageField& in : *doc.page_fields) {
      if (in.field >= field_count) {
        return Corrupt("page field names a field the table does not have");
      }
      pivot::PivotPageField entry;
      entry.field_index = in.field;
      entry.item_index = in.item == kNoPageItem ? std::nullopt : std::optional<std::uint32_t>(in.item);
      entry.hierarchy = static_cast<std::int32_t>(in.hierarchy);
      table.mutable_page_fields().push_back(std::move(entry));
    }
  }
  RETURN_IF_ERROR(ReadDataFields(doc, table));
  RETURN_IF_ERROR(CheckAxes(doc, table));
  RETURN_IF_ERROR(ReadStyle(doc.style, table));
  ReadFormats(doc, table);
  if (!doc.filters.empty()) {
    ASSIGN_OR_RETURN(const std::string filters, FiltersXml(doc.filters, field_count, doc.data_fields.size()));
    if (!read_pivot_filters_xml(filters, &table)) {
      return Corrupt("filters do not spell as OOXML");
    }
  }
  table.set_hide_values_row(doc.view14.has_value());
  table.mutable_raw_passthrough_xml() = TableExtLst(doc.view14.has_value(), doc.version_info.has_value());
  lift_pivot_sort_scopes(table);
  return table;
}

Expected<std::vector<std::uint8_t>, Error> pivot_table_to_xlsb(const pivot::PivotTable& table, ExcelHost host,
                                                               std::size_t sheet_index,
                                                               std::optional<PivotRenderedSpan> rendered_span,
                                                               const PivotItemsXml* items) {
  PivotTableDoc doc;
  if (!WriteTable(table, host, rendered_span, items, doc)) {
    return not_writable(kNotWritableKind, sheet_index);
  }
  return encode_pivot_table(doc);
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon
