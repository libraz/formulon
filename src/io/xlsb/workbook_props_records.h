//
// MS-XLSB workbook property records and the model state they carry:
//
//   * BrtWbProp (153, [MS-XLSB] 2.4.868): a u32 flag word, dwThemeVersion and
//     the CodeName. Maps to `<workbookPr>`; the model keeps that element
//     verbatim (`Workbook::workbook_pr_xml`) with `date1904` mirrored into
//     `Workbook::date1904`.
//   * BrtCalcProp (157, [MS-XLSB] 2.4.318): recalcID, fAutoRecalc,
//     cCalcCount, xnumDelta, cUserThreadCount and a u16 flag word. Maps to
//     `<calcPr>`: `CalcMode`, `IterativeOptions` and `CalcOptions`.
//   * BrtBookView (158, 2.4.313) inside BrtBeginBookViews (135) /
//     BrtEndBookViews (136): window position and size, iTabRatio, itabFirst,
//     itabCur and a flag byte. One per `<workbookView>` of the retained
//     `<bookViews>` (`Workbook::book_views_xml`).
//
// `<workbookPr dateCompatibility allowRefreshQuery>` have no BrtWbProp field.

#ifndef FORMULON_IO_XLSB_WORKBOOK_PROPS_RECORDS_H_
#define FORMULON_IO_XLSB_WORKBOOK_PROPS_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <vector>

#include "calc_settings.h"
#include "io/zip_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtCalcProp = 157;
inline constexpr std::uint16_t kBrtBookView = 158;

/// The `<calcPr>` state one BrtCalcProp carries.
struct XlsbCalcProps {
  CalcMode mode = CalcMode::kAuto;
  IterativeOptions iterative;
  CalcOptions options;
};

/// Decodes a BrtCalcProp payload. A recalcID of 0 reads as an absent
/// `calcId`; the iteration count is clamped like the OOXML reader's.
/// Returns false, leaving `out` untouched, on a truncated payload.
bool decode_calc_prop(ByteSpan payload, XlsbCalcProps& out);

/// True when `props` holds the schema defaults, i.e. `<calcPr>` is omitted.
bool calc_props_are_default(const XlsbCalcProps& props);

/// Emits BrtCalcProp for `props`. An absent `calcId` writes recalcID 0,
/// which asks the reader's application for a full recalculation on load.
void emit_calc_prop(std::vector<std::uint8_t>& dst, const XlsbCalcProps& props);

/// Decodes a BrtWbProp payload into a `<workbookPr>` element (empty when
/// every field holds its schema default) and the 1904 flag. Returns false when
/// even the flag word is truncated.
bool decode_wb_prop(ByteSpan payload, std::string& workbook_pr_xml, bool& date1904);

/// Emits BrtWbProp from a `<workbookPr>` element; `date1904` overrides the
/// element's own attribute. An empty element writes the default flags and
/// theme version 202300. Returns false when the element carries an
/// attribute BrtWbProp has no field for (the record is still written).
bool emit_wb_prop(std::vector<std::uint8_t>& dst, const std::string& workbook_pr_xml, bool date1904);

/// Decodes a BrtBookView payload and appends its `<workbookView>` element
/// to `xml`. Returns false on a truncated payload.
bool decode_book_view(ByteSpan payload, std::string& xml);

/// Windows the workbook part writes for `book_views_xml`: one per
/// `<workbookView>`, and one when there is none.
std::uint32_t workbook_view_count(const std::string& book_views_xml);

/// Emits the BOOKVIEWS block: one BrtBookView per `<workbookView>`, or the
/// default window when there is none. The first window's itabCur is
/// `active_sheet`; a sheet index past `sheet_count` writes 0. Returns false
/// when a view carries something BrtBookView has no field for (`<extLst>`,
/// an unknown attribute); the revision `xr2:uid` is not counted.
bool emit_book_views(std::vector<std::uint8_t>& dst, const std::string& book_views_xml, std::uint32_t active_sheet,
                     std::size_t sheet_count);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_WORKBOOK_PROPS_RECORDS_H_
