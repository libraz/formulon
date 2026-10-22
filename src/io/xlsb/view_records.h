// Worksheet-view codec: BrtBeginWsView (137) / BrtPane (151) / BrtSel (152) /
// BrtEndWsView (138) to and from `SheetView`, for the modelled view and the
// extra windows kept as `SheetView::extra_views_xml`.

#ifndef FORMULON_IO_XLSB_VIEW_RECORDS_H_
#define FORMULON_IO_XLSB_VIEW_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

struct SheetView;

namespace io {
namespace xlsb {

/// Decodes a BrtBeginWsView into `view`'s display fields, scroll origin,
/// zooms and view mode, and its iWbkView into `*workbook_view`.
Expected<void, Error> decode_ws_view(const XlsbRecord& rec, SheetView& view, std::uint32_t* workbook_view);

/// Emits BrtBeginWsView, BrtPane, the BrtSels and BrtEndWsView of `view`
/// shown in workbook window `workbook_view`.
void emit_ws_view(std::vector<std::uint8_t>& dst, const SheetView& view, std::uint32_t workbook_view);

/// Emits one view per `<sheetView>` of `view.extra_views_xml` whose window
/// is among the `workbook_views` BrtBookView records the workbook part
/// writes, in document order.
void emit_extra_ws_views(std::vector<std::uint8_t>& dst, const SheetView& view, std::uint32_t workbook_views = 1U);

/// Appends the `<sheetView>` of an extra window decoded from XLSB to
/// `main.extra_views_xml`.
void append_extra_view(SheetView& main, const SheetView& extra, std::uint32_t workbook_view);

/// Parts of `view` and its extra windows the view records cannot carry: a
/// `<pivotSelection>` or view `<extLst>`, an unknown view mode, and a window
/// outside the `workbook_views` the workbook part writes.
std::size_t unwritten_view_part_count(const SheetView& view, std::uint32_t workbook_views = 1U);

/// Decodes one BrtPane into `view`'s freeze or split fields. False for an
/// unfrozen pane with no valid split; an error when a frozen pane's counts
/// are not in-grid whole numbers.
Expected<bool, Error> decode_pane(const XlsbRecord& rec, SheetView& view);

/// Decodes one BrtSel and appends it to `view.selections`, an A1 active cell
/// or sqref left empty and the top-left pane's A1 selection not appended,
/// as Excel's XLSX spells them. False, appending nothing, when the record is
/// truncated, inconsistent or leaves the grid.
bool decode_selection(const XlsbRecord& rec, SheetView& view);

/// Emits the BrtPane of a frozen or split view, then one BrtSel per pane
/// the view has, as `resolve_view_selections` orders and fills them.
void emit_pane_and_selections(std::vector<std::uint8_t>& dst, const SheetView& view);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_VIEW_RECORDS_H_
