//
// Worksheet view and layout reader: viewport state, sheet format
// defaults, `<cols>` spans and `<row>`-level overrides.

#ifndef FORMULON_IO_SHEET_LAYOUT_READER_H_
#define FORMULON_IO_SHEET_LAYOUT_READER_H_

#include <cstddef>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Reads non-cell worksheet metadata from the parsed `sheet*.xml`
/// document and writes it onto `workbook.sheet(sheet_index)`. Currently
/// covers the viewport state mirrored by `Sheet::view()` (zoom scale,
/// frozen panes, tab visibility) and the layout overrides mirrored by
/// `Sheet::layout()` (`<cols>` spans, `<row>`-level height / hidden /
/// outline overrides).
///
/// Behaviour:
///   * `<sheetView zoomScale="...">` — clamped to `[10, 400]`, defaulting
///     to `SheetView::kDefaultZoomScale` when absent or out of range.
///   * `<sheetView><pane state="frozen" xSplit ySplit/></sheetView>` —
///     populates `freeze_rows` / `freeze_cols`. `state="frozenSplit"`
///     is equally frozen and is read the same way; the model does not
///     carry the split position, so the writer re-emits both spellings
///     as `state="frozen"`. A `state="split"` pane leaves both at `0`.
///   * `<sheetPr><tabHidden val="1"/></sheetPr>` — sets
///     `view().tab_hidden`. The workbook-side `<sheet state="hidden">`
///     path (handled in `read_ooxml`) merges OR-style: either signal
///     marks the sheet as hidden.
///   * `<cols>/<col min max width style hidden outlineLevel/>` — appended to
///     `layout().columns`. Width/style attribute presence is preserved,
///     including explicit zero values; valid hidden/outline-only spans are
///     retained. A `customWidth=1` / `bestFit=1` marker without any stored
///     layout state remains a no-op for the model.
///   * `<row r ht hidden outlineLevel s customFormat ...>` — appended to
///     `layout().row_overrides` when a metric/visibility/outline override is
///     present or when `customFormat=1`; row `s` is ignored otherwise.
///
/// Returns `kIoSheetCorrupt` when the worksheet root is malformed.
/// Missing optional sub-elements are not errors.
Expected<void, Error> read_sheet_view_and_layout(const pugi::xml_document& sheet_doc, std::size_t sheet_index,
                                                 Workbook& workbook);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_LAYOUT_READER_H_
