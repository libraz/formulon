//
// One `<sheetView>` element to and from `SheetView`, shared by the OOXML
// reader and writer and by the XLSB codec of a sheet's extra windows.

#ifndef FORMULON_IO_SHEET_VIEW_XML_H_
#define FORMULON_IO_SHEET_VIEW_XML_H_

#include <cstdint>
#include <string>

#include "pugixml.hpp"
#include "sheet.h"

namespace formulon {
namespace io {

/// Folds a `<sheetView>` element's attributes, `<pane>`, `<selection>`s and
/// raw `<extLst>` into `view`, an unknown `view` mode read as normal.
/// Elements after it are not read.
void read_sheet_view_element(const pugi::xml_node& sheet_view, SheetView& view);

/// The `<sheetView>` element of `view` for workbook window
/// `workbook_view_id`, attributes at their schema defaults omitted.
/// `view.extra_views_xml` is not part of it.
std::string build_sheet_view_element_xml(const SheetView& view, std::uint32_t workbook_view_id);

/// The `<pane>` and `<selection>` children of `view`'s element, as
/// `build_sheet_view_element_xml` writes them; with `listed_only`, only the
/// selections of panes `view.selections` names, as a custom view holds them.
std::string build_pane_and_selections_xml(const SheetView& view, bool listed_only);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_VIEW_XML_H_
