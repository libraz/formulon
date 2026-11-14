//
// The custom workbook views of a `xl/workbook.bin` body: one BrtUserBookView
// (397) per view, after BrtCalcProp and before BrtFileRecover / BrtEndBook.
// The record names its view by a 16-byte GUID at payload offset 24; each sheet
// carrying the view holds a BrtBeginUserShView (423) naming the same GUID at
// payload offset 0 in its custom-view block (422..425). Excel refuses a sheet
// custom view whose workbook view is missing. An `.xlsb` source's records are
// kept as `Workbook::xlsb_user_book_views()` and re-emitted verbatim, each only
// while a written sheet still carries its sheet side.

#ifndef FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_
#define FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/zip_reader.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The framed BrtUserBookView records of a `xl/workbook.bin` body, in order;
/// empty when the body has none.
std::vector<std::uint8_t> read_book_user_views(ByteSpan workbook_bin);

/// Splices `workbook.xlsb_user_book_views()` into the generated `workbook_bin`
/// ahead of BrtFileRecover, the future-record blocks and BrtEndBook. A view no
/// sheet's retained tail names is left out; returns how many were.
std::size_t splice_book_user_views(std::vector<std::uint8_t>& workbook_bin, const Workbook& workbook);

/// `<customWorkbookView>` elements an `.xlsx` source kept as a raw
/// `<customWorkbookViews>` child. No record here is written from them, so
/// they and every sheet's `<customSheetViews>` are left out together.
std::size_t unwritten_xml_book_views(const Workbook& workbook);

/// `<customSheetView>` elements `sheet` kept as a raw `<customSheetViews>`
/// child; left out of an `.xlsb` save with the workbook views.
std::size_t unwritten_xml_sheet_views(const Sheet& sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_
