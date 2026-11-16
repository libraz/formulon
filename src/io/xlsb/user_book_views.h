//
// The custom workbook views of a `xl/workbook.bin` body: one BrtUserBookView
// (397) per view, after BrtCalcProp and before BrtFileRecover / BrtEndBook.
// The record names its view by a 16-byte GUID at payload offset 24; each sheet
// carrying the view holds a BrtBeginUserShView (423) naming the same GUID at
// payload offset 0 in its custom-view block (422..425). Excel refuses a sheet
// custom view whose workbook view is missing. An `.xlsb` source's records are
// kept as `Workbook::xlsb_user_book_views()` and re-emitted verbatim, each only
// while a written sheet still carries its sheet side. An `.xlsx` source's
// views are encoded at save time (`plan_xml_custom_views`).

#ifndef FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_
#define FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_

#include <cstddef>
#include <cstdint>
#include <string>
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

/// The custom views an `.xlsx` source kept as raw XML, encoded for an
/// `.xlsb` save. A view goes out only when its workbook view and every sheet
/// view naming its GUID encode, and at least one sheet carries it; otherwise
/// both sides are left out. Empty for a workbook that kept BrtUserBookView
/// records, which are written instead.
struct XmlCustomViewPlan {
  /// The BrtUserBookView records, in document order.
  std::vector<std::uint8_t> book_records;
  /// Per sheet, the BrtBeginUserShViews .. BrtEndUserShViews block; empty
  /// when the sheet writes none.
  std::vector<std::vector<std::uint8_t>> sheet_blocks;
  /// Workbook views left out with their sheet views.
  std::size_t refused_book_views = 0;
  /// Per sheet, sheet views naming no workbook view, which are left out.
  std::vector<std::size_t> orphan_sheet_views;
};

XmlCustomViewPlan plan_xml_custom_views(const Workbook& workbook);

/// The custom views an `.xlsb` source kept as records, decoded for an
/// `.xlsx` save, on the same terms as `XmlCustomViewPlan`. Empty for a
/// workbook that kept `<customWorkbookViews>` XML, which is written instead.
struct RecordCustomViewPlan {
  /// The `<customWorkbookViews>` element; empty when none goes out.
  std::string book_xml;
  /// Per sheet, the `<customSheetViews>` element; empty when none goes out.
  std::vector<std::string> sheet_xml;
  /// Workbook views left out with their sheet views.
  std::size_t refused_book_views = 0;
};

RecordCustomViewPlan plan_record_custom_views(const Workbook& workbook);

/// Splices the BrtUserBookView records into the generated `workbook_bin`
/// ahead of BrtFileRecover, the future-record blocks and BrtEndBook:
/// `workbook.xlsb_user_book_views()` when it holds any, each only while a
/// sheet's retained tail names it, else `generated`. Returns how many
/// retained views were left out.
std::size_t splice_book_user_views(std::vector<std::uint8_t>& workbook_bin, const Workbook& workbook,
                                   const std::vector<std::uint8_t>& generated);

/// `<customWorkbookView>` elements an `.xlsx` source kept as a raw
/// `<customWorkbookViews>` child.
std::size_t xml_book_view_count(const Workbook& workbook);

/// `<customSheetView>` elements `sheet` kept as a raw `<customSheetViews>`
/// child.
std::size_t xml_sheet_view_count(const Sheet& sheet);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_USER_BOOK_VIEWS_H_
