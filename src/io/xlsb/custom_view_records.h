//
// Custom views between their `.xlsx` and `.xlsb` forms. A workbook's
// `<customWorkbookView>` is one BrtUserBookView (397); a sheet's
// `<customSheetView>` is a BrtBeginUserShView (423) .. BrtEndUserShView (424)
// pair holding the view's pane, selection, print and AutoFilter records, inside the
// sheet's BrtBeginUserShViews (422) .. BrtEndUserShViews (425) block.
// Field layouts, defaults and child order are as Excel writes them. A view
// carrying anything else is refused whole, so that the caller can leave out
// both sides of it: Excel does not open a sheet view without its workbook
// view. Both directions are here: the `.xlsx` form encoded for an `.xlsb`
// save, and the records decoded for an `.xlsx` save.

#ifndef FORMULON_IO_XLSB_CUSTOM_VIEW_RECORDS_H_
#define FORMULON_IO_XLSB_CUSTOM_VIEW_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/record.h"
#include "pugixml.hpp"

namespace formulon {
namespace io {
namespace xlsb {

/// Appends the BrtUserBookView of a `<customWorkbookView>`. False, appending
/// nothing, for an attribute or value the record does not hold.
bool encode_book_view(const pugi::xml_node& view, std::vector<std::uint8_t>& dst);

/// Appends the BrtBeginUserShView .. BrtEndUserShView records of a
/// `<customSheetView>`. Without `print_settings` (its workbook view's
/// `includePrintSettings="0"`) the view keeps no print records, as Excel
/// writes it. A view's `<autoFilter>` is written last, as the sheet's AutoFilter
/// block. False, appending nothing, for an attribute, child or value the
/// records do not hold, an AutoFilter with a setting the block cannot hold among them.
bool encode_sheet_view(const pugi::xml_node& view, bool print_settings, std::vector<std::uint8_t>& dst);

/// The `<customWorkbookView>` of a BrtUserBookView payload, attributes at
/// their defaults left out as Excel writes them; nullopt for a payload not in
/// the layout Excel writes.
std::optional<std::string> decode_book_view(ByteSpan payload);

/// The `<customSheetView>` of one view's records, from its BrtBeginUserShView
/// through its BrtEndUserShView in `records` (framed in `body`); nullopt for
/// a view holding a record or bit Excel was not seen to write, or a print
/// record the `.xlsx` form here cannot state exactly (an automatic page break
/// among them). The filter's revision-uid wrapper is not carried over.
std::optional<std::string> decode_sheet_view(const std::vector<FramedRecord>& records,
                                             const std::vector<std::uint8_t>& body);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CUSTOM_VIEW_RECORDS_H_
