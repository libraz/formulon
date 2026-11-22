// Declarations shared between the XLSB reader's package pipeline
// (`reader.cpp`) and its external-link loading (`reader_external_links.cpp`).

#ifndef FORMULON_IO_XLSB_READER_INTERNAL_H_
#define FORMULON_IO_XLSB_READER_INTERNAL_H_

#include <vector>

#include "io/xlsb/ptg_reader.h"
#include "io/xlsb/workbook_bin_reader.h"
#include "io/zip_reader.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {
namespace reader_detail {

/// Renumbers the external books by first appearance in the `BrtExternSheet` list; books no entry
/// names follow in stored order. `books` and `ranges` are rewritten to the same new numbers.
void NumberExternalBooksByFirstUse(std::vector<XlsbSupBook>& books, std::vector<XlsbSheetRange>& ranges);

/// Loads every external link part the supporting-book list names, in `[N]` order, into `wb`.
void LoadExternalLinkParts(const ZipReader& zip, Workbook& wb, const std::vector<XlsbSupBook>& books);

/// Projects the loaded external links onto the two name tables `decode_ptgs` consults.
XlsbExternalBooks CollectExternalBookTables(const Workbook& wb);

}  // namespace reader_detail
}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_READER_INTERNAL_H_
