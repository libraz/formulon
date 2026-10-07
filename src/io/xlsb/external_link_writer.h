//
// MS-XLSB external-link encoding: one `xl/externalLinks/externalLink<N>.bin`
// part and its rels per link the workbook records, generated from the model
// rather than carried from the source package. The record layout is the one
// `external_link_reader.cpp` decodes, reproduced byte for byte against
// Excel-written parts.
//
// Every workbook link record is written, in `ExternalLinkRecord::index`
// order (OLE and DDE links are not), and its
// 1-based position in that order is the `<N>` of its part and the position
// of its `BrtSupBookSrc` in `xl/workbook.bin`. Record indices may have gaps;
// positions never do.
//
// Defined names are written book-scope only: the part's name records carry
// no scope this writer has measured.

#ifndef FORMULON_IO_XLSB_EXTERNAL_LINK_WRITER_H_
#define FORMULON_IO_XLSB_EXTERNAL_LINK_WRITER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "external_link.h"
#include "io/xlsb/ptg_writer.h"
#include "workbook.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The workbook's external book links in the order a save writes them.
std::vector<const ExternalLinkRecord*> written_external_links(const Workbook& wb);

/// Package path of the part written at 1-based `position`.
std::string external_link_part_path(std::size_t position);

/// Encodes `link`'s part. `tables` is its entry in the workbook's
/// `SheetRangeTable`, whose `names` decide each name's `ilbl`.
std::vector<std::uint8_t> build_external_link_bin(const ExternalLinkRecord& link, const XlsbLinkTables& tables);

/// The part's rels: the target relationship and, when the link records one,
/// the absolute URL beside it.
std::string build_external_link_rels(const ExternalLinkRecord& link);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_EXTERNAL_LINK_WRITER_H_
