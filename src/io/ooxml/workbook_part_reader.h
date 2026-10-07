// Workbook-part settings read from `<workbook>`: the calc options and the
// elements captured raw for verbatim re-emission.

#ifndef FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_
#define FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_

#include "pugixml.hpp"

namespace formulon {

class Workbook;

namespace io {
namespace ooxml {

/// Applies `<calcPr>` and the raw-captured workbook elements
/// (`<fileVersion>`, `<fileSharing>`, `<workbookPr>`, `<workbookProtection>`,
/// `<bookViews>`, `<extLst>`) of `wb_root` to `wb`.
void apply_workbook_part_settings(const pugi::xml_node& wb_root, Workbook& wb);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_
