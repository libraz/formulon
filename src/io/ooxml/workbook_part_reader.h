// Workbook-part settings read from `<workbook>`: the calc options and the
// elements captured raw for verbatim re-emission.

#ifndef FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_
#define FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_

#include <cstddef>
#include <string_view>

#include "pugixml.hpp"

namespace formulon {

class Workbook;

namespace io {
namespace ooxml {

/// The ECMA-376 `CT_Workbook` child sequence. Index into `kOrder` is the slot
/// a retained `Workbook::raw_children()` entry records.
namespace workbook_child {

inline constexpr std::string_view kOrder[] = {
    "fileVersion",    "fileSharing",
    "workbookPr",     "workbookProtection",
    "bookViews",      "sheets",
    "functionGroups", "externalReferences",
    "definedNames",   "calcPr",
    "oleSize",        "customWorkbookViews",
    "pivotCaches",    "smartTagPr",
    "smartTagTypes",  "webPublishing",
    "fileRecoveryPr", "webPublishObjects",
    "extLst",
};

/// Number of names in `kOrder`, and the value `slot_of` returns for a name
/// outside the sequence.
inline constexpr std::size_t kCount = sizeof(kOrder) / sizeof(kOrder[0]);

/// Slot of `name` in `kOrder`, or `kCount` when the schema has no such child.
inline std::size_t slot_of(std::string_view name) noexcept {
  for (std::size_t i = 0; i < kCount; ++i) {
    if (kOrder[i] == name) {
      return i;
    }
  }
  return kCount;
}

}  // namespace workbook_child

/// Applies `<calcPr>` and the raw-captured workbook elements
/// (`<fileVersion>`, `<fileSharing>`, `<workbookPr>`, `<workbookProtection>`,
/// `<bookViews>`, `<extLst>`) of `wb_root` to `wb`, and retains every other
/// child the model does not fold in, with its slot, in `raw_children()`.
void apply_workbook_part_settings(const pugi::xml_node& wb_root, Workbook& wb);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKBOOK_PART_READER_H_
