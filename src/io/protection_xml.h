//
// Conversion between the protection records in `protection.h` and the XML
// the model retains for them.

#ifndef FORMULON_IO_PROTECTION_XML_H_
#define FORMULON_IO_PROTECTION_XML_H_

#include <string>
#include <string_view>
#include <vector>

#include "protection.h"
#include "sheet_passthrough.h"

namespace formulon::io {

/// Decodes a `<workbookProtection>` element; empty or unparseable XML yields
/// an all-default record.
WorkbookProtection parse_workbook_protection(std::string_view xml);

/// `existing` with the modelled attributes replaced by `protection`;
/// attributes the record does not model survive. Empty when the result would
/// carry no attribute at all (the element is dropped).
std::string apply_workbook_protection(std::string_view existing, const WorkbookProtection& protection);

/// The `<protectedRange>` entries in the retained `<protectedRanges>` child.
std::vector<ProtectedRange> read_protected_ranges(const WorksheetRawExtensions& children);

/// Replaces the retained `<protectedRanges>` child with `ranges`, in its
/// schema slot; an empty list removes it.
void write_protected_ranges(WorksheetRawExtensions& children, const std::vector<ProtectedRange>& ranges);

}  // namespace formulon::io

#endif  // FORMULON_IO_PROTECTION_XML_H_
