//
// Conversion between the protection records in `protection.h` and the XML
// the model retains for them.

#ifndef FORMULON_IO_PROTECTION_XML_H_
#define FORMULON_IO_PROTECTION_XML_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "protection.h"
#include "sheet_passthrough.h"
#include "sheet_view.h"

namespace formulon::io {

/// One boolean attribute of `<sheetProtection>`: its member and ECMA-376
/// §18.3.1.85 schema default (eleven action flags default to true, the rest
/// to false).
struct SheetProtectionFlag {
  const char* attr;
  bool SheetProtection::*member;
  bool default_value;
};

/// The boolean `<sheetProtection>` attributes in Excel's emission order.
inline constexpr SheetProtectionFlag kSheetProtectionFlags[] = {
    {"sheet", &SheetProtection::sheet, false},
    {"objects", &SheetProtection::objects, false},
    {"scenarios", &SheetProtection::scenarios, false},
    {"formatCells", &SheetProtection::format_cells, true},
    {"formatColumns", &SheetProtection::format_columns, true},
    {"formatRows", &SheetProtection::format_rows, true},
    {"insertColumns", &SheetProtection::insert_columns, true},
    {"insertRows", &SheetProtection::insert_rows, true},
    {"insertHyperlinks", &SheetProtection::insert_hyperlinks, true},
    {"deleteColumns", &SheetProtection::delete_columns, true},
    {"deleteRows", &SheetProtection::delete_rows, true},
    {"selectLockedCells", &SheetProtection::select_locked_cells, false},
    {"sort", &SheetProtection::sort, true},
    {"autoFilter", &SheetProtection::auto_filter, true},
    {"pivotTables", &SheetProtection::pivot_tables, true},
    {"selectUnlockedCells", &SheetProtection::select_unlocked_cells, false},
};

/// Appends the `algorithmName` / `hashValue` / `saltValue` / `spinCount`
/// attributes of a modern password hash, each only when set.
void append_password_hash_attrs(std::string& out, const std::string& algorithm_name, const std::string& hash_value,
                                const std::string& salt_value, std::uint32_t spin_count);

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
