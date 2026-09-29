//
// MS-XLSB sheet and workbook protection records. Field layouts were
// measured from Excel 365 (macOS) re-saves of OOXML workbooks, checked
// against the `<sheetProtection>` / `<workbookProtection>` Excel wrote for
// the same book:
//
//   * BrtSheetProtection (535): protpwd u16, then 16 u32 booleans --
//     fLocked, then the fifteen action flags in `SheetProtection` order
//     (objects .. pivotTables, selectUnlockedCells last). Each action flag
//     is the inverse of its OOXML attribute (1 = action allowed).
//   * BrtSheetProtectionIso (678, precedes 535): spin count u32, the same
//     16 booleans, hash and salt as u32-length byte blobs, then the
//     algorithm name as an XLWideString.
//   * BrtBookProtection (534): protpwd u16, a u16 Mac Excel always writes
//     as 0, and a u16 flag word whose bit 0 is lockStructure.
//   * BrtBookProtectionIso (677, precedes 534): book and revision spin
//     counts, the flag word, then the book's hash / salt / algorithm and
//     the revision's hash / salt / nullable algorithm.
//
// Mac Excel drops lockWindows and lockRevision on save, so their bits
// are unmeasured and a record setting them is refused.

#ifndef FORMULON_IO_XLSB_PROTECTION_RECORDS_H_
#define FORMULON_IO_XLSB_PROTECTION_RECORDS_H_

#include <cstdint>
#include <string>
#include <vector>

#include "io/zip_reader.h"
#include "sheet.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

inline constexpr std::uint16_t kBrtSheetProtection = 535;
inline constexpr std::uint16_t kBrtSheetProtectionIso = 678;
inline constexpr std::uint16_t kBrtBookProtection = 534;
inline constexpr std::uint16_t kBrtBookProtectionIso = 677;

/// Folds a BrtSheetProtection payload into `out`. Returns false, leaving
/// `out` untouched, for a payload whose width or values were not measured.
/// A record with fLocked clear leaves the sheet unprotected: Excel writes
/// one for every sheet and omits `<sheetProtection>` from .xlsx for it.
bool decode_sheet_protection(ByteSpan payload, SheetProtection& out);

/// Folds a BrtSheetProtectionIso payload's password fields into `out`.
/// Same refusal contract as `decode_sheet_protection`.
bool decode_sheet_protection_iso(ByteSpan payload, SheetProtection& out);

/// Emits BrtSheetProtectionIso (when a modern hash is set) and
/// BrtSheetProtection for an enabled `protection`; nothing otherwise.
/// Fails on a hash, salt or legacy password the records cannot carry.
Expected<void, Error> emit_sheet_protection(std::vector<std::uint8_t>& dst, const SheetProtection& protection);

/// Workbook protection decoded to the `<workbookProtection>` element the
/// model carries (`Workbook::workbook_protection_xml`), attributes in the
/// order Excel writes them. `iso` may be empty when the book has no
/// modern hash. Returns an empty string for an unprotected book, and
/// false for content that was not measured.
bool decode_book_protection(ByteSpan payload, ByteSpan iso, std::string& xml);

/// Emits BrtBookProtectionIso (when a modern hash is set) and
/// BrtBookProtection for a `<workbookProtection>` element; nothing for an
/// empty one. Returns false when the element carries lockWindows,
/// lockRevision or a revisions password, which have no measured field and
/// are left out. Fails on a hash, salt or password the records cannot carry.
Expected<bool, Error> emit_book_protection(std::vector<std::uint8_t>& dst, const std::string& xml);

/// Base64 (RFC 4648, padded) of `bytes`; the spelling OOXML hash and salt
/// attributes use.
std::string base64_encode(const std::uint8_t* data, std::size_t size);

/// Inverse of `base64_encode`. Returns false on anything but canonical
/// padded base64.
bool base64_decode(const std::string& text, std::vector<std::uint8_t>& out);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PROTECTION_RECORDS_H_
