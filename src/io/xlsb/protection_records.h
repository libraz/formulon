//
// MS-XLSB sheet and workbook protection records. Field layouts follow
// Excel 365 (macOS) re-saves of OOXML workbooks, checked
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
//   * BrtBookProtection (534, [MS-XLSB] 2.4.311): protpwdBook u16,
//     protpwdRev u16, then BookProtectionFlags (2.5.7) u16 -- bit 0
//     lockStructure, bit 1 lockWindows, bit 2 lockRevision.
//   * BrtBookProtectionIso (677, 2.4.312, precedes 534): book and revision
//     spin counts, the flag word, then the book's and the revision's
//     IsoPasswordData (2.5.80: hash, salt, nullable algorithm). The 534
//     record that follows it carries zero verifiers.
//
// Excel (macOS and Windows 16.0.20228) drops lockWindows, lockRevision and
// the revisions password on its own save in both formats; the records still
// define them, so they are written and read like the other fields.

#ifndef FORMULON_IO_XLSB_PROTECTION_RECORDS_H_
#define FORMULON_IO_XLSB_PROTECTION_RECORDS_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/record_ids.h"
#include "io/zip_reader.h"
#include "protection.h"
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
/// `out` untouched, for a payload whose width or values are not supported.
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
/// order `apply_workbook_protection` writes them. `iso` may be empty when
/// the book has no modern hash. Returns an empty string for an unprotected
/// book, and false for a malformed payload.
bool decode_book_protection(ByteSpan payload, ByteSpan iso, std::string& xml);

/// Emits BrtBookProtectionIso (when a modern hash is set) and
/// BrtBookProtection for a `<workbookProtection>` element; nothing for an
/// element with no protection. Returns false when a legacy verifier sits
/// beside a modern hash: the 534 record after an Iso record must carry zero
/// verifiers, so the legacy one is left out. Fails on a hash, salt or
/// password the records cannot carry.
Expected<bool, Error> emit_book_protection(std::vector<std::uint8_t>& dst, const std::string& xml);

/// Base64 (RFC 4648, padded) of `bytes`; the spelling OOXML hash and salt
/// attributes use.
std::string base64_encode(const std::uint8_t* data, std::size_t size);

/// Inverse of `base64_encode`. Returns false on anything but canonical
/// padded base64.
bool base64_decode(const std::string& text, std::vector<std::uint8_t>& out);

// Range protection, as Excel 365 (macOS) writes it:
//   * BrtRangeProtection (536): protPwd u16, UncheckedSqRfX (i32 count, then
//     four-u32 ranges), RangeProtectionTitleSDRel (XLWideString title, u32
//     dwsdRel, that many bytes of self-relative SECURITY_DESCRIPTOR, which
//     Excel encodes from the OOXML SDDL text: header, DACL, owner, group).
//   * BrtRangeProtectionIso (680, precedes a 536 with protPwd 0 and the same
//     ranges and title): dwSpinCount u32, UncheckedSqRfX,
//     RangeProtectionTitleSDRel, then IsoPasswordData (hash and salt as
//     u32-length blobs, algorithm as an XLNullableWideString).
//   * BrtRangeProtection14 (1103), for a range of more than
//     `kMaxProtectedRangeRects` areas, alone in a BrtFRTBegin / BrtFRTEnd
//     block among the worksheet future records: FRTHeader (flags u32 2, one
//     FRTSqref with flags u32 2 and an UncheckedSqRfX of any length), protPwd
//     u16, RangeProtectionTitleSDRel.

inline constexpr std::uint16_t kBrtRangeProtection14 = 1103;

/// How SDDL text encodes to a self-relative security descriptor.
enum class SddlEncode : std::uint8_t {
  kEncoded,
  /// Names a SID that has no fixed value (a domain-relative or unknown alias), or has no
  /// upper-case letter while its upper-case spelling encodes; the range is written without
  /// a descriptor, as Excel does.
  kMalformed,
  /// Outside the encoded form `O:<alias>G:<alias>[D:(A|D;;<rights>;;;<alias>)...]`, its
  /// sections also accepted as `G:..O:..D:..` or `D:..O:..G:..`, with fixed-value SID aliases
  /// in either case and two-letter right tokens.
  kUnsupported,
};

/// Encodes `sddl` into `out` (set only for `kEncoded`).
SddlEncode encode_security_descriptor(std::string_view sddl, std::vector<std::uint8_t>& out);

/// SDDL text of a self-relative descriptor of the encoded form, with the rights of
/// each ACE spelled canonically: the composite token with exactly that mask, else
/// the one-bit tokens in ascending bit order. The owner, group and DACL may sit in
/// any order after the header. False for any other descriptor, including a mask
/// that has no token spelling.
bool decode_security_descriptor(ByteSpan sd, std::string& out);

/// The spelling Excel keeps for `sddl`: the canonical text of a descriptor it encodes, empty for one it
/// drops, and `sddl` itself otherwise (including a mask only a hex spelling names).
std::string canonical_security_descriptor(std::string_view sddl);

/// How a BrtRangeProtection record (with its Iso record) decoded.
enum class RangeProtectionRead : std::uint8_t {
  kDecoded,
  /// Well-formed, but carries a security descriptor that does not decode to
  /// SDDL text, or differs between the pair; the records are kept verbatim.
  kSecurityDescriptor,
  /// Malformed, a mismatched pair, an empty or over-long title, or a range
  /// count outside 1..`kMaxProtectedRangeRects`.
  kUndecodable,
};

/// Decodes a BrtRangeProtection `payload` and the BrtRangeProtectionIso `iso`
/// before it (empty when none). `out` is set only for `kDecoded`.
RangeProtectionRead decode_range_protection(ByteSpan payload, ByteSpan iso, ProtectedRange& out);

/// Emits ([BrtRangeProtectionIso] BrtRangeProtection) for each range of at
/// most `kMaxProtectedRangeRects` areas, the Iso record when a modern hash is
/// set; longer ranges are left to `add_range_protections14`. Fails with `kIoFeatureNotWritable`
/// (context `kind=` and `sheet_index=`) on a range the records cannot carry:
/// SDDL text outside the encoded form, a name outside 1..`kMaxProtectedRangeNameUnits` UTF-16 units,
/// an sqref that does not parse, or more than `kMaxProtectedRangeRects`
/// ranges under a modern hash. Fails with `kInvalidArgument` on a hash, salt
/// or legacy password the records cannot carry. A malformed descriptor is
/// left out, as Excel does.
Expected<void, Error> emit_range_protections(std::vector<std::uint8_t>& dst, const std::vector<ProtectedRange>& ranges,
                                             std::size_t sheet_index);

/// Inserts one FRT block of BrtRangeProtection14 records, one per range of
/// more than `kMaxProtectedRangeRects` areas, into the worksheet tail
/// `records`: after the x14 CF, DVals14 and sparkline-group containers,
/// ahead of any later future record and a trailing revision-uid block. Same
/// failure contract as `emit_range_protections`.
Expected<void, Error> add_range_protections14(std::vector<std::uint8_t>& records,
                                              const std::vector<ProtectedRange>& ranges, std::size_t sheet_index);

/// Moves the ranges of every FRT block in the worksheet tail `records` that
/// holds only decodable BrtRangeProtection14 records into `out`, removing the
/// block; any other block stays as read.
void take_range_protections14(std::vector<std::uint8_t>& records, std::vector<ProtectedRange>& out);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PROTECTION_RECORDS_H_
