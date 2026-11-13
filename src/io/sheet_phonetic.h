// The worksheet's own `<phoneticPr>` (the furigana defaults for cells without
// one), held as a raw worksheet child and shared by the xlsx and xlsb paths.

#ifndef FORMULON_IO_SHEET_PHONETIC_H_
#define FORMULON_IO_SHEET_PHONETIC_H_

#include <optional>

#include "phonetic.h"

namespace formulon {

class Sheet;
class Workbook;

namespace io {

/// The sheet's `<phoneticPr>`, when its raw children hold one.
std::optional<PhoneticProperties> sheet_phonetic_properties(const Sheet& sheet);

/// Stores `props` as the sheet's `<phoneticPr>` child, written as Excel
/// writes it: `fontId`, then `type` and `alignment` only where they differ
/// from the schema defaults.
void set_sheet_phonetic_properties(Sheet& sheet, const PhoneticProperties& props);

/// Points each sheet `<phoneticPr>` whose `fontId` lies past the font table
/// at font 0, keeping its type and alignment, as Excel does when it loads
/// the book.
void clamp_sheet_phonetic_fonts(Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_SHEET_PHONETIC_H_
