// Worksheet print records the XLSB writer derives from the model's print
// settings: BrtPrintOptions and BrtMargins, spliced at the head of the
// retained after-hyperlinks tail slot in place of any retained copies.

#ifndef FORMULON_IO_XLSB_PRINT_RECORDS_H_
#define FORMULON_IO_XLSB_PRINT_RECORDS_H_

#include <cstdint>
#include <vector>

#include "sheet_layout.h"

namespace formulon {
namespace io {
namespace xlsb {

/// True when the model holds the sheet's print options or margins, which the
/// writer then emits itself.
bool model_owns_print_records(const SheetPrintSettings& print);

/// True when `print` holds a setting the XLSB writer does not write from the
/// model yet, which the save reports as a deferred `print_settings` feature.
bool has_deferred_print_settings(const SheetPrintSettings& print);

/// When `model_owns_print_records(print)`, replaces the BrtPrintOptions and
/// BrtMargins records of `after_hyperlinks` with ones written from `print`:
/// BrtPrintOptions always (gridLinesSet on unless stated off), BrtMargins
/// when the sheet has `<pageMargins>`. A non-empty buffer that does not split
/// into records is left untouched.
void splice_print_records(std::vector<std::uint8_t>& after_hyperlinks, const SheetPrintSettings& print);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PRINT_RECORDS_H_
