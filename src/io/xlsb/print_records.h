// Worksheet print records the XLSB writer derives from the model's print
// settings: print options, margins, page setup, header and footer, and the
// manual row and column breaks, spliced at the head of the retained
// after-hyperlinks tail slot in stream order. Each family the model holds
// replaces its retained copies; one it does not hold keeps them.

#ifndef FORMULON_IO_XLSB_PRINT_RECORDS_H_
#define FORMULON_IO_XLSB_PRINT_RECORDS_H_

#include <cstdint>
#include <vector>

#include "sheet_layout.h"

namespace formulon {
namespace io {
namespace xlsb {

/// True when the model holds any of the sheet's print record families, which
/// the writer then emits itself.
bool model_owns_print_records(const SheetPrintSettings& print);

/// True when `print` holds a setting the XLSB writer does not write from the
/// model yet, which the save reports as a deferred `print_settings` feature.
bool has_deferred_print_settings(const SheetPrintSettings& print);

/// When `model_owns_print_records(print)`, rewrites the print records at the
/// head of `after_hyperlinks`: each family the model holds is written from
/// `print`, the others keep their retained records, and BrtPrintOptions is
/// written with its defaults when neither holds one. A `<pageSetup>` stating
/// only `usePrinterDefaults="0"` and a break list without manual breaks write
/// nothing. A non-empty buffer that does not split into records is left
/// untouched.
void splice_print_records(std::vector<std::uint8_t>& after_hyperlinks, const SheetPrintSettings& print);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PRINT_RECORDS_H_
