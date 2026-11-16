// Worksheet print records the XLSB writer derives from the model's print
// settings: print options, margins, page setup, header and footer, and the
// manual row and column breaks, spliced at the head of the retained
// after-hyperlinks tail slot in stream order. Each family the model holds
// replaces its retained copies; one it does not hold keeps them.
// On load, the print options, page setup, margins, header and footer and manual breaks are read back into the model
// as well; a family the reader cannot state exactly stays retained only, and an `.xlsx` save counts it.

#ifndef FORMULON_IO_XLSB_PRINT_RECORDS_H_
#define FORMULON_IO_XLSB_PRINT_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <vector>

#include "io/xlsb/record.h"
#include "sheet_layout.h"

namespace formulon {
namespace io {
namespace xlsb {

/// True when the model holds any of the sheet's print record families, which
/// the writer then emits itself.
bool model_owns_print_records(const SheetPrintSettings& print);

/// Reads the BrtPrintOptions, BrtPageSetup, BrtMargins, header and footer, and
/// manual break records retained in a loaded sheet's `after_hyperlinks` slot
/// into `print_options_xml`, `page_setup_xml` / `page_margins_xml` (and their
/// structured views), `header_footer_xml` and the manual row / column breaks,
/// as the `.xlsx` twin states them. The print options, header and footer and
/// break families are read only when the writer would re-emit exactly their
/// retained bytes (no automatic break, unknown flag bit or non-canonical
/// string); a record the reader cannot state exactly stays retained only.
void decode_print_records(const std::vector<std::uint8_t>& after_hyperlinks, SheetPrintSettings& print);

/// The `<printOptions>` of a BrtPrintOptions payload, empty when every
/// option is at its default; nullopt for a payload whose bits Excel was not
/// seen to write in a custom view (gridLinesSet cleared among them).
std::optional<std::string> decode_print_options(ByteSpan payload);

/// The print record families retained in `after_hyperlinks` that `print`
/// does not hold, which an `.xlsx` save cannot write. Print options at their
/// defaults are not counted: they state nothing.
std::uint32_t unwritten_print_families(const std::vector<std::uint8_t>& after_hyperlinks,
                                       const SheetPrintSettings& print);

/// Emits the print records of a custom view in the order Excel writes them
/// inside BrtBeginUserShView: row breaks, column breaks, margins, print
/// options (always, with the defaults when `print` states none), page setup,
/// header and footer. A family `print` does not hold is left out.
void emit_view_print_records(std::vector<std::uint8_t>& dst, const SheetPrintSettings& print);

/// Emits a chartsheet's margins and header and footer records, in that order,
/// for each family `print` holds.
void emit_chartsheet_print_records(std::vector<std::uint8_t>& dst, const SheetPrintSettings& print);

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
