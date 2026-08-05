//
// MS-XLSB `xl/styles.bin` writer.  This is deliberately separate from the
// package writer: style records are a binary record stream, while the package
// envelope and relationships remain XML.

#ifndef FORMULON_IO_XLSB_STYLES_WRITER_H_
#define FORMULON_IO_XLSB_STYLES_WRITER_H_

#include <cstdint>
#include <vector>

#include "io/styles_reader.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Emits a complete, self-contained `xl/styles.bin` part for `table`.
///
/// Empty font/fill/border/XF collections are normalised to one default record,
/// as required by the XLSB collection contracts.  The writer preserves the
/// modelled font, fill, border, number-format, XF, and named-style fields;
/// OOXML-only extensions such as differential formats are intentionally out
/// of scope for this binary part and remain handled by the package passthrough
/// path.
std::vector<std::uint8_t> write_styles_bin(const StylesTable& table);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_STYLES_WRITER_H_
