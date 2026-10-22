//
// MS-XLSB `xl/styles.bin` writer.  This is deliberately separate from the
// package writer: style records are a binary record stream, while the package
// envelope and relationships remain XML.

#ifndef FORMULON_IO_XLSB_STYLES_WRITER_H_
#define FORMULON_IO_XLSB_STYLES_WRITER_H_

#include <cstdint>
#include <vector>

#include "styles.h"

namespace formulon {
namespace io {
namespace xlsb {

/// Emits a complete, self-contained `xl/styles.bin` part for `table`.
///
/// Empty font/fill/border/XF collections are normalised to one default record,
/// as required by the XLSB collection contracts.  The writer preserves the
/// modelled font, fill (pattern and gradient), border, number-format, XF,
/// named-style and differential-format fields. Gradient stops past the 256
/// `cNumStop` allows are dropped.
std::vector<std::uint8_t> write_styles_bin(const StylesTable& table);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_STYLES_WRITER_H_
