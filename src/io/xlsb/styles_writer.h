//
// MS-XLSB `xl/styles.bin` writer.  This is deliberately separate from the
// package writer: style records are a binary record stream, while the package
// envelope and relationships remain XML.

#ifndef FORMULON_IO_XLSB_STYLES_WRITER_H_
#define FORMULON_IO_XLSB_STYLES_WRITER_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <vector>

#include "io/dxf_save_plan.h"
#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class CellStyleSavePlan;

namespace io {
namespace xlsb {

/// Emits a complete, self-contained `xl/styles.bin` part for `table`.
///
/// Empty font/fill/border/XF collections are normalised to one default record,
/// as required by the XLSB collection contracts.  The writer preserves the
/// modelled font, fill (pattern and gradient), border, number-format, XF,
/// named-style and differential-format fields. Every gradient stop is written.
/// `plan` (null: every dxf in model order) selects and orders the dxfs written; a model dxf may repeat.
/// `style_plan` (null: as the model holds them) selects, names and orders the
/// BrtStyle records and style BrtXFs and remaps each cell BrtXF's parent.
std::vector<std::uint8_t> write_styles_bin(const StylesTable& table, const DxfSavePlan* plan = nullptr,
                                           const CellStyleSavePlan* style_plan = nullptr);

/// The BrtDXF count of a source styles part whose BrtDXF records form one list
/// directly under its only BrtBeginDXFs (zero when it has neither); nullopt for
/// any other layout or a part that does not frame whole.
std::optional<std::size_t> styles_bin_dxf_count(const std::vector<std::uint8_t>& raw);

/// The source styles part `raw` with its dxf list written in `plan.order()`,
/// BrtBeginDXFs counting it, and every dxf slot mapped through `plan`; every
/// other record is byte-equal. Fails for a dxf list `styles_bin_dxf_count`
/// rejects or a dxf slot that cannot be located.
Expected<std::vector<std::uint8_t>, Error> compact_styles_bin(const std::vector<std::uint8_t>& raw,
                                                              const DxfSavePlan& plan);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_STYLES_WRITER_H_
