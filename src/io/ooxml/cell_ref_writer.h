//
// Cell-reference formatters for the OOXML writer. Bijective base-26
// column letters plus the `<row><col>` joiners used by `<mergeCell>`,
// `<hyperlink>`, and other range-shaped XML elements. Internal to
// `src/io/ooxml/`; not part of the public API.

#ifndef FORMULON_IO_OOXML_CELL_REF_WRITER_H_
#define FORMULON_IO_OOXML_CELL_REF_WRITER_H_

#include <cstdint>
#include <string>

namespace formulon {
struct MergeRange;
namespace io {

/// Bijective base-26 column letters (`0 -> A`, `25 -> Z`, `26 -> AA`,
/// ...).
void AppendColumnLettersForRef(std::string& out, std::uint32_t col);

/// Appends the Excel A1-style cell reference for the given 0-based
/// (row, col) (e.g. `(0, 0)` -> `"A1"`).
void AppendCellRefForRef(std::string& out, std::uint32_t row, std::uint32_t col);

/// Appends the Excel A1-style range reference for `r`. Collapsed to a
/// single cell when first/last coincide.
void AppendRangeRef(std::string& out, const MergeRange& r);

/// As above, from the 0-based corner coordinates.
void AppendRangeRef(std::string& out, std::uint32_t first_row, std::uint32_t first_col, std::uint32_t last_row,
                    std::uint32_t last_col);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_CELL_REF_WRITER_H_
