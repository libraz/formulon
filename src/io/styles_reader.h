//
// `xl/styles.xml` reader. Decodes the OOXML styles part into a flat
// in-memory `StylesTable` carrying every record kind the engine needs
// to round-trip a workbook: fonts, fills, borders, custom number-format
// strings, and the per-cell xf table that ties them together.
//
// The record types and `builtin_num_fmt(id)` live in `styles.h`.

#ifndef FORMULON_IO_STYLES_READER_H_
#define FORMULON_IO_STYLES_READER_H_

#include <cstdint>
#include <string>
#include <vector>

#include "styles.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {

/// Parses an OOXML styles part.
///
/// Behaviour:
///   * Empty `<styleSheet/>` (no children) yields a table whose `fonts`,
///     `fills`, `borders`, and `cell_xfs` each contain a single default
///     entry so `xf_index = 0` is always resolvable.
///   * Sections present but empty (`<fonts count="0"/>`) similarly fall
///     back to the single-default-record shape.
///   * `<colors>`, `<tableStyles>`, and `<extLst>` are retained as raw XML;
///     other unrecognised children are accepted but ignored.
///   * Unknown attributes inside a recognised element are ignored.
///   * Every stored index resolves. A `fontId` / `fillId` / `borderId` /
///     `xfId` naming a record the part does not define is rewritten to
///     the default record `0`, matching Excel's own fallback: such a
///     workbook loads, its getters succeed, and the writer never
///     re-emits the dangling reference.
///
/// Errors:
///   * `kIoXmlParse` — pugixml could not parse the document.
///   * `kIoContentTypeInvalid` — the document parses but its root is
///     not `<styleSheet>`.
///   * `kIoSheetCorrupt` — a style `<xf>` carries a malformed alignment
///     boolean / enum / integer or an alignment value outside its OOXML
///     range. The error context identifies the table, xf index, and attribute.
Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STYLES_READER_H_
