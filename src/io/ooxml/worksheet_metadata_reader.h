// Worksheet metadata half of the OOXML reader: every non-cell `<worksheet>`
// element, plus the shell that lets the SAX path parse it as a DOM.

#ifndef FORMULON_IO_OOXML_WORKSHEET_METADATA_READER_H_
#define FORMULON_IO_OOXML_WORKSHEET_METADATA_READER_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/package_diagnostics.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace io {
namespace ooxml {

/// Returns a copy of `sheet_xml` with the `<sheetData>` element's children
/// removed, so the SAX path can parse the small non-cell metadata as a DOM.
/// Returns the input unchanged when there is no non-empty `<sheetData>`.
std::vector<std::uint8_t> build_worksheet_shell_bytes(const std::vector<std::uint8_t>& sheet_xml);

/// Reads every non-cell worksheet element (siblings of `<sheetData>`) from
/// `doc` into sheet `i`. Shared between the DOM path (full document) and
/// the SAX path (metadata shell). Dropped overlay entries land in
/// `diagnostics`.
Expected<void, Error> apply_worksheet_metadata(const pugi::xml_document& doc, std::size_t i, Workbook& wb,
                                               ReadDiagnostics* diagnostics);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WORKSHEET_METADATA_READER_H_
