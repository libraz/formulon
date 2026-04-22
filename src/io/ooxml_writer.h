// Copyright 2026 libraz. Licensed under the MIT License.
//
// OOXML (.xlsx) package writer. The M1 slice of this writer only supports
// empty workbooks — i.e. the output ZIP contains the minimum set of parts
// that Excel 365 will open without complaint. Later milestones extend this
// function in place with shared strings, styles, defined names, tables, and
// the full cell store. See backup/plans/04-xlsx-io.md for the complete
// contract.

#ifndef FORMULON_IO_OOXML_WRITER_H_
#define FORMULON_IO_OOXML_WRITER_H_

#include <cstdint>
#include <vector>

#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon {
namespace io {

/// Serialises `wb` into an in-memory `.xlsx` byte stream.
///
/// The M1 slice emits exactly six parts into a miniz-produced ZIP:
///
///   * `[Content_Types].xml`
///   * `_rels/.rels`
///   * `xl/workbook.xml`
///   * `xl/_rels/workbook.xml.rels`
///   * `xl/worksheets/sheet<N>.xml` (one per sheet, 1-based)
///   * `xl/styles.xml`
///
/// Sheet IDs and relationship IDs are assigned sequentially per sheet, with
/// the styles relationship following the last worksheet. Non-ASCII sheet
/// names are emitted as UTF-8 bytes with the five XML-critical characters
/// (`& < > " '`) escaped.
///
/// Returns `FormulonErrorCode::kIoWriteFailed` on any miniz failure; the
/// error context identifies the offending part.
Expected<std::vector<std::uint8_t>, Error> write_ooxml(const Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_WRITER_H_
