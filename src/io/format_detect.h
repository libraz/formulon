//
// Workbook container-format detection.
//
// The C ABI load boundary receives raw bytes with no filename, so the
// reader cannot route on extension. Both `.xlsx` (OOXML) and `.xlsb`
// (MS-XLSB) are ZIP/OPC packages; they differ in which workbook part the
// package declares: `.xlsx` ships `xl/workbook.xml`, `.xlsb` ships the
// binary `xl/workbook.bin`. This module peeks the package's central
// directory to decide which reader to dispatch.

#ifndef FORMULON_IO_FORMAT_DETECT_H_
#define FORMULON_IO_FORMAT_DETECT_H_

#include <cstdint>

#include "io/zip_reader.h"
#include "workbook_format.h"

namespace formulon {
namespace io {

/// Inspects `bytes` and returns the detected container format.
///
/// The probe opens the ZIP central directory and checks for the
/// presence of the binary workbook part (`xl/workbook.bin` => Xlsb)
/// versus the XML workbook part (`xl/workbook.xml` => Ooxml). When the
/// bytes are not a readable ZIP, or contain neither part, returns
/// `Unknown` so the caller can fall back to the OOXML path (which owns
/// the authoritative "not a workbook" diagnostics). The probe never
/// fails: an unreadable buffer simply yields `Unknown`.
WorkbookFormat detect_workbook_format(ByteSpan bytes);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_FORMAT_DETECT_H_
