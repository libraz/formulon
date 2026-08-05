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

namespace formulon {
namespace io {

/// Detected container format of a workbook byte stream.
enum class WorkbookFormat : std::uint8_t {
  /// Not a recognised OPC package, or a package that declares neither an
  /// xlsx nor an xlsb workbook part. The caller should still attempt the
  /// OOXML reader so its richer diagnostics surface (encryption, zip
  /// corruption, etc.).
  Unknown = 0,
  /// OOXML `.xlsx` / `.xlsm`: the package contains `xl/workbook.xml`.
  Ooxml = 1,
  /// MS-XLSB `.xlsb`: the package contains the binary `xl/workbook.bin`.
  Xlsb = 2,
};

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
