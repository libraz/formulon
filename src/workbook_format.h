//
// `WorkbookFormat`: detected container format of a workbook byte stream.
// Detection lives in `io/format_detect.h`.

#ifndef FORMULON_WORKBOOK_FORMAT_H_
#define FORMULON_WORKBOOK_FORMAT_H_

#include <cstdint>

namespace formulon {

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

}  // namespace formulon

#endif  // FORMULON_WORKBOOK_FORMAT_H_
