//
// A chartsheet part read from an .xlsb (`xl/chartsheets/sheetN.bin`) as the
// `<chartsheet>` XML an .xlsx save writes for it. The records are laid out as
// Excel saves them (BrtBeginSheet, BrtCsProp, the chartsheet views, the
// protection, the print records, BrtDrawing, FRT blocks, BrtEndSheet); a
// record without a measured XML form is left out and counted.

#ifndef FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
#define FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_

#include <cstdint>
#include <string>

#include "io/xlsb/record.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace xlsb {

/// The `<chartsheet>` part an .xlsx save writes for a chartsheet's records.
struct ChartsheetXml {
  std::string xml;
  /// Records or settings the XML leaves out because their form is not known.
  std::uint32_t unconverted = 0;
};

/// Converts the records of a chartsheet part. Fails with
/// `kIoXlsbRecordTruncated` when a record overruns the part.
Expected<ChartsheetXml, Error> chartsheet_xml_from_records(ByteSpan part);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
