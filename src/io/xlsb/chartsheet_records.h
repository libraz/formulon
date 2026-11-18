//
// A chartsheet part read from an .xlsb (`xl/chartsheets/sheetN.bin`) as the
// `<chartsheet>` XML an .xlsx save writes for it. The records are laid out as
// Excel saves them (BrtBeginSheet, BrtCsProp, the chartsheet views, the
// protection, the print records, BrtDrawing, FRT blocks, BrtEndSheet); a
// record without a measured XML form is left out and counted.

#ifndef FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
#define FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/xlsb/record.h"
#include "passthrough_part.h"
#include "theme.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Workbook;
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

/// The chartsheet part an .xlsb save writes for a `<chartsheet>`.
struct ChartsheetBin {
  std::vector<std::uint8_t> bytes;
  /// Elements or attributes the records leave out because their form is not known.
  std::uint32_t unconverted = 0;
};

/// Converts a `<chartsheet>` part, rendering a theme tab colour against
/// `theme`. Fails with `kIoSheetCorrupt` when it does not parse or its root is
/// not `chartsheet`.
Expected<ChartsheetBin, Error> chartsheet_records_from_xml(std::string_view xml, const Theme& theme);

/// The zoom of a chartsheet's first view: `zoomScale` (100 when absent) and
/// `zoomToFit`, BrtBeginCsView's zoom and fZoomToFit.
struct ChartsheetZoom {
  std::uint32_t zoom_scale = 100;
  bool zoom_to_fit = false;
  bool operator==(const ChartsheetZoom& other) const {
    return zoom_scale == other.zoom_scale && zoom_to_fit == other.zoom_to_fit;
  }
  bool operator!=(const ChartsheetZoom& other) const { return !(*this == other); }
};

/// The zoom of a `<chartsheet>` part's first `<sheetView>`; nullopt when it has none.
std::optional<ChartsheetZoom> chartsheet_zoom_from_xml(std::string_view xml);

/// The zoom of a chartsheet part's first BrtBeginCsView; nullopt when it has
/// none or a record overruns the part.
std::optional<ChartsheetZoom> chartsheet_zoom_from_records(ByteSpan part);

/// `xml` with its first `<sheetView>` stating `zoom`: `zoomScale` (left out
/// at 100) and `zoomToFit`. Every byte outside those two attributes is kept;
/// a part without a view is returned unchanged.
std::string chartsheet_xml_with_zoom(std::string_view xml, ChartsheetZoom zoom);

/// `part` with its first BrtBeginCsView stating `zoom`; every other byte is
/// kept, and a part without the record is returned unchanged.
std::vector<std::uint8_t> chartsheet_records_with_zoom(ByteSpan part, ChartsheetZoom zoom);

/// The retained chartsheet parts (`.xml` or `.bin`) whose stored zoom differs
/// from their sheet's view, restated to it; every other part is left out.
std::vector<PassthroughPart> rezoomed_chartsheet_parts(const Workbook& wb);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
