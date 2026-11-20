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
#include "sheet_layout.h"
#include "styles.h"
#include "theme.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Sheet;
struct SheetView;
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

/// A chartsheet's first view: `tabSelected`, `zoomScale` (100 when absent)
/// and `zoomToFit`, BrtBeginCsView's fSelected, zoom and fZoomToFit.
struct ChartsheetView {
  std::uint32_t zoom_scale = 100;
  bool zoom_to_fit = false;
  bool tab_selected = false;
  bool operator==(const ChartsheetView& other) const {
    return zoom_scale == other.zoom_scale && zoom_to_fit == other.zoom_to_fit && tab_selected == other.tab_selected;
  }
  bool operator!=(const ChartsheetView& other) const { return !(*this == other); }
};

/// The view of a `<chartsheet>` part's first `<sheetView>`; nullopt when it has none.
std::optional<ChartsheetView> chartsheet_view_from_xml(std::string_view xml);

/// The view of a chartsheet part's first BrtBeginCsView; nullopt when it has
/// none or a record overruns the part.
std::optional<ChartsheetView> chartsheet_view_from_records(ByteSpan part);

/// The protection a chartsheet part states, as stored; the model type.
using formulon::ChartsheetProtection;

/// The protection of a `<chartsheet>` part; nullopt when it has no `<sheetProtection>` or the part does not parse.
std::optional<ChartsheetProtection> chartsheet_protection_from_xml(std::string_view xml);

/// The protection of a chartsheet part's BrtCsProtection; nullopt when it has none, it is all zero with no
/// BrtCsProtectionIso, or a record overruns the part.
std::optional<ChartsheetProtection> chartsheet_protection_from_records(ByteSpan part);

/// `xml` with its first `<sheetView>` stating `view`: `tabSelected` (left out
/// when false), `zoomScale` (left out at 100) and `zoomToFit`. Every byte
/// outside those attributes is kept; a part without a view is returned unchanged.
std::string chartsheet_xml_with_view(std::string_view xml, ChartsheetView view);

/// `part` with its first BrtBeginCsView stating `view`; every other byte is
/// kept, and a part without the record is returned unchanged.
std::vector<std::uint8_t> chartsheet_records_with_view(ByteSpan part, ChartsheetView view);

/// `xml` with its `<sheetPr>` stating `color` as `<tabColor>`, or closing empty
/// for no colour; a part without `<sheetPr>` gains one. Every byte outside
/// `<sheetPr>` is kept; a part already reporting `color`, or one that does not
/// parse, is returned unchanged.
std::string chartsheet_xml_with_tab_color(std::string_view xml, const ColorSpec& color);

/// `part` with its first BrtCsProp's tab colour stating `color`, a theme or
/// indexed colour rendered against `theme`; every other byte is kept, and a
/// part already reporting `color` or without the record is returned unchanged.
std::vector<std::uint8_t> chartsheet_records_with_tab_color(ByteSpan part, const ColorSpec& color, const Theme& theme);

/// The tab colour of a chartsheet part's first BrtCsProp, read whatever else
/// the record states (a code name, a tint); `Kind::kNone` for the automatic
/// colour, nullopt when the part has no such record.
std::optional<ColorSpec> chartsheet_tab_color_from_records(ByteSpan part);

/// `xml` with the `<pageMargins>`, `<pageSetup>`, `<headerFooter>` and `<sheetProtection>` that differ from `sheet`'s
/// model restated to it: each is replaced where it stands, inserted at its schema position when absent, and dropped
/// when the model holds none. A `<pageSetup>` always carries `orientation`, right after `paperSize`.
/// `stored_protection` is the protection the part states. Every byte outside the restated elements is kept.
std::string chartsheet_xml_with_settings(std::string_view xml, const Sheet& sheet,
                                         const std::optional<ChartsheetProtection>& stored_protection);

/// `part` with the margins, page setup, header and footer and protection records that differ from `sheet`'s model
/// restated to it: each group is replaced where it stands, or inserted at the order Excel writes when absent.
/// Unprotected is the all-zero BrtCsProtection. Every other record is kept byte for byte.
std::vector<std::uint8_t> chartsheet_records_with_settings(ByteSpan part, const Sheet& sheet);

/// Projects a `<chartsheet>` part onto `sheet`: its first view's tab
/// selection and zoom, and its `sheetPr`, `pageMargins`, `pageSetup`,
/// `headerFooter` and `sheetProtection`. A part that does not parse leaves `sheet` unchanged.
void apply_chartsheet_part(std::string_view xml, Sheet& sheet);

/// The part of a sheet `view` a chartsheet's view can hold.
ChartsheetView chartsheet_view_of(const SheetView& view);

/// The retained chartsheet parts (`.xml` or `.bin`) whose stored view, tab
/// colour, margins, page setup, header and footer or protection differs from
/// their sheet's, restated to it; every other part is left out.
std::vector<PassthroughPart> restated_chartsheet_parts(const Workbook& wb);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
