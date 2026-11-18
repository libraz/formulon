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

/// The protection a chartsheet part states, as stored: `<sheetProtection>`, or BrtCsProtection with an optional
/// BrtCsProtectionIso. Empty strings and a zero `spin_count` stand for a value the part does not state.
struct ChartsheetProtection {
  std::string algorithm_name;
  /// Base64 text, as `<sheetProtection hashValue>` holds it.
  std::string hash_value;
  /// Base64 text, as `<sheetProtection saltValue>` holds it.
  std::string salt_value;
  std::uint32_t spin_count = 0;
  /// The legacy 16-bit password hash as four hex digits.
  std::string legacy_password_hash;
  bool content_locked = false;
  bool objects_locked = false;
  bool operator==(const ChartsheetProtection& other) const {
    return algorithm_name == other.algorithm_name && hash_value == other.hash_value && salt_value == other.salt_value &&
           spin_count == other.spin_count && legacy_password_hash == other.legacy_password_hash &&
           content_locked == other.content_locked && objects_locked == other.objects_locked;
  }
  bool operator!=(const ChartsheetProtection& other) const { return !(*this == other); }
};

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

/// Projects a `<chartsheet>` part onto `sheet`: its first view's tab
/// selection and zoom, and its `sheetPr`, `pageMargins`, `pageSetup` and
/// `headerFooter`. A part that does not parse leaves `sheet` unchanged.
void apply_chartsheet_part(std::string_view xml, Sheet& sheet);

/// The part of a sheet `view` a chartsheet's view can hold.
ChartsheetView chartsheet_view_of(const SheetView& view);

/// The retained chartsheet parts (`.xml` or `.bin`) whose stored view differs
/// from their sheet's view, restated to it; every other part is left out.
std::vector<PassthroughPart> restated_chartsheet_parts(const Workbook& wb);

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_CHARTSHEET_RECORDS_H_
