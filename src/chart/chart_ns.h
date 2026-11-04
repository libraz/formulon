//
// XML namespaces, content types and extension URIs the chart and drawing editors write.
// Values are byte-identical to what Excel emits.

#ifndef FORMULON_CHART_CHART_NS_H_
#define FORMULON_CHART_CHART_NS_H_

#include <string>

namespace formulon {
namespace chart_ns {

/// DrawingML chart namespace (`c:`).
inline constexpr const char* kNsC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/// DrawingML main namespace (`a:`).
inline constexpr const char* kNsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// SpreadsheetDrawing namespace (`xdr:`).
inline constexpr const char* kNsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
/// Content type of a worksheet drawing part.
inline constexpr const char* kCtDrawing = "application/vnd.openxmlformats-officedocument.drawing+xml";
/// Office 2013 chart extension namespace (`c15:`).
inline constexpr const char* kNsC15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
/// Office 2016 chartex namespace (`cx:`).
inline constexpr const char* kNsCx = "http://schemas.microsoft.com/office/drawing/2014/chartex";

/// `c:ext/@uri` of the `c15:` extensions (filtered series and titles, formula refs).
inline constexpr const char* kUriC15 = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
/// `c:ext/@uri` of the data-label leader-line extension.
inline constexpr const char* kUriLeader = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";
/// Office 2017 chart extension namespace (`c16r3:`).
inline constexpr const char* kNsC16r3 = "http://schemas.microsoft.com/office/drawing/2017/03/chart";
/// `c:ext/@uri` of the data-display options extension (`dispNaAsBlank`).
inline constexpr const char* kUriDataDisplayOptions16 = "{56B9EC1D-385E-4148-901F-78D8002777C0}";

/// The `c:ext` element turning `c16r3:dispNaAsBlank` on; Excel writes no element for off.
inline std::string data_display_ext_xml() {
  return std::string("<c:ext uri=\"") + kUriDataDisplayOptions16 + "\" xmlns:c16r3=\"" + kNsC16r3 +
         "\"><c16r3:dataDisplayOptions16><c16r3:dispNaAsBlank val=\"1\"/></c16r3:dataDisplayOptions16></c:ext>";
}

}  // namespace chart_ns
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_NS_H_
