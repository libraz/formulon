//
// XML namespaces and extension URIs the chart editors write into chart parts.
// Values are byte-identical to what Excel emits.

#ifndef FORMULON_CHART_CHART_NS_H_
#define FORMULON_CHART_CHART_NS_H_

namespace formulon {
namespace chart_ns {

/// DrawingML chart namespace (`c:`).
inline constexpr const char* kNsC = "http://schemas.openxmlformats.org/drawingml/2006/chart";
/// DrawingML main namespace (`a:`).
inline constexpr const char* kNsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
/// Office 2013 chart extension namespace (`c15:`).
inline constexpr const char* kNsC15 = "http://schemas.microsoft.com/office/drawing/2012/chart";
/// Office 2016 chartex namespace (`cx:`).
inline constexpr const char* kNsCx = "http://schemas.microsoft.com/office/drawing/2014/chartex";

/// `c:ext/@uri` of the `c15:` extensions (filtered series and titles, formula refs).
inline constexpr const char* kUriC15 = "{02D57815-91ED-43cb-92C2-25804820EDAC}";
/// `c:ext/@uri` of the data-label leader-line extension.
inline constexpr const char* kUriLeader = "{CE6537A1-D6FC-4f65-9D91-7224C49458BB}";

}  // namespace chart_ns
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_NS_H_
