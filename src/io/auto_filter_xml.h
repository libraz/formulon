//
// `<autoFilter>` element (ECMA-376 §18.3.2, `CT_AutoFilter`) reader and
// writer for the AutoFilter model in `auto_filter.h`.
//
// Attributes the model does not name on `autoFilter`, `filterColumn` and
// `sortState` (Excel's `xr:uid`, namespace declarations) and unknown
// `filterColumn` children are retained as raw XML. A fragment that does not
// parse into the model at all -- malformed XML, a ref that is not a plain A1
// rectangle, an unknown enumeration value -- is kept verbatim as `opaque_xml`
// and round-trips unchanged.
//
// Excel writes date-group criteria only inside a `filterColumn` extension
// (`xlrd2:filterColumn`, uri {1AD28BCE-077C-4C59-8B6E-1921CE8616D4}); the
// reader lifts them into `ValueFilters::date_groups` and the writer emits
// them back in that form, whichever form the source used.

#ifndef FORMULON_IO_AUTO_FILTER_XML_H_
#define FORMULON_IO_AUTO_FILTER_XML_H_

#include <string>
#include <string_view>

#include "auto_filter.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// Parses one `<autoFilter>` element. Returns `kAutoFilterInvalid` when the
/// fragment is not a single well-formed `autoFilter` element the model can
/// represent.
Expected<AutoFilter, Error> parse_auto_filter_xml(std::string_view xml);

/// Serializes `filter` as an `<autoFilter>` element. An opaque filter
/// returns its retained bytes.
std::string serialize_auto_filter(const AutoFilter& filter);

/// The serialized element of `filter`, or empty when it is null.
std::string auto_filter_xml(const AutoFilter* filter);

/// Reads a non-empty `<autoFilter>` element; a fragment the model cannot
/// represent comes back opaque.
AutoFilter auto_filter_from_xml(std::string_view xml);

}  // namespace formulon::io

#endif  // FORMULON_IO_AUTO_FILTER_XML_H_
