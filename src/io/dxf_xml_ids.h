//
// dxf references inside verbatim XML: a markup tokenizer (not a DOM) that
// finds every dxf attribute of a fragment or of an attribute list, and
// rewrites their values through a `DxfPartCursor` with every other byte kept.
//
// A dxf attribute is one named exactly `dxfId`, `headerRowDxfId`,
// `dataDxfId`, `totalsRowDxfId`, `headerRowBorderDxfId`, `tableBorderDxfId`
// or `totalsRowBorderDxfId` (a prefixed attribute name never matches), on any
// element except `slicerStyleElement` / `timelineStyleElement` under any
// prefix: those index the extension dxf collections. Namespaces are not
// resolved. A value that is not an unsigned decimal integer is not a dxf
// reference and is neither reported nor touched. Comments, CDATA, processing
// instructions, end tags and text are skipped. An unterminated tag, quote or
// skipped construct is an error.

#ifndef FORMULON_IO_DXF_XML_IDS_H_
#define FORMULON_IO_DXF_XML_IDS_H_

#include <cstdint>
#include <string>
#include <string_view>
#include <vector>

#include "io/dxf_save_plan.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon::io {

/// One dxf attribute found in markup; the views point into the scanned text.
struct DxfXmlId {
  std::string_view element;    ///< Owning element as written (prefix included); empty for an attribute list.
  std::string_view attribute;  ///< Attribute name.
  std::uint32_t value = 0;
};

/// Every dxf attribute of `fragment` (any number of roots), in document order.
Expected<std::vector<DxfXmlId>, Error> visit_dxf_xml_ids(std::string_view fragment);

/// Every dxf attribute of `attrs` (` name="value"` pairs), in order.
Expected<std::vector<DxfXmlId>, Error> visit_dxf_attr_ids(std::string_view attrs);

/// `fragment` with each dxf value, in document order, replaced by the saved
/// index `cursor` hands out, or the attribute and its leading whitespace
/// removed when it hands out nothing.
Expected<std::string, Error> rewrite_dxf_xml_ids(std::string_view fragment, DxfPartCursor& cursor);

/// `attrs` rewritten as `rewrite_dxf_xml_ids` rewrites a fragment.
Expected<std::string, Error> rewrite_dxf_attr_ids(std::string_view attrs, DxfPartCursor& cursor);

}  // namespace formulon::io

#endif  // FORMULON_IO_DXF_XML_IDS_H_
