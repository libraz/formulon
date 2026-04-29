// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the minimal styles reader. See styles_reader.h for
// the public contract.

#include "io/styles_reader.h"

#include <cstddef>
#include <string>
#include <utility>

#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace io {
namespace {

/// Counts the direct children of `parent` named `child_name`. Returns 0
/// when `parent` is empty (a missing parent means "no entries", not an
/// error: many minimal styleSheets omit the `<numFmts>` block).
std::size_t CountChildren(const pugi::xml_node& parent, const char* child_name) {
  std::size_t n = 0;
  for (pugi::xml_node c = parent.child(child_name); c; c = c.next_sibling(child_name)) {
    ++n;
  }
  return n;
}

}  // namespace

Expected<StylesTable, Error> read_styles(const std::vector<std::uint8_t>& styles_bytes) {
  pugi::xml_document doc;
  pugi::xml_parse_result parse =
      doc.load_buffer(styles_bytes.data(), styles_bytes.size(), pugi::parse_default, pugi::encoding_utf8);
  if (!parse) {
    std::string ctx("context=styles_reader part=xl/styles.xml desc=");
    ctx.append(parse.description());
    return make_error(FormulonErrorCode::kIoXmlParse, "styles.xml: pugixml parse failed", std::move(ctx));
  }
  pugi::xml_node root = doc.child("styleSheet");
  if (!root) {
    return make_error(FormulonErrorCode::kIoContentTypeInvalid, "styles.xml: missing <styleSheet> root",
                      "context=styles_reader part=xl/styles.xml");
  }

  StylesTable table;
  if (pugi::xml_node cell_xfs = root.child("cellXfs"); cell_xfs) {
    table.cell_xfs_count = CountChildren(cell_xfs, "xf");
  }
  if (pugi::xml_node num_fmts = root.child("numFmts"); num_fmts) {
    table.num_fmts_count = CountChildren(num_fmts, "numFmt");
  }
  return table;
}

}  // namespace io
}  // namespace formulon
