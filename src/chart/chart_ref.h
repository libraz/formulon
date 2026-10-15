//
// Reference sites inside chart parts: enumeration and rewriting. Elements
// match by namespace URI under any prefix. Sites are the `c:f` elements under
// series `c:tx|c:cat|c:val|c:xVal|c:yVal|c:bubbleSize`, title and axis-title
// `c:strRef`, `c15:f` in `c:extLst`, `c15:sqref` under `c15:fullRef` /
// `c15:formulaRef`, and chartex `cx:f`.

#ifndef FORMULON_CHART_CHART_REF_H_
#define FORMULON_CHART_CHART_REF_H_

#include <cstddef>
#include <vector>

#include "parser/ast_shift.h"
#include "pugixml.hpp"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

namespace chart_ref {

/// Every reference-bearing text element of `doc`, in document order.
std::vector<pugi::xml_node> enumerate_sites(const pugi::xml_document& doc);

/// Rewrites the text of every site of `doc` through `transform`; a site that
/// does not parse is left as is. A deleted area keeps its sheet qualifier
/// (`Sheet1!#REF!`) while its sheet survives, and a series name or category
/// with a deleted area moves into `c15:filteredSeriesTitle` /
/// `c15:filteredCategoryTitle`. Returns the number of sites changed.
std::size_t rewrite_sites(pugi::xml_document& doc, const parser::RefTransform& transform);

/// Runs `rewrite_sites` over every chart of every sheet (chartsheets included)
/// and stores each part whose text changed.
Expected<void, Error> rewrite_all(Workbook& wb, const parser::RefTransform& transform);

}  // namespace chart_ref
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_REF_H_
