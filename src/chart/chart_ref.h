//
// Reference sites inside chart parts: enumeration and rewriting. Elements
// match by namespace URI under any prefix. Sites are the `c:f` elements under
// series `c:tx|c:cat|c:val|c:xVal|c:yVal|c:bubbleSize`, title and axis-title
// `c:strRef`, `c15:f` in `c:extLst`, `c15:sqref` under `c15:fullRef` /
// `c15:formulaRef`, and chartex `cx:f`.

#ifndef FORMULON_CHART_CHART_REF_H_
#define FORMULON_CHART_CHART_REF_H_

#include <cstddef>
#include <cstdint>
#include <functional>
#include <string>
#include <string_view>
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

/// The text of every site except `c15:sqref` in every chart part not on the
/// sheet at `exclude_sheet`.
Expected<std::vector<std::string>, Error> collect_formula_texts(const Workbook& wb, std::size_t exclude_sheet);

/// The part path of every chart reachable from a sheet's drawing, chartsheets included; a path appears once.
std::vector<std::string> chart_part_paths(const Workbook& wb);

/// Passes `visit` the text of every site except `c15:sqref` of each chart on sheet `sheet`: the charts in drawing
/// anchor order, the sites of one chart in document order. Returns false when a chart part could not be loaded.
bool for_each_sheet_chart_site(const Workbook& wb, std::size_t sheet,
                               const std::function<void(std::string_view)>& visit);

/// `text` with the book prefix `[n]` of each reference replaced by `[remap[n]]`, an `n` past `remap` kept. Only a
/// prefix outside structured brackets and string literals counts; in a quoted sheet name it must open the name.
std::string respell_book_indices(std::string_view text, const std::vector<std::uint32_t>& remap);

/// Runs `respell_book_indices` over every site of every chart and stores each part whose text changed.
void respell_all_book_indices(Workbook& wb, const std::vector<std::uint32_t>& remap);

/// Runs `rewrite_sites` over every chart of every sheet (chartsheets included)
/// and stores each part whose text changed.
Expected<void, Error> rewrite_all(Workbook& wb, const parser::RefTransform& transform);

}  // namespace chart_ref
}  // namespace formulon

#endif  // FORMULON_CHART_CHART_REF_H_
