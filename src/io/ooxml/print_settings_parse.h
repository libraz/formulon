//
// Shared derivation of `SheetPrintSettings`' structured views from the
// raw print-settings XML the reader captures verbatim.
//
// The raw fragments (`page_setup_xml`, `page_margins_xml`, `sheet_pr_xml`,
// ...) are the writer's source of truth; `PageSetup` / `PageMargins` /
// `ManualBreak` are additive projections for consumers that need typed
// access — chiefly the paginator, which reads the structured views only.
//
// Both producers of those fragments live behind these functions: the OOXML
// reader on load, and the C-ABI print setters on mutation. Sharing one
// implementation is what makes "set the raw XML, then paginate" observe the
// new settings without either side re-deriving the parse.

#ifndef FORMULON_IO_OOXML_PRINT_SETTINGS_PARSE_H_
#define FORMULON_IO_OOXML_PRINT_SETTINGS_PARSE_H_

#include <string_view>
#include <vector>

#include "pugixml.hpp"
#include "sheet.h"

namespace formulon {
namespace io {
namespace ooxml {

/// Populates the structured `PageSetup` fields from a `<pageSetup>` node.
/// Missing attributes keep whatever `out` already holds, so callers that
/// need "absent means default" must reset the struct first.
/// `fit_to_page` is NOT touched here — it lives in `<sheetPr>`; use
/// `read_fit_to_page`.
void apply_structured_page_setup(const pugi::xml_node& page_setup, PageSetup& out);

/// Populates the structured `PageMargins` fields from a `<pageMargins>`
/// node. Same "missing attributes keep the current value" contract.
///
/// A margin outside the shared non-negative-double lexical space keeps the
/// current value too. The paginator subtracts the margins from the paper to
/// get the printable body, so an infinite or NaN one collapses that body and
/// breaks a page before every single track; a negative one inflates the body
/// past the paper. Neither is a margin, and the raw XML string is re-emitted
/// verbatim regardless, so nothing the file states is lost.
void apply_structured_page_margins(const pugi::xml_node& page_margins, PageMargins& out);

/// Appends the `<brk>` children of a `<rowBreaks>` / `<colBreaks>` node to
/// `out`. OOXML's `id` is already the 0-based index the break precedes, so
/// it is carried over unchanged. The `count` / `manualBreakCount` wrapper
/// attributes are ignored — only the `<brk>` entries are honoured.
///
/// `out` is left strictly increasing by `id` and no longer than
/// `kMaxManualBreaksPerAxis`, which is the shape every consumer of the
/// break vectors assumes and the shape the mutation API maintains. A
/// document is free to spell its `<brk>` children in any order and to
/// repeat an `id`; the first entry for each `id` wins, and any excess
/// past the axis cap is dropped.
void read_manual_breaks(const pugi::xml_node& breaks_node, std::vector<ManualBreak>& out);

/// Reads `<sheetPr><pageSetUpPr fitToPage>`. Returns false when either
/// element or the attribute is absent, matching ECMA-376's default.
bool read_fit_to_page(const pugi::xml_node& sheet_pr);

/// Parses `fragment` as a standalone XML document and re-derives every
/// structured view a print-settings fragment can feed, dispatching on
/// `element_name` (`pageSetup` / `pageMargins` / `sheetPr`). Fragments with
/// no structured projection (`printOptions`, `headerFooter`) are accepted
/// and leave `settings` untouched.
///
/// An empty `fragment` resets that element's structured view to its default
/// instead of parsing, which is what "setting the empty string removes the
/// element" has to mean for the paginator.
///
/// The caller is expected to have validated the fragment already (see
/// `c_api/parts/xml_fragment.h`); a parse failure here is treated as an
/// empty fragment rather than reported, because there is no recovery this
/// layer could offer.
void refresh_structured_views(std::string_view element_name, std::string_view fragment, SheetPrintSettings& settings);

}  // namespace ooxml
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_PRINT_SETTINGS_PARSE_H_
