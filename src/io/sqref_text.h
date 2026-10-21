//
// The A1 text form of an sqref (`A1:B2 D4`) shared by the retained-XML
// coordinate rewriters.

#ifndef FORMULON_IO_SQREF_TEXT_H_
#define FORMULON_IO_SQREF_TEXT_H_

#include <string>
#include <string_view>
#include <vector>

#include "sheet.h"

namespace formulon::io {

/// Appends the space-separated A1 ranges of `text` to `out`. False, with
/// `out` partly filled, when a token does not parse or `text` names none.
bool parse_sqref_text(std::string_view text, std::vector<MergeRange>& out);

/// Appends the ranges of a `<selection sqref>`: A1 forms plus whole-row
/// (`3:4`) and whole-column (`A:B`) tokens widened to the full sheet. A token
/// that does not parse, is inverted or leaves the grid is skipped.
void parse_selection_sqref_text(std::string_view text, std::vector<MergeRange>& out);

/// The A1 text of `ranges`, a single cell written without `:`.
std::string format_sqref_text(const std::vector<MergeRange>& ranges);

}  // namespace formulon::io

#endif  // FORMULON_IO_SQREF_TEXT_H_
