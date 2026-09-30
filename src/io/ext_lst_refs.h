//
// Coordinate rewriting inside a raw worksheet `<extLst>` for row / column
// structural edits.
//
// The extension payload is retained verbatim, but its `<xm:sqref>` and
// `<xm:f>` elements name cells: the x14 conditional-format ranges, each
// sparkline's location and source, the x14 data-validation ranges. Excel
// moves them with the edit (measured on Windows Excel 365: an x14 CF
// `xm:sqref` mirrors the legacy union sqref and moves with it, a sparkline's
// `xm:f` and `xm:sqref` move like ordinary references), so these helpers
// let the structural-edit pass apply the same mapping the modelled
// structures use. Only those two element kinds are touched; everything
// else keeps its bytes, and an extension nothing changed is not
// re-serialised.

#ifndef FORMULON_IO_EXT_LST_REFS_H_
#define FORMULON_IO_EXT_LST_REFS_H_

#include <functional>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "sheet.h"

namespace formulon::io {

/// Edits one sqref's rectangles in place; leaving it empty removes the
/// element that owns the sqref.
using SqrefRemap = std::function<void(std::vector<MergeRange>&)>;

/// Returns the rewritten text of one formula (no leading `=`), or nothing
/// when it is unchanged.
using FormulaRemap = std::function<std::optional<std::string>(std::string_view)>;

/// Maps every `<xm:sqref>` in `xml` through `remap`. An emptied sqref
/// removes its owner -- the `<x14:sparkline>` (and its group once no
/// sparkline is left), the `<x14:conditionalFormatting>`, or whatever
/// other element holds it -- and any container that leaves empty; a
/// container's `count` attribute is kept in step. An sqref that does not
/// parse is left alone. Returns true when `xml` changed.
bool remap_ext_lst_sqrefs(std::string& xml, const SqrefRemap& remap);

/// Maps the text of every `<xm:f>` in `xml` through `remap`. Returns true
/// when `xml` changed.
bool remap_ext_lst_formulas(std::string& xml, const FormulaRemap& remap);

}  // namespace formulon::io

#endif  // FORMULON_IO_EXT_LST_REFS_H_
