//
// Which external links a save keeps: Excel drops the links nothing in the workbook references.

#ifndef FORMULON_IO_EXTERNAL_LINK_USE_H_
#define FORMULON_IO_EXTERNAL_LINK_USE_H_

#include <cstdint>
#include <vector>

namespace formulon {

class Workbook;

namespace io {

/// The `ExternalLinkRecord::index` values a save of `wb` writes, ascending. Excel's save drops every link nothing in
/// the workbook references and renumbers the rest densely. A holder written verbatim cannot follow a renumbering,
/// so every link up to the highest index such a holder names is kept.
std::vector<std::uint32_t> written_external_link_indices(const Workbook& wb);

/// Renumbers `wb`'s external links in first-use order: defined names in stored order, then the cell formulas of the
/// active sheet (`Workbook::active_sheet`), then of the other sheets in tab order, row-major within a sheet. A link
/// a holder that cannot follow a renumbering names (conditional-format and validation formulas, tables, extension
/// lists, raw parts, OLE / DDE links, retained xlsb sheet tails) or that has no target keeps its index, as does every
/// lower one; the others take the remaining indices in first-use order, unused links last. Leaves `wb` untouched
/// when that is the order already. An .xlsx load calls it; an .xlsb load keeps the stored order.
void renumber_external_links_by_first_use(Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_EXTERNAL_LINK_USE_H_
