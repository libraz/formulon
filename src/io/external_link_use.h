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

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_EXTERNAL_LINK_USE_H_
