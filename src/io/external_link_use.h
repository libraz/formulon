//
// Which external links a save keeps: Excel drops the trailing links nothing in the workbook references.

#ifndef FORMULON_IO_EXTERNAL_LINK_USE_H_
#define FORMULON_IO_EXTERNAL_LINK_USE_H_

#include <cstdint>

namespace formulon {

class Workbook;

namespace io {

/// The highest `ExternalLinkRecord::index` a save of `wb` writes: Excel's save drops the links after the last one
/// anything in the workbook references. 0 when no link is referenced.
std::uint32_t last_written_external_link(const Workbook& wb);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_EXTERNAL_LINK_USE_H_
