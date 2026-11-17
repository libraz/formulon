// Adjust guides a drawing part may not carry: Excel refuses to open a file whose `a:avLst` names a guide its
// preset does not define, or names one twice, and drops any guide of a preset that defines none.

#ifndef FORMULON_IO_OOXML_DRAWING_GUIDES_H_
#define FORMULON_IO_OOXML_DRAWING_GUIDES_H_

#include <cstddef>
#include <optional>
#include <string_view>

#include "passthrough_part.h"

namespace formulon {
namespace drawing_guides {

/// A copy of `part` with every adjust guide its preset does not define, and every repeated one, left out; unset
/// when `content_type` (the part's effective type) is not a drawing's or the part has none to leave out, so the
/// part is written as it is. `*refused` grows by the guides left out of a preset that defines guides, which Excel
/// refuses to open; a guide of a preset without any is left out uncounted, as Excel leaves it out of its own save.
std::optional<PassthroughPart> without_undefined_guides(const PassthroughPart& part, std::string_view content_type,
                                                        std::size_t* refused);

}  // namespace drawing_guides
}  // namespace formulon

#endif  // FORMULON_IO_OOXML_DRAWING_GUIDES_H_
