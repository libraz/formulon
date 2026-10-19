//
// The `PtgAttrSpace` type a whitespace run carries, shared by the Ptg reader
// and writer so the two map whitespace slots the same way.

#ifndef FORMULON_IO_XLSB_PTG_SPACE_H_
#define FORMULON_IO_XLSB_PTG_SPACE_H_

#include <cstdint>

#include "io/xlsb/ptg.h"
#include "parser/ast_format.h"

namespace formulon {
namespace io {
namespace xlsb {

/// `PtgAttrSpace` types by `parser::WhitespaceSlot`: `[slot][0]` for a run of
/// spaces, `[slot][1]` for a run of line feeds.
inline constexpr std::uint8_t kPtgSpaceTypeBySlot[4][2] = {
    {kPtgSpaceSpaces, kPtgSpaceLineFeeds},                  // kBeforeText
    {kPtgSpaceSpaces, kPtgSpaceLineFeeds},                  // kBeforeOperator
    {kPtgSpaceBeforeClose, kPtgSpaceLineFeedsBeforeClose},  // kBeforeClose
    {kPtgSpaceBeforeOpen, kPtgSpaceBeforeOpen},             // kBeforeOpen
};

/// The `PtgAttrSpace` type of a run at `slot`: line feeds when `feeds`.
inline std::uint8_t ptg_space_type(parser::WhitespaceSlot slot, bool feeds) noexcept {
  return kPtgSpaceTypeBySlot[static_cast<std::uint8_t>(slot)][feeds ? 1 : 0];
}

}  // namespace xlsb
}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_XLSB_PTG_SPACE_H_
