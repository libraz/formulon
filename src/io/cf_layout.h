//
// The conditional-format blocks a save writes. Excel groups the blocks of one
// sheet by sqref, orders the rules of a block by priority and orders the
// blocks by their first area, whatever order the model holds them in. The
// xlsx writer, the xlsb writer and the xlsx dxf collector all walk this view,
// so the written order and the dxf order cannot disagree.

#ifndef FORMULON_IO_CF_LAYOUT_H_
#define FORMULON_IO_CF_LAYOUT_H_

#include <optional>
#include <string>
#include <vector>

#include "cf/cf_types.h"

namespace formulon::io {

/// One block of the saved layout. The pointers borrow from the model blocks
/// passed to `cf_saved_layout` and live as long as they do.
struct CfBlockView {
  const std::vector<cf::CFCellRange>* sqref = nullptr;
  bool pivot_scope = false;
  const std::optional<std::string>* ext_lst_raw = nullptr;
  std::vector<const cf::CFRule*> rules;  ///< By priority, ties in model order.
};

/// The blocks a save writes for `formats`. Model blocks with an identical
/// sqref sequence and equal `pivot_scope`, none carrying `ext_lst_raw`,
/// combine into one view; a block with `ext_lst_raw` stays its own view. Views
/// are ordered by the first area of the sqref as (first column, first row,
/// last column, last row), ties in model order.
std::vector<CfBlockView> cf_saved_layout(const std::vector<cf::ConditionalFormat>& formats);

}  // namespace formulon::io

#endif  // FORMULON_IO_CF_LAYOUT_H_
