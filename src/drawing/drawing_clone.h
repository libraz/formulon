//
// Copies the pictures and charts of one worksheet's drawing onto another
// sheet, rebinding each copied chart's series references.
//
// `plan_drawing_copy` does every fallible read and refusal without touching the
// workbook; `apply_drawing_copy` restores the captured objects onto the target.
// Neither takes a workbook lock, so both run under the caller's.

#ifndef FORMULON_DRAWING_DRAWING_CLONE_H_
#define FORMULON_DRAWING_DRAWING_CLONE_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "parser/ast_shift.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

class Workbook;

/// Snapshots of a sheet's pictures and charts, back to front.
struct DrawingCopyPlan {
  std::vector<std::vector<std::uint8_t>> snapshots;
};

/// Captures every picture and chart of `source_sheet`'s drawing. A sheet
/// without a drawing yields an empty plan. Fails with `kNotImplemented` for a
/// drawing holding any other object (shape, group, connector, text box,
/// slicer frame), a pivot chart, a framed slicer, or a drawing retained from
/// an `.xlsb` package.
Expected<DrawingCopyPlan, Error> plan_drawing_copy(const Workbook& wb, std::size_t source_sheet);

/// Restores `plan` onto `target_sheet` as copies with fresh ids and part
/// paths, then rewrites the references of the chart parts that restore added
/// through `transform`.
Expected<void, Error> apply_drawing_copy(Workbook& wb, std::size_t target_sheet, const DrawingCopyPlan& plan,
                                         const parser::RefTransform& transform);

}  // namespace formulon

#endif  // FORMULON_DRAWING_DRAWING_CLONE_H_
