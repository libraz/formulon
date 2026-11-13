//
// The main-table dxf references an .xlsx save writes, collected from the
// same holders and predicates the xlsx emitters use. The uses of each part
// follow the dxf attributes the emitted part holds, in document order, which
// is what lets the writer rewrite the part reference by reference.
//
// Every CF rule is a user, an x14-only one included: Excel keeps its inline
// format in `<dxfs>` as well, with no index in the worksheet. Raw XML is
// scanned with `dxf_xml_ids.h`; a fragment it cannot read blocks the plan
// (identity). The styles part's `<extLst>` is not scanned: its slicer and
// timeline style elements index the extension dxfs. It blocks the plan when
// such an element names a `dxfId` and no `x14:dxfs` / `x15:dxfs` exists.

#ifndef FORMULON_IO_DXF_USES_XLSX_H_
#define FORMULON_IO_DXF_USES_XLSX_H_

#include "io/dxf_save_plan.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Workbook;
namespace io {
struct EmissionPlan;

/// Every main-table dxf reference `write_ooxml` writes for `wb` under `plan`,
/// keyed by part: a worksheet body `{sheet, kWorksheet, 0}`, the pivot and
/// table parts by their index in `plan`'s per-sheet lists, the styles part
/// `{-1, kStyles, 0}`.
Expected<DxfUseSet, Error> collect_xlsx_dxf_uses(const Workbook& wb, const EmissionPlan& plan);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DXF_USES_XLSX_H_
