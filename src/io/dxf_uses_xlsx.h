//
// The main-table dxf references an .xlsx save writes, collected from the
// same holders and predicates the xlsx emitters use, plus the helpers those
// emitters apply to write a reference through a `DxfSavePlan`.
//
// Every CF rule is a user, an x14-only one included: Excel keeps its inline
// format in `<dxfs>` as well. Raw XML is scanned with `dxf_xml_ids.h`; a
// fragment it cannot read blocks the plan (identity). The styles part's `<extLst>` is not scanned: its slicer
// and timeline style elements index the extension dxfs. It blocks the plan
// when such an element names a `dxfId` and no `x14:dxfs` / `x15:dxfs` exists.

#ifndef FORMULON_IO_DXF_USES_XLSX_H_
#define FORMULON_IO_DXF_USES_XLSX_H_

#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <vector>

#include "io/dxf_save_plan.h"
#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
class Workbook;
namespace io {
struct EmissionPlan;

/// Every main-table dxf reference `write_ooxml` writes for `wb` under `plan`.
Expected<DxfUseSet, Error> collect_xlsx_dxf_uses(const Workbook& wb, const EmissionPlan& plan);

/// The id a modelled reference is written as: unchanged under a null plan,
/// else `plan->map`, which is nullopt for an id the save does not write.
std::optional<std::uint32_t> saved_dxf_id(std::optional<std::uint32_t> model_id, const DxfSavePlan* plan);

/// `xml` with its dxf attributes written through `plan`; unchanged under a
/// null plan or when the markup cannot be read (the plan is then identity).
std::string saved_dxf_xml(std::string_view xml, const DxfSavePlan* plan);

/// `attrs` (` name="value"` pairs) written as `saved_dxf_xml` writes a fragment.
std::string saved_dxf_attrs(std::string_view attrs, const DxfSavePlan* plan);

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_DXF_USES_XLSX_H_
