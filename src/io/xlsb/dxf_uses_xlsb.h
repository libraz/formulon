//
// The dxf references an xlsb save writes. Every stream the writer emits is
// encoded with model dxf indices first; the collector reads the dxf slots of
// exactly those bytes, so a dxf is used iff some written record names it, and
// the plan built from the uses is then applied to the same bytes. A written
// stream holding a dxf-bearing record no reader decodes (a retained pivot
// table part), a slot that cannot be located, or a raw styles part whose dxf
// list is not the model's keeps the dxf table as it is.

#ifndef FORMULON_IO_XLSB_DXF_USES_XLSB_H_
#define FORMULON_IO_XLSB_DXF_USES_XLSB_H_

#include <cstddef>
#include <cstdint>
#include <vector>

#include "io/dxf_save_plan.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "workbook.h"

namespace formulon::io::xlsb {

/// One record stream an xlsb save writes; the records from byte `begin` on
/// are the ones that can carry a dxf slot.
struct XlsbDxfStream {
  std::vector<std::uint8_t>* bytes = nullptr;
  std::size_t begin = 0;
  std::int32_t sheet = -1;  ///< Tab index; -1 when no sheet relates the part.
  std::uint32_t part = 0;   ///< Table ordinal within the sheet; 0 for a sheet body.
};

/// The streams an xlsb save writes that can name a model dxf.
struct XlsbEmission {
  std::vector<XlsbDxfStream> streams;
  /// The source styles.bin written in place of a generated one; null when generated.
  std::vector<std::uint8_t>* raw_styles = nullptr;
  /// A retained pivot table part is written.
  bool retained_pivot = false;
};

/// The dxf references `emission` writes against `wb`'s dxf table.
DxfUseSet collect_xlsb_dxf_uses(const Workbook& wb, const XlsbEmission& emission);

/// Maps every dxf slot of `emission` through `plan`, a slot naming no model dxf
/// becoming kNoDxf, and writes the raw styles part's dxf list in
/// `plan.order()`. Under an identity plan a stream whose slots all name model
/// dxfs keeps every byte. Not called for a blocked use set, whose ids stay as
/// written.
Expected<void, Error> apply_xlsb_dxf_plan(XlsbEmission& emission, const DxfSavePlan& plan);

}  // namespace formulon::io::xlsb

#endif  // FORMULON_IO_XLSB_DXF_USES_XLSB_H_
