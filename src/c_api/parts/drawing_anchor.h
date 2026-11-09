//
// Drawing anchor resolution shared by the drawing parts: the add rules of
// `fm_drawing_anchor_t`, the merge rules of `fm_drawing_anchor_patch_t`, and
// the read of an object's anchor. Defined in `drawing.cpp`.

#ifndef FORMULON_C_API_PARTS_DRAWING_ANCHOR_H_
#define FORMULON_C_API_PARTS_DRAWING_ANCHOR_H_

#include <cstddef>
#include <cstdint>

#include "c_api/formulon_c.h"
#include "drawing/drawing_xml.h"

namespace formulon {
class Workbook;
namespace c_api {
namespace parts {

/// A resolved anchor in the engine's terms. `to` is meaningful for a two-cell
/// anchor, `cx` / `cy` for every kind (a two-cell anchor's box size).
struct AnchorPlan {
  AnchorKind kind = AnchorKind::kOneCell;
  EditAs edit_as = EditAs::kTwoCell;
  AnchorPoint from;
  AnchorPoint to;
  std::int64_t cx = 0;
  std::int64_t cy = 0;
};

/// Resolves the anchor of an add. `default_kind` stands for
/// `FM_ANCHOR_KIND_DEFAULT`; `size_required` refuses a size zero in both directions where the
/// object has no default one. A two-cell anchor with a non-zero `to` marker
/// takes its size from the box. Returns the status to report.
fm_status_t plan_add_anchor(const Workbook& wb, std::size_t sheet, const fm_drawing_anchor_t& anchor,
                            AnchorKind default_kind, bool size_required, AnchorPlan* out);

/// Applies what the engine's add entries do not take from a placement to the
/// new object `object_id`: `hidden` (`FM_TRISTATE_TRUE`) and, when
/// `with_description`, its non-empty `description`. Returns the status to report.
fm_status_t apply_placement_extras(Workbook& wb, std::size_t sheet, std::uint32_t object_id,
                                   const fm_drawing_placement_spec_t& placement, bool with_description);

/// Fills `*out` with the anchor of `obj` as the reads report it.
void fill_drawing_anchor(const Workbook& wb, std::size_t sheet, const DrawingObject& obj, fm_drawing_anchor_t* out);

}  // namespace parts
}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_PARTS_DRAWING_ANCHOR_H_
