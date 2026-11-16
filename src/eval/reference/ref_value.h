//
// Construction helpers for `Ref` values: a one-area reference from parsed
// endpoints, a 3-D span, a union and a grid. They resolve sheet qualifiers
// against the context's workbook and delegate the storage choice to
// `make_ref`.

#ifndef FORMULON_EVAL_REFERENCE_REF_VALUE_H_
#define FORMULON_EVAL_REFERENCE_REF_VALUE_H_

#include <cstdint>
#include <string_view>

#include "eval/declared_rect.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {

class EvalContext;

/// The `RefArea` covering `rect` on the sheets `sheet_first`..`sheet_last`
/// (workbook order; `RefArea::kHomeSheet` for the formula's own sheet). A
/// whole-axis rectangle sets `whole_col` / `whole_row` by the axis it spans.
RefArea make_ref_area(const DeclaredRect& rect, std::uint16_t sheet_first, std::uint16_t sheet_last,
                      RefOrigin origin) noexcept;

/// A `kArea` reference over the rectangle `[lhs : rhs]` (pass the same
/// reference twice for a single cell or a whole-axis token). A bare side
/// inherits the other side's sheet qualifier. A qualifier that does not name
/// a sheet of the workbook, or two different ones, gives `#REF!`; an endpoint
/// pair that declares no rectangle gives `#VALUE!`. A bare reference gets
/// `RefArea::kHomeSheet`. Returns the error as a `Value`.
Value make_area_ref(const parser::Reference& lhs, const parser::Reference& rhs, RefOrigin origin, Arena& arena,
                    const EvalContext& ctx);

/// A `kSpan3D` reference over the rectangle `[lhs : rhs]` on every sheet from
/// `sheet_begin` to `sheet_end`. A missing end sheet beside a present begin
/// sheet is `#NAME?`; any other missing sheet is `#REF!`. A span whose two
/// sheets coincide is a `kArea`.
Value make_span3d_ref(std::string_view sheet_begin, std::string_view sheet_end, const parser::Reference& lhs,
                      const parser::Reference& rhs, RefOrigin origin, Arena& arena, const EvalContext& ctx);

/// A `kUnion` of `count` own-workbook areas (one area is returned as a
/// `kArea`).
Value make_union_ref(Arena& arena, const RefArea* areas, std::uint32_t count);

/// A `kGrid` of `rows` x `cols` own-workbook elements, row-major. An element
/// is a rectangle on one sheet or an error element.
Value make_grid_ref(Arena& arena, const RefArea* elements, std::uint32_t rows, std::uint32_t cols);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_REFERENCE_REF_VALUE_H_
