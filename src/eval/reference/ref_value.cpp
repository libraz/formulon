//
// Construction helpers for `Ref` values. See `ref_value.h`.

#include "eval/reference/ref_value.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// Sheet indices at or above this collide with the `RefArea` markers.
constexpr std::size_t kSheetIndexLimit = RefArea::kErrorElement;

// The index of the sheet `name` (empty = the formula's own sheet) in `*out`.
bool sheet_index_of(const EvalContext& ctx, std::string_view name, std::uint16_t* out) {
  if (name.empty()) {
    *out = RefArea::kHomeSheet;
    return true;
  }
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return false;
  }
  const std::size_t index = wb->sheet_index_by_name(name);
  if (index == Workbook::kNoSheet || index >= kSheetIndexLimit) {
    return false;
  }
  *out = static_cast<std::uint16_t>(index);
  return true;
}

}  // namespace

RefArea make_ref_area(const DeclaredRect& rect, std::uint16_t sheet_first, std::uint16_t sheet_last,
                      RefOrigin origin) noexcept {
  RefArea a{};
  a.row_first = rect.row_first;
  a.row_last = rect.row_last;
  a.col_first = static_cast<std::uint16_t>(rect.col_first & 0x3FFFU);
  a.col_last = static_cast<std::uint16_t>(rect.col_last & 0x3FFFU);
  a.origin = static_cast<std::uint16_t>(static_cast<std::uint32_t>(origin) & 0x3U);
  a.whole_col = static_cast<std::uint16_t>(rect.whole_axis && rect.rows() == Sheet::kMaxRows);
  a.whole_row = static_cast<std::uint16_t>(rect.whole_axis && rect.cols() == Sheet::kMaxCols);
  a.sheet_first = sheet_first;
  a.sheet_last = sheet_last;
  return a;
}

bool ref_sheet_name(const EvalContext& ctx, std::uint16_t index, std::string_view* name) {
  if (index == RefArea::kHomeSheet) {
    *name = {};
    return true;
  }
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr || index >= wb->sheet_count()) {
    return false;
  }
  *name = wb->sheet(index).name();
  return true;
}

Value make_rect_ref(std::string_view sheet, const DeclaredRect& rect, RefOrigin origin, Arena& arena,
                    const EvalContext& ctx) {
  std::uint16_t index = RefArea::kHomeSheet;
  if (!sheet_index_of(ctx, sheet, &index)) {
    return Value::error(ErrorCode::Ref);
  }
  const RefArea area = make_ref_area(rect, index, index, origin);
  return make_ref(arena, RefForm::kArea, 0, &area, 1);
}

Value make_area_ref(const parser::Reference& lhs, const parser::Reference& rhs, RefOrigin origin, Arena& arena,
                    const EvalContext& ctx) {
  const Expected<DeclaredRect, ErrorCode> rect = ctx.declared_range_rect_or_bare(lhs, rhs);
  if (!rect) {
    return Value::error(rect.error());
  }
  const Expected<std::string_view, ErrorCode> sheet = merge_sheet_qualifiers(lhs.sheet, rhs.sheet, ctx.workbook());
  if (!sheet) {
    return Value::error(ErrorCode::Ref);
  }
  return make_rect_ref(sheet.value(), rect.value(), origin, arena, ctx);
}

Value make_span3d_ref(std::string_view sheet_begin, std::string_view sheet_end, const parser::Reference& lhs,
                      const parser::Reference& rhs, RefOrigin origin, Arena& arena, const EvalContext& ctx) {
  std::uint16_t begin = 0;
  std::uint16_t end = 0;
  const bool have_begin = !sheet_begin.empty() && sheet_index_of(ctx, sheet_begin, &begin);
  const bool have_end = !sheet_end.empty() && sheet_index_of(ctx, sheet_end, &end);
  if (have_begin && !have_end) {
    return Value::error(ErrorCode::Name);
  }
  if (!have_begin || !have_end) {
    return Value::error(ErrorCode::Ref);
  }
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(lhs, rhs);
  if (!rect) {
    return Value::error(rect.error());
  }
  const RefArea area = make_ref_area(rect.value(), std::min(begin, end), std::max(begin, end), origin);
  return make_ref(arena, RefForm::kArea, 0, &area, 1);
}

Value make_union_ref(Arena& arena, const RefArea* areas, std::uint32_t count) {
  if (count == 0U) {
    return Value::error(ErrorCode::Value);
  }
  return make_ref(arena, count == 1U ? RefForm::kArea : RefForm::kUnion, 0, areas, count);
}

Value make_grid_ref(Arena& arena, const RefArea* elements, std::uint32_t rows, std::uint32_t cols) {
  return make_ref(arena, RefForm::kGrid, 0, elements, rows * cols, rows, cols);
}

}  // namespace eval
}  // namespace formulon
