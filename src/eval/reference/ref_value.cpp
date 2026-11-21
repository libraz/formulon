//
// Construction helpers for `Ref` values. See `ref_value.h`.

#include "eval/reference/ref_value.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

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

Value grid_element(const RefView& view, std::size_t index, Arena& arena) {
  Value own = Value::blank();
  if (grid_element_value(view, index, &own)) {
    return own;
  }
  const std::uint16_t book = view.element_books != nullptr ? view.element_books[index] : view.book;
  return make_ref(arena, RefForm::kArea, book, &view.areas[index], 1U);
}

namespace {

// Copies a rows x cols grid out of `view`; `from_index(r, c)` names the source element index for each target cell.
template <typename From>
Value grid_copy(const RefView& view, std::uint32_t rows, std::uint32_t cols, Arena& arena, From from_index) {
  const std::size_t n = static_cast<std::size_t>(rows) * cols;
  std::vector<RefArea> areas(n);
  std::vector<std::uint16_t> books(n);
  std::vector<Value> values(n, Value::blank());
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      const std::size_t from = from_index(r, c);
      const std::size_t to = static_cast<std::size_t>(r) * cols + c;
      areas[to] = view.areas[from];
      books[to] = view.element_books != nullptr ? view.element_books[from] : view.book;
      if (view.element_values != nullptr) {
        values[to] = view.element_values[from];
      }
    }
  }
  return make_ref(arena, RefForm::kGrid, view.book, areas.data(), static_cast<std::uint32_t>(n), rows, cols,
                  view.element_books != nullptr ? books.data() : nullptr,
                  view.element_values != nullptr ? values.data() : nullptr);
}

}  // namespace

Value grid_slice(const Value& grid, std::uint32_t row_first, std::uint32_t col_first, std::uint32_t rows,
                 std::uint32_t cols, Arena& arena) {
  const RefView view = grid.as_ref();
  return grid_copy(view, rows, cols, arena, [&](std::uint32_t r, std::uint32_t c) {
    return static_cast<std::size_t>(row_first + r) * view.cols + col_first + c;
  });
}

Value grid_lanes(const Value& grid, const std::vector<std::uint32_t>& lanes, bool by_col, Arena& arena) {
  const RefView view = grid.as_ref();
  const std::uint32_t rows = by_col ? view.rows : static_cast<std::uint32_t>(lanes.size());
  const std::uint32_t cols = by_col ? static_cast<std::uint32_t>(lanes.size()) : view.cols;
  return grid_copy(view, rows, cols, arena, [&](std::uint32_t r, std::uint32_t c) {
    return by_col ? static_cast<std::size_t>(r) * view.cols + lanes[c]
                  : static_cast<std::size_t>(lanes[r]) * view.cols + c;
  });
}

bool grid_element_value(const RefView& view, std::size_t index, Value* out) {
  const RefArea& element = view.areas[index];
  if (element.is_error_element()) {
    *out = Value::error(element.error_code());
    return true;
  }
  if (element.is_value_element()) {
    *out = view.element_values[index];
    return true;
  }
  return false;
}

}  // namespace eval
}  // namespace formulon
