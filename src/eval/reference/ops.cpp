//
// The reference operators over `Ref` values. See `ops.h`.

#include "eval/reference/ops.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "eval/eval_context.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

enum class Operator : std::uint8_t { kRange, kIntersect, kUnion };

RefOrigin origin_of(const RefArea& a) noexcept {
  return static_cast<RefOrigin>(a.origin);
}

// `RefArea` stores these in bit-fields; the masks state the widths.
RefOrigin merged_origin(const RefArea& a, const RefArea& b) noexcept {
  if (origin_of(a) == RefOrigin::kSpelled || origin_of(b) == RefOrigin::kSpelled) {
    return RefOrigin::kSpelled;
  }
  if (origin_of(a) == RefOrigin::kDerived || origin_of(b) == RefOrigin::kDerived) {
    return RefOrigin::kDerived;
  }
  return RefOrigin::kImplicit;
}

// The index of the formula's own sheet, or `kHomeSheet` when the context does
// not locate it in a workbook.
std::uint16_t home_sheet_index(const EvalContext& ctx) noexcept {
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return RefArea::kHomeSheet;
  }
  for (std::size_t i = 0; i < wb->sheet_count() && i < RefArea::kErrorElement; ++i) {
    if (&wb->sheet(i) == current) {
      return static_cast<std::uint16_t>(i);
    }
  }
  return RefArea::kHomeSheet;
}

// Checks that `v` is a `Ref` of a form the operator takes. On failure `*out`
// is the error to return: an error operand as it is, `#NAME?` for a 3-D span
// under `,`, `#VALUE!` otherwise.
bool take_operand(const Value& v, Operator op, RefView* view, Value* out) {
  if (!v.is_ref()) {
    *out = v.is_error() ? v : Value::error(ErrorCode::Value);
    return false;
  }
  *view = v.as_ref();
  switch (view->form) {
    case RefForm::kArea:
      return true;
    case RefForm::kUnion:
      if (op == Operator::kRange) {
        break;
      }
      return true;
    case RefForm::kSpan3D:
      *out = Value::error(op == Operator::kUnion ? ErrorCode::Name : ErrorCode::Value);
      return false;
    case RefForm::kGrid:
      break;
  }
  *out = Value::error(ErrorCode::Value);
  return false;
}

// The sheet a `:` box over `a` and `b` lies on; false when they name two
// different sheets.
bool range_sheet(const RefArea& a, const RefArea& b, std::uint16_t book, const EvalContext& ctx, std::uint16_t* out) {
  if (a.sheet_first == b.sheet_first) {
    *out = a.sheet_first;
    return true;
  }
  const bool a_bare = a.sheet_first == RefArea::kHomeSheet;
  const bool b_bare = b.sheet_first == RefArea::kHomeSheet;
  if (a_bare == b_bare) {
    return false;
  }
  const RefArea& qualified = a_bare ? b : a;
  const bool own_sheet = book == 0 && qualified.sheet_first == home_sheet_index(ctx);
  if (origin_of(qualified) == RefOrigin::kDerived && !own_sheet) {
    return false;
  }
  *out = qualified.sheet_first;
  return true;
}

// The sheet an intersection of `a` and `b` lies on; false when they name two
// different sheets.
bool intersect_sheet(const RefArea& a, const RefArea& b, std::uint16_t* out) {
  if (a.sheet_first == b.sheet_first || b.sheet_first == RefArea::kHomeSheet) {
    *out = a.sheet_first;
    return true;
  }
  if (a.sheet_first == RefArea::kHomeSheet) {
    *out = b.sheet_first;
    return true;
  }
  return false;
}

bool spans_all_rows(const RefArea& a) noexcept {
  return a.row_first == 0U && a.row_last == Sheet::kMaxRows - 1U;
}

bool spans_all_cols(const RefArea& a) noexcept {
  return a.col_first == 0U && a.col_last == Sheet::kMaxCols - 1U;
}

Value finish(Arena& arena, RefForm form, std::uint16_t book, const RefArea* areas, std::uint32_t count, bool dynamic) {
  const Value result = make_ref(arena, form, book, areas, count);
  return dynamic ? result.with_ref_dynamic() : result;
}

}  // namespace

Value ref_range(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx) {
  RefView a{};
  RefView b{};
  Value err = Value::blank();
  if (!take_operand(lhs, Operator::kRange, &a, &err) || !take_operand(rhs, Operator::kRange, &b, &err)) {
    return err;
  }
  std::uint16_t sheet = RefArea::kHomeSheet;
  if (a.book != b.book || !range_sheet(a.areas[0], b.areas[0], a.book, ctx, &sheet)) {
    return Value::error(ErrorCode::Value);
  }
  const RefArea& x = a.areas[0];
  const RefArea& y = b.areas[0];
  RefArea box{};
  box.row_first = std::min<std::uint32_t>(x.row_first, y.row_first);
  box.row_last = std::max<std::uint32_t>(x.row_last, y.row_last);
  box.col_first = static_cast<std::uint16_t>(std::min<std::uint32_t>(x.col_first, y.col_first) & RefArea::kColumnMask);
  box.col_last = static_cast<std::uint16_t>(std::max<std::uint32_t>(x.col_last, y.col_last) & RefArea::kColumnMask);
  box.origin = static_cast<std::uint16_t>(static_cast<std::uint32_t>(merged_origin(x, y)) & RefArea::kOriginMask);
  box.whole_col = static_cast<std::uint16_t>((x.whole_col || y.whole_col) && spans_all_rows(box));
  box.whole_row = static_cast<std::uint16_t>((x.whole_row || y.whole_row) && spans_all_cols(box));
  box.sheet_first = sheet;
  box.sheet_last = sheet;
  return finish(arena, RefForm::kArea, a.book, &box, 1, lhs.ref_is_dynamic() || rhs.ref_is_dynamic());
}

Value ref_intersect(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx) {
  (void)ctx;
  RefView a{};
  RefView b{};
  Value err = Value::blank();
  if (!take_operand(lhs, Operator::kIntersect, &a, &err) || !take_operand(rhs, Operator::kIntersect, &b, &err)) {
    return err;
  }
  if (a.book != b.book) {
    return Value::error(ErrorCode::Value);
  }
  std::vector<RefArea> cells;
  for (std::uint32_t i = 0; i < a.area_count; ++i) {
    for (std::uint32_t j = 0; j < b.area_count; ++j) {
      const RefArea& x = a.areas[i];
      const RefArea& y = b.areas[j];
      std::uint16_t sheet = RefArea::kHomeSheet;
      if (!intersect_sheet(x, y, &sheet)) {
        return Value::error(ErrorCode::Ref);
      }
      const std::uint32_t row_first = std::max<std::uint32_t>(x.row_first, y.row_first);
      const std::uint32_t row_last = std::min<std::uint32_t>(x.row_last, y.row_last);
      const std::uint32_t col_first = std::max<std::uint32_t>(x.col_first, y.col_first);
      const std::uint32_t col_last = std::min<std::uint32_t>(x.col_last, y.col_last);
      if (row_first > row_last || col_first > col_last) {
        continue;
      }
      RefArea part{};
      part.row_first = row_first;
      part.row_last = row_last;
      part.col_first = static_cast<std::uint16_t>(col_first & RefArea::kColumnMask);
      part.col_last = static_cast<std::uint16_t>(col_last & RefArea::kColumnMask);
      part.origin = static_cast<std::uint16_t>(static_cast<std::uint32_t>(merged_origin(x, y)) & RefArea::kOriginMask);
      part.whole_col = static_cast<std::uint16_t>(x.whole_col && y.whole_col);
      part.whole_row = static_cast<std::uint16_t>(x.whole_row && y.whole_row);
      part.sheet_first = sheet;
      part.sheet_last = sheet;
      cells.push_back(part);
    }
  }
  if (cells.empty()) {
    return Value::error(ErrorCode::Null);
  }
  const std::uint32_t count = static_cast<std::uint32_t>(cells.size());
  return finish(arena, count == 1U ? RefForm::kArea : RefForm::kUnion, a.book, cells.data(), count,
                lhs.ref_is_dynamic() || rhs.ref_is_dynamic());
}

Value ref_union(const Value& lhs, const Value& rhs, Arena& arena, const EvalContext& ctx) {
  (void)ctx;
  RefView a{};
  RefView b{};
  Value err = Value::blank();
  // A 3-D operand outranks a grid one, whichever side each is on.
  const bool lhs_span = lhs.is_ref() && lhs.ref_form() == RefForm::kSpan3D;
  const bool rhs_span = rhs.is_ref() && rhs.ref_form() == RefForm::kSpan3D;
  if (lhs_span || rhs_span) {
    return Value::error(ErrorCode::Name);
  }
  if (!take_operand(lhs, Operator::kUnion, &a, &err) || !take_operand(rhs, Operator::kUnion, &b, &err)) {
    return err;
  }
  if (a.book != b.book) {
    return Value::error(ErrorCode::Value);
  }
  for (std::uint32_t i = 0; i < a.area_count; ++i) {
    for (std::uint32_t j = 0; j < b.area_count; ++j) {
      const RefArea& x = a.areas[i];
      const RefArea& y = b.areas[j];
      const bool spelled = origin_of(x) == RefOrigin::kSpelled && origin_of(y) == RefOrigin::kSpelled;
      if (spelled && x.sheet_first != RefArea::kHomeSheet && y.sheet_first != RefArea::kHomeSheet &&
          x.sheet_first != y.sheet_first) {
        return Value::error(ErrorCode::Value);
      }
    }
  }
  std::vector<RefArea> areas;
  areas.reserve(static_cast<std::size_t>(a.area_count) + b.area_count);
  areas.insert(areas.end(), a.areas, a.areas + a.area_count);
  areas.insert(areas.end(), b.areas, b.areas + b.area_count);
  return finish(arena, RefForm::kUnion, a.book, areas.data(), static_cast<std::uint32_t>(areas.size()),
                lhs.ref_is_dynamic() || rhs.ref_is_dynamic());
}

}  // namespace eval
}  // namespace formulon
