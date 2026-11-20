//
// Reading a `Ref` as values. See `deref.h`.

#include "eval/reference/deref.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/external_ref.h"
#include "eval/range_args.h"
#include "eval/reference/ref_value.h"
#include "eval/tail_array.h"
#include "eval/tree_walker/walker_internal.h"
#include "external_book.h"
#include "external_link.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/index_sort.h"
#include "utils/resource_budget.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// The cached supporting workbook `book` (1-based) names, or null.
const ExternalBook* external_book_of(const EvalContext& ctx, std::uint16_t book) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr || book == 0U || book > wb->external_links().size()) {
    return nullptr;
  }
  const ExternalLinkRecord& link = wb->external_links()[book - 1U];
  return link.kind == ExternalLinkRecord::Kind::kExternalBook ? &link.book : nullptr;
}

DeclaredRect rect_of(const RefArea& a) noexcept {
  DeclaredRect r;
  r.row_first = a.row_first;
  r.row_last = a.row_last;
  r.col_first = a.col_first;
  r.col_last = a.col_last;
  r.whole_axis = a.whole_col != 0U || a.whole_row != 0U;
  return r;
}

// A rectangle of more than one line that runs to the grid's last row (or last column) from inside the grid,
// short of a whole axis (`A2:A1048576`): it reads as a whole axis does, its populated head followed by blanks.
bool reaches_grid_edge(const RefArea& a, TailAxis* axis) noexcept {
  if (a.whole_col != 0U || a.whole_row != 0U) {
    return false;
  }
  const bool rows = a.row_first != 0U && a.row_first != a.row_last && a.row_last == Sheet::kMaxRows - 1U;
  const bool cols = a.col_first != 0U && a.col_first != a.col_last && a.col_last == Sheet::kMaxCols - 1U;
  *axis = rows ? TailAxis::kRows : TailAxis::kCols;
  return rows || cols;
}

bool is_one_cell(const RefArea& a) noexcept {
  return a.row_first == a.row_last && a.col_first == a.col_last && a.sheet_first == a.sheet_last;
}

// The corner pair `EvalContext::expand_range` takes for `a` on sheet `sheet`.
// A rectangle spanning a whole axis is flagged so it reads as one.
void corner_refs(const RefArea& a, std::string_view sheet, parser::Reference* lo, parser::Reference* hi) {
  set_rect_corner_refs(sheet, a.row_first, a.col_first, a.row_last, a.col_last, lo, hi);
  const bool full_col = a.whole_col != 0U || (a.row_first == 0U && a.row_last == Sheet::kMaxRows - 1U);
  const bool full_row = !full_col && (a.whole_row != 0U || (a.col_first == 0U && a.col_last == Sheet::kMaxCols - 1U));
  lo->is_full_col = hi->is_full_col = full_col;
  lo->is_full_row = hi->is_full_row = full_row;
}

// A cell read from a supporting workbook's cache, a blank as a range read.
Value external_cell(const ExternalBook& book, std::uint32_t sheet, std::uint32_t row, std::uint32_t col, Arena& arena) {
  const Value v = read_external_cell(book, sheet, row, col, arena);
  return v.is_blank() ? Value::blank(BlankGridProjection::kReferenceGridZero) : v;
}

// The rectangle a read of `a` walks in the cache of `book`: a whole axis is
// clipped to the cached extent, or to its first line when nothing is cached.
DeclaredRect external_walked(const ExternalBook& book, std::uint32_t sheet, const RefArea& a) {
  const DeclaredRect declared = rect_of(a);
  DeclaredRect walked = declared;
  if (!declared.whole_axis) {
    return walked;
  }
  std::uint32_t last_row = 0;
  std::uint32_t last_col = 0;
  const bool cached = book.cached_extent(sheet, &last_row, &last_col);
  if (declared.rows() == Sheet::kMaxRows) {
    walked.row_first = 0;
    walked.row_last = cached ? last_row : 0U;
  } else {
    walked.col_first = 0;
    walked.col_last = cached ? last_col : 0U;
  }
  return walked;
}

std::uint16_t element_book(const RefView& view, std::size_t index) noexcept {
  return view.element_books != nullptr ? view.element_books[index] : view.book;
}

// Reads the areas of one `Ref` and collects the dynamic reads they make.
class Reader {
 public:
  Reader(Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx, ReadCount count, bool dynamic)
      : arena_(arena),
        registry_(registry),
        ctx_(count == ReadCount::kCounted ? ctx : ctx.without_read_observer()),
        dynamic_(dynamic) {}

  // Appends the cells of `a` (every sheet of a span) to `out`. `rows` / `cols`
  // receive the shape of a one-sheet read and may be null.
  bool read_cells(const RefArea& a, std::uint16_t book, std::vector<Value>* out, std::uint32_t* rows,
                  std::uint32_t* cols, ErrorCode* err) {
    const std::uint16_t first = std::min(a.sheet_first, a.sheet_last);
    const std::uint16_t last = std::max(a.sheet_first, a.sheet_last);
    const Workbook* wb = ctx_.workbook();
    for (std::uint32_t s = first; s <= last; ++s) {
      // A chartsheet or macro sheet inside a span is left out of it.
      if (first != last && book == 0U &&
          (wb->sheet(s).kind() == SheetKind::kChartsheet || wb->sheet(s).kind() == SheetKind::kMacrosheet)) {
        continue;
      }
      if (!sheet_cells(a, static_cast<std::uint16_t>(s), book, out, rows, cols, err)) {
        return false;
      }
    }
    return true;
  }

  // Reads the one-sheet area `a` as the walker reads the same range.
  Shaped read_shaped(const RefArea& a, std::uint16_t book) {
    if (book != 0U) {
      return external_shaped(a, book);
    }
    std::string_view name;
    if (!ref_sheet_name(ctx_, a.sheet_first, &name)) {
      return Shaped{Value::error(ErrorCode::Ref), nullptr};
    }
    parser::Reference lo{};
    parser::Reference hi{};
    corner_refs(a, name, &lo, &hi);
    note(a.sheet_first, rect_of(a));
    if (lo.is_full_col || lo.is_full_row) {
      return read_whole_axis(lo, hi, arena_, registry_, ctx_);
    }
    if (TailAxis axis = TailAxis::kRows; reaches_grid_edge(a, &axis)) {
      return read_edge_tail(a, name, axis);
    }
    if (is_one_cell(a)) {
      return Shaped{ctx_.resolve_ref(lo, arena_, registry_), nullptr};
    }
    return Shaped{materialize_rectangle(lo, hi, arena_, registry_, ctx_), nullptr};
  }

  // The populated head of the edge-reaching `a` on sheet `name`: the band's whole axis is walked as a whole
  // column (row) is, and the head is its part from `a`'s first row (column); empty when none of `a` is populated.
  Expected<std::optional<DeclaredRect>, ErrorCode> edge_head(const RefArea& a, std::string_view name, TailAxis axis) {
    parser::Reference lo{};
    parser::Reference hi{};
    set_rect_corner_refs(name, a.row_first, a.col_first, a.row_last, a.col_last, &lo, &hi);
    (axis == TailAxis::kRows ? lo.is_full_col : lo.is_full_row) = true;
    (axis == TailAxis::kRows ? hi.is_full_col : hi.is_full_row) = true;
    const auto walked = ctx_.walked_range_rect(lo, hi);
    if (!walked) {
      return walked.error();
    }
    if (!walked.value().has_value()) {
      return std::optional<DeclaredRect>{};
    }
    DeclaredRect head = rect_of(a);
    std::uint32_t& last = axis == TailAxis::kRows ? head.row_last : head.col_last;
    const std::uint32_t first = axis == TailAxis::kRows ? head.row_first : head.col_first;
    last = axis == TailAxis::kRows ? walked.value()->row_last : walked.value()->col_last;
    if (last < first) {
      return std::optional<DeclaredRect>{};
    }
    return std::optional<DeclaredRect>{head};
  }

  // Reads the edge-reaching `a` as a whole axis reads: its populated head, then blanks to its declared size.
  Shaped read_edge_tail(const RefArea& a, std::string_view name, TailAxis axis) {
    const auto head = edge_head(a, name, axis);
    if (!head) {
      return Shaped{Value::error(head.error()), nullptr};
    }
    std::vector<Value> cells;
    std::uint32_t head_len = 0U;
    if (head.value().has_value()) {
      const DeclaredRect& h = *head.value();
      parser::Reference lo{};
      parser::Reference hi{};
      set_rect_corner_refs(name, h.row_first, h.col_first, h.row_last, h.col_last, &lo, &hi);
      auto expanded = ctx_.expand_range(lo, hi, arena_, registry_);
      if (!expanded) {
        return Shaped{Value::error(expanded.error()), nullptr};
      }
      cells = std::move(expanded.value());
      head_len = axis == TailAxis::kRows ? h.rows() : h.cols();
    }
    Value* stored = cells.empty() ? nullptr : arena_.create_array<Value>(cells.size());
    if (!cells.empty() && stored == nullptr) {
      return Shaped{Value::error(ErrorCode::Num), nullptr};
    }
    for (std::size_t i = 0; i < cells.size(); ++i) {
      stored[i] = cells[i];
    }
    const DeclaredRect declared = rect_of(a);
    return make_reference_tail(arena_, declared.rows(), declared.cols(), head_len, axis, stored);
  }

  // Reads element `i` of the grid `view`: an error or value element is itself,
  // an area its top-left cell.
  Value read_element(const RefView& view, std::size_t i) {
    Value own = Value::blank();
    return grid_element_value(view, i, &own) ? own : read_top_left(view.areas[i], element_book(view, i));
  }

  // Reads the top-left cell of `a` (its first sheet); an error element is its
  // error.
  Value read_top_left(const RefArea& a, std::uint16_t book) {
    if (a.is_error_element()) {
      return Value::error(a.error_code());
    }
    const std::uint16_t sheet = std::min(a.sheet_first, a.sheet_last);
    if (book != 0U) {
      const ExternalBook* ext = external_book_of(ctx_, book);
      if (ext == nullptr || !ext->sheet_readable(sheet)) {
        return Value::error(ErrorCode::Ref);
      }
      return external_cell(*ext, sheet, a.row_first, a.col_first, arena_);
    }
    std::string_view name;
    if (!ref_sheet_name(ctx_, sheet, &name)) {
      return Value::error(ErrorCode::Ref);
    }
    parser::Reference cell{};
    cell.sheet = name;
    cell.row = a.row_first;
    cell.col = a.col_first;
    DeclaredRect rect;
    rect.row_first = rect.row_last = a.row_first;
    rect.col_first = rect.col_last = a.col_first;
    note(sheet, rect);
    return ctx_.resolve_ref(cell, arena_, registry_);
  }

  // Emits the collected dynamic reads: each distinct (sheet, rectangle) once.
  void flush() {
    if (notes_.empty()) {
      return;
    }
    std::vector<std::uint32_t> order;
    const auto earlier = [this](std::uint32_t i, std::uint32_t j) { return less(notes_[i], notes_[j]); };
    sorted_index_order(order, static_cast<std::uint32_t>(notes_.size()), make_index_less(earlier));
    const Note* previous = nullptr;
    for (const std::uint32_t index : order) {
      const Note& n = notes_[index];
      if (previous != nullptr && !less(*previous, n)) {
        continue;
      }
      previous = &n;
      std::string_view name;
      if (ref_sheet_name(ctx_, n.sheet, &name)) {
        ctx_.note_dynamic_read(name, n.rect);
      }
    }
    notes_.clear();
  }

 private:
  struct Note {
    std::uint16_t sheet;
    DeclaredRect rect;
  };

  static bool less(const Note& a, const Note& b) noexcept {
    if (a.sheet != b.sheet) {
      return a.sheet < b.sheet;
    }
    if (a.rect.row_first != b.rect.row_first) {
      return a.rect.row_first < b.rect.row_first;
    }
    if (a.rect.row_last != b.rect.row_last) {
      return a.rect.row_last < b.rect.row_last;
    }
    if (a.rect.col_first != b.rect.col_first) {
      return a.rect.col_first < b.rect.col_first;
    }
    if (a.rect.col_last != b.rect.col_last) {
      return a.rect.col_last < b.rect.col_last;
    }
    return a.rect.whole_axis < b.rect.whole_axis;
  }

  void note(std::uint16_t sheet, const DeclaredRect& rect) {
    if (dynamic_) {
      notes_.push_back(Note{sheet, rect});
    }
  }

  bool sheet_cells(const RefArea& area, std::uint16_t sheet, std::uint16_t book, std::vector<Value>* out,
                   std::uint32_t* rows, std::uint32_t* cols, ErrorCode* err) {
    RefArea a = area;
    a.sheet_first = sheet;
    a.sheet_last = sheet;
    if (book != 0U) {
      return external_cells(a, book, out, rows, cols, err);
    }
    std::string_view name;
    if (!ref_sheet_name(ctx_, sheet, &name)) {
      *err = ErrorCode::Ref;
      return false;
    }
    parser::Reference lo{};
    parser::Reference hi{};
    corner_refs(a, name, &lo, &hi);
    // An edge-reaching rectangle narrows to its populated head, as a whole axis does.
    if (TailAxis axis = TailAxis::kRows; reaches_grid_edge(a, &axis)) {
      const auto head = edge_head(a, name, axis);
      if (!head) {
        *err = head.error();
        return false;
      }
      if (!head.value().has_value()) {
        note(sheet, rect_of(a));
        if (rows != nullptr) {
          *rows = axis == TailAxis::kRows ? 0U : a.row_last - a.row_first + 1U;
        }
        if (cols != nullptr) {
          *cols = axis == TailAxis::kRows ? static_cast<std::uint32_t>(a.col_last) - a.col_first + 1U : 0U;
        }
        return true;
      }
      const DeclaredRect& h = *head.value();
      set_rect_corner_refs(name, h.row_first, h.col_first, h.row_last, h.col_last, &lo, &hi);
    }
    auto expanded = ctx_.expand_range(lo, hi, arena_, registry_, rows, cols);
    if (!expanded) {
      *err = expanded.error();
      return false;
    }
    note(sheet, rect_of(a));
    out->insert(out->end(), expanded.value().begin(), expanded.value().end());
    return true;
  }

  bool external_cells(const RefArea& a, std::uint16_t book, std::vector<Value>* out, std::uint32_t* rows,
                      std::uint32_t* cols, ErrorCode* err) {
    const ExternalBook* ext = external_book_of(ctx_, book);
    if (ext == nullptr || !ext->sheet_readable(a.sheet_first)) {
      *err = ErrorCode::Ref;
      return false;
    }
    const DeclaredRect walked = external_walked(*ext, a.sheet_first, a);
    if (static_cast<std::uint64_t>(walked.rows()) * walked.cols() > kMaxRangeExpansionCells) {
      *err = ErrorCode::Calc;
      return false;
    }
    for (std::uint32_t r = walked.row_first; r <= walked.row_last; ++r) {
      for (std::uint32_t c = walked.col_first; c <= walked.col_last; ++c) {
        out->push_back(external_cell(*ext, a.sheet_first, r, c, arena_));
      }
    }
    if (rows != nullptr) {
      *rows = walked.rows();
    }
    if (cols != nullptr) {
      *cols = walked.cols();
    }
    return true;
  }

  Shaped external_shaped(const RefArea& a, std::uint16_t book) {
    const ExternalBook* ext = external_book_of(ctx_, book);
    if (ext == nullptr || !ext->sheet_readable(a.sheet_first)) {
      return Shaped{Value::error(ErrorCode::Ref), nullptr};
    }
    if (is_one_cell(a)) {
      return Shaped{adopt_text_into(arena_, ext->cached_cell(a.sheet_first, a.row_first, a.col_first)), nullptr};
    }
    std::vector<Value> cells;
    std::uint32_t rows = 0;
    std::uint32_t cols = 0;
    ErrorCode err = ErrorCode::Ref;
    if (!external_cells(a, book, &cells, &rows, &cols, &err)) {
      return Shaped{Value::error(err), nullptr};
    }
    if (a.whole_col == 0U && a.whole_row == 0U) {
      ArrayValue* arr = array_from_values(rows, cols, cells.data(), cells.size(), arena_);
      return Shaped{arr == nullptr ? Value::error(ErrorCode::Num) : Value::array(arr), nullptr};
    }
    Value* head = arena_.create_array<Value>(cells.size());
    if (head == nullptr) {
      return Shaped{Value::error(ErrorCode::Num), nullptr};
    }
    std::copy(cells.begin(), cells.end(), head);
    const DeclaredRect declared = rect_of(a);
    const bool by_rows = declared.rows() == Sheet::kMaxRows;
    return make_reference_tail(arena_, declared.rows(), declared.cols(), by_rows ? rows : cols,
                               by_rows ? TailAxis::kRows : TailAxis::kCols, head);
  }

  Arena& arena_;
  const FunctionRegistry& registry_;
  const EvalContext ctx_;
  const bool dynamic_;
  std::vector<Note> notes_;
};

}  // namespace

Expected<RefCells, ErrorCode> deref_cells(const Value& ref, bool allow_span3d, Arena& arena,
                                          const FunctionRegistry& registry, const EvalContext& ctx, ReadCount count) {
  if (!ref.is_ref()) {
    return ErrorCode::Value;
  }
  const RefView view = ref.as_ref();
  if (view.form == RefForm::kGrid) {
    return ErrorCode::Value;
  }
  if (view.form == RefForm::kSpan3D && !allow_span3d) {
    return ErrorCode::Ref;
  }
  Reader reader(arena, registry, ctx, count, ref.ref_is_dynamic());
  const bool one_rectangle = view.form == RefForm::kArea;
  RefCells out;
  ErrorCode err = ErrorCode::Ref;
  bool ok = true;
  for (std::uint32_t i = 0; ok && i < view.area_count; ++i) {
    ok = reader.read_cells(view.areas[i], view.book, &out.cells, one_rectangle ? &out.rows : nullptr,
                           one_rectangle ? &out.cols : nullptr, &err);
  }
  reader.flush();
  if (!ok) {
    return err;
  }
  return out;
}

Shaped deref_shaped(const Value& ref, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                    ReadCount count) {
  if (!ref.is_ref()) {
    return Shaped{Value::error(ErrorCode::Value), nullptr};
  }
  const RefView view = ref.as_ref();
  switch (view.form) {
    case RefForm::kGrid:
      return Shaped{materialise_grid(ref, GridArg::kElementError, nullptr, arena, registry, ctx, count), nullptr};
    case RefForm::kUnion:
      return Shaped{Value::error(ErrorCode::Value), nullptr};
    case RefForm::kSpan3D:
      return Shaped{Value::error(ErrorCode::Ref), nullptr};
    case RefForm::kArea:
      break;
  }
  Reader reader(arena, registry, ctx, count, ref.ref_is_dynamic());
  const Shaped read = reader.read_shaped(view.areas[0], view.book);
  reader.flush();
  return read;
}

Value deref_top_left(const Value& ref, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx,
                     ReadCount count) {
  if (!ref.is_ref()) {
    return Value::error(ErrorCode::Value);
  }
  const RefView view = ref.as_ref();
  if (view.area_count == 0U) {
    return Value::error(ErrorCode::Value);
  }
  Reader reader(arena, registry, ctx, count, ref.ref_is_dynamic());
  const Value v = view.form == RefForm::kGrid ? reader.read_element(view, 0)
                                              : reader.read_top_left(view.areas[0], element_book(view, 0));
  reader.flush();
  return v;
}

Value materialise_grid(const Value& ref, GridArg cls, const FunctionDef* def, Arena& arena,
                       const FunctionRegistry& registry, const EvalContext& ctx, ReadCount count) {
  switch (cls) {
    case GridArg::kPerRef:
    case GridArg::kKeepRef:
    case GridArg::kReshape:
    case GridArg::kLambdaElement:
      return ref;
    case GridArg::kWholeValue:
    case GridArg::kWholeNA:
    case GridArg::kElementError:
    case GridArg::kPerValue:
    case GridArg::kTopLeft:
    case GridArg::kOpaque:
      break;
  }
  if (!ref.is_ref() || ref.ref_form() != RefForm::kGrid) {
    return ref;
  }
  if (cls == GridArg::kWholeValue) {
    return Value::error(ErrorCode::Value);
  }
  if (cls == GridArg::kWholeNA) {
    return Value::error(ErrorCode::NA);
  }
  const RefView view = ref.as_ref();
  const std::size_t n = static_cast<std::size_t>(view.rows) * view.cols;
  if (n == 0U) {
    return Value::error(ErrorCode::Value);
  }
  Reader reader(arena, registry, ctx, count, ref.ref_is_dynamic());

  if (cls == GridArg::kTopLeft) {
    std::vector<Value> kept;
    for (std::size_t i = 0; i < n; ++i) {
      Value v = reader.read_element(view, i);
      if (v.is_blank()) {
        v = Value::blank(BlankGridProjection::kReferenceGridZero);
      }
      Value range_err = Value::blank();
      if (def == nullptr) {
        kept.push_back(v);
      } else if (!append_range_sourced_value(*def, v, &kept, &range_err)) {
        reader.flush();
        return range_err;
      }
    }
    reader.flush();
    if (kept.empty()) {
      kept.push_back(Value::blank());
    }
    Value* cells = nullptr;
    ArrayValue* arr = allocate_array_value(static_cast<std::uint32_t>(kept.size()), 1U, arena, cells);
    if (arr == nullptr) {
      return Value::error(ErrorCode::Num);
    }
    std::copy(kept.begin(), kept.end(), cells);
    return Value::array(arr);
  }

  Value* cells = nullptr;
  ArrayValue* arr = allocate_array_value(view.rows, view.cols, arena, cells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    const RefArea& element = view.areas[i];
    if (grid_element_value(view, i, &cells[i])) {
      continue;
    }
    switch (cls) {
      case GridArg::kOpaque:
        cells[i] = Value::blank(BlankGridProjection::kOpaqueElement);
        break;
      case GridArg::kPerValue:
        cells[i] = is_one_cell(element) ? reader.read_top_left(element, element_book(view, i))
                                        : Value::error(ErrorCode::Value);
        break;
      default:
        cells[i] = Value::error(ErrorCode::Value);
        break;
    }
  }
  reader.flush();
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
