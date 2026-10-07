
#include "eval/external_ref.h"

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/declared_rect.h"
#include "eval/eval_context.h"
#include "eval/lazy_impls.h"
#include "external_book.h"
#include "external_link.h"
#include "parser/ast.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/expected.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

/// Re-interns a Text result into the evaluation arena so the returned
/// value does not borrow the workbook's cache.
Value ReifyCached(Value cached, Arena& arena) {
  if (!cached.is_text()) {
    return cached;
  }
  const std::string_view interned = arena.intern(cached.as_text());
  return Value::text(interned);
}

/// Materialises `[row_first..row_last] x [col_first..col_last]` of
/// `book`'s sheet `sheet` as an Array.
Value MaterializeRect(const ExternalBook& book, std::uint32_t sheet, std::uint32_t row_first, std::uint32_t row_last,
                      std::uint32_t col_first, std::uint32_t col_last, Arena& arena) {
  const std::uint32_t rows = row_last - row_first + 1U;
  const std::uint32_t cols = col_last - col_first + 1U;
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::size_t k = 0;
  for (std::uint32_t r = row_first; r <= row_last; ++r) {
    for (std::uint32_t c = col_first; c <= col_last; ++c) {
      buffer[k] = ReifyCached(book.cached_cell(sheet, r, c), arena);
      ++k;
    }
  }
  return Value::array(arr);
}

/// The rectangle `first`..`last` spans on `sheet`, a whole column or row
/// clipped to the sheet's cached extent. False for a whole column or row
/// over a sheet that caches nothing.
bool TailRect(const ExternalBook& book, std::uint32_t sheet, const parser::Reference& first,
              const parser::Reference& last, std::uint32_t* row_first, std::uint32_t* row_last,
              std::uint32_t* col_first, std::uint32_t* col_last) {
  // Endpoint order is normalised the same way a local range is, so
  // `Data!A3:A1` and `Data!A1:A3` denote one rectangle.
  *row_first = first.row < last.row ? first.row : last.row;
  *row_last = first.row < last.row ? last.row : first.row;
  *col_first = first.col < last.col ? first.col : last.col;
  *col_last = first.col < last.col ? last.col : first.col;
  if (!first.is_full_col && !first.is_full_row) {
    return true;
  }
  std::uint32_t extent_row = 0;
  std::uint32_t extent_col = 0;
  if (!book.cached_extent(sheet, &extent_row, &extent_col)) {
    return false;
  }
  if (first.is_full_col) {
    *row_first = 0;
    *row_last = extent_row;
  } else {
    *col_first = 0;
    *col_last = extent_col;
  }
  return true;
}

/// Whether every sheet of the 3-D span `node` names is listed and cached,
/// with the span's first and last sheet in `*lo` / `*hi`.
bool SpanReadable(const ExternalBook& book, const parser::AstNode& node, std::uint32_t* lo, std::uint32_t* hi) {
  const std::uint32_t begin = book.sheet_index(node.as_external_ref_sheet());
  const std::uint32_t end = book.sheet_index(node.as_external_ref_sheet_end());
  if (begin == ExternalBook::kNoSheet || end == ExternalBook::kNoSheet) {
    return false;
  }
  const std::uint32_t first = begin < end ? begin : end;
  const std::uint32_t last = begin < end ? end : begin;
  for (std::uint32_t sheet = first; sheet <= last; ++sheet) {
    if (!book.sheet_has_data(sheet)) {
      return false;
    }
  }
  *lo = first;
  *hi = last;
  return true;
}

const ExternalBook* BookFor(const parser::AstNode& node, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return nullptr;
  }
  const ExternalLinkRecord* link = wb->find_external_link(node.as_external_ref_path(), node.as_external_ref_book());
  return link == nullptr ? nullptr : &link->book;
}

}  // namespace

Value resolve_external_ref(const parser::AstNode& node, Arena& arena, const EvalContext& ctx) {
  const ExternalBook* found = BookFor(node, ctx);
  if (found == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  const ExternalBook& book = *found;
  const std::string_view sheet_name = node.as_external_ref_sheet();

  if (const std::string_view name = node.as_external_ref_name(); !name.empty()) {
    // `[Book]Sheet!Name` is the name local to that sheet; `Book!Name` the
    // book-scope one.
    std::uint32_t scope = ExternalBook::kNoSheet;
    if (!sheet_name.empty()) {
      scope = book.sheet_index(sheet_name);
      if (scope == ExternalBook::kNoSheet) {
        return Value::error(ErrorCode::Name);
      }
    }
    return resolve_external_book_name(book, scope, name, arena);
  }

  if (!node.as_external_ref_sheet_end().empty()) {
    // Scalar context, as a local 3-D reference: a span of sheets is #REF!.
    return Value::error(ErrorCode::Ref);
  }
  const std::uint32_t sheet = book.sheet_index(sheet_name);
  if (sheet == ExternalBook::kNoSheet || !book.sheet_has_data(sheet)) {
    return Value::error(ErrorCode::Ref);
  }
  const parser::Reference& first = node.as_external_ref_cell();
  if (!node.as_external_ref_is_range() && !first.is_full_col && !first.is_full_row) {
    return ReifyCached(book.cached_cell(sheet, first.row, first.col), arena);
  }
  std::uint32_t row_first = 0;
  std::uint32_t row_last = 0;
  std::uint32_t col_first = 0;
  std::uint32_t col_last = 0;
  if (!TailRect(book, sheet, first, node.as_external_ref_cell_end(), &row_first, &row_last, &col_first, &col_last)) {
    // Nothing cached: the clipped rectangle is the corner cell alone.
    row_last = row_first;
    col_last = col_first;
  }
  return MaterializeRect(book, sheet, row_first, row_last, col_first, col_last, arena);
}

Shaped resolve_external_ref_shaped(const parser::AstNode& node, Arena& arena, const EvalContext& ctx) {
  Shaped out;
  out.value = resolve_external_ref(node, arena, ctx);
  return out;
}

Expected<ExternalRect, ErrorCode> resolve_external_rect(const parser::AstNode& node, const EvalContext& ctx) {
  const parser::Reference* lhs = nullptr;
  const parser::Reference* rhs = nullptr;
  if (!external_ref_declared_endpoints(node, &lhs, &rhs)) {
    return ErrorCode::Value;
  }
  const ExternalBook* book = BookFor(node, ctx);
  if (book == nullptr) {
    return ErrorCode::Ref;
  }
  const std::uint32_t sheet = book->sheet_index(node.as_external_ref_sheet());
  if (sheet == ExternalBook::kNoSheet || !book->sheet_has_data(sheet)) {
    return ErrorCode::Ref;
  }
  const Expected<DeclaredRect, ErrorCode> declared = declared_rect(*lhs, *rhs);
  if (!declared) {
    return declared.error();
  }
  ExternalRect out;
  out.book = book;
  out.sheet = sheet;
  out.declared = declared.value();
  out.walked = declared.value();
  if (!TailRect(*book, sheet, *lhs, *rhs, &out.walked.row_first, &out.walked.row_last, &out.walked.col_first,
                &out.walked.col_last)) {
    // Nothing cached: a full read walks the first line, as it does over an
    // unpopulated local whole axis.
    out.walked = declared.value();
    if (out.walked.rows() == Sheet::kMaxRows) {
      out.walked.row_last = out.walked.row_first;
    } else {
      out.walked.col_last = out.walked.col_first;
    }
  }
  return out;
}

bool external_whole_axis_footprint(const parser::AstNode& node, std::uint32_t* out_rows, std::uint32_t* out_cols) {
  const parser::Reference* lhs = nullptr;
  const parser::Reference* rhs = nullptr;
  if (!external_ref_declared_endpoints(node, &lhs, &rhs)) {
    return false;
  }
  const Expected<DeclaredRect, ErrorCode> rect = declared_rect(*lhs, *rhs);
  if (!rect || !rect.value().whole_axis) {
    return false;
  }
  *out_rows = rect.value().rows();
  *out_cols = rect.value().cols();
  return true;
}

Value read_external_cell(const ExternalBook& book, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                         Arena& arena) {
  return ReifyCached(book.cached_cell(sheet, row, col), arena);
}

Value resolve_external_book_name(const ExternalBook& book, std::uint32_t scope_sheet, std::string_view name,
                                 Arena& arena) {
  const ExternalBookName* entry = book.find_name(name, scope_sheet);
  if (entry == nullptr || !entry->exists) {
    return Value::error(ErrorCode::Name);
  }
  if (!entry->resolvable || !book.sheet_has_data(entry->sheet)) {
    return Value::error(ErrorCode::Ref);
  }
  if (!entry->is_range) {
    return ReifyCached(book.cached_cell(entry->sheet, entry->row, entry->col), arena);
  }
  return MaterializeRect(book, entry->sheet, entry->row, entry->row_end, entry->col, entry->col_end, arena);
}

bool is_three_d_reference(const parser::AstNode& node) noexcept {
  return node.kind() == parser::NodeKind::Ref3D ||
         (node.kind() == parser::NodeKind::ExternalRef && !node.as_external_ref_sheet_end().empty());
}

Value eval_selected_arm(const parser::AstNode& arm, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (is_three_d_reference(arm)) {
    return Value::error(ErrorCode::Value);
  }
  return eval_node(arm, arena, registry, ctx);
}

bool collect_external_ref3d_cells(const parser::AstNode& node, Arena& arena, const EvalContext& ctx,
                                  std::vector<Value>* out) {
  const ExternalBook* found = BookFor(node, ctx);
  if (found == nullptr) {
    return false;
  }
  const ExternalBook& book = *found;
  std::uint32_t lo = 0;
  std::uint32_t hi = 0;
  if (!SpanReadable(book, node, &lo, &hi)) {
    return false;
  }
  const parser::Reference& first = node.as_external_ref_cell();
  const parser::Reference& last = node.as_external_ref_cell_end();
  for (std::uint32_t sheet = lo; sheet <= hi; ++sheet) {
    std::uint32_t row_first = 0;
    std::uint32_t row_last = 0;
    std::uint32_t col_first = 0;
    std::uint32_t col_last = 0;
    if (!TailRect(book, sheet, first, node.as_external_ref_is_range() ? last : first, &row_first, &row_last, &col_first,
                  &col_last)) {
      continue;
    }
    for (std::uint32_t r = row_first; r <= row_last; ++r) {
      for (std::uint32_t c = col_first; c <= col_last; ++c) {
        out->push_back(ReifyCached(book.cached_cell(sheet, r, c), arena));
      }
    }
  }
  return true;
}

}  // namespace eval
}  // namespace formulon
