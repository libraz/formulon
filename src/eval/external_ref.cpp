
#include "eval/external_ref.h"

#include <cstddef>
#include <cstdint>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/eval_context.h"
#include "io/external_book.h"
#include "io/external_links.h"
#include "parser/ast.h"
#include "utils/arena.h"
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
Value MaterializeRect(const io::ExternalBook& book, std::uint32_t sheet, std::uint32_t row_first,
                      std::uint32_t row_last, std::uint32_t col_first, std::uint32_t col_last, Arena& arena) {
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

}  // namespace

Value resolve_external_ref(const parser::AstNode& node, Arena& arena, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  if (wb == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  // `[N]` is 1-based and selects the N-th `<externalReference>`, which is
  // the N-th entry of this list because the reader builds it in document
  // order.
  const std::vector<io::ExternalLinkRecord>& links = wb->external_links();
  const std::uint32_t book_index = node.as_external_ref_book();
  if (book_index == 0 || book_index > links.size()) {
    return Value::error(ErrorCode::Ref);
  }
  const io::ExternalBook& book = links[book_index - 1U].book;

  if (const std::string_view name = node.as_external_ref_name(); !name.empty()) {
    const io::ExternalBookName* entry = book.find_name(name);
    if (entry == nullptr) {
      return Value::error(ErrorCode::Name);
    }
    if (!entry->resolvable) {
      return Value::error(ErrorCode::Ref);
    }
    if (!entry->is_range) {
      return ReifyCached(book.cached_cell(entry->sheet, entry->row, entry->col), arena);
    }
    return MaterializeRect(book, entry->sheet, entry->row, entry->row_end, entry->col, entry->col_end, arena);
  }

  const std::uint32_t sheet = book.sheet_index(node.as_external_ref_sheet());
  if (sheet == io::ExternalBook::kNoSheet) {
    return Value::error(ErrorCode::Ref);
  }
  const parser::Reference& first = node.as_external_ref_cell();
  if (!node.as_external_ref_is_range()) {
    return ReifyCached(book.cached_cell(sheet, first.row, first.col), arena);
  }
  // Endpoint order is normalised the same way a local range is, so
  // `[1]Data!A3:A1` and `[1]Data!A1:A3` denote one rectangle.
  const parser::Reference& last = node.as_external_ref_cell_end();
  const std::uint32_t row_first = first.row < last.row ? first.row : last.row;
  const std::uint32_t row_last = first.row < last.row ? last.row : first.row;
  const std::uint32_t col_first = first.col < last.col ? first.col : last.col;
  const std::uint32_t col_last = first.col < last.col ? last.col : first.col;
  return MaterializeRect(book, sheet, row_first, row_last, col_first, col_last, arena);
}

}  // namespace eval
}  // namespace formulon
