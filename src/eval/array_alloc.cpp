
#include "eval/array_alloc.h"

#include <cstddef>
#include <cstdint>

#include "eval/tail_array.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/checked_mul.h"
#include "value.h"

namespace formulon {
namespace eval {

ArrayValue* allocate_array_value(std::uint32_t rows, std::uint32_t cols, Arena& arena, Value*& out_buffer,
                                 std::uint64_t max_cells) {
  out_buffer = nullptr;
  if (rows == 0U || cols == 0U || rows > Sheet::kMaxRows || cols > Sheet::kMaxCols) {
    return nullptr;
  }
  const auto total = checked_mul_size_t(rows, cols);
  if (!total || static_cast<std::uint64_t>(total.value()) > max_cells) {
    return nullptr;
  }
  Value* buffer = arena.create_array<Value>(total.value());
  if (buffer == nullptr) {
    return nullptr;
  }
  ArrayValue* out = arena.create<ArrayValue>();
  if (out == nullptr) {
    return nullptr;
  }
  out->rows = rows;
  out->cols = cols;
  out->cells = buffer;
  out_buffer = buffer;
  return out;
}

ArrayValue* array_from_values(std::uint32_t rows, std::uint32_t cols, const Value* values, std::size_t count,
                              Arena& arena) {
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return nullptr;
  }
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = i < count ? values[i] : Value::blank();
  }
  return arr;
}

const TailArray* make_tail_array(Arena& arena, std::uint32_t rows, std::uint32_t cols, std::uint32_t head,
                                 TailAxis axis, const Value* cells, const Value* tail, bool from_reference) {
  return arena.create<TailArray>(TailArray{rows, cols, head, axis, cells, tail, from_reference});
}

const Value& tail_array_at(const TailArray& ta, std::uint32_t r, std::uint32_t c) {
  if (ta.axis == TailAxis::kRows) {
    return r < ta.head ? ta.cells[static_cast<std::size_t>(r) * ta.cols + c] : ta.tail[c];
  }
  return c < ta.head ? ta.cells[static_cast<std::size_t>(r) * ta.head + c] : ta.tail[r];
}

Value densify(const Shaped& s, Arena& arena) {
  if (s.tail_array == nullptr) {
    return s.value;
  }
  const TailArray& ta = *s.tail_array;
  if (ta.from_reference && static_cast<std::uint64_t>(ta.rows) * ta.cols > kMaxRangeExpansionCells) {
    return Value::error(ErrorCode::Calc);
  }
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(ta.rows, ta.cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::uint32_t r = 0; r < ta.rows; ++r) {
    for (std::uint32_t c = 0; c < ta.cols; ++c) {
      buffer[static_cast<std::size_t>(r) * ta.cols + c] = tail_array_at(ta, r, c);
    }
  }
  return Value::array(arr);
}

}  // namespace eval
}  // namespace formulon
