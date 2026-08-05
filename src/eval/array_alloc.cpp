
#include "eval/array_alloc.h"

#include <cstddef>
#include <cstdint>

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

}  // namespace eval
}  // namespace formulon
