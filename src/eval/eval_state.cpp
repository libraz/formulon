// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of `EvalState`. See `eval_state.h` for the public contract.

#include "eval/eval_state.h"

#include <cassert>
#include <cstdint>

#include "sheet.h"

namespace formulon {
namespace eval {

bool EvalState::push_cell(const Sheet* sheet, std::uint32_t row, std::uint32_t col) {
  const CellKey key{sheet, row, col};
  for (const CellKey& entry : stack_) {
    if (entry == key) {
      return false;
    }
  }
  stack_.push_back(key);
  return true;
}

void EvalState::pop_cell(const Sheet* sheet, std::uint32_t row, std::uint32_t col) {
  // The braced initializer is wrapped in an extra pair of parens so the
  // commas inside don't look like argument separators to the `assert`
  // function-like macro.
  const CellKey expected{sheet, row, col};
  assert(!stack_.empty() && stack_.back() == expected);
  (void)expected;
  stack_.pop_back();
}

const Value* EvalState::lookup_memo(const Sheet* sheet, std::uint32_t row,
                                    std::uint32_t col) const noexcept {
  const auto found = memo_.find(CellKey{sheet, row, col});
  if (found == memo_.end()) {
    return nullptr;
  }
  return &found->second;
}

void EvalState::memoize(const Sheet* sheet, std::uint32_t row, std::uint32_t col,
                        Value value) {
  // `unordered_map::operator[]` requires a default-constructible mapped
  // type; `Value`'s default ctor is private. `insert_or_assign` constructs
  // the pair from the supplied Value directly. `Value` is trivially
  // copyable, so no move is needed.
  memo_.insert_or_assign(CellKey{sheet, row, col}, value);
}

}  // namespace eval
}  // namespace formulon
