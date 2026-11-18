// The argument storage and table row of the generated call table of every
// sheet-scoped C ABI function (`sheet_index_calls.inc`, from
// tools/codegen/gen_sheet_index_calls.py). Each including test file gets its own
// copy in its anonymous namespace; the .inc follows the include.

#ifndef FORMULON_TESTS_C_API_SHEET_INDEX_TABLE_H_
#define FORMULON_TESTS_C_API_SHEET_INDEX_TABLE_H_

#include <cstddef>
#include <cstdint>
#include <type_traits>

#include "c_api/formulon_c.h"

namespace {

template <typename T, typename = void>
struct HasStructSize : std::false_type {};
template <typename T>
struct HasStructSize<T, std::void_t<decltype(T::struct_size)>> : std::true_type {};

// Zeroed storage for every pointer argument of one call. `arg` hands out slot
// `I` as a `T*`, or NULL when the test has masked that slot. A record slot
// carries its `struct_size`.
struct ZeroSlots {
  static constexpr std::size_t kCount = 8;
  static constexpr std::size_t kBytes = 4096;

  alignas(std::max_align_t) unsigned char storage[kCount][kBytes] = {};
  std::uint32_t null_mask = 0;

  template <typename T, std::size_t I>
  T* at() {
    static_assert(I < kCount, "a sheet-scoped function takes more pointers than ZeroSlots holds");
    static_assert(sizeof(T) <= kBytes, "a pointee does not fit one slot");
    T* slot = reinterpret_cast<T*>(storage[I]);
    if constexpr (HasStructSize<T>::value) {
      const_cast<std::remove_const_t<T>*>(slot)->struct_size = static_cast<std::uint32_t>(sizeof(T));
    }
    return slot;
  }

  template <typename T, std::size_t I>
  T* arg() {
    return (null_mask & (1U << I)) != 0 ? nullptr : at<T, I>();
  }
};

// The sheet kinds a call serves, as the ledger's `sheet_kinds` states them.
enum class Serves { kAll, kWorksheetOrChartsheet, kWorksheet };

struct SheetIndexCall {
  const char* name;
  std::size_t pointer_count;
  std::uint32_t text_mask;  // pointer slots that are `const char*` arguments
  std::uint32_t rect_mask;  // pointer slots that are single `const fm_cell_rect_t*` arguments
  fm_status_t (*call)(fm_workbook_t* wb, std::size_t sheet_index, ZeroSlots& slots);
  void (*release)(ZeroSlots& slots);
  void (*poison)(ZeroSlots& slots);  // makes every field of each input record invalid
  Serves serves;
};

}  // namespace

#endif  // FORMULON_TESTS_C_API_SHEET_INDEX_TABLE_H_
