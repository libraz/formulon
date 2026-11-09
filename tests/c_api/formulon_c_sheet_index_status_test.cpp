// Sheet-index validation across the C ABI. Every sheet-scoped function, called
// through the generated table in `sheet_index_calls.inc`, checks its arguments
// in one order -- handle, NULL pointers, sheet index, everything else -- and
// reports an out-of-range sheet as FM_STATUS_SHEET_INDEX_OUT_OF_RANGE. A
// successful call leaves the last-error state empty. Every `fm_cell_rect_t`
// argument rejects an inverted or off-grid rectangle as invalid. A record
// argument's fields are ordinary arguments, checked after the sheet.

#include <cstddef>
#include <cstdint>
#include <cstring>
#include <limits>
#include <type_traits>
#include <vector>

#include "formulon_c_record_arg.h"
#include "formulon_c_setters.h"
#include "formulon_c_test_helpers.h"
#include "gtest/gtest.h"

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

struct SheetIndexCall {
  const char* name;
  std::size_t pointer_count;
  std::uint32_t text_mask;  // pointer slots that are `const char*` arguments
  std::uint32_t rect_mask;  // pointer slots that are single `const fm_cell_rect_t*` arguments
  fm_status_t (*call)(fm_workbook_t* wb, std::size_t sheet_index, ZeroSlots& slots);
  void (*release)(ZeroSlots& slots);
  void (*poison)(ZeroSlots& slots);  // makes every field of each input record invalid
};

#include "sheet_index_calls.inc"

// One fresh single-sheet workbook per call keeps a mutation from one entry
// point out of the next one's arguments.
fm_status_t Invoke(const SheetIndexCall& entry, std::size_t sheet_index, std::uint32_t null_mask,
                   bool null_handle = false) {
  WorkbookGuard wb;
  EXPECT_EQ(fm_workbook_create_default(&wb.handle), 0);
  ZeroSlots slots;
  slots.null_mask = null_mask;
  const fm_status_t rc = entry.call(null_handle ? nullptr : wb.handle, sheet_index, slots);
  if (entry.release != nullptr) {
    entry.release(slots);
  }
  return rc;
}

std::vector<std::size_t> OutOfRangeSheets() {
  std::vector<std::size_t> sheets = {1U, std::size_t{1} << 31U};
  if (sizeof(std::size_t) > sizeof(std::uint32_t)) {
    // Truncated to 32 bits this would name sheet 0.
    sheets.push_back(static_cast<std::size_t>(std::uint64_t{1} << 32U));
  }
  sheets.push_back(std::numeric_limits<std::size_t>::max());
  return sheets;
}

}  // namespace

TEST(FormulonCApiSheetIndexStatus, TableCoversTheSheetScope) {
  EXPECT_GT(sizeof(kSheetIndexCalls) / sizeof(kSheetIndexCalls[0]), 200U);
}

TEST(FormulonCApiSheetIndexStatus, OutOfRangeSheetIsSheetIndexOutOfRange) {
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    for (const std::size_t sheet : OutOfRangeSheets()) {
      EXPECT_EQ(Invoke(entry, sheet, 0U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE)
          << entry.name << " sheet_index=" << sheet << ": " << fm_get_last_error_message();
      EXPECT_STRNE(fm_get_last_error_message(), "") << entry.name;
    }
  }
}

// Text validity is an ordinary argument check, so it runs after the sheet.
TEST(FormulonCApiSheetIndexStatus, MalformedUtf8IsCheckedAfterTheSheet) {
  std::size_t text_calls = 0;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    if (entry.text_mask == 0U) {
      continue;
    }
    ++text_calls;
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ZeroSlots slots;
    for (std::size_t slot = 0; slot < ZeroSlots::kCount; ++slot) {
      if ((entry.text_mask & (1U << slot)) != 0U) {
        slots.storage[slot][0] = 0xFFU;
      }
    }
    const fm_status_t rc = entry.call(wb.handle, 1U, slots);
    if (entry.release != nullptr) {
      entry.release(slots);
    }
    EXPECT_EQ(rc, FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << entry.name << ": " << fm_get_last_error_message();
  }
  EXPECT_GT(text_calls, 0U);
}

// Malformed text, integers at their type's maximum and NaN doubles in a record
// argument still lose to an out-of-range sheet.
TEST(FormulonCApiSheetIndexStatus, MalformedRecordIsCheckedAfterTheSheet) {
  std::size_t record_calls = 0;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    if (entry.poison == nullptr) {
      continue;
    }
    ++record_calls;
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ZeroSlots slots;
    entry.poison(slots);
    const fm_status_t rc = entry.call(wb.handle, 1U, slots);
    if (entry.release != nullptr) {
      entry.release(slots);
    }
    EXPECT_EQ(rc, FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << entry.name << ": " << fm_get_last_error_message();
  }
  EXPECT_GT(record_calls, 0U);
}

TEST(FormulonCApiSheetIndexStatus, NullHandleIsCheckedBeforeTheSheet) {
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    EXPECT_EQ(Invoke(entry, 1U, 0U, /*null_handle=*/true), FM_STATUS_BINDING_NULL_POINTER) << entry.name;
  }
}

// A pointer the function requires is one it rejects as NULL on a valid sheet;
// with the sheet out of range too, the NULL still wins.
TEST(FormulonCApiSheetIndexStatus, RequiredPointerIsCheckedBeforeTheSheet) {
  std::size_t required = 0;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    for (std::size_t slot = 0; slot < entry.pointer_count; ++slot) {
      const std::uint32_t mask = 1U << slot;
      if (Invoke(entry, 0U, mask) != FM_STATUS_BINDING_NULL_POINTER) {
        continue;
      }
      ++required;
      EXPECT_EQ(Invoke(entry, 1U, mask), FM_STATUS_BINDING_NULL_POINTER) << entry.name << " pointer #" << slot;
    }
  }
  EXPECT_GT(required, 0U);
}

TEST(FormulonCApiSheetIndexStatus, SuccessClearsTheLastError) {
  std::size_t succeeded = 0;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    WorkbookGuard wb;
    ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
    ZeroSlots slots;
    ASSERT_EQ(entry.call(wb.handle, 1U, slots), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << entry.name;
    ASSERT_STRNE(fm_get_last_error_message(), "");

    // One unrelated successful call clears the rejection.
    std::size_t count = 0;
    ASSERT_EQ(fm_workbook_get_sheet_count(wb.handle, &count), 0);
    EXPECT_STREQ(fm_get_last_error_message(), "") << "after " << entry.name;
    EXPECT_STREQ(fm_get_last_error_context(), "") << "after " << entry.name;

    // So does the function itself, whenever zeroed arguments are accepted.
    ASSERT_EQ(entry.call(wb.handle, 1U, slots), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE) << entry.name;
    ZeroSlots fresh;
    const fm_status_t rc = entry.call(wb.handle, 0U, fresh);
    if (entry.release != nullptr) {
      entry.release(fresh);
    }
    if (rc == FM_STATUS_OK) {
      ++succeeded;
      EXPECT_STREQ(fm_get_last_error_message(), "") << entry.name;
      EXPECT_STREQ(fm_get_last_error_context(), "") << entry.name;
    }
  }
  EXPECT_GT(succeeded, 0U);
}

// The workbook-scope functions that address a sheet by its index report a bad
// one with the same code.
TEST(FormulonCApiSheetIndexStatus, SheetCollectionIndexIsSheetIndexOutOfRange) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  constexpr std::size_t kBad = 1U;
  const char* name = nullptr;
  std::size_t index = 0;
  const char* columns[] = {"A"};

  EXPECT_EQ(fm_workbook_get_sheet_name(wb.handle, kBad, &name), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_set_active_sheet(wb.handle, kBad), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_remove_sheet(wb.handle, kBad), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_rename_sheet(wb.handle, kBad, "Other"), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_move_sheet(wb.handle, kBad, 0U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_move_sheet(wb.handle, 0U, kBad + 1U), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_copy_sheet(wb.handle, kBad, 0U, nullptr, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(fm_workbook_copy_sheet(wb.handle, 0U, kBad + 1U, nullptr, &index), FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(AddTable(wb.handle, kBad, record_arg(fm_cell_rect_t{0, 0, 1, 0}), "T", columns, 1U, nullptr, 1, 0, &index),
            FM_STATUS_SHEET_INDEX_OUT_OF_RANGE);
  EXPECT_EQ(count_of(fm_workbook_get_sheet_count, wb.handle), 1U);
}

// Rectangles every `fm_cell_rect_t` argument refuses: corners in reverse order
// on either axis, and a corner one past the last row or column.
const fm_cell_rect_t kBadRects[] = {
    {2U, 0U, 1U, 0U},
    {0U, 2U, 0U, 1U},
    {0U, 0U, 1048576U, 0U},
    {0U, 0U, 0U, 16384U},
};

TEST(FormulonCApiCellRect, EveryRectArgumentRejectsAnInvertedOrOffGridRect) {
  std::size_t rect_calls = 0;
  for (const SheetIndexCall& entry : kSheetIndexCalls) {
    if (entry.rect_mask == 0U) {
      continue;
    }
    ++rect_calls;
    for (const fm_cell_rect_t& bad : kBadRects) {
      WorkbookGuard wb;
      ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
      ZeroSlots slots;
      for (std::size_t slot = 0; slot < ZeroSlots::kCount; ++slot) {
        if ((entry.rect_mask & (1U << slot)) != 0U) {
          std::memcpy(slots.storage[slot], &bad, sizeof(bad));
        }
      }
      const fm_status_t rc = entry.call(wb.handle, 0U, slots);
      if (entry.release != nullptr) {
        entry.release(slots);
      }
      EXPECT_EQ(rc, FM_STATUS_INVALID_ARGUMENT) << entry.name << " {" << bad.first_row << "," << bad.first_col << ","
                                                << bad.last_row << "," << bad.last_col << "}";
    }
  }
  EXPECT_GE(rect_calls, 13U);
}

// The rectangle arguments outside the sheet-scoped table: a selection range and
// the workbook-scope reference formatter.
TEST(FormulonCApiCellRect, SelectionAndReferenceRejectAnInvertedOrOffGridRect) {
  WorkbookGuard wb;
  ASSERT_EQ(fm_workbook_create_default(&wb.handle), 0);
  for (const fm_cell_rect_t& bad : kBadRects) {
    EXPECT_EQ(fm_sheet_set_selection(wb.handle, 0U, 0U, 0U, &bad, 1U), FM_STATUS_INVALID_ARGUMENT);
    const char* text = nullptr;
    EXPECT_EQ(fm_workbook_format_reference(wb.handle, nullptr, &bad, 0, &text), FM_STATUS_INVALID_ARGUMENT);
  }
}
