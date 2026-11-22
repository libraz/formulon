//
// C ABI - manual page breaks: set, clear and enumerate a sheet's row and
// column breaks.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "c_api/parts/struct_v1_sizes.h"
#include "sheet.h"
#include "utils/resource_budget.h"
#include "workbook.h"

using formulon::c_api::parts::check_column_index;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_row_index;
using formulon::c_api::parts::check_sheet;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::invalid_argument;
using formulon::c_api::parts::null_argument;
using formulon::c_api::parts::SheetUse;

namespace {

using formulon::ManualBreak;
using formulon::Sheet;
using formulon::SheetPrintSettings;

// Perpendicular span Excel writes for a break inserted across the whole
// sheet: a row break spans every column, a column break every row.
constexpr std::uint32_t kWholeSheetColSpanMax = Sheet::kMaxCols - 1U;
constexpr std::uint32_t kWholeSheetRowSpanMax = Sheet::kMaxRows - 1U;

/// Upserts a manual break into an axis vector, keeping it sorted by index.
fm_status_t upsert_break(std::vector<ManualBreak>& breaks, std::uint32_t index, std::uint32_t first, std::uint32_t last,
                         bool manual) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    pos->min = first;
    pos->max = last;
    pos->manual = manual;
    return 0;
  }
  if (breaks.size() >= formulon::kMaxManualBreaksPerAxis) {
    return invalid_argument(
        "The axis already holds the most manual page breaks",
        formulon::limit_context("manual_breaks", breaks.size() + 1U, formulon::kMaxManualBreaksPerAxis));
  }
  ManualBreak entry;
  entry.id = index;
  entry.min = first;
  entry.max = last;
  entry.manual = manual;
  breaks.insert(pos, entry);
  return 0;
}

/// Erases the break at `index`; a no-op when the axis has none there.
void erase_break(std::vector<ManualBreak>& breaks, std::uint32_t index) {
  const auto pos = std::lower_bound(breaks.begin(), breaks.end(), index,
                                    [](const ManualBreak& entry, std::uint32_t key) { return entry.id < key; });
  if (pos != breaks.end() && pos->id == index) {
    breaks.erase(pos);
  }
}

/// Shared body of `fm_sheet_set_row_break` / `fm_sheet_set_column_break`.
/// A null record removes the break; `first == last == 0` takes the whole span.
fm_status_t set_break(fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t index, bool rows,
                      const std::uint32_t* first_in, const std::uint32_t* last_in, const int32_t* manual_in) {
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  if (auto rc = rows ? check_row_index(index) : check_column_index(index); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  std::vector<ManualBreak>& breaks = rows ? print.manual_row_breaks : print.manual_col_breaks;
  if (first_in == nullptr) {
    erase_break(breaks, index);
    return 0;
  }
  const std::uint32_t span_max = rows ? kWholeSheetColSpanMax : kWholeSheetRowSpanMax;
  std::uint32_t first = *first_in;
  std::uint32_t last = *last_in;
  if (first == 0U && last == 0U) {
    last = span_max;
  }
  if (first > last || last > span_max) {
    return invalid_argument("first and last must be within the break span",
                            "first=" + std::to_string(first) + " last=" + std::to_string(last));
  }
  if (*manual_in < FM_TRISTATE_DEFAULT || *manual_in > FM_TRISTATE_TRUE) {
    return invalid_argument(("manual " + std::to_string(*manual_in) + " is not a tristate").c_str(),
                            "manual=" + std::to_string(*manual_in));
  }
  return upsert_break(breaks, index, first, last, *manual_in != FM_TRISTATE_FALSE);
}

// Counts one axis. The `size_t` return has no room for a status, so an
// argument error is reported as `0` plus a populated thread-local
// diagnostic; `clear_last_error` above is what makes the empty message a
// reliable "this zero means the axis is empty" signal. Callers that
// enumerate by count-then-`break_at` must consult the diagnostic, since
// the `break_at` loop that would otherwise surface the status never runs
// on a zero count.
fm_status_t break_count(const fm_workbook_t* wb, std::size_t sheet_index, bool rows, std::size_t* out_count) {
  clear_last_error();
  if (out_count == nullptr) {
    return null_argument("out_count");
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  *out_count = rows ? print.manual_row_breaks.size() : print.manual_col_breaks.size();
  return 0;
}

template <typename Record>
fm_status_t break_at(const fm_workbook_t* wb, std::size_t sheet_index, std::size_t index, bool rows, Record* out,
                     std::uint32_t Record::*position) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  const SheetPrintSettings& print = wb->workbook().sheet(sheet_index).print_settings();
  const std::vector<ManualBreak>& breaks = rows ? print.manual_row_breaks : print.manual_col_breaks;
  if (auto rc = check_index(index, breaks.size(), "index"); rc != 0) {
    return rc;
  }
  const ManualBreak& entry = breaks[index];
  out->*position = entry.id;
  out->first = entry.min;
  out->last = entry.max;
  out->manual = entry.manual ? FM_TRISTATE_TRUE : FM_TRISTATE_FALSE;
  return 0;
}

}  // namespace

/* -------------------------------------------------------------------------- */
/* Manual page breaks                                                         */
/* -------------------------------------------------------------------------- */

extern "C" fm_status_t fm_sheet_set_row_break(fm_workbook_t* wb, size_t sheet_index, uint32_t row,
                                              const fm_row_break_t* brk) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (brk != nullptr) {
    if (auto rc = check_record_size(*brk); rc != 0) {
      return rc;
    }
    return set_break(wb, sheet_index, row, /*rows=*/true, &brk->first, &brk->last, &brk->manual);
  }
  return set_break(wb, sheet_index, row, /*rows=*/true, nullptr, nullptr, nullptr);
}

extern "C" fm_status_t fm_sheet_set_column_break(fm_workbook_t* wb, size_t sheet_index, uint32_t col,
                                                 const fm_column_break_t* brk) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  if (brk != nullptr) {
    if (auto rc = check_record_size(*brk); rc != 0) {
      return rc;
    }
    return set_break(wb, sheet_index, col, /*rows=*/false, &brk->first, &brk->last, &brk->manual);
  }
  return set_break(wb, sheet_index, col, /*rows=*/false, nullptr, nullptr, nullptr);
}

extern "C" fm_status_t fm_sheet_clear_breaks(fm_workbook_t* wb, size_t sheet_index) {
  clear_last_error();
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheet); rc != 0) {
    return rc;
  }
  SheetPrintSettings& print = wb->workbook().sheet(sheet_index).mutable_print_settings();
  print.manual_row_breaks.clear();
  print.manual_col_breaks.clear();
  return 0;
}

extern "C" fm_status_t fm_sheet_get_row_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  return break_count(wb, sheet_index, /*rows=*/true, out_count);
}

extern "C" fm_status_t fm_sheet_get_row_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                 fm_row_break_t* out) {
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  return break_at(wb, sheet_index, index, /*rows=*/true, out, &fm_row_break_t::row);
}

extern "C" fm_status_t fm_sheet_get_column_break_count(const fm_workbook_t* wb, size_t sheet_index, size_t* out_count) {
  if (out_count == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  return break_count(wb, sheet_index, /*rows=*/false, out_count);
}

extern "C" fm_status_t fm_sheet_get_column_break_at(const fm_workbook_t* wb, size_t sheet_index, size_t index,
                                                    fm_column_break_t* out) {
  if (out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_sheet(wb, sheet_index, SheetUse::kWorksheetOrChartsheet); rc != 0) {
    return rc;
  }
  return break_at(wb, sheet_index, index, /*rows=*/false, out, &fm_column_break_t::col);
}
