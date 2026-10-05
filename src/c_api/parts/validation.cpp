//
// C ABI - data-validation evaluation: checking a proposed value against the
// rule covering a cell, and enumerating the cells whose value fails theirs.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <memory>
#include <string>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/cell_range.h"
#include "c_api/parts/common.h"
#include "sheet.h"
#include "utils/error.h"
#include "validation_eval.h"
#include "value.h"
#include "workbook.h"

using formulon::c_api::parts::cell_range_page_size;
using formulon::c_api::parts::check_cell_cursor;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;
using formulon::c_api::parts::set_last_error;
using formulon::c_api::parts::value_from_fm;

namespace {

// Cells visited per scan step while looking for invalid ones.
constexpr std::uint32_t kInvalidScanChunk = 4096U;

// One populated cell copied out of the sheet lock. Text payloads keep
// aliasing sheet storage, which no mutation can touch during the call.
struct Candidate {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::string formula_text;
  formulon::Value value = formulon::Value::blank();
};

}  // namespace

extern "C" fm_status_t fm_sheet_validate_value(const fm_workbook_t* wb, size_t sheet_index, uint32_t row, uint32_t col,
                                               const fm_value_t* proposed, fm_validation_outcome* out) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_validate_value";
  if (proposed == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer,
                             "fm_sheet_validate_value: NULL argument");
  }
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  formulon::Value value = formulon::Value::blank();
  if (auto rc = value_from_fm(*proposed, &value, kApi); rc != 0) {
    return rc;
  }
  const formulon::Workbook& book = wb->workbook();
  auto outcome = formulon::validate_value(book, book.sheet(sheet_index), row, col, value);
  if (!outcome) {
    return set_last_error(outcome.error());
  }
  *out = fm_validation_outcome{};
  out->has_rule = outcome.value().has_rule ? 1 : 0;
  out->valid = outcome.value().valid ? 1 : 0;
  out->rule_index = outcome.value().rule_index;
  out->error_style = outcome.value().error_style;
  return 0;
}

extern "C" fm_status_t fm_sheet_list_invalid_cells(const fm_workbook_t* wb, size_t sheet_index, uint64_t cursor,
                                                   uint32_t limit, fm_cell_range_t** out) {
  clear_last_error();
  constexpr const char* kApi = "fm_sheet_list_invalid_cells";
  if (out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_sheet_list_invalid_cells: NULL out");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index, kApi); rc != 0) {
    return rc;
  }
  if (auto rc = check_cell_cursor(cursor, kApi); rc != 0) {
    return rc;
  }
  const formulon::Workbook& book = wb->workbook();
  const formulon::Sheet& sheet = book.sheet(sheet_index);
  auto handle = std::unique_ptr<fm_cell_range_t>(new fm_cell_range_t{});

  // Only cells inside some rule's ranges can fail one, so the scan covers
  // their bounding box.
  bool any_range = false;
  formulon::MergeRange box;
  for (const formulon::DataValidation& rule : sheet.validations()) {
    for (const formulon::MergeRange& r : rule.ranges) {
      if (!any_range) {
        box = r;
        any_range = true;
        continue;
      }
      box.first_row = std::min(box.first_row, r.first_row);
      box.first_col = std::min(box.first_col, r.first_col);
      box.last_row = std::max(box.last_row, r.last_row);
      box.last_col = std::max(box.last_col, r.last_col);
    }
  }
  if (!any_range) {
    *out = handle.release();
    return 0;
  }

  const std::uint32_t page = cell_range_page_size(limit);
  std::uint64_t scan = cursor;
  std::vector<Candidate> chunk;
  while (scan != formulon::Sheet::kCellCursorEnd) {
    chunk.clear();
    const std::uint64_t next = sheet.cells_in_range(
        box.first_row, box.first_col, box.last_row, box.last_col, scan, kInvalidScanChunk,
        [](const formulon::Sheet::RangeCell& cell, void* ctx) {
          static_cast<std::vector<Candidate>*>(ctx)->push_back(
              Candidate{cell.row, cell.col, std::string(cell.formula_text), cell.value});
        },
        &chunk);
    for (const Candidate& candidate : chunk) {
      auto outcome = formulon::validate_value(book, sheet, candidate.row, candidate.col, candidate.value);
      if (!outcome) {
        return set_last_error(outcome.error());
      }
      if (!outcome.value().has_rule || outcome.value().valid) {
        continue;
      }
      const std::uint64_t position =
          static_cast<std::uint64_t>(candidate.row) * formulon::Sheet::kMaxCols + candidate.col;
      if (handle->entries.size() == page) {
        // The page is full: the first further invalid cell is where the
        // next page resumes.
        handle->next_cursor = position;
        *out = handle.release();
        return 0;
      }
      formulon::Sheet::RangeCell cell;
      cell.row = candidate.row;
      cell.col = candidate.col;
      cell.formula_text = candidate.formula_text;
      cell.value = candidate.value;
      handle->append(cell);
    }
    scan = next;
  }
  *out = handle.release();
  return 0;
}
