//
// The error a workbook file stores as a cell's value.
//
// A file's cell value can only hold the legacy error set. Excel 365 keeps
// every newer error as `#VALUE!` in the cell and carries the real one in a
// rich value (cell `vm` metadata plus the `richData` parts), which this
// writer does not produce; it stores `#GETTING_DATA` as `#N/A`. Measured
// for #SPILL!, #CALC! and #GETTING_DATA in both .xlsx and .xlsb; the other
// newer errors cannot be produced by entering a formula and follow the
// same fallback. Every writer of a cell value goes through this one rule.

#ifndef FORMULON_IO_STORED_CELL_ERROR_H_
#define FORMULON_IO_STORED_CELL_ERROR_H_

#include "value.h"

namespace formulon {
namespace io {

/// Returns the error a file stores as the cell value for `e`.
constexpr ErrorCode stored_cell_error(ErrorCode e) noexcept {
  switch (e) {
    case ErrorCode::Null:
    case ErrorCode::Div0:
    case ErrorCode::Value:
    case ErrorCode::Ref:
    case ErrorCode::Name:
    case ErrorCode::Num:
    case ErrorCode::NA:
      return e;
    case ErrorCode::GettingData:
      return ErrorCode::NA;
    case ErrorCode::Spill:
    case ErrorCode::Calc:
    case ErrorCode::Field:
    case ErrorCode::Blocked:
    case ErrorCode::Connect:
    case ErrorCode::External:
    case ErrorCode::Busy:
    case ErrorCode::Python:
    case ErrorCode::Unknown:
      return ErrorCode::Value;
  }
  return ErrorCode::Value;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STORED_CELL_ERROR_H_
