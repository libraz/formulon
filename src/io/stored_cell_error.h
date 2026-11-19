//
// The error a workbook file stores as a cell's value.
//
// A file's cell value can only hold the legacy error set. Excel 365 keeps
// every newer error as `#VALUE!` in the cell and carries the real one in a
// rich value (cell `vm` metadata plus the `richData` parts, see
// `io/rich_value_errors.h`); it stores `#GETTING_DATA` as `#N/A`. This holds
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
    // Inferred from #SPILL! / #CALC!: Excel cannot be made to produce these.
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

/// Returns the value a file stores in the cell for `v`: a picture is
/// written as a `#VALUE!` error cell whose `vm` carries the picture.
inline Value stored_cell_value(const Value& v) noexcept {
  return v.is_picture() ? Value::error(ErrorCode::Value) : v;
}

}  // namespace io
}  // namespace formulon

#endif  // FORMULON_IO_STORED_CELL_ERROR_H_
