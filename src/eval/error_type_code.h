#ifndef FORMULON_EVAL_ERROR_TYPE_CODE_H_
#define FORMULON_EVAL_ERROR_TYPE_CODE_H_

#include <optional>

#include "value.h"

namespace formulon::eval {

// The ERROR.TYPE code (`#NULL!` 1 .. `#N/A` 7, `#GETTING_DATA` 8, `#SPILL!` 9,
// `#FIELD!` 13, `#CALC!` 14, as Mac Excel 365 returns them) of an error; empty
// for the errors whose code is unmeasured.
inline std::optional<int> error_type_code(ErrorCode code) noexcept {
  switch (code) {
    case ErrorCode::Null:
      return 1;
    case ErrorCode::Div0:
      return 2;
    case ErrorCode::Value:
      return 3;
    case ErrorCode::Ref:
      return 4;
    case ErrorCode::Name:
      return 5;
    case ErrorCode::Num:
      return 6;
    case ErrorCode::NA:
      return 7;
    case ErrorCode::GettingData:
      return 8;
    case ErrorCode::Spill:
      return 9;
    case ErrorCode::Field:
      return 13;
    case ErrorCode::Calc:
      return 14;
    default:
      return std::nullopt;
  }
}

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_ERROR_TYPE_CODE_H_
