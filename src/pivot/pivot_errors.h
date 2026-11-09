#ifndef FORMULON_PIVOT_PIVOT_ERRORS_H_
#define FORMULON_PIVOT_PIVOT_ERRORS_H_

#include <string>
#include <string_view>
#include <utility>

#include "utils/error.h"

namespace formulon::pivot {

/// An `Error` of code `kInvalidArgument`.
inline Error invalid(const char* message) {
  return make_error(FormulonErrorCode::kInvalidArgument, message);
}

/// An `Error` of code `kInvalidArgument` carrying `context`.
inline Error invalid(const char* message, std::string_view context) {
  return make_error(FormulonErrorCode::kInvalidArgument, message, std::string(context));
}

/// An `Error` of code `kAlreadyExists` carrying `context`.
inline Error name_taken(const char* message, std::string context) {
  return make_error(FormulonErrorCode::kAlreadyExists, message, std::move(context));
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_PIVOT_ERRORS_H_
