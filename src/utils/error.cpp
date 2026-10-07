//
// Out-of-line `make_error` overloads for literal message and context text.

#include "utils/error.h"

#include <string>
#include <utility>

namespace formulon {

Error make_error(FormulonErrorCode code, const char* message) {
  Error err;
  err.code = code;
  err.message = message;
  return err;
}

Error make_error(FormulonErrorCode code, const char* message, const char* context) {
  Error err;
  err.code = code;
  err.message = message;
  err.context = context;
  return err;
}

Error make_error(FormulonErrorCode code, const char* message, std::string context) {
  Error err;
  err.code = code;
  err.message = message;
  err.context = std::move(context);
  return err;
}

}  // namespace formulon
