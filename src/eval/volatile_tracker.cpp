//
// Implementation of the `VolatileTracker` name classifiers, which defer to
// the parser-level lists so the file writers mark the same functions.

#include "eval/volatile_tracker.h"

#include <string_view>

#include "parser/ast.h"

namespace formulon::eval {

bool VolatileTracker::is_volatile_function(std::string_view name) {
  return parser::is_volatile_function_name(name);
}

bool VolatileTracker::is_dynamic_reference_function(std::string_view name) {
  return parser::is_dynamic_reference_function_name(name);
}

}  // namespace formulon::eval
