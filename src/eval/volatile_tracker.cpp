//
// Implementation of the `VolatileTracker` name classifiers. The set of
// volatile functions is fixed by Excel's contract; which of them resolve
// their reads at evaluation time follows from their argument shapes.

#include "eval/volatile_tracker.h"

#include <string_view>

#include "utils/strings.h"

namespace formulon::eval {

bool VolatileTracker::is_volatile_function(std::string_view name) {
  // Function names may reach this classifier in any case (a hand-typed
  // `=now()` keeps its lexeme as written), so comparisons are
  // case-insensitive over ASCII letters. Switch on the upper-cased first
  // character to keep the linear comparison list small.
  if (name.empty()) {
    return false;
  }
  using strings::ascii_to_upper;
  using strings::case_insensitive_eq;
  switch (ascii_to_upper(name.front())) {
    case 'C':
      return case_insensitive_eq(name, "CELL");
    case 'I':
      return case_insensitive_eq(name, "INDIRECT") || case_insensitive_eq(name, "INFO");
    case 'N':
      return case_insensitive_eq(name, "NOW");
    case 'O':
      return case_insensitive_eq(name, "OFFSET");
    case 'R':
      return case_insensitive_eq(name, "RAND") || case_insensitive_eq(name, "RANDBETWEEN") ||
             case_insensitive_eq(name, "RANDARRAY");
    case 'T':
      return case_insensitive_eq(name, "TODAY");
    default:
      return false;
  }
}

bool VolatileTracker::is_dynamic_reference_function(std::string_view name) {
  // `INDIRECT` builds its reference from text and `OFFSET` derives one
  // from a base plus runtime displacements, so neither read is visible to
  // the static dependency extractor. Every other volatile function reads
  // the host environment (clock, RNG, workbook metadata) or a reference
  // its caller spelled out, both of which the graph already describes.
  if (name.empty()) {
    return false;
  }
  using strings::ascii_to_upper;
  using strings::case_insensitive_eq;
  switch (ascii_to_upper(name.front())) {
    case 'I':
      return case_insensitive_eq(name, "INDIRECT");
    case 'O':
      return case_insensitive_eq(name, "OFFSET");
    default:
      return false;
  }
}

}  // namespace formulon::eval
