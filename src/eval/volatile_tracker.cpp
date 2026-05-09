// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Implementation of `VolatileTracker::is_volatile_function`. The set of
// volatile functions is fixed by Excel's contract.

#include "eval/volatile_tracker.h"

#include <string_view>

namespace formulon::eval {

bool VolatileTracker::is_volatile_function(std::string_view name) {
  // Switch on the first character to keep the linear comparison list
  // small. Names are already ASCII-uppercase per the precondition.
  if (name.empty()) {
    return false;
  }
  switch (name.front()) {
    case 'C':
      return name == "CELL";
    case 'I':
      return name == "INDIRECT" || name == "INFO";
    case 'N':
      return name == "NOW";
    case 'O':
      return name == "OFFSET";
    case 'R':
      return name == "RAND" || name == "RANDBETWEEN" || name == "RANDARRAY";
    case 'T':
      return name == "TODAY";
    default:
      return false;
  }
}

}  // namespace formulon::eval
