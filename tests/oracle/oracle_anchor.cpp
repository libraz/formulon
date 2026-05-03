// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#include "tests/oracle/oracle_anchor.h"

#include "value.h"

namespace formulon {
namespace tests {
namespace oracle {

const Value& anchor_or_self(const Value& v) {
  if (!v.is_array()) {
    return v;
  }
  if (v.as_array_rows() > 0 && v.as_array_cols() > 0) {
    return v.as_array_cells()[0];
  }
  return v;
}

}  // namespace oracle
}  // namespace tests
}  // namespace formulon
