#pragma once

// The C ABI takes input records as `const T*`. This lets a test pass a
// temporary record, which lives until the end of the full-expression.
template <typename T>
const T* record_arg(const T& record) {
  return &record;
}
