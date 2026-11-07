//
// Session-scoped identity of a model object that lives by value in a
// container (a table in `Workbook::tables()`, a PivotTable on its sheet).
//
// The id is never saved. It is assigned lazily by the C API from one
// workbook-wide counter, so every creation path (load, add, copy) leaves it
// zero until someone asks. The special members encode what identity means
// across the container operations the engine performs:
//   - copy construction yields zero: a copy is a new object (copySheet,
//     `clone()`) or a snapshot, never the original;
//   - move construction and move assignment carry the value and zero the
//     source, so vector reallocation, the shift `erase` performs and a
//     filtered rebuild keep the id on the element it belongs to;
//   - copy assignment keeps the target's value: overwriting an object's
//     content does not change which object it is.

#ifndef FORMULON_UTILS_RUNTIME_ID_H_
#define FORMULON_UTILS_RUNTIME_ID_H_

#include <cstdint>

namespace formulon {

/// Lazily assigned, never-persisted object identity. Zero means "not yet
/// assigned".
class RuntimeId {
 public:
  RuntimeId() noexcept = default;
  RuntimeId(const RuntimeId& /*other*/) noexcept {}
  RuntimeId(RuntimeId&& other) noexcept : value_(other.value_) { other.value_ = 0; }
  RuntimeId& operator=(const RuntimeId& /*other*/) noexcept { return *this; }
  RuntimeId& operator=(RuntimeId&& other) noexcept {
    if (this != &other) {
      value_ = other.value_;
      other.value_ = 0;
    }
    return *this;
  }
  ~RuntimeId() = default;

  /// The assigned id, or zero.
  std::uint32_t value() const noexcept { return value_; }

  /// Stores `value`; used to assign a fresh id and to restore one across a
  /// rollback that replaced the owning object.
  void assign(std::uint32_t value) noexcept { value_ = value; }

 private:
  std::uint32_t value_ = 0;
};

}  // namespace formulon

#endif  // FORMULON_UTILS_RUNTIME_ID_H_
