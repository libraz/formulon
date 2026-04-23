// Copyright 2026 libraz. Licensed under the MIT License.
//
// `NameEnv` is the evaluator's lexical-scope stack for LET bindings (and,
// eventually, LAMBDA parameters). It is a singly-linked list of
// `(name, value, prev)` frames allocated in the caller's `Arena`, read-only
// once built, and lookup is case-insensitive over ASCII letters to match
// Excel's name-matching rules.
//
// The intentional design properties:
//   * Immutable frames: extending the environment produces a new head frame
//     and leaves the tail untouched, so a parent binding never sees a later
//     shadow. This is how `LET(x, 1, x, x+10, x)` returns 11 -- the second
//     `x` binds a new frame whose initialiser reads the outer `x` from the
//     parent chain.
//   * Arena-allocated: every frame lives in the evaluator arena, so the
//     environment's lifetime is bounded by the `evaluate()` call. No manual
//     teardown required.
//   * O(depth) lookup: lookup walks from head to tail. LET depth is
//     practically small (Excel's corpus peaks in the single digits), so this
//     is cheaper than maintaining a hash map and keeps the size budget tight.
//
// Thread safety: `NameEnv` is a stack-local view that is only read on the
// thread currently evaluating a formula. Lookup takes `const NameEnv*`; the
// environment is never mutated after construction. Two sibling LET frames in
// different calc-graph nodes therefore never interact.

#ifndef FORMULON_EVAL_NAME_ENV_H_
#define FORMULON_EVAL_NAME_ENV_H_

#include <cstddef>
#include <string_view>

#include "utils/arena.h"
#include "utils/strings.h"
#include "value.h"

namespace formulon {
namespace eval {

/// Immutable, arena-allocated linked-frame environment.
///
/// Construct the root environment via the default constructor (zero bindings),
/// then grow it with `extend()` for each LET binding. The resulting head can
/// be handed to the evaluator as a pointer; `nullptr` is a valid "no bindings"
/// sentinel that `lookup` understands.
class NameEnv {
 public:
  /// Builds an empty environment (no bindings).
  NameEnv() noexcept = default;

  /// Returns a pointer to the bound Value for `name`, or `nullptr` when the
  /// name is not in scope. Case-insensitive over ASCII letters; non-ASCII
  /// bytes compare verbatim. Walks frames from head to tail so a freshly
  /// extended binding shadows earlier ones with the same name.
  const Value* lookup(std::string_view name) const noexcept {
    for (const Binding* b = head_; b != nullptr; b = b->prev) {
      if (strings::case_insensitive_eq(b->name, name)) {
        return &b->value;
      }
    }
    return nullptr;
  }

  /// Returns a new environment with `(name, value)` pushed on top. The
  /// returned env shares its tail with `*this`; the underlying `Binding`
  /// node lives in `arena`. Returns a copy of `*this` unchanged if arena
  /// allocation fails (the caller surfaces the resulting `#NAME?` on lookup).
  NameEnv extend(std::string_view name, Value value, Arena& arena) const noexcept {
    auto* frame = arena.create<Binding>();
    if (frame == nullptr) {
      return *this;
    }
    // Intern the name so the environment owns the byte storage independently
    // of the AST node that produced it. LET binding names already live in
    // the parser arena for the duration of evaluation, but interning keeps
    // the invariant explicit for future callers (e.g. LAMBDA).
    frame->name = arena.intern(name);
    frame->value = value;
    frame->prev = head_;
    NameEnv next;
    next.head_ = frame;
    return next;
  }

  /// Returns the number of bindings currently visible. O(depth); intended
  /// for tests and diagnostics only.
  std::size_t depth() const noexcept {
    std::size_t n = 0;
    for (const Binding* b = head_; b != nullptr; b = b->prev) {
      ++n;
    }
    return n;
  }

 private:
  struct Binding {
    // `Value`'s default constructor is private; initialise explicitly via
    // `blank()` so this struct is default-constructible for placement-new in
    // `Arena::create<Binding>()`.
    Binding() noexcept : value(Value::blank()) {}
    std::string_view name;
    Value value;
    const Binding* prev = nullptr;
  };

  const Binding* head_ = nullptr;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_NAME_ENV_H_
