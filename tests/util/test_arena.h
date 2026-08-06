//
// Test-side Arena helpers.
//
// Production code parses formulas into an Arena (`src/utils/arena.h`) and
// evaluates them into a second Arena that owns runtime-derived text. Unit
// tests across `tests/unit/eval/` repeat the same setup boilerplate dozens
// of times:
//
//   static thread_local Arena parse_arena;
//   static thread_local Arena eval_arena;
//   parse_arena.reset();
//   eval_arena.reset();
//
// (248 occurrences across the test tree as of the time this header was
// introduced.) The thread-local form keeps allocator state warm between
// gtest cases inside one binary while the explicit `reset()` calls make the
// behaviour observable: each test sees a freshly emptied arena even though
// the underlying chunks are reused.
//
// This header centralises that pattern. Tests pull it in via
// `tests/util/test_eval_helpers.h` (the evaluation wrappers) or directly
// when they need an Arena for their own AST construction.

#ifndef FORMULON_TESTS_UTIL_TEST_ARENA_H_
#define FORMULON_TESTS_UTIL_TEST_ARENA_H_

#include "utils/arena.h"

namespace formulon {
namespace test {

// RAII-style holder for a per-test Arena.
//
// Construct one in a test function, then call `get()` to hand its inner
// arena to the parser or evaluator. The destructor releases every owned
// chunk; tests that want to retain memory between calls should use
// `test_parse_arena()` / `test_eval_arena()` below instead.
class TestArena {
 public:
  TestArena() noexcept = default;
  ~TestArena() = default;

  TestArena(const TestArena&) = delete;
  TestArena& operator=(const TestArena&) = delete;

  // Returns the inner Arena. The returned reference is valid for the
  // lifetime of `*this`.
  Arena& get() noexcept { return arena_; }
  const Arena& get() const noexcept { return arena_; }

 private:
  Arena arena_{};
};

// Returns a thread-local Arena intended for parser use. The arena is
// reset on every call so the caller sees a freshly emptied allocator,
// but the underlying chunks are retained for reuse across calls within
// the same thread. Drop-in replacement for the open-coded
// `static thread_local Arena parse_arena; parse_arena.reset();` pattern.
Arena& test_parse_arena() noexcept;

// Returns a thread-local Arena intended for evaluator use. Same
// reset-but-reuse semantics as `test_parse_arena()`. Kept distinct so
// tests that hold an AST node and a freshly evaluated Value across the
// same statement do not invalidate either by sharing one arena.
Arena& test_eval_arena() noexcept;

// Backwards-compatible alias. Returns the same arena as
// `test_eval_arena()`; provided so existing call sites that just say
// "give me an arena" do not have to pick a name.
inline Arena& test_arena() noexcept {
  return test_eval_arena();
}

}  // namespace test
}  // namespace formulon

#endif  // FORMULON_TESTS_UTIL_TEST_ARENA_H_
