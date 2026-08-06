//
// Implementation of the test-side Arena accessors declared in
// `test_arena.h`. The thread-local storage lives here rather than the
// header so every TU that includes the header sees the same Arena
// instance per thread, instead of a copy per TU.

#include "util/test_arena.h"

namespace formulon {
namespace test {

Arena& test_parse_arena() noexcept {
  static thread_local Arena arena;
  arena.reset();
  return arena;
}

Arena& test_eval_arena() noexcept {
  static thread_local Arena arena;
  arena.reset();
  return arena;
}

}  // namespace test
}  // namespace formulon
