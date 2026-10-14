// Runs a test body on a thread with a small stack, so a walk that recurses
// once per nesting level overflows natively as it would in WASM.

#ifndef FORMULON_TESTS_UNIT_SMALL_STACK_H_
#define FORMULON_TESTS_UNIT_SMALL_STACK_H_

#include <pthread.h>

#include <cstddef>

#include "gtest/gtest.h"

namespace formulon {
namespace testing_support {

/// Runs `body` to completion on a thread whose stack is `stack_bytes`.
template <typename Body>
void RunOnSmallStack(Body body, std::size_t stack_bytes = 256U * 1024U) {
  pthread_attr_t attr;
  ASSERT_EQ(pthread_attr_init(&attr), 0);
  ASSERT_EQ(pthread_attr_setstacksize(&attr, stack_bytes), 0);
  pthread_t thread;
  const auto trampoline = [](void* arg) -> void* {
    (*static_cast<Body*>(arg))();
    return nullptr;
  };
  ASSERT_EQ(pthread_create(&thread, &attr, trampoline, &body), 0);
  ASSERT_EQ(pthread_join(thread, nullptr), 0);
  pthread_attr_destroy(&attr);
}

}  // namespace testing_support
}  // namespace formulon

#endif  // FORMULON_TESTS_UNIT_SMALL_STACK_H_
