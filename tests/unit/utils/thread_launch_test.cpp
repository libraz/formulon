//
// Unit tests for the non-throwing thread launcher.
//
// The launcher exists because every target is built `-fno-exceptions`,
// where `std::thread`'s constructor turns an OS refusal into process
// termination. The tests below cover the two halves of that contract: a
// launch that succeeds behaves like an ordinary joinable thread, and a
// launch the OS refuses comes back as an `Error` the caller can act on.
// Real thread exhaustion is not reproducible inside a test, so the
// refusal is driven through the launcher's injection hook.

#include "utils/thread_launch.h"

#include <atomic>
#include <cstdint>
#include <utility>

#include "gtest/gtest.h"
#include "utils/error.h"

namespace formulon {
namespace {

/// Clears any injected failure when the test leaves scope, including on
/// the assertion-failure path, so one test cannot poison the next.
struct InjectionGuard {
  ~InjectionGuard() { clear_thread_launch_failure_injection(); }
};

void IncrementCounter(void* arg) {
  static_cast<std::atomic<std::uint32_t>*>(arg)->fetch_add(1U, std::memory_order_relaxed);
}

TEST(ThreadLaunch, LaunchedThreadRunsItsEntryAndJoins) {
  std::atomic<std::uint32_t> ran{0U};
  ThreadStart start{&IncrementCounter, &ran};

  auto thread_or = launch_thread(start);
  ASSERT_TRUE(static_cast<bool>(thread_or)) << thread_or.error().message;
  Thread thread = std::move(thread_or.value());
  EXPECT_TRUE(thread.joinable());

  thread.join();
  EXPECT_FALSE(thread.joinable());
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 1U);

  // Joining again is a documented no-op, which is what makes the
  // destructor's own join safe after an explicit one.
  thread.join();
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 1U);
}

TEST(ThreadLaunch, DestructorJoinsAThreadTheOwnerNeverJoined) {
  // The entry's write must be visible once the `Thread` has been
  // destroyed; if the destructor did not join, this would be a read of a
  // value still being produced.
  std::atomic<std::uint32_t> ran{0U};
  ThreadStart start{&IncrementCounter, &ran};
  {
    auto thread_or = launch_thread(start);
    ASSERT_TRUE(static_cast<bool>(thread_or)) << thread_or.error().message;
    Thread thread = std::move(thread_or.value());
  }
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 1U);
}

TEST(ThreadLaunch, MoveAssignmentTransfersOwnershipOfOneThread) {
  std::atomic<std::uint32_t> ran{0U};
  ThreadStart first{&IncrementCounter, &ran};
  ThreadStart second{&IncrementCounter, &ran};

  auto first_or = launch_thread(first);
  ASSERT_TRUE(static_cast<bool>(first_or)) << first_or.error().message;
  Thread thread = std::move(first_or.value());

  auto second_or = launch_thread(second);
  ASSERT_TRUE(static_cast<bool>(second_or)) << second_or.error().message;
  // Overwriting joins the first thread rather than dropping it.
  thread = std::move(second_or.value());
  EXPECT_TRUE(thread.joinable());

  thread.join();
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 2U);
}

TEST(ThreadLaunch, StartRecordWithoutAnEntryIsRejected) {
  ThreadStart start;
  auto thread_or = launch_thread(start);
  ASSERT_FALSE(static_cast<bool>(thread_or));
  EXPECT_EQ(thread_or.error().code, FormulonErrorCode::kInvalidArgument);
}

TEST(ThreadLaunch, RefusedLaunchIsReportedAsAnErrorNotATermination) {
  InjectionGuard guard;
  std::atomic<std::uint32_t> ran{0U};
  ThreadStart start{&IncrementCounter, &ran};

  set_thread_launch_failure_after(0U);
  auto refused = launch_thread(start);
  ASSERT_FALSE(static_cast<bool>(refused));
  // Thread-count and memory limits both surface as resource exhaustion.
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kOutOfMemory);
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 0U) << "a refused launch must not have run the entry";

  clear_thread_launch_failure_injection();
  auto allowed = launch_thread(start);
  ASSERT_TRUE(static_cast<bool>(allowed)) << allowed.error().message;
  allowed.value().join();
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 1U);
}

TEST(ThreadLaunch, InjectionLetsTheFirstLaunchesThroughBeforeFailing) {
  // A pool that asks for more workers than it gets is the case the
  // scheduler degrades around, so the hook has to be able to express
  // "the Nth launch is the one that fails" and not just "all of them".
  InjectionGuard guard;
  std::atomic<std::uint32_t> ran{0U};
  ThreadStart start{&IncrementCounter, &ran};

  set_thread_launch_failure_after(2U);
  for (std::uint32_t i = 0; i < 2U; ++i) {
    auto thread_or = launch_thread(start);
    ASSERT_TRUE(static_cast<bool>(thread_or)) << "launch " << i << ": " << thread_or.error().message;
    thread_or.value().join();
  }

  auto refused = launch_thread(start);
  ASSERT_FALSE(static_cast<bool>(refused));
  EXPECT_EQ(refused.error().code, FormulonErrorCode::kOutOfMemory);
  EXPECT_EQ(ran.load(std::memory_order_relaxed), 2U);
}

}  // namespace
}  // namespace formulon
