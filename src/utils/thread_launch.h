//
// Non-throwing OS thread launch.
//
// Every Formulon target builds with `-fno-exceptions`, which makes
// `std::thread`'s constructor unusable for anything that wants to survive
// a launch failure: libstdc++ reports the failure by throwing
// `std::system_error`, and with exceptions disabled that call site
// terminates the process instead. A machine under thread pressure (an
// `RLIMIT_NPROC` ceiling, an exhausted Emscripten pthread pool, a
// container CPU quota) is exactly where the parallel scheduler is most
// likely to be asked for more workers than the OS will hand out, so the
// engine launches threads through this wrapper and degrades on failure
// rather than aborting the host.
//
// The wrapper is deliberately smaller than `std::thread`: launch, join,
// and nothing else. Detaching is not offered — an unjoined worker
// outliving the state it borrows is the failure mode this layer exists to
// keep out of the engine.

#ifndef FORMULON_UTILS_THREAD_LAUNCH_H_
#define FORMULON_UTILS_THREAD_LAUNCH_H_

#include <cstdint>

#if !defined(_WIN32)
#include <pthread.h>
#endif

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {

/// Thread body. Runs on the new thread and returns to end it.
using ThreadEntry = void (*)(void* arg);

/// The (entry, arg) pair a launched thread reads on start-up.
///
/// The launcher performs no allocation, so the caller owns this record
/// and must keep it at a stable address until the thread is joined.
/// Storing it inside whatever per-worker slot the caller already keeps is
/// the intended usage; a `std::vector` element qualifies only if the
/// vector is reserved up front and never grows afterwards.
struct ThreadStart {
  ThreadEntry entry = nullptr;
  void* arg = nullptr;
};

/// A joinable OS thread.
///
/// Move-only, and joined by the destructor if the owner has not joined it
/// already, so a thread cannot outlive the `Thread` that represents it.
/// A default-constructed (or moved-from) instance represents no thread.
class Thread {
 public:
  Thread() noexcept = default;
  ~Thread();

  Thread(Thread&& other) noexcept;
  Thread& operator=(Thread&& other) noexcept;

  Thread(const Thread&) = delete;
  Thread& operator=(const Thread&) = delete;

  /// True while this object owns a thread that has not been joined.
  bool joinable() const noexcept { return joinable_; }

  /// Blocks until the thread's entry returns. A no-op when not joinable,
  /// so it is safe to call before the destructor's own join.
  void join() noexcept;

 private:
  friend Expected<Thread, Error> launch_thread(ThreadStart& start);

  bool joinable_ = false;
#if defined(_WIN32)
  void* handle_ = nullptr;
#else
  pthread_t handle_{};
#endif
};

/// Starts a thread running `start.entry(start.arg)`.
///
/// Returns `kOutOfMemory` when the OS refuses for lack of resources
/// (`EAGAIN` / `ENOMEM`, which is what a thread-count or memory limit
/// surfaces as) and `kInternalError` for any other refusal. The error's
/// context carries the platform error number. On failure no thread was
/// created and `start` is untouched.
Expected<Thread, Error> launch_thread(ThreadStart& start);

/// Test-only: makes the next `launch_thread` calls fail.
///
/// The next `successes` launches behave normally; every launch after that
/// fails with `kOutOfMemory` until the injection is cleared. This exists
/// so the scheduler's degradation path can be exercised without having to
/// drive a real machine into thread exhaustion, which no test can do
/// reproducibly. Set and clear it from a single thread while no launch is
/// in flight.
void set_thread_launch_failure_after(std::uint32_t successes);

/// Test-only: restores normal launch behaviour.
void clear_thread_launch_failure_injection();

}  // namespace formulon

#endif  // FORMULON_UTILS_THREAD_LAUNCH_H_
