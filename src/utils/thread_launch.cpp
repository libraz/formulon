//
// Implementation of the non-throwing thread launcher. See
// thread_launch.h for the contract.
//
// Two backends: `_beginthreadex` on Windows (rather than
// `CreateThread`, which leaves the CRT's per-thread state
// uninitialised), and `pthread_create` everywhere else. Emscripten uses
// the pthread backend, where a request that exceeds `PTHREAD_POOL_SIZE`
// comes back as `EAGAIN` — the same failure the native limits produce,
// so the callers need no per-platform handling.

#include "utils/thread_launch.h"

#include <atomic>
#include <cstdint>
#include <string>
#include <utility>

#if defined(_WIN32)
#include <errno.h>
#include <process.h>
#include <windows.h>
#else
#include <cerrno>
#endif

#include "utils/error.h"
#include "utils/expected.h"

namespace formulon {
namespace {

/// Remaining successful launches before injected failures begin, or -1
/// when no injection is active. Relaxed throughout: the value is written
/// by a test between launches, never concurrently with one, and a stale
/// read would only mis-time an injected failure rather than corrupt
/// anything.
std::atomic<std::int64_t> g_launches_until_failure{-1};

/// Returns true when this launch should fail without touching the OS,
/// consuming one of the allowance's remaining successes otherwise.
bool ConsumeInjectedFailure() {
  const std::int64_t remaining = g_launches_until_failure.load(std::memory_order_relaxed);
  if (remaining < 0) {
    return false;
  }
  if (remaining == 0) {
    return true;
  }
  g_launches_until_failure.store(remaining - 1, std::memory_order_relaxed);
  return false;
}

Error LaunchError(int err) {
  // `EAGAIN` is what both a per-user thread ceiling and an exhausted
  // Emscripten pthread pool report, and `ENOMEM` covers a failed stack
  // allocation; both are resource exhaustion the caller can retry or
  // degrade around, so they map to `kOutOfMemory` rather than to a
  // generic internal fault.
  const FormulonErrorCode code =
      (err == EAGAIN || err == ENOMEM) ? FormulonErrorCode::kOutOfMemory : FormulonErrorCode::kInternalError;
  std::string ctx("context=thread_launch errno=");
  ctx.append(std::to_string(err));
  return make_error(code, "failed to start a thread", std::move(ctx));
}

#if defined(_WIN32)
unsigned __stdcall ThreadTrampoline(void* raw) {
  auto* start = static_cast<ThreadStart*>(raw);
  start->entry(start->arg);
  return 0U;
}
#else
void* ThreadTrampoline(void* raw) {
  auto* start = static_cast<ThreadStart*>(raw);
  start->entry(start->arg);
  return nullptr;
}
#endif

}  // namespace

Thread::~Thread() {
  join();
}

Thread::Thread(Thread&& other) noexcept : joinable_(other.joinable_), handle_(other.handle_) {
  other.joinable_ = false;
}

Thread& Thread::operator=(Thread&& other) noexcept {
  if (this == &other) {
    return *this;
  }
  // Joining first keeps the invariant that a `Thread` never drops a
  // running thread on the floor, matching the destructor.
  join();
  joinable_ = other.joinable_;
  handle_ = other.handle_;
  other.joinable_ = false;
  return *this;
}

void Thread::join() noexcept {
  if (!joinable_) {
    return;
  }
  joinable_ = false;
#if defined(_WIN32)
  WaitForSingleObject(handle_, INFINITE);
  CloseHandle(handle_);
  handle_ = nullptr;
#else
  pthread_join(handle_, nullptr);
#endif
}

Expected<Thread, Error> launch_thread(ThreadStart& start) {
  if (start.entry == nullptr) {
    return make_error(FormulonErrorCode::kInvalidArgument, "thread start record has no entry point",
                      "context=thread_launch");
  }
  if (ConsumeInjectedFailure()) {
    return LaunchError(EAGAIN);
  }

  Thread thread;
#if defined(_WIN32)
  const uintptr_t handle = _beginthreadex(nullptr, 0U, &ThreadTrampoline, &start, 0U, nullptr);
  if (handle == 0U) {
    return LaunchError(errno);
  }
  thread.handle_ = reinterpret_cast<void*>(handle);
#else
  // `pthread_create` reports through its return value; `errno` is not
  // set, which is the whole reason this path can stay non-throwing.
  const int create_rc = pthread_create(&thread.handle_, nullptr, &ThreadTrampoline, &start);
  if (create_rc != 0) {
    return LaunchError(create_rc);
  }
#endif
  thread.joinable_ = true;
  return thread;
}

void set_thread_launch_failure_after(std::uint32_t successes) {
  g_launches_until_failure.store(static_cast<std::int64_t>(successes), std::memory_order_relaxed);
}

void clear_thread_launch_failure_injection() {
  g_launches_until_failure.store(-1, std::memory_order_relaxed);
}

}  // namespace formulon
