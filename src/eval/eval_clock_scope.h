#ifndef FORMULON_EVAL_EVAL_CLOCK_SCOPE_H_
#define FORMULON_EVAL_EVAL_CLOCK_SCOPE_H_

#include <optional>

#include "utils/date_time.h"

namespace formulon::eval {
namespace detail {

inline thread_local std::optional<date_time::CivilTime> g_current_eval_pinned_now;

}  // namespace detail

/// The wall clock of the evaluation running on this thread: the workbook's
/// pinned reading while a `ScopedEvalClock` carries one, else the host clock.
/// Serves the paths that hold no `EvalContext` (implicit text coercion,
/// criteria, eager date builtins).
inline date_time::CivilTime current_eval_now() noexcept {
  return detail::g_current_eval_pinned_now.has_value() ? *detail::g_current_eval_pinned_now
                                                       : date_time::host_civil_time();
}

class ScopedEvalClock {
 public:
  explicit ScopedEvalClock(const std::optional<date_time::CivilTime>& pinned_now) noexcept
      : previous_(detail::g_current_eval_pinned_now) {
    detail::g_current_eval_pinned_now = pinned_now;
  }

  ~ScopedEvalClock() noexcept { detail::g_current_eval_pinned_now = previous_; }

  ScopedEvalClock(const ScopedEvalClock&) = delete;
  ScopedEvalClock& operator=(const ScopedEvalClock&) = delete;
  ScopedEvalClock(ScopedEvalClock&&) = delete;
  ScopedEvalClock& operator=(ScopedEvalClock&&) = delete;

 private:
  std::optional<date_time::CivilTime> previous_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_EVAL_CLOCK_SCOPE_H_
