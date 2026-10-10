#ifndef FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
#define FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_

#include "excel_profile.h"

namespace formulon::eval {
namespace detail {

inline thread_local ExcelProfile g_current_eval_profile = default_excel_profile();
inline thread_local int g_eval_profile_scope_depth = 0;

}  // namespace detail

inline ExcelProfile current_eval_profile() noexcept {
  return detail::g_current_eval_profile;
}

/// True while a `ScopedEvalProfile` is live on this thread, i.e. when
/// `current_eval_profile()` reflects an entry point rather than the default.
inline bool eval_profile_scope_active() noexcept {
  return detail::g_eval_profile_scope_depth > 0;
}

class ScopedEvalProfile {
 public:
  explicit ScopedEvalProfile(ExcelProfile profile) noexcept : previous_(detail::g_current_eval_profile) {
    detail::g_current_eval_profile = profile;
    ++detail::g_eval_profile_scope_depth;
  }

  ~ScopedEvalProfile() noexcept {
    --detail::g_eval_profile_scope_depth;
    detail::g_current_eval_profile = previous_;
  }

  ScopedEvalProfile(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile& operator=(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile(ScopedEvalProfile&&) = delete;
  ScopedEvalProfile& operator=(ScopedEvalProfile&&) = delete;

 private:
  ExcelProfile previous_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
