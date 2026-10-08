#ifndef FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
#define FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_

#include "excel_profile.h"

namespace formulon::eval {
namespace detail {

inline thread_local ExcelProfile g_current_eval_profile = default_excel_profile();

}  // namespace detail

inline ExcelProfile current_eval_profile() noexcept {
  return detail::g_current_eval_profile;
}

class ScopedEvalProfile {
 public:
  explicit ScopedEvalProfile(ExcelProfile profile) noexcept : previous_(detail::g_current_eval_profile) {
    detail::g_current_eval_profile = profile;
  }

  ~ScopedEvalProfile() noexcept { detail::g_current_eval_profile = previous_; }

  ScopedEvalProfile(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile& operator=(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile(ScopedEvalProfile&&) = delete;
  ScopedEvalProfile& operator=(ScopedEvalProfile&&) = delete;

 private:
  ExcelProfile previous_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
