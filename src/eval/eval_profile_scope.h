#ifndef FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
#define FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_

#include "excel_profile.h"
#include "utils/text_ops.h"

namespace formulon::eval {
namespace detail {

inline thread_local ExcelProfile g_current_eval_profile = default_excel_profile();
inline thread_local int g_eval_profile_scope_depth = 0;
// The position unit of LEN, MID, FIND, SEARCH and REPLACE: code points, a
// workbook's effective default, unless the evaluated workbook is version 1.
inline thread_local TextIndexUnit g_current_text_index_unit = TextIndexUnit::kCodePoints;

}  // namespace detail

inline ExcelProfile current_eval_profile() noexcept {
  return detail::g_current_eval_profile;
}

/// True while a `ScopedEvalProfile` is live on this thread, i.e. when
/// `current_eval_profile()` reflects an entry point rather than the default.
inline bool eval_profile_scope_active() noexcept {
  return detail::g_eval_profile_scope_depth > 0;
}

inline TextIndexUnit current_text_index_unit() noexcept {
  return detail::g_current_text_index_unit;
}

/// Sets the profile, and with the two-argument form the workbook's text
/// position unit, for the evaluation on this thread. The one-argument form
/// keeps the unit already in force (code points outside any workbook).
class ScopedEvalProfile {
 public:
  explicit ScopedEvalProfile(ExcelProfile profile) noexcept
      : ScopedEvalProfile(profile, detail::g_current_text_index_unit) {}

  ScopedEvalProfile(ExcelProfile profile, TextIndexUnit unit) noexcept
      : previous_(detail::g_current_eval_profile), previous_unit_(detail::g_current_text_index_unit) {
    detail::g_current_eval_profile = profile;
    detail::g_current_text_index_unit = unit;
    ++detail::g_eval_profile_scope_depth;
  }

  ~ScopedEvalProfile() noexcept {
    --detail::g_eval_profile_scope_depth;
    detail::g_current_eval_profile = previous_;
    detail::g_current_text_index_unit = previous_unit_;
  }

  ScopedEvalProfile(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile& operator=(const ScopedEvalProfile&) = delete;
  ScopedEvalProfile(ScopedEvalProfile&&) = delete;
  ScopedEvalProfile& operator=(ScopedEvalProfile&&) = delete;

 private:
  ExcelProfile previous_;
  TextIndexUnit previous_unit_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_EVAL_PROFILE_SCOPE_H_
