//
// Shared profile-aware evaluation helpers for unit tests. Keep profile
// selection visible at call sites so Mac/Windows behavior expectations do not
// silently track the runtime default.

#ifndef FORMULON_TESTS_UNIT_EVAL_TEST_EVAL_HELPERS_H_
#define FORMULON_TESTS_UNIT_EVAL_TEST_EVAL_HELPERS_H_

#include "eval/compat.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "sheet.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace test {

inline constexpr ExcelProfile mac_profile() noexcept {
  return mac_365_ja_jp_profile();
}

inline constexpr ExcelProfile win_profile() noexcept {
  return win_365_ja_jp_profile();
}

inline Workbook workbook_with_profile(ExcelProfile profile) {
  Workbook wb = Workbook::create();
  wb.set_excel_profile(profile);
  return wb;
}

inline Workbook mac_workbook() {
  return workbook_with_profile(mac_profile());
}

inline Workbook win_workbook() {
  return workbook_with_profile(win_profile());
}

inline EvalContext context_with_profile(ExcelProfile profile) {
  return EvalContext().with_excel_profile(profile);
}

inline EvalContext mac_context() {
  return context_with_profile(mac_profile());
}

inline EvalContext win_context() {
  return context_with_profile(win_profile());
}

inline EvalContext host_context(ExcelHost host) {
  return context_with_profile(profile_from_host(host));
}

inline EvalContext context_with_profile(const Workbook& wb, const Sheet& sheet, EvalState& state,
                                        ExcelProfile profile) {
  return EvalContext(wb, sheet, state).with_excel_profile(profile);
}

inline EvalContext mac_context(const Workbook& wb, const Sheet& sheet, EvalState& state) {
  return context_with_profile(wb, sheet, state, mac_profile());
}

inline EvalContext win_context(const Workbook& wb, const Sheet& sheet, EvalState& state) {
  return context_with_profile(wb, sheet, state, win_profile());
}

inline EvalContext workbook_context(const Workbook& wb, const Sheet& sheet, EvalState& state) {
  return EvalContext(wb, sheet, state).with_excel_profile(wb.excel_profile());
}

}  // namespace test
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_TESTS_UNIT_EVAL_TEST_EVAL_HELPERS_H_
