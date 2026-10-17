//
// The per-thread mark that makes a formula hold the number 0 instead of the
// value its evaluation produced. Excel leaves a formula at 0 when it
// abandons it (an evaluated array past `kMaxEvaluatedArrayElements`) and
// when the formula's cell is the only member of a cycle with iterative
// calculation off. Either can arise arbitrarily deep inside the evaluation,
// under any enclosing function, so it is a mark the top-level evaluation
// reads rather than an error value the enclosing functions could inspect.

#ifndef FORMULON_EVAL_FORMULA_ABANDON_H_
#define FORMULON_EVAL_FORMULA_ABANDON_H_

namespace formulon::eval {
namespace detail {

inline thread_local bool g_formula_abandoned = false;

}  // namespace detail

/// Marks the formula being evaluated on this thread as abandoned; its top-level
/// evaluation then yields the number 0 (see `ScopedFormulaAbandon`).
inline void abandon_formula() noexcept {
  detail::g_formula_abandoned = true;
}

/// Bounds one top-level formula evaluation: clears the abandonment mark on
/// entry and restores the enclosing formula's on exit, so a formula read
/// through a reference abandons only itself.
class ScopedFormulaAbandon {
 public:
  ScopedFormulaAbandon() noexcept : previous_(detail::g_formula_abandoned) { detail::g_formula_abandoned = false; }
  ~ScopedFormulaAbandon() noexcept { detail::g_formula_abandoned = previous_; }

  ScopedFormulaAbandon(const ScopedFormulaAbandon&) = delete;
  ScopedFormulaAbandon& operator=(const ScopedFormulaAbandon&) = delete;
  ScopedFormulaAbandon(ScopedFormulaAbandon&&) = delete;
  ScopedFormulaAbandon& operator=(ScopedFormulaAbandon&&) = delete;

  /// True when the formula evaluated in this scope was abandoned.
  bool abandoned() const noexcept { return detail::g_formula_abandoned; }

 private:
  bool previous_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_FORMULA_ABANDON_H_
