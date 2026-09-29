//
// Workbook-level calculation settings: `CalcMode` mirrors Excel's
// `<calcPr calcMode="...">` attribute and `IterativeOptions` the "Enable
// iterative calculation" workbook option. They live outside `eval/` and
// `io/` so `Workbook`, the OOXML reader/writer, the recalc engine and the C
// ABI can all name them without including the full `workbook.h`.

#ifndef FORMULON_CALC_SETTINGS_H_
#define FORMULON_CALC_SETTINGS_H_

#include <cstdint>

namespace formulon {

/// Excel calc-mode enum. Mirrors the `calcMode` attribute on `<calcPr>`
/// (`auto` / `manual` / `autoNoTable`).
///
/// Plain metadata: the engine itself does not gate evaluation on this
/// setting (every `recalc()` call honours all dirty cells). The value
/// is preserved as round-trip metadata and surfaced through the
/// bindings so a host UI can mirror Excel's user-visible state.
enum class CalcMode : std::uint8_t {
  kAuto = 0,
  kManual = 1,
  kAutoNoTable = 2,
};

/// Excel's default iteration cap when iterative calc is enabled.
constexpr std::uint32_t kDefaultMaxIterations = 100U;
/// Excel's own upper bound on the iteration count, as enforced by its
/// "Enable iterative calculation" dialog. Every path that populates
/// `IterativeOptions` from outside the engine clamps to this, because a
/// larger cap is not merely slow: the solver has no wall-clock limit and
/// no default cancellation hook, so a workbook asking for billions of
/// sweeps is an unrecoverable hang rather than a long calculation.
/// Clamping here matches the reference implementation instead of
/// trading fidelity for safety.
constexpr std::uint32_t kMaxIterationsCap = 32767U;
/// Excel's default absolute convergence threshold.
constexpr double kDefaultMaxChange = 0.001;

/// User-facing knobs governing iterative-calc behaviour. Mirrors Excel's
/// "Enable iterative calculation" workbook option. The defaults match
/// Excel: iterative calc disabled, capped at `kDefaultMaxIterations`
/// passes, `kDefaultMaxChange` absolute max-change threshold.
struct IterativeOptions {
  /// When false, circular SCCs surface `#REF!` (the legacy behaviour).
  /// When true, the recalc engine hands them to `run_iterative_solve`.
  bool enabled = false;
  /// Maximum number of evaluation passes the solver will run before
  /// giving up. Must be at least 1 to make any progress; the solver
  /// silently treats 0 as 1 to avoid an empty-loop edge case, and
  /// silently treats anything above `kMaxIterationsCap` as that cap.
  std::uint32_t max_iterations = kDefaultMaxIterations;
  /// Absolute convergence threshold. The solver stops as soon as
  /// `max(|v_n - v_{n-1}|) < max_change` across all numeric members. A
  /// non-positive value forces the solver to run for the full
  /// `max_iterations` since strict-less-than against any non-negative
  /// delta will never hold.
  double max_change = kDefaultMaxChange;
};

}  // namespace formulon

#endif  // FORMULON_CALC_SETTINGS_H_
