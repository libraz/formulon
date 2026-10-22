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
/// Round-trip metadata surfaced through the bindings so a host UI can
/// mirror Excel's user-visible state. The one effect on evaluation:
/// `kAutoNoTable` makes `Workbook::recalc()` skip What-If data tables,
/// which `Workbook::recalc_data_tables()` then computes on demand.
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

/// Cell-reference style of the `<calcPr refMode>` attribute: the notation
/// the host UI shows formulas in.
enum class RefMode : std::uint8_t {
  kA1 = 0,
  kR1C1 = 1,
};

/// The `<calcPr>` attributes other than `calcMode` and the iterative trio.
/// Round-trip metadata the host UI mirrors; the engine itself evaluates the
/// same way whatever they hold. The defaults are the schema's, so a
/// workbook that never set them writes none of them.
struct CalcOptions {
  /// `fullPrecision`: false is "set precision as displayed".
  bool full_precision = true;
  bool calc_on_save = true;  ///< `calcOnSave`.
  RefMode ref_mode = RefMode::kA1;
  bool has_calc_id = false;     ///< `calcId` was present.
  std::uint32_t calc_id = 0;    ///< `calcId`: the calc engine version that last computed the file.
  bool concurrent_calc = true;  ///< `concurrentCalc`.
  bool has_concurrent_manual_count = false;
  std::uint32_t concurrent_manual_count = 0;  ///< `concurrentManualCount`.
  bool force_full_calc = false;               ///< `forceFullCalc`.
  bool full_calc_on_load = false;             ///< `fullCalcOnLoad`.
  bool calc_completed = true;                 ///< `calcCompleted`.

  /// Whether every attribute holds its schema default, i.e. nothing to write.
  bool is_default() const noexcept {
    return full_precision && calc_on_save && ref_mode == RefMode::kA1 && !has_calc_id && concurrent_calc &&
           !has_concurrent_manual_count && !force_full_calc && !full_calc_on_load && calc_completed;
  }
};

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
