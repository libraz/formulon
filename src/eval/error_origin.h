//
// The per-thread record of where the error a formula is producing came from.
// Excel saves a `#CALC!`, a `#SPILL!` and some `#VALUE!` results with a rich
// detail that depends on whether the formula made the error itself (and how)
// or received it from another cell. Producers note their origin on their
// error path only; a function that consumes an error forgets it; and the
// recalc commit reads what is left (see `evaluate_cell_for_recalc`). Reading
// a cell notes nothing, so an empty slot at commit means the error arrived
// from elsewhere.

#ifndef FORMULON_EVAL_ERROR_ORIGIN_H_
#define FORMULON_EVAL_ERROR_ORIGIN_H_

#include <cstdint>

#include "sheet_metadata.h"

namespace formulon::eval {

/// What the slot holds: nothing, a measured producer's detail, a producer
/// whose detail is not known, or a spill refusal made at the evaluating anchor.
enum class ErrorOrigin : std::uint8_t { kNone, kMeasured, kUnknown, kRefusedHere };

/// A producer whose rich detail is measured made an error.
void note_error_origin(CellErrorDetail detail) noexcept;

/// `note_error_origin` for a measured `#CALC!` with only a subType.
void note_calc_sub_type(std::uint32_t sub_type) noexcept;

/// A producer of `#CALC!` (or of a measured family's `#VALUE!`) whose detail
/// is not known made an error; the cell is saved with a plain error.
void note_unknown_error_origin() noexcept;

/// The evaluating anchor refused its own spill and recorded the footprint.
void note_spill_refused_here() noexcept;

/// An error-consuming function consumed an error, so no earlier origin
/// reaches the result through it.
void forget_error_origin() noexcept;

/// The slot's current state.
ErrorOrigin error_origin() noexcept;

/// The detail of the last measured note; meaningful while `error_origin()` is
/// `kMeasured`.
CellErrorDetail noted_error_detail() noexcept;

/// The slot's state at one moment, for the left error operand to win.
struct ErrorOriginSnapshot {
  ErrorOrigin origin = ErrorOrigin::kNone;
  CellErrorDetail detail;
};

/// The slot as it stands; taken when the first error operand is produced.
ErrorOriginSnapshot snapshot_error_origin() noexcept;

/// Puts back `snapshot` once that first error operand becomes the result.
void restore_error_origin(const ErrorOriginSnapshot& snapshot) noexcept;

/// Bounds one evaluation that owns its result's origin: empties the slot on
/// entry and restores the enclosing evaluation's on exit.
class ErrorOriginScope {
 public:
  ErrorOriginScope() noexcept;
  ~ErrorOriginScope() noexcept;

  ErrorOriginScope(const ErrorOriginScope&) = delete;
  ErrorOriginScope& operator=(const ErrorOriginScope&) = delete;
  ErrorOriginScope(ErrorOriginScope&&) = delete;
  ErrorOriginScope& operator=(ErrorOriginScope&&) = delete;

 private:
  ErrorOrigin saved_origin_;
  CellErrorDetail saved_detail_;
};

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_ERROR_ORIGIN_H_
