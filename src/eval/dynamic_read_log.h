//
// Per-recalc record of the rectangles OFFSET / INDIRECT resolved to.
//
// A formula text cannot say which cells an OFFSET or INDIRECT reads, so the
// recalc engine observes the reads as the formulas run and learns graph
// edges from them after each wave. Every evaluated cell receives an
// ordinal in evaluation order; a reader whose target was committed at or
// after the reader's own ordinal read a value that was not final yet (or
// read itself) and has to be scheduled again behind the learned edge.
//
// The log also keeps the value a dynamic reader held before this recalc
// first overwrote it, which is what a cell that turns out to be circular
// shows again: Excel leaves a circular cell at its last value.
//
// Not thread-safe; the parallel scheduler evaluates every observed reader
// on its calling thread, so one log serves a whole pass. The workbook's
// sheet list cannot change during a recalc (every mutation takes the engine
// lock the pass holds), so a read's sheet is resolved when it is recorded.

#ifndef FORMULON_EVAL_DYNAMIC_READ_LOG_H_
#define FORMULON_EVAL_DYNAMIC_READ_LOG_H_

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <vector>

#include "eval/declared_rect.h"
#include "eval/dep_graph.h"
#include "value.h"

namespace formulon {

class Sheet;
class Workbook;

namespace eval {

class VolatileTracker;
struct EvaluateCellOptions;

/// One rectangle an OFFSET / INDIRECT in `reader` resolved to, on sheet
/// `sheet_id` (at or past the sheet count when the qualifier names no sheet).
struct DynamicRead {
  CellNodeId reader;
  std::uint64_t ordinal = 0;
  std::size_t sheet_id = 0;
  DeclaredRect rect;
};

class DynamicReadLog {
 public:
  /// `volatiles` and `graph` decide which cells are observed; all three
  /// must outlive the log.
  DynamicReadLog(const Workbook& workbook, const VolatileTracker& volatiles, const DepGraph& graph) noexcept
      : workbook_(workbook), volatiles_(volatiles), graph_(graph) {}

  /// Whether `cell` resolves references at evaluation time: it holds an
  /// OFFSET / INDIRECT, or it learned a dynamic-reference edge before.
  bool observes(CellNodeId cell) const;

  /// When `observes(cell)`, attributes the reads of `cell`'s next
  /// evaluation to this log at `ordinal` by installing the observer in
  /// `opts`; with `sheet`, also keeps the value `cell` shows before that
  /// evaluation.
  void observe(CellNodeId cell, std::uint64_t ordinal, EvaluateCellOptions& opts, const Sheet* sheet);

  /// Stops attributing reads to the cell `observe` named.
  void end() noexcept { active_ = false; }

  /// Records a read by the current reader; ignored outside `observe`/`end`.
  void record(std::string_view sheet, const DeclaredRect& rect);

  /// `DynamicReadCallback` adapter; `user_data` is the log.
  static void record_trampoline(void* user_data, std::string_view sheet, const DeclaredRect& rect);

  /// Notes that `cell` was committed at `ordinal`; the latest one wins.
  void note_commit(CellNodeId cell, std::uint64_t ordinal);

  /// Tags `cell` as a member of cyclic component `component` that the
  /// iterative solver evaluated; reads between members of one such
  /// component are the solver's contract, not stale reads.
  void note_iterative_member(CellNodeId cell, std::uint64_t component);

  /// Writes back the value each of `cells` showed before this recalc first
  /// evaluated it, where one was kept.
  void restore_prior_values(Workbook& workbook, const std::vector<CellNodeId>& cells) const;

  /// Drops the wave records and keeps the prior values.
  void clear_wave();

  const std::unordered_map<CellNodeId, std::uint64_t, CellNodeIdHash>& readers() const noexcept { return readers_; }
  const std::vector<DynamicRead>& reads() const noexcept { return reads_; }
  std::optional<std::uint64_t> commit_ordinal(CellNodeId cell) const;
  std::optional<std::uint64_t> iterative_component(CellNodeId cell) const;

 private:
  // A text `value` views `text`, which the map node keeps at a stable
  // address.
  struct PriorValue {
    Value value;
    std::string text;
  };
  using CellIndex = std::unordered_map<CellNodeId, std::uint64_t, CellNodeIdHash>;

  /// Keeps `prior` as the value `cell` showed before this recalc changed
  /// it, unless one is already kept. Only scalars are kept.
  void keep_prior_value(CellNodeId cell, const Value& prior);

  const Workbook& workbook_;
  const VolatileTracker& volatiles_;
  const DepGraph& graph_;
  CellIndex readers_;
  std::vector<DynamicRead> reads_;
  CellIndex commits_;
  CellIndex iterative_members_;
  std::unordered_map<CellNodeId, PriorValue, CellNodeIdHash> prior_values_;
  CellNodeId current_{};
  std::uint64_t current_ordinal_ = 0;
  bool active_ = false;
};

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DYNAMIC_READ_LOG_H_
