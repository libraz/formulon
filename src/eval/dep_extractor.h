// Copyright 2026 libraz. Licensed under the MIT License.
//
// AST -> dependency-graph adapter.
//
// `extract_deps()` walks a parser AST and gathers two pieces of information
// the recalc engine needs at registration time:
//
//   1. The set of *cell* references the formula reads, expressed as
//      `CellNodeId`s. Plain `NodeKind::Ref` nodes contribute one cell each;
//      `NodeKind::RangeOp` rectangles are flattened into the cells they
//      cover. Whole-column / whole-row references are NOT flattened (the
//      sheet-wide expansion would explode to ~16M cells); instead the
//      tracker promotes the formula to a volatile-ish status so dependents
//      always recompute on a sheet change. Sheet qualifiers are resolved
//      against the bound `Workbook` so cross-sheet references land on the
//      correct `sheet_id`.
//   2. Whether any function call in the tree is one of the nine Excel
//      Volatile functions (per `VolatileTracker::is_volatile_function`).
//      The recalc engine uses this to add the formula's cell to the
//      always-dirty seed set.
//
// External / Defined-name / Structured (table) references are intentionally
// out of scope at this stage: `extract_deps()` skips them silently. They
// will be wired in once the workbook layer learns to resolve them; see the
// `// TODO` markers in the implementation.
//
// LAMBDA bodies are not descended into here. Captures and parameter lookups
// happen at evaluator time via the LET / LAMBDA name environment, so static
// analysis of the body before binding would either over-approximate (treat
// every captured name as a workbook ref) or be flat-out incorrect.

#ifndef FORMULON_EVAL_DEP_EXTRACTOR_H_
#define FORMULON_EVAL_DEP_EXTRACTOR_H_

#include <cstdint>
#include <vector>

#include "eval/dep_graph.h"
#include "parser/ast.h"

namespace formulon {

class Workbook;

namespace eval {

/// Output of a single `extract_deps()` walk.
///
/// `cell_deps` is the deduplicated list of cells the formula reads. Order is
/// not guaranteed; callers that need a stable order must sort externally.
/// `is_volatile` is true when any function-call node in the AST names one
/// of the nine Excel Volatile functions, OR when the formula touches a
/// whole-column / whole-row reference (those are conservatively treated as
/// volatile because we cannot enumerate every cell on the sheet).
struct ExtractedDeps {
  std::vector<CellNodeId> cell_deps;
  bool is_volatile = false;
};

/// Walks `node` and reports the cell dependencies and volatility status of
/// the formula it represents.
///
/// `current_sheet_id` is the sheet that owns the formula being analysed:
/// unqualified references resolve to that sheet. `workbook` is consulted
/// for sheet-qualified references; passing a workbook with no matching
/// sheet causes the qualified reference to be skipped (it is neither a
/// dependency nor a volatility trigger). The function never mutates
/// `workbook`.
///
/// `RangeOp` rectangles are flattened cell-by-cell unless an endpoint is a
/// whole-column / whole-row reference, in which case the range is omitted
/// from `cell_deps` and `is_volatile` is set to true. Endpoints that are
/// not plain `Ref` nodes (e.g. OFFSET / INDIRECT call results) are ignored;
/// dynamic ranges are out of scope for static dependency analysis.
ExtractedDeps extract_deps(const parser::AstNode& node, std::uint16_t current_sheet_id, const Workbook& workbook);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DEP_EXTRACTOR_H_
