//
// AST -> dependency-graph adapter.
//
// `extract_deps()` walks a parser AST and gathers two pieces of information
// the recalc engine needs at registration time:
//
//   1. The set of *cell* references the formula reads, expressed as
//      `CellNodeId`s. Plain `NodeKind::Ref` nodes contribute one cell each;
//      small `NodeKind::RangeOp` rectangles are flattened into the cells they
//      cover. Whole-column / whole-row references, and any rectangle above
//      `kMaxMaterializedDependencyCells`, are represented as compact
//      rectangles instead, so mutations within their bounds dirty the formula
//      without either enumerating the area or turning the formula into an
//      Excel-volatile one. Sheet qualifiers are resolved against the bound
//      `Workbook` so cross-sheet references land on the correct `sheet_id`.
//   2. Whether any function call in the tree is one of the nine Excel
//      Volatile functions (per `VolatileTracker::is_volatile_function`),
//      and separately whether any of them resolves its reads at
//      evaluation time (per `is_dynamic_reference_function`). The recalc
//      engine uses the first to add the formula's cell to the
//      always-dirty seed set and the second to decide whether the cell
//      may be evaluated concurrently with its layer.
//
// Defined-name (`NameRef`) nodes are resolved against the workbook's
// `defined_names()` list: a sheet-scoped definition matching the current
// sheet wins, otherwise the workbook-scoped definition (case-insensitive
// match) is used. The matched formula is parsed in a transient arena and
// walked recursively, so cells, ranges, volatility, and nested defined
// names all surface naturally. Cycles (`Loop = Loop + 1`) are detected by
// a name stack and broken silently.
//
// Structured (table) references (`Table[Col]`, `Table[#All]`,
// `Table[[#Headers],[ColA]:[ColB]]`, ...) are resolved at extract time
// against the workbook's `tables()` metadata: the bracket payload is
// parsed by the same machinery the evaluator uses, the resulting
// rectangle is pinned to the table's owning sheet, and the cells inside
// are emitted as direct deps. The implicit-intersection shorthand
// `Table[@Col]` is skipped statically because the row depends on the
// formula cell's row context that this layer does not know — the
// evaluator surfaces the actual dep when the implicit intersection
// resolves at eval time. Table-resize events must trigger a dep
// re-extraction at the recalc-engine layer; this layer does not maintain
// a live "table-shape" dep, matching how `RangeOp` is treated.
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
/// A compact, inclusive rectangle dependency. It is used for whole-column /
/// whole-row references and for any bounded rectangle wider than
/// `kMaxMaterializedDependencyCells`, whose individual cells must not be
/// materialised in the graph. `dependent` is added by the recalc engine when
/// it registers a formula; extraction only supplies the referenced rectangle.
struct CellRangeDependency {
  std::uint16_t sheet_id = 0;
  std::uint32_t row_first = 0;
  std::uint32_t row_last = 0;
  std::uint32_t col_first = 0;
  std::uint32_t col_last = 0;

  bool contains(CellNodeId cell) const noexcept {
    return cell.sheet_id == sheet_id && cell.row >= row_first && cell.row <= row_last && cell.col >= col_first &&
           cell.col <= col_last;
  }

  friend bool operator==(const CellRangeDependency& lhs, const CellRangeDependency& rhs) noexcept {
    return lhs.sheet_id == rhs.sheet_id && lhs.row_first == rhs.row_first && lhs.row_last == rhs.row_last &&
           lhs.col_first == rhs.col_first && lhs.col_last == rhs.col_last;
  }
  friend bool operator!=(const CellRangeDependency& lhs, const CellRangeDependency& rhs) noexcept {
    return !(lhs == rhs);
  }
};

/// Hash for `CellRangeDependency`, so identical rectangles authored by many
/// formulas (a lookup dragged down a column) can be interned into one entry.
/// Mixes the five coordinates with the same multiply-and-add strategy
/// `CellNodeIdHash` uses; the constant is truncated to `size_t` so the WASM
/// 32-bit build stays warning-clean alongside the 64-bit native build.
struct CellRangeDependencyHash {
  std::size_t operator()(const CellRangeDependency& range) const noexcept {
    constexpr std::size_t kMix = static_cast<std::size_t>(0x9e3779b97f4a7c15ULL);
    std::size_t hash = static_cast<std::size_t>(range.sheet_id);
    hash = (hash * kMix) + static_cast<std::size_t>(range.row_first);
    hash = (hash * kMix) + static_cast<std::size_t>(range.row_last);
    hash = (hash * kMix) + static_cast<std::size_t>(range.col_first);
    hash = (hash * kMix) + static_cast<std::size_t>(range.col_last);
    return hash;
  }
};

/// Normalized inclusive workbook-order sheet span read by a Ref3D node.
/// Endpoint names are resolved before this metadata is emitted, so callers
/// can match a structural edit against the span without reparsing formula
/// text. `sheet_first <= sheet_last` even when the authored endpoint order
/// is reversed.
struct ThreeDSheetSpanDependency {
  std::uint16_t sheet_first = 0;
  std::uint16_t sheet_last = 0;

  friend bool operator==(const ThreeDSheetSpanDependency& lhs, const ThreeDSheetSpanDependency& rhs) noexcept {
    return lhs.sheet_first == rhs.sheet_first && lhs.sheet_last == rhs.sheet_last;
  }
};

/// `is_volatile` is true only when any function-call node in the AST names
/// one of the nine Excel Volatile functions. Full-row / full-column refs use
/// `range_deps` instead.
struct ExtractedDeps {
  std::vector<CellNodeId> cell_deps;
  std::vector<CellRangeDependency> range_deps;
  /// Deduplicated normalized spans contributed by direct and expanded Ref3D
  /// nodes, including those reached through defined names / named Lambdas.
  std::vector<ThreeDSheetSpanDependency> three_d_spans;
  bool is_volatile = false;
  /// True when a volatile call in the AST resolves the cells it reads at
  /// evaluation time, so `cell_deps` / `range_deps` cannot describe that
  /// read. Implies `is_volatile`.
  bool has_dynamic_reference = false;
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
/// `RangeOp` rectangles are flattened cell-by-cell only while they cover at
/// most `kMaxMaterializedDependencyCells` cells. A wider rectangle — and any
/// rectangle with a whole-column / whole-row endpoint — is retained in
/// `range_deps` without enumerating its cells, so registration memory is
/// bounded by the formula text rather than by the referenced area. The same
/// ceiling applies to structured (table) references and to the per-sheet
/// rectangle of a 3-D reference. Endpoints that are not plain `Ref` nodes
/// (e.g. OFFSET / INDIRECT call results) are ignored; dynamic ranges are out
/// of scope for static dependency analysis.
ExtractedDeps extract_deps(const parser::AstNode& node, std::uint16_t current_sheet_id, const Workbook& workbook);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_DEP_EXTRACTOR_H_
