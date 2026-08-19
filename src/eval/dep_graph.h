//
// Dependency graph used by the recalc engine. Tracks "this cell reads that
// cell" relationships across an entire workbook (i.e. the edges may span
// sheets) and answers the queries the recalc engine needs:
//
//   * "Which cells must be re-evaluated when X changes?"
//     -> `dependents_of(X)` — used to BFS-propagate dirtiness through a
//     cascade of formulas.
//   * "What does this formula read?" -> `dependencies_of(X)` — used when a
//     cell is rewritten (its old outgoing edges are dropped, its new ones
//     are added).
//   * "How do I evaluate the cells in an order that respects dependencies?"
//     -> `tarjan_scc()` returns strongly connected components in
//     **reverse-topological order** (leaves first, roots last) so the
//     evaluator can run SCCs from inputs to outputs and detect cycles.
//
// The graph is intentionally minimal: nodes are just `CellNodeId`
// (sheet+row+col triples) and edges are stored as a forward + reverse
// adjacency list. Range / DefinedName / TableColumn nodes will be modeled
// as separate node kinds in a follow-up bundle when the recalc engine
// starts wiring those in; today every node is a cell.
//
// The implementation uses an iterative Tarjan to stay safe on the WASM
// stack (Excel's 1,048,576-row maximum makes deep call chains plausible
// in pathological inputs).

#ifndef FORMULON_EVAL_DEP_GRAPH_H_
#define FORMULON_EVAL_DEP_GRAPH_H_

#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

namespace formulon::eval {

/// Workbook-wide cell coordinate used as a vertex in the dependency graph.
///
/// `sheet_id` selects the sheet (assigned by the workbook), `row` is in
/// `[0, 1048576)`, and `col` is in `[0, 16384)`. The triple uniquely
/// identifies a cell across the entire workbook.
struct CellNodeId {
  std::uint16_t sheet_id = 0;
  std::uint32_t row = 0;
  std::uint32_t col = 0;

  friend bool operator==(CellNodeId lhs, CellNodeId rhs) noexcept {
    return lhs.sheet_id == rhs.sheet_id && lhs.row == rhs.row && lhs.col == rhs.col;
  }
  friend bool operator!=(CellNodeId lhs, CellNodeId rhs) noexcept { return !(lhs == rhs); }
};

/// Orders cells by sheet, then row, then column — Excel's top-left-first
/// calculation order, and the canonical order every pinned cell sequence
/// in the recalc path is emitted in.
///
/// A named type rather than a lambda at each call site: the four places
/// that need this order would otherwise each instantiate their own copy of
/// the sort body.
struct CellNodeIdOrder {
  bool operator()(CellNodeId lhs, CellNodeId rhs) const noexcept {
    if (lhs.sheet_id != rhs.sheet_id) {
      return lhs.sheet_id < rhs.sheet_id;
    }
    if (lhs.row != rhs.row) {
      return lhs.row < rhs.row;
    }
    return lhs.col < rhs.col;
  }
};

/// Hash for `CellNodeId` suitable for `std::unordered_map` /
/// `std::unordered_set`.
///
/// Mirrors the `CellAddressHash` mixing strategy: Excel's coordinate ranges
/// (`row < 2^21`, `col < 2^14`, plus a small `sheet_id`) fit comfortably
/// inside 64 bits, so a multiply-and-add mix is collision-free in the
/// usable range. `noexcept` because the field accesses cannot throw.
struct CellNodeIdHash {
  // Mixing constants. `kSheetMix` is an odd 32-bit prime that decorrelates
  // adjacent sheet ids; `kRowMix` matches the multiplier used by
  // `CellAddressHash` so single-sheet workbooks hash exactly like the
  // per-sheet hash for the same `(row, col)`.
  static constexpr std::size_t kSheetMix = 1315423911U;
  static constexpr std::size_t kRowMix = 31U;

  std::size_t operator()(CellNodeId node) const noexcept {
    auto mix = static_cast<std::size_t>(node.sheet_id);
    mix = (mix * kSheetMix) + static_cast<std::size_t>(node.row);
    mix = (mix * kRowMix) + static_cast<std::size_t>(node.col);
    return mix;
  }
};

/// Sparse directed graph of cell-to-cell dependencies.
///
/// "Forward" edges go from a cell to the cells it reads (its
/// dependencies); "reverse" edges go from a cell to the cells that read it
/// (its dependents). Both adjacency lists are kept in sync by
/// `add_dependency` / `clear_dependencies_of` / `remove_node`.
///
/// This class is **not** thread-safe; the recalc engine drives mutation
/// from a single thread (locking will be layered on top by the scheduler
/// in a later phase).
class DepGraph {
 public:
  /// The source of a directed dependency edge. Authored edges come from a
  /// formula's static dependency extraction; spill-footprint edges are
  /// derived from committed dynamic-array geometry. Both sources share one
  /// union adjacency pair.
  enum class DependencySource : std::uint8_t {
    kAuthored = 1U,
    kSpillFootprint = 2U,
  };

  using Edge = std::pair<CellNodeId, CellNodeId>;

  /// Source-level changes made by `replace_dependencies`. These contain
  /// pairs whose source bit changed, even when another source keeps the
  /// pair in the union adjacency.
  struct DependencyDelta {
    std::vector<Edge> added;
    std::vector<Edge> removed;
  };

  DepGraph() = default;

  // Move-only.
  DepGraph(const DepGraph&) = delete;
  DepGraph& operator=(const DepGraph&) = delete;
  DepGraph(DepGraph&&) noexcept = default;
  DepGraph& operator=(DepGraph&&) noexcept = default;
  ~DepGraph() = default;

  /// Records that `dependent` reads `dependency` (i.e. an edge
  /// `dependent -> dependency` in the forward graph and the symmetric
  /// reverse edge). Idempotent: registering the same pair twice leaves the
  /// adjacency lists unchanged.
  void add_dependency(CellNodeId dependent, CellNodeId dependency);

  /// Replaces all edges owned by `source` with `desired`. The operation is
  /// idempotent and preserves a union adjacency when an edge has multiple
  /// source bits. `added` identifies newly acquired source ownership, which
  /// is useful for scheduling a targeted recalc wave.
  DependencyDelta replace_dependencies(DependencySource source, const std::vector<Edge>& desired);

  /// Returns whether the edge currently carries `source` provenance.
  bool has_dependency_source(CellNodeId dependent, CellNodeId dependency, DependencySource source) const noexcept;

  /// Returns the number of directed pairs carrying `source` provenance.
  /// This is maintained incrementally and is useful to skip a spill snapshot
  /// when no compact-range watcher or spill-owned edge exists.
  std::size_t source_edge_count(DependencySource source) const noexcept;

  /// Returns whether at least one edge carries `source` provenance.
  bool has_source_edges(DependencySource source) const noexcept { return source_edge_count(source) != 0U; }

  /// Drops every outgoing edge of `dependent`.
  ///
  /// Both the forward entry for `dependent` and the matching reverse
  /// entries on each former dependency are erased, so calling
  /// `dependents_of(d)` afterwards no longer reports `dependent` for any
  /// `d` that `dependent` previously read. The reverse edges *into*
  /// `dependent` (i.e. who reads `dependent`) are left intact — call
  /// `remove_node` if the node itself is going away.
  void clear_dependencies_of(CellNodeId dependent);

  /// Removes `node` from the graph entirely.
  ///
  /// Drops all of `node`'s outgoing and incoming edges and erases its
  /// adjacency list buckets. Safe to call on a node that has never been
  /// touched (no-op).
  void remove_node(CellNodeId node);

  /// Returns the cells that read `node` (i.e. `node`'s reverse-graph
  /// neighbors). Order is unspecified but stable for a given graph state.
  std::vector<CellNodeId> dependents_of(CellNodeId node) const;

  /// Returns the cells that `node` reads (i.e. `node`'s forward-graph
  /// neighbors). Order is unspecified but stable for a given graph state.
  std::vector<CellNodeId> dependencies_of(CellNodeId node) const;

  /// Borrowed version of `dependencies_of`, for recalc internals that only
  /// need to traverse an adjacency list. The returned reference stays valid
  /// until this graph is mutated.
  const std::vector<CellNodeId>& dependencies_of_ref(CellNodeId node) const noexcept;

  /// Computes the strongly connected components of the graph using an
  /// iterative Tarjan algorithm. Each inner vector is one SCC; SCCs are
  /// emitted in **reverse-topological order**: a component appears before
  /// any component that depends on it. So if `A` reads `B`, the SCC
  /// containing `B` is emitted before the SCC containing `A`. Within a
  /// component, the order of cells is unspecified.
  ///
  /// Singleton SCCs include cells that are isolated as well as cells in
  /// non-cyclic positions; a multi-cell SCC indicates a cycle (e.g. a
  /// circular reference). A self-loop on a cell `X` produces a singleton
  /// SCC `{X}` — callers that need to flag self-loops separately must
  /// check `dependencies_of(X)` for `X`.
  ///
  /// Complexity: O(V + E).
  std::vector<std::vector<CellNodeId>> tarjan_scc() const;

  /// Computes SCCs for the induced subgraph containing only `nodes`.
  /// Edges to cells outside the set are ignored. This keeps incremental
  /// recalc proportional to its dirty closure instead of the workbook-wide
  /// dependency graph. Every supplied node is emitted, including nodes with
  /// no recorded edges.
  std::vector<std::vector<CellNodeId>> tarjan_scc_subset(
      const std::unordered_set<CellNodeId, CellNodeIdHash>& nodes) const;

  /// Returns a topological ordering of the cells: every cell appears
  /// before any cell that depends on it. The ordering is derived from
  /// `tarjan_scc()` (each SCC contributes its cells in unspecified order).
  ///
  /// When the graph contains a cycle, the cycle's cells appear in the
  /// returned ordering at an arbitrary position relative to one another;
  /// callers that need an SCC-aware ordering should use `tarjan_scc()`
  /// directly.
  std::vector<CellNodeId> topological_order() const;

  /// Number of cells with at least one outgoing or incoming edge.
  /// O(1) — backed by a counter kept in sync by every mutator.
  std::size_t node_count() const noexcept { return node_count_; }

  /// Whether the graph has no recorded edges.
  bool empty() const noexcept { return forward_.empty() && reverse_.empty(); }

 private:
  using AdjacencyMap = std::unordered_map<CellNodeId, std::vector<CellNodeId>, CellNodeIdHash>;

  std::vector<std::vector<CellNodeId>> tarjan_scc_impl(
      const std::unordered_set<CellNodeId, CellNodeIdHash>* nodes) const;

  void add_dependency_source(CellNodeId dependent, CellNodeId dependency, DependencySource source);
  bool remove_dependency_source(CellNodeId dependent, CellNodeId dependency, DependencySource source);

  /// Hash for an ordered pair of `CellNodeId`s. Used to dedupe directed
  /// edges in O(1) on `add_dependency`. Combines the two component
  /// hashes with a boost-style mix so swapping the endpoints produces a
  /// different hash (the relation is asymmetric: `a -> b` is a
  /// different edge from `b -> a`). The golden-ratio constant is sized
  /// to `size_t` so the WASM 32-bit build stays warning-clean alongside
  /// the 64-bit native build.
  struct EdgeHash {
    std::size_t operator()(const std::pair<CellNodeId, CellNodeId>& edge) const noexcept {
      const std::size_t h1 = CellNodeIdHash{}(edge.first);
      const std::size_t h2 = CellNodeIdHash{}(edge.second);
      // Boost's `hash_combine` mix: stable, well-known, no platform
      // dependence on hash strength. The golden-ratio constant is
      // truncated to whatever `size_t` can hold (32 bits on WASM, 64
      // bits on native) — the mix is still avalanche-correct.
      constexpr std::size_t kGolden = static_cast<std::size_t>(0x9e3779b97f4a7c15ULL);
      return h1 ^ (h2 + kGolden + (h1 << 6) + (h1 >> 2));
    }
  };

  // Forward edges: `forward_[a]` is the list of cells `a` reads.
  AdjacencyMap forward_;
  // Reverse edges: `reverse_[b]` is the list of cells that read `b`.
  AdjacencyMap reverse_;
  // Parallel edge index for O(1) source lookup. Each entry is `(dependent,
  // dependency)` and a bit mask of `DependencySource` values. An adjacency
  // pair is removed only when its mask reaches zero.
  std::unordered_map<Edge, std::uint8_t, EdgeHash> edge_sources_;
  // Spill-footprint pairs have their own index so a replacement scans only
  // derived ownership, never authored/static edges. The set contains every
  // pair whose source mask carries kSpillFootprint, including pairs that
  // also carry kAuthored.
  std::unordered_set<Edge, EdgeHash> spill_footprint_edges_;
  // Authored ownership is counted separately so `source_edge_count` stays
  // O(1) for both source kinds; unlike spill ownership it does not need a
  // replacement iteration index.
  std::size_t authored_edge_count_ = 0U;
  // Distinct-node counter kept in sync by `add_dependency` /
  // `clear_dependencies_of` / `remove_node`. Backs the `noexcept` /
  // O(1) `node_count()` accessor; without it, that accessor would
  // need to allocate and walk both adjacency maps every call.
  std::size_t node_count_ = 0;
};

/// Returns whether an SCC is a real cycle: two or more nodes, or a
/// singleton whose sole cell has a self-dependency.
bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) noexcept;

}  // namespace formulon::eval

#endif  // FORMULON_EVAL_DEP_GRAPH_H_
