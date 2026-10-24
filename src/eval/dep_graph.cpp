//
// Implementation of `DepGraph`. See `dep_graph.h` for the public contract.

#include "eval/dep_graph.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <unordered_map>
#include <utility>
#include <vector>

#include "utils/index_sort.h"

namespace formulon::eval {
namespace {

// Erases the first occurrence of `target` from `vec`; no-op when absent.
void erase_first(std::vector<CellNodeId>& vec, CellNodeId target) {
  auto pos = std::find(vec.begin(), vec.end(), target);
  if (pos != vec.end()) {
    vec.erase(pos);
  }
}

}  // namespace

namespace {

constexpr std::uint8_t source_bit(DepGraph::DependencySource source) noexcept {
  return static_cast<std::uint8_t>(source);
}

}  // namespace

void DepGraph::add_dependency(CellNodeId dependent, CellNodeId dependency) {
  add_dependency_source(dependent, dependency, DependencySource::kAuthored);
}

void DepGraph::add_dependency_source(CellNodeId dependent, CellNodeId dependency, DependencySource source) {
  const Edge edge{dependent, dependency};
  const std::uint8_t bit = source_bit(source);
  auto existing = edge_sources_.find(edge);
  if (existing != edge_sources_.end()) {
    if ((existing->second & bit) == 0U && source == DependencySource::kSpillFootprint) {
      spill_footprint_edges_.insert(edge);
    }
    if ((existing->second & bit) == 0U && source == DependencySource::kAuthored) {
      ++authored_edge_count_;
    }
    if ((existing->second & bit) == 0U && source == DependencySource::kDynamicReference) {
      dynamic_by_dependent_[dependent].push_back(dependency);
      ++dynamic_edge_count_;
    }
    existing->second = static_cast<std::uint8_t>(existing->second | bit);
    return;
  }

  const bool dependent_was_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  const bool dependency_was_known = forward_.count(dependency) != 0U || reverse_.count(dependency) != 0U;
  edge_sources_.emplace(edge, bit);
  if (source == DependencySource::kSpillFootprint) {
    spill_footprint_edges_.insert(edge);
  } else if (source == DependencySource::kAuthored) {
    ++authored_edge_count_;
  } else if (source == DependencySource::kDynamicReference) {
    dynamic_by_dependent_[dependent].push_back(dependency);
    ++dynamic_edge_count_;
  }
  forward_[dependent].push_back(dependency);
  reverse_[dependency].push_back(dependent);
  if (!dependent_was_known && !is_range_node(dependent)) {
    ++node_count_;
  }
  if (!dependency_was_known && dependent != dependency && !is_range_node(dependency)) {
    ++node_count_;
  }
}

bool DepGraph::remove_dependency_source(CellNodeId dependent, CellNodeId dependency, DependencySource source) {
  const Edge edge{dependent, dependency};
  const std::uint8_t bit = source_bit(source);
  auto existing = edge_sources_.find(edge);
  if (existing == edge_sources_.end() || (existing->second & bit) == 0U) {
    return false;
  }
  if (source == DependencySource::kSpillFootprint) {
    spill_footprint_edges_.erase(edge);
  } else if (source == DependencySource::kAuthored) {
    --authored_edge_count_;
  } else if (source == DependencySource::kDynamicReference) {
    auto dyn = dynamic_by_dependent_.find(dependent);
    if (dyn != dynamic_by_dependent_.end()) {
      erase_first(dyn->second, dependency);
      if (dyn->second.empty()) {
        dynamic_by_dependent_.erase(dyn);
      }
    }
    --dynamic_edge_count_;
  }
  existing->second = static_cast<std::uint8_t>(existing->second & static_cast<std::uint8_t>(~bit));
  if (existing->second != 0U) {
    return false;
  }
  edge_sources_.erase(existing);

  const bool dependent_was_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  const bool dependency_was_known = forward_.count(dependency) != 0U || reverse_.count(dependency) != 0U;
  auto fwd = forward_.find(dependent);
  if (fwd != forward_.end()) {
    erase_first(fwd->second, dependency);
    if (fwd->second.empty()) {
      forward_.erase(fwd);
    }
  }
  auto rev = reverse_.find(dependency);
  if (rev != reverse_.end()) {
    erase_first(rev->second, dependent);
    if (rev->second.empty()) {
      reverse_.erase(rev);
    }
  }

  const bool dependent_is_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  const bool dependency_is_known = forward_.count(dependency) != 0U || reverse_.count(dependency) != 0U;
  if (dependent_was_known && !dependent_is_known && !is_range_node(dependent)) {
    --node_count_;
  }
  if (dependency != dependent && dependency_was_known && !dependency_is_known && !is_range_node(dependency)) {
    --node_count_;
  }
  return true;
}

DepGraph::DependencyDelta DepGraph::replace_dependencies(DependencySource source, const std::vector<Edge>& desired) {
  DependencyDelta delta;
  const std::uint8_t bit = source_bit(source);
  const std::size_t current_count = source_edge_count(source);
  if (current_count == 0U && desired.empty()) {
    return delta;
  }
  std::unordered_set<Edge, EdgeHash> wanted;
  wanted.reserve(desired.size());
  for (const Edge& edge : desired) {
    wanted.insert(edge);
  }

  std::vector<Edge> stale;
  if (source == DependencySource::kSpillFootprint) {
    stale.reserve(spill_footprint_edges_.size());
    for (const Edge& edge : spill_footprint_edges_) {
      if (wanted.count(edge) == 0U) {
        stale.push_back(edge);
      }
    }
  } else {
    stale.reserve(edge_sources_.size());
    for (const auto& [edge, mask] : edge_sources_) {
      if ((mask & bit) != 0U && wanted.count(edge) == 0U) {
        stale.push_back(edge);
      }
    }
  }
  for (const Edge& edge : stale) {
    // Report the removed source even when another source retains the adjacency.
    (void)remove_dependency_source(edge.first, edge.second, source);
    delta.removed.push_back(edge);
  }

  for (const Edge& edge : wanted) {
    if (!has_dependency_source(edge.first, edge.second, source)) {
      delta.added.push_back(edge);
      add_dependency_source(edge.first, edge.second, source);
    }
  }
  return delta;
}

bool DepGraph::has_dependency_source(CellNodeId dependent, CellNodeId dependency,
                                     DependencySource source) const noexcept {
  const auto pos = edge_sources_.find(Edge{dependent, dependency});
  return pos != edge_sources_.end() && (pos->second & source_bit(source)) != 0U;
}

std::size_t DepGraph::source_edge_count(DependencySource source) const noexcept {
  if (source == DependencySource::kSpillFootprint) {
    return spill_footprint_edges_.size();
  }
  if (source == DependencySource::kDynamicReference) {
    return dynamic_edge_count_;
  }
  return source == DependencySource::kAuthored ? authored_edge_count_ : 0U;
}

bool DepGraph::add_dynamic_dependency(CellNodeId dependent, CellNodeId dependency) {
  if (has_dependency_source(dependent, dependency, DependencySource::kDynamicReference)) {
    return false;
  }
  add_dependency_source(dependent, dependency, DependencySource::kDynamicReference);
  return true;
}

void DepGraph::clear_dynamic_dependencies_of(CellNodeId dependent) {
  const auto pos = dynamic_by_dependent_.find(dependent);
  if (pos == dynamic_by_dependent_.end()) {
    return;
  }
  const bool dependent_was_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  std::vector<CellNodeId> dependencies = std::move(pos->second);
  dynamic_by_dependent_.erase(pos);
  clear_outgoing_sources(dependent, dependencies, source_bit(DependencySource::kDynamicReference), dependent_was_known);
}

bool DepGraph::is_dynamic_only_edge(CellNodeId dependent, CellNodeId dependency) const noexcept {
  const auto pos = edge_sources_.find(Edge{dependent, dependency});
  return pos != edge_sources_.end() && pos->second == source_bit(DependencySource::kDynamicReference);
}

void DepGraph::remove_all_sources(CellNodeId dependent, CellNodeId dependency, std::uint8_t mask) {
  for (const DependencySource source :
       {DependencySource::kAuthored, DependencySource::kSpillFootprint, DependencySource::kDynamicReference}) {
    if ((mask & source_bit(source)) != 0U) {
      remove_dependency_source(dependent, dependency, source);
    }
  }
}

void DepGraph::remove_edge_sources(CellNodeId dependent, CellNodeId dependency) {
  const auto pos = edge_sources_.find(Edge{dependent, dependency});
  if (pos != edge_sources_.end()) {
    remove_all_sources(dependent, dependency, pos->second);
  }
}

void DepGraph::clear_outgoing_sources(CellNodeId dependent, const std::vector<CellNodeId>& dependencies,
                                      std::uint8_t source_mask, bool dependent_was_known) {
  for (CellNodeId dependency : dependencies) {
    const auto edge = edge_sources_.find(Edge{dependent, dependency});
    if (edge == edge_sources_.end()) {
      continue;
    }
    const std::uint8_t removed_mask = static_cast<std::uint8_t>(edge->second & source_mask);
    if (removed_mask == 0U) {
      continue;
    }
    if ((removed_mask & source_bit(DependencySource::kAuthored)) != 0U) {
      --authored_edge_count_;
    }
    if ((removed_mask & source_bit(DependencySource::kSpillFootprint)) != 0U) {
      spill_footprint_edges_.erase(Edge{dependent, dependency});
    }
    if ((removed_mask & source_bit(DependencySource::kDynamicReference)) != 0U) {
      --dynamic_edge_count_;
    }

    edge->second = static_cast<std::uint8_t>(edge->second & static_cast<std::uint8_t>(~removed_mask));
    if (edge->second != 0U) {
      continue;
    }
    edge_sources_.erase(edge);

    // Targets are distinct, so each reverse bucket is edited once; the source list is compacted below.
    const auto reverse = reverse_.find(dependency);
    if (reverse != reverse_.end()) {
      erase_first(reverse->second, dependent);
      if (reverse->second.empty()) {
        reverse_.erase(reverse);
      }
    }

    if (dependency != dependent && !is_range_node(dependency) && forward_.count(dependency) == 0U &&
        reverse_.count(dependency) == 0U) {
      --node_count_;
    }
  }

  const auto forward = forward_.find(dependent);
  if (forward != forward_.end()) {
    std::vector<CellNodeId>& adjacency = forward->second;
    const auto retained_end = std::remove_if(adjacency.begin(), adjacency.end(), [&](CellNodeId dependency) {
      return edge_sources_.find(Edge{dependent, dependency}) == edge_sources_.end();
    });
    adjacency.erase(retained_end, adjacency.end());
    if (adjacency.empty()) {
      forward_.erase(forward);
    }
  }

  if (dependent_was_known && !is_range_node(dependent) && forward_.count(dependent) == 0U &&
      reverse_.count(dependent) == 0U) {
    --node_count_;
  }
}

void DepGraph::clear_dependencies_of(CellNodeId dependent) {
  auto fwd = forward_.find(dependent);
  if (fwd == forward_.end()) {
    return;
  }
  const bool dependent_was_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  std::vector<CellNodeId> dependencies = std::move(fwd->second);
  forward_.erase(fwd);
  dynamic_by_dependent_.erase(dependent);
  constexpr std::uint8_t kAllSourceBits = static_cast<std::uint8_t>(source_bit(DependencySource::kAuthored) |
                                                                    source_bit(DependencySource::kSpillFootprint) |
                                                                    source_bit(DependencySource::kDynamicReference));
  clear_outgoing_sources(dependent, dependencies, kAllSourceBits, dependent_was_known);
}

void DepGraph::remove_node(CellNodeId node) {
  // Detach both buckets whole so no per-edge removal shifts a high-degree list.
  clear_dependencies_of(node);
  auto rev = reverse_.find(node);
  if (rev == reverse_.end()) {
    return;
  }
  const bool node_was_known = forward_.count(node) != 0U || reverse_.count(node) != 0U;
  std::vector<CellNodeId> dependents = std::move(rev->second);
  reverse_.erase(rev);
  for (CellNodeId dependent : dependents) {
    remove_edge_sources(dependent, node);
  }
  if (node_was_known && !is_range_node(node)) {
    --node_count_;
  }
}

std::vector<CellNodeId> DepGraph::dependents_of(CellNodeId node) const {
  return dependents_of_ref(node);
}

std::vector<CellNodeId> DepGraph::dependencies_of(CellNodeId node) const {
  return dependencies_of_ref(node);
}

namespace {

template <typename Map>
const std::vector<CellNodeId>& adjacency_or_empty(const Map& adjacency, CellNodeId node) noexcept {
  static const std::vector<CellNodeId> kEmpty;
  const auto pos = adjacency.find(node);
  return pos == adjacency.end() ? kEmpty : pos->second;
}

}  // namespace

const std::vector<CellNodeId>& DepGraph::dependencies_of_ref(CellNodeId node) const noexcept {
  return adjacency_or_empty(forward_, node);
}

const std::vector<CellNodeId>& DepGraph::dependents_of_ref(CellNodeId node) const noexcept {
  return adjacency_or_empty(reverse_, node);
}

bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) noexcept {
  if (component.size() != 1U) {
    return component.size() > 1U;
  }
  const CellNodeId only = component.front();
  const std::vector<CellNodeId>& dependencies = graph.dependencies_of_ref(only);
  return std::find(dependencies.begin(), dependencies.end(), only) != dependencies.end();
}

bool closes_through_dynamic_edge(const std::vector<CellNodeId>& component, const DepGraph& graph) {
  // Dynamic-only edges are rare, so membership is only searched for them.
  for (const CellNodeId member : component) {
    for (const CellNodeId dependency : graph.dependencies_of_ref(member)) {
      if (graph.is_dynamic_only_edge(member, dependency) &&
          std::find(component.begin(), component.end(), dependency) != component.end()) {
        return true;
      }
    }
  }
  return false;
}

std::vector<CellNodeId> cells_of_component(const std::vector<CellNodeId>& component) {
  std::vector<CellNodeId> cells;
  cells.reserve(component.size());
  for (const CellNodeId node : component) {
    if (!is_range_node(node)) {
      cells.push_back(node);
    }
  }
  return cells;
}

bool DepGraph::rect_cell_depends_on(CellNodeId node, std::uint32_t row_first, std::uint32_t col_first,
                                    std::uint32_t row_last, std::uint32_t col_last) const {
  std::unordered_map<CellNodeId, char, CellNodeIdHash> seen;
  std::vector<CellNodeId> queue{node};
  seen.emplace(node, 1);
  for (std::size_t head = 0; head < queue.size(); ++head) {
    for (const CellNodeId dependent : dependents_of_ref(queue[head])) {
      if (!seen.emplace(dependent, 1).second) {
        continue;
      }
      if (dependent.sheet_id == node.sheet_id && dependent.row >= row_first && dependent.row <= row_last &&
          dependent.col >= col_first && dependent.col <= col_last) {
        return true;
      }
      queue.push_back(dependent);
    }
  }
  return false;
}

std::vector<std::vector<CellNodeId>> DepGraph::tarjan_scc() const {
  return tarjan_scc_impl(nullptr);
}

std::vector<std::vector<CellNodeId>> DepGraph::tarjan_scc_subset(
    const std::unordered_set<CellNodeId, CellNodeIdHash>& nodes) const {
  return tarjan_scc_impl(&nodes);
}

std::vector<std::vector<CellNodeId>> DepGraph::tarjan_scc_impl(
    const std::unordered_set<CellNodeId, CellNodeIdHash>* nodes) const {
  // Iterative Tarjan to avoid blowing the WASM stack on deep dependency
  // chains. Algorithm summary:
  //
  //   index[node]   — DFS discovery index of `node`.
  //   lowlink[node] — smallest index reachable from `node` through tree
  //                   or back edges that are still on the explicit
  //                   `dfs_component_stack`.
  //   on_stack[node] — whether `node` is currently on
  //                    `dfs_component_stack`.
  //
  // When a node finishes (we have visited every neighbor) and
  // `lowlink == index`, pop everything down to that node from
  // `dfs_component_stack` to form an SCC.
  //
  // The classic recursive version uses C++ recursion for the DFS. Here we
  // use a `Frame` stack that records the "neighbor iterator position" so
  // we can resume after a recursive descent without using the C++ stack.

  // Collect every distinct node so the outer loop is well-defined.
  std::vector<CellNodeId> all_nodes;
  if (nodes != nullptr) {
    all_nodes.reserve(nodes->size());
    for (CellNodeId node : *nodes) {
      all_nodes.push_back(node);
    }
  } else {
    all_nodes.reserve(forward_.size() + reverse_.size());
    std::unordered_map<CellNodeId, char, CellNodeIdHash> seen;
    seen.reserve(forward_.size() + reverse_.size());
    for (const auto& entry : forward_) {
      if (seen.emplace(entry.first, 1).second) {
        all_nodes.push_back(entry.first);
      }
    }
    for (const auto& entry : reverse_) {
      if (seen.emplace(entry.first, 1).second) {
        all_nodes.push_back(entry.first);
      }
    }
  }

  std::unordered_map<CellNodeId, std::int64_t, CellNodeIdHash> index_of;
  std::unordered_map<CellNodeId, std::int64_t, CellNodeIdHash> lowlink_of;
  std::unordered_map<CellNodeId, char, CellNodeIdHash> on_stack;
  index_of.reserve(all_nodes.size());
  lowlink_of.reserve(all_nodes.size());
  on_stack.reserve(all_nodes.size());

  std::vector<CellNodeId> dfs_component_stack;  // Tarjan's S
  dfs_component_stack.reserve(all_nodes.size());

  // Frame for the explicit stack: (node, index-into-its-forward-list).
  struct Frame {
    CellNodeId node;
    std::size_t neighbor_idx;
  };
  std::vector<Frame> frames;
  frames.reserve(all_nodes.size());

  std::vector<std::vector<CellNodeId>> sccs;
  std::int64_t next_index = 0;

  // Empty adjacency (used when a node has no outgoing edges) — kept on
  // the stack so we can return a stable reference.
  static const std::vector<CellNodeId> kEmptyNeighbors;

  auto neighbors_of = [&](CellNodeId node) -> const std::vector<CellNodeId>& {
    auto pos = forward_.find(node);
    if (pos == forward_.end()) {
      return kEmptyNeighbors;
    }
    return pos->second;
  };

  for (CellNodeId root : all_nodes) {
    if (index_of.count(root) != 0U) {
      continue;
    }

    // Push the root frame.
    index_of[root] = next_index;
    lowlink_of[root] = next_index;
    ++next_index;
    dfs_component_stack.push_back(root);
    on_stack[root] = 1;
    frames.push_back(Frame{root, 0});

    while (!frames.empty()) {
      Frame& top = frames.back();
      const std::vector<CellNodeId>& adj = neighbors_of(top.node);
      if (top.neighbor_idx < adj.size()) {
        CellNodeId neighbor = adj[top.neighbor_idx++];
        if (nodes != nullptr && nodes->count(neighbor) == 0U) {
          continue;
        }
        if (index_of.count(neighbor) == 0U) {
          // Tree edge: descend.
          index_of[neighbor] = next_index;
          lowlink_of[neighbor] = next_index;
          ++next_index;
          dfs_component_stack.push_back(neighbor);
          on_stack[neighbor] = 1;
          frames.push_back(Frame{neighbor, 0});
        } else if (on_stack.count(neighbor) != 0U) {
          // Back edge into an ancestor that's still on the component
          // stack: tighten the current node's lowlink.
          std::int64_t& cur_low = lowlink_of[top.node];
          std::int64_t neighbor_index = index_of[neighbor];
          if (neighbor_index < cur_low) {
            cur_low = neighbor_index;
          }
        }
        // Cross / forward edges into a node that is no longer on the
        // component stack: ignore (its SCC has already been closed).
      } else {
        // All neighbors processed: maybe close the SCC rooted at this
        // node, then pop the frame and propagate the lowlink to the
        // parent.
        CellNodeId finished = top.node;
        std::int64_t finished_low = lowlink_of[finished];
        std::int64_t finished_index = index_of[finished];
        frames.pop_back();

        if (finished_low == finished_index) {
          // Root of an SCC — pop everything down to `finished`.
          std::vector<CellNodeId> component;
          while (true) {
            CellNodeId popped = dfs_component_stack.back();
            dfs_component_stack.pop_back();
            on_stack.erase(popped);
            component.push_back(popped);
            if (popped == finished) {
              break;
            }
          }
          // Pop order follows the DFS, which starts from whatever order the
          // hash containers enumerated the graph in — that order differs
          // between standard-library implementations. Members of a cyclic
          // component are committed in this order by the Gauss-Seidel
          // iterative solver, and a tautological cycle (`A1=B1+1`,
          // `B1=A1-1`) settles on a different pair depending on which
          // member goes first, so the order has to be pinned. Address order
          // also matches Excel's top-left-first calculation chain.
          std::sort(component.begin(), component.end(), CellNodeIdOrder{});
          sccs.push_back(std::move(component));
        }

        if (!frames.empty()) {
          CellNodeId parent = frames.back().node;
          std::int64_t& parent_low = lowlink_of[parent];
          if (finished_low < parent_low) {
            parent_low = finished_low;
          }
        }
      }
    }
  }

  // Tarjan emits SCCs in reverse-topological order, but which of several
  // valid orders depends on hash enumeration and adjacency insertion history.
  // The order decides which of two overlapping spill producers with no edge
  // between them commits first, so it is re-derived from the edge set and
  // the addresses alone.
  return canonical_component_order(std::move(sccs), nodes);
}

std::vector<std::vector<CellNodeId>> DepGraph::canonical_component_order(
    std::vector<std::vector<CellNodeId>> sccs, const std::unordered_set<CellNodeId, CellNodeIdHash>* nodes) const {
  const std::size_t count = sccs.size();
  // Reuses the value type of Tarjan's maps, so no new map is instantiated.
  std::unordered_map<CellNodeId, std::int64_t, CellNodeIdHash> component_of;
  for (std::size_t i = 0; i < count; ++i) {
    for (const CellNodeId member : sccs[i]) {
      component_of.emplace(member, static_cast<std::int64_t>(i));
    }
  }
  std::vector<std::vector<std::size_t>> successors(count);
  std::vector<std::size_t> in_degree(count, 0U);
  for (std::size_t i = 0; i < count; ++i) {
    for (const CellNodeId member : sccs[i]) {
      for (const CellNodeId dependency : dependencies_of_ref(member)) {
        if (nodes != nullptr && nodes->count(dependency) == 0U) {
          continue;
        }
        const auto it = component_of.find(dependency);
        if (it != component_of.end() && static_cast<std::size_t>(it->second) != i) {
          successors[static_cast<std::size_t>(it->second)].push_back(i);
          ++in_degree[i];
        }
      }
    }
  }

  // Kahn layers, as the parallel scheduler builds them, each taken in
  // address order: the serial walk then commits in the order the pooled
  // driver settles a layer in. Members are sorted, so a component's front is
  // its lowest address; a range node sorts after every cell.
  // Fronts of different components are different nodes, so the order is total.
  const auto earlier = [&sccs](std::uint32_t lhs, std::uint32_t rhs) {
    return CellNodeIdOrder{}(sccs[lhs].front(), sccs[rhs].front());
  };
  std::vector<std::uint32_t> frontier;
  for (std::size_t i = 0; i < count; ++i) {
    if (in_degree[i] == 0U) {
      frontier.push_back(static_cast<std::uint32_t>(i));
    }
  }
  std::vector<std::vector<CellNodeId>> ordered;
  ordered.reserve(count);
  std::vector<std::uint32_t> next_frontier;
  while (!frontier.empty()) {
    sort_index_order(frontier, make_index_less(earlier));
    next_frontier.clear();
    for (const std::uint32_t current : frontier) {
      for (const std::size_t successor : successors[current]) {
        if (--in_degree[successor] == 0U) {
          next_frontier.push_back(static_cast<std::uint32_t>(successor));
        }
      }
      ordered.push_back(std::move(sccs[current]));
    }
    frontier.swap(next_frontier);
  }
  return ordered;
}

std::vector<CellNodeId> DepGraph::topological_order() const {
  // `tarjan_scc()` emits leaves-first; flatten in that order so the
  // returned vector also satisfies "every node appears before any node
  // that depends on it".
  std::vector<std::vector<CellNodeId>> sccs = tarjan_scc();
  std::vector<CellNodeId> order;
  std::size_t total = 0;
  for (const auto& component : sccs) {
    total += component.size();
  }
  order.reserve(total);
  for (auto& component : sccs) {
    for (CellNodeId cell : component) {
      order.push_back(cell);
    }
  }
  return order;
}

}  // namespace formulon::eval
