//
// Implementation of `DepGraph`. See `dep_graph.h` for the public contract.

#include "eval/dep_graph.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <unordered_map>
#include <utility>
#include <vector>

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

void DepGraph::add_dependency(CellNodeId dependent, CellNodeId dependency) {
  // O(1) dedup via the parallel edge index. Idempotent on a duplicate
  // edge so the surrounding recalc engine can re-record dependencies
  // freely without quadratic blow-up on dense formulas (`SUM(A1:Z99)`
  // would otherwise scan ~26*99 entries per duplicate insertion).
  if (!edges_.insert(std::make_pair(dependent, dependency)).second) {
    return;
  }
  // Record key existence BEFORE mutation so the counter sees a
  // monotone delta. A key counts as "present" if it lives in either
  // adjacency map; the union models the conceptual node set.
  const bool dependent_was_known = forward_.count(dependent) != 0U || reverse_.count(dependent) != 0U;
  const bool dependency_was_known = forward_.count(dependency) != 0U || reverse_.count(dependency) != 0U;
  forward_[dependent].push_back(dependency);
  reverse_[dependency].push_back(dependent);
  if (!dependent_was_known) {
    ++node_count_;
  }
  // Avoid double-counting the self-loop case where dependent ==
  // dependency: it is a single new node, not two.
  if (!dependency_was_known && dependent != dependency) {
    ++node_count_;
  }
}

void DepGraph::clear_dependencies_of(CellNodeId dependent) {
  auto fwd_pos = forward_.find(dependent);
  if (fwd_pos == forward_.end()) {
    return;
  }
  // Drop the matching reverse entry from each former dependency, and
  // erase that node's reverse bucket entirely if it becomes empty AND the
  // node itself has no outgoing edges either (i.e. it would otherwise be
  // an orphan). Track each dependency that becomes unreferenced so the
  // node-count counter can be adjusted exactly once at the end.
  for (CellNodeId dependency : fwd_pos->second) {
    edges_.erase(std::make_pair(dependent, dependency));
    auto rev_pos = reverse_.find(dependency);
    if (rev_pos == reverse_.end()) {
      continue;
    }
    erase_first(rev_pos->second, dependent);
    if (rev_pos->second.empty()) {
      reverse_.erase(rev_pos);
      // `dependency` is fully gone from both maps only when it also has
      // no outgoing edges. Skip the self-loop case (`dependency ==
      // dependent`) because the dependent's own node-count delta is
      // handled below after its forward bucket is erased.
      if (forward_.find(dependency) == forward_.end() && dependency != dependent) {
        --node_count_;
      }
    }
  }
  forward_.erase(fwd_pos);
  // `dependent` may still be referenced by `reverse_` (someone reads
  // it). Only decrement when it has truly left the node set.
  if (reverse_.find(dependent) == reverse_.end()) {
    --node_count_;
  }
}

void DepGraph::remove_node(CellNodeId node) {
  // Capture pre-state so the final node-count adjustment knows whether
  // `node` was ever in the graph.
  const bool node_was_known = forward_.count(node) != 0U || reverse_.count(node) != 0U;

  // Outgoing edges: drop `node` from each former dependency's reverse list.
  auto fwd_pos = forward_.find(node);
  if (fwd_pos != forward_.end()) {
    for (CellNodeId dependency : fwd_pos->second) {
      edges_.erase(std::make_pair(node, dependency));
      auto rev_pos = reverse_.find(dependency);
      if (rev_pos == reverse_.end()) {
        continue;
      }
      erase_first(rev_pos->second, node);
      if (rev_pos->second.empty() && forward_.find(dependency) == forward_.end()) {
        reverse_.erase(rev_pos);
        if (dependency != node) {
          --node_count_;
        }
      }
    }
    forward_.erase(fwd_pos);
  }

  // Incoming edges: drop `node` from each dependent's forward list.
  auto rev_pos = reverse_.find(node);
  if (rev_pos != reverse_.end()) {
    for (CellNodeId dependent : rev_pos->second) {
      edges_.erase(std::make_pair(dependent, node));
      auto fwd_other = forward_.find(dependent);
      if (fwd_other == forward_.end()) {
        continue;
      }
      erase_first(fwd_other->second, node);
      if (fwd_other->second.empty() && reverse_.find(dependent) == reverse_.end()) {
        forward_.erase(fwd_other);
        if (dependent != node) {
          --node_count_;
        }
      }
    }
    reverse_.erase(rev_pos);
  }

  // Final adjustment: `node` itself is now absent from both maps. The
  // per-edge erasures above only decrement when *other* nodes become
  // unreferenced, never `node` itself, so do that here.
  if (node_was_known) {
    --node_count_;
  }
}

std::vector<CellNodeId> DepGraph::dependents_of(CellNodeId node) const {
  auto pos = reverse_.find(node);
  if (pos == reverse_.end()) {
    return {};
  }
  return pos->second;
}

std::vector<CellNodeId> DepGraph::dependencies_of(CellNodeId node) const {
  return dependencies_of_ref(node);
}

const std::vector<CellNodeId>& DepGraph::dependencies_of_ref(CellNodeId node) const noexcept {
  static const std::vector<CellNodeId> kEmptyDependencies;
  auto pos = forward_.find(node);
  if (pos == forward_.end()) {
    return kEmptyDependencies;
  }
  return pos->second;
}

bool is_cyclic_component(const std::vector<CellNodeId>& component, const DepGraph& graph) noexcept {
  if (component.size() != 1U) {
    return component.size() > 1U;
  }
  const CellNodeId only = component.front();
  const std::vector<CellNodeId>& dependencies = graph.dependencies_of_ref(only);
  return std::find(dependencies.begin(), dependencies.end(), only) != dependencies.end();
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

  // Tarjan emits SCCs in reverse-topological order by construction:
  // deeper-first DFS post-order means leaves close before roots. Keep that
  // ordering as the documented contract.
  return sccs;
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
