//
// Unit tests for the workbook-wide cell dependency graph.

#include "eval/dep_graph.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <unordered_set>
#include <vector>

#include "gtest/gtest.h"

namespace formulon::eval {
namespace {

CellNodeId Make(std::uint16_t sheet, std::uint32_t row, std::uint32_t col) {
  return CellNodeId{sheet, row, col};
}

// Sorts a vector<CellNodeId> by (sheet, row, col) so order-insensitive
// comparisons are easy.
std::vector<CellNodeId> Sorted(std::vector<CellNodeId> v) {
  std::sort(v.begin(), v.end(), [](CellNodeId a, CellNodeId b) {
    if (a.sheet_id != b.sheet_id)
      return a.sheet_id < b.sheet_id;
    if (a.row != b.row)
      return a.row < b.row;
    return a.col < b.col;
  });
  return v;
}

// Position of `id` inside `order`, or -1 if not present.
int IndexOf(const std::vector<CellNodeId>& order, CellNodeId id) {
  for (std::size_t i = 0; i < order.size(); ++i) {
    if (order[i] == id) {
      return static_cast<int>(i);
    }
  }
  return -1;
}

// Index of the SCC containing `id` in `sccs`, or -1.
int SccIndexOf(const std::vector<std::vector<CellNodeId>>& sccs, CellNodeId id) {
  for (std::size_t i = 0; i < sccs.size(); ++i) {
    for (CellNodeId v : sccs[i]) {
      if (v == id)
        return static_cast<int>(i);
    }
  }
  return -1;
}

TEST(DepGraph, EmptyGraph) {
  DepGraph g;
  EXPECT_TRUE(g.empty());
  EXPECT_EQ(g.node_count(), 0u);
  EXPECT_TRUE(g.tarjan_scc().empty());
  EXPECT_TRUE(g.topological_order().empty());
  EXPECT_TRUE(g.dependents_of(Make(0, 0, 0)).empty());
  EXPECT_TRUE(g.dependencies_of(Make(0, 0, 0)).empty());
}

TEST(DepGraph, SingleEdge) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  g.add_dependency(a, b);

  EXPECT_FALSE(g.empty());
  EXPECT_EQ(g.node_count(), 2u);
  EXPECT_EQ(g.dependencies_of(a), std::vector<CellNodeId>{b});
  EXPECT_EQ(g.dependents_of(b), std::vector<CellNodeId>{a});
  EXPECT_TRUE(g.dependents_of(a).empty());
  EXPECT_TRUE(g.dependencies_of(b).empty());
  EXPECT_EQ(g.dependencies_of_ref(a), std::vector<CellNodeId>{b});
  EXPECT_TRUE(g.dependencies_of_ref(b).empty());
}

TEST(DepGraph, ClassifiesCyclicComponents) {
  DepGraph g;
  const CellNodeId a = Make(0, 0, 0);
  const CellNodeId b = Make(0, 0, 1);
  const CellNodeId c = Make(0, 0, 2);
  g.add_dependency(a, a);
  g.add_dependency(b, c);
  g.add_dependency(c, b);

  EXPECT_FALSE(is_cyclic_component({}, g));
  EXPECT_TRUE(is_cyclic_component({a}, g));
  EXPECT_FALSE(is_cyclic_component({b}, g));
  EXPECT_TRUE(is_cyclic_component({b, c}, g));
}

TEST(DepGraph, IdempotentAddDependency) {
  DepGraph g;
  CellNodeId a = Make(0, 1, 1);
  CellNodeId b = Make(0, 1, 2);
  g.add_dependency(a, b);
  g.add_dependency(a, b);
  g.add_dependency(a, b);

  // Both adjacency lists must contain exactly one entry.
  EXPECT_EQ(g.dependencies_of(a).size(), 1u);
  EXPECT_EQ(g.dependents_of(b).size(), 1u);
}

TEST(DepGraph, SourceProvenancePreservesStaticOverlap) {
  DepGraph g;
  const CellNodeId watcher = Make(0, 1, 0);
  const CellNodeId producer = Make(0, 2, 0);
  g.add_dependency(watcher, producer);

  const auto delta = g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, {{watcher, producer}});
  ASSERT_EQ(delta.added.size(), 1U);
  EXPECT_TRUE(g.has_dependency_source(watcher, producer, DepGraph::DependencySource::kAuthored));
  EXPECT_TRUE(g.has_dependency_source(watcher, producer, DepGraph::DependencySource::kSpillFootprint));
  EXPECT_EQ(g.dependencies_of(watcher), std::vector<CellNodeId>{producer});

  const auto removed = g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, {});
  ASSERT_EQ(removed.removed.size(), 1U);
  EXPECT_TRUE(g.has_dependency_source(watcher, producer, DepGraph::DependencySource::kAuthored));
  EXPECT_FALSE(g.has_dependency_source(watcher, producer, DepGraph::DependencySource::kSpillFootprint));
  EXPECT_EQ(g.dependencies_of(watcher), std::vector<CellNodeId>{producer});
  EXPECT_EQ(g.source_edge_count(DepGraph::DependencySource::kSpillFootprint), 0U);
  EXPECT_EQ(g.source_edge_count(DepGraph::DependencySource::kAuthored), 1U);
}

TEST(DepGraph, SourceReplacementIsIdempotentAndRemovesDerivedOnlyEdges) {
  DepGraph g;
  const CellNodeId watcher = Make(0, 1, 0);
  const CellNodeId first = Make(0, 2, 0);
  const CellNodeId second = Make(0, 3, 0);
  const std::vector<DepGraph::Edge> desired = {{watcher, first}, {watcher, second}};
  EXPECT_EQ(g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, desired).added.size(), 2U);
  EXPECT_TRUE(g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, desired).added.empty());
  const auto delta = g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, {{watcher, second}});
  ASSERT_EQ(delta.removed.size(), 1U);
  const auto dependencies = g.dependencies_of(watcher);
  EXPECT_NE(std::find(dependencies.begin(), dependencies.end(), second), dependencies.end());
  EXPECT_EQ(std::find(dependencies.begin(), dependencies.end(), first), dependencies.end());
}

TEST(DepGraph, ClearAndRemoveNodeDropEveryProvenanceSource) {
  DepGraph g;
  const CellNodeId watcher = Make(0, 1, 0);
  const CellNodeId producer = Make(0, 2, 0);
  g.add_dependency(watcher, producer);
  g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, {{watcher, producer}});
  g.clear_dependencies_of(watcher);
  EXPECT_TRUE(g.empty());
  EXPECT_EQ(g.node_count(), 0U);

  g.replace_dependencies(DepGraph::DependencySource::kSpillFootprint, {{watcher, producer}});
  g.remove_node(producer);
  EXPECT_TRUE(g.empty());
  EXPECT_EQ(g.node_count(), 0U);
}

TEST(DepGraph, CrossSheetEdge) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);  // Sheet1!A1
  CellNodeId b = Make(1, 0, 0);  // Sheet2!A1 (different sheet)
  g.add_dependency(a, b);

  EXPECT_EQ(g.dependencies_of(a), std::vector<CellNodeId>{b});
  EXPECT_EQ(g.dependents_of(b), std::vector<CellNodeId>{a});
  // The same row/col on a different sheet is *not* the same node.
  CellNodeId b_other_sheet = Make(2, 0, 0);
  EXPECT_TRUE(g.dependencies_of(b_other_sheet).empty());
  EXPECT_TRUE(g.dependents_of(b_other_sheet).empty());
}

TEST(DepGraph, ClearDependenciesOfDropsReverseEdges) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  g.add_dependency(a, b);
  g.add_dependency(a, c);
  // Sanity.
  EXPECT_EQ(Sorted(g.dependencies_of(a)), Sorted({b, c}));
  EXPECT_EQ(g.dependents_of(b), std::vector<CellNodeId>{a});
  EXPECT_EQ(g.dependents_of(c), std::vector<CellNodeId>{a});

  g.clear_dependencies_of(a);

  // a's outgoing edges are gone.
  EXPECT_TRUE(g.dependencies_of(a).empty());
  // The reverse edges into b and c that pointed at a must also be gone.
  EXPECT_TRUE(g.dependents_of(b).empty());
  EXPECT_TRUE(g.dependents_of(c).empty());
}

TEST(DepGraph, ClearDependenciesPreservesIncomingEdges) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  // c depends on a; a depends on b. Clearing a's deps should not drop the
  // edge c->a.
  g.add_dependency(a, b);
  g.add_dependency(c, a);

  g.clear_dependencies_of(a);

  EXPECT_TRUE(g.dependencies_of(a).empty());
  EXPECT_TRUE(g.dependents_of(b).empty());
  EXPECT_EQ(g.dependents_of(a), std::vector<CellNodeId>{c});
  EXPECT_EQ(g.dependencies_of(c), std::vector<CellNodeId>{a});
}

TEST(DepGraph, ClearSelfLoopRemovesTheNowEmptyReverseBucket) {
  DepGraph g;
  const CellNodeId a = Make(0, 0, 0);
  g.add_dependency(a, a);
  ASSERT_FALSE(g.empty());
  ASSERT_EQ(g.node_count(), 1u);

  g.clear_dependencies_of(a);

  EXPECT_TRUE(g.dependencies_of(a).empty());
  EXPECT_TRUE(g.dependents_of(a).empty());
  EXPECT_TRUE(g.empty());
  EXPECT_EQ(g.node_count(), 0u);
}

TEST(DepGraph, RemoveNodeDropsAllEdges) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  // a -> b, c -> a, c -> b
  g.add_dependency(a, b);
  g.add_dependency(c, a);
  g.add_dependency(c, b);

  g.remove_node(a);

  EXPECT_TRUE(g.dependencies_of(a).empty());
  EXPECT_TRUE(g.dependents_of(a).empty());
  // Edges that did not touch a must still be there.
  EXPECT_EQ(g.dependencies_of(c), std::vector<CellNodeId>{b});
  EXPECT_EQ(g.dependents_of(b), std::vector<CellNodeId>{c});
}

TEST(DepGraph, RemoveNodeOnUnknown) {
  DepGraph g;
  // Should be a no-op rather than blowing up.
  g.remove_node(Make(7, 42, 9));
  EXPECT_TRUE(g.empty());
}

TEST(DepGraph, TarjanLinearChainReverseTopological) {
  // a -> b -> c. Tarjan emits leaves first, so the SCC containing `c`
  // must come before the SCC containing `b`, which must come before the
  // SCC containing `a`.
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  g.add_dependency(a, b);
  g.add_dependency(b, c);

  auto sccs = g.tarjan_scc();
  ASSERT_EQ(sccs.size(), 3u);
  for (const auto& comp : sccs) {
    EXPECT_EQ(comp.size(), 1u);
  }
  int idx_a = SccIndexOf(sccs, a);
  int idx_b = SccIndexOf(sccs, b);
  int idx_c = SccIndexOf(sccs, c);
  ASSERT_GE(idx_a, 0);
  ASSERT_GE(idx_b, 0);
  ASSERT_GE(idx_c, 0);
  EXPECT_LT(idx_c, idx_b);
  EXPECT_LT(idx_b, idx_a);
}

TEST(DepGraph, TarjanSubsetUsesOnlySelectedDirtyClosure) {
  // a -> b -> c -> d. Selecting b/c must not allocate or traverse the
  // unrelated endpoints; the induced subgraph preserves b -> c ordering.
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  CellNodeId d = Make(0, 0, 3);
  g.add_dependency(a, b);
  g.add_dependency(b, c);
  g.add_dependency(c, d);

  const std::unordered_set<CellNodeId, CellNodeIdHash> selected = {b, c};
  const auto sccs = g.tarjan_scc_subset(selected);
  ASSERT_EQ(sccs.size(), 2U);
  EXPECT_EQ(SccIndexOf(sccs, a), -1);
  EXPECT_EQ(SccIndexOf(sccs, d), -1);
  EXPECT_LT(SccIndexOf(sccs, c), SccIndexOf(sccs, b));
}

TEST(DepGraph, TarjanSelfLoopIsSingleton) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  g.add_dependency(a, a);

  auto sccs = g.tarjan_scc();
  ASSERT_EQ(sccs.size(), 1u);
  ASSERT_EQ(sccs[0].size(), 1u);
  EXPECT_EQ(sccs[0][0], a);

  // The self-loop is observable through `dependencies_of` so callers can
  // distinguish a singleton with a self-edge from one without.
  EXPECT_EQ(g.dependencies_of(a), std::vector<CellNodeId>{a});
  EXPECT_EQ(g.dependents_of(a), std::vector<CellNodeId>{a});
}

TEST(DepGraph, TarjanSimpleCycleTwoNodes) {
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  g.add_dependency(a, b);
  g.add_dependency(b, a);

  auto sccs = g.tarjan_scc();
  ASSERT_EQ(sccs.size(), 1u);
  EXPECT_EQ(Sorted(sccs[0]), Sorted({a, b}));
}

TEST(DepGraph, TarjanLargerCycle) {
  // a -> b -> c -> a, and an extra acyclic edge d -> a so there is a
  // distinct second SCC at depth 0.
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  CellNodeId d = Make(0, 0, 3);
  g.add_dependency(a, b);
  g.add_dependency(b, c);
  g.add_dependency(c, a);
  g.add_dependency(d, a);

  auto sccs = g.tarjan_scc();
  ASSERT_EQ(sccs.size(), 2u);

  // Find the cycle SCC (3 members) and the singleton.
  std::vector<CellNodeId> cycle;
  std::vector<CellNodeId> singleton;
  for (const auto& comp : sccs) {
    if (comp.size() == 3u) {
      cycle = comp;
    } else if (comp.size() == 1u) {
      singleton = comp;
    }
  }
  EXPECT_EQ(Sorted(cycle), Sorted({a, b, c}));
  ASSERT_EQ(singleton.size(), 1u);
  EXPECT_EQ(singleton[0], d);

  // The singleton {d} reads into the cycle, so the cycle's SCC must come
  // first in reverse-topological order.
  int cycle_idx = SccIndexOf(sccs, a);
  int singleton_idx = SccIndexOf(sccs, d);
  ASSERT_GE(cycle_idx, 0);
  ASSERT_GE(singleton_idx, 0);
  EXPECT_LT(cycle_idx, singleton_idx);
}

TEST(DepGraph, TarjanTwoIndependentSccs) {
  // Two disjoint cycles. Order of the two SCCs relative to each other is
  // unspecified; we just check that there are two of size 2.
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(1, 0, 0);
  CellNodeId d = Make(1, 0, 1);
  g.add_dependency(a, b);
  g.add_dependency(b, a);
  g.add_dependency(c, d);
  g.add_dependency(d, c);

  auto sccs = g.tarjan_scc();
  ASSERT_EQ(sccs.size(), 2u);
  for (const auto& comp : sccs) {
    EXPECT_EQ(comp.size(), 2u);
  }
}

TEST(DepGraph, TopologicalOrderRespectsDependencies) {
  // a -> b -> c. In topo order, c (leaf) appears before b before a.
  DepGraph g;
  CellNodeId a = Make(0, 0, 0);
  CellNodeId b = Make(0, 0, 1);
  CellNodeId c = Make(0, 0, 2);
  g.add_dependency(a, b);
  g.add_dependency(b, c);

  auto order = g.topological_order();
  ASSERT_EQ(order.size(), 3u);
  EXPECT_LT(IndexOf(order, c), IndexOf(order, b));
  EXPECT_LT(IndexOf(order, b), IndexOf(order, a));
}

TEST(DepGraph, FanoutDependentsAllPresent) {
  // Many dependents of a single node — verifies dependents_of returns
  // exactly the registered set without duplicates.
  DepGraph g;
  CellNodeId hub = Make(0, 0, 0);
  constexpr std::uint32_t kFanout = 256;

  std::unordered_set<std::uint32_t> expected_cols;
  for (std::uint32_t i = 1; i <= kFanout; ++i) {
    g.add_dependency(Make(0, 1, i), hub);
    expected_cols.insert(i);
  }
  // Idempotent re-registration.
  for (std::uint32_t i = 1; i <= kFanout; ++i) {
    g.add_dependency(Make(0, 1, i), hub);
  }

  std::vector<CellNodeId> dependents = g.dependents_of(hub);
  EXPECT_EQ(dependents.size(), kFanout);

  std::unordered_set<std::uint32_t> got_cols;
  for (CellNodeId d : dependents) {
    EXPECT_EQ(d.sheet_id, 0u);
    EXPECT_EQ(d.row, 1u);
    got_cols.insert(d.col);
  }
  EXPECT_EQ(got_cols, expected_cols);
}

TEST(DepGraph, NodeIdEqualityAndHash) {
  CellNodeId a = Make(3, 4, 5);
  CellNodeId b = Make(3, 4, 5);
  CellNodeId c = Make(3, 4, 6);
  EXPECT_EQ(a, b);
  EXPECT_NE(a, c);
  CellNodeIdHash h;
  EXPECT_EQ(h(a), h(b));
  // Hash inequality across different ids is not strictly required, but
  // a sane mix should not collide on these tiny coordinates.
  EXPECT_NE(h(a), h(c));
}

}  // namespace
}  // namespace formulon::eval
