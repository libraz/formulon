//
// C ABI - trace precedents / dependents (BFS over the recalc dep graph).
//
// Bridges `RecalcEngine::dep_graph()` over the stable C ABI. The opaque
// `fm_cell_nodes_t` owns the BFS-expanded result so the index accessors
// can return data without re-walking the graph. Depth is capped at 32
// to keep cyclic graphs from blowing up the queue.

#include <cstddef>
#include <cstdint>
#include <memory>
#include <new>
#include <string>
#include <unordered_set>
#include <utility>
#include <vector>

#include "c_api/formulon_c.h"
#include "c_api/parts/common.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_sheet_u32;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::set_binding_error;

struct fm_cell_nodes {
  std::vector<formulon::eval::CellNodeId> nodes;
};

namespace {

constexpr std::uint32_t kMaxTraceDepth = 32U;

// Effective depth: 0 / 1 -> 1-step (direct neighbors only); larger
// values are capped at `kMaxTraceDepth` to keep BFS bounded.
std::uint32_t effective_depth(std::uint32_t depth) {
  if (depth <= 1) {
    return 1U;
  }
  return depth > kMaxTraceDepth ? kMaxTraceDepth : depth;
}

// BFS from `seed` using `next_neighbors(node) -> vector<CellNodeId>`
// up to `depth` hops. The seed itself is excluded from the result. The
// result preserves first-encountered order, which matches the
// dependency-graph adjacency list order at depth 1 and stays
// deterministic for deeper traversals.
template <typename NextFn>
std::vector<formulon::eval::CellNodeId> bfs_collect(formulon::eval::CellNodeId seed, std::uint32_t depth,
                                                    NextFn&& next_neighbors) {
  std::vector<formulon::eval::CellNodeId> ordered;
  std::unordered_set<formulon::eval::CellNodeId, formulon::eval::CellNodeIdHash> seen;
  std::vector<formulon::eval::CellNodeId> frontier;
  std::vector<formulon::eval::CellNodeId> next_frontier;
  seen.insert(seed);
  frontier.push_back(seed);
  for (std::uint32_t hop = 0; hop < depth; ++hop) {
    next_frontier.clear();
    for (formulon::eval::CellNodeId node : frontier) {
      for (formulon::eval::CellNodeId nbr : next_neighbors(node)) {
        if (seen.insert(nbr).second) {
          ordered.push_back(nbr);
          next_frontier.push_back(nbr);
        }
      }
    }
    if (next_frontier.empty()) {
      break;
    }
    frontier.swap(next_frontier);
  }
  return ordered;
}

template <bool kPrecedents>
fm_status_t trace_impl(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row, std::uint32_t col,
                       std::uint32_t depth, fm_cell_nodes_t** out) {
  clear_last_error();
  constexpr const char* fn_name = kPrecedents ? "fm_workbook_precedents" : "fm_workbook_dependents";
  if (out == nullptr) {
    return set_binding_error(
        formulon::FormulonErrorCode::kBindingNullPointer,
        kPrecedents ? "fm_workbook_precedents: NULL argument" : "fm_workbook_dependents: NULL argument");
  }
  *out = nullptr;
  if (auto rc = check_sheet_u32(wb, sheet, fn_name); rc != 0) {
    return rc;
  }
  const auto& workbook = wb->workbook();
  const auto& engine = workbook.recalc_engine();
  const auto& graph = engine.dep_graph();
  formulon::eval::CellNodeId seed{static_cast<std::uint16_t>(sheet), row, col};
  // A reference wide enough to be registered as a compact rectangle owns no
  // per-cell graph edge, so the raw adjacency lists alone would hide
  // `=SUM(A1:A5000)` from both trace directions. Fold the rectangle's
  // content-clipped expansion into the neighbour set.
  auto neighbors = [&](formulon::eval::CellNodeId node) {
    if constexpr (kPrecedents) {
      auto nodes = graph.dependencies_of(node);
      const auto compact = engine.compact_range_precedents_of(node, workbook);
      nodes.insert(nodes.end(), compact.begin(), compact.end());
      return nodes;
    } else {
      auto nodes = graph.dependents_of(node);
      const auto compact = engine.compact_range_dependents_of(node);
      nodes.insert(nodes.end(), compact.begin(), compact.end());
      return nodes;
    }
  };
  auto nodes = bfs_collect(seed, effective_depth(depth), neighbors);

  auto handle = std::unique_ptr<fm_cell_nodes_t>(new fm_cell_nodes_t{});
  handle->nodes = std::move(nodes);
  *out = handle.release();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_workbook_precedents(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  return trace_impl<true>(wb, sheet, row, col, depth, out);
}

extern "C" fm_status_t fm_workbook_dependents(const fm_workbook_t* wb, std::uint32_t sheet, std::uint32_t row,
                                              std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  return trace_impl<false>(wb, sheet, row, col, depth, out);
}

extern "C" void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes) {
  delete nodes;
}

extern "C" size_t fm_cell_nodes_count(const fm_cell_nodes_t* nodes) {
  if (nodes == nullptr) {
    return 0;
  }
  return nodes->nodes.size();
}

extern "C" fm_status_t fm_cell_nodes_at(const fm_cell_nodes_t* nodes, std::size_t idx, fm_cell_node_t* out) {
  clear_last_error();
  if (nodes == nullptr || out == nullptr) {
    return set_binding_error(formulon::FormulonErrorCode::kBindingNullPointer, "fm_cell_nodes_at: NULL argument");
  }
  if (idx >= nodes->nodes.size()) {
    return set_binding_error(formulon::FormulonErrorCode::kInvalidArgument, "fm_cell_nodes_at: idx out of range",
                             "idx=" + std::to_string(idx) + " count=" + std::to_string(nodes->nodes.size()));
  }
  out->sheet = static_cast<std::uint32_t>(nodes->nodes[idx].sheet_id);
  out->row = nodes->nodes[idx].row;
  out->col = nodes->nodes[idx].col;
  return 0;
}
