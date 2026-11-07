//
// C ABI - trace precedents / dependents (BFS over the recalc dep graph).
//
// Bridges `RecalcEngine::dep_graph()` over the stable C ABI. The opaque
// `fm_cell_nodes_t` owns the BFS-expanded result so the index accessors
// can return data without re-walking the graph. Depth is capped at 32
// to keep cyclic graphs from blowing up the queue.

#include <algorithm>
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
#include "c_api/parts/struct_v1_sizes.h"
#include "eval/dep_graph.h"
#include "eval/recalc_engine.h"
#include "utils/error.h"
#include "workbook.h"

using formulon::c_api::parts::check_cell_coord;
using formulon::c_api::parts::check_index;
using formulon::c_api::parts::check_record_size;
using formulon::c_api::parts::check_sheet_index;
using formulon::c_api::parts::clear_last_error;
using formulon::c_api::parts::ensure_calculated;
using formulon::c_api::parts::null_argument;

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

// Shared argument checks of the two trace entry points; `*out` is NULL on failure.
fm_status_t check_trace_args(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row, std::uint32_t col,
                             fm_cell_nodes_t** out) {
  clear_last_error();
  if (out == nullptr) {
    return null_argument("out");
  }
  *out = nullptr;
  if (auto rc = check_sheet_index(wb, sheet_index); rc != 0) {
    return rc;
  }
  return check_cell_coord(row, col);
}

// The direction is a runtime flag rather than a template parameter so the
// BFS body and its hash set are linked once for both entry points.
fm_status_t trace_impl(bool precedents, const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                       std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  const auto& workbook = wb->workbook();
  const auto& engine = workbook.recalc_engine();
  const auto& graph = engine.dep_graph();
  formulon::eval::CellNodeId seed{static_cast<std::uint16_t>(sheet_index), row, col};
  // A reference wide enough to be registered as a compact rectangle owns no
  // per-cell graph edge, so the raw adjacency lists alone would hide
  // `=SUM(A1:A5000)` from both trace directions. Fold the rectangle's
  // content-clipped expansion into the neighbour set in place of the
  // rectangle's virtual graph node, which is not a cell.
  auto neighbors = [&](formulon::eval::CellNodeId node) {
    auto nodes = precedents ? graph.dependencies_of(node) : graph.dependents_of(node);
    nodes.erase(std::remove_if(nodes.begin(), nodes.end(), formulon::eval::is_range_node), nodes.end());
    const auto compact =
        precedents ? engine.compact_range_precedents_of(node, workbook) : engine.compact_range_dependents_of(node);
    nodes.insert(nodes.end(), compact.begin(), compact.end());
    return nodes;
  };
  auto nodes = bfs_collect(seed, effective_depth(depth), neighbors);

  auto handle = std::unique_ptr<fm_cell_nodes_t>(new fm_cell_nodes_t{});
  handle->nodes = std::move(nodes);
  *out = handle.release();
  return 0;
}

}  // namespace

extern "C" fm_status_t fm_sheet_get_precedents(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                               std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  if (auto rc = check_trace_args(wb, sheet_index, row, col, out); rc != 0) {
    return rc;
  }
  // INDIRECT / OFFSET edges exist only once the formulas are calculated.
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  return trace_impl(true, wb, sheet_index, row, col, depth, out);
}

extern "C" fm_status_t fm_sheet_get_dependents(const fm_workbook_t* wb, std::size_t sheet_index, std::uint32_t row,
                                               std::uint32_t col, std::uint32_t depth, fm_cell_nodes_t** out) {
  if (auto rc = check_trace_args(wb, sheet_index, row, col, out); rc != 0) {
    return rc;
  }
  // INDIRECT / OFFSET edges exist only once the formulas are calculated.
  if (auto rc = ensure_calculated(wb); rc != 0) {
    return rc;
  }
  return trace_impl(false, wb, sheet_index, row, col, depth, out);
}

extern "C" void fm_cell_nodes_destroy(fm_cell_nodes_t* nodes) {
  delete nodes;
}

extern "C" fm_status_t fm_cell_nodes_get_node_count(const fm_cell_nodes_t* nodes, size_t* out_count) {
  clear_last_error();
  if (nodes == nullptr || out_count == nullptr) {
    return null_argument();
  }
  *out_count = nodes->nodes.size();
  return 0;
}

extern "C" fm_status_t fm_cell_nodes_get_node_at(const fm_cell_nodes_t* nodes, std::size_t idx, fm_cell_node_t* out) {
  clear_last_error();
  if (nodes == nullptr || out == nullptr) {
    return null_argument();
  }
  if (auto rc = check_index(idx, nodes->nodes.size(), "index"); rc != 0) {
    return rc;
  }
  if (auto rc = check_record_size(*out); rc != 0) {
    return rc;
  }
  out->sheet = static_cast<std::uint32_t>(nodes->nodes[idx].sheet_id);
  out->row = nodes->nodes[idx].row;
  out->col = nodes->nodes[idx].col;
  return 0;
}
