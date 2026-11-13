//
// Implementation of the saved conditional-format layout. See cf_layout.h.

// @size-budget: 2 KB

#include "io/cf_layout.h"

#include <cstddef>
#include <cstdint>
#include <string>
#include <tuple>
#include <unordered_map>

#include "utils/index_sort.h"

namespace formulon::io {
namespace {

/// Identity of the blocks that combine: the sqref sequence and the pivot flag.
std::string GroupKey(const cf::ConditionalFormat& block) {
  std::string key;
  key.reserve(block.sqref.size() * 16U + 1U);
  key.push_back(block.pivot_scope ? 'p' : 'n');
  for (const cf::CFCellRange& r : block.sqref) {
    for (const std::uint32_t v : {r.first.col, r.first.row, r.last.col, r.last.row}) {
      key.append(reinterpret_cast<const char*>(&v), sizeof(v));
    }
  }
  return key;
}

auto OrderKey(const CfBlockView& view) {
  if (view.sqref->empty()) {
    return std::make_tuple(0U, 0U, 0U, 0U);
  }
  const cf::CFCellRange& r = view.sqref->front();
  return std::make_tuple(static_cast<std::uint32_t>(r.first.col), static_cast<std::uint32_t>(r.first.row),
                         static_cast<std::uint32_t>(r.last.col), static_cast<std::uint32_t>(r.last.row));
}

}  // namespace

std::vector<CfBlockView> cf_saved_layout(const std::vector<cf::ConditionalFormat>& formats) {
  std::vector<CfBlockView> views;
  views.reserve(formats.size());
  std::unordered_map<std::string, std::size_t> group_of;
  for (const cf::ConditionalFormat& block : formats) {
    std::size_t at = views.size();
    if (block.ext_lst_raw.has_value()) {
      views.emplace_back();
    } else {
      const auto [it, added] = group_of.try_emplace(GroupKey(block), views.size());
      at = it->second;
      if (added) {
        views.emplace_back();
      }
    }
    CfBlockView& view = views[at];
    view.sqref = &block.sqref;
    view.pivot_scope = block.pivot_scope;
    view.ext_lst_raw = &block.ext_lst_raw;
    for (const cf::CFRule& rule : block.rules) {
      view.rules.push_back(&rule);
    }
  }
  for (CfBlockView& view : views) {
    sort_by_index(view.rules, [](const cf::CFRule* a, const cf::CFRule* b) { return a->priority < b->priority; });
  }
  sort_by_index(views, [](const CfBlockView& a, const CfBlockView& b) { return OrderKey(a) < OrderKey(b); });
  return views;
}

}  // namespace formulon::io
