//
// Implementation of the load-time conditional-format merge. See cf_merge.h
// for the contract. Rules are grouped by an anchor-independent definition
// key, so only rules of one group are ever paired; ranges are compared by
// coordinate compression, never cell by cell.

#include "io/cf_merge.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <string_view>
#include <unordered_map>
#include <unordered_set>
#include <utility>
#include <vector>

#include "cell.h"
#include "cf/cf_types.h"
#include "io/styles_writer.h"
#include "io/xlsb/cf_records.h"
#include "parser/ast_format_r1c1.h"
#include "parser/parser.h"
#include "pugixml.hpp"
#include "sheet.h"
#include "styles.h"
#include "utils/arena.h"
#include "utils/index_sort.h"
#include "workbook.h"

// @size-budget: 6 KB

namespace formulon::io {
namespace {

using Sqref = std::vector<cf::CFCellRange>;

/// Definition group of each rule, shaped like the sheet's blocks.
using Groups = std::vector<std::vector<std::size_t>>;

constexpr std::size_t kNoGroup = SIZE_MAX;

struct RuleRef {
  std::size_t block;
  std::size_t rule;
};

/// Inclusive bounds of a cell set.
struct Box {
  std::uint32_t top;
  std::uint32_t left;
  std::uint32_t bottom;
  std::uint32_t right;
};

enum class Shape : std::uint8_t { kNone, kRectangle, kEqualSet };

bool IsMergingType(cf::RuleType type) {
  switch (type) {
    case cf::RuleType::Expression:
    case cf::RuleType::CellIs:
    case cf::RuleType::ContainsText:
    case cf::RuleType::NotContainsText:
    case cf::RuleType::BeginsWith:
    case cf::RuleType::EndsWith:
    case cf::RuleType::ContainsBlanks:
    case cf::RuleType::NotContainsBlanks:
    case cf::RuleType::ContainsErrors:
    case cf::RuleType::NotContainsErrors:
    case cf::RuleType::TimePeriod:
      return true;
    default:
      return false;
  }
}

template <typename T>
void AppendRaw(std::string& key, const T& value) {
  key.append(reinterpret_cast<const char*>(&value), sizeof value);
}

void AppendText(std::string& key, std::string_view text) {
  AppendRaw(key, text.size());
  key.append(text);
}

/// Appends `formula` as R1C1 text against `anchor`; false when it does not parse.
bool AppendFormula(std::string& key, const std::optional<std::string>& formula, CellAddress anchor, Arena& arena) {
  key.push_back(formula.has_value() ? '\1' : '\0');
  if (!formula.has_value()) {
    return true;
  }
  const parser::AstNode* root = parser::parse_strict(*formula, arena);
  if (root == nullptr) {
    return false;
  }
  AppendText(key, parser::format_formula_r1c1(*root, anchor.row, anchor.col));
  return true;
}

/// The definition two rules must share to merge, or nothing for a rule that
/// never merges.
std::optional<std::string> DefinitionKey(const cf::ConditionalFormat& block, const cf::CFRule& rule,
                                         const std::vector<DifferentialFormat>& dxfs, Arena& arena) {
  if (!IsMergingType(rule.type) || block.pivot_scope || block.ext_lst_raw.has_value() || rule.ext_lst_raw.has_value() ||
      block.sqref.empty() || (rule.dxf_id.has_value() && *rule.dxf_id >= dxfs.size())) {
    return std::nullopt;
  }
  std::string key;
  AppendRaw(key, rule.type);
  AppendRaw(key, rule.stop_if_true);
  key.push_back(rule.op.has_value() ? '\1' : '\0');
  AppendRaw(key, rule.op.value_or(cf::CellIsOperator::LessThan));
  key.push_back(rule.text.has_value() ? '\1' : '\0');
  AppendText(key, rule.text.value_or(std::string()));
  key.push_back(rule.time_period.has_value() ? '\1' : '\0');
  AppendRaw(key, rule.time_period.value_or(cf::TimePeriod{}));
  key.push_back(rule.rank.has_value() ? '\1' : '\0');
  AppendRaw(key, rule.rank.value_or(0));
  AppendRaw(key, rule.percent);
  AppendRaw(key, rule.bottom);
  AppendRaw(key, rule.above_average);
  AppendRaw(key, rule.equal_average);
  key.push_back(rule.std_dev.has_value() ? '\1' : '\0');
  AppendRaw(key, rule.std_dev.value_or(0.0));
  const CellAddress anchor = cf::sqref_anchor(block.sqref);
  const bool parsed =
      AppendFormula(key, rule.formula1, anchor, arena) && AppendFormula(key, rule.formula2, anchor, arena);
  arena.reset();
  if (!parsed) {
    return std::nullopt;
  }
  AppendText(key, rule.dxf_id.has_value() ? dxf_body_fragment(dxfs[*rule.dxf_id]) : std::string("none"));
  return key;
}

Box BoundingBox(const Sqref& sqref) {
  Box box{sqref.front().first.row, sqref.front().first.col, sqref.front().last.row, sqref.front().last.col};
  for (const cf::CFCellRange& range : sqref) {
    box.top = std::min(box.top, range.first.row);
    box.left = std::min(box.left, range.first.col);
    box.bottom = std::max(box.bottom, range.last.row);
    box.right = std::max(box.right, range.last.col);
  }
  return box;
}

bool Intersects(const Sqref& a, const Sqref& b) {
  for (const cf::CFCellRange& x : a) {
    for (const cf::CFCellRange& y : b) {
      if (x.first.row <= y.last.row && y.first.row <= x.last.row && x.first.col <= y.last.col &&
          y.first.col <= x.last.col) {
        return true;
      }
    }
  }
  return false;
}

/// Sorted distinct values of `cuts`.
void SortUnique(std::vector<std::uint32_t>& cuts) {
  sort_ascending(cuts);
  cuts.erase(std::unique(cuts.begin(), cuts.end()), cuts.end());
}

std::size_t CutIndex(const std::vector<std::uint32_t>& cuts, std::uint32_t value) {
  return static_cast<std::size_t>(std::lower_bound(cuts.begin(), cuts.end(), value) - cuts.begin());
}

/// Whether the cell sets of `a` and `b` are equal, union to their bounding
/// box, or neither. Both are cut at every area edge; each cell of the
/// resulting grid is wholly in or out of each set.
Shape ClassifyUnion(const Sqref& a, const Sqref& b) {
  const Box x = BoundingBox(a);
  const Box y = BoundingBox(b);
  // Sets that neither overlap nor touch cannot make one rectangle, nor be equal.
  if (x.top > y.bottom + 1U || y.top > x.bottom + 1U || x.left > y.right + 1U || y.left > x.right + 1U) {
    return Shape::kNone;
  }
  std::vector<std::uint32_t> rows;
  std::vector<std::uint32_t> cols;
  for (const Sqref* sqref : {&a, &b}) {
    for (const cf::CFCellRange& range : *sqref) {
      rows.push_back(range.first.row);
      rows.push_back(range.last.row + 1U);
      cols.push_back(range.first.col);
      cols.push_back(range.last.col + 1U);
    }
  }
  SortUnique(rows);
  SortUnique(cols);
  const std::size_t width = cols.size() - 1U;
  std::vector<std::uint8_t> cover((rows.size() - 1U) * width, 0U);
  const auto mark = [&](const Sqref& sqref, std::uint8_t bit) {
    for (const cf::CFCellRange& range : sqref) {
      const std::size_t row_end = CutIndex(rows, range.last.row + 1U);
      const std::size_t col_end = CutIndex(cols, range.last.col + 1U);
      for (std::size_t r = CutIndex(rows, range.first.row); r < row_end; ++r) {
        for (std::size_t c = CutIndex(cols, range.first.col); c < col_end; ++c) {
          cover[r * width + c] = static_cast<std::uint8_t>(cover[r * width + c] | bit);
        }
      }
    }
  };
  mark(a, 1U);
  mark(b, 2U);
  bool equal = true;
  bool full = true;
  for (const std::uint8_t cell : cover) {
    equal = equal && (cell == 0U || cell == 3U);
    full = full && cell != 0U;
  }
  return equal ? Shape::kEqualSet : (full ? Shape::kRectangle : Shape::kNone);
}

/// The sheet's rules in model order, in priority order, and per group.
struct SheetState {
  std::vector<RuleRef> refs;
  std::vector<std::uint32_t> by_priority;
  std::vector<std::vector<std::uint32_t>> members;
};

std::int32_t Priority(const std::vector<cf::ConditionalFormat>& formats, RuleRef ref) {
  return formats[ref.block].rules[ref.rule].priority;
}

void BuildState(const std::vector<cf::ConditionalFormat>& formats, const Groups& groups, SheetState& state) {
  state.refs.clear();
  for (std::size_t b = 0; b < formats.size(); ++b) {
    for (std::size_t r = 0; r < formats[b].rules.size(); ++r) {
      state.refs.push_back(RuleRef{b, r});
    }
  }
  const auto less = [&formats, &state](std::uint32_t x, std::uint32_t y) {
    return Priority(formats, state.refs[x]) < Priority(formats, state.refs[y]);
  };
  sorted_index_order(state.by_priority, static_cast<std::uint32_t>(state.refs.size()), make_index_less(less));
  for (std::vector<std::uint32_t>& members : state.members) {
    members.clear();
  }
  for (const std::uint32_t index : state.by_priority) {
    const std::size_t group = groups[state.refs[index].block][state.refs[index].rule];
    if (group != kNoGroup) {
      state.members[group].push_back(index);
    }
  }
}

/// Condition (c): no rule other than the pair with a priority from A's to
/// B's intersects B's range.
bool NoIntermediate(const std::vector<cf::ConditionalFormat>& formats, const SheetState& state, std::uint32_t a,
                    std::uint32_t b) {
  const std::int32_t low = Priority(formats, state.refs[a]);
  const std::int32_t high = Priority(formats, state.refs[b]);
  const Sqref& target = formats[state.refs[b].block].sqref;
  auto it = std::lower_bound(state.by_priority.begin(), state.by_priority.end(), low,
                             [&formats, &state](std::uint32_t index, std::int32_t priority) {
                               return Priority(formats, state.refs[index]) < priority;
                             });
  for (; it != state.by_priority.end() && Priority(formats, state.refs[*it]) <= high; ++it) {
    if (*it != a && *it != b && Intersects(formats[state.refs[*it].block].sqref, target)) {
      return false;
    }
  }
  return true;
}

/// Folds B into A: A takes the merged range, B and a block it leaves empty go.
void Merge(std::vector<cf::ConditionalFormat>& formats, Groups& groups, RuleRef a, RuleRef b, Shape shape) {
  Sqref merged = formats[a.block].sqref;
  if (shape == Shape::kRectangle) {
    const Box x = BoundingBox(formats[a.block].sqref);
    const Box y = BoundingBox(formats[b.block].sqref);
    merged = Sqref{cf::CFCellRange{{std::min(x.top, y.top), std::min(x.left, y.left)},
                                   {std::max(x.bottom, y.bottom), std::max(x.right, y.right)}}};
  }
  const bool moved = merged != formats[a.block].sqref;
  if (moved && cf::sqref_anchor(merged) != cf::sqref_anchor(formats[a.block].sqref)) {
    // The new anchor is a cell of B's range, which B's text is anchored at.
    cf::CFRule& kept = formats[a.block].rules[a.rule];
    const cf::CFRule& folded = formats[b.block].rules[b.rule];
    kept.formula1 = folded.formula1;
    kept.formula2 = folded.formula2;
  }

  const auto b_rule = static_cast<std::ptrdiff_t>(b.rule);
  formats[b.block].rules.erase(formats[b.block].rules.begin() + b_rule);
  groups[b.block].erase(groups[b.block].begin() + b_rule);
  if (formats[b.block].rules.empty()) {
    formats.erase(formats.begin() + static_cast<std::ptrdiff_t>(b.block));
    groups.erase(groups.begin() + static_cast<std::ptrdiff_t>(b.block));
    if (b.block < a.block) {
      --a.block;
    }
  } else if (b.block == a.block && b.rule < a.rule) {
    --a.rule;
  }

  if (!moved) {
    return;
  }
  cf::ConditionalFormat& home = formats[a.block];
  if (home.rules.size() == 1U) {
    home.sqref = std::move(merged);
    return;
  }
  cf::ConditionalFormat own;
  own.sqref = std::move(merged);
  own.pivot_scope = home.pivot_scope;
  const auto a_rule = static_cast<std::ptrdiff_t>(a.rule);
  own.rules.push_back(std::move(home.rules[a.rule]));
  const std::size_t group = groups[a.block][a.rule];
  home.rules.erase(home.rules.begin() + a_rule);
  groups[a.block].erase(groups[a.block].begin() + a_rule);
  const auto after = static_cast<std::ptrdiff_t>(a.block + 1U);
  formats.insert(formats.begin() + after, std::move(own));
  groups.insert(groups.begin() + after, std::vector<std::size_t>{group});
}

/// Merges the first mergeable pair of `group` in (priority, model position)
/// order; false when no pair merges.
bool MergeFirstPair(std::vector<cf::ConditionalFormat>& formats, Groups& groups, const SheetState& state,
                    std::size_t group) {
  const std::vector<std::uint32_t>& members = state.members[group];
  for (std::size_t i = 0; i < members.size(); ++i) {
    for (std::size_t j = i + 1U; j < members.size(); ++j) {
      const RuleRef a = state.refs[members[i]];
      const RuleRef b = state.refs[members[j]];
      const Shape shape = ClassifyUnion(formats[a.block].sqref, formats[b.block].sqref);
      if (shape != Shape::kNone && NoIntermediate(formats, state, members[i], members[j])) {
        Merge(formats, groups, a, b, shape);
        return true;
      }
    }
  }
  return false;
}

std::unordered_set<std::string> ModelRuleIds(const std::vector<cf::ConditionalFormat>& formats) {
  std::unordered_set<std::string> ids;
  for (const cf::ConditionalFormat& format : formats) {
    for (const cf::CFRule& rule : format.rules) {
      if (!rule.id.empty()) {
        ids.insert(rule.id);
      }
    }
  }
  return ids;
}

/// True when the sheet's retained `<extLst>` still holds an x14 rule whose id
/// names no model rule (one the reader could not adopt).
bool ExtensionHoldsUnmodelledCf(const std::string& ext_lst_xml, const std::vector<cf::ConditionalFormat>& formats) {
  if (ext_lst_xml.find("conditionalFormatting") == std::string::npos) {
    return false;
  }
  pugi::xml_document doc;
  if (!doc.load_string(ext_lst_xml.c_str())) {
    return true;
  }
  const std::unordered_set<std::string> ids = ModelRuleIds(formats);
  for (pugi::xml_node ext : doc.child("extLst").children("ext")) {
    for (pugi::xml_node block : ext.child("x14:conditionalFormattings").children("x14:conditionalFormatting")) {
      for (pugi::xml_node rule : block.children("x14:cfRule")) {
        if (ids.count(rule.attribute("id").value()) == 0U) {
          return true;
        }
      }
    }
  }
  return false;
}

void MergeSheet(Sheet& sheet, const std::vector<DifferentialFormat>& dxfs, Arena& arena) {
  std::vector<cf::ConditionalFormat>& formats = sheet.mutable_conditional_formats();
  std::size_t rule_count = 0;
  for (const cf::ConditionalFormat& format : formats) {
    rule_count += format.rules.size();
  }
  if (rule_count < 2U || ExtensionHoldsUnmodelledCf(sheet.ext_lst_xml(), formats) ||
      xlsb::tail_holds_unmodelled_cf(sheet.xlsb_tail(), formats)) {
    return;
  }
  std::unordered_map<std::string, std::size_t> group_of;
  Groups groups(formats.size());
  for (std::size_t b = 0; b < formats.size(); ++b) {
    groups[b].assign(formats[b].rules.size(), kNoGroup);
    for (std::size_t r = 0; r < formats[b].rules.size(); ++r) {
      std::optional<std::string> key = DefinitionKey(formats[b], formats[b].rules[r], dxfs, arena);
      if (key.has_value()) {
        const std::size_t next = group_of.size();
        groups[b][r] = group_of.emplace(std::move(*key), next).first->second;
      }
    }
  }
  SheetState state;
  state.members.resize(group_of.size());
  BuildState(formats, groups, state);
  // A merge changes ranges another group's condition (c) reads, so passes
  // repeat until one merges nothing.
  bool merged = true;
  while (merged) {
    merged = false;
    for (std::size_t group = 0; group < state.members.size(); ++group) {
      while (state.members[group].size() >= 2U && MergeFirstPair(formats, groups, state, group)) {
        merged = true;
        BuildState(formats, groups, state);
      }
    }
  }
}

}  // namespace

void merge_cf_rules_on_load(Workbook& wb) {
  Arena arena;
  const std::vector<DifferentialFormat>& dxfs = wb.styles().dxfs;
  for (std::size_t i = 0; i < wb.sheet_count(); ++i) {
    MergeSheet(wb.sheet(i), dxfs, arena);
  }
}

}  // namespace formulon::io
