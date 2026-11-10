//
// `<rowItems>` / `<colItems>` lines of a projected pivot, read off the layout
// plan's axis entries and grand-total strips.
//
// A compact / outline row axis gives every line one `<x>` at its own level
// (`r` = that level). Tabular rows and every column axis list a leaf's path
// from the first level it does not share with the previous leaf. Subtotal and
// blank lines name only the item they close. An axis whose shape has no known
// encoding is refused rather than guessed, leaving the other axis intact.
//
// @size-budget: 12 KB

#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/pivot_layout.h"
#include "pivot/pivot_layout_internal.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "utils/error.h"

namespace formulon::pivot::detail {
namespace {

PivotLineType line_type_for(SubtotalFn fn) {
  switch (fn) {
    case SubtotalFn::Sum:
      return PivotLineType::kSum;
    case SubtotalFn::Count:
      return PivotLineType::kCountA;
    case SubtotalFn::Average:
      return PivotLineType::kAverage;
    case SubtotalFn::Max:
      return PivotLineType::kMax;
    case SubtotalFn::Min:
      return PivotLineType::kMin;
    case SubtotalFn::Product:
      return PivotLineType::kProduct;
    case SubtotalFn::CountNumbers:
      return PivotLineType::kCount;
    case SubtotalFn::StdDev:
      return PivotLineType::kStdDev;
    case SubtotalFn::StdDevP:
      return PivotLineType::kStdDevP;
    case SubtotalFn::Var:
      return PivotLineType::kVar;
    case SubtotalFn::VarP:
      return PivotLineType::kVarP;
  }
  return PivotLineType::kDefault;
}

/// What one axis contributes to its item list.
struct AxisView {
  const char* name = "";
  const std::vector<AxisEntry>* entries = nullptr;
  const AxisShape* shape = nullptr;
  const std::vector<std::uint32_t>* field_order = nullptr;
  bool values_here = false;    ///< The Values level is one of this axis' levels.
  bool stored_values = false;  ///< The model records a Values position on this axis.
  bool group_rows = false;     ///< Compact / outline row axis.
  std::size_t data_fields = 0;
  std::size_t grand_lines = 0;  ///< Grand-total lines closing the list.
};

Error refuse(const AxisView& axis, const char* shape) {
  return make_error(FormulonErrorCode::kNotImplemented, std::string("pivot line items: ") + shape,
                    std::string("axis=") + axis.name);
}

bool has_form_override(const PivotField& field) {
  for (const auto& attr : field.passthrough_attrs) {
    if (attr.first == "compact" || attr.first == "outline" || attr.first == "compactData" ||
        attr.first == "outlineData") {
      return true;
    }
  }
  return false;
}

std::size_t real_node_count(const AxisEntry& entry) {
  std::size_t count = 0;
  for (const AxisHierarchyNode* node : entry.nodes) {
    count += node != nullptr ? 1U : 0U;
  }
  return count;
}

/// The refusal for an axis whose shape no measured encoding covers.
std::optional<Error> axis_refusal(const LayoutPlan& plan, const AxisView& axis, bool rows) {
  if (plan.legacy) {
    return refuse(axis, "the legacy projection has no Excel layout");
  }
  if (plan.data_field_count == 0 && plan.row_depth == 0 && plan.col_depth == 0) {
    return refuse(axis, "a report with no data fields and no row or column fields");
  }
  if (axis.stored_values && !axis.values_here) {
    return refuse(axis, "a stored Values position the layout places no Values level at");
  }
  const AxisShape& shape = *axis.shape;
  if (shape.values_at.has_value() && shape.depth >= *shape.values_at + 2) {
    return refuse(axis, "a real field with children below the Values level");
  }
  for (std::size_t depth = 0; depth < shape.depth; ++depth) {
    const PivotField* field = depth < shape.fields.size() ? shape.fields[depth] : nullptr;
    if (field == nullptr) {
      continue;
    }
    const bool subtotalled = depth + 1 < shape.depth && !shape.subtotal_fns[depth].empty();
    if (subtotalled && !rows) {
      return refuse(axis, "custom subtotal functions on the column axis");
    }
    if (subtotalled && shape.values_at.has_value() && *shape.values_at > depth) {
      return refuse(axis, "custom subtotal functions with the Values level below the field");
    }
    if (rows && has_form_override(*field)) {
      return refuse(axis, "a row field whose form differs from the table's");
    }
    if (field->show_all) {
      return refuse(axis, "a field listing items without data");
    }
  }
  if (!rows && !plan.table.grand_totals_rows()) {
    return refuse(axis, "column grand totals turned off");
  }
  for (const AxisEntry& entry : *axis.entries) {
    if (entry.kind == EntryKind::Subtotal && entry.heads_group && !entry.nodes.empty() &&
        entry.nodes.back() == nullptr) {
      return refuse(axis, "a collapsed item with the Values level below it");
    }
    if (!axis.group_rows && entry.kind == EntryKind::Subtotal && entry.heads_group && real_node_count(entry) >= 2) {
      return refuse(axis, "a collapsed inner item in tabular form or on the column axis");
    }
    if (axis.group_rows && entry.kind == EntryKind::Subtotal && !entry.heads_group && !entry.custom_fn.has_value() &&
        entry.subtotal_label.empty()) {
      return refuse(axis, "a subtotal below its group in compact or outline form");
    }
    if (axis.group_rows && entry.kind == EntryKind::Blank && real_node_count(entry) >= 2) {
      return refuse(axis, "a blank row after an inner field in compact or outline form");
    }
  }
  return std::nullopt;
}

/// The `<x>` for level `level` of `entry`: the Values level, or the item there.
PivotLineX x_at(const AxisEntry& entry, std::size_t level, const AxisView& axis) {
  PivotLineX x;
  const AxisHierarchyNode* node = entry.nodes[level];
  if (node == nullptr) {
    x.values_level = true;
    x.data_index = static_cast<std::uint32_t>(entry.data_field.value_or(0));
    return x;
  }
  const std::optional<std::size_t>& values_at = axis.shape->values_at;
  const std::size_t depth = values_at.has_value() && *values_at < level ? level - 1 : level;
  x.field_index = (*axis.field_order)[depth];
  x.node = node;
  return x;
}

bool same_step(const AxisEntry& a, const AxisEntry& b, std::size_t level) {
  if (a.nodes[level] != b.nodes[level]) {
    return false;
  }
  return a.nodes[level] != nullptr || a.data_field == b.data_field;
}

std::vector<PivotLine> axis_lines(const AxisView& axis) {
  std::vector<PivotLine> lines;
  if (axis.shape->depth == 0) {
    // No real field: one line per data field when Values sits here, else one empty line.
    if (!axis.values_here) {
      lines.emplace_back();
      return lines;
    }
    for (std::size_t df = 0; df < axis.data_fields; ++df) {
      PivotLine line;
      line.i = static_cast<std::uint32_t>(df);
      line.xs.emplace_back();
      line.xs.back().values_level = true;
      line.xs.back().data_index = line.i;
      lines.push_back(std::move(line));
    }
    return lines;
  }
  const AxisEntry* prev_leaf = nullptr;
  for (const AxisEntry& entry : *axis.entries) {
    PivotLine line;
    line.i = static_cast<std::uint32_t>(entry.data_field.value_or(0));
    const std::size_t last = entry.nodes.size() - 1;
    const bool plain = entry.kind == EntryKind::Leaf || entry.kind == EntryKind::Group ||
                       (entry.kind == EntryKind::Subtotal && entry.heads_group);
    if (entry.kind == EntryKind::Blank) {
      line.type = PivotLineType::kBlank;
    } else if (!plain) {
      line.type = entry.custom_fn.has_value() ? line_type_for(*entry.custom_fn) : PivotLineType::kDefault;
    }
    std::size_t first = last;
    if (plain && !axis.group_rows) {
      first = 0;
      while (prev_leaf != nullptr && first < last && first < prev_leaf->nodes.size() &&
             same_step(*prev_leaf, entry, first)) {
        ++first;
      }
      prev_leaf = &entry;
    }
    line.r = static_cast<std::uint32_t>(first);
    for (std::size_t level = first; level <= last; ++level) {
      line.xs.push_back(x_at(entry, level, axis));
    }
    lines.push_back(std::move(line));
  }
  for (std::size_t b = 0; b < axis.grand_lines; ++b) {
    PivotLine line;
    line.type = PivotLineType::kGrand;
    line.i = axis.values_here ? static_cast<std::uint32_t>(b) : 0U;
    line.xs.emplace_back();
    lines.push_back(std::move(line));
  }
  return lines;
}

}  // namespace

PivotLineItems build_line_items(const LayoutPlan& plan) {
  AxisView rows;
  rows.name = "row";
  rows.entries = &plan.row_entries;
  rows.shape = &plan.row_shape;
  rows.field_order = &plan.table.row_field_order();
  rows.values_here = plan.values_on_rows;
  rows.stored_values = plan.table.row_values_position().has_value();
  rows.group_rows = plan.row_shape.group_rows;
  rows.data_fields = plan.data_field_count;
  rows.grand_lines = plan.total_strip_rows;
  AxisView cols;
  cols.name = "column";
  cols.entries = &plan.col_entries;
  cols.shape = &plan.col_shape;
  cols.field_order = &plan.table.col_field_order();
  cols.values_here = plan.values_on_cols;
  cols.stored_values = plan.table.col_values_position().has_value();
  cols.data_fields = plan.data_field_count;
  cols.grand_lines = plan.total_strip_cols;

  PivotLineItems items;
  items.row_refusal = axis_refusal(plan, rows, /*rows=*/true);
  if (!items.row_refusal) {
    items.rows = axis_lines(rows);
  }
  items.col_refusal = axis_refusal(plan, cols, /*rows=*/false);
  if (!items.col_refusal) {
    items.cols = axis_lines(cols);
  }
  return items;
}

}  // namespace formulon::pivot::detail
