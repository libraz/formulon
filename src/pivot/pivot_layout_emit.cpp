//
// Cell emission for the pivot layout: page block, header block, body rows,
// grand-total strips and merged ranges, projected from a `LayoutPlan`.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string>
#include <utility>
#include <vector>

#include "pivot/field_lookup.h"
#include "pivot/pivot_layout.h"
#include "pivot/pivot_layout_internal.h"
#include "pivot/pivot_locale.h"
#include "pivot/pivot_result.h"
#include "pivot/pivot_table.h"
#include "pivot/pivot_types.h"
#include "pivot/value_order.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon::pivot::detail {
namespace {

const std::string& data_field_name(const PivotTable& table, std::size_t index) {
  static const std::string kNone;
  return index < table.data_fields().size() ? table.data_fields()[index].name : kNone;
}

const std::string& data_field_format(const PivotTable& table, std::size_t index) {
  static const std::string kNone;
  return index < table.data_fields().size() ? table.data_fields()[index].number_format : kNone;
}

// A show-as data field states its total cells itself; a blank one stays blank instead of
// being re-summed from the raw cells.
bool total_is_derived(const PivotTable& table, std::size_t index) {
  return index < table.data_fields().size() && table.data_fields()[index].show_as != ShowValuesAs::Normal;
}

Value text_value(PivotCells& cells, std::string text) {
  cells.text_storage.push_back(std::move(text));
  return Value::text(cells.text_storage.back());
}

Value reify_value(PivotCells& cells, const Value& value) {
  if (!value.is_text()) {
    return value;
  }
  return text_value(cells, std::string(value.as_text()));
}

void append_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, Value value, PivotCellKind kind,
                 PivotCellRole role, std::uint32_t depth, std::string field_name = {}, std::string number_format = {}) {
  PivotCell cell;
  cell.row = row;
  cell.col = col;
  cell.value = value;
  cell.kind = kind;
  cell.role = role;
  cell.depth = depth;
  cell.field_name = std::move(field_name);
  cell.number_format = std::move(number_format);
  cells.cells.push_back(std::move(cell));
}

void append_blank(PivotCells& cells, std::uint32_t row, std::uint32_t col, std::uint32_t depth = 0) {
  append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::None, depth);
}

// A depth-0 cell of data field `df`, carrying its name and number format.
void append_df_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, Value value, PivotCellKind kind,
                    PivotCellRole role, const PivotTable& table, std::size_t df) {
  append_cell(cells, row, col, value, kind, role, 0, data_field_name(table, df), data_field_format(table, df));
}

PivotCellRole value_role(PivotCellKind kind) {
  switch (kind) {
    case PivotCellKind::RowSubtotal:
    case PivotCellKind::ColSubtotal:
      return PivotCellRole::SubtotalValue;
    case PivotCellKind::GrandTotal:
      return PivotCellRole::GrandTotalValue;
    default:
      return PivotCellRole::DataValue;
  }
}

// The cell Excel writes for `value`: an error under `showError` shows the error
// caption, and an empty cell shows a non-empty missing caption, where a lone
// "-" is written as the number 0.
Value shown_value(const PivotTable& table, PivotCells& cells, const Value& value) {
  if (value.is_error() && table.show_error()) {
    const std::string caption = table.error_caption().value_or(std::string());
    return caption.empty() ? Value::blank() : text_value(cells, caption);
  }
  if (value.is_blank() && table.missing_caption().has_value() && !table.missing_caption()->empty()) {
    return *table.missing_caption() == "-" ? Value::number(0.0) : text_value(cells, *table.missing_caption());
  }
  return reify_value(cells, value);
}

// A value cell of data field `df`, as Excel shows `value`.
void append_value_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col, const Value& value, PivotCellKind kind,
                       const PivotTable& table, std::size_t df) {
  append_df_cell(cells, row, col, shown_value(table, cells, value), kind, value_role(kind), table, df);
}

// A value cell left empty whatever the captions say: a group heading row's or
// an inserted blank row's.
void append_empty_value_cell(PivotCells& cells, std::uint32_t row, std::uint32_t col) {
  append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::DataValue, 0);
}

// `count` empty value cells in `row`, starting at `first_col`.
void append_empty_value_cells(PivotCells& cells, std::uint32_t row, std::uint32_t first_col, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    append_empty_value_cell(cells, row, first_col + static_cast<std::uint32_t>(i));
  }
}

std::string plan_subtotal_label(const LayoutPlan& plan, const std::string& group) {
  return plan.options.subtotal_prefix + group + plan.subtotal_suffix;
}

// A custom subtotal row puts its function word where the locale puts the
// automatic subtotal word: after the item, or before it in prefix locales.
std::string entry_subtotal_label(const LayoutPlan& plan, const AxisEntry& entry, const std::string& group) {
  if (!entry.subtotal_label.empty()) {
    return entry.subtotal_label;
  }
  if (entry.custom_fn.has_value()) {
    const std::string word(
        subtotal_function_label(aggregation_from_subtotal_fn(*entry.custom_fn), plan.options.profile));
    return plan.options.subtotal_prefix.empty() ? group + " " + word : word + " " + group;
  }
  return plan_subtotal_label(plan, group);
}

std::string data_total_label(const LayoutPlan& plan, std::size_t df) {
  return plan.options.data_total_prefix + data_field_name(plan.table, df) + plan.options.data_total_suffix;
}

// Total of a strip of cells: the first source error wins, an all-numeric strip sums, and anything
// else (or no number at all) is blank.
struct StripTotal {
  double sum = 0.0;
  bool summable = true;
  bool any_number = false;
  const Value* source_error = nullptr;

  // Returns false once an error cell ends the strip.
  bool add(const Value& v) {
    if (v.is_error()) {
      source_error = &v;
      return false;
    }
    if (v.is_number()) {
      sum += v.as_number();
      any_number = true;
    } else if (!v.is_blank()) {
      summable = false;
    }
    return true;
  }

  Value value() const {
    if (source_error != nullptr) {
      return *source_error;
    }
    return summable && any_number ? Value::number(sum) : Value::blank();
  }
};

// A field caption, or nothing while `showHeaders` hides them.
std::string field_caption(const LayoutPlan& plan, std::string text) {
  return plan.table.show_headers() ? std::move(text) : std::string();
}

// Compact form indents a level-`level` row label by `indent + 1` per level.
void set_compact_indent(const LayoutPlan& plan, PivotCells& cells, std::size_t level) {
  cells.cells.back().label_indent = static_cast<std::uint32_t>(level) * (plan.table.indent() + 1U);
}

// The field-order index of axis level `level` once the Values level at `values_at` is skipped;
// nullopt at the Values level itself.
std::optional<std::size_t> field_level(const std::optional<std::size_t>& values_at, std::size_t level) {
  if (values_at.has_value()) {
    if (level == *values_at) {
      return std::nullopt;
    }
    if (level > *values_at) {
      --level;
    }
  }
  return level;
}

// Display name of the row / column level at `level`: a field's display
// name, or the stored Values caption for the Values level.
std::string level_name(const LayoutPlan& plan, const std::vector<std::uint32_t>& order,
                       const std::optional<std::size_t>& values_at, std::size_t level) {
  const PivotTable& table = plan.table;
  const std::optional<std::size_t> real = field_level(values_at, level);
  if (!real.has_value()) {
    return table.data_caption();
  }
  level = *real;
  if (level < order.size() && order[level] < table.fields().size()) {
    return pivot_field_display_name(table.fields()[order[level]]);
  }
  return std::string();
}

std::string row_level_name(const LayoutPlan& plan, std::size_t level) {
  return level < plan.row_levels ? level_name(plan, plan.table.row_field_order(), plan.row_shape.values_at, level)
                                 : std::string();
}

std::string col_level_name(const LayoutPlan& plan, std::size_t level) {
  return level < plan.col_levels ? level_name(plan, plan.table.col_field_order(), plan.col_shape.values_at, level)
                                 : std::string();
}

// The field at `shape`'s level `level`, or null at the Values level.
const PivotField* axis_level_field(const AxisShape& shape, std::size_t level) {
  const std::optional<std::size_t> real = field_level(shape.values_at, level);
  if (!real.has_value()) {
    return nullptr;
  }
  return *real < shape.fields.size() ? shape.fields[*real] : nullptr;
}

// Number format of the item at `entry.labels[level]`, for a cell drawing that label bare.
std::string label_format(const AxisEntry& entry, std::size_t level) {
  return level < entry.nodes.size() && entry.nodes[level] != nullptr ? entry.nodes[level]->number_format
                                                                     : std::string();
}

// The number behind the item at `entry.labels[level]`, when it is one.
std::optional<Value> label_item_value(const AxisEntry& entry, std::size_t level) {
  if (level >= entry.nodes.size() || entry.nodes[level] == nullptr) {
    return std::nullopt;
  }
  const AxisHierarchyNode& node = *entry.nodes[level];
  if (node.identity_is_group_item || !node.identity.has_value() || !node.identity->is_number()) {
    return std::nullopt;
  }
  return node.identity;
}

void emit_text_or_blank(PivotCells& cells, std::uint32_t row, std::uint32_t col, const std::string& text,
                        PivotCellKind kind, PivotCellRole role, std::uint32_t depth, std::string field_name = {},
                        std::string number_format = {}) {
  if (text.empty()) {
    append_cell(cells, row, col, Value::blank(), PivotCellKind::Blank, PivotCellRole::None, depth,
                std::move(field_name));
  } else {
    append_cell(cells, row, col, text_value(cells, text), kind, role, depth, std::move(field_name),
                std::move(number_format));
  }
}

// The data field a cell at (`row_entry`, `col_entry`) reports: whichever
// axis carries the Values level decides, and a single field is field 0.
std::size_t cell_data_field(const AxisEntry& row_entry, const AxisEntry& col_entry) {
  if (row_entry.data_field) {
    return *row_entry.data_field;
  }
  return col_entry.data_field.value_or(0);
}

// Under `gridDropZones` a lone data field without column fields, or without
// row fields, moves its name to the report's top-left corner; its value
// column reads the locale's standalone total word instead.
bool lone_data_field_in_drop_zone(const LayoutPlan& plan) {
  return plan.table.grid_drop_zones() && !plan.legacy && plan.data_field_count == 1 &&
         (plan.col_depth == 0 || plan.row_depth == 0);
}

// The locale's bare "total" word: the subtotal affix without its separator.
std::string total_word(const LayoutPlan& plan) {
  const bool prefixed = plan.options.subtotal_suffix.empty() && !plan.options.subtotal_prefix.empty();
  const std::string& affix = prefixed ? plan.options.subtotal_prefix : plan.subtotal_suffix;
  const std::size_t first = affix.find_first_not_of(' ');
  if (first == std::string::npos) {
    return std::string();
  }
  return affix.substr(first, affix.find_last_not_of(' ') - first + 1);
}

// Label of a column with no column level above it: its data field's name,
// or nothing when the data fields sit on the rows.
std::string flat_col_header(const LayoutPlan& plan, std::size_t c_entry) {
  const AxisEntry& entry = plan.col_entries[c_entry];
  if (lone_data_field_in_drop_zone(plan)) {
    return total_word(plan);
  }
  if (entry.data_field) {
    return data_field_name(plan.table, *entry.data_field);
  }
  return plan.data_field_count == 1 ? data_field_name(plan.table, 0) : std::string();
}

}  // namespace

// The page block: field name, then the item it is showing. Both are
// header furniture, so neither earns a `PivotCellKind` of its own — the
// vocabulary is mirrored by the C ABI and the WASM enum, and a page cell
// is not something a renderer has to draw differently from a header.
// Every other cell of the block, the separator row included, is an explicit
// blank so the rendered extent is the rectangle Excel reports.
void emit_page_block(const LayoutPlan& plan, PivotCells& cells) {
  const std::vector<PivotPageSelection>& pages = plan.result.page_selections;
  if (plan.page_rows == 0) {
    return;
  }
  const std::size_t stride = kPageHeaderCols + 1;
  for (std::size_t r = 0; r < plan.page_rows; ++r) {
    const std::uint32_t row = cells.top + static_cast<std::uint32_t>(r);
    for (std::size_t c = 0; c < plan.total_cols; ++c) {
      const std::uint32_t col = plan.left + static_cast<std::uint32_t>(c);
      const std::size_t line = plan.table.page_over_then_down() ? r : c / stride;
      const std::size_t pos = plan.table.page_over_then_down() ? c / stride : r;
      const std::size_t i = line * plan.page_lines + pos;
      if (r + 1 == plan.page_rows || c % stride == kPageHeaderCols || pos >= plan.page_lines || i >= pages.size()) {
        append_blank(cells, row, col);
      } else if (c % stride == 0) {
        emit_text_or_blank(cells, row, col, pages[i].field_label, PivotCellKind::Header, PivotCellRole::PageCaption, 0,
                           pages[i].field_label);
      } else {
        append_cell(cells, row, col, text_value(cells, pages[i].item_label), PivotCellKind::Header,
                    PivotCellRole::PageValue, 0, pages[i].field_label);
      }
    }
  }
}

namespace {

// Emits one column-level label row at `depth`. Excel shows a non-leaf
// level's label once per contiguous run and blanks the repeats; the leaf-
// most level always shows its own value. A subtotal column carries its
// label on its own group's level and blanks below it. The grand-totals-
// rows strip is left blank except on `total_header_row`.
void emit_col_hierarchy_row(const LayoutPlan& plan, PivotCells& cells, std::uint32_t row, std::size_t depth,
                            bool blank_repeats) {
  const PivotTable& table = plan.table;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const bool leaf_depth = depth + 1 >= plan.col_levels;
  for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
    const AxisEntry& entry = col_entries[c_entry];
    // A collapsed item's columns are labelled like a leaf's.
    const bool subtotal_here =
        entry.kind == EntryKind::Subtotal && !entry.heads_group && depth + 1 == entry.labels.size();
    std::string label;
    if (depth < entry.labels.size()) {
      label = entry.labels[depth];
    }
    const PivotField* field = axis_level_field(plan.col_shape, depth);
    if (subtotal_here) {
      label = entry_subtotal_label(plan, entry, label);
    } else if (blank_repeats && !leaf_depth && c_entry > 0 && (field == nullptr || !field->fill_down_labels)) {
      const std::vector<std::string>& prev = col_entries[c_entry - 1].labels;
      if (prev.size() > depth && entry.labels.size() > depth &&
          std::equal(prev.begin(), prev.begin() + static_cast<std::ptrdiff_t>(depth + 1), entry.labels.begin())) {
        label.clear();
      }
    }
    const std::uint32_t col = plan.data_left + static_cast<std::uint32_t>(c_entry);
    if (!subtotal_here && plan.col_shape.values_at == depth && entry.data_field) {
      // The Values level heads its column with the data field's name.
      emit_text_or_blank(cells, row, col, label, PivotCellKind::Header, PivotCellRole::DataHeader,
                         static_cast<std::uint32_t>(depth), data_field_name(table, *entry.data_field),
                         data_field_format(table, *entry.data_field));
      continue;
    }
    emit_text_or_blank(cells, row, col, label,
                       entry.kind == EntryKind::Subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::ColLabel,
                       subtotal_here ? PivotCellRole::SubtotalLabel : PivotCellRole::ColItem,
                       static_cast<std::uint32_t>(depth), col_level_name(plan, depth),
                       subtotal_here ? std::string() : label_format(entry, depth));
    if (!subtotal_here && !label.empty()) {
      cells.cells.back().item_value = label_item_value(entry, depth);
    }
  }
  if (row != plan.total_header_row) {
    for (std::size_t s = 0; s < plan.total_strip_cols; ++s) {
      append_blank(cells, row, plan.total_left + static_cast<std::uint32_t>(s));
    }
  }
}

// The corner of a pivot with column fields names its data field; with
// several, the Values level names them instead and the corner is blank.
void emit_corner(const LayoutPlan& plan, PivotCells& cells) {
  if (plan.data_field_count == 1 && (plan.row_depth > 0 || lone_data_field_in_drop_zone(plan))) {
    append_df_cell(cells, plan.header_top, plan.left, text_value(cells, data_field_name(plan.table, 0)),
                   PivotCellKind::Header, PivotCellRole::FieldHeader, plan.table, 0);
  } else {
    append_blank(cells, plan.header_top, plan.left);
  }
}

// The `gridDropZones` row: the Values caption over the first data column when
// the data fields head the columns.
void emit_drop_zone_row(const LayoutPlan& plan, PivotCells& cells) {
  if (plan.drop_zone_rows == 0) {
    return;
  }
  const std::size_t width = plan.row_header_cols + plan.data_cols + plan.total_strip_cols;
  for (std::size_t c = 0; c < width; ++c) {
    const std::uint32_t col = plan.left + static_cast<std::uint32_t>(c);
    if (col == plan.left && lone_data_field_in_drop_zone(plan)) {
      append_df_cell(cells, plan.top, col, text_value(cells, data_field_name(plan.table, 0)), PivotCellKind::Header,
                     PivotCellRole::FieldHeader, plan.table, 0);
    } else if (col == plan.data_left && plan.values_on_cols) {
      emit_text_or_blank(cells, plan.top, col, field_caption(plan, plan.table.data_caption()), PivotCellKind::Header,
                         PivotCellRole::FieldHeader, 0);
    } else {
      append_blank(cells, plan.top, col);
    }
  }
}

void emit_header_multi_col(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t top = plan.header_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t data_cols = plan.data_cols;
  const std::size_t col_levels = plan.col_levels;
  if (plan.col_depth == 0) {
    // One header row: the row level names, then the data field names.
    for (std::size_t d = 0; d < row_header_cols; ++d) {
      emit_text_or_blank(cells, top, left + static_cast<std::uint32_t>(d), field_caption(plan, row_level_name(plan, d)),
                         PivotCellKind::Header, PivotCellRole::FieldHeader, static_cast<std::uint32_t>(d));
    }
    for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
      const std::size_t df = col_entries[c_entry].data_field.value_or(0);
      emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(plan, c_entry),
                         PivotCellKind::Header, PivotCellRole::DataHeader, 0, data_field_name(table, df),
                         data_field_format(table, df));
    }
  } else {
    // Caption row: the corner, then each column level's own name -- the
    // Values level under the stored caption -- where Compact draws its
    // single "Column Labels" placeholder.
    if (plan.caption_row) {
      emit_corner(plan, cells);
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_blank(cells, top, left + static_cast<std::uint32_t>(d));
      }
      for (std::size_t i = 0; i < plan.corner_extent; ++i) {
        emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(i),
                           field_caption(plan, col_level_name(plan, i)), PivotCellKind::Header,
                           PivotCellRole::FieldHeader, 0);
      }
    }
    // The last column-level row also carries the row level names.
    for (std::size_t depth = 0; depth < col_levels; ++depth) {
      const std::uint32_t row = plan.col_rows_top + static_cast<std::uint32_t>(depth);
      for (std::size_t d = 0; d < row_header_cols; ++d) {
        emit_text_or_blank(cells, row, left + static_cast<std::uint32_t>(d),
                           depth + 1 == col_levels ? field_caption(plan, row_level_name(plan, d)) : std::string(),
                           PivotCellKind::Header, PivotCellRole::FieldHeader, static_cast<std::uint32_t>(d));
      }
      emit_col_hierarchy_row(plan, cells, row, depth, true);
    }
  }
}

void emit_header_compact(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotLayoutOptions& options = plan.options;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t top = plan.header_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t data_cols = plan.data_cols;
  const std::size_t col_levels = plan.col_levels;
  if (plan.col_depth == 0) {
    // One header row: the "Row Labels" placeholder, then the data field
    // names (blank when the data fields sit on the rows). Without row
    // fields under `gridDropZones` the placeholder gives way to the corner.
    if (plan.grid_without_rows) {
      emit_corner(plan, cells);
    } else if (row_header_cols > 0) {
      emit_text_or_blank(cells, top, left, field_caption(plan, options.row_labels_label), PivotCellKind::Header,
                         PivotCellRole::FieldHeader, 0);
    }
    for (std::size_t c_entry = 0; c_entry < data_cols; ++c_entry) {
      const std::size_t df = col_entries[c_entry].data_field.value_or(0);
      emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(c_entry), flat_col_header(plan, c_entry),
                         PivotCellKind::Header, PivotCellRole::DataHeader, 0, data_field_name(table, df),
                         data_field_format(table, df));
    }
  } else {
    // Caption row: the corner, then the "Column Labels" placeholder over
    // the first column. Without row fields the single data field's name
    // moves down to the data row instead (see the row loop below). Without
    // row or data fields, Outline and Tabular name the column field there.
    if (plan.caption_row) {
      if (row_header_cols > 0) {
        emit_corner(plan, cells);
      }
      const std::string caption = plan.valueless && plan.row_levels == 0 && table.layout() != PivotLayout::Compact
                                      ? col_level_name(plan, 0)
                                      : options.column_labels_label;
      for (std::size_t i = 0; i < plan.corner_extent; ++i) {
        emit_text_or_blank(cells, top, data_left + static_cast<std::uint32_t>(i),
                           i == 0 ? field_caption(plan, caption) : std::string(), PivotCellKind::Header,
                           PivotCellRole::FieldHeader, 0);
      }
    }
    for (std::size_t depth = 0; depth < col_levels; ++depth) {
      const std::uint32_t row = plan.col_rows_top + static_cast<std::uint32_t>(depth);
      if (row_header_cols > 0 && (plan.row_levels == 0 || depth + 1 < col_levels)) {
        append_blank(cells, row, left);
      }
      emit_col_hierarchy_row(plan, cells, row, depth, true);
    }
    if (row_header_cols > 0 && plan.row_levels > 0) {
      emit_text_or_blank(cells, plan.row_header_row, left, field_caption(plan, options.row_labels_label),
                         PivotCellKind::Header, PivotCellRole::FieldHeader, 0);
    }
  }
}

void emit_header_legacy(const LayoutPlan& plan, PivotCells& cells) {
  // English projection: column levels stack from the top, each label
  // repeated per column, over a row of row level names.
  for (std::size_t depth = 0; depth < plan.row_header_cols; ++depth) {
    emit_text_or_blank(cells, plan.row_header_row, plan.left + static_cast<std::uint32_t>(depth),
                       field_caption(plan, row_level_name(plan, depth)), PivotCellKind::Header,
                       PivotCellRole::FieldHeader, static_cast<std::uint32_t>(depth));
  }
  if (plan.col_levels == 0) {
    for (std::size_t c_entry = 0; c_entry < plan.data_cols; ++c_entry) {
      emit_text_or_blank(cells, plan.top, plan.data_left + static_cast<std::uint32_t>(c_entry),
                         flat_col_header(plan, c_entry), PivotCellKind::ColLabel, PivotCellRole::DataHeader, 0);
    }
  }
  for (std::size_t depth = 0; depth < plan.col_levels; ++depth) {
    emit_col_hierarchy_row(plan, cells, plan.top + static_cast<std::uint32_t>(depth), depth, false);
  }
}

}  // namespace

// Column hierarchy rows plus the row-level header, per layout mode.
void emit_header_block(const LayoutPlan& plan, PivotCells& cells) {
  emit_drop_zone_row(plan, cells);
  if (plan.multi_col_layout) {
    emit_header_multi_col(plan, cells);
  } else if (plan.compact) {
    emit_header_compact(plan, cells);
  } else {
    emit_header_legacy(plan, cells);
  }
}

// Row labels and data cells. Tabular tracks the previous leaf row's label
// so a parent path shared with the leaf row above is blanked, unless the
// parent's field fills its label down.
void emit_body_rows(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::uint32_t left = plan.left;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t data_left = plan.data_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const bool multi_col_layout = plan.multi_col_layout;
  const bool outline = plan.outline;
  const bool compact = plan.compact;
  const bool tabular = plan.tabular;
  std::vector<std::string> prev_row_labels(row_header_cols);
  std::vector<bool> prev_row_labels_filled(row_header_cols, false);
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const AxisEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    if (entry.kind == EntryKind::Blank) {
      // Compact form writes the label cell of a blank row; the other layouts leave it out.
      for (std::size_t c = 0; c < row_header_cols; ++c) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(c), Value::blank(), PivotCellKind::Blank,
                    compact ? PivotCellRole::BlankRow : PivotCellRole::None, 0);
        if (compact) {
          set_compact_indent(plan, cells, 0);
        }
      }
      append_empty_value_cells(cells, row, data_left, col_entries.size());
      continue;
    }
    const std::vector<std::string>& labels = entry.labels;
    const bool subtotal = entry.kind == EntryKind::Subtotal;
    const std::size_t last_depth = labels.empty() ? 0 : labels.size() - 1;
    // Leading levels this row shares with the previous leaf row; a level
    // under a changed outer item starts a new block and shows its label.
    std::size_t shared = 0;
    while (shared < labels.size() && shared < row_header_cols && prev_row_labels_filled[shared] &&
           prev_row_labels[shared] == labels[shared]) {
      ++shared;
    }
    for (std::size_t depth = 0; depth < row_header_cols; ++depth) {
      std::string label;
      // The level whose item `label` shows as is, for its number format.
      std::optional<std::size_t> bare_level;
      PivotCellRole role = PivotCellRole::SubtotalLabel;
      if (multi_col_layout) {
        const PivotField* field = axis_level_field(plan.row_shape, depth);
        const bool fill_down = field != nullptr && field->fill_down_labels && depth < last_depth;
        if (subtotal || entry.kind == EntryKind::Group || outline) {
          // Subtotal and group rows, and every outline row, label only
          // their own level and the fill-down levels above it. A subtotal
          // takes the localized wording unless it doubles as the group's
          // heading row.
          if (depth == last_depth && depth < labels.size()) {
            label = labels[depth];
            if (subtotal && !entry.heads_group) {
              label = entry_subtotal_label(plan, entry, label);
            } else {
              bare_level = depth;
            }
          } else if (fill_down) {
            label = labels[depth];
            bare_level = depth;
          }
        } else if (depth < labels.size()) {
          // Tabular leaf rows repeat the path but blank a parent whose
          // whole path matches the previous leaf row's; the leaf-most level
          // always shows.
          const std::string& candidate = labels[depth];
          if (fill_down || depth + 1 == labels.size() || depth >= shared) {
            label = candidate;
            bare_level = depth;
          }
        }
      } else if (compact) {
        // Compact form shows each row's own level in the single label
        // column; a subtotal heading its group is the bare group label.
        if (subtotal && !entry.subtotal_label.empty()) {
          label = entry.subtotal_label;
        } else if (subtotal && !entry.heads_group && !labels.empty()) {
          label = entry_subtotal_label(plan, entry, labels.back());
        } else if (!labels.empty()) {
          label = labels.back();
          bare_level = last_depth;
        } else if (r_entry == 0 && plan.grid_without_rows) {
          // Under `gridDropZones` the single data row reads the locale's total word.
          label = total_word(plan);
          role = PivotCellRole::FieldHeader;
        } else if (plan.row_levels == 0 && plan.col_depth > 0 && r_entry == 0) {
          // No-row-fields pivot: the data field's name heads the single
          // data row.
          label = data_field_name(table, 0);
          role = PivotCellRole::FieldHeader;
        }
      } else if (depth < labels.size()) {
        label = labels[depth];
        if (subtotal && depth == last_depth && !entry.heads_group) {
          label = entry_subtotal_label(plan, entry, label);
        } else {
          bare_level = depth;
        }
      }
      const PivotCellKind kind = subtotal ? PivotCellKind::RowSubtotal : PivotCellKind::RowLabel;
      // Compact folds every level into one column, so the cell reports the
      // row's hierarchy level rather than its physical column.
      const std::size_t level = compact ? last_depth : depth;
      if (bare_level.has_value()) {
        role = plan.row_shape.values_at == *bare_level ? PivotCellRole::DataRowLabel : PivotCellRole::RowItem;
      }
      if (label.empty()) {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), Value::blank(), kind, PivotCellRole::None,
                    static_cast<std::uint32_t>(level), row_level_name(plan, level));
      } else {
        append_cell(cells, row, left + static_cast<std::uint32_t>(depth), text_value(cells, label), kind, role,
                    static_cast<std::uint32_t>(level), row_level_name(plan, level),
                    bare_level.has_value() ? label_format(entry, *bare_level) : std::string());
        if (bare_level.has_value()) {
          cells.cells.back().item_value = label_item_value(entry, *bare_level);
        }
        if (compact && role != PivotCellRole::FieldHeader) {
          set_compact_indent(plan, cells, level);
        }
      }
      if (tabular && entry.kind == EntryKind::Leaf && depth < labels.size()) {
        prev_row_labels[depth] = labels[depth];
        prev_row_labels_filled[depth] = true;
      }
    }
    if (entry.kind == EntryKind::Group || plan.valueless) {
      // A heading row without its subtotal, or a report without data fields, spans the data area as blanks.
      append_empty_value_cells(cells, row, data_left, col_entries.size());
      continue;
    }
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& col_entry = col_entries[c_entry];
      const std::size_t df = cell_data_field(entry, col_entry);
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      const Value* value = nullptr;
      PivotCellKind kind = PivotCellKind::Data;
      if (subtotal) {
        const RowSubtotal& row_subtotal = result.row_subtotals[entry.subtotal_index];
        kind = PivotCellKind::RowSubtotal;
        if (col_entry.kind == EntryKind::Subtotal &&
            col_entry.subtotal_index < row_subtotal.col_subtotal_values.size() &&
            df < row_subtotal.col_subtotal_values[col_entry.subtotal_index].size()) {
          value = &row_subtotal.col_subtotal_values[col_entry.subtotal_index][df];
        } else if (col_entry.kind == EntryKind::Leaf && col_entry.leaf_index < row_subtotal.col_values.size() &&
                   df < row_subtotal.col_values[col_entry.leaf_index].size()) {
          value = &row_subtotal.col_values[col_entry.leaf_index][df];
        } else if (df < row_subtotal.values.size()) {
          value = &row_subtotal.values[df];
        }
      } else if (col_entry.kind == EntryKind::Subtotal) {
        const ColSubtotal& col_subtotal = result.col_subtotals[col_entry.subtotal_index];
        kind = PivotCellKind::ColSubtotal;
        if (entry.leaf_index < col_subtotal.values.size() && df < col_subtotal.values[entry.leaf_index].size()) {
          value = &col_subtotal.values[entry.leaf_index][df];
        }
      } else {
        value = &result.values[entry.leaf_index][col_entry.leaf_index][df];
      }
      append_value_cell(cells, row, col, value != nullptr ? *value : Value::blank(), kind, table, df);
    }
  }
}

// Row totals: the strip at the right, from the evaluator's re-aggregated
// per-row totals so the projection remains cache-only.
void emit_row_totals_strip(const LayoutPlan& plan, PivotCells& cells) {
  if (!plan.emit_grand_totals_rows_strip) {
    return;
  }
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisLeaf>& col_leaves = plan.col_leaves;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t total_left = plan.total_left;
  const std::size_t total_strip_cols = plan.total_strip_cols;
  const bool values_on_cols = plan.values_on_cols;
  for (std::size_t r_entry = 0; r_entry < row_entries.size(); ++r_entry) {
    const AxisEntry& entry = row_entries[r_entry];
    const std::uint32_t row = data_top + static_cast<std::uint32_t>(r_entry);
    if (entry.kind == EntryKind::Group || entry.kind == EntryKind::Blank || plan.valueless) {
      append_empty_value_cells(cells, row, total_left, total_strip_cols);
      continue;
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = entry.data_field.value_or(values_on_cols ? s : 0);
      const std::uint32_t col = total_left + static_cast<std::uint32_t>(s);
      if (entry.kind == EntryKind::Subtotal) {
        // `RowSubtotal::values[df]` aggregates the group across the whole
        // column axis, which is exactly what this strip states.
        const RowSubtotal& subtotal = result.row_subtotals[entry.subtotal_index];
        append_value_cell(cells, row, col, df < subtotal.values.size() ? subtotal.values[df] : Value::blank(),
                          PivotCellKind::RowSubtotal, table, df);
        continue;
      }
      // Use the evaluator's per-row-leaf re-aggregation rather than
      // summing this row's cells: a non-additive function (Average/Max/
      // Min/StdDev/Var) cannot be recovered from the cell aggregates.
      const std::size_t leaf = entry.leaf_index;
      if (leaf < result.row_leaf_totals.size() && df < result.row_leaf_totals[leaf].size() &&
          !result.row_leaf_totals[leaf][df].is_blank()) {
        append_value_cell(cells, row, col, result.row_leaf_totals[leaf][df], PivotCellKind::GrandTotal, table, df);
        continue;
      }
      if (total_is_derived(table, df)) {
        append_value_cell(cells, row, col, Value::blank(), PivotCellKind::GrandTotal, table, df);
        continue;
      }
      // Every row of this strip states a total: a source error wins, an
      // all-numeric row sums, and an aggregate that cannot be summed
      // reports blank rather than leaving a hole in the total column.
      StripTotal strip;
      for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
        if (!strip.add(result.values[leaf][c_leaf][df])) {
          break;
        }
      }
      append_value_cell(cells, row, col, strip.value(), PivotCellKind::GrandTotal, table, df);
    }
  }
  for (std::size_t s = 0; s < total_strip_cols; ++s) {
    const std::size_t df = values_on_cols ? s : 0;
    append_df_cell(cells, plan.total_header_row, total_left + static_cast<std::uint32_t>(s),
                   text_value(cells, values_on_cols ? data_total_label(plan, s) : plan.grand_total_label),
                   PivotCellKind::Header, PivotCellRole::GrandTotalLabel, table, df);
  }
}

// Column totals: the strip at the bottom.
void emit_col_totals_strip(const LayoutPlan& plan, PivotCells& cells) {
  const PivotTable& table = plan.table;
  const PivotResult& result = plan.result;
  const std::vector<AxisEntry>& row_entries = plan.row_entries;
  const std::vector<AxisEntry>& col_entries = plan.col_entries;
  const std::vector<AxisLeaf>& row_leaves = plan.row_leaves;
  const std::vector<AxisLeaf>& col_leaves = plan.col_leaves;
  const std::uint32_t left = plan.left;
  const std::uint32_t data_top = plan.data_top;
  const std::uint32_t data_left = plan.data_left;
  const std::uint32_t total_left = plan.total_left;
  const std::size_t row_header_cols = plan.row_header_cols;
  const std::size_t total_strip_cols = plan.total_strip_cols;
  for (std::size_t b = 0; b < plan.total_strip_rows; ++b) {
    const std::optional<std::size_t> row_df = plan.values_on_rows ? std::optional<std::size_t>(b) : std::nullopt;
    const std::uint32_t total_row = data_top + static_cast<std::uint32_t>(row_entries.size() + b);
    append_cell(cells, total_row, left, text_value(cells, row_df ? data_total_label(plan, b) : plan.grand_total_label),
                PivotCellKind::GrandTotal, PivotCellRole::GrandTotalLabel, 0);
    if (plan.compact) {
      set_compact_indent(plan, cells, 0);
    }
    // Tabular / Outline leave the remaining row-label columns blank so the
    // rendered grid is rectangular.
    if (plan.multi_col_layout) {
      for (std::size_t d = 1; d < row_header_cols; ++d) {
        append_cell(cells, total_row, left + static_cast<std::uint32_t>(d), Value::blank(), PivotCellKind::GrandTotal,
                    PivotCellRole::None, static_cast<std::uint32_t>(d));
      }
    }
    if (plan.valueless) {
      append_empty_value_cells(cells, total_row, data_left, col_entries.size() + total_strip_cols);
      continue;
    }
    for (std::size_t c_entry = 0; c_entry < col_entries.size(); ++c_entry) {
      const AxisEntry& col_entry = col_entries[c_entry];
      const std::size_t df = row_df.value_or(col_entry.data_field.value_or(0));
      const std::uint32_t col = data_left + static_cast<std::uint32_t>(c_entry);
      const bool col_subtotal = col_entry.kind == EntryKind::Subtotal;
      // Prefer the evaluator's re-aggregated totals: summing per-row
      // aggregates is wrong for Average, Max, Min, StdDev and Var.
      if (plan.col_depth == 0 && !col_subtotal && df < result.grand_totals.size() &&
          !result.grand_totals[df].is_blank()) {
        append_value_cell(cells, total_row, col, result.grand_totals[df], PivotCellKind::GrandTotal, table, df);
        continue;
      }
      if (!col_subtotal && col_entry.leaf_index < result.col_leaf_totals.size() &&
          df < result.col_leaf_totals[col_entry.leaf_index].size() &&
          !result.col_leaf_totals[col_entry.leaf_index][df].is_blank()) {
        append_value_cell(cells, total_row, col, result.col_leaf_totals[col_entry.leaf_index][df],
                          PivotCellKind::GrandTotal, table, df);
        continue;
      }
      const PivotCellKind kind = col_subtotal ? PivotCellKind::ColSubtotal : PivotCellKind::GrandTotal;
      if (total_is_derived(table, df)) {
        const ColSubtotal* subtotal = col_subtotal ? &result.col_subtotals[col_entry.subtotal_index] : nullptr;
        append_value_cell(cells, total_row, col,
                          subtotal != nullptr && df < subtotal->total.size() ? subtotal->total[df] : Value::blank(),
                          kind, table, df);
        continue;
      }
      // Grand-total row x subtotal column: non-Sum takes the evaluator's 2-D combination.
      if (col_subtotal && df < table.data_fields().size()) {
        const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
        const Aggregation aggregation =
            subtotal.aggregation.has_value() ? *subtotal.aggregation : table.data_fields()[df].aggregation;
        if ((aggregation != Aggregation::Sum || result.sum_totals_as_state) && df < subtotal.total.size() &&
            !subtotal.total[df].is_blank()) {
          append_value_cell(cells, total_row, col, subtotal.total[df], kind, table, df);
          continue;
        }
      }
      StripTotal strip;
      for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
        const Value* value = nullptr;
        if (col_subtotal) {
          const ColSubtotal& subtotal = result.col_subtotals[col_entry.subtotal_index];
          if (r_leaf < subtotal.values.size() && df < subtotal.values[r_leaf].size()) {
            value = &subtotal.values[r_leaf][df];
          }
        } else {
          value = &result.values[r_leaf][col_entry.leaf_index][df];
        }
        if (value == nullptr) {
          continue;
        }
        if (!strip.add(*value)) {
          break;
        }
      }
      append_value_cell(cells, total_row, col, strip.value(), kind, table, df);
    }
    for (std::size_t s = 0; s < total_strip_cols; ++s) {
      const std::size_t df = row_df.value_or(plan.values_on_cols ? s : 0);
      Value v = Value::blank();
      if (df < result.grand_totals.size() && !result.grand_totals[df].is_blank()) {
        v = result.grand_totals[df];
      } else if (df == 0 && !result.grand_total.is_blank()) {
        v = result.grand_total;
      } else if (!total_is_derived(table, df)) {
        double sum = 0.0;
        bool numeric = true;
        bool any_number = false;
        for (std::size_t r_leaf = 0; r_leaf < row_leaves.size(); ++r_leaf) {
          for (std::size_t c_leaf = 0; c_leaf < col_leaves.size(); ++c_leaf) {
            const Value& cell = result.values[r_leaf][c_leaf][df];
            if (cell.is_number()) {
              sum += cell.as_number();
              any_number = true;
            } else if (!cell.is_blank()) {
              numeric = false;
            }
          }
        }
        if (numeric && any_number) {
          v = Value::number(sum);
        }
      }
      append_value_cell(cells, total_row, total_left + static_cast<std::uint32_t>(s), v, PivotCellKind::GrandTotal,
                        table, df);
    }
  }
}

namespace {

// True when row entry `entry` lies inside the block of `node`, the item at
// level `level`, below the item's own label row; its own subtotal does not.
bool inside_block(const AxisEntry& entry, std::size_t level, const AxisHierarchyNode* node) {
  return entry.nodes.size() > level + 1 && entry.nodes[level] == node;
}

}  // namespace

// `mergeItem` in tabular form: an outer item's label spans the rows of its
// block when it has two or more, and a subtotal or grand-total label spans
// the label columns from its own to the last.
void emit_merges(const LayoutPlan& plan, PivotCells& cells) {
  if (!plan.tabular || !plan.table.merge_item()) {
    return;
  }
  const std::vector<AxisEntry>& entries = plan.row_entries;
  const std::size_t width = plan.row_header_cols;
  const auto col_at = [&](std::size_t c) { return plan.left + static_cast<std::uint32_t>(c); };
  const auto row_at = [&](std::size_t r) { return plan.data_top + static_cast<std::uint32_t>(r); };
  for (std::size_t r = 0; r < entries.size(); ++r) {
    const AxisEntry& entry = entries[r];
    for (std::size_t level = 0; level + 1 < entry.nodes.size() && level < width; ++level) {
      const AxisHierarchyNode* node = entry.nodes[level];
      if (node == nullptr || (r > 0 && inside_block(entries[r - 1], level, node))) {
        continue;
      }
      std::size_t end = r + 1;
      while (end < entries.size() && inside_block(entries[end], level, node)) {
        ++end;
      }
      if (end - r >= 2) {
        cells.merges.push_back(PivotRect{row_at(r), col_at(level), row_at(end - 1), col_at(level)});
      }
    }
    const std::size_t own = entry.labels.empty() ? 0 : entry.labels.size() - 1;
    if (entry.kind == EntryKind::Subtotal && !entry.heads_group && own + 1 < width) {
      cells.merges.push_back(PivotRect{row_at(r), col_at(own), row_at(r), col_at(width - 1)});
    }
  }
  for (std::size_t b = 0; width >= 2 && b < plan.total_strip_rows; ++b) {
    cells.merges.push_back(
        PivotRect{row_at(entries.size() + b), col_at(0), row_at(entries.size() + b), col_at(width - 1)});
  }
}

}  // namespace formulon::pivot::detail
