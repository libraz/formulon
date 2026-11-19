//
// Same-sheet Cut / Paste and Copy / Paste of a cell rectangle for `Workbook`.
//
// Both operations are restricted by a preflight to plain content: values and
// formulas made of unqualified single-cell references, moved onto an empty,
// unformatted, object-free destination. Every check runs before the first
// write, so the mutation stage cannot fail and a refused call changes
// nothing. A destination holding any object is refused, so no object is ever
// deleted by a move.

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <initializer_list>
#include <mutex>
#include <optional>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "auto_filter.h"
#include "cell.h"
#include "cf/cf_types.h"
#include "defined_name.h"
#include "eval/recalc_engine.h"
#include "io/ooxml_defs.h"
#include "merge_range.h"
#include "parser/ast.h"
#include "parser/ast_shift.h"
#include "parser/parser.h"
#include "parser/ref_transforms.h"
#include "parser/reference.h"
#include "pivot/pivot_cache.h"
#include "row_cells.h"
#include "sheet.h"
#include "sheet_layout.h"
#include "sheet_metadata.h"
#include "sheet_view.h"
#include "table.h"
#include "threaded_comment.h"
#include "unknown_relationship.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/formula_text.h"
#include "utils/status_macros.h"
#include "value.h"
#include "workbook.h"
#include "workbook_sheet_mutation.h"

namespace formulon {

namespace {

using parser::AstNode;
using parser::NodeKind;

constexpr std::uint32_t kLastRow = Sheet::kMaxRows - 1U;
constexpr std::uint32_t kLastCol = Sheet::kMaxCols - 1U;

enum class Op : std::uint8_t { kMove, kCopy };

// The source and destination rectangles of one operation.
struct Rects {
  MergeRange src;
  MergeRange dst;
};

// One source cell with its destination, formulas already rewritten.
struct SourceCell {
  std::uint32_t row = 0;
  std::uint32_t col = 0;
  std::uint32_t dest_row = 0;
  std::uint32_t dest_col = 0;
  bool is_formula = false;
  bool is_text = false;
  // Formula text for a formula, the string for a text value.
  std::string text;
  Value value = Value::blank();
};

// The references a formula may hold: when `inside` is set (move), unqualified
// single cells inside it; otherwise (copy) unqualified single cells that stay
// in the grid after a shift of (`row_delta`, `col_delta`).
struct SourceRule {
  const MergeRange* inside = nullptr;
  std::int64_t row_delta = 0;
  std::int64_t col_delta = 0;
};

// Where an unqualified reference resolves (`home`) and the sheet whose
// rectangles are collected (`target`).
struct RefScope {
  const Workbook& wb;
  std::size_t home;
  std::size_t target;
};

const char* VerbOf(Op op) {
  return op == Op::kMove ? "moving" : "copying";
}

Error Unsupported(Op op, const std::string& what) {
  return make_error(FormulonErrorCode::kNotImplemented, std::string(VerbOf(op)) + " " + what + " is not supported");
}

Error Touching(Op op, const char* object) {
  return Unsupported(op, std::string("cells touching ") + object);
}

MergeRange CellRect(std::uint32_t row, std::uint32_t col) {
  return MergeRange{row, col, row, col};
}

bool Hits(const MergeRange& r, const Rects& m) {
  return intersects(r, m.src) || intersects(r, m.dst);
}

// ---------------------------------------------------------------------------
// Reference collection
// ---------------------------------------------------------------------------

MergeRange RefRect(const parser::Reference& ref) {
  if (ref.is_full_col) {
    return MergeRange{0U, ref.col, kLastRow, ref.col};
  }
  if (ref.is_full_row) {
    return MergeRange{ref.row, 0U, ref.row, kLastCol};
  }
  return CellRect(ref.row, ref.col);
}

bool OnTarget(std::string_view sheet, const RefScope& scope) {
  return sheet.empty() ? scope.home == scope.target : scope.wb.sheet_index_by_name(sheet) == scope.target;
}

void CollectRects(const AstNode& node, const RefScope& scope, std::vector<MergeRange>& out);

// A `:` chain reads as one box over its operands; a bare operand takes the
// qualifier of the first qualified one.
void CollectRangeOp(const AstNode& node, const RefScope& scope, std::vector<MergeRange>& out) {
  std::string_view owner;
  for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
    const AstNode& operand = node.as_range_operand(i);
    if (operand.kind() == NodeKind::Ref && !operand.as_ref().sheet.empty()) {
      owner = operand.as_ref().sheet;
      break;
    }
  }
  std::vector<MergeRange> parts;
  for (std::uint32_t i = 0; i < node.as_range_arity(); ++i) {
    const AstNode& operand = node.as_range_operand(i);
    if (operand.kind() == NodeKind::Ref) {
      const parser::Reference& ref = operand.as_ref();
      if (OnTarget(ref.sheet.empty() ? owner : ref.sheet, scope)) {
        parts.push_back(RefRect(ref));
      }
    } else {
      CollectRects(operand, scope, parts);
    }
  }
  if (parts.empty()) {
    return;
  }
  MergeRange box = parts.front();
  for (const MergeRange& part : parts) {
    box.first_row = std::min(box.first_row, part.first_row);
    box.first_col = std::min(box.first_col, part.first_col);
    box.last_row = std::max(box.last_row, part.last_row);
    box.last_col = std::max(box.last_col, part.last_col);
  }
  out.push_back(box);
}

// Appends every rectangle of `scope.target` the formula can read through a
// written-out reference. A 3-D reference whose span holds the target counts
// as the whole sheet; names and structured references name no rectangle here
// (a name's body and a table's range are checked on their own).
void CollectRects(const AstNode& node, const RefScope& scope, std::vector<MergeRange>& out) {
  switch (node.kind()) {
    case NodeKind::Ref:
      if (OnTarget(node.as_ref().sheet, scope)) {
        out.push_back(RefRect(node.as_ref()));
      }
      return;
    case NodeKind::SpillRef:
      if (const AstNode* anchor = node.as_spill_ref_anchor_expr(); anchor != nullptr) {
        CollectRects(*anchor, scope, out);
      } else if (OnTarget(node.as_spill_ref().sheet, scope)) {
        out.push_back(RefRect(node.as_spill_ref()));
      }
      return;
    case NodeKind::Ref3D: {
      const std::size_t first = scope.wb.sheet_index_by_name(node.as_ref3d_sheet_begin());
      const std::size_t last = scope.wb.sheet_index_by_name(node.as_ref3d_sheet_end());
      const bool unresolved = first == Workbook::kNoSheet || last == Workbook::kNoSheet;
      if (unresolved || (scope.target >= std::min(first, last) && scope.target <= std::max(first, last))) {
        out.push_back(MergeRange{0U, 0U, kLastRow, kLastCol});
      }
      return;
    }
    case NodeKind::RangeOp:
      CollectRangeOp(node, scope, out);
      return;
    case NodeKind::Literal:
    case NodeKind::ErrorLiteral:
    case NodeKind::ErrorPlaceholder:
    case NodeKind::NameRef:
    case NodeKind::StructuredRef:
    case NodeKind::ExternalRef:
      return;
    default:
      break;
  }
  parser::any_child_node(node, [&](const AstNode& child) {
    CollectRects(child, scope, out);
    return false;
  });
}

// The tree of `text`, or null when it does not parse; `arena` owns it.
const AstNode* ParseFormula(std::string_view text, Arena& arena, bool name_body) {
  const std::string_view body = strip_formula_prefix(text);
  const AstNode* root =
      name_body ? parser::parse_strict(body, arena, parser::name_body_options()) : parser::parse_strict(body, arena);
  if (root == nullptr || !parser::ast_depth_within_limit(*root, parser::kMaxFormulaAstDepth)) {
    return nullptr;
  }
  return root;
}

// Describes the first part of `node` outside what a moved or copied formula
// may hold under `rule`, or null when it is all supported.
const char* UnsupportedShape(const AstNode& node, const SourceRule& rule) {
  switch (node.kind()) {
    case NodeKind::Literal:
      return nullptr;
    case NodeKind::ErrorLiteral:
      return node.as_ref_error_site() == nullptr ? nullptr : "formulas with deleted references";
    case NodeKind::Ref: {
      const parser::Reference& ref = node.as_ref();
      if (!ref.sheet.empty()) {
        return "formulas with sheet-qualified references";
      }
      if (ref.is_full_col || ref.is_full_row) {
        return "formulas with whole row or column references";
      }
      if (rule.inside != nullptr) {
        if (!contains(*rule.inside, CellRect(ref.row, ref.col))) {
          return "formulas pointing outside the moved cells";
        }
      } else {
        const std::int64_t col = static_cast<std::int64_t>(ref.col) + (ref.col_abs ? 0 : rule.col_delta);
        const std::int64_t row = static_cast<std::int64_t>(ref.row) + (ref.row_abs ? 0 : rule.row_delta);
        if (col < 0 || col > static_cast<std::int64_t>(kLastCol) || row < 0 ||
            row > static_cast<std::int64_t>(kLastRow)) {
          return "formulas whose references leave the grid";
        }
      }
      return nullptr;
    }
    case NodeKind::RangeOp:
    case NodeKind::UnionOp:
    case NodeKind::IntersectOp:
      return "formulas with range references";
    case NodeKind::Ref3D:
      return "formulas with 3-D references";
    case NodeKind::ExternalRef:
      return "formulas with external references";
    case NodeKind::StructuredRef:
      return "formulas with table references";
    case NodeKind::NameRef:
      return "formulas with defined names";
    case NodeKind::SpillRef:
      return "formulas with spilled-range references";
    case NodeKind::Lambda:
    case NodeKind::LetBinding:
    case NodeKind::LambdaCall:
      return "formulas with LET or LAMBDA";
    case NodeKind::ErrorPlaceholder:
      return "formulas that do not parse";
    case NodeKind::Call:
      if (parser::is_dynamic_reference_function_name(node.as_call_name())) {
        return "formulas that build references";
      }
      break;
    default:
      break;
  }
  const char* found = nullptr;
  parser::any_child_node(node, [&](const AstNode& child) {
    found = UnsupportedShape(child, rule);
    return found != nullptr;
  });
  return found;
}

// Parses `text` and, when `check_hits`, refuses a reference into either
// rectangle of `m` on `target`.
Expected<void, Error> CheckFormulaRefs(const Workbook& wb, std::string_view text, std::size_t home, std::size_t target,
                                       const Rects& m, Op op, bool name_body, bool check_hits) {
  Arena arena;
  const AstNode* root = ParseFormula(text, arena, name_body);
  if (root == nullptr) {
    return Unsupported(op, "cells in a workbook with formulas that do not parse");
  }
  if (!check_hits) {
    return Expected<void, Error>::Ok();
  }
  std::vector<MergeRange> rects;
  CollectRects(*root, RefScope{wb, home, target}, rects);
  for (const MergeRange& rect : rects) {
    if (intersects(rect, m.src)) {
      return Unsupported(op, "cells that a formula, name or rule refers to");
    }
    if (intersects(rect, m.dst)) {
      return Unsupported(op, "cells onto a destination that a formula, name or rule refers to");
    }
  }
  return Expected<void, Error>::Ok();
}

// Move only: every formula of the workbook parses, and no reference outside
// the source rectangle (cell formulas, defined-name bodies, rule formulas on
// any sheet) reaches the source or the destination.
Expected<void, Error> CheckWorkbookReferences(const Workbook& wb, std::size_t sheet_index, const Rects& m, Op op) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    for (const CellAddress at : sheet.all_formula_cells()) {
      const Cell* cell = sheet.cell_at(at.row, at.col);
      if (cell == nullptr) {
        continue;
      }
      // A source formula is held to the stricter source rule instead.
      const bool in_source = s == sheet_index && contains(m.src, CellRect(at.row, at.col));
      RETURN_IF_ERROR(CheckFormulaRefs(wb, cell->formula_text, s, sheet_index, m, op, false, !in_source));
    }
    for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
      for (const cf::CFRule& rule : format.rules) {
        for (const std::optional<std::string>* formula : {&rule.formula1, &rule.formula2}) {
          if (formula->has_value() && !(*formula)->empty()) {
            RETURN_IF_ERROR(CheckFormulaRefs(wb, **formula, s, sheet_index, m, op, false, true));
          }
        }
      }
    }
    for (const DataValidation& validation : sheet.validations()) {
      for (const std::string* formula : {&validation.formula1, &validation.formula2}) {
        if (!formula->empty()) {
          RETURN_IF_ERROR(CheckFormulaRefs(wb, *formula, s, sheet_index, m, op, false, true));
        }
      }
    }
  }
  for (const DefinedName& name : wb.defined_names()) {
    if (!name.formula.empty()) {
      RETURN_IF_ERROR(CheckFormulaRefs(wb, name.formula, sheet_index, sheet_index, m, op, true, true));
    }
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// Sheet objects and formats
// ---------------------------------------------------------------------------

// True when a stored cell or a blank style run in `r` has a non-default xf.
bool HasOwnFormat(const Sheet& sheet, const MergeRange& r) {
  const std::uint64_t end_col = static_cast<std::uint64_t>(r.last_col) + 1U;
  for (const auto& entry : sheet.rows()) {
    if (entry.first < r.first_row || entry.first > r.last_row) {
      continue;
    }
    const auto runs = entry.second.runs_in(r.first_col, end_col);
    for (const StyleRun* run = runs.first; run != runs.second; ++run) {
      if (run->xf != 0U) {
        return true;
      }
    }
    for (const RowCells::Entry stored : entry.second.entries_in(r.first_col, end_col)) {
      if (stored.cell.xf_index != 0U) {
        return true;
      }
    }
  }
  return false;
}

// Names what makes the rows or columns of `r` unsupported: a non-default row
// or column xf, a custom size, or hidden / outlined rows and columns.
const char* AxisConflict(const SheetLayout& layout, const MergeRange& r) {
  for (const RowLayout& row : layout.row_overrides) {
    if (row.row < r.first_row || row.row > r.last_row) {
      continue;
    }
    if (row.has_style && row.style_xf != 0U) {
      return "cells in rows or columns that carry formats";
    }
    if (row.custom_height || row.hidden || row.outline_level != 0U) {
      return "cells in rows or columns with custom sizes or visibility";
    }
  }
  for (const ColumnLayout& col : layout.columns) {
    if (col.last < r.first_col || col.first > r.last_col) {
      continue;
    }
    if (col.has_style && col.style_xf != 0U) {
      return "cells in rows or columns that carry formats";
    }
    if (col.custom_width || col.hidden || col.outline_level != 0U) {
      return "cells in rows or columns with custom sizes or visibility";
    }
  }
  return nullptr;
}

// True when a space-separated A1 list holds a range touching `m`, or one that
// does not parse.
bool SqrefTouches(std::string_view sqref, const Rects& m) {
  std::size_t from = 0;
  while (from < sqref.size()) {
    std::size_t space = sqref.find(' ', from);
    if (space == std::string_view::npos) {
      space = sqref.size();
    }
    const std::string_view token = sqref.substr(from, space - from);
    from = space + 1U;
    if (token.empty()) {
      continue;
    }
    std::string plain;
    for (const char ch : token) {
      if (ch != '$') {
        plain.push_back(ch);
      }
    }
    const std::optional<MergeRange> rect = parse_a1_rectangle(plain);
    if (!rect.has_value() || Hits(*rect, m)) {
      return true;
    }
  }
  return false;
}

// A sheet name as written in a reference, without its quotes.
std::string Unquoted(std::string_view name) {
  if (name.size() < 2U || name.front() != '\'' || name.back() != '\'') {
    return std::string(name);
  }
  std::string out;
  for (std::size_t i = 1; i + 1U < name.size(); ++i) {
    out.push_back(name[i]);
    if (name[i] == '\'' && i + 2U < name.size() && name[i + 1U] == '\'') {
      ++i;
    }
  }
  return out;
}

// A PivotTable cache sourced from a range touching `m` on the sheet.
bool PivotSourceTouches(const Workbook& wb, std::size_t sheet_index, const Rects& m) {
  for (const auto& cache : wb.pivot_caches()) {
    if (cache == nullptr || !cache->worksheet_source().present) {
      continue;
    }
    const pivot::WorksheetSource& source = cache->worksheet_source();
    if (!source.name.empty()) {
      return true;
    }
    std::string_view ref = source.ref;
    std::string sheet_name = source.sheet;
    const std::size_t bang = ref.rfind('!');
    if (bang != std::string_view::npos) {
      sheet_name = Unquoted(ref.substr(0, bang));
      ref = ref.substr(bang + 1U);
    }
    if (sheet_name.empty()) {
      return true;
    }
    const std::size_t index = wb.sheet_index_by_name(sheet_name);
    if (index != sheet_index) {
      continue;
    }
    std::string plain;
    for (const char ch : ref) {
      if (ch != '$') {
        plain.push_back(ch);
      }
    }
    const std::optional<MergeRange> rect = parse_a1_rectangle(plain);
    if (!rect.has_value() || Hits(*rect, m)) {
      return true;
    }
  }
  return false;
}

// True when any sheet of the workbook anchors a drawing, chart or chartsheet.
bool WorkbookHoldsDrawings(const Workbook& wb) {
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    const Sheet& sheet = wb.sheet(s);
    if (!sheet.drawing_rel_target().empty() || sheet.is_opaque_ooxml_sheet()) {
      return true;
    }
    for (const UnknownRelationship& rel : sheet.unknown_relationships()) {
      if (rel.type == io::kRelDrawing) {
        return true;
      }
    }
  }
  return false;
}

// Refuses every object on the sheet that touches the source or destination.
// Objects held only as raw XML (x14 rules, sparklines, slicers, charts,
// drawings, retained worksheet children) refuse on the sheet as a whole.
Expected<void, Error> CheckSheetObjects(const Workbook& wb, std::size_t sheet_index, const Rects& m, Op op) {
  const Sheet& sheet = wb.sheet(sheet_index);
  for (const MergeRange& merge : sheet.merges()) {
    if (Hits(merge, m)) {
      return Touching(op, "a merged range");
    }
  }
  for (const CellComment& note : sheet.comments()) {
    if (Hits(CellRect(note.row, note.col), m)) {
      return Touching(op, "a note");
    }
  }
  for (const ThreadedComment& comment : sheet.threaded_comments()) {
    if (Hits(CellRect(comment.row, comment.col), m)) {
      return Touching(op, "a threaded comment");
    }
  }
  for (const Hyperlink& link : sheet.hyperlinks()) {
    for (const MergeRange& r : {m.src, m.dst}) {
      if (link.intersects(r.first_row, r.first_col, r.last_row, r.last_col)) {
        return Touching(op, "a hyperlink");
      }
    }
  }
  for (const cf::ConditionalFormat& format : sheet.conditional_formats()) {
    for (const cf::CFCellRange& r : format.sqref) {
      if (Hits(MergeRange{r.first.row, r.first.col, r.last.row, r.last.col}, m)) {
        return Touching(op, "a conditional format");
      }
    }
  }
  for (const DataValidation& validation : sheet.validations()) {
    for (const MergeRange& r : validation.ranges) {
      if (Hits(r, m)) {
        return Touching(op, "a data validation");
      }
    }
  }
  // A data table's block and a PivotTable's body extend past their recorded
  // ranges, so either refuses on the sheet as a whole.
  if (!sheet.data_tables().empty()) {
    return Unsupported(op, "cells on a sheet with a data table");
  }
  if (!sheet.pivot_tables().empty()) {
    return Unsupported(op, "cells on a sheet with a PivotTable");
  }
  if (PivotSourceTouches(wb, sheet_index, m)) {
    return Touching(op, "a PivotTable source");
  }
  if (sheet.has_auto_filter()) {
    const AutoFilter* filter = sheet.auto_filter();
    if (filter == nullptr || filter->is_opaque() || Hits(filter->range, m)) {
      return Touching(op, "an AutoFilter");
    }
  }
  for (const TableMetadata& table : wb.tables()) {
    if (table.sheet_index != sheet_index) {
      continue;
    }
    const std::optional<MergeRange> ref = parse_a1_rectangle(table.ref);
    if (!ref.has_value() || Hits(*ref, m)) {
      return Touching(op, "a table");
    }
  }
  for (const MergeRange& r : {m.src, m.dst}) {
    const std::uint32_t rows = r.last_row - r.first_row + 1U;
    const std::uint32_t cols = r.last_col - r.first_col + 1U;
    if (!sheet.committed_spill_anchors_intersecting(r.first_row, r.first_col, rows, cols).empty() ||
        !sheet.blocked_spill_anchors_intersecting(r.first_row, r.first_col, rows, cols).empty()) {
      return Touching(op, "a dynamic-array spill");
    }
  }
  for (const CellAddress at : sheet.all_formula_cells()) {
    const Cell* cell = sheet.cell_at(at.row, at.col);
    if (cell == nullptr || cell->array_block_rows == 0U) {
      continue;
    }
    const MergeRange block{at.row, at.col, at.row + cell->array_block_rows - 1U,
                           at.col + static_cast<std::uint32_t>(cell->array_block_cols) - 1U};
    if (Hits(block, m)) {
      return Touching(op, "an array formula");
    }
  }
  const SheetPrintSettings& print = sheet.print_settings();
  for (const MergeRange& r : {m.src, m.dst}) {
    for (const ManualBreak& page_break : print.manual_row_breaks) {
      if (page_break.id >= r.first_row && page_break.id <= static_cast<std::uint64_t>(r.last_row) + 1U) {
        return Touching(op, "a page break");
      }
    }
    for (const ManualBreak& page_break : print.manual_col_breaks) {
      if (page_break.id >= r.first_col && page_break.id <= static_cast<std::uint64_t>(r.last_col) + 1U) {
        return Touching(op, "a page break");
      }
    }
  }
  for (const SheetSelection& selection : sheet.view().selections) {
    if (SqrefTouches(selection.active_cell, m) || SqrefTouches(selection.sqref, m)) {
      return Touching(op, "the selection");
    }
  }
  if (const char* conflict = AxisConflict(sheet.layout(), m.src); conflict != nullptr) {
    return Unsupported(op, conflict);
  }
  if (const char* conflict = AxisConflict(sheet.layout(), m.dst); conflict != nullptr) {
    return Unsupported(op, conflict);
  }
  // The checks below hold the sheet or the workbook as a whole: the holder is
  // raw XML that is not parsed for its ranges.
  if (!sheet.ext_lst_xml().empty()) {
    return Unsupported(op, "cells on a sheet with extension data");
  }
  if (!sheet.raw_extensions().empty()) {
    return Unsupported(op, "cells on a sheet with retained worksheet elements");
  }
  if (WorkbookHoldsDrawings(wb)) {
    return Unsupported(op, "cells in a workbook with drawings or charts");
  }
  for (std::size_t s = 0; s < wb.sheet_count(); ++s) {
    if (s != sheet_index && !wb.sheet(s).ext_lst_xml().empty()) {
      return Unsupported(op, "cells in a workbook with extension data on another sheet");
    }
  }
  return Expected<void, Error>::Ok();
}

// ---------------------------------------------------------------------------
// Source cells
// ---------------------------------------------------------------------------

// The populated cells of `src` in row-major order.
Expected<std::vector<SourceCell>, Error> ReadSource(const Sheet& sheet, const MergeRange& src, Op op) {
  const std::uint64_t end_col = static_cast<std::uint64_t>(src.last_col) + 1U;
  std::vector<SourceCell> cells;
  for (const auto& entry : sheet.rows()) {
    if (entry.first < src.first_row || entry.first > src.last_row) {
      continue;
    }
    for (const RowCells::Entry stored : entry.second.entries_in(src.first_col, end_col)) {
      const Cell& cell = stored.cell;
      if (!cell.phonetic_runs.empty() || cell.show_phonetic) {
        return Unsupported(op, "cells that carry phonetic text");
      }
      if (cell.array_block_rows != 0U) {
        return Unsupported(op, "array formulas");
      }
      SourceCell out;
      out.row = entry.first;
      out.col = stored.col;
      if (!cell.formula_text.empty()) {
        out.is_formula = true;
        out.text = cell.formula_text;
      } else if (cell.cached_value.is_blank()) {
        continue;
      } else if (cell.cached_value.is_text()) {
        out.is_text = true;
        out.text = std::string(cell.cached_value.as_text());
      } else if (cell.cached_value.is_number() || cell.cached_value.is_boolean() || cell.cached_value.is_error()) {
        out.value = cell.cached_value;
      } else {
        return Unsupported(op, "cells holding anything but a plain value or formula");
      }
      cells.push_back(std::move(out));
    }
  }
  std::sort(cells.begin(), cells.end(),
            [](const SourceCell& a, const SourceCell& b) { return a.row != b.row ? a.row < b.row : a.col < b.col; });
  return cells;
}

// Runs the whole preflight and returns the cells to write, formulas rewritten
// for the destination. Reads only; the workbook is unchanged.
Expected<std::vector<SourceCell>, Error> PlanRangeMove(const Workbook& wb, std::size_t sheet_index,
                                                       const MergeRange& src, std::uint32_t dst_row,
                                                       std::uint32_t dst_col, Op op) {
  RETURN_IF_ERROR(check_sheet_index(sheet_index, wb.sheet_count()));
  if (src.first_row > src.last_row || src.first_col > src.last_col || src.last_row >= Sheet::kMaxRows ||
      src.last_col >= Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kInvalidArgument, "range is outside the grid");
  }
  const std::uint64_t height = static_cast<std::uint64_t>(src.last_row) - src.first_row + 1U;
  const std::uint64_t width = static_cast<std::uint64_t>(src.last_col) - src.first_col + 1U;
  if (dst_row + height > Sheet::kMaxRows || dst_col + width > Sheet::kMaxCols) {
    return make_error(FormulonErrorCode::kInvalidArgument, "row and col place the destination outside the grid");
  }
  const Rects m{src, MergeRange{dst_row, dst_col, dst_row + static_cast<std::uint32_t>(height) - 1U,
                                dst_col + static_cast<std::uint32_t>(width) - 1U}};

  if (op == Op::kCopy && dst_row != src.first_row && dst_col != src.first_col) {
    return Unsupported(op, "with both a row and a column offset");
  }
  if ((src.first_col == 0U && src.last_col == kLastCol) || (src.first_row == 0U && src.last_row == kLastRow)) {
    return Unsupported(op, "whole rows or columns");
  }
  if (intersects(m.src, m.dst)) {
    return Unsupported(op, "cells onto an overlapping destination");
  }
  const Sheet& sheet = wb.sheet(sheet_index);
  if (sheet.is_opaque_ooxml_sheet()) {
    return Unsupported(op, "cells on a sheet that is not a worksheet");
  }
  if (sheet.protection().enabled) {
    return Unsupported(op, "cells on a protected sheet");
  }
  RETURN_IF_ERROR(CheckSheetObjects(wb, sheet_index, m, op));
  if (HasOwnFormat(sheet, m.src)) {
    return Unsupported(op, "cells that carry formats");
  }
  if (HasOwnFormat(sheet, m.dst)) {
    return Unsupported(op, "cells onto formatted cells");
  }
  if (sheet.populated_extent(m.dst.first_row, m.dst.first_col, m.dst.last_row, m.dst.last_col).has_value()) {
    return Unsupported(op, "cells onto a non-empty destination");
  }
  if (op == Op::kMove) {
    RETURN_IF_ERROR(CheckWorkbookReferences(wb, sheet_index, m, op));
  }

  ASSIGN_OR_RETURN(std::vector<SourceCell> cells, ReadSource(sheet, src, op));
  SourceRule rule;
  rule.inside = op == Op::kMove ? &m.src : nullptr;
  rule.row_delta = static_cast<std::int64_t>(dst_row) - static_cast<std::int64_t>(src.first_row);
  rule.col_delta = static_cast<std::int64_t>(dst_col) - static_cast<std::int64_t>(src.first_col);
  for (SourceCell& cell : cells) {
    cell.dest_row = dst_row + (cell.row - src.first_row);
    cell.dest_col = dst_col + (cell.col - src.first_col);
    if (!cell.is_formula) {
      continue;
    }
    Arena arena;
    const AstNode* root = ParseFormula(cell.text, arena, false);
    if (root == nullptr) {
      return Unsupported(op, "formulas that do not parse");
    }
    if (const char* shape = UnsupportedShape(*root, rule); shape != nullptr) {
      return Unsupported(op, shape);
    }
    const parser::FormulaRewriteResult rewritten =
        op == Op::kMove
            ? parser::rewrite_formula(cell.text, parser::MoveTransform(src.first_row, src.first_col, src.last_row,
                                                                       src.last_col, dst_row, dst_col))
            : parser::rewrite_formula(cell.text, parser::CopyTransform(static_cast<std::int32_t>(rule.row_delta),
                                                                       static_cast<std::int32_t>(rule.col_delta)));
    if (rewritten.error) {
      return Unsupported(op, "formulas that cannot be rewritten");
    }
    cell.text = rewritten.text;
  }
  return cells;
}

}  // namespace

Expected<void, Error> Workbook::move_cells(std::size_t sheet_index, const MergeRange& src, std::uint32_t dst_row,
                                           std::uint32_t dst_col) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  ASSIGN_OR_RETURN(const std::vector<SourceCell> cells,
                   PlanRangeMove(*this, sheet_index, src, dst_row, dst_col, Op::kMove));
  // The source cells go first: they are blanked through the value setter, so
  // their formulas leave the dependency graph, then dropped from the store.
  for (const SourceCell& cell : cells) {
    put_cell_value_locked(sheet_index, cell.row, cell.col, Value::blank());
    sheets_[sheet_index].erase_cell(cell.row, cell.col);
  }
  for (const SourceCell& cell : cells) {
    if (cell.is_formula) {
      put_cell_formula_locked(sheet_index, cell.dest_row, cell.dest_col, cell.text, FormulaSource::kRendered, 0U, 0U);
    } else if (cell.is_text) {
      put_cell_text_locked(sheet_index, cell.dest_row, cell.dest_col, cell.text);
    } else {
      put_cell_value_locked(sheet_index, cell.dest_row, cell.dest_col, cell.value);
    }
  }
  return Expected<void, Error>::Ok();
}

Expected<void, Error> Workbook::copy_cells(std::size_t sheet_index, const MergeRange& src, std::uint32_t dst_row,
                                           std::uint32_t dst_col) {
  std::lock_guard<std::mutex> guard(engine_->mutex_for_compound_mutation());
  ASSIGN_OR_RETURN(const std::vector<SourceCell> cells,
                   PlanRangeMove(*this, sheet_index, src, dst_row, dst_col, Op::kCopy));
  for (const SourceCell& cell : cells) {
    if (cell.is_formula) {
      put_cell_formula_locked(sheet_index, cell.dest_row, cell.dest_col, cell.text, FormulaSource::kRendered, 0U, 0U);
    } else if (cell.is_text) {
      put_cell_text_locked(sheet_index, cell.dest_row, cell.dest_col, cell.text);
    } else {
      put_cell_value_locked(sheet_index, cell.dest_row, cell.dest_col, cell.value);
    }
  }
  return Expected<void, Error>::Ok();
}

}  // namespace formulon
