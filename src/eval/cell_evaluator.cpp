//
// Implementation of `evaluate_cell_for_recalc`. See `cell_evaluator.h`
// for the contract and the rationale for centralising this logic.

#include "eval/cell_evaluator.h"

#include <algorithm>
#include <cstddef>
#include <cstdint>
#include <optional>
#include <string_view>
#include <vector>

#include "cell.h"
#include "eval/array_alloc.h"
#include "eval/builtin_names.h"
#include "eval/dep_extractor.h"
#include "eval/dep_graph.h"
#include "eval/dynamic_read_log.h"
#include "eval/error_origin.h"
#include "eval/eval_context.h"
#include "eval/eval_state.h"
#include "eval/function_registry.h"
#include "eval/iterative_solver.h"
#include "eval/recalc_engine.h"
#include "eval/spill_committer.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "parser/parser.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/formula_text.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {

namespace {

// Fills a legacy array formula's `rows` x `cols` block from `result` by
// Excel's array fill rule: a scalar or a one-cell axis repeats across the
// block, and a position the result cannot supply is #N/A.
Value fit_array_block(const Value& result, std::uint32_t rows, std::uint32_t cols, Arena& arena) {
  Value scalar_slot = Value::blank();
  const ArrayView view = as_array_view(result, &scalar_slot);
  Value* buf = nullptr;
  ArrayValue* out = allocate_array_value(rows, cols, arena, buf);
  if (out == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  std::size_t i = 0;
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c, ++i) {
      const Value* cell = broadcast_cell(view, r, c);
      buf[i] = cell != nullptr ? *cell : Value::error(ErrorCode::NA);
    }
  }
  return Value::array(out);
}

// True when `cell` caches a `#VALUE!` that carries a rich detail.
bool holds_rich_value_error(const Workbook& workbook, CellNodeId cell) {
  if (cell.sheet_id >= workbook.sheet_count()) {
    return false;
  }
  const Sheet& sheet = workbook.sheet(cell.sheet_id);
  const std::optional<CellErrorDetail> detail = sheet.error_detail_at(cell.row, cell.col);
  if (!detail.has_value() || detail->code != ErrorCode::Value) {
    return false;
  }
  const Value cached = sheet.resolve_cell_value(cell.row, cell.col);
  return cached.is_error() && cached.as_error() == ErrorCode::Value;
}

// True when the references `node` makes statically reach a cell holding a rich `#VALUE!`.
bool references_rich_value_error(const Workbook& workbook, CellNodeId formula_cell, const parser::AstNode& node) {
  const ExtractedDeps deps = extract_deps(node, formula_cell, workbook);
  for (const CellNodeId cell : deps.cell_deps) {
    if (cell != formula_cell && holds_rich_value_error(workbook, cell)) {
      return true;
    }
  }
  for (const CellRangeDependency& range : deps.range_deps) {
    if (range.sheet_id >= workbook.sheet_count()) {
      continue;
    }
    for (const auto& [address, detail] : workbook.sheet(range.sheet_id).error_details()) {
      const CellNodeId cell{range.sheet_id, address.row, address.col};
      if (detail.code == ErrorCode::Value && cell != formula_cell && range.contains(cell) &&
          holds_rich_value_error(workbook, cell)) {
        return true;
      }
    }
  }
  return false;
}

// True when `node` names one of `params`.
bool mentions_param(const parser::AstNode& node, const std::vector<std::string_view>& params) {
  if (node.kind() == parser::NodeKind::NameRef) {
    for (const std::string_view param : params) {
      if (strings::case_insensitive_eq(node.as_name(), param)) {
        return true;
      }
    }
    return false;
  }
  return parser::any_child_node(node, [&](const parser::AstNode& child) { return mentions_param(child, params); });
}

// A `#VALUE!` that arrived from a precedent holding a rich one, shaped by the
// formula's root: an operator by its Ptg id, a function by its name and the
// first argument that reads such a precedent, anything else as propagated. A
// called LAMBDA is shaped by its body, whose parameters stand for the call's
// arguments.
CellErrorDetail propagated_value_detail(const Workbook& workbook, const FunctionRegistry& registry,
                                        CellNodeId formula_cell, const parser::AstNode& root) {
  CellErrorDetail detail{ErrorCode::Value, CellErrorDetail::Shape::kPropagated, 0U, 0U};
  const parser::AstNode* shaped = &root;
  std::vector<std::string_view> rich_params;
  if (root.kind() == parser::NodeKind::LambdaCall) {
    const parser::AstNode& callee = root.as_lambda_call_callee();
    if (callee.kind() != parser::NodeKind::Lambda) {
      return detail;
    }
    const std::uint32_t bound = std::min(root.as_lambda_call_arity(), callee.as_lambda_param_count());
    for (std::uint32_t i = 0; i < bound; ++i) {
      if (references_rich_value_error(workbook, formula_cell, root.as_lambda_call_arg(i))) {
        rich_params.push_back(callee.as_lambda_param(i));
      }
    }
    shaped = &callee.as_lambda_body();
  }
  if (shaped->kind() == parser::NodeKind::BinaryOp) {
    detail.shape = CellErrorDetail::Shape::kPropagatedPtg;
    detail.root_binop = static_cast<std::uint8_t>(shaped->as_binary_op(shaped->as_binary_arity() - 2U)) + 1U;
    return detail;
  }
  if (shaped->kind() != parser::NodeKind::Call) {
    return detail;
  }
  std::string_view name = find_lazy_canonical_name(shaped->as_call_name());
  if (name.empty()) {
    const FunctionDef* def = registry.lookup(shaped->as_call_name());
    name = def != nullptr ? def->canonical_name : std::string_view();
  }
  if (name.empty()) {
    return detail;
  }
  detail.shape = CellErrorDetail::Shape::kPropagatedIftab;
  detail.root_function = name;
  const std::uint32_t arity = std::min<std::uint32_t>(shaped->as_call_arity(), UINT8_MAX);
  for (std::uint32_t i = 0; i < arity; ++i) {
    const parser::AstNode& arg = shaped->as_call_arg(i);
    if ((!rich_params.empty() && mentions_param(arg, rich_params)) ||
        references_rich_value_error(workbook, formula_cell, arg)) {
      detail.root_argpos = static_cast<std::uint8_t>(i + 1U);
      break;
    }
  }
  return detail;
}

// The detail of a `#VALUE!` with no origin of its own, or nullopt when no
// precedent holds a rich one.
std::optional<CellErrorDetail> received_value_detail(const Workbook& workbook, const FunctionRegistry& registry,
                                                     const Sheet& sheet, std::uint32_t row, std::uint32_t col,
                                                     const parser::AstNode& root) {
  const std::size_t sheet_index = workbook.sheet_index_by_name(sheet.name());
  if (sheet_index >= workbook.sheet_count()) {
    return std::nullopt;
  }
  const CellNodeId cell{static_cast<std::uint16_t>(sheet_index), row, col};
  bool received = false;
  for (const CellNodeId precedent : workbook.recalc_engine().dep_graph().dependencies_of_ref(cell)) {
    if (precedent != cell && holds_rich_value_error(workbook, precedent)) {
      received = true;
      break;
    }
  }
  if (!received && !references_rich_value_error(workbook, cell, root)) {
    return std::nullopt;
  }
  return propagated_value_detail(workbook, registry, cell, root);
}

// The detail `value` is saved with, from the origin slot of its evaluation.
// A `#SPILL!` refused at this anchor carries its footprint instead, and a
// committed spill whose elements carry a measured origin records it for them.
CellResult classify_cell_result(const Workbook& workbook, const FunctionRegistry& registry, const Sheet& sheet,
                                std::uint32_t row, std::uint32_t col, const parser::AstNode* root, Value value) {
  CellResult result;
  result.value = value;
  const ErrorOrigin origin = error_origin();
  if (origin == ErrorOrigin::kMeasured && sheet.spill_region_at_anchor(row, col) != nullptr) {
    CellErrorDetail detail = noted_error_detail();
    detail.elements = true;
    result.detail = detail;
    result.clear = false;
    return result;
  }
  if (!value.is_error()) {
    return result;
  }
  const ErrorCode code = value.as_error();
  const bool measured = origin == ErrorOrigin::kMeasured && noted_error_detail().code == code;
  if (code == ErrorCode::Spill) {
    if (!sheet.blocked_spill_footprint_at(row, col).has_value()) {
      result.detail = CellErrorDetail{ErrorCode::Spill, CellErrorDetail::Shape::kPropagated, 0U, 0U};
    }
  } else if (code == ErrorCode::Calc) {
    if (measured) {
      result.detail = noted_error_detail();
    } else if (origin != ErrorOrigin::kUnknown) {
      result.detail = CellErrorDetail{ErrorCode::Calc, CellErrorDetail::Shape::kPropagated, 0U, 0U};
    }
  } else if (code == ErrorCode::Value) {
    if (measured) {
      result.detail = noted_error_detail();
    } else if (origin == ErrorOrigin::kNone && root != nullptr) {
      result.detail = received_value_detail(workbook, registry, sheet, row, col, *root);
    }
  }
  result.clear = !result.detail.has_value();
  return result;
}

}  // namespace

void commit_cell_result(Sheet& sheet, std::uint32_t row, std::uint32_t col, const CellResult& result) {
  sheet.set_cell_cached_value(row, col, result.value);
  if (result.detail.has_value()) {
    sheet.set_error_detail(row, col, *result.detail);
  } else if (result.clear) {
    sheet.clear_error_detail(row, col);
  }
}

bool stage_formula_cell(const Sheet& sheet, std::uint32_t row, std::uint32_t col, Cell& staged) {
  Sheet::CellRead read;
  sheet.read_formula_cell(row, col, read);
  if (!read.is_formula()) {
    return false;
  }
  staged.formula_text.assign(read.formula_text());
  staged.array_block_rows = read.array_block_rows();
  staged.array_block_cols = static_cast<std::uint16_t>(read.array_block_cols());
  return true;
}

CellResult evaluate_cell_for_recalc(Workbook& workbook, Sheet& sheet, const Cell& cell_data, std::uint32_t row,
                                    std::uint32_t col, const FunctionRegistry& registry, Arena& arena,
                                    const EvaluateCellOptions& opts) {
  const ErrorOriginScope origin_scope;
  // Strip the leading '=' before parsing (the parser expects an expression,
  // not an assignment). `formula_text` is guaranteed non-empty by the
  // caller because we only feed formula cells through here. The shared
  // helper keeps this strip in lockstep across every recalc entry.
  const std::string_view src = strip_formula_prefix(cell_data.formula_text);

  parser::AstNode* root = parse_formula_entry(src, arena);
  if (root == nullptr) {
    // Hard parse failure, or a valid prefix followed by unparseable trailing
    // tokens. Either way we refuse to evaluate a recovered prefix as if it
    // were the whole formula. #NAME? matches the existing recursive resolver
    // behaviour in `EvalContext::resolve_ref`. It is committed like any other
    // result, so a legacy array block keeps its rectangle.
    return classify_cell_result(
        workbook, registry, sheet, row, col, nullptr,
        SpillCommitter(&sheet, row, col, opts.spill_release_callback, opts.spill_release_user_data)
            .commit(Value::error(ErrorCode::Name)));
  }

  // Build an EvalContext that authorises spill writes on `sheet` and
  // anchors the formula at the cell being evaluated. The context is
  // workbook-bound and state-less, so a formula-cell read short-circuits to
  // that cell's cached value instead of recursing into its formula text.
  // Both recalc paths need that: iterative calculation reads each SCC
  // member's last committed value, and the dependency-ordered path has
  // already computed every predecessor, so recursing would re-enter the
  // tree walker once per hop and exhaust its recursion budget on a long
  // chain. Ad-hoc evaluation, which does want recursive resolution, builds
  // its own `EvalContext` rather than coming through this recalc helper.
  EvalContext ctx = EvalContext::workbook_only(workbook, sheet)
                        .with_date1904(workbook.date1904())
                        .with_pinned_now(workbook.pinned_now())
                        .with_mutable_sheet(sheet)
                        .with_spill_release_callback(opts.spill_release_callback, opts.spill_release_user_data)
                        .with_dynamic_read_callback(opts.dynamic_read_callback, opts.dynamic_read_user_data)
                        .with_formula_cell(row, col);
  // The recalc engine owns iterative-calc resolution (SCC detection + the
  // iterative solver). Each per-cell evaluation it issues must be a single
  // pass; suppress the `evaluate()`-level fixed-point driver so the two
  // mechanisms do not double-iterate or fight over divergence accounting.
  ctx = ctx.with_iterative_driver_suppressed();
  if (opts.read_observer != nullptr) {
    ctx = ctx.with_state(*opts.read_observer);
  }

  // A legacy array formula fills its own block at a fixed size, never
  // spilling past it.
  if (cell_data.array_block_rows != 0U) {
    const Value whole = evaluate_array_formula(*root, arena, registry, ctx);
    return classify_cell_result(workbook, registry, sheet, row, col, root,
                                ctx.dispatch_array_result(fit_array_block(whole, cell_data.array_block_rows,
                                                                          cell_data.array_block_cols, arena)));
  }
  Value result = evaluate(*root, arena, registry, ctx);
  // If the top-level evaluator produced an Array (e.g. a SEQUENCE() at the
  // anchor), commit the spill and return the anchor scalar. Mirrors the
  // logic in `EvalContext::resolve_ref` for recursive Array results.
  result = ctx.dispatch_array_result(result);
  return classify_cell_result(workbook, registry, sheet, row, col, root, result);
}

Value CyclicComponentEvaluator::evaluate(CellNodeId cell, EvalState* observer) {
  last_cell_.reset();
  if (cell.sheet_id >= workbook_.sheet_count()) {
    return Value::error(ErrorCode::Ref);
  }
  Sheet& sheet = workbook_.sheet(cell.sheet_id);
  if (!stage_formula_cell(sheet, cell.row, cell.col, staged_)) {
    // No formula text: nothing to evaluate. Blank keeps a value for the
    // solver to compare against.
    return Value::blank();
  }
  // Committed Text scalars survive the reset: `set_cell_cached_value`
  // deep-copies them into the cell, and `commit_spill` does the same for
  // Array results.
  arena_.reset();
  EvaluateCellOptions opts;
  opts.spill_release_callback = release_callback_;
  opts.spill_release_user_data = release_user_data_;
  opts.read_observer = observer;
  if (log_ != nullptr) {
    log_->observe(cell, ordinal_, opts, nullptr);
    log_->note_iterative_member(cell, ordinal_);
  }
  const CellResult result =
      evaluate_cell_for_recalc(workbook_, sheet, staged_, cell.row, cell.col, registry_, arena_, opts);
  if (log_ != nullptr) {
    log_->end();
  }
  last_cell_ = cell;
  last_detail_ = result.detail;
  return result.value;
}

void CyclicComponentEvaluator::commit(CellNodeId cell, Value value) {
  if (cell.sheet_id >= workbook_.sheet_count()) {
    return;
  }
  Sheet& sheet = workbook_.sheet(cell.sheet_id);
  std::unique_lock<std::mutex> guard;
  if (commit_mutex_ != nullptr) {
    guard = std::unique_lock<std::mutex>(*commit_mutex_);
  }
  CellResult result;
  result.value = value;
  if (last_cell_ == cell) {
    result.detail = last_detail_;
  }
  result.clear = !result.detail.has_value();
  commit_cell_result(sheet, cell.row, cell.col, result);
  if (log_ != nullptr) {
    log_->note_commit(cell, ordinal_);
  }
}

bool settle_component_by_reads(const std::vector<CellNodeId>& cells, CyclicComponentEvaluator& evaluator) {
  const Workbook& workbook = evaluator.workbook();
  EvalState observer;
  for (std::size_t i = 0; i < cells.size(); ++i) {
    const CellNodeId c = cells[i];
    if (c.sheet_id < workbook.sheet_count()) {
      observer.observe_reads_of(&workbook.sheet(c.sheet_id), c.row, c.col, i);
    }
  }
  return settle_cycle_by_reads(
      cells.size(),
      [&](std::size_t i, std::vector<std::size_t>* reads) {
        observer.clear_observed_reads();
        const Value v = evaluator.evaluate(cells[i], &observer);
        *reads = observer.observed_reads();
        return v;
      },
      [&](std::size_t i, Value v) { evaluator.commit(cells[i], v); });
}

Value unresolved_cycle_value(const std::vector<CellNodeId>& component) {
  return cell_count_of_component(component) == 1U ? Value::number(0.0) : Value::error(ErrorCode::Ref);
}

void prepare_iterative_component_seeds(Workbook& workbook, const std::vector<CellNodeId>& cells) {
  if (!workbook.recalc_engine().disabled_cycle_refs_pending()) {
    return;
  }
  const std::size_t sheet_count = workbook.sheet_count();
  for (const CellNodeId cell : cells) {
    if (cell.sheet_id >= sheet_count) {
      continue;
    }
    Sheet& sheet = workbook.sheet(cell.sheet_id);
    Sheet::CellRead read;
    sheet.read_formula_cell(cell.row, cell.col, read);
    if (!read.is_formula()) {
      continue;
    }
    const Value cached = read.value();
    if (!cached.is_error() || cached.as_error() != ErrorCode::Ref) {
      continue;
    }
    sheet.set_cell_cached_value(cell.row, cell.col, Value::blank());
  }
}

}  // namespace eval
}  // namespace formulon
