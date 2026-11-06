//
// Recursive node visitor for the tree-walk evaluator. Holds the `eval_node`
// switch (declared with external linkage in `eval/lazy_impls.h` so lazy-impl
// TUs can recurse back into it). The public `evaluate()` overloads live in
// `tree_walker/evaluate_top.cpp`, the bare-range spill positions in
// `tree_walker/spill_position.cpp`.
//
// The function-call dispatch path (`dispatch_call`, range-argument
// expansion) lives in `tree_walker/dispatch.cpp`, the runtime lambda
// invocation (`invoke_lambda`) in `tree_walker/lambda_invoke.cpp`; the
// array-broadcasting helpers (`broadcast_binop`, `broadcast_unary`,
// `apply_binop_per_cell`) live in `tree_walker/broadcast.cpp`. The
// TUs split the original monolithic `tree_walker.cpp` while
// preserving file-local helpers and the existing `formulon::eval`
// namespace shape (anonymous helpers are local to each TU).
//
// See `tree_walker.h` for the public contract.

#include <algorithm>
#include <cmath>
#include <cstddef>
#include <cstdint>
#include <limits>
#include <optional>
#include <string_view>
#include <vector>

#include "eval/array_alloc.h"
#include "eval/coerce.h"
#include "eval/declared_rect.h"
#include "eval/defined_name_resolve.h"
#include "eval/dep_graph.h"
#include "eval/dynamic_array/anchor.h"
#include "eval/eval_clock_scope.h"
#include "eval/eval_context.h"
#include "eval/eval_profile_scope.h"
#include "eval/eval_state.h"
#include "eval/external_ref.h"
#include "eval/function_registry.h"
#include "eval/implicit_intersection.h"
#include "eval/iterative_solver.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/name_env_resolve.h"
#include "eval/range_resolvers.h"
#include "eval/recalc_engine.h"
#include "eval/spill_anchor.h"
#include "eval/structured_ref.h"
#include "eval/structured_ref_project.h"
#include "eval/tree_walker.h"
#include "eval/tree_walker/broadcast.h"
#include "eval/tree_walker/depth_guard.h"
#include "eval/tree_walker/dispatch.h"
#include "eval/tree_walker/walker_internal.h"
#include "eval/tree_walker_lazy_table.h"
#include "parser/ast.h"
#include "sheet.h"
#include "sheet_name.h"
#include "utils/arena.h"
#include "utils/cancellation_snap.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
// Public entry point declared in `eval/tree_walker.h`. Routes through
// the lazy-table seam; the actual array is owned by
// `tree_walker_lazy_table.cpp`.
const char* const* lazy_form_names() {
  return lazy_table_names();
}

namespace {

// Everything `eval_node_shaped` does not carry at declared size: the node
// kinds whose value is never a whole column / row.
Value eval_node_plain(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      return node.as_literal();

    case parser::NodeKind::ErrorLiteral:
      return Value::error(node.as_error_literal());

    case parser::NodeKind::ErrorPlaceholder:
      // Panic-mode skipped this subtree at parse time; we cannot do better
      // than #NAME? since the original tokens are unavailable.
      return Value::error(ErrorCode::Name);

    case parser::NodeKind::SpillRef: {
      // Excel's `=A1#` operator: yields the entire spill region anchored at
      // the referenced cell as a `Value::Array`. Resolution rules:
      //   * Unbound context (no current sheet)             -> #NAME?
      //   * Qualified anchor with no workbook bound        -> #REF!
      //   * Qualified anchor with missing target sheet     -> #REF!
      //   * Anchor row/col >= Sheet::kMax{Rows,Cols}       -> #REF!
      //   * No spill region anchored at this address       -> #REF!
      //   * Computed anchor that names no single cell      -> #REF!
      //
      // The returned ArrayValue header lives in the eval `arena`; its cells
      // buffer holds shallow copies of the SpillRegion's `Value`s. Text
      // payloads point into `SpillRegion::owned_strings`, which lives on
      // Sheet and outlives any single evaluation arena (zero-copy reuse).
      std::string_view anchor_sheet;
      std::uint32_t anchor_row = 0;
      std::uint32_t anchor_col = 0;
      ErrorCode spill_err = ErrorCode::Ref;
      if (!resolve_spill_anchor_node(node, arena, registry, ctx, &anchor_sheet, &anchor_row, &anchor_col, &spill_err)) {
        return Value::error(spill_err);
      }
      ArrayValue* arr = project_spill_at_anchor(anchor_sheet, anchor_row, anchor_col, arena, ctx, &spill_err);
      if (arr == nullptr) {
        return Value::error(spill_err);
      }
      return Value::array(arr);
    }

    case parser::NodeKind::Ref3D:
      // A 3-D reference (`Sheet2:Sheet3!A1`) denotes one cell across a span
      // of sheets — a range shape. Read as a value, Excel answers `#REF!`
      // whether or not the endpoint sheets exist. Range-aware aggregators
      // intercept this node in `dispatch_call` before reaching here, and
      // `IF` / `CHOOSE` turn a selected 3-D arm into `#VALUE!` themselves.
      return Value::error(ErrorCode::Ref);

    case parser::NodeKind::StructuredRef: {
      // Resolve the table reference (`Table[Col]`, `Table[#All]`, ...) to
      // a concrete rectangle and read it through the shared projection.
      bool arena_exhausted = false;
      return project_structured_ref(node.as_structured_ref_table(), node.as_structured_ref_column(), arena, registry,
                                    ctx, &arena_exhausted);
    }

    case parser::NodeKind::Lambda: {
      // Build a runtime closure capturing the current name environment so a
      // body using outer LET-bound names (e.g. `LET(y, 100, LAMBDA(x, x+y))`)
      // sees `y` at call time even though the LET frame has gone out of
      // lexical scope. Param string_views are re-copied into the eval arena
      // even though the parser-arena view they reference also outlives the
      // value: the copy keeps the lifetime story uniform with the rest of
      // the LambdaValue payload.
      auto* lv = arena.create<LambdaValue>();
      if (lv == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      const std::uint32_t n = node.as_lambda_param_count();
      std::string_view* params = nullptr;
      if (n > 0) {
        params = arena.create_array<std::string_view>(n);
        if (params == nullptr) {
          return Value::error(ErrorCode::Num);
        }
        for (std::uint32_t i = 0; i < n; ++i) {
          params[i] = node.as_lambda_param(i);
        }
      }
      lv->params = params;
      lv->param_count = n;
      lv->optional_count = node.as_lambda_optional_count();
      lv->body = &node.as_lambda_body();
      lv->name_scope_sheet = ctx.name_scope_sheet();
      // Copy the caller's NameEnv into the arena: the live `NameEnv` value at
      // `ctx.name_env()` typically lives on a parent eval_node frame that
      // disappears once that frame returns, but every `Binding*` it points
      // at is arena-allocated and survives. Cloning the small wrapper struct
      // into the arena keeps the closure's reach into the binding chain
      // valid for the lifetime of the LambdaValue.
      if (const NameEnv* parent = ctx.name_env(); parent != nullptr) {
        auto* env_copy = arena.create<NameEnv>();
        if (env_copy == nullptr) {
          return Value::error(ErrorCode::Num);
        }
        *env_copy = *parent;
        lv->captured_env = env_copy;
      } else {
        lv->captured_env = nullptr;
      }
      return Value::lambda(lv);
    }

    case parser::NodeKind::LambdaCall: {
      // Calling a reference (`A1(1)`, `Sheet1!LOG10(100)`, `(A1:A2)(1)`) is
      // #REF! in Excel, whatever the cells hold.
      if (is_reference_shape(node.as_lambda_call_callee())) {
        return Value::error(ErrorCode::Ref);
      }
      // Evaluate the callee expression. Excel rejects calling a non-lambda
      // with #VALUE! (e.g. `(1+2)(3)` — when the parser admits the form).
      const Value callee = eval_node(node.as_lambda_call_callee(), arena, registry, ctx);
      if (callee.is_error()) {
        return callee;
      }
      if (!callee.is_lambda()) {
        return Value::error(ErrorCode::Value);
      }
      // invoke_lambda takes the arena slots directly and binds omitted optionals.
      return invoke_lambda(callee.as_lambda(), node.as_lambda_call_arity(), node.as_lambda_call_args(), arena, registry,
                           ctx);
    }

    case parser::NodeKind::IntersectOp:
      // A nested position, so the rectangle is materialised rather than
      // weighed against a spill footprint — the same split the `RangeOp`
      // case above draws against `evaluate()`.
      return eval_intersect_op(node, arena, registry, ctx, /*spill_position=*/false);

    case parser::NodeKind::ArrayLiteral: {
      // A brace literal is a first-class dynamic array in modern Excel, so
      // the tree walker can spill it, broadcast it through an operator, and
      // let `@` reduce it through the common implicit-intersection path.
      const std::uint32_t rows = node.as_array_rows();
      const std::uint32_t cols = node.as_array_cols();
      Value* cells = nullptr;
      ArrayValue* array = allocate_derived_array_value(rows, cols, arena, cells);
      if (array == nullptr) {
        return Value::error(ErrorCode::Num);
      }
      for (std::uint32_t row = 0; row < rows; ++row) {
        for (std::uint32_t col = 0; col < cols; ++col) {
          cells[static_cast<std::size_t>(row) * cols + col] =
              eval_node(node.as_array_element(row, col), arena, registry, ctx);
        }
      }
      return Value::array(array);
    }

    // -- Unsupported range-producing operator ------------------------------
    case parser::NodeKind::UnionOp:
      return Value::error(ErrorCode::Value);

    default:
      break;
  }
  return Value::error(ErrorCode::Value);
}

// A `:` whose endpoint is a reference-returning call (`A1:INDEX(...)`,
// `A1:OFFSET(...)`) names its rectangle only once resolved; the union then
// spills like any bounded range.
Value eval_range_of_calls(const parser::AstNode& range, Arena& arena, const FunctionRegistry& registry,
                          const EvalContext& ctx) {
  ErrorCode err = ErrorCode::Value;
  parser::Reference top_left{};
  parser::Reference bottom_right{};
  if (!resolve_range_endpoints(range, arena, registry, ctx, &top_left.sheet, &top_left.row, &top_left.col,
                               &bottom_right.row, &bottom_right.col, &err)) {
    return Value::error(err);
  }
  bottom_right.sheet = top_left.sheet;
  if (top_left.row == bottom_right.row && top_left.col == bottom_right.col) {
    return ctx.resolve_ref(top_left, arena, registry);
  }
  return materialize_rectangle(top_left, bottom_right, arena, registry, ctx);
}

}  // namespace

// First element of a shaped result, read without expanding it.
Value first_element(const Shaped& s) {
  if (s.tail_array != nullptr) {
    return tail_array_at(*s.tail_array, 0U, 0U);
  }
  return implicit_intersect_value(s.value);
}

namespace {

// Operand of a broadcast binary op at output cell (r, c), coerced to a
// number; empty when the operand cannot supply that cell or is not numeric.
std::optional<double> broadcast_operand_number(const Value& v, std::uint32_t r, std::uint32_t c) {
  Value slot = Value::blank();
  const Value* cell = broadcast_cell(as_array_view(v, &slot), r, c);
  if (cell == nullptr) {
    return std::nullopt;
  }
  const auto n = coerce_to_number(*cell);
  return n ? std::optional<double>(n.value()) : std::nullopt;
}

// `out` with the cancellation snap applied to each numeric cell of an
// array-valued `+` / `-` result.
Value snap_array_result(const Value& out, const Value& lhs, const Value& rhs, Arena& arena) {
  const ArrayValue* src = out.as_array();
  Value* cells = nullptr;
  ArrayValue* snapped = allocate_derived_array_value(src->rows, src->cols, arena, cells);
  if (snapped == nullptr) {
    return out;
  }
  for (std::uint32_t r = 0; r < src->rows; ++r) {
    for (std::uint32_t c = 0; c < src->cols; ++c) {
      const std::size_t i = static_cast<std::size_t>(r) * src->cols + c;
      cells[i] = src->cells[i];
      if (!cells[i].is_number()) {
        continue;
      }
      const std::optional<double> a = broadcast_operand_number(lhs, r, c);
      const std::optional<double> b = broadcast_operand_number(rhs, r, c);
      if (a && b) {
        cells[i] = Value::number(snap_cancellation(*a, cells[i].as_number()));
      }
    }
  }
  return Value::array(snapped);
}

// A binary operator chain, folded left to right. `snap_cancellation_result`
// applies Excel's near-zero cancellation snap to the last `+` / `-` result,
// cell by cell when it is an array.
Shaped eval_binary_op(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                      const EvalContext& ctx, bool snap_cancellation_result) {
  // Left-most wins between scalars; a scalar error against an array operand
  // fills every position of the broadcast instead (`ROWS(1/0+{1;2;3})` is 3).
  Shaped lhs = eval_node_shaped(node.as_binary_operand(0), arena, registry, ctx);
  const std::uint32_t arity = node.as_binary_arity();
  for (std::uint32_t i = 1; i < arity; ++i) {
    const Shaped rhs = eval_node_shaped(node.as_binary_operand(i), arena, registry, ctx);
    const bool lhs_array = lhs.tail_array != nullptr || lhs.value.is_array();
    const bool rhs_array = rhs.tail_array != nullptr || rhs.value.is_array();
    if (!rhs_array && lhs.tail_array == nullptr && lhs.value.is_error()) {
      return lhs;
    }
    if (!lhs_array && rhs.tail_array == nullptr && rhs.value.is_error()) {
      return rhs;
    }
    // Cellwise broadcast when either operand is an array (SpillRef #,
    // TRANSPOSE, SEQUENCE, a whole column / row, ...); two scalars take
    // the per-cell fast path.
    Shaped out = broadcast_binop(node.as_binary_op(i - 1U), lhs, rhs, arena);
    if (snap_cancellation_result && i + 1U == arity && out.tail_array == nullptr && lhs.tail_array == nullptr &&
        rhs.tail_array == nullptr) {
      if (out.value.is_array()) {
        out.value = snap_array_result(out.value, lhs.value, rhs.value, arena);
      } else if (out.value.is_number()) {
        const std::optional<double> a = broadcast_operand_number(lhs.value, 0U, 0U);
        const std::optional<double> b = broadcast_operand_number(rhs.value, 0U, 0U);
        if (a && b) {
          out.value = Value::number(snap_cancellation(*a, out.value.as_number()));
        }
      }
    }
    lhs = out;
  }
  return lhs;
}

}  // namespace

// The value of a formula's root node. An unparenthesised root `+` / `-` is
// the one operator position where Excel snaps a near-zero cancellation to 0.
Shaped eval_root_shaped(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  if (node.kind() == parser::NodeKind::BinaryOp && node.paren_depth() == 0U &&
      (node.as_binary_op(0) == parser::BinOp::Add || node.as_binary_op(0) == parser::BinOp::Sub)) {
    EvalDepthGuard depth_guard(ctx.eval_depth_counter(), kMaxEvalDepth);
    if (depth_guard.exceeded()) {
      return Shaped{Value::error(ErrorCode::Calc), nullptr};
    }
    return eval_binary_op(node, arena, registry, ctx, /*snap_cancellation_result=*/true);
  }
  return eval_node_shaped(node, arena, registry, ctx);
}

// Defined with external linkage (declared in `eval/lazy_impls.h`) so the
// per-family lazy-impl TUs can recurse into the evaluator. The scalar
// operator helpers it calls below -- `apply_unary`, `apply_arithmetic`,
// `apply_concat`, `apply_comparison` -- live in `eval/scalar_ops.h` and are
// reachable via ordinary unqualified lookup. `dispatch_call` lives in
// `tree_walker/dispatch.cpp` and is declared in
// `eval/tree_walker/dispatch.h`; the broadcast helpers live in
// `tree_walker/broadcast.cpp` and are declared in
// `eval/tree_walker/broadcast.h`.
Value eval_node(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry, const EvalContext& ctx) {
  return densify(eval_node_shaped(node, arena, registry, ctx), arena);
}

Shaped eval_node_shaped(const parser::AstNode& node, Arena& arena, const FunctionRegistry& registry,
                        const EvalContext& ctx) {
  // Bounds linear cell-chain recursion through `EvalContext::resolve_ref`
  // and any other re-entrant evaluator path. See `kMaxEvalDepth`.
  EvalDepthGuard depth_guard(ctx.eval_depth_counter(), kMaxEvalDepth);
  if (depth_guard.exceeded()) {
    return Shaped{Value::error(ErrorCode::Calc), nullptr};
  }
  switch (node.kind()) {
    case parser::NodeKind::RangeOp: {
      // Excel 365 dynamic-array semantics: a bare bounded range used in a
      // value context spills. It evaluates to the whole rectangle as a
      // Value::Array, which bubbles up to the cell entry point (committing a
      // spill) or to an enclosing operator's cellwise broadcast, rather than
      // collapsing to a single implicit-intersection cell. Legacy implicit
      // intersection is reached only through the explicit `@` / SINGLE
      // wrapper (the ImplicitIntersection case below), never here.
      //
      // A rectangle spanning a whole grid axis (`A:C`, `1:3`, `A1:A1048576`)
      // is read at its declared size as a `TailArray`; a single-cell
      // (`A1:A1`) range degrades to the scalar.
      //
      // Verified Mac semantics: tests/oracle/cases/implicit_intersection.yaml.
      const auto& lhs = node.as_range_operand(0);
      const auto& rhs = node.as_range_operand(1);
      if (node.as_range_arity() != 2U || lhs.kind() != parser::NodeKind::Ref || rhs.kind() != parser::NodeKind::Ref) {
        return Shaped{eval_range_of_calls(node, arena, registry, ctx), nullptr};
      }
      const auto& lhs_ref = lhs.as_ref();
      const auto& rhs_ref = rhs.as_ref();
      if (has_whole_axis_endpoint(lhs_ref, rhs_ref) || is_full_axis_range(lhs_ref, rhs_ref)) {
        return read_whole_axis(lhs_ref, rhs_ref, arena, registry, ctx);
      }
      const std::uint32_t r1 = std::min(lhs_ref.row, rhs_ref.row);
      const std::uint32_t r2 = std::max(lhs_ref.row, rhs_ref.row);
      const std::uint32_t c1 = std::min(lhs_ref.col, rhs_ref.col);
      const std::uint32_t c2 = std::max(lhs_ref.col, rhs_ref.col);

      parser::Reference top_left{};
      top_left.sheet = lhs_ref.sheet;
      top_left.row = r1;
      top_left.col = c1;
      if (r1 == r2 && c1 == c2) {
        return Shaped{ctx.resolve_ref(top_left, arena, registry), nullptr};
      }

      parser::Reference bottom_right{};
      bottom_right.sheet = lhs_ref.sheet;
      bottom_right.row = r2;
      bottom_right.col = c2;
      return Shaped{materialize_rectangle(top_left, bottom_right, arena, registry, ctx), nullptr};
    }

    case parser::NodeKind::ImplicitIntersection: {
      const auto& operand = node.as_implicit_intersection_operand();
      // Implicit intersection on a reference: project the formula cell onto
      // the declared rectangle. All three spellings that declare one --
      // bounded `Ref:Ref`, full-axis `Ref:Ref`, and the single `Ref` the
      // parser folds `A:A` / `1:1` into -- go through the projection shared
      // with `_xlfn.SINGLE`, so `=@A:B` and `=@A1:B3` agree wherever they
      // denote the same rectangle.
      if (ctx.has_formula_cell()) {
        parser::Reference target{};
        switch (project_implicit_intersection(operand, ctx.formula_row(), ctx.formula_col(), &target)) {
          case IntersectionProjection::kCell:
            return Shaped{ctx.resolve_ref(target, arena, registry), nullptr};
          case IntersectionProjection::kNoCell:
            return Shaped{Value::error(ErrorCode::Value), nullptr};
          case IntersectionProjection::kNotStaticReference:
            break;
        }
        if (Value projected = Value::blank(); project_reference_result(operand, arena, registry, ctx, &projected)) {
          return Shaped{projected, nullptr};
        }
      } else if (operand.kind() == parser::NodeKind::RangeOp) {
        // No formula-cell context (top-level evaluator entry) -> degrade to
        // top-left, matching the bare-range fallback. Production calls
        // through Workbook always supply a formula cell, so this branch
        // only fires for parser-driven smoke tests.
        const Shaped operand_value = eval_node_shaped(operand, arena, registry, ctx);
        if (operand_value.tail_array != nullptr) {
          return Shaped{first_element(operand_value), nullptr};
        }
        return operand_value;
      }
      // Dynamic arrays produced by a call, spill reference, or expression no
      // longer retain static range coordinates. Excel's `@` takes their
      // top-left element instead of allowing the value to spill.
      return Shaped{first_element(eval_node_shaped(operand, arena, registry, ctx)), nullptr};
    }

    case parser::NodeKind::UnaryOp: {
      // Eager scalar unary; broadcast cellwise when the operand evaluates to
      // an array (e.g. `=-A1#`). The array result then bubbles up to the cell
      // entry point where dispatch_array_result decides whether to commit a
      // spill.
      const Shaped operand = eval_node_shaped(node.as_unary_operand(), arena, registry, ctx);
      return broadcast_unary(node.as_unary_op(), operand, arena);
    }

    case parser::NodeKind::BinaryOp:
      return eval_binary_op(node, arena, registry, ctx, /*snap_cancellation_result=*/false);

    case parser::NodeKind::Call:
      return dispatch_call(node, arena, registry, ctx);

    case parser::NodeKind::Ref: {
      const parser::Reference& ref = node.as_ref();
      if (ref.is_full_col || ref.is_full_row) {
        return read_whole_axis(ref, ref, arena, registry, ctx);
      }
      return Shaped{ctx.resolve_ref(ref, arena, registry), nullptr};
    }

    case parser::NodeKind::NameRef: {
      // Resolution order: lexical scope (LET / LAMBDA bindings) wins over a
      // workbook / sheet-scoped defined name, so `=LET(Rate, 2, Rate)` reads
      // the binding, not a `Rate` defined name. When no binding matches, fall
      // through to defined-name resolution, which returns `#NAME?` itself when
      // the name is undefined in scope. `Sheet1!Name` is never a lexical
      // binding and is looked up in Sheet1's scope.
      if (const std::string_view sheet = node.as_name_sheet(); !sheet.empty()) {
        return resolve_sheet_defined_name_shaped(sheet, node.as_name(), arena, registry, ctx);
      }
      const NameEnv* env = ctx.name_env();
      if (env != nullptr) {
        const auto read = [&](const parser::AstNode& ref) { return eval_node_shaped(ref, arena, registry, ctx); };
        if (const auto* binding = env->lookup(node.as_name()); binding != nullptr) {
          return NameEnv::read_shaped(*binding, read);
        }
      }
      // An undefined built-in name is that function as a value (`=TYPE(SUM)`).
      if (find_defined_name(ctx, node.as_name()) == nullptr &&
          (lookup_on_current_host(registry, node.as_name()) != nullptr ||
           find_lazy_impl(strip_future_prefix(node.as_name())) != nullptr)) {
        auto* lv = arena.create<LambdaValue>();
        if (lv == nullptr) {
          return Shaped{Value::error(ErrorCode::Num), nullptr};
        }
        lv->builtin = node.as_name();
        return Shaped{Value::lambda(lv), nullptr};
      }
      return resolve_defined_name_shaped(node.as_name(), arena, registry, ctx);
    }

    case parser::NodeKind::LetBinding: {
      // Sequential (left-to-right) bind-then-body. Excel semantics:
      //   * Each binding initialiser evaluates in the scope of previously
      //     bound names, so `LET(x, 1, y, x+2, y)` returns 3.
      //   * Error values DO flow into the environment -- downstream
      //     expressions (including `IFERROR` inside the body) may catch
      //     them: `LET(x, 1/0, IFERROR(x, 99))` returns 99.
      //   * Names are ASCII-case-insensitive and a later binding with the
      //     same name shadows earlier ones in subsequent expressions.
      //   * A reference initialiser binds as a reference, by the rule every
      //     LAMBDA argument follows too (see `eval_binding_source`).
      NameEnv env;
      const NameEnv* parent = ctx.name_env();
      // Start from whatever the caller supplied; extending `NameEnv` makes
      // `env` point at a new head frame while preserving the parent chain.
      if (parent != nullptr) {
        env = *parent;
      }
      const std::uint32_t count = node.as_let_binding_count();
      for (std::uint32_t i = 0; i < count; ++i) {
        const parser::AstNode* bound_ast = nullptr;
        const Value v =
            eval_binding_source(node.as_let_binding_expr(i), arena, registry, ctx.with_name_env(&env), &bound_ast);
        env = env.extend(node.as_let_binding_name(i), v, bound_ast, arena);
      }
      const EvalContext body_ctx = ctx.with_name_env(&env);
      return eval_node_shaped(node.as_let_body(), arena, registry, body_ctx);
    }

    case parser::NodeKind::ExternalRef:
      // `[0]!Name` names a defined name of this workbook itself.
      if (parser::is_self_book_name_ref(node)) {
        return resolve_self_book_defined_name_shaped(node.as_external_ref_name(), arena, registry, ctx);
      }
      // Read straight out of the external-link cache. Unlike `Ref3D` this
      // needs no sheet resolution against the workbook: the target lives
      // in another file whose grid the cache already holds, so a
      // rectangle materialises here rather than being routed through
      // `expand_range`.
      return resolve_external_ref_shaped(node, arena, ctx);

    default:
      return Shaped{eval_node_plain(node, arena, registry, ctx), nullptr};
  }
}

}  // namespace eval
}  // namespace formulon
