// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Stack-machine VM implementation. See `vm.h` for the public contract and
// `bytecode.h` for the IR shape.
//
// Design:
//   * Single switch over `OpCode`. Plain dispatch: Emscripten does not
//     synthesise a threaded interpreter from computed-goto / tail-call
//     idioms, and the bundled WASM size budget rewards clarity over a
//     hand-tuned dispatch.
//   * Operand stack lives in a `std::vector<Value>` reserved up front. We
//     do not lift it into the arena because the WASM toolchain emits
//     smaller code for the standard vector path.
//   * LET slots live in a single `std::vector<Value>` keyed by the
//     compiler-assigned 24-bit slot index. The compiler never re-uses
//     slots, so the vector grows monotonically and the VM never has to
//     reset / scope it across nested LET / LAMBDA bodies.
//   * Lambda invocation drives a sub-loop on the same `code` stream. The
//     compiler splices lambda bodies inline (after a Jump that the parent
//     stream takes); a `CallLambda` jumps into the body, the body's
//     `Return` pops the frame, and execution continues at the parent's
//     return pc. Lambda values are represented as ordinary
//     `eval::LambdaValue`s whose `body` slot is repurposed to point at a
//     VM-internal closure record (see `VmClosure` below). The VM never
//     dereferences `body` as an AST node, and the tree-walker never sees
//     a VM-emitted lambda, so the cast-pun is contained to this TU.
//
// IFERROR / IFNA short-circuit: the bytecode IR pre-evaluates both args
// (matching the documented `Call` shape in `compiler.cpp`); the VM
// inspects the primary post-hoc and selects either the primary or the
// fallback. This drifts from the tree-walker only when the *fallback*
// itself raises a different error and the primary is fine — see the
// parity-test corpus for the practical avoidance pattern.

#include "eval/vm.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/bytecode.h"
#include "eval/coerce.h"
#include "eval/compiler_emit.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/scalar_ops.h"
#include "eval/structured_ref.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "sheet.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "utils/status_macros.h"
#include "utils/strings.h"
#include "value.h"
#include "workbook.h"

namespace formulon {
namespace eval {
namespace {

// Hard cap on operand stack depth. Defends against a corrupted bytecode
// stream that emits unbounded `Union` / array-literal pushes. The limit is
// generous relative to anything well-formed compile() emits (deeply nested
// formulas peak in the low hundreds).
constexpr std::size_t kMaxStackDepth = 65536;

// VM-internal closure record. Lives in the eval arena. Pointed to by a
// `LambdaValue` whose `body` slot is repurposed (see file header). The VM
// never crosses with the tree-walker's AST-bound lambdas.
struct VmClosure {
  std::uint32_t body_pc = 0;  // first instruction of the body
};

Error make_vm_error(FormulonErrorCode code, const char* msg) {
  return make_error(code, std::string(msg));
}

// Lazy IFERROR / IFNA classification. The bytecode emits a `Call` with the
// canonical name; we route those two through a post-hoc selector rather
// than the function registry (the registry has no entry for them).
enum class LazyCallKind : std::uint8_t { None, IfError, IfNa };

LazyCallKind classify_lazy_call(std::string_view name) noexcept {
  if (strings::case_insensitive_eq(name, "IFERROR")) {
    return LazyCallKind::IfError;
  }
  if (strings::case_insensitive_eq(name, "IFNA")) {
    return LazyCallKind::IfNa;
  }
  return LazyCallKind::None;
}

// Builds a string_view-backed copy of `s` that lives in `arena`. Used when
// we must intern a temporary string (e.g. concat result) so the `Value::Text`
// payload remains valid after the underlying buffer goes out of scope.
std::string_view intern_arena_string(Arena& arena, std::string_view s) {
  if (s.empty()) {
    return std::string_view{};
  }
  char* buf = arena.create_array<char>(s.size());
  if (buf == nullptr) {
    return std::string_view{};
  }
  std::memcpy(buf, s.data(), s.size());
  return std::string_view(buf, s.size());
}

// Coerces a Value to Excel-truthy boolean for `JumpIfFalse`. Errors fall
// through as TRUE (i.e. the jump is NOT taken) so the error continues to
// flow through whatever the taken branch returns. Mirrors the tree-walker's
// IF-on-error contract: `=IF(1/0, "a", "b")` returns `#DIV/0!`, not "b".
//
// Returns nullopt-equivalent via the `out_err` channel when coercion itself
// fails (e.g. `Text("hello")` to bool); the VM then treats that as TRUE so
// the surrounding error flow stays unchanged.
bool coerce_to_truthy_for_jump(const Value& v, bool* out_is_error) {
  *out_is_error = false;
  if (v.is_error()) {
    *out_is_error = true;
    return true;  // arbitrary; caller short-circuits on error before consulting
  }
  auto b = coerce_to_bool(v);
  if (!b) {
    return true;  // unparseable text -> propagate via JumpIfFalse-not-taken
  }
  return b.value();
}

// Resolves a structured-ref bracket payload through the workbook. Mirrors
// tree_walker.cpp's `NodeKind::StructuredRef` branch and funnels the
// resolved rectangle through `EvalContext::expand_range` so cross-sheet
// resolution + cycle detection stays in one place.
Value resolve_struct_ref_op(std::string_view table_name, std::string_view column_payload, Arena& arena,
                            const FunctionRegistry& registry, const EvalContext& ctx) {
  const Workbook* wb = ctx.workbook();
  const Sheet* current = ctx.current_sheet();
  if (wb == nullptr || current == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  auto sel_or = parse_structured_ref_payload(column_payload);
  if (!sel_or) {
    return Value::error(sel_or.error());
  }
  StructuredRefSelector sel = std::move(sel_or).value();
  sel.table_name = table_name;
  std::uint32_t current_sheet_index = 0;
  for (std::size_t i = 0; i < wb->sheet_count(); ++i) {
    if (&wb->sheet(i) == current) {
      current_sheet_index = static_cast<std::uint32_t>(i);
      break;
    }
  }
  const std::uint32_t current_row = ctx.has_formula_cell() ? ctx.formula_row() : EvalContext::kNoFormulaCell;
  auto rect_or = resolve_structured_ref(sel, *wb, current_sheet_index, current_row);
  if (!rect_or) {
    return Value::error(rect_or.error());
  }
  const StructuredRefRange rect = std::move(rect_or).value();
  parser::Reference lhs{};
  lhs.sheet = rect.sheet_name;
  lhs.row = rect.row_first;
  lhs.col = rect.col_first;
  parser::Reference rhs{};
  rhs.sheet = rect.sheet_name;
  rhs.row = rect.row_last;
  rhs.col = rect.col_last;
  if (rect.row_first == rect.row_last && rect.col_first == rect.col_last) {
    return ctx.resolve_ref(lhs, arena, registry);
  }
  auto cells = ctx.expand_range(lhs, rhs, arena, registry);
  if (!cells) {
    return Value::error(cells.error());
  }
  const std::uint32_t rows = rect.row_last - rect.row_first + 1u;
  const std::uint32_t cols = rect.col_last - rect.col_first + 1u;
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  Value* buffer = arena.create_array<Value>(total);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < total && i < cells.value().size(); ++i) {
    buffer[i] = cells.value()[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = rows;
  arr->cols = cols;
  arr->cells = buffer;
  return Value::array(arr);
}

// Mirror of tree_walker.cpp's SpillRef branch: resolve the spill region at
// the anchor and project it as a `Value::Array`.
Value resolve_spill_ref_op(const parser::Reference& r, Arena& arena, const EvalContext& ctx) {
  const Sheet* current = ctx.current_sheet();
  if (current == nullptr) {
    return Value::error(ErrorCode::Name);
  }
  const Sheet* target = current;
  if (!r.sheet.empty()) {
    const Workbook* wb = ctx.workbook();
    if (wb == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
    target = wb->sheet_by_name(r.sheet);
    if (target == nullptr) {
      return Value::error(ErrorCode::Ref);
    }
  }
  if (r.row >= Sheet::kMaxRows || r.col >= Sheet::kMaxCols) {
    return Value::error(ErrorCode::Ref);
  }
  const SpillRegion* region = target->spill_region_at_anchor(r.row, r.col);
  if (region == nullptr) {
    return Value::error(ErrorCode::Ref);
  }
  const std::size_t n = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
  Value* buffer = arena.create_array<Value>(n);
  if (buffer == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = region->cells[i];
  }
  ArrayValue* arr = arena.create<ArrayValue>();
  if (arr == nullptr) {
    return Value::error(ErrorCode::Num);
  }
  arr->rows = region->rows;
  arr->cols = region->cols;
  arr->cells = buffer;
  return Value::array(arr);
}

// Per-execute() VM state. Lives on the call stack of `execute()`; the
// dispatch loop receives a pointer to it.
struct VmState {
  std::vector<Value> stack;
  std::vector<Value> let_slots;
  // Lambda-call frame: snapshot of the call's argument slots. Pushed by
  // CallLambda, popped by Return when the frame is the active body.
  struct Frame {
    std::uint32_t return_pc = 0;
    std::vector<Value> args;
  };
  std::vector<Frame> frames;
};

// Pop helper: returns kVmStackUnderflow when the stack is shorter than `n`.
Expected<void, Error> require_stack_depth(const VmState& s, std::size_t n) {
  if (s.stack.size() < n) {
    return make_vm_error(FormulonErrorCode::kVmStackUnderflow, "operand stack underflow");
  }
  return {};
}

// Push helper: returns kVmStackOverflow when the stack would exceed the
// hard cap. Defends against malformed bytecode emitting unbounded pushes.
Expected<void, Error> push_value(VmState& s, Value v) {
  if (s.stack.size() >= kMaxStackDepth) {
    return make_vm_error(FormulonErrorCode::kVmStackOverflow, "operand stack overflow");
  }
  s.stack.push_back(v);
  return {};
}

void pop_values(VmState& s, std::size_t n) noexcept {
  for (std::size_t i = 0; i < n; ++i) {
    s.stack.pop_back();
  }
}

// Pool accessors (`const_at` / `name_at` / `ref_at`) live in
// `compiler_emit.h` so VM and tooling share a single bounds-check
// implementation.

// Drives the dispatch loop starting at `start_pc`. Returns when a
// `Return` / `Halt` is hit at depth 0 (i.e. with no active lambda frame).
// Lambda invocation re-enters this function for the body via a recursive
// call, with the body's pc as the start.
Expected<Value, Error> dispatch(const ByteCode& bc, Arena& arena, const FunctionRegistry& registry,
                                const EvalContext& ctx, VmState& s, std::uint32_t start_pc) {
  if (bc.code.empty()) {
    return make_vm_error(FormulonErrorCode::kVmEmptyBytecode, "empty bytecode body");
  }
  std::uint32_t pc = start_pc;
  const std::uint32_t code_size = static_cast<std::uint32_t>(bc.code.size());

  while (pc < code_size) {
    const Instruction& ins = bc.code[pc];
    switch (ins.op) {
      case OpCode::LoadConst: {
        auto v = const_at(bc, ins.a);
        if (!v) {
          return v.error();
        }
        // Text constants borrow from `bc.string_storage`. The bytecode body
        // typically lives only for the duration of `execute()`, so a Text
        // payload escaping the call (e.g. into a parity-test result) would
        // dangle. Re-intern into `arena`, which the caller guarantees to
        // outlive every returned Value (matches the tree-walker's text-
        // literal handling, which also borrows from the parser arena).
        Value pushed = *v.value();
        if (pushed.kind() == ValueKind::Text) {
          pushed = Value::text(intern_arena_string(arena, pushed.as_text()));
        }
        RETURN_IF_ERROR(push_value(s, pushed));
        ++pc;
        break;
      }

      case OpCode::LoadRef: {
        auto r = ref_at(bc, ins.a);
        if (!r) {
          return r.error();
        }
        RETURN_IF_ERROR(push_value(s, ctx.resolve_ref(*r.value(), arena, registry)));
        ++pc;
        break;
      }

      case OpCode::LoadRange: {
        // The compiler emits the two endpoint subexpressions before this
        // opcode, so the stack already carries the endpoint values. The IR
        // does not preserve the Ref payloads at this point, so the VM
        // collapses to the lhs (top-left analog) — matching the tree-walker
        // RangeOp scalar fallback when no formula cell is bound. Range-aware
        // aggregator parity is a documented limitation tracked for Bundle
        // 5.3+ optimisation.
        RETURN_IF_ERROR(require_stack_depth(s, 2));
        const Value rhs = s.stack.back();
        s.stack.pop_back();
        (void)rhs;
        // Top of stack is now the lhs — we keep it as-is.
        ++pc;
        break;
      }

      case OpCode::LoadName: {
        auto n = name_at(bc, ins.a);
        if (!n) {
          return n.error();
        }
        // Tree-walker mirror: defined-name lookup is not yet wired, so an
        // unbound identifier surfaces #NAME?. A runtime NameEnv (from a
        // captured lambda environment) takes precedence.
        Value v = Value::error(ErrorCode::Name);
        if (const NameEnv* env = ctx.name_env(); env != nullptr) {
          if (const Value* bound = env->lookup(*n.value()); bound != nullptr) {
            v = *bound;
          }
        }
        RETURN_IF_ERROR(push_value(s, v));
        ++pc;
        break;
      }

      case OpCode::LoadStructRef: {
        auto table_n = name_at(bc, ins.a);
        if (!table_n) {
          return table_n.error();
        }
        const std::uint32_t col_idx = ins.b & 0xFFFFu;
        auto col_n = name_at(bc, col_idx);
        if (!col_n) {
          return col_n.error();
        }
        const Value v = resolve_struct_ref_op(*table_n.value(), *col_n.value(), arena, registry, ctx);
        RETURN_IF_ERROR(push_value(s, v));
        ++pc;
        break;
      }

      case OpCode::LoadSpillRef: {
        auto r = ref_at(bc, ins.a);
        if (!r) {
          return r.error();
        }
        RETURN_IF_ERROR(push_value(s, resolve_spill_ref_op(*r.value(), arena, ctx)));
        ++pc;
        break;
      }

      case OpCode::LoadExternalRef: {
        // External workbook refs are not yet supported anywhere in the
        // engine; mirror the tree-walker's `#NAME?` surfacing.
        RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Name)));
        ++pc;
        break;
      }

      case OpCode::LoadLet: {
        if (ins.a >= s.let_slots.size()) {
          return make_vm_error(FormulonErrorCode::kVmLetSlotMissing, "LoadLet slot not populated");
        }
        RETURN_IF_ERROR(push_value(s, s.let_slots[ins.a]));
        ++pc;
        break;
      }

      case OpCode::StoreLet: {
        RETURN_IF_ERROR(require_stack_depth(s, 1));
        if (ins.a >= s.let_slots.size()) {
          s.let_slots.resize(ins.a + 1U, Value::blank());
        }
        s.let_slots[ins.a] = s.stack.back();
        s.stack.pop_back();
        ++pc;
        break;
      }

      case OpCode::LoadLambdaArg: {
        if (s.frames.empty() || ins.a >= s.frames.back().args.size()) {
          return make_vm_error(FormulonErrorCode::kVmLetSlotMissing, "LoadLambdaArg outside an active lambda frame");
        }
        RETURN_IF_ERROR(push_value(s, s.frames.back().args[ins.a]));
        ++pc;
        break;
      }

      case OpCode::Call: {
        const std::uint32_t arity = ins.b;
        RETURN_IF_ERROR(require_stack_depth(s, arity));
        auto name_p = name_at(bc, ins.a);
        if (!name_p) {
          return name_p.error();
        }
        const std::string_view name = *name_p.value();
        // Lazy IFERROR / IFNA: the compiler emits these as eager Calls with
        // both arguments already on the stack. Inspect the primary post-hoc
        // and pick either the primary or the fallback. This drifts from the
        // tree-walker's true short-circuit only when the *fallback* itself
        // raises a different error and the primary is fine; the parity
        // corpus avoids that pattern.
        const LazyCallKind lazy = classify_lazy_call(name);
        if (lazy != LazyCallKind::None) {
          if (arity != 2) {
            // Pop and replace with #VALUE! to match the tree-walker's
            // arity-check failure surface.
            pop_values(s, arity);
            RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
            ++pc;
            break;
          }
          const Value fallback = s.stack.back();
          s.stack.pop_back();
          const Value primary = s.stack.back();
          s.stack.pop_back();
          Value out = primary;
          if (lazy == LazyCallKind::IfError) {
            if (primary.is_error()) {
              out = fallback;
            }
          } else {  // IfNa
            const bool primary_is_na = primary.is_error() && primary.as_error() == ErrorCode::NA;
            if (primary_is_na) {
              out = fallback;
            }
            // Mirror eval_ifna_lazy's blank->0 coercion for both paths.
            if (out.is_blank()) {
              out = Value::number(0.0);
            }
          }
          RETURN_IF_ERROR(push_value(s, out));
          ++pc;
          break;
        }
        // Eager dispatch through the function registry. Argument-AST
        // introspection (range-aware aggregators, AST-driven lazy forms
        // like CHOOSE / OFFSET / SUMIF) is not represented in the bytecode
        // IR; the VM simply hands the registry impl whatever values arrived
        // on the stack. Bundle 5.3+ optimisation may bring the IR closer to
        // tree-walker parity for those families.
        const FunctionDef* def = registry.lookup(name);
        if (def == nullptr) {
          pop_values(s, arity);
          RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Name)));
          ++pc;
          break;
        }
        if (arity < def->min_arity || arity > def->max_arity) {
          pop_values(s, arity);
          RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
          ++pc;
          break;
        }
        // Build the argv vector from the top `arity` stack slots, oldest
        // first (i.e. left-to-right argument order). For range-aware
        // aggregators (`accepts_ranges == true`), Value::Array arguments
        // are flattened in row-major order with the same provenance-aware
        // filtering the tree-walker applies to RangeOp / SpillRef args.
        // This is the only place the VM mirrors the tree-walker's
        // per-arg expansion logic — the bytecode IR doesn't represent
        // RangeOp endpoints, so a `=SUM({1,2,3})` style array literal is
        // the principal flattening surface for the VM.
        std::vector<Value> argv;
        argv.reserve(arity);
        for (std::uint32_t i = 0; i < arity; ++i) {
          const Value& slot = s.stack[s.stack.size() - arity + i];
          if (def->accepts_ranges && slot.is_array()) {
            const ArrayValue* arr = slot.as_array();
            const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
            for (std::size_t k = 0; k < n; ++k) {
              const Value& cell = arr->cells[k];
              if (def->range_filter_numeric_only && cell.kind() != ValueKind::Number) {
                continue;
              }
              if (def->range_filter_bool_coercible && cell.kind() != ValueKind::Number &&
                  cell.kind() != ValueKind::Bool) {
                continue;
              }
              if (def->range_filter_a_coerce) {
                if (cell.kind() == ValueKind::Blank) {
                  continue;
                }
                if (cell.kind() == ValueKind::Bool) {
                  argv.push_back(Value::number(cell.as_boolean() ? 1.0 : 0.0));
                  continue;
                }
                if (cell.kind() == ValueKind::Text) {
                  argv.push_back(Value::number(0.0));
                  continue;
                }
              }
              argv.push_back(cell);
            }
          } else {
            argv.push_back(slot);
          }
        }
        pop_values(s, arity);
        // Default tree-walker rule: short-circuit on the first error arg.
        // Functions that opt out (`propagate_errors == false`) must inspect
        // their error inputs; we pass the raw values through.
        if (def->propagate_errors) {
          bool short_circuit = false;
          Value err_v = Value::blank();
          for (const Value& a : argv) {
            if (a.is_error()) {
              err_v = a;
              short_circuit = true;
              break;
            }
          }
          if (short_circuit) {
            RETURN_IF_ERROR(push_value(s, err_v));
            ++pc;
            break;
          }
        }
        // Pass the post-flattening size (may exceed the original `arity`
        // when an Array argument was unpacked).
        const Value out = def->impl(argv.data(), static_cast<std::uint32_t>(argv.size()), arena);
        RETURN_IF_ERROR(push_value(s, out));
        ++pc;
        break;
      }

      case OpCode::CallLambda: {
        const std::uint32_t arity = ins.a;
        RETURN_IF_ERROR(require_stack_depth(s, arity + 1U));
        // Stack layout: [..., callee, arg0, ..., arg{arity-1}].
        std::vector<Value> args;
        args.reserve(arity);
        for (std::uint32_t i = 0; i < arity; ++i) {
          args.push_back(s.stack[s.stack.size() - arity + i]);
        }
        pop_values(s, arity);
        const Value callee = s.stack.back();
        s.stack.pop_back();
        if (!callee.is_lambda()) {
          if (callee.is_error()) {
            RETURN_IF_ERROR(push_value(s, callee));
          } else {
            RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
          }
          ++pc;
          break;
        }
        const LambdaValue* lv = callee.as_lambda();
        const std::uint32_t required = lv->param_count - lv->optional_count;
        if (arity < required || arity > lv->param_count) {
          RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
          ++pc;
          break;
        }
        // Excel-style left-to-right argument evaluation already happened
        // when the args were pushed; first error short-circuits the call.
        bool err_short = false;
        Value err_v = Value::blank();
        for (const Value& a : args) {
          if (a.is_error()) {
            err_short = true;
            err_v = a;
            break;
          }
        }
        if (err_short) {
          RETURN_IF_ERROR(push_value(s, err_v));
          ++pc;
          break;
        }
        // Pad missing trailing args with Blank to fill the optional slots.
        while (args.size() < lv->param_count) {
          args.push_back(Value::blank());
        }
        // Recurse into the body. We reuse the parent's `let_slots` (the
        // compiler-issued slot numbering is global within a single
        // compile() output, so reuse is safe). The body's `Return` exits
        // the recursive dispatch and leaves its result on top of stack
        // here.
        const auto* closure = reinterpret_cast<const VmClosure*>(lv->body);
        VmState::Frame frame;
        frame.return_pc = pc + 1U;
        frame.args = std::move(args);
        s.frames.push_back(std::move(frame));
        auto sub = dispatch(bc, arena, registry, ctx, s, closure->body_pc);
        if (!sub) {
          return sub.error();
        }
        // The body's Return left its result already pushed on the stack
        // (Return semantics below); pop our frame and continue.
        s.frames.pop_back();
        ++pc;
        break;
      }

      case OpCode::BinaryOp: {
        RETURN_IF_ERROR(require_stack_depth(s, 2));
        const Value rhs = s.stack.back();
        s.stack.pop_back();
        const Value lhs = s.stack.back();
        s.stack.pop_back();
        const auto op = static_cast<parser::BinOp>(ins.a);
        // Mirror eval_node's BinaryOp branch: error short-circuits left-to-
        // right, scalar otherwise. (Array broadcast intentionally not
        // mirrored: the bytecode IR collapses RangeOp endpoints to scalars
        // before this opcode runs, so a Value::Array can only arise from
        // SpillRef / ArrayLiteral, which the parity corpus avoids.)
        if (lhs.is_error()) {
          RETURN_IF_ERROR(push_value(s, lhs));
          ++pc;
          break;
        }
        if (rhs.is_error()) {
          RETURN_IF_ERROR(push_value(s, rhs));
          ++pc;
          break;
        }
        Value out = Value::blank();
        switch (op) {
          case parser::BinOp::Add:
          case parser::BinOp::Sub:
          case parser::BinOp::Mul:
          case parser::BinOp::Div:
          case parser::BinOp::Pow: {
            auto ln = coerce_to_number(lhs);
            if (!ln) {
              out = Value::error(ln.error());
              break;
            }
            auto rn = coerce_to_number(rhs);
            if (!rn) {
              out = Value::error(rn.error());
              break;
            }
            out = apply_arithmetic(op, ln.value(), rn.value());
            break;
          }
          case parser::BinOp::Concat:
            out = apply_concat(lhs, rhs, arena);
            break;
          case parser::BinOp::Eq:
          case parser::BinOp::NotEq:
          case parser::BinOp::Lt:
          case parser::BinOp::LtEq:
          case parser::BinOp::Gt:
          case parser::BinOp::GtEq:
            out = apply_comparison(op, lhs, rhs);
            break;
        }
        RETURN_IF_ERROR(push_value(s, out));
        ++pc;
        break;
      }

      case OpCode::UnaryOp: {
        RETURN_IF_ERROR(require_stack_depth(s, 1));
        const Value operand = s.stack.back();
        s.stack.pop_back();
        const auto op = static_cast<parser::UnaryOp>(ins.a);
        Value out = operand.is_error() ? operand : apply_unary(op, operand);
        RETURN_IF_ERROR(push_value(s, out));
        ++pc;
        break;
      }

      case OpCode::Concat: {
        RETURN_IF_ERROR(require_stack_depth(s, 2));
        const Value rhs = s.stack.back();
        s.stack.pop_back();
        const Value lhs = s.stack.back();
        s.stack.pop_back();
        if (lhs.is_error()) {
          RETURN_IF_ERROR(push_value(s, lhs));
          ++pc;
          break;
        }
        if (rhs.is_error()) {
          RETURN_IF_ERROR(push_value(s, rhs));
          ++pc;
          break;
        }
        RETURN_IF_ERROR(push_value(s, apply_concat(lhs, rhs, arena)));
        ++pc;
        break;
      }

      case OpCode::MakeArray: {
        const std::uint32_t rows = ins.a;
        const std::uint32_t cols = ins.b;
        const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
        RETURN_IF_ERROR(require_stack_depth(s, n));
        Value* buf = arena.create_array<Value>(n);
        if (buf == nullptr) {
          return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "arena exhausted in MakeArray");
        }
        // Row-major: oldest stack slot is (0,0); top of stack is the last
        // cell.
        for (std::size_t i = 0; i < n; ++i) {
          buf[i] = s.stack[s.stack.size() - n + i];
        }
        pop_values(s, n);
        ArrayValue* arr = arena.create<ArrayValue>();
        if (arr == nullptr) {
          return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "arena exhausted in MakeArray");
        }
        arr->rows = rows;
        arr->cols = cols;
        arr->cells = buf;
        RETURN_IF_ERROR(push_value(s, Value::array(arr)));
        ++pc;
        break;
      }

      case OpCode::MakeLambda: {
        // Layout (from compile_lambda):
        //   pc:   MakeLambda name_start (optional<<16 | param_count)
        //   pc+1: Jump <Lend>      (skip body when the parent stream runs)
        //   pc+2: <body>
        //         Return
        //   Lend:
        //
        // The parent stream takes the Jump on the next iteration; we just
        // need to construct the closure value here. The body pc is pc + 2.
        const std::uint32_t param_count = ins.b & 0xFFFFu;
        const std::uint32_t optional_count = (ins.b >> 16) & 0xFFFFu;
        const std::uint32_t name_start = ins.a;

        auto* lv = arena.create<LambdaValue>();
        if (lv == nullptr) {
          return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "arena exhausted in MakeLambda");
        }
        std::string_view* params = nullptr;
        if (param_count > 0) {
          params = arena.create_array<std::string_view>(param_count);
          if (params == nullptr) {
            return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "arena exhausted in MakeLambda");
          }
          for (std::uint32_t i = 0; i < param_count; ++i) {
            const std::uint32_t idx = name_start + i;
            if (idx >= bc.names.size()) {
              return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "MakeLambda param name index out of range");
            }
            params[i] = intern_arena_string(arena, bc.names[idx]);
          }
        }
        lv->params = params;
        lv->param_count = param_count;
        lv->optional_count = optional_count;
        // Construct the VM closure record; cast its address into the
        // `body` slot. Same arena, so lifetime is unified.
        auto* closure = arena.create<VmClosure>();
        if (closure == nullptr) {
          return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "arena exhausted in MakeLambda");
        }
        closure->body_pc = pc + 2U;
        // NOTE: storing a non-AST pointer in `LambdaValue::body`. The VM
        // never derefs `body` as an AST node; the tree-walker never sees
        // VM-emitted lambdas. The cast-pun is contained to this TU.
        lv->body = reinterpret_cast<const parser::AstNode*>(closure);
        lv->captured_env = nullptr;
        RETURN_IF_ERROR(push_value(s, Value::lambda(lv)));
        ++pc;
        break;
      }

      case OpCode::Union: {
        // Excel comma-as-union (`(A1, B2)`) is not a Value-level operator
        // in scalar context; the tree-walker surfaces #VALUE! here. Mirror
        // that by popping the operands and pushing #VALUE!.
        const std::uint32_t count = ins.a;
        RETURN_IF_ERROR(require_stack_depth(s, count));
        pop_values(s, count);
        RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
        ++pc;
        break;
      }

      case OpCode::Intersect: {
        // Space-as-intersection: also requires Ref payloads the IR has
        // already discarded. Surface #VALUE! to mirror the tree-walker's
        // fallback when the operands aren't Refs.
        RETURN_IF_ERROR(require_stack_depth(s, 2));
        s.stack.pop_back();
        s.stack.pop_back();
        RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
        ++pc;
        break;
      }

      case OpCode::ImplicitIntersection: {
        // Identity for non-RangeOp operands matches the tree-walker's
        // pass-through; range collapse already happened during LoadRange.
        RETURN_IF_ERROR(require_stack_depth(s, 1));
        ++pc;
        break;
      }

      case OpCode::Jump: {
        if (ins.a >= code_size) {
          return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget, "Jump target out of range");
        }
        pc = ins.a;
        break;
      }

      case OpCode::JumpIfFalse: {
        RETURN_IF_ERROR(require_stack_depth(s, 1));
        const Value cond = s.stack.back();
        s.stack.pop_back();
        if (ins.a >= code_size) {
          return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget, "JumpIfFalse target out of range");
        }
        // Error in cond: do NOT take the branch; push the error back and
        // fall through so the taken-branch path can carry it. Tree-walker
        // mirror: `=IF(1/0,"a","b")` returns `#DIV/0!`.
        if (cond.is_error()) {
          RETURN_IF_ERROR(push_value(s, cond));
          // Skip the taken branch: jump straight to the merge point. The
          // compiler shape is `then; Jump Lend; Lfalse:`. Easiest mirror is
          // to force-take the false branch so the error flows through it;
          // the false branch will be the unconditional false-result,
          // `LoadConst FALSE`, which we then overwrite with the error after
          // the merge. To stay simple, we instead bypass the IF entirely:
          // pop the just-pushed error and re-push at the merge point. The
          // safest path is just propagating the error by leaving it on the
          // stack and jumping to the false target — the false branch will
          // execute and produce its own value, which we then have to
          // suppress. Net: route to the false target and drop the false
          // branch's output by replacing it with the error.
          //
          // Simpler: Just take the false branch (jump). After the IF
          // merges, replace the top of stack with the error. We can't do
          // that reliably without scanning, so we adopt the approach of
          // tree-walker BinaryOp short-circuit by aborting the IF: jump
          // unconditionally past the merge by walking to the next
          // instruction whose op is Jump (the unconditional one paired with
          // JumpIfFalse). Bytecode shape lets us do that: the JumpIfFalse
          // target IS the false branch start; right before it sits the
          // unconditional Jump whose target is the merge point. Read that
          // Jump target.
          //
          // ins.a = pc of false branch start. The instruction at ins.a - 1
          // (in the well-formed IR) is the unconditional Jump to merge.
          if (ins.a == 0u) {
            return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget, "IF false-branch target at pc 0");
          }
          const Instruction& merge_jump = bc.code[ins.a - 1U];
          if (merge_jump.op != OpCode::Jump || merge_jump.a >= code_size) {
            return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget,
                                 "IF false-branch is not preceded by a Jump-to-merge");
          }
          pc = merge_jump.a;
          break;
        }
        bool dummy = false;
        const bool truthy = coerce_to_truthy_for_jump(cond, &dummy);
        if (!truthy) {
          pc = ins.a;
        } else {
          ++pc;
        }
        break;
      }

      case OpCode::Return: {
        if (s.stack.empty()) {
          return make_vm_error(FormulonErrorCode::kVmStackUnderflow, "Return on empty stack");
        }
        if (s.frames.empty()) {
          return s.stack.back();
        }
        // Inside a lambda body: leave the result on the stack and unwind
        // back to the caller's CallLambda path.
        return s.stack.back();
      }

      case OpCode::Halt: {
        if (s.stack.empty()) {
          return make_vm_error(FormulonErrorCode::kVmStackUnderflow, "Halt on empty stack");
        }
        if (s.frames.empty()) {
          return s.stack.back();
        }
        return s.stack.back();
      }
    }
  }
  return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "fell off the end of the instruction stream");
}

}  // namespace

Expected<Value, Error> execute(const ByteCode& bc, Arena& eval_arena, const FunctionRegistry& registry,
                               const EvalContext& ctx) {
  if (bc.code.empty()) {
    return make_vm_error(FormulonErrorCode::kVmEmptyBytecode, "execute() called on an empty bytecode body");
  }
  VmState s;
  s.stack.reserve(64);
  return dispatch(bc, eval_arena, registry, ctx, s, /*start_pc=*/0U);
}

}  // namespace eval
}  // namespace formulon
