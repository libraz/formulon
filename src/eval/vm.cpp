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
//     compiler-assigned 24-bit slot index. Slot numbering is unique per
//     compile() output but is reused across recursion depths, so each
//     `CallLambda` snapshots the slot vector before recursing into a lambda
//     body and restores it afterwards. This isolates a self-recursive
//     LAMBDA's inner LET bindings from the caller's still-live ones.
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

#include "eval/array_alloc.h"
#include "eval/bytecode.h"
#include "eval/coerce.h"
#include "eval/compiler_emit.h"
#include "eval/datetime_lazy.h"
#include "eval/defined_name_resolve.h"
#include "eval/eval_context.h"
#include "eval/function_registry.h"
#include "eval/implicit_intersection.h"
#include "eval/lambda_value.h"
#include "eval/lazy_impls.h"
#include "eval/name_env.h"
#include "eval/range_args.h"
#include "eval/scalar_ops.h"
#include "eval/structured_ref.h"
#include "eval/tree_walker/dispatch.h"
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

// Hard cap on lambda-call recursion depth. A runaway self-recursive LAMBDA
// (e.g. `LET(f, LAMBDA(n, f(n)), f(0))`) re-enters `dispatch` once per
// CallLambda; without a cap the native call stack overflows and the process
// crashes. On overflow we surface `#CALC!` and stop recursing, matching the
// tree-walker's runaway-lambda contract. The value mirrors the tree-walker's
// `kMaxLambdaDepth` (see eval/tree_walker/depth_guard.h) so both evaluators
// reject the same recursion depth.
constexpr std::uint32_t kMaxLambdaDepth = 256;

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
// payload remains valid after the underlying buffer goes out of scope. Arena
// exhaustion surfaces as a `kOutOfMemory` VM fault rather than a silent empty
// string, so a truncated Text payload can never masquerade as a valid one.
Expected<std::string_view, Error> intern_arena_string(Arena& arena, std::string_view s) {
  if (s.empty()) {
    return std::string_view{};
  }
  char* buf = arena.create_array<char>(s.size());
  if (buf == nullptr) {
    return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted interning a string payload");
  }
  std::memcpy(buf, s.data(), s.size());
  return std::string_view(buf, s.size());
}

// Coerces a Value to an Excel-truthy boolean for `JumpIfFalse` using the
// same `coerce_to_bool` helper the tree-walker's `eval_if_lazy` runs, so the
// two evaluators agree on every IF-condition edge case. On a coercion failure
// (e.g. a non-numeric text condition like "hello"), `*out_err` is set to the
// surfaced error code and the boolean return is unspecified; the caller must
// short-circuit the IF with that error rather than picking a branch. Matches
// `=IF("hello", 1, 2)` -> `#VALUE!`.
bool coerce_to_truthy_for_jump(const Value& v, ErrorCode* out_err, bool* out_has_err) {
  *out_has_err = false;
  auto b = coerce_to_bool(v);
  if (!b) {
    *out_has_err = true;
    *out_err = b.error();
    return true;  // unspecified; caller consults *out_has_err first
  }
  return b.value();
}

// Resolves a structured-ref bracket payload through the workbook. Mirrors
// tree_walker.cpp's `NodeKind::StructuredRef` branch and funnels the
// resolved rectangle through `EvalContext::expand_range` so cross-sheet
// resolution + cycle detection stays in one place.
Expected<Value, Error> resolve_struct_ref_op(std::string_view table_name, std::string_view column_payload, Arena& arena,
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
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted resolving a structured reference");
  }
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const std::vector<Value>& expanded = cells.value();
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = i < expanded.size() ? expanded[i] : Value::blank();
  }
  return Value::array(arr);
}

// Mirror of tree_walker.cpp's SpillRef branch: resolve the spill region at
// the anchor and project it as a `Value::Array`.
Expected<Value, Error> resolve_spill_ref_op(const parser::Reference& r, Arena& arena, const EvalContext& ctx) {
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
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(region->rows, region->cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted resolving a spilled reference");
  }
  const std::size_t n = static_cast<std::size_t>(region->rows) * static_cast<std::size_t>(region->cols);
  for (std::size_t i = 0; i < n; ++i) {
    buffer[i] = region->cells[i];
  }
  return Value::array(arr);
}

// Expands a bare `Ref:Ref` rectangle into a `Value::Array` in row-major
// order. Mirrors the tree-walker dispatcher's RangeOp argument expansion so
// range-aware aggregators (`SUM(A1:B2)`) see the same cells. Endpoint
// ordering is normalised by `expand_range`; Excel-visible faults (`#REF!`
// for out-of-range / whole-column endpoints) flow back as an error `Value`
// while arena exhaustion surfaces as a `kOutOfMemory` VM fault.
Expected<Value, Error> resolve_range_op(const parser::Reference& lhs, const parser::Reference& rhs, Arena& arena,
                                        const FunctionRegistry& registry, const EvalContext& ctx) {
  auto cells = ctx.expand_range(lhs, rhs, arena, registry);
  if (!cells) {
    return Value::error(cells.error());
  }
  const std::uint32_t r1 = lhs.row < rhs.row ? lhs.row : rhs.row;
  const std::uint32_t r2 = lhs.row < rhs.row ? rhs.row : lhs.row;
  const std::uint32_t c1 = lhs.col < rhs.col ? lhs.col : rhs.col;
  const std::uint32_t c2 = lhs.col < rhs.col ? rhs.col : lhs.col;
  const std::uint32_t rows = r2 - r1 + 1U;
  const std::uint32_t cols = c2 - c1 + 1U;
  Value* buffer = nullptr;
  ArrayValue* arr = allocate_array_value(rows, cols, arena, buffer, kMaxDerivedArrayCells);
  if (arr == nullptr) {
    return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted expanding a range");
  }
  const std::size_t total = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
  const std::vector<Value>& expanded = cells.value();
  for (std::size_t i = 0; i < total; ++i) {
    buffer[i] = i < expanded.size() ? expanded[i] : Value::blank();
  }
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
          ASSIGN_OR_RETURN(auto interned, intern_arena_string(arena, pushed.as_text()));
          pushed = Value::text(interned);
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
        // `LoadRange` is a marker emitted after the two endpoint
        // subexpressions (see compile_range). For a bare `Ref:Ref` range the
        // two preceding instructions are `LoadRef`s whose `a` operands index
        // the refs pool; recover the endpoints from there and expand the
        // rectangle into a `Value::Array` so range-aware aggregators
        // (`SUM(A1:B2)`) receive every cell, matching the tree-walker's
        // RangeOp argument expansion. The endpoint values the two `LoadRef`s
        // already pushed are scalar and get discarded here.
        RETURN_IF_ERROR(require_stack_depth(s, 2));
        if (pc < 2U || bc.code[pc - 1U].op != OpCode::LoadRef || bc.code[pc - 2U].op != OpCode::LoadRef) {
          // Complex range endpoints (e.g. `OFFSET(...):B5`) are not
          // reconstructable from the IR; mirror the tree-walker's #VALUE!
          // fallback for non-Ref range endpoints.
          s.stack.pop_back();
          s.stack.pop_back();
          RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
          ++pc;
          break;
        }
        auto lo = ref_at(bc, bc.code[pc - 2U].a);
        if (!lo) {
          return lo.error();
        }
        auto hi = ref_at(bc, bc.code[pc - 1U].a);
        if (!hi) {
          return hi.error();
        }
        s.stack.pop_back();
        s.stack.pop_back();
        ASSIGN_OR_RETURN(auto range_val, resolve_range_op(*lo.value(), *hi.value(), arena, registry, ctx));
        RETURN_IF_ERROR(push_value(s, range_val));
        ++pc;
        break;
      }

      case OpCode::LoadName: {
        auto n = name_at(bc, ins.a);
        if (!n) {
          return n.error();
        }
        // Tree-walker mirror: a runtime NameEnv (LET / captured lambda env)
        // takes precedence; otherwise resolve a workbook / sheet-scoped
        // defined name through the shared helper so the VM and tree-walker
        // agree. An unbound identifier surfaces #NAME? from the resolver.
        Value v = Value::error(ErrorCode::Name);
        bool bound_in_env = false;
        if (const NameEnv* env = ctx.name_env(); env != nullptr) {
          if (const Value* bound = env->lookup(*n.value()); bound != nullptr) {
            v = *bound;
            bound_in_env = true;
          }
        }
        if (!bound_in_env) {
          v = resolve_defined_name(*n.value(), arena, registry, ctx);
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
        ASSIGN_OR_RETURN(auto v, resolve_struct_ref_op(*table_n.value(), *col_n.value(), arena, registry, ctx));
        RETURN_IF_ERROR(push_value(s, v));
        ++pc;
        break;
      }

      case OpCode::LoadSpillRef: {
        auto r = ref_at(bc, ins.a);
        if (!r) {
          return r.error();
        }
        ASSIGN_OR_RETURN(auto spill_val, resolve_spill_ref_op(*r.value(), arena, ctx));
        RETURN_IF_ERROR(push_value(s, spill_val));
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
        // A visible defined name shadows the lazy table and registry just as
        // it does in the tree-walker. The compiler has already evaluated the
        // argument expressions, so resolve the definition and invoke an
        // AST-backed LambdaValue through the value-based bridge. VM-internal
        // closure records never enter this path; they are reached only by
        // CallLambda and remain isolated from the tree-walker helper.
        if (find_defined_name(ctx, name) != nullptr) {
          std::vector<Value> defined_args;
          defined_args.reserve(arity);
          for (std::uint32_t i = 0; i < arity; ++i) {
            defined_args.push_back(s.stack[s.stack.size() - arity + i]);
          }
          pop_values(s, arity);
          const Value defined = resolve_defined_name(name, arena, registry, ctx);
          Value out = defined;
          if (!defined.is_error()) {
            if (!defined.is_lambda()) {
              out = Value::error(ErrorCode::Value);
            } else {
              // `execute()` historically permits an uninstrumented context;
              // provide local depth counters for tree-walker evaluation of a
              // named Lambda so recursive definitions still terminate. When
              // the caller supplied counters, preserve those instead.
              std::uint32_t eval_depth = 0;
              std::uint32_t lambda_depth = 0;
              const EvalContext lambda_ctx = ctx.with_depth_counters(
                  ctx.eval_depth_counter() != nullptr ? ctx.eval_depth_counter() : &eval_depth,
                  ctx.lambda_depth_counter() != nullptr ? ctx.lambda_depth_counter() : &lambda_depth);
              out =
                  invoke_lambda_values(defined.as_lambda(), arity, defined_args.empty() ? nullptr : defined_args.data(),
                                       arena, registry, lambda_ctx);
            }
          }
          RETURN_IF_ERROR(push_value(s, out));
          ++pc;
          break;
        }
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
        // Date1904-sensitive calendar family (DATE / YEAR / EDATE / TODAY /
        // ...). These are NOT in the eager registry — they route through the
        // tree-walker's lazy table so the workbook date epoch reaches them.
        // The VM has no call AST, so it reuses the shared date1904-aware impl
        // directly with `ctx.date1904()`. They are scalar-only, so the args
        // are the top `arity` stack slots verbatim with the left-most error
        // short-circuiting (calendar functions never opt out of that rule).
        if (const DateEntry* date = find_date_entry(name)) {
          if (arity < date->min_arity || arity > date->max_arity) {
            pop_values(s, arity);
            RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
            ++pc;
            break;
          }
          std::vector<Value> date_argv;
          date_argv.reserve(arity);
          for (std::uint32_t i = 0; i < arity; ++i) {
            date_argv.push_back(s.stack[s.stack.size() - arity + i]);
          }
          pop_values(s, arity);
          bool date_short_circuit = false;
          Value date_err = Value::blank();
          for (const Value& a : date_argv) {
            if (a.is_error()) {
              date_err = a;
              date_short_circuit = true;
              break;
            }
          }
          if (date_short_circuit) {
            RETURN_IF_ERROR(push_value(s, date_err));
            ++pc;
            break;
          }
          // NOW / TODAY need the context's wall-clock reading rather than
          // arguments, so a pinned workbook stays deterministic here too.
          const Value out = date->clock_impl != nullptr
                                ? date->clock_impl(ctx.wall_clock(), ctx.date1904())
                                : date->impl(date_argv.empty() ? nullptr : date_argv.data(),
                                             static_cast<std::uint32_t>(date_argv.size()), arena, ctx.date1904());
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
        bool short_circuit = false;
        Value propagated_error = Value::blank();
        for (std::uint32_t i = 0; i < arity; ++i) {
          const Value& slot = s.stack[s.stack.size() - arity + i];
          // Keep error precedence in source order across mixed scalar and
          // array arguments. A scalar error belongs to this slot, so detect
          // it before moving on to a later range. Range cells are handled
          // at their own row-major position below.
          if (def->propagate_errors && slot.is_error()) {
            propagated_error = slot;
            short_circuit = true;
            break;
          }
          if (def->accepts_ranges && slot.is_array()) {
            const ArrayValue* arr = slot.as_array();
            const std::size_t n = static_cast<std::size_t>(arr->rows) * static_cast<std::size_t>(arr->cols);
            for (std::size_t k = 0; k < n; ++k) {
              if (!append_range_sourced_value(*def, arr->cells[k], &argv, &propagated_error)) {
                short_circuit = true;
                break;
              }
            }
            if (short_circuit) {
              break;
            }
          } else {
            argv.push_back(slot);
          }
        }
        pop_values(s, arity);
        if (short_circuit) {
          RETURN_IF_ERROR(push_value(s, propagated_error));
          ++pc;
          break;
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
        // Recursion-depth guard: each CallLambda re-enters `dispatch`, and
        // `s.frames` carries one entry per active lambda body. A runaway
        // self-recursive LAMBDA would otherwise overflow the native stack and
        // crash. Surface `#CALC!` and stop recursing once the depth would
        // exceed the cap, matching the tree-walker's runaway-lambda contract.
        if (s.frames.size() >= kMaxLambdaDepth) {
          RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Calc)));
          ++pc;
          break;
        }
        // Recurse into the body. LET slot numbering is unique per compile()
        // output but is *reused across recursion depths*: a self-recursive
        // LAMBDA whose body contains a LET re-runs the same StoreLet slots on
        // every level, so without isolation an inner call would clobber the
        // caller's still-live bindings (e.g. a binding read after the
        // recursive call would see the deepest write). Snapshot the LET
        // slots before recursing and restore them afterwards so each
        // invocation gets its own scope. The body's result is returned on the
        // operand stack, not via let_slots, so restoring here is safe.
        const auto* closure = reinterpret_cast<const VmClosure*>(lv->body);
        std::vector<Value> saved_let_slots = s.let_slots;
        VmState::Frame frame;
        frame.return_pc = pc + 1U;
        frame.args = std::move(args);
        s.frames.push_back(std::move(frame));
        auto sub = dispatch(bc, arena, registry, ctx, s, closure->body_pc);
        if (!sub) {
          return sub.error();
        }
        // The body's Return left its result already pushed on the stack
        // (Return semantics below); pop our frame, restore the caller's LET
        // bindings, and continue.
        s.frames.pop_back();
        s.let_slots = std::move(saved_let_slots);
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
        Value* buf = nullptr;
        ArrayValue* arr = allocate_array_value(rows, cols, arena, buf, kMaxDerivedArrayCells);
        if (arr == nullptr) {
          return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted in MakeArray");
        }
        const std::size_t n = static_cast<std::size_t>(rows) * static_cast<std::size_t>(cols);
        RETURN_IF_ERROR(require_stack_depth(s, n));
        // Row-major: oldest stack slot is (0,0); top of stack is the last
        // cell.
        for (std::size_t i = 0; i < n; ++i) {
          buf[i] = s.stack[s.stack.size() - n + i];
        }
        pop_values(s, n);
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
          return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted in MakeLambda");
        }
        std::string_view* params = nullptr;
        if (param_count > 0) {
          params = arena.create_array<std::string_view>(param_count);
          if (params == nullptr) {
            return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted in MakeLambda");
          }
          for (std::uint32_t i = 0; i < param_count; ++i) {
            const std::uint32_t idx = name_start + i;
            if (idx >= bc.names.size()) {
              return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "MakeLambda param name index out of range");
            }
            ASSIGN_OR_RETURN(params[i], intern_arena_string(arena, bc.names[idx]));
          }
        }
        lv->params = params;
        lv->param_count = param_count;
        lv->optional_count = optional_count;
        // Construct the VM closure record; cast its address into the
        // `body` slot. Same arena, so lifetime is unified.
        auto* closure = arena.create<VmClosure>();
        if (closure == nullptr) {
          return make_vm_error(FormulonErrorCode::kOutOfMemory, "arena exhausted in MakeLambda");
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
        RETURN_IF_ERROR(require_stack_depth(s, 1));
        // Mirror the tree-walker's implicit-intersection projection for a bare
        // `Ref:Ref` range operand (`@A1:A3`). That operand lowers to
        // `LoadRef; LoadRef; LoadRange`, so when the three preceding
        // instructions match that shape we recover the endpoints from the refs
        // pool and project onto the formula cell:
        //   * single-column range -> the formula row (or #VALUE! when outside);
        //   * single-row range    -> the formula column (or #VALUE!);
        //   * 2D range            -> the (row, col) intersection cell, or
        //                            #VALUE! when the formula cell is outside;
        //   * no formula-cell anchor -> degrade to the top-left cell.
        // Any other operand (single Ref, call, scalar) is identity
        // pass-through, matching the tree-walker's non-RangeOp branch.
        if (pc >= 3U && bc.code[pc - 1U].op == OpCode::LoadRange && bc.code[pc - 2U].op == OpCode::LoadRef &&
            bc.code[pc - 3U].op == OpCode::LoadRef) {
          auto lo = ref_at(bc, bc.code[pc - 3U].a);
          if (!lo) {
            return lo.error();
          }
          auto hi = ref_at(bc, bc.code[pc - 2U].a);
          if (!hi) {
            return hi.error();
          }
          const parser::Reference& ra = *lo.value();
          const parser::Reference& rb = *hi.value();
          const std::uint32_t r1 = ra.row < rb.row ? ra.row : rb.row;
          const std::uint32_t r2 = ra.row < rb.row ? rb.row : ra.row;
          const std::uint32_t c1 = ra.col < rb.col ? ra.col : rb.col;
          const std::uint32_t c2 = ra.col < rb.col ? rb.col : ra.col;
          s.stack.pop_back();  // discard the range Array pushed by LoadRange
          parser::Reference target{};
          target.sheet = ra.sheet;
          if (!ctx.has_formula_cell()) {
            // No anchor (top-level eval): degrade to the top-left cell, the
            // same fallback the tree-walker's RangeOp branch takes.
            target.row = r1;
            target.col = c1;
            RETURN_IF_ERROR(push_value(s, ctx.resolve_ref(target, arena, registry)));
            ++pc;
            break;
          }
          const std::uint32_t fr = ctx.formula_row();
          const std::uint32_t fc = ctx.formula_col();
          bool outside = false;
          if (c1 == c2) {
            outside = fr < r1 || fr > r2;
            target.row = fr;
            target.col = c1;
          } else if (r1 == r2) {
            outside = fc < c1 || fc > c2;
            target.row = r1;
            target.col = fc;
          } else {
            outside = fr < r1 || fr > r2 || fc < c1 || fc > c2;
            target.row = fr;
            target.col = fc;
          }
          if (outside) {
            RETURN_IF_ERROR(push_value(s, Value::error(ErrorCode::Value)));
          } else {
            RETURN_IF_ERROR(push_value(s, ctx.resolve_ref(target, arena, registry)));
          }
          ++pc;
          break;
        }
        // Calls and other non-static operands retain only their value in the
        // bytecode stream. Collapse an Array result to its top-left element
        // so VM behavior matches the tree walker and `_xlfn.SINGLE`.
        s.stack.back() = implicit_intersect_value(s.stack.back());
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
        // Coerce the condition through the shared `coerce_to_bool` helper.
        // A non-coercible condition (e.g. the text "hello") surfaces
        // `#VALUE!` exactly like the tree-walker's `eval_if_lazy`; we route
        // that error through the false-branch suppression path so the IF
        // result is the error, not a branch value.
        ErrorCode cond_err = ErrorCode::Value;
        bool cond_has_err = false;
        const bool truthy = coerce_to_truthy_for_jump(cond, &cond_err, &cond_has_err);
        if (cond_has_err) {
          if (ins.a == 0u) {
            return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget, "IF false-branch target at pc 0");
          }
          const Instruction& merge_jump = bc.code[ins.a - 1U];
          if (merge_jump.op != OpCode::Jump || merge_jump.a >= code_size) {
            return make_vm_error(FormulonErrorCode::kVmInvalidJumpTarget,
                                 "IF false-branch is not preceded by a Jump-to-merge");
          }
          RETURN_IF_ERROR(push_value(s, Value::error(cond_err)));
          pc = merge_jump.a;
          break;
        }
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

      default:
        // Every real opcode advances `pc` itself; an unrecognised opcode
        // would otherwise leave `pc` unchanged and spin the dispatch loop
        // forever. Fail fast instead of hanging.
        return make_vm_error(FormulonErrorCode::kVmInvalidOpcode, "unknown opcode in instruction stream");
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
