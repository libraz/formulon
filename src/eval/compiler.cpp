// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// AST -> ByteCode compiler implementation.
//
// See `compiler.h` for the public contract and `bytecode.h` for the IR
// shape. Notable design points:
//
//   - The compiler is purely additive: it never optimises (constant
//     folding / range canonicalisation live in `optimizer.{h,cpp}`,
//     Bundle 5.3) and never inspects the function registry. Lazy-family
//     dispatch (`IF`, `IFERROR`, `IFNA`) is hard-coded by name; everything
//     else lowers to an eager `Call`.
//   - Every operand is bounds-checked against the 24-bit `Instruction::a`
//     budget. Overflow surfaces as `kVm*Overflow` errors, never as a
//     truncated word.
//   - Lambda bodies are spliced inline: a `MakeLambda` is followed by a
//     `Jump` past the body bytes, then the body itself, then a `Halt` /
//     `Return`. The VM (Bundle 5.2) jumps over the body when executing
//     the parent stream and into it when the closure is called.
//   - `LET` allocates one slot per binding in declaration order. Slots
//     are scoped to the enclosing body; the compiler does not currently
//     reuse slots across non-overlapping bindings (a simple optimisation
//     we may add later).
//
// Equality-of-compiles guarantee: lowering is deterministic in AST shape
// only. Two compiles of two ASTs that are structurally equal up to node
// identity produce byte-identical `code` / `constants` / `names` / `refs`
// streams. The `source_pos` map is the only exception (see header).

#include "eval/compiler.h"

#include <cstdint>
#include <cstring>
#include <string>
#include <string_view>
#include <utility>
#include <vector>

#include "eval/bytecode.h"
#include "eval/optimizer.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

// `status_macros.h` uses `__COUNTER__`, which Emscripten's bundled clang
// flags as a C2y extension under `-Werror`. The compiler TU expands these
// macros heavily, so we re-implement the two helpers here using a
// `__LINE__`-derived name. The behaviour is identical to the shared
// macros; only the unique-id strategy differs. Scoped to this TU.
#define FM_COMP_CONCAT_INNER(a, b) a##b
#define FM_COMP_CONCAT(a, b) FM_COMP_CONCAT_INNER(a, b)
#define FM_COMP_UNIQUE(prefix) FM_COMP_CONCAT(prefix, __LINE__)

#define FM_RETURN_IF_ERROR(expr) \
  do {                           \
    auto _fm_status = (expr);    \
    if (!_fm_status) {           \
      return _fm_status.error(); \
    }                            \
  } while (0)

#define FM_ASSIGN_OR_RETURN_IMPL(tmp, lhs, expr) \
  auto tmp = (expr);                             \
  if (!tmp) {                                    \
    return tmp.error();                          \
  }                                              \
  lhs = std::move(tmp.value())

#define FM_ASSIGN_OR_RETURN(lhs, expr) FM_ASSIGN_OR_RETURN_IMPL(FM_COMP_UNIQUE(_fm_tmp_), lhs, expr)

namespace formulon {
namespace eval {

namespace {

/// Names recognised as lazy / short-circuit forms by the compiler. Any
/// other call name is lowered to an eager `Call` opcode.
///
/// The matching is ASCII case-insensitive (Excel-equivalent name-folding).
enum class LazyKind : std::uint8_t {
  None,
  If,
  IfError,
  IfNa,
};

/// Returns the lazy classification of a call name, or `LazyKind::None`
/// if the name does not match a short-circuit form.
LazyKind classify_lazy(std::string_view name) noexcept {
  // ASCII upper-case fold; allocation-free.
  auto eq_ci = [](std::string_view a, std::string_view b) noexcept {
    if (a.size() != b.size()) {
      return false;
    }
    for (std::size_t i = 0; i < a.size(); ++i) {
      char ca = a[i];
      char cb = b[i];
      if (ca >= 'a' && ca <= 'z') {
        ca = static_cast<char>(ca - ('a' - 'A'));
      }
      if (cb >= 'a' && cb <= 'z') {
        cb = static_cast<char>(cb - ('a' - 'A'));
      }
      if (ca != cb) {
        return false;
      }
    }
    return true;
  };
  if (eq_ci(name, "IF")) {
    return LazyKind::If;
  }
  if (eq_ci(name, "IFERROR")) {
    return LazyKind::IfError;
  }
  if (eq_ci(name, "IFNA")) {
    return LazyKind::IfNa;
  }
  return LazyKind::None;
}

/// LET binding scope: stack of `(name, slot)` pairs. Innermost binding
/// wins on lookup (Excel allows shadowing).
struct LetScope {
  std::vector<std::string_view> names;
  std::vector<std::uint32_t> slots;
};

/// Lambda parameter scope: declared parameters bind to argument slots
/// in declaration order. Like `LetScope`, innermost wins on lookup.
struct LambdaScope {
  std::vector<std::string_view> names;
  std::vector<std::uint32_t> slots;
};

/// Mutable per-body compilation state. One instance is created for the
/// outer formula body; lambdas push a fresh nested instance for their
/// own body.
struct BodyState {
  ByteCode* out = nullptr;
  std::uint32_t next_let_slot = 0;
  LetScope let_scope;
  LambdaScope lambda_scope;
};

/// Top-level compiler context. Owns scratch storage that persists across
/// nested-body lowering (e.g. the `arena` for temporary helpers).
struct CompilerContext {
  Arena* arena = nullptr;
  ByteCode bc;
};

struct LexicalBinding {
  enum class Kind : std::uint8_t {
    None,
    Let,
    LambdaArg,
  };

  Kind kind = Kind::None;
  std::uint32_t slot = 0;
};

Error make_compile_error(FormulonErrorCode code, const char* msg) {
  return make_error(code, std::string(msg));
}

Expected<std::uint32_t, Error> emit(BodyState& bs, const parser::AstNode& src, OpCode op, std::uint32_t a,
                                    std::uint32_t b);

LexicalBinding lookup_lexical_binding(const BodyState& bs, std::string_view name) noexcept {
  for (std::size_t i = bs.let_scope.names.size(); i > 0; --i) {
    if (bs.let_scope.names[i - 1] == name) {
      return LexicalBinding{LexicalBinding::Kind::Let, bs.let_scope.slots[i - 1]};
    }
  }
  for (std::size_t i = bs.lambda_scope.names.size(); i > 0; --i) {
    if (bs.lambda_scope.names[i - 1] == name) {
      return LexicalBinding{LexicalBinding::Kind::LambdaArg, bs.lambda_scope.slots[i - 1]};
    }
  }
  return {};
}

Expected<void, Error> emit_load_binding(BodyState& bs, const parser::AstNode& node, const LexicalBinding& binding) {
  if (binding.kind == LexicalBinding::Kind::Let) {
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadLet, binding.slot, 0));
    return {};
  }
  if (binding.kind == LexicalBinding::Kind::LambdaArg) {
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadLambdaArg, binding.slot, 0));
    return {};
  }
  return make_compile_error(FormulonErrorCode::kVmCompileFailed, "internal compiler error: missing lexical binding");
}

// --------------------------------------------------------------------------
// Pool helpers. Each returns the pool index of the freshly inserted entry
// (no de-duplication — the compiler is shape-deterministic, not
// canonicalising). Bounds-check against `Instruction::kMaxA`.
// --------------------------------------------------------------------------

Expected<std::uint32_t, Error> intern_text_for_constant(BodyState& bs, std::string_view text) {
  ByteCode& bc = *bs.out;
  bc.string_storage.emplace_back(text);
  // Index of the freshly inserted string; storage is reserved up-front so the
  // backing data() pointer remains stable for the lifetime of `bc`.
  return static_cast<std::uint32_t>(bc.string_storage.size() - 1);
}

Expected<std::uint32_t, Error> push_constant(BodyState& bs, Value v) {
  ByteCode& bc = *bs.out;
  if (bc.constants.size() > Instruction::kMaxA) {
    return make_compile_error(FormulonErrorCode::kVmConstPoolOverflow, "constants pool exceeds 24-bit operand budget");
  }
  // For Text values, intern the bytes into `string_storage` so the ByteCode
  // owns its character storage.
  if (v.kind() == ValueKind::Text) {
    FM_ASSIGN_OR_RETURN(auto idx, intern_text_for_constant(bs, v.as_text()));
    const std::string& slot = bc.string_storage[idx];
    v = Value::text(std::string_view(slot.data(), slot.size()));
  }
  bc.constants.push_back(v);
  return static_cast<std::uint32_t>(bc.constants.size() - 1);
}

Expected<std::uint32_t, Error> push_name(BodyState& bs, std::string_view name) {
  ByteCode& bc = *bs.out;
  if (bc.names.size() > Instruction::kMaxA) {
    return make_compile_error(FormulonErrorCode::kVmNamePoolOverflow, "names pool exceeds 24-bit operand budget");
  }
  bc.names.emplace_back(name);
  return static_cast<std::uint32_t>(bc.names.size() - 1);
}

// Stores a Reference into the refs pool, re-interning its sheet view into
// the ByteCode's string_storage so the resulting `ByteCode` is
// self-contained.
Expected<std::uint32_t, Error> push_ref(BodyState& bs, const parser::Reference& r) {
  ByteCode& bc = *bs.out;
  if (bc.refs.size() > Instruction::kMaxA) {
    return make_compile_error(FormulonErrorCode::kVmConstPoolOverflow, "refs pool exceeds 24-bit operand budget");
  }
  parser::Reference copy = r;
  if (!r.sheet.empty()) {
    bc.string_storage.emplace_back(r.sheet);
    const std::string& slot = bc.string_storage.back();
    copy.sheet = std::string_view(slot.data(), slot.size());
  } else {
    copy.sheet = std::string_view{};
  }
  bc.refs.push_back(copy);
  return static_cast<std::uint32_t>(bc.refs.size() - 1);
}

// --------------------------------------------------------------------------
// Instruction emit. Every emit grows `code` and `source_pos` in lockstep
// so the VM can map any pc back to a node identity.
// --------------------------------------------------------------------------

Expected<std::uint32_t, Error> emit(BodyState& bs, const parser::AstNode& src, OpCode op, std::uint32_t a = 0,
                                    std::uint32_t b = 0) {
  ByteCode& bc = *bs.out;
  if (a > Instruction::kMaxA) {
    return make_compile_error(FormulonErrorCode::kVmInstructionLimit, "instruction operand exceeds 24-bit budget");
  }
  if (bc.code.size() > Instruction::kMaxA) {
    return make_compile_error(FormulonErrorCode::kVmInstructionLimit, "code length exceeds 24-bit jump budget");
  }
  Instruction ins{};
  ins.op = op;
  ins.a = a;
  ins.b = b;
  bc.code.push_back(ins);
  // Use the low 32 bits of the AST node pointer as the source-position
  // attribution. Stable within a single parse but not across parses;
  // diagnostic emitters must not assume otherwise.
  const auto bits = reinterpret_cast<std::uintptr_t>(&src);
  bc.source_pos.push_back(static_cast<std::uint32_t>(bits));
  return static_cast<std::uint32_t>(bc.code.size() - 1);
}

// Forward declaration: the per-kind handlers all call back into compile_node.
Expected<void, Error> compile_node(BodyState& bs, const parser::AstNode& node);

// --------------------------------------------------------------------------
// Per-kind lowering helpers.
// --------------------------------------------------------------------------

Expected<void, Error> compile_literal(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto idx, push_constant(bs, node.as_literal()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadConst, idx));
  return {};
}

Expected<void, Error> compile_ref(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto idx, push_ref(bs, node.as_ref()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadRef, idx));
  return {};
}

Expected<void, Error> compile_spill_ref(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto idx, push_ref(bs, node.as_spill_ref()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadSpillRef, idx));
  return {};
}

Expected<void, Error> compile_external_ref(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto sheet_n, push_name(bs, node.as_external_ref_sheet()));
  FM_ASSIGN_OR_RETURN(auto refs_idx, push_ref(bs, node.as_external_ref_cell()));
  // Pack book_id (low 16) and refs_idx (high 16) into `b`; sheet_n in `a`.
  const std::uint32_t book_id = node.as_external_ref_book_id();
  if (book_id > 0xFFFFu || refs_idx > 0xFFFFu) {
    return make_compile_error(FormulonErrorCode::kVmCompileFailed,
                              "external ref book_id or refs index exceeds 16-bit budget");
  }
  const std::uint32_t b = (refs_idx << 16) | book_id;
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadExternalRef, sheet_n, b));
  return {};
}

Expected<void, Error> compile_structured_ref(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto table_n, push_name(bs, node.as_structured_ref_table()));
  FM_ASSIGN_OR_RETURN(auto col_n, push_name(bs, node.as_structured_ref_column()));
  if (col_n > 0xFFFFu) {
    return make_compile_error(FormulonErrorCode::kVmNamePoolOverflow,
                              "structured-ref column name index exceeds 16 bits");
  }
  const std::uint32_t modifier = static_cast<std::uint32_t>(node.as_structured_ref_modifier());
  // a = table name, b = (modifier << 16) | column-name index.
  const std::uint32_t b = (modifier << 16) | col_n;
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadStructRef, table_n, b));
  return {};
}

Expected<void, Error> compile_name_ref(BodyState& bs, const parser::AstNode& node) {
  std::string_view name = node.as_name();
  // Resolve against active LET / Lambda scope first (innermost wins).
  const LexicalBinding binding = lookup_lexical_binding(bs, name);
  if (binding.kind != LexicalBinding::Kind::None) {
    FM_RETURN_IF_ERROR(emit_load_binding(bs, node, binding));
    return {};
  }
  // Fall through to a workbook-scope name lookup at run time.
  FM_ASSIGN_OR_RETURN(auto idx, push_name(bs, name));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadName, idx));
  return {};
}

Expected<void, Error> compile_unary(BodyState& bs, const parser::AstNode& node) {
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_unary_operand()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::UnaryOp, static_cast<std::uint32_t>(node.as_unary_op())));
  return {};
}

Expected<void, Error> compile_binary(BodyState& bs, const parser::AstNode& node) {
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_binary_lhs()));
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_binary_rhs()));
  const auto op = node.as_binary_op();
  if (op == parser::BinOp::Concat) {
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Concat));
    return {};
  }
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::BinaryOp, static_cast<std::uint32_t>(op)));
  return {};
}

Expected<void, Error> compile_range(BodyState& bs, const parser::AstNode& node) {
  // Range endpoints may be arbitrary subexpressions (e.g. `OFFSET(...):B5`),
  // but in the common case they are bare cell refs. We compile each endpoint
  // recursively; the VM (Bundle 5.2) is responsible for promoting two
  // single-cell refs on the stack into a rectangular range value via the
  // `LoadRange` opcode reading both stack slots.
  //
  // For the simple case `Ref:Ref` we emit a single `LoadRange` with two
  // adjacent refs-pool entries; the more general case falls back to
  // emitting a `BinaryOp` style op which the VM handles via its range
  // operator. Because the VM is not yet implemented we keep the IR shape
  // simple: always go through the two-subexpression path and let the VM
  // see two independent operands followed by a reserved BinaryOp slot.
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_range_lhs()));
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_range_rhs()));
  // Range is encoded as a special BinaryOp variant. The placeholder uses
  // `0xFF` as the op code so the VM can dispatch on it without colliding
  // with `parser::BinOp` values. A future revision may switch to a
  // dedicated `LoadRange` form once the VM lands.
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadRange, 0xFFu));
  return {};
}

Expected<void, Error> compile_union(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t arity = node.as_union_arity();
  for (std::uint32_t i = 0; i < arity; ++i) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_union_child(i)));
  }
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Union, arity));
  return {};
}

Expected<void, Error> compile_intersect(BodyState& bs, const parser::AstNode& node) {
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_intersect_lhs()));
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_intersect_rhs()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Intersect));
  return {};
}

Expected<void, Error> compile_implicit_intersection(BodyState& bs, const parser::AstNode& node) {
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_implicit_intersection_operand()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::ImplicitIntersection));
  return {};
}

// IF(cond, then, else?) — short-circuit lowering:
//   compile(cond)
//   JumpIfFalse Lfalse
//   compile(then)
//   Jump Lend
//   Lfalse: compile(else)        ; or LoadConst FALSE if else is omitted
//   Lend:
Expected<void, Error> compile_if(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t arity = node.as_call_arity();
  if (arity < 2 || arity > 3) {
    return make_compile_error(FormulonErrorCode::kVmCompileFailed, "IF requires 2 or 3 arguments");
  }
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(0)));
  FM_ASSIGN_OR_RETURN(auto jif_pc, emit(bs, node, OpCode::JumpIfFalse, 0));
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(1)));
  FM_ASSIGN_OR_RETURN(auto jmp_pc, emit(bs, node, OpCode::Jump, 0));
  // Patch JumpIfFalse to land here.
  bs.out->code[jif_pc].a = static_cast<std::uint32_t>(bs.out->code.size());
  if (arity == 3) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(2)));
  } else {
    // IF without else returns FALSE per Excel semantics.
    FM_ASSIGN_OR_RETURN(auto idx, push_constant(bs, Value::boolean(false)));
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadConst, idx));
  }
  // Patch the unconditional jump to land here.
  bs.out->code[jmp_pc].a = static_cast<std::uint32_t>(bs.out->code.size());
  return {};
}

// IFERROR(expr, fallback) / IFNA(expr, fallback).
//
// These need a runtime "is this value an error of kind X?" check that the
// generic JumpIfFalse cannot express; we therefore emit a Call to the
// dedicated lazy form. The Call carries the function name in the names
// pool exactly as a regular call would, but with a dedicated opcode
// `Call` and arity == 2. The VM dispatch table (Bundle 5.2) will route
// these names to a lazy implementation that re-evaluates only the
// fallback when the error is suppressed.
//
// Rationale for not using JumpIfFalse here: Excel's IFERROR / IFNA do not
// short-circuit on a boolean coercion of the first arg; they short-circuit
// on whether the *value* is an error (and which error). The reference VM
// will detect this case by an internal "trapping eval" of the first
// argument, which is implementation-defined; Bundle 5.1's job is just to
// preserve the call shape so the VM can pick the right strategy.
Expected<void, Error> compile_iferror_or_ifna(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t arity = node.as_call_arity();
  if (arity != 2) {
    return make_compile_error(FormulonErrorCode::kVmCompileFailed, "IFERROR / IFNA require exactly 2 arguments");
  }
  for (std::uint32_t i = 0; i < arity; ++i) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(i)));
  }
  FM_ASSIGN_OR_RETURN(auto name_idx, push_name(bs, node.as_call_name()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Call, name_idx, arity));
  return {};
}

Expected<void, Error> compile_call(BodyState& bs, const parser::AstNode& node) {
  const auto kind = classify_lazy(node.as_call_name());
  if (kind == LazyKind::If) {
    return compile_if(bs, node);
  }
  if (kind == LazyKind::IfError || kind == LazyKind::IfNa) {
    return compile_iferror_or_ifna(bs, node);
  }
  // Name-bound lambda dispatch: when the call's name matches a LET slot or a
  // lambda parameter that resolves at compile time to the lexical scope,
  // emit a LoadLet / LoadLambdaArg followed by `CallLambda` so the runtime
  // closure value is invoked directly. This mirrors the tree-walker's
  // `dispatch_call` lookup against `ctx.name_env()` for `=LET(f, LAMBDA(...),
  // f(7))` style formulas. Names that resolve to neither scope fall through
  // to the generic registry-driven `Call` opcode.
  const std::string_view call_name = node.as_call_name();
  const LexicalBinding binding = lookup_lexical_binding(bs, call_name);
  if (binding.kind != LexicalBinding::Kind::None) {
    FM_RETURN_IF_ERROR(emit_load_binding(bs, node, binding));
    const std::uint32_t lambda_arity = node.as_call_arity();
    for (std::uint32_t k = 0; k < lambda_arity; ++k) {
      FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(k)));
    }
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::CallLambda, lambda_arity));
    return {};
  }
  const std::uint32_t arity = node.as_call_arity();
  for (std::uint32_t i = 0; i < arity; ++i) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_call_arg(i)));
  }
  FM_ASSIGN_OR_RETURN(auto name_idx, push_name(bs, node.as_call_name()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Call, name_idx, arity));
  return {};
}

Expected<void, Error> compile_array_literal(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t rows = node.as_array_rows();
  const std::uint32_t cols = node.as_array_cols();
  for (std::uint32_t r = 0; r < rows; ++r) {
    for (std::uint32_t c = 0; c < cols; ++c) {
      FM_RETURN_IF_ERROR(compile_node(bs, node.as_array_element(r, c)));
    }
  }
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::MakeArray, rows, cols));
  return {};
}

Expected<void, Error> compile_let(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t bindings = node.as_let_binding_count();
  // Snapshot the LET scope so we can pop our slots when the body is done.
  const std::size_t saved = bs.let_scope.names.size();
  for (std::uint32_t i = 0; i < bindings; ++i) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_let_binding_expr(i)));
    if (bs.next_let_slot > Instruction::kMaxA) {
      return make_compile_error(FormulonErrorCode::kVmLetSlotOverflow, "LET slot index exceeds 24-bit budget");
    }
    const std::uint32_t slot = bs.next_let_slot++;
    FM_RETURN_IF_ERROR(emit(bs, node, OpCode::StoreLet, slot));
    bs.let_scope.names.push_back(node.as_let_binding_name(i));
    bs.let_scope.slots.push_back(slot);
  }
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_let_body()));
  // Pop the bindings we added (later siblings should not see them).
  bs.let_scope.names.resize(saved);
  bs.let_scope.slots.resize(saved);
  return {};
}

Expected<void, Error> compile_lambda(BodyState& bs, const parser::AstNode& node) {
  const std::uint32_t param_count = node.as_lambda_param_count();
  const std::uint32_t optional_count = node.as_lambda_optional_count();
  if (param_count > 0xFFFFu || optional_count > 0xFFFFu) {
    return make_compile_error(FormulonErrorCode::kVmLambdaParamOverflow,
                              "lambda parameter count exceeds 16-bit budget");
  }
  // Push parameter names so a NameRef inside the body resolves to a
  // LoadLambdaArg. Slots are 0..param_count-1.
  const std::size_t saved = bs.lambda_scope.names.size();
  for (std::uint32_t i = 0; i < param_count; ++i) {
    bs.lambda_scope.names.push_back(node.as_lambda_param(i));
    bs.lambda_scope.slots.push_back(i);
  }
  // Stash the param-array start index in the names pool so the VM can read
  // the parameter names back out (useful for error messages and TCO arg
  // rebinding). The first emitted index is the array start.
  std::uint32_t name_start = 0;
  for (std::uint32_t i = 0; i < param_count; ++i) {
    FM_ASSIGN_OR_RETURN(auto idx, push_name(bs, node.as_lambda_param(i)));
    if (i == 0) {
      name_start = idx;
    }
  }
  if (param_count == 0) {
    // No params: anchor name_start at the next-allocated slot for clarity.
    name_start = static_cast<std::uint32_t>(bs.out->names.size());
  }

  // Emit `MakeLambda` followed by an unconditional Jump over the body bytes.
  // The body bytes are spliced inline so the parent stream skips them at
  // run time. Layout:
  //
  //   MakeLambda name_start (param_count | optional_count<<16)
  //   Jump <Lend>
  //   <body>
  //   Halt          ; or Return
  //   Lend:
  FM_ASSIGN_OR_RETURN(auto mk_pc, emit(bs, node, OpCode::MakeLambda, name_start, (optional_count << 16) | param_count));
  (void)mk_pc;  // The VM walks the next two slots; no explicit reference needed.
  FM_ASSIGN_OR_RETURN(auto jmp_pc, emit(bs, node, OpCode::Jump, 0));
  // Body starts at the next pc.
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_lambda_body()));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::Return));
  bs.out->code[jmp_pc].a = static_cast<std::uint32_t>(bs.out->code.size());

  // Pop the params we pushed.
  bs.lambda_scope.names.resize(saved);
  bs.lambda_scope.slots.resize(saved);
  return {};
}

Expected<void, Error> compile_lambda_call(BodyState& bs, const parser::AstNode& node) {
  // Compile the callee first (it ends up below the args on the stack), then
  // the args left-to-right, then `CallLambda arity`.
  FM_RETURN_IF_ERROR(compile_node(bs, node.as_lambda_call_callee()));
  const std::uint32_t arity = node.as_lambda_call_arity();
  for (std::uint32_t i = 0; i < arity; ++i) {
    FM_RETURN_IF_ERROR(compile_node(bs, node.as_lambda_call_arg(i)));
  }
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::CallLambda, arity));
  return {};
}

Expected<void, Error> compile_error_literal(BodyState& bs, const parser::AstNode& node) {
  FM_ASSIGN_OR_RETURN(auto idx, push_constant(bs, Value::error(node.as_error_literal())));
  FM_RETURN_IF_ERROR(emit(bs, node, OpCode::LoadConst, idx));
  return {};
}

// --------------------------------------------------------------------------
// Dispatch.
// --------------------------------------------------------------------------

Expected<void, Error> compile_node(BodyState& bs, const parser::AstNode& node) {
  switch (node.kind()) {
    case parser::NodeKind::Literal:
      return compile_literal(bs, node);
    case parser::NodeKind::Ref:
      return compile_ref(bs, node);
    case parser::NodeKind::SpillRef:
      return compile_spill_ref(bs, node);
    case parser::NodeKind::ExternalRef:
      return compile_external_ref(bs, node);
    case parser::NodeKind::Ref3D:
      // 3-D references are resolved by the tree-walker only; the bytecode
      // VM runs in parity mode and does not implement them.
      return make_compile_error(FormulonErrorCode::kVmUnsupportedNode, "3-D reference not supported by the VM");
    case parser::NodeKind::StructuredRef:
      return compile_structured_ref(bs, node);
    case parser::NodeKind::NameRef:
      return compile_name_ref(bs, node);
    case parser::NodeKind::UnaryOp:
      return compile_unary(bs, node);
    case parser::NodeKind::BinaryOp:
      return compile_binary(bs, node);
    case parser::NodeKind::RangeOp:
      return compile_range(bs, node);
    case parser::NodeKind::UnionOp:
      return compile_union(bs, node);
    case parser::NodeKind::IntersectOp:
      return compile_intersect(bs, node);
    case parser::NodeKind::ImplicitIntersection:
      return compile_implicit_intersection(bs, node);
    case parser::NodeKind::Call:
      return compile_call(bs, node);
    case parser::NodeKind::ArrayLiteral:
      return compile_array_literal(bs, node);
    case parser::NodeKind::Lambda:
      return compile_lambda(bs, node);
    case parser::NodeKind::LetBinding:
      return compile_let(bs, node);
    case parser::NodeKind::LambdaCall:
      return compile_lambda_call(bs, node);
    case parser::NodeKind::ErrorLiteral:
      return compile_error_literal(bs, node);
    case parser::NodeKind::ErrorPlaceholder:
      return make_compile_error(FormulonErrorCode::kVmUnsupportedNode,
                                "compiler reached ErrorPlaceholder; parser must report the underlying error first");
  }
  return make_compile_error(FormulonErrorCode::kVmUnsupportedNode, "unknown AST node kind");
}

}  // namespace

// --------------------------------------------------------------------------
// Public entry point.
// --------------------------------------------------------------------------

Expected<ByteCode, Error> compile(const parser::AstNode& root, Arena& arena) {
  CompilerContext ctx;
  ctx.arena = &arena;
  // Reserve generously so the borrow contract on `string_storage` (stable
  // pointers across the lifetime of the ByteCode) holds even for large
  // formulas; this is a heuristic upper bound, not a hard limit.
  ctx.bc.code.reserve(16);
  ctx.bc.constants.reserve(8);
  ctx.bc.names.reserve(8);
  ctx.bc.refs.reserve(4);
  ctx.bc.string_storage.reserve(16);
  ctx.bc.source_pos.reserve(16);

  BodyState bs;
  bs.out = &ctx.bc;
  FM_RETURN_IF_ERROR(compile_node(bs, root));
  FM_RETURN_IF_ERROR(emit(bs, root, OpCode::Return));
  return std::move(ctx.bc);
}

Expected<ByteCode, Error> compile_and_optimize(const parser::AstNode& root, Arena& arena) {
  FM_ASSIGN_OR_RETURN(auto bc, compile(root, arena));
  return optimize(std::move(bc), arena);
}

// --------------------------------------------------------------------------
// Equality helper (test-only).
// --------------------------------------------------------------------------

bool bytecode_shapes_equal(const ByteCode& lhs, const ByteCode& rhs) noexcept {
  if (lhs.code.size() != rhs.code.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.code.size(); ++i) {
    const auto& a = lhs.code[i];
    const auto& b = rhs.code[i];
    if (a.op != b.op || a.a != b.a || a.b != b.b) {
      return false;
    }
  }
  if (lhs.constants.size() != rhs.constants.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.constants.size(); ++i) {
    if (!(lhs.constants[i] == rhs.constants[i])) {
      return false;
    }
  }
  if (lhs.names.size() != rhs.names.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.names.size(); ++i) {
    if (lhs.names[i] != rhs.names[i]) {
      return false;
    }
  }
  if (lhs.refs.size() != rhs.refs.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.refs.size(); ++i) {
    const auto& a = lhs.refs[i];
    const auto& b = rhs.refs[i];
    if (a.sheet != b.sheet || a.col != b.col || a.row != b.row || a.col_abs != b.col_abs || a.row_abs != b.row_abs ||
        a.is_full_col != b.is_full_col || a.is_full_row != b.is_full_row) {
      return false;
    }
  }
  if (lhs.string_storage.size() != rhs.string_storage.size()) {
    return false;
  }
  for (std::size_t i = 0; i < lhs.string_storage.size(); ++i) {
    if (lhs.string_storage[i] != rhs.string_storage[i]) {
      return false;
    }
  }
  return true;
}

}  // namespace eval
}  // namespace formulon
