// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Per-opcode metadata table for the bytecode VM and compiler.
//
// `OpCode` itself lives in `bytecode.h` alongside the IR shape; that header
// already carries the canonical `opcode_name()` mnemonic helper. This file
// adds a richer, constexpr-indexed metadata layer reusable by:
//
//   * the VM's dispatch / validation paths (e.g. "does this opcode pop
//     a stack slot?"),
//   * the compiler's emit-side assertions (e.g. "if I emit a `Call`, am
//     I encoding the names-pool index in `a` and arity in `b`?"),
//   * any future disassembler or bytecode-equality tooling that wants
//     to know operand interpretation without re-implementing the per-
//     opcode switch.
//
// The table is a single `constexpr` array indexed by `OpCode`. Looking up
// metadata for an opcode is a single bounds-checked array read; the
// compiler will fold the call away for constant inputs.
//
// Header-only. Including this file pulls in `bytecode.h` for `OpCode`.

#ifndef FORMULON_EVAL_OPCODE_META_H_
#define FORMULON_EVAL_OPCODE_META_H_

#include <array>
#include <cstdint>

#include "eval/bytecode.h"

namespace formulon {
namespace eval {

/// Coarse classification of an opcode's role. One opcode belongs to at most
/// one class; the partition is exhaustive over the current opcode catalog.
///
/// The classification is intentionally coarse: it captures the dispatch
/// pattern (push a value, pop a value, branch, call, etc.) but not the
/// exact operand layout (which `OpcodeMeta` describes separately).
enum class OpClass : std::uint8_t {
  /// Pushes a value computed from one or more pool entries
  /// (`LoadConst` / `LoadRef` / `LoadName` / `LoadStructRef` /
  /// `LoadSpillRef` / `LoadExternalRef` / `LoadLet` / `LoadLambdaArg`).
  Load,
  /// Pops the top of stack and stores it into a slot (`StoreLet`).
  Store,
  /// Combines values already on the stack into a single result
  /// (`LoadRange` / `BinaryOp` / `UnaryOp` / `Concat` /
  /// `MakeArray` / `MakeLambda` / `Union` / `Intersect` /
  /// `ImplicitIntersection`).
  Combine,
  /// Routes through the function registry (`Call`) or a lambda value
  /// (`CallLambda`).
  Call,
  /// Branches the program counter (`Jump` / `JumpIfFalse`).
  Jump,
  /// Terminates the current frame (`Return` / `Halt`).
  Terminate,
};

/// How an opcode interprets its `Instruction::a` operand.
///
/// "Pool" entries name the side-pool the index refers to. "Inline" entries
/// carry a small integer (arity, op-tag, slot number) rather than a pool
/// reference. "Target" carries an absolute instruction index. "None" means
/// the operand is unused (the compiler always writes zero).
enum class OperandA : std::uint8_t {
  None,
  ConstantsIndex,
  NamesIndex,
  RefsIndex,
  Slot,
  Target,
  InlineCount,
  InlineOpTag,
};

/// How an opcode interprets its `Instruction::b` operand. Most opcodes do
/// not use `b`; the ones that do encode either a small inline count
/// (arity, rows/cols) or a packed bitfield documented inline in
/// `bytecode.h`.
enum class OperandB : std::uint8_t {
  None,
  InlineCount,
  Packed,
};

/// Per-opcode metadata record. Populated once at compile time via the
/// `kOpcodeMeta` table below. All fields are POD and `constexpr`.
struct OpcodeMeta {
  /// Display mnemonic, stable across runs. Identical to `opcode_name(op)`;
  /// duplicated here so the metadata record is self-contained.
  const char* name;
  /// Coarse role of the opcode in dispatch.
  OpClass cls;
  /// Operand-A semantics.
  OperandA a;
  /// Operand-B semantics.
  OperandB b;
};

namespace detail {

constexpr std::size_t kOpcodeTableSize = static_cast<std::size_t>(OpCode::Halt) + 1U;

/// Builds the metadata table. Defined inline so the compiler can fold any
/// `opcode_meta(op)` call with a constant `op` at the call site.
constexpr std::array<OpcodeMeta, kOpcodeTableSize> make_opcode_table() noexcept {
  std::array<OpcodeMeta, kOpcodeTableSize> t{};
  t[static_cast<std::size_t>(OpCode::LoadConst)] =
      OpcodeMeta{"LoadConst", OpClass::Load, OperandA::ConstantsIndex, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadRef)] =
      OpcodeMeta{"LoadRef", OpClass::Load, OperandA::RefsIndex, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadRange)] =
      OpcodeMeta{"LoadRange", OpClass::Combine, OperandA::RefsIndex, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadName)] =
      OpcodeMeta{"LoadName", OpClass::Load, OperandA::NamesIndex, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadStructRef)] =
      OpcodeMeta{"LoadStructRef", OpClass::Load, OperandA::NamesIndex, OperandB::Packed};
  t[static_cast<std::size_t>(OpCode::LoadSpillRef)] =
      OpcodeMeta{"LoadSpillRef", OpClass::Load, OperandA::RefsIndex, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadExternalRef)] =
      OpcodeMeta{"LoadExternalRef", OpClass::Load, OperandA::NamesIndex, OperandB::Packed};
  t[static_cast<std::size_t>(OpCode::LoadLet)] = OpcodeMeta{"LoadLet", OpClass::Load, OperandA::Slot, OperandB::None};
  t[static_cast<std::size_t>(OpCode::StoreLet)] =
      OpcodeMeta{"StoreLet", OpClass::Store, OperandA::Slot, OperandB::None};
  t[static_cast<std::size_t>(OpCode::LoadLambdaArg)] =
      OpcodeMeta{"LoadLambdaArg", OpClass::Load, OperandA::Slot, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Call)] =
      OpcodeMeta{"Call", OpClass::Call, OperandA::NamesIndex, OperandB::InlineCount};
  t[static_cast<std::size_t>(OpCode::CallLambda)] =
      OpcodeMeta{"CallLambda", OpClass::Call, OperandA::InlineCount, OperandB::None};
  t[static_cast<std::size_t>(OpCode::BinaryOp)] =
      OpcodeMeta{"BinaryOp", OpClass::Combine, OperandA::InlineOpTag, OperandB::None};
  t[static_cast<std::size_t>(OpCode::UnaryOp)] =
      OpcodeMeta{"UnaryOp", OpClass::Combine, OperandA::InlineOpTag, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Concat)] = OpcodeMeta{"Concat", OpClass::Combine, OperandA::None, OperandB::None};
  t[static_cast<std::size_t>(OpCode::MakeArray)] =
      OpcodeMeta{"MakeArray", OpClass::Combine, OperandA::InlineCount, OperandB::InlineCount};
  t[static_cast<std::size_t>(OpCode::MakeLambda)] =
      OpcodeMeta{"MakeLambda", OpClass::Combine, OperandA::NamesIndex, OperandB::Packed};
  t[static_cast<std::size_t>(OpCode::Union)] =
      OpcodeMeta{"Union", OpClass::Combine, OperandA::InlineCount, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Intersect)] =
      OpcodeMeta{"Intersect", OpClass::Combine, OperandA::None, OperandB::None};
  t[static_cast<std::size_t>(OpCode::ImplicitIntersection)] =
      OpcodeMeta{"ImplicitIntersection", OpClass::Combine, OperandA::None, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Jump)] = OpcodeMeta{"Jump", OpClass::Jump, OperandA::Target, OperandB::None};
  t[static_cast<std::size_t>(OpCode::JumpIfFalse)] =
      OpcodeMeta{"JumpIfFalse", OpClass::Jump, OperandA::Target, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Return)] =
      OpcodeMeta{"Return", OpClass::Terminate, OperandA::None, OperandB::None};
  t[static_cast<std::size_t>(OpCode::Halt)] = OpcodeMeta{"Halt", OpClass::Terminate, OperandA::None, OperandB::None};
  return t;
}

}  // namespace detail

/// The constexpr metadata table itself. Indexed by `static_cast<size_t>(op)`.
/// One row per defined opcode; the array length is fixed at compile time and
/// covers every value of `OpCode`.
inline constexpr std::array<OpcodeMeta, detail::kOpcodeTableSize> kOpcodeMeta = detail::make_opcode_table();

/// Returns the metadata record for `op`. Behaviour is undefined for invalid
/// opcode values (i.e. integer casts beyond `OpCode::Halt`); callers that
/// receive opcodes off the wire must validate the range before calling.
constexpr const OpcodeMeta& opcode_meta(OpCode op) noexcept {
  return kOpcodeMeta[static_cast<std::size_t>(op)];
}

/// Convenience: returns true when `op` is a branching instruction (`Jump`
/// or `JumpIfFalse`).
constexpr bool is_jump(OpCode op) noexcept {
  return opcode_meta(op).cls == OpClass::Jump;
}

/// Convenience: returns true when `op` is a Load (pushes a value built
/// from a pool entry or slot).
constexpr bool is_load(OpCode op) noexcept {
  return opcode_meta(op).cls == OpClass::Load;
}

/// Convenience: returns true when `op` is a function or lambda call.
constexpr bool is_call(OpCode op) noexcept {
  return opcode_meta(op).cls == OpClass::Call;
}

/// Convenience: returns true when `op` terminates the current frame.
constexpr bool is_terminate(OpCode op) noexcept {
  return opcode_meta(op).cls == OpClass::Terminate;
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_OPCODE_META_H_
