//
// Bytecode IR for the (forthcoming) stack-machine VM.
//
// `compile()` (in `compiler.h`) lowers a parser `AstNode` tree into a flat
// `ByteCode` body that the VM will consume. The IR is a small
// stack-oriented opcode set.
//
// Layout:
//   - `Instruction` is a 64-bit POD: `(opcode:8, a:24, b:32)`.
//     The `a`/`b` operands are unsigned indices into one of the side pools
//     (`constants`, `names`) or, for branch ops, an absolute instruction
//     index. Their interpretation is per-opcode and documented inline.
//   - `ByteCode` owns:
//       * `code` -- the linear instruction stream.
//       * `constants` -- a pool of `Value` literals lifted out of the
//         instruction stream. Reference variants (`Ref`, `Array`,
//         `Lambda`) are not stored here; only `Number`, `Bool`, `Text`,
//         `Error`, `Blank` literals appear.
//       * `names` -- a pool of identifier strings (function names, defined-
//         name targets, `LET` binding names, lambda parameter names, table
//         / column names for structured refs). Strings are deep-copied into
//         the pool so a `ByteCode` is self-contained.
//       * `refs` -- a pool of `parser::Reference` payloads for `LoadRef`
//         and the two endpoints of a range. Holds anchor cells for
//         `LoadSpillRef` as well.
//       * `source_pos` -- a parallel-to-`code` map of AST node identity
//         (we use node pointer hashes truncated to 32 bits) for diagnostic
//         attribution. Exact format is opaque to the VM; only the compiler
//         and diagnostic emitter agree on it.
//
// Trivially-copyable POD where possible; no exceptions, no virtuals, no
// `std::function`. The ByteCode body itself is not trivially copyable
// because it owns its `std::vector` storage, but the per-instruction word
// is.
//
// This file is header-only. The compiler (`compiler.{h,cpp}`) is the sole
// producer of `ByteCode` instances and the VM (Bundle 5.2) will be the
// sole consumer.

#ifndef FORMULON_EVAL_BYTECODE_H_
#define FORMULON_EVAL_BYTECODE_H_

#include <cstdint>
#include <deque>
#include <string>
#include <type_traits>
#include <vector>

#include "parser/reference.h"
#include "value.h"

namespace formulon {
namespace eval {

/// Stack-machine opcode catalog.
///
/// Operand mnemonics in the comment after each opcode use this convention:
///   - `K` is a constant-pool index (`a` field).
///   - `N` is a names-pool index (`a` field).
///   - `R` is a refs-pool index (`a` field).
///   - `arity` / `count` / `slot` are small unsigned counts (`a` or `b`).
///   - `target` is an absolute instruction index (`a` field).
///   - `op` is a sub-opcode for `BinaryOp` / `UnaryOp` / `Compare`.
///
/// Encodings deliberately keep at most one pool reference per instruction
/// so the operand layout stays uniform.
enum class OpCode : std::uint8_t {
  /// `LoadConst K` -- push `constants[K]` onto the operand stack.
  LoadConst = 0,

  /// `LoadRef R` -- push a `Ref` value resolving `refs[R]` (single cell or
  /// range endpoint) at run time.
  LoadRef = 1,

  /// `LoadRange` -- range marker emitted after the two endpoint
  /// subexpressions of a `NodeKind::RangeOp`. The `a` operand is the
  /// sentinel `0xFF` (not a pool index); the endpoints are carried by the
  /// two preceding instructions. For the common `Ref:Ref` shape those are
  /// two `LoadRef`s, and both the VM and the optimizer's
  /// range-canonicalisation pass read the refs-pool indices back from those
  /// `LoadRef` operands to reconstruct the rectangle. The VM expands it and
  /// pushes an `Array` value covering every cell in row-major order; when
  /// the endpoints are not both `LoadRef` (a complex range expression) the
  /// VM surfaces `#VALUE!`.
  LoadRange = 2,

  /// `LoadName N` -- resolve `names[N]` against the active name
  /// environment, push the resulting value (or a deferred lookup that the
  /// VM materialises into a body bytecode at first use).
  LoadName = 3,

  /// `LoadStructRef a=N(table) b=lo16(N column)|hi16(modifier)` -- a
  /// structured (table) reference. The encoding is documented in
  /// `compile_structured_ref()` in `compiler.cpp`.
  LoadStructRef = 4,

  /// `LoadSpillRef R` -- a spilled-range reference (`A1#`); same anchor
  /// payload as `LoadRef` but the VM expands the spill region.
  LoadSpillRef = 5,

  /// `LoadExternalRef a=N(sheet) b=lo16(book_id)|hi16(refs index)` -- an
  /// external workbook reference. The book id, sheet view, and cell
  /// reference are split across pools to keep the operand budget within
  /// 56 bits.
  LoadExternalRef = 6,

  /// `LoadLet a=slot` -- push the value bound at LET slot `a`.
  LoadLet = 7,

  /// `StoreLet a=slot` -- pop the top of stack and store it in LET slot
  /// `a`. Emitted as the second half of a `LetBinding` lowering.
  StoreLet = 8,

  /// `LoadLambdaArg a=slot` -- push the lambda call argument bound at
  /// slot `a` of the active call frame.
  LoadLambdaArg = 9,

  /// `Call a=N(name) b=arity` -- pop `arity` operands, look up
  /// `names[N]` in the function registry, push the result. Eager (Excel
  /// non-short-circuit) functions only.
  Call = 10,

  /// `CallLambda a=arity` -- pop `arity` operands plus a lambda value
  /// underneath them, push the result. The lambda value is the
  /// `arity+1`-th deepest stack slot.
  CallLambda = 11,

  /// `BinaryOp a=op` -- pop two operands, apply binary operator
  /// `parser::BinOp(a)`, push the result. Operators include arithmetic,
  /// concat, and the six comparisons. Excel error propagation is the
  /// VM's responsibility.
  BinaryOp = 12,

  /// `UnaryOp a=op` -- pop one operand, apply `parser::UnaryOp(a)`,
  /// push the result.
  UnaryOp = 13,

  /// `Concat` -- pop two operands, push their `&` concatenation. Sugar
  /// for `BinaryOp(BinOp::Concat)` to keep `&` lowering cheap to detect.
  Concat = 14,

  /// `MakeArray a=rows b=cols` -- pop `rows*cols` operands (row-major,
  /// last cell is top of stack) and push a single `Array` value.
  MakeArray = 15,

  /// `MakeLambda a=N(name array start) b=lo16(param_count)|hi16(optional_count)` --
  /// build a closure value over the immediately following sub-bytecode
  /// body. The body bytes are spliced inline via a `Jump` so the parent
  /// stream skips over them at run time. Detailed encoding lives in
  /// `compile_lambda()`.
  MakeLambda = 16,

  /// `Union a=count` -- pop `count` operands and push a single union
  /// value (Excel comma operator inside ranges).
  Union = 17,

  /// `Intersect` -- pop two operands and push their intersection (Excel
  /// space operator).
  Intersect = 18,

  /// `ImplicitIntersection` -- pop one operand, apply Excel's `@`
  /// implicit-intersection coercion, push the result.
  ImplicitIntersection = 19,

  /// `Jump target` -- unconditional branch to absolute instruction
  /// `target` (`a` field, encoded as 24-bit; for targets above 2^24 the
  /// compiler reports `kVmInstructionLimit`).
  Jump = 20,

  /// `JumpIfFalse target` -- pop one operand. If it coerces to FALSE
  /// (Excel rules: zero-numeric / FALSE / empty becomes false; error
  /// propagates by leaving the value on the stack and skipping the
  /// branch -- the VM contract is documented in Bundle 5.2's vm.cpp),
  /// branch to `target`; otherwise fall through.
  JumpIfFalse = 21,

  /// `Return` -- terminate execution; top of stack is the result.
  Return = 22,

  /// `Halt` -- internal sentinel emitted at the end of every body so the
  /// VM never falls off the end of the instruction stream. Indistinct
  /// from `Return` for the reference VM but kept separate so optimiser
  /// passes can tell synthetic from user-emitted RETs.
  Halt = 23,
};

/// Returns a stable textual mnemonic for `op`, useful for diagnostics.
/// The pointer references a static string literal.
constexpr const char* opcode_name(OpCode op) noexcept {
  switch (op) {
    case OpCode::LoadConst:
      return "LoadConst";
    case OpCode::LoadRef:
      return "LoadRef";
    case OpCode::LoadRange:
      return "LoadRange";
    case OpCode::LoadName:
      return "LoadName";
    case OpCode::LoadStructRef:
      return "LoadStructRef";
    case OpCode::LoadSpillRef:
      return "LoadSpillRef";
    case OpCode::LoadExternalRef:
      return "LoadExternalRef";
    case OpCode::LoadLet:
      return "LoadLet";
    case OpCode::StoreLet:
      return "StoreLet";
    case OpCode::LoadLambdaArg:
      return "LoadLambdaArg";
    case OpCode::Call:
      return "Call";
    case OpCode::CallLambda:
      return "CallLambda";
    case OpCode::BinaryOp:
      return "BinaryOp";
    case OpCode::UnaryOp:
      return "UnaryOp";
    case OpCode::Concat:
      return "Concat";
    case OpCode::MakeArray:
      return "MakeArray";
    case OpCode::MakeLambda:
      return "MakeLambda";
    case OpCode::Union:
      return "Union";
    case OpCode::Intersect:
      return "Intersect";
    case OpCode::ImplicitIntersection:
      return "ImplicitIntersection";
    case OpCode::Jump:
      return "Jump";
    case OpCode::JumpIfFalse:
      return "JumpIfFalse";
    case OpCode::Return:
      return "Return";
    case OpCode::Halt:
      return "Halt";
  }
  return "?";
}

/// 64-bit packed instruction word.
///
/// Field layout (little-endian on every target, but the bit-packing uses
/// arithmetic shifts so it is layout-independent):
///   - `op` (`std::uint8_t`)            : opcode, lower 8 bits.
///   - `a`  (`std::uint32_t`)           : first operand, 24 bits used.
///                                         Stored in a 32-bit field for
///                                         alignment. The compiler
///                                         enforces the 24-bit budget;
///                                         the upper 8 bits MUST be 0.
///   - `b`  (`std::uint32_t`)           : second operand, full 32 bits.
///
/// The struct is trivially copyable so it can be stored in a `std::vector`
/// without per-element constructors firing.
struct Instruction {
  /// Opcode tag.
  OpCode op = OpCode::Halt;
  /// Reserved padding so the explicit layout total reaches 8 bytes.
  /// Always zero; reserved for a future short third operand.
  std::uint8_t pad = 0;
  /// Reserved padding (16 bits) -- future expansion (e.g. flags on op).
  /// Always zero; the compiler never writes a non-zero value.
  std::uint16_t flags = 0;
  /// First operand (24-bit budget; upper 8 bits must be zero).
  std::uint32_t a = 0;
  /// Second operand (full 32-bit budget). Often unused; defaults to zero.
  std::uint32_t b = 0;

  /// Maximum value the `a` operand may carry before
  /// `kVmInstructionLimit` / `kVmConstPoolOverflow` fires.
  static constexpr std::uint32_t kMaxA = (1u << 24) - 1u;
};

static_assert(std::is_trivially_copyable_v<Instruction>, "Instruction must be trivially copyable");
static_assert(sizeof(Instruction) == 12 || sizeof(Instruction) == 16,
              "Instruction word should be a small multiple of 4 bytes");

/// Compiled formula body.
///
/// `ByteCode` is a self-contained value: every string referenced by a
/// `Call` / `LoadName` / `LoadStructRef` / `MakeLambda` is deep-copied
/// into `names`, every `Reference` is copied into `refs`, every literal
/// `Value` (and its underlying text storage) lives in `constants` /
/// `string_storage`. Two `ByteCode`s compiled from equivalent ASTs
/// compare equal under `operator==` (see implementation note below).
///
/// The structure is movable but not copyable; equality is provided as a
/// free function for test-only use.
struct ByteCode {
  /// Linear instruction stream. The VM dispatches on `code[pc].op`.
  std::vector<Instruction> code;

  /// Pool of literal values referenced by `LoadConst`. `Text` literals
  /// borrow from `string_storage`; the borrow is valid for the full
  /// lifetime of this `ByteCode`.
  std::vector<Value> constants;

  /// Pool of identifier names referenced by `Call` / `LoadName` /
  /// `LoadStructRef` / `MakeLambda`. Owns the bytes; views into this
  /// vector are stable across moves but invalidated on rebuild.
  std::vector<std::string> names;

  /// Pool of references referenced by `LoadRef` / `LoadRange` /
  /// `LoadSpillRef` / `LoadExternalRef`. The `sheet` view inside each
  /// `parser::Reference` borrows from `string_storage` so the
  /// `ByteCode` is fully self-contained.
  std::vector<parser::Reference> refs;

  /// Backing storage for borrow-only string views inside `constants`
  /// (Text values) and `refs` (sheet names). A `std::deque` is used
  /// deliberately: unlike `std::vector`, appending never relocates
  /// existing elements, so the `string_view`s that `constants` and `refs`
  /// hold into these strings stay valid for the full lifetime of the
  /// `ByteCode` no matter how many entries are interned. Callers must
  /// treat this pool as opaque and never mutate it.
  std::deque<std::string> string_storage;

  /// One entry per `code[i]` recording an opaque AST node identity used
  /// by the diagnostic emitter to map run-time errors back to a source
  /// span. The encoding is private to the compiler / diagnostic layer.
  std::vector<std::uint32_t> source_pos;
};

/// Trivial-shape equality, useful for tests that want to assert two
/// compiles of the same AST produce identical streams.
///
/// Two `ByteCode`s are equal iff:
///   - their `code` vectors have identical instruction words;
///   - their `constants` pools match element-wise via `Value::operator==`;
///   - their `names`, `refs`, and `string_storage` pools match
///     element-wise.
///
/// `source_pos` is intentionally excluded: the compiler may attribute
/// instructions to different parser nodes when the same logical formula
/// is reparsed (because AST node addresses differ). Tests that need to
/// inspect attribution should compare `source_pos` directly.
bool bytecode_shapes_equal(const ByteCode& lhs, const ByteCode& rhs) noexcept;

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BYTECODE_H_
