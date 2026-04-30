// Copyright 2026 libraz. Licensed under the MIT License.
//
// ByteCode optimiser implementation.
//
// See `optimizer.h` for the public contract and the design rationale
// for each pass. The four passes are stand-alone static functions and
// are run from `optimize()` in the documented fixed order:
//
//   constant fold -> name inline -> range canonicalise -> branch hoist
//
// Each pass produces a fresh `ByteCode` rather than mutating the input
// in place. That keeps the per-pass logic linear (no in-place index
// shuffling) and makes it cheap for a future bundle to gate any
// individual pass behind a flag without disturbing the others.
//
// Jump remapping. Whenever a pass shortens the instruction stream, the
// pass builds a parallel `old_to_new[]` array mapping each input pc to
// its position in the output stream. After the rewrite, every `Jump` /
// `JumpIfFalse` operand is rewritten through that map. Source-position
// attribution (`bc.source_pos`) is kept in lockstep with `bc.code`.

#include "eval/optimizer.h"

#include <algorithm>
#include <cstdint>
#include <cstring>
#include <utility>
#include <vector>

#include "eval/bytecode.h"
#include "eval/coerce.h"
#include "eval/scalar_ops.h"
#include "parser/ast.h"
#include "parser/reference.h"
#include "utils/arena.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

// Local copies of the status macros used by `compiler.cpp`. The shared
// macros in `utils/status_macros.h` lean on `__COUNTER__`, which under
// `-Werror` is rejected by Emscripten's bundled clang as a C2y
// extension. We mirror the compiler-tu strategy and use `__LINE__` here.
#define FM_OPT_CONCAT_INNER(a, b) a##b
#define FM_OPT_CONCAT(a, b) FM_OPT_CONCAT_INNER(a, b)
#define FM_OPT_UNIQUE(prefix) FM_OPT_CONCAT(prefix, __LINE__)

#define FM_OPT_RETURN_IF_ERROR(expr) \
  do {                               \
    auto _fm_status = (expr);        \
    if (!_fm_status) {               \
      return _fm_status.error();     \
    }                                \
  } while (0)

#define FM_OPT_ASSIGN_OR_RETURN_IMPL(tmp, lhs, expr) \
  auto tmp = (expr);                                 \
  if (!tmp) {                                        \
    return tmp.error();                              \
  }                                                  \
  lhs = std::move(tmp.value())

#define FM_OPT_ASSIGN_OR_RETURN(lhs, expr) FM_OPT_ASSIGN_OR_RETURN_IMPL(FM_OPT_UNIQUE(_fm_opt_tmp_), lhs, expr)

namespace formulon {
namespace eval {

namespace {

/// Returns true if `op` writes to a 24-bit absolute pc target in `a`.
bool is_jump_op(OpCode op) noexcept {
  return op == OpCode::Jump || op == OpCode::JumpIfFalse;
}

/// Returns the "logical lexicographic" order of two References, matching
/// the ordering used by `tree_walker.cpp::eval_range`'s run-time
/// normalisation: same sheet, then column, then row. Sheet quote / abs
/// flags do not participate in ordering — they are layout-formatting
/// hints, not addressing data.
///
/// The comparison is total: it returns -1, 0, or +1, and never reports
/// equal for two Reference values that differ in `is_full_col` /
/// `is_full_row` (those never participate in the canonicalisation
/// pattern, so we keep the relation strict for safety).
int compare_refs(const parser::Reference& a, const parser::Reference& b) noexcept {
  if (a.sheet != b.sheet) {
    return a.sheet < b.sheet ? -1 : 1;
  }
  if (a.is_full_col != b.is_full_col) {
    return a.is_full_col ? -1 : 1;
  }
  if (a.is_full_row != b.is_full_row) {
    return a.is_full_row ? -1 : 1;
  }
  if (a.col != b.col) {
    return a.col < b.col ? -1 : 1;
  }
  if (a.row != b.row) {
    return a.row < b.row ? -1 : 1;
  }
  return 0;
}

/// Returns true iff `a` and `b` denote the same cell ignoring round-trip
/// cosmetic flags (`sheet_quoted`). Used to detect degenerate ranges.
bool refs_address_equal(const parser::Reference& a, const parser::Reference& b) noexcept {
  return a.sheet == b.sheet && a.col == b.col && a.row == b.row && a.col_abs == b.col_abs && a.row_abs == b.row_abs &&
         a.is_full_col == b.is_full_col && a.is_full_row == b.is_full_row;
}

/// Returns the constant-pool index for `v` after appending it to `out`.
/// Bounds-checks against the 24-bit operand budget. Text values are
/// re-interned into `out.string_storage` so the resulting payload is
/// self-contained, mirroring `compiler.cpp::push_constant`.
Expected<std::uint32_t, Error> append_constant(ByteCode& out, Value v) {
  if (out.constants.size() > Instruction::kMaxA) {
    return make_error(FormulonErrorCode::kVmConstPoolOverflow,
                      "optimizer: constants pool exceeds 24-bit operand budget");
  }
  if (v.kind() == ValueKind::Text) {
    out.string_storage.emplace_back(v.as_text());
    const std::string& slot = out.string_storage.back();
    v = Value::text(std::string_view(slot.data(), slot.size()));
  }
  out.constants.push_back(v);
  return static_cast<std::uint32_t>(out.constants.size() - 1);
}

/// Rewrites the `a` operand of every Jump / JumpIfFalse in `bc.code`
/// through `old_to_new`. The map MUST cover every input pc (size ==
/// original code length); positions absorbed by a fold/canonicalise map
/// to the surviving instruction's new pc, so jumps that landed on those
/// positions still resolve correctly.
void remap_jumps(ByteCode& bc, const std::vector<std::uint32_t>& old_to_new) noexcept {
  for (auto& ins : bc.code) {
    if (!is_jump_op(ins.op)) {
      continue;
    }
    const std::uint32_t old_target = ins.a;
    if (old_target < old_to_new.size()) {
      ins.a = old_to_new[old_target];
    } else if (old_target == old_to_new.size()) {
      // Jumps to the one-past-the-last position are valid (e.g. an IF
      // whose else branch is the very last instruction). The new
      // one-past-the-last pc is the new code length.
      // The caller fills this in from outside the loop.
    }
  }
}

/// Folds a binary-op result. Returns a `Value` if folding is safe; the
/// outer pass is responsible for pooling it. Returns nullopt-equivalent
/// (false in `Expected<>::has_value()` is reserved for engine errors;
/// here we use a discriminator: `is_blank()` on the returned `Value`
/// is impossible for a real fold result so we return Value::blank()
/// ONLY when we choose to skip the fold). To keep the contract clean we
/// instead return a small struct.
struct FoldResult {
  bool ok = false;
  Value value = Value::blank();
};

FoldResult try_fold_binary(parser::BinOp op, const Value& lhs, const Value& rhs, Arena& arena) {
  // Guardrails: never fold across a non-scalar constant. The current
  // compiler never emits a `LoadConst` of an Array / Lambda / Ref, but
  // we defend against it anyway so a future change to the compiler
  // cannot silently break the optimiser.
  const ValueKind lk = lhs.kind();
  const ValueKind rk = rhs.kind();
  auto is_scalar = [](ValueKind k) noexcept {
    return k == ValueKind::Number || k == ValueKind::Bool || k == ValueKind::Text || k == ValueKind::Error ||
           k == ValueKind::Blank;
  };
  if (!is_scalar(lk) || !is_scalar(rk)) {
    return {};
  }
  switch (op) {
    case parser::BinOp::Add:
    case parser::BinOp::Sub:
    case parser::BinOp::Mul:
    case parser::BinOp::Div:
    case parser::BinOp::Pow: {
      // Mirror the VM's BinaryOp handler exactly: error short-circuits
      // left-to-right, then numeric coerce both sides, then apply.
      if (lhs.is_error()) {
        return {true, lhs};
      }
      if (rhs.is_error()) {
        return {true, rhs};
      }
      auto ln = coerce_to_number(lhs);
      if (!ln) {
        return {true, Value::error(ln.error())};
      }
      auto rn = coerce_to_number(rhs);
      if (!rn) {
        return {true, Value::error(rn.error())};
      }
      return {true, apply_arithmetic(op, ln.value(), rn.value())};
    }
    case parser::BinOp::Concat:
      // Concat result is a `Text` whose backing bytes live in `arena`.
      // The optimiser does not have a stable arena that outlives the
      // returned `ByteCode`, so we deliberately skip this fold. (See
      // optimizer.h header note.)
      (void)arena;
      return {};
    case parser::BinOp::Eq:
    case parser::BinOp::NotEq:
    case parser::BinOp::Lt:
    case parser::BinOp::LtEq:
    case parser::BinOp::Gt:
    case parser::BinOp::GtEq:
      if (lhs.is_error()) {
        return {true, lhs};
      }
      if (rhs.is_error()) {
        return {true, rhs};
      }
      return {true, apply_comparison(op, lhs, rhs)};
  }
  return {};
}

FoldResult try_fold_unary(parser::UnaryOp op, const Value& operand) {
  const ValueKind k = operand.kind();
  if (k != ValueKind::Number && k != ValueKind::Bool && k != ValueKind::Text && k != ValueKind::Error &&
      k != ValueKind::Blank) {
    return {};
  }
  return {true, apply_unary(op, operand)};
}

// ---------------------------------------------------------------------------
// Pass 1: constant folding.
// ---------------------------------------------------------------------------

Expected<ByteCode, Error> fold_constants_pass(ByteCode in, Arena& arena, OptimizerStats* stats) {
  ByteCode out;
  out.constants = std::move(in.constants);
  out.names = std::move(in.names);
  out.refs = std::move(in.refs);
  out.string_storage = std::move(in.string_storage);
  out.code.reserve(in.code.size());
  out.source_pos.reserve(in.source_pos.size());

  std::vector<std::uint32_t> old_to_new(in.code.size() + 1, 0);

  // Prebuild a "is this input pc a jump target" bitmap so the fold check
  // is O(1) per instruction (overall pass: O(N)). A linear in-loop scan
  // would otherwise make the pass O(N^2) for synthesised long expression
  // chains.
  std::vector<bool> is_target(in.code.size() + 1, false);
  for (const auto& j : in.code) {
    if (is_jump_op(j.op) && j.a < is_target.size()) {
      is_target[j.a] = true;
    }
  }

  for (std::size_t i = 0; i < in.code.size(); ++i) {
    const Instruction& ins = in.code[i];
    // Try the binary-fold pattern first. We need three consecutive ops
    // and the previous two output entries to be `LoadConst`. We must
    // also guarantee that none of the three absorbed positions is a
    // jump target — folding them away would orphan an inbound jump.
    //
    // Chained folds: the rejection composes correctly. If a fold-1
    // succeeds at position `i1`, none of the input pcs in
    // `[i1-2, i1]` were jump targets. A subsequent fold-2 at `i2` can
    // only succeed if `[i2-2, i2]` are also non-targets. The set of
    // absorbed pcs in fold-2 includes `[i1-2, i1]` transitively only
    // through the surviving constant slot, but the fold result Value
    // is what the jump would have observed at the original target —
    // and since no jump targeted those positions, no observer cares.
    if ((ins.op == OpCode::BinaryOp || ins.op == OpCode::Concat) && out.code.size() >= 2 &&
        out.code[out.code.size() - 1].op == OpCode::LoadConst &&
        out.code[out.code.size() - 2].op == OpCode::LoadConst) {
      const std::size_t lhs_in_pc = i - 2;
      const std::size_t rhs_in_pc = i - 1;
      const bool any_target = is_target[lhs_in_pc] || is_target[rhs_in_pc] || is_target[i];
      if (!any_target) {
        const std::uint32_t k_lhs = out.code[out.code.size() - 2].a;
        const std::uint32_t k_rhs = out.code[out.code.size() - 1].a;
        const Value lhs = out.constants[k_lhs];
        const Value rhs = out.constants[k_rhs];
        const auto op = static_cast<parser::BinOp>(
            ins.op == OpCode::Concat ? static_cast<std::uint32_t>(parser::BinOp::Concat) : ins.a);
        FoldResult fr = try_fold_binary(op, lhs, rhs, arena);
        if (fr.ok) {
          // Pop the two predecessor LoadConsts and emit a single
          // LoadConst of the folded value. Source-position attribution
          // is taken from the BinaryOp / Concat itself, since that is
          // the user-visible expression position the folded value
          // belongs to.
          out.code.pop_back();
          out.code.pop_back();
          out.source_pos.pop_back();
          out.source_pos.pop_back();
          // Drop the corresponding old_to_new entries we set when
          // mirroring those LoadConsts.
          old_to_new[lhs_in_pc] = static_cast<std::uint32_t>(out.code.size());
          old_to_new[rhs_in_pc] = static_cast<std::uint32_t>(out.code.size());
          FM_OPT_ASSIGN_OR_RETURN(auto k_new, append_constant(out, fr.value));
          Instruction folded{};
          folded.op = OpCode::LoadConst;
          folded.a = k_new;
          out.code.push_back(folded);
          out.source_pos.push_back(in.source_pos[i]);
          old_to_new[i] = static_cast<std::uint32_t>(out.code.size() - 1);
          if (stats != nullptr) {
            ++stats->constants_folded;
          }
          continue;
        }
      }
    }

    if (ins.op == OpCode::UnaryOp && !out.code.empty() && out.code.back().op == OpCode::LoadConst) {
      const std::size_t operand_in_pc = i - 1;
      const bool any_target = is_target[operand_in_pc] || is_target[i];
      if (!any_target) {
        const std::uint32_t k_op = out.code.back().a;
        const Value operand = out.constants[k_op];
        const auto op = static_cast<parser::UnaryOp>(ins.a);
        FoldResult fr = try_fold_unary(op, operand);
        if (fr.ok) {
          out.code.pop_back();
          out.source_pos.pop_back();
          old_to_new[operand_in_pc] = static_cast<std::uint32_t>(out.code.size());
          FM_OPT_ASSIGN_OR_RETURN(auto k_new, append_constant(out, fr.value));
          Instruction folded{};
          folded.op = OpCode::LoadConst;
          folded.a = k_new;
          out.code.push_back(folded);
          out.source_pos.push_back(in.source_pos[i]);
          old_to_new[i] = static_cast<std::uint32_t>(out.code.size() - 1);
          if (stats != nullptr) {
            ++stats->constants_folded;
          }
          continue;
        }
      }
    }

    // Default: copy through.
    old_to_new[i] = static_cast<std::uint32_t>(out.code.size());
    out.code.push_back(ins);
    out.source_pos.push_back(in.source_pos[i]);
  }

  // One-past-the-last entry: jumps that landed at the original end now
  // land at the new end.
  old_to_new.back() = static_cast<std::uint32_t>(out.code.size());

  remap_jumps(out, old_to_new);
  return out;
}

// ---------------------------------------------------------------------------
// Pass 2: name inlining (stub; reserved for a workbook-aware overload).
// ---------------------------------------------------------------------------

Expected<ByteCode, Error> inline_names_pass(ByteCode in, Arena& /*arena*/, OptimizerStats* /*stats*/) {
  // No-op. The real implementation needs `Workbook&` to resolve a
  // defined name to its target literal, and `compile()` does not have
  // that today. Bundle 5.4+ may add a `compile_with_workbook()` overload
  // that propagates the workbook here.
  return in;
}

// ---------------------------------------------------------------------------
// Pass 3: range canonicalisation.
// ---------------------------------------------------------------------------

Expected<ByteCode, Error> canonicalize_ranges_pass(ByteCode in, Arena& /*arena*/, OptimizerStats* stats) {
  ByteCode out;
  out.constants = std::move(in.constants);
  out.names = std::move(in.names);
  out.refs = std::move(in.refs);
  out.string_storage = std::move(in.string_storage);
  out.code.reserve(in.code.size());
  out.source_pos.reserve(in.source_pos.size());

  std::vector<std::uint32_t> old_to_new(in.code.size() + 1, 0);

  std::vector<bool> is_target(in.code.size() + 1, false);
  for (const auto& j : in.code) {
    if (is_jump_op(j.op) && j.a < is_target.size()) {
      is_target[j.a] = true;
    }
  }

  for (std::size_t i = 0; i < in.code.size(); ++i) {
    const Instruction& ins = in.code[i];
    if (ins.op == OpCode::LoadRange && ins.a == 0xFFu && out.code.size() >= 2 &&
        out.code[out.code.size() - 1].op == OpCode::LoadRef && out.code[out.code.size() - 2].op == OpCode::LoadRef) {
      const std::size_t lhs_in_pc = i - 2;
      const std::size_t rhs_in_pc = i - 1;
      const bool any_target = is_target[lhs_in_pc] || is_target[rhs_in_pc] || is_target[i];
      if (!any_target) {
        const std::uint32_t r_lhs = out.code[out.code.size() - 2].a;
        const std::uint32_t r_rhs = out.code[out.code.size() - 1].a;
        const parser::Reference& lhs = out.refs[r_lhs];
        const parser::Reference& rhs = out.refs[r_rhs];
        // Degenerate range: A:A collapses to a single LoadRef.
        if (refs_address_equal(lhs, rhs)) {
          out.code.pop_back();
          out.source_pos.pop_back();
          // Drop the `LoadRange` itself by skipping the default emit
          // below; map both the second LoadRef and LoadRange pcs to the
          // surviving first LoadRef new pc.
          old_to_new[lhs_in_pc] = static_cast<std::uint32_t>(out.code.size() - 1);
          old_to_new[rhs_in_pc] = static_cast<std::uint32_t>(out.code.size() - 1);
          old_to_new[i] = static_cast<std::uint32_t>(out.code.size() - 1);
          if (stats != nullptr) {
            ++stats->ranges_canonicalized;
          }
          continue;
        }
        // Non-degenerate: ensure the two endpoints are in canonical
        // order. If not, swap their `a` operands in the two LoadRef
        // instructions; the refs pool entries themselves do not move.
        if (compare_refs(lhs, rhs) > 0) {
          out.code[out.code.size() - 2].a = r_rhs;
          out.code[out.code.size() - 1].a = r_lhs;
          if (stats != nullptr) {
            ++stats->ranges_canonicalized;
          }
        }
      }
    }
    old_to_new[i] = static_cast<std::uint32_t>(out.code.size());
    out.code.push_back(ins);
    out.source_pos.push_back(in.source_pos[i]);
  }

  old_to_new.back() = static_cast<std::uint32_t>(out.code.size());
  remap_jumps(out, old_to_new);
  return out;
}

// ---------------------------------------------------------------------------
// Pass 4: branch hoisting (skeleton).
// ---------------------------------------------------------------------------

Expected<ByteCode, Error> hoist_branches_pass(ByteCode in, Arena& /*arena*/, OptimizerStats* stats) {
  // Recognise the shape:
  //
  //   <cond>
  //   JumpIfFalse Lfalse
  //   LoadConst K1
  //   Jump Lend
  //   Lfalse: LoadConst K2
  //   Lend:
  //
  // and (for now) only count it. Rewriting requires a `Select` opcode
  // we have not yet introduced; leaving the pattern in place keeps VM
  // semantics intact.
  for (std::size_t i = 0; i + 4 < in.code.size(); ++i) {
    if (in.code[i].op != OpCode::JumpIfFalse) {
      continue;
    }
    const std::uint32_t lfalse = in.code[i].a;
    if (in.code[i + 1].op != OpCode::LoadConst || in.code[i + 2].op != OpCode::Jump) {
      continue;
    }
    const std::uint32_t lend = in.code[i + 2].a;
    if (lfalse != static_cast<std::uint32_t>(i + 3)) {
      continue;
    }
    if (in.code[i + 3].op != OpCode::LoadConst) {
      continue;
    }
    if (lend != static_cast<std::uint32_t>(i + 4)) {
      continue;
    }
    if (stats != nullptr) {
      ++stats->branch_hoist_opportunities;
    }
  }
  return in;
}

}  // namespace

// ---------------------------------------------------------------------------
// Public entry point.
// ---------------------------------------------------------------------------

Expected<ByteCode, Error> optimize(ByteCode bc, Arena& arena, OptimizerStats* stats) {
  FM_OPT_ASSIGN_OR_RETURN(ByteCode after_fold, fold_constants_pass(std::move(bc), arena, stats));
  FM_OPT_ASSIGN_OR_RETURN(ByteCode after_inline, inline_names_pass(std::move(after_fold), arena, stats));
  FM_OPT_ASSIGN_OR_RETURN(ByteCode after_canon, canonicalize_ranges_pass(std::move(after_inline), arena, stats));
  FM_OPT_ASSIGN_OR_RETURN(ByteCode after_hoist, hoist_branches_pass(std::move(after_canon), arena, stats));
  return after_hoist;
}

}  // namespace eval
}  // namespace formulon
