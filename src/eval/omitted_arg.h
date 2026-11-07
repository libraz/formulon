//
// Identifies a syntactically omitted function argument (`f(a,,c)`). This is
// distinct from an evaluated blank reference: omission carries the callee's
// documented default (`FunctionDef::omitted_defaults` for eager built-ins)
// rather than ordinary blank coercion.
//
// Also hosts the former Analysis-ToolPak argument rule, shared by the eager
// dispatcher (`FunctionDef::analysis_toolpak_args`) and the lazy impls that
// evaluate their own arguments: a boolean argument is `#VALUE!`, an omitted
// required slot is `#N/A`, and an omitted optional slot keeps its default.

#ifndef FORMULON_EVAL_OMITTED_ARG_H_
#define FORMULON_EVAL_OMITTED_ARG_H_

#include <cstdint>

#include "eval/function_registry.h"
#include "parser/ast.h"
#include "value.h"

namespace formulon {
namespace eval {

inline bool is_omitted_arg(const parser::AstNode& arg) {
  return arg.kind() == parser::NodeKind::Literal && arg.as_literal().is_blank();
}

/// The default `def` declares for an omitted slot `index`, or null.
inline const OmittedArgDefault* omitted_arg_default(const FunctionDef& def, std::uint32_t index) noexcept {
  for (std::uint32_t i = 0; i < def.omitted_default_count; ++i) {
    if (def.omitted_defaults[i].slot == index) {
      return &def.omitted_defaults[i];
    }
  }
  return nullptr;
}

/// Writes the value an omitted slot `index` of `def` takes into `*out`.
/// Returns false when the slot has no number or boolean default.
inline bool omitted_arg_value(const FunctionDef& def, std::uint32_t index, Value* out) {
  const OmittedArgDefault* d = omitted_arg_default(def, index);
  if (d == nullptr || d->kind == OmittedArgDefault::Kind::kAbsent) {
    return false;
  }
  *out = d->kind == OmittedArgDefault::Kind::kBool ? Value::boolean(d->number != 0.0) : Value::number(d->number);
  return true;
}

/// `arity` less the trailing omitted slots of `call` whose default is
/// `kAbsent`, so the callee sees them as not supplied.
inline std::uint32_t trim_absent_omitted_args(const parser::AstNode& call, const FunctionDef& def,
                                              std::uint32_t arity) {
  while (arity > def.min_arity && is_omitted_arg(call.as_call_arg(arity - 1U))) {
    const OmittedArgDefault* d = omitted_arg_default(def, arity - 1U);
    if (d == nullptr || d->kind != OmittedArgDefault::Kind::kAbsent) {
      break;
    }
    --arity;
  }
  return arity;
}

/// Whether an omitted slot at `index` is a required argument under the
/// Analysis-ToolPak rule: below `min_arity`, or any slot of a variadic.
inline bool atp_slot_required(std::uint32_t index, std::uint32_t min_arity, std::uint32_t max_arity) noexcept {
  return index < min_arity || max_arity == kVariadic;
}

/// Number of leading argument slots of `call` an Analysis-ToolPak function
/// evaluates. Trailing omitted optional slots are dropped so the callee
/// applies its defaults (`WEEKNUM(10,)` is `WEEKNUM(10)`).
inline std::uint32_t atp_evaluated_arity(const parser::AstNode& call, std::uint32_t min_arity,
                                         std::uint32_t max_arity) {
  std::uint32_t arity = call.as_call_arity();
  if (max_arity == kVariadic) {
    return arity;
  }
  while (arity > min_arity && is_omitted_arg(call.as_call_arg(arity - 1U))) {
    --arity;
  }
  return arity;
}

/// The error an Analysis-ToolPak argument slot surfaces before the callee
/// sees it, or `Value::blank()` when the slot passes. `v` is the slot's
/// evaluated scalar value; an omitted slot passes `Value::blank()`. A
/// blank-cell reference passes and coerces to 0 as usual. A `logical`
/// parameter (ACCRINT's `calc_method`) accepts a boolean.
inline Value atp_arg_error(const parser::AstNode& arg, const Value& v, bool required, bool logical = false) {
  if (is_omitted_arg(arg)) {
    return required ? Value::error(ErrorCode::NA) : Value::blank();
  }
  return v.is_boolean() && !logical ? Value::error(ErrorCode::Value) : Value::blank();
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_OMITTED_ARG_H_
