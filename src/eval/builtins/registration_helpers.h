// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.

#ifndef FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_

#include <cstddef>
#include <cstdint>
#include <string_view>

#include "eval/function_registry.h"

namespace formulon {
namespace eval {
namespace builtins_detail {

using BuiltinImpl = Value (*)(const Value*, std::uint32_t, Arena&);

struct BuiltinRegistration {
  std::string_view name;
  std::uint32_t min_arity;
  std::uint32_t max_arity;
  BuiltinImpl impl;
  bool propagate_errors = true;
  bool accepts_ranges = false;
  bool range_filter_numeric_only = false;
  bool range_filter_bool_coercible = false;
  bool range_filter_a_coerce = false;
  FunctionDef::BlankScalarPolicy blank_scalar_policy = FunctionDef::BlankScalarPolicy::Allow;
  ErrorCode blank_scalar_error = ErrorCode::Value;
};

inline void register_builtin_functions(FunctionRegistry& registry, const BuiltinRegistration* defs, std::size_t count) {
  for (std::size_t i = 0; i < count; ++i) {
    const BuiltinRegistration& def = defs[i];
    FunctionDef fn{def.name, def.min_arity, def.max_arity, def.impl, def.propagate_errors};
    fn.accepts_ranges = def.accepts_ranges;
    fn.range_filter_numeric_only = def.range_filter_numeric_only;
    fn.range_filter_bool_coercible = def.range_filter_bool_coercible;
    fn.range_filter_a_coerce = def.range_filter_a_coerce;
    fn.blank_scalar_policy = def.blank_scalar_policy;
    fn.blank_scalar_error = def.blank_scalar_error;
    registry.register_function(fn);
  }
}

}  // namespace builtins_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_
