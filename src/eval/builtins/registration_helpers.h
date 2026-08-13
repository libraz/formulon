
#ifndef FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_
#define FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_

#include <array>
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
  // The auto policy follows the dispatcher: range-aware eager built-ins
  // reduce their flattened inputs, while ordinary eager built-ins broadcast
  // array-valued children. A registration may override this for a less
  // common shape, including an intrinsic array producer.
  FunctionDef::ResultShape result_shape = FunctionDef::ResultShape::kAuto;
};

inline bool registration_name_is(std::string_view name, std::string_view expected) noexcept {
  if (name.size() != expected.size()) {
    return false;
  }
  for (std::size_t i = 0; i < name.size(); ++i) {
    const char lhs = name[i] >= 'a' && name[i] <= 'z' ? static_cast<char>(name[i] - ('a' - 'A')) : name[i];
    if (lhs != expected[i]) {
      return false;
    }
  }
  return true;
}

template <std::size_t N>
inline bool registration_name_in(std::string_view name, const std::array<std::string_view, N>& names) noexcept {
  for (const std::string_view expected : names) {
    if (registration_name_is(name, expected)) {
      return true;
    }
  }
  return false;
}

inline FunctionDef::ResultShape infer_builtin_result_shape(std::string_view name, bool accepts_ranges) noexcept {
  static constexpr std::array<std::string_view, 17> kBroadcast = {
      "ABS",  "EXP",  "INT",   "LN",    "LOG", "N",        "ROUND",   "ROUNDDOWN", "ROUNDUP",
      "SIGN", "SQRT", "VALUE", "POWER", "MOD", "QUOTIENT", "CEILING", "FLOOR"};
  static constexpr std::array<std::string_view, 16> kReduce = {
      "SUM",        "SUMSQ",  "PRODUCT", "AVERAGE", "MIN",   "MAX",    "COUNT", "COUNTA",
      "COUNTBLANK", "CONCAT", "ROWS",    "COLUMNS", "AREAS", "SHEETS", "GCD",   "LCM"};
  if (registration_name_in(name, kBroadcast)) {
    return FunctionDef::ResultShape::kBroadcast;
  }
  if (registration_name_in(name, kReduce)) {
    return FunctionDef::ResultShape::kReduce;
  }
  return accepts_ranges ? FunctionDef::ResultShape::kReduce : FunctionDef::ResultShape::kBroadcast;
}

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
    fn.result_shape = def.result_shape == FunctionDef::ResultShape::kAuto
                          ? infer_builtin_result_shape(def.name, def.accepts_ranges)
                          : def.result_shape;
    registry.register_function(fn);
  }
}

}  // namespace builtins_detail
}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_REGISTRATION_HELPERS_H_
