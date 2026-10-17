//
// Implementation of the Cube-category built-in stubs.
//
// The Cube family (CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
// CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE) talks to an OLAP
// data model through Excel's connection layer. Formulon is a standalone
// calculation engine with no connection infrastructure, so every CUBE*
// function returns `#N/A`, which is what Excel 365 returns for a CUBE*
// formula whose connection does not exist.
//
// Argument evaluation: each stub's `impl` simply returns the fixed error
// once invoked. The dispatcher (`tree_walker.cpp`) pre-evaluates every
// argument with `propagate_errors = true` (the default), so an error in
// any argument short-circuits to `#DIV/0!` / `#VALUE!` / etc. before this
// body runs. That preserves Excel's "errors beat stubs" invariant.

#include "eval/builtins/cube.h"

#include <cstdint>

#include "eval/builtins/registration_helpers.h"
#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Single shared impl: every CUBE* function resolves to `#N/A` once its
// arguments have all evaluated cleanly. Using one function pointer keeps
// the catalog footprint minimal and makes the "stub" contract literal in
// the source.
Value CubeNoConnectionStub(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::NA);
}

}  // namespace

void register_cube_builtins(FunctionRegistry& registry) {
  // Arities as documented by Microsoft for Excel 365. Enforced by the
  // dispatcher's min/max arity check before the stub body runs.
  static constexpr builtins_detail::BuiltinRegistration functions[] = {
      {"CUBEKPIMEMBER", 3u, 4u, &CubeNoConnectionStub},
      {"CUBEMEMBER", 2u, 3u, &CubeNoConnectionStub},
      {"CUBEMEMBERPROPERTY", 3u, 3u, &CubeNoConnectionStub},
      {"CUBERANKEDMEMBER", 3u, 4u, &CubeNoConnectionStub},
      {"CUBESET", 2u, 5u, &CubeNoConnectionStub},
      {"CUBESETCOUNT", 1u, 1u, &CubeNoConnectionStub},
      // CUBEVALUE is variadic: connection plus 0..N member_expression tuples.
      {"CUBEVALUE", 1u, kVariadic, &CubeNoConnectionStub},
  };
  builtins_detail::register_builtin_functions(registry, functions, sizeof(functions) / sizeof(functions[0]));
}

}  // namespace eval
}  // namespace formulon
