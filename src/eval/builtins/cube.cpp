// Copyright 2026 libraz. Licensed under the MIT License.
//
// Implementation of the Cube-category built-in stubs.
//
// The Cube family (CUBEKPIMEMBER, CUBEMEMBER, CUBEMEMBERPROPERTY,
// CUBERANKEDMEMBER, CUBESET, CUBESETCOUNT, CUBEVALUE) talks to an OLAP
// data model through Excel's connection layer. Formulon is a standalone
// calculation engine with no connection infrastructure, so every CUBE*
// function returns `#NAME?`.
//
// Why #NAME? rather than #VALUE!: Mac Excel 365 itself surfaces `#NAME?`
// for CUBE* formulas when no cube connection exists (the function name is
// recognised, but the semantic runtime is missing). Matching Excel's
// no-connection output keeps Formulon's behaviour deterministic and
// oracle-friendly should a connectionless fixture ever be introduced.
//
// Argument evaluation: each stub's `impl` simply returns the fixed error
// once invoked. The dispatcher (`tree_walker.cpp`) pre-evaluates every
// argument with `propagate_errors = true` (the default), so an error in
// any argument short-circuits to `#DIV/0!` / `#VALUE!` / etc. before this
// body runs. That preserves Excel's "errors beat stubs" invariant.

#include "eval/builtins/cube.h"

#include <cstdint>

#include "eval/function_registry.h"
#include "utils/arena.h"
#include "value.h"

namespace formulon {
namespace eval {
namespace {

// Single shared impl: every CUBE* function resolves to `#NAME?` once its
// arguments have all evaluated cleanly. Using one function pointer keeps
// the catalog footprint minimal and makes the "stub" contract literal in
// the source.
Value CubeNameStub(const Value* /*args*/, std::uint32_t /*arity*/, Arena& /*arena*/) {
  return Value::error(ErrorCode::Name);
}

}  // namespace

void register_cube_builtins(FunctionRegistry& registry) {
  // Arities as documented by Microsoft for Excel 365. Enforced by the
  // dispatcher's min/max arity check before the stub body runs.
  registry.register_function(FunctionDef{"CUBEKPIMEMBER", 3u, 4u, &CubeNameStub});
  registry.register_function(FunctionDef{"CUBEMEMBER", 2u, 3u, &CubeNameStub});
  registry.register_function(FunctionDef{"CUBEMEMBERPROPERTY", 3u, 3u, &CubeNameStub});
  registry.register_function(FunctionDef{"CUBERANKEDMEMBER", 3u, 4u, &CubeNameStub});
  registry.register_function(FunctionDef{"CUBESET", 2u, 5u, &CubeNameStub});
  registry.register_function(FunctionDef{"CUBESETCOUNT", 1u, 1u, &CubeNameStub});
  // CUBEVALUE is variadic: connection plus 0..N member_expression tuples.
  registry.register_function(FunctionDef{"CUBEVALUE", 1u, kVariadic, &CubeNameStub});
}

}  // namespace eval
}  // namespace formulon
