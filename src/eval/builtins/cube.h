// Copyright 2026 libraz. Licensed under the MIT License.
//
// Registers Excel's Cube-category built-ins into a FunctionRegistry. All
// seven CUBE* functions require an OLAP cube connection that Formulon does
// not establish; each is registered as a deterministic #NAME? stub that
// matches what Excel itself returns when no cube data model is available.
//
// Registered functions:
//   * CUBEKPIMEMBER(connection, kpi_name, kpi_property, [caption])  (3-4)
//   * CUBEMEMBER(connection, member_expression, [caption])          (2-3)
//   * CUBEMEMBERPROPERTY(connection, member_expression, property)   (3)
//   * CUBERANKEDMEMBER(connection, set_expression, rank, [caption]) (3-4)
//   * CUBESET(connection, set, [caption], [sort_order], [sort_by])  (2-5)
//   * CUBESETCOUNT(set)                                              (1)
//   * CUBEVALUE(connection, [member_expression], ...)               (1..*)
//
// All are eager scalar (dispatched via the default FunctionDef path); no
// tree-walker changes are needed.

#ifndef FORMULON_EVAL_BUILTINS_CUBE_H_
#define FORMULON_EVAL_BUILTINS_CUBE_H_

namespace formulon {
namespace eval {

class FunctionRegistry;

/// Registers the Cube-category built-ins (all seven as deterministic
/// `#NAME?` stubs) into `registry`. Intended to be invoked from
/// `register_builtins`.
void register_cube_builtins(FunctionRegistry& registry);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_BUILTINS_CUBE_H_
