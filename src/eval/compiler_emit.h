//
// Shared bytecode side-pool accessors.
//
// The VM (`vm.cpp`) and any future bytecode tooling (disassembler,
// optimiser, parity-test inspector) all need to read individual entries
// out of a `ByteCode`'s three side pools (constants, names, refs) with a
// bounds check that surfaces malformed bytecode as a structured `Error`.
//
// The compiler-side emit helpers (`push_constant` / `push_name` /
// `push_ref`) are intentionally NOT extracted here because their
// signatures are tightly coupled to the compiler's per-body `BodyState`
// struct (which carries the active `ByteCode*` plus LET / lambda scopes
// and the slot allocator). Hoisting them to a header would require
// hoisting `BodyState` too, which would either bloat the public API or
// force an awkward template indirection. Per W8-2, this header therefore
// extracts only the read-side accessors that genuinely have no compiler-
// specific dependencies.
//
// Header-only. All functions are `inline` so the linker collapses them to
// a single definition across TUs.

#ifndef FORMULON_EVAL_COMPILER_EMIT_H_
#define FORMULON_EVAL_COMPILER_EMIT_H_

#include <cstdint>
#include <string>

#include "eval/bytecode.h"
#include "parser/reference.h"
#include "utils/error.h"
#include "utils/expected.h"
#include "value.h"

namespace formulon {
namespace eval {

namespace detail {

/// Internal helper: builds an "invalid bytecode" error with the supplied
/// message. Kept private so the public accessors below stay a single line.
inline Error make_pool_oob_error(const char* msg) {
  return make_error(FormulonErrorCode::kVmInvalidOpcode, msg);
}

}  // namespace detail

/// Returns a borrowed pointer to `bc.constants[idx]`, or an
/// `kVmInvalidOpcode` error if `idx` is out of range.
///
/// The pointer is valid for the lifetime of `bc`. Callers that want to
/// hold on to the underlying `Value` beyond that lifetime (e.g. across
/// VM boundaries) must deep-copy the text payload via the arena — see
/// `vm.cpp`'s `LoadConst` handler for the canonical pattern.
inline Expected<const Value*, Error> const_at(const ByteCode& bc, std::uint32_t idx) {
  if (idx >= bc.constants.size()) {
    return detail::make_pool_oob_error("constants index out of range");
  }
  return &bc.constants[idx];
}

/// Returns a borrowed pointer to `bc.names[idx]`, or an
/// `kVmInvalidOpcode` error if `idx` is out of range.
///
/// The pointer is valid for the lifetime of `bc`; the underlying string
/// storage is owned by the names pool itself.
inline Expected<const std::string*, Error> name_at(const ByteCode& bc, std::uint32_t idx) {
  if (idx >= bc.names.size()) {
    return detail::make_pool_oob_error("names index out of range");
  }
  return &bc.names[idx];
}

/// Returns a borrowed pointer to `bc.refs[idx]`, or an `kVmInvalidOpcode`
/// error if `idx` is out of range.
///
/// The pointer is valid for the lifetime of `bc`. The embedded
/// `parser::Reference::sheet` view borrows from `bc.string_storage` per
/// the `ByteCode` self-contained-borrow contract documented in
/// `bytecode.h`.
inline Expected<const parser::Reference*, Error> ref_at(const ByteCode& bc, std::uint32_t idx) {
  if (idx >= bc.refs.size()) {
    return detail::make_pool_oob_error("refs index out of range");
  }
  return &bc.refs[idx];
}

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COMPILER_EMIT_H_
