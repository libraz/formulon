//
// Status codes the language bindings raise themselves, before or instead of
// a C-ABI call. The WASM and Node layers share these values; they mirror
// `formulon::FormulonErrorCode` without including it.

#ifndef FORMULON_C_API_BINDING_STATUS_H_
#define FORMULON_C_API_BINDING_STATUS_H_

#include <cstdint>

namespace formulon {
namespace c_api {

/// A JS wrapper was used after its workbook handle was destroyed. This is
/// the binding-layer `kBindingInvalidHandle` code, not the C ABI's 7001
/// `kBindingNullPointer` argument-validation error.
inline constexpr std::int32_t kBindingInvalidHandle = 7000;

/// A JS callable handed to the binding threw. Reported on the envelope of
/// the call that ran it.
inline constexpr std::int32_t kBindingCallbackException = 7003;

/// `FormulonErrorCode::kInvalidArgument`: the status of every argument the
/// binding itself rejects before any C call.
inline constexpr std::int32_t kInvalidArgument = 2;

}  // namespace c_api
}  // namespace formulon

#endif  // FORMULON_C_API_BINDING_STATUS_H_
