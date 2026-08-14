// Generated Unicode simple-case-fold data interface.

#ifndef FORMULON_SHEET_NAME_CASEFOLD_DATA_H_
#define FORMULON_SHEET_NAME_CASEFOLD_DATA_H_

#include <cstdint>

namespace formulon {
namespace sheet_names {
namespace detail {

/// Applies the generated Unicode CaseFolding C/S mapping to one scalar.
/// Scalars without a simple mapping are returned unchanged.
std::uint32_t fold_scalar(std::uint32_t scalar) noexcept;

}  // namespace detail
}  // namespace sheet_names
}  // namespace formulon

#endif  // FORMULON_SHEET_NAME_CASEFOLD_DATA_H_
