// Worksheet-name identity helpers.
//
// Worksheet names use Unicode simple case folding for identity. This is
// intentionally separate from the ASCII-only string helpers: those helpers
// are also used for function names, logical literals, table names, and other
// byte-oriented contracts that must not acquire Unicode semantics.

#ifndef FORMULON_SHEET_NAME_H_
#define FORMULON_SHEET_NAME_H_

#include <string_view>

namespace formulon {
namespace sheet_names {

/// Returns true when `lhs` and `rhs` have the same worksheet identity under
/// Unicode simple case folding (Unicode CaseFolding C/S mappings, no
/// normalization or full/expanding mappings).
///
/// The comparison decodes both strings as strict UTF-8 and returns false for
/// malformed input, including when both malformed inputs have identical
/// bytes. Structural worksheet-name validation rejects malformed UTF-8 before
/// names enter a workbook, so the exact-byte fast path is useful for valid
/// ASCII and for callers that already own the validation contract.
bool equal(std::string_view lhs, std::string_view rhs) noexcept;

/// Returns true when `name` is well-formed UTF-8. Unicode scalar values are
/// decoded strictly: overlong encodings, surrogate code points, truncated
/// sequences, stray continuation bytes, and values above U+10FFFF are
/// rejected.
bool valid_utf8(std::string_view name) noexcept;

}  // namespace sheet_names
}  // namespace formulon

#endif  // FORMULON_SHEET_NAME_H_
