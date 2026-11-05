//
// Deterministic ids for conditional-format rules that travel through the x14
// extension, in the shape Excel writes them.

#ifndef FORMULON_CF_CF_X14_ID_H_
#define FORMULON_CF_CF_X14_ID_H_

#include <cstddef>
#include <cstdint>
#include <string>
#include <unordered_set>

namespace formulon::cf {

/// Builds `{00000000-000E-0000-SSSS-0000NNNNNNNN}` where `SSSS` is the
/// zero-based sheet index as a little-endian u16 and `NNNNNNNN` the ordinal as
/// a little-endian u32 (uppercase hex). While the id is in `used_ids`, the
/// ordinal is incremented.
std::string x14_rule_id(std::size_t sheet_index, std::uint32_t ordinal,
                        const std::unordered_set<std::string>& used_ids);

}  // namespace formulon::cf

#endif  // FORMULON_CF_CF_X14_ID_H_
