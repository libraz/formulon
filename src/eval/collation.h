// Locale-aware text collation for Excel text comparison.
//
// On Mac profiles a string becomes an order-preserving byte key built from the
// measured tables in `collation_table.cpp`: primary, secondary, dynamic and
// compare-only tertiary weights, then the zh word-sort symbols and a trailing
// ignorable mark, the parts joined by 0x01. Comparing two strings is comparing
// their keys bytewise. Windows profiles keep the engine's byte model: ASCII
// case-insensitive unsigned byte order, with half-width voicing marks composed
// for the dynamic and classic-exact strengths.

#ifndef FORMULON_EVAL_COLLATION_H_
#define FORMULON_EVAL_COLLATION_H_

#include <cstdint>
#include <string>
#include <string_view>

#include "excel_profile.h"

namespace formulon {
namespace eval {

/// How much of the collation order a comparison sees.
///   * `kCompare`: operators, criteria, SWITCH, approximate MATCH / VLOOKUP.
///   * `kDynamic`: SORT, UNIQUE, GROUPBY, XLOOKUP / XMATCH, pivots; ignores the
///     compare-only tertiary level (zh full/half width).
///   * `kClassicExact`: exact MATCH / VLOOKUP / HLOOKUP; on Mac, no expansions
///     and no half-width voicing composition (the Windows host composes it).
enum class CollationStrength : std::uint8_t { kCompare, kDynamic, kClassicExact };

/// Order-preserving key of `s`: `collate(a, b)` has the sign of
/// `collate_key_compare(collation_key(a), collation_key(b))`.
std::string collation_key(std::string_view s, ExcelProfile profile, CollationStrength strength);

/// Returns -1, 0 or 1 as `a` sorts before, equal to or after `b`.
int collate(std::string_view a, std::string_view b, ExcelProfile profile, CollationStrength strength);

/// Compares two keys from `collation_key` (same profile and strength); -1, 0 or 1.
int collate_key_compare(std::string_view key_a, std::string_view key_b) noexcept;

/// D-function prefix rule: truncate `cell` to the code-point count of `criterion`,
/// then compare at `kCompare` for equality.
bool collate_prefix_equal(std::string_view cell, std::string_view criterion, ExcelProfile profile);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COLLATION_H_
