// Locale-aware text collation for Excel text comparison.
//
// On Mac profiles a string becomes an order-preserving byte key built from the
// tables in `collation_table.cpp`: primary, secondary, dynamic and
// compare-only tertiary weights, then the zh word-sort symbols and a trailing
// ignorable mark, the parts joined by 0x01. Comparing two strings is comparing
// their keys bytewise. Windows profiles build the same keys from the same
// tables under the win-365-ja_JP host rules (collation.cpp).

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
///     compare-only tertiary level (zh and Windows full/half width).
///   * `kClassicExact`: exact MATCH / VLOOKUP / HLOOKUP; no expansions, no
///     half-width voicing composition and no combining-mark composition.
enum class CollationStrength : std::uint8_t { kCompare, kDynamic, kClassicExact };

/// Order-preserving key of `s`: `collate(a, b)` has the sign of
/// `collate_key_compare(collation_key(a), collation_key(b))`.
std::string collation_key(std::string_view s, ExcelProfile profile, CollationStrength strength);

/// Returns -1, 0 or 1 as `a` sorts before, equal to or after `b`.
int collate(std::string_view a, std::string_view b, ExcelProfile profile, CollationStrength strength);

/// Compares two keys from `collation_key` (same profile and strength); -1, 0 or 1.
int collate_key_compare(std::string_view key_a, std::string_view key_b) noexcept;

/// One fixed operand of repeated `collate(cell, query, ...)` calls. The query's key is built once, so `compare(cell)`
/// returns exactly what `collate(cell, query, profile, strength)` does while keying only the cell.
class CollateQuery {
 public:
  CollateQuery(std::string_view query, ExcelProfile profile, CollationStrength strength);

  int compare(std::string_view cell) const;

  /// False when the query takes the ASCII fast path, where `collate` itself is as cheap as `compare`.
  bool keyed() const noexcept { return !ascii_; }

 private:
  std::string query_;
  // Empty when the query takes the ASCII fast path, whose comparison needs no key.
  std::string key_;
  ExcelProfile profile_;
  CollationStrength strength_;
  bool ascii_;
};

/// D-function prefix rule: truncate `cell` to the code-point count of `criterion`,
/// then compare at `kCompare` for equality.
bool collate_prefix_equal(std::string_view cell, std::string_view criterion, ExcelProfile profile);

}  // namespace eval
}  // namespace formulon

#endif  // FORMULON_EVAL_COLLATION_H_
