//
// Cross-kind ordering and display helpers shared by every TU in the
// pivot evaluator. Header-only because the comparator (`ValueLess`) is
// the key type of `std::map<Value, HierNode, ValueLess>` and the
// hierarchy builder, filter engine, and aggregator all need to render
// the same labels via `display_string`.
//
// Excel's default sort on a row/column field is ascending by the
// field's display value. The cache stores discrete values as the typed
// `Value` they came from, so we need a total order that:
//
//   * groups same-kind values together (so we don't interleave numbers
//     into strings);
//   * sorts numerically within `Number`;
//   * sorts case-sensitively within `Text` (good enough for MVP — the
//     primary oracle is ja-JP which Excel collates byte-wise on the
//     wire; the locale-aware compare will arrive with date grouping);
//   * is total across kinds so `std::map` stays well-formed.
//
// Kind ordering: Number < Text < Bool < Error < everything else, sourced
// from the shared `excel_kind_rank` so the pivot comparator and the
// GROUPBY / SORT comparator (`eval/groupby_pivotby/common.cpp`) cannot drift
// on the relative position of Bool vs Text. The cache only ever produces the
// first four kinds for discrete fields, but the fall-through keeps the
// comparator total in case a future cache field surfaces an Array or similar.

#ifndef FORMULON_PIVOT_VALUE_ORDER_H_
#define FORMULON_PIVOT_VALUE_ORDER_H_

#include <cstdint>
#include <string>

#include "eval/jp_fold.h"
#include "utils/double_format.h"
#include "utils/error.h"
#include "value.h"
#include "value_sort_order.h"

namespace formulon::pivot {

inline bool value_less(const Value& a, const Value& b) noexcept {
  const int ra = excel_kind_rank(a.kind());
  const int rb = excel_kind_rank(b.kind());
  if (ra != rb) {
    return ra < rb;
  }
  switch (a.kind()) {
    case ValueKind::Number:
      return a.as_number() < b.as_number();
    case ValueKind::Bool:
      // false < true.
      return !a.as_boolean() && b.as_boolean();
    case ValueKind::Text:
      // Match the GROUPBY / SORT comparator (`groupby_pivotby/common.cpp`)
      // by folding through `fold_jp_text` before the byte compare, so the
      // two families cannot diverge on Japanese text collation (half-width
      // kana, hiragana/katakana, full-width ASCII).
      return eval::fold_jp_text(a.as_text()) < eval::fold_jp_text(b.as_text());
    case ValueKind::Error:
      return static_cast<std::uint16_t>(a.as_error()) < static_cast<std::uint16_t>(b.as_error());
    default:
      // Same kind, no in-kind ordering defined — treat as equal.
      return false;
  }
}

struct ValueLess {
  bool operator()(const Value& a, const Value& b) const noexcept { return value_less(a, b); }
};

/// Renders a cache `Value` to the same string that would appear in a
/// `PivotItem::name`. Mirrors what the OOXML reader would have written
/// out: numbers as their canonical decimal, bools as TRUE/FALSE, errors
/// as their `#…` token, blanks as the empty string. This is sufficient
/// for matching against `PivotItem::name`, which is the only place we
/// use it (manual-filter visibility check).
inline std::string display_string(const Value& v) {
  switch (v.kind()) {
    case ValueKind::Blank:
      return std::string{};
    case ValueKind::Number: {
      // Use the shared shortest-round-trip renderer rather than
      // `std::to_string`, which appends six fixed decimals and used to make
      // a numeric item named "1.5" fail to match GETPIVOTDATA's label.
      // It also avoids casting large doubles to long long.
      std::string out;
      format_double(out, v.as_number());
      return out;
    }
    case ValueKind::Bool:
      return v.as_boolean() ? "TRUE" : "FALSE";
    case ValueKind::Text:
      return std::string{v.as_text()};
    case ValueKind::Error:
      return std::string{display_name(v.as_error())};
    default:
      return std::string{};
  }
}

}  // namespace formulon::pivot

#endif  // FORMULON_PIVOT_VALUE_ORDER_H_
