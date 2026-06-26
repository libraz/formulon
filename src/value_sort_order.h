// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Canonical Excel cross-kind ordering rank, shared by every comparator that
// needs to sort heterogeneous `Value` kinds the way Excel's ascending value
// sort does. Keeping the kind-to-bucket mapping in one place stops the
// GROUPBY / SORT comparator (`eval/groupby_pivotby/common.cpp`) and the pivot
// comparator (`pivot/value_order.h`) from drifting apart on the relative
// position of Bool vs Text.
//
// Excel's ascending value order is: numbers, then text, then logical (FALSE
// before TRUE), then errors, then blanks. Array / Ref / Lambda never appear
// in user-visible sortable data; they are ranked last so the comparator stays
// total for `std::map` and `std::sort` consumers.

#ifndef FORMULON_VALUE_SORT_ORDER_H_
#define FORMULON_VALUE_SORT_ORDER_H_

#include "value.h"

namespace formulon {

/// Maps a `ValueKind` to its Excel ascending-sort bucket. Lower rank sorts
/// first. The order is Number < Text < Bool < Error < Blank < (Array / Ref /
/// Lambda). Within a bucket the caller compares the typed payload itself.
inline int excel_kind_rank(ValueKind k) noexcept {
  switch (k) {
    case ValueKind::Number:
      return 0;
    case ValueKind::Text:
      return 1;
    case ValueKind::Bool:
      return 2;
    case ValueKind::Error:
      return 3;
    case ValueKind::Blank:
      return 4;
    case ValueKind::Array:
      return 5;
    case ValueKind::Ref:
      return 6;
    case ValueKind::Lambda:
      return 7;
  }
  return 8;
}

}  // namespace formulon

#endif  // FORMULON_VALUE_SORT_ORDER_H_
